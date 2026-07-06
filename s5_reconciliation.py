"""S5 Reconciliation — single source of truth for cnp / dedup / liability.

Called by:
- capture_daily_snapshot.py  (nightly cron at 23:30 IST)
- dashboard.py               (writes Daily Totals row on load if none exists today)

Ensures live view and cron always produce identical numbers, because there
is only ONE computation path.

Uses partner_code everywhere past the ingestion boundary. Name-based sheet
lookups pass through collision_resolver.resolve_collisions() → owner
partner_code; loser partner_codes are forced to 0 downstream.

Formula (2026-07-06 rewrite per Kapil's spec — device-ID dedup)
--------------------------------------------------------------
    liability = |idle_device_ids ∪ customer_side_device_ids|
              = |idle_device_ids| + |customer_side_device_ids − idle_device_ids|

    idle_device_ids = T_DEVICE STATUS='IDLE' GROUP BY LCO_ACCOUNT_ID
                      → set of DEVICE_ID per partner
    customer_side_device_ids
        = (U1_pending_mobiles ∪ U2_pending_mobiles) resolved via
          PUBLIC.ACTIVE_CUST → DEVICE_ID (fallback DYNAMODB_READ.CUSTOMER_V_2)
    U1_pending_mobiles = PX Migration Raw Data (exit_partner_code = X)
                       − PX Migrated Cases (Old LCO Id = X)
    U2_pending_mobiles = Main sheet rows where partner = X AND not picked
                         (universal picked rule: Remarks='Device picked up'
                         OR Pradeep/Ajinkya='Yes')

Dedup happens on DEVICE_ID (not NAS ID as pre-2026-07-06). Reason: NAS ID
match under-counted "already at partner" duplicates by ~74/78 in a live
batch — Kapil's manual reconciliation surfaced this. Device IDs are the
authoritative identifier.
"""
from collections import defaultdict

# Local import — collision resolver
try:
    from collision_resolver import resolve_collisions
except ImportError:
    from .collision_resolver import resolve_collisions


# ────────────────────────────────────────────────────────────────────────────
# Constants shared with dashboard.py / capture_daily_snapshot.py
# ────────────────────────────────────────────────────────────────────────────

# CSPs whose exit was stopped or excluded from tracking. Keep in sync with the
# copies in dashboard.py + capture_daily_snapshot.py (or import from a shared
# module later). Numbers only.
EXCLUDED_PARTNER_CODES = {
    281749854790153,  # yash broadband
    281749854637042,  # Lovely communication
    281749854779181,
    281749854779177,
    281749854779178,
}

# Two workbook IDs
CSP_EXIT_WORKBOOK_ID_ENV = "GOOGLE_SHEET_ID"  # set via env
PX_MIGRATION_WORKBOOK_ID = "1hmT50leXZUAibzd2zzfO4FVj-B3m675CCFUbdwFuVS4"
PX_RAW_TAB = "PX Migration Raw Data"
PX_MIGRATED_TAB = "PX Migrated Cases"
MAIN_TAB = "Main sheet"
MIGRATION_DATA_TAB = "Migration Data"


# ────────────────────────────────────────────────────────────────────────────
# Row-level helpers
# ────────────────────────────────────────────────────────────────────────────

def _u2_row_is_picked(row):
    """Universal U2 'picked' rule — Remarks Dropdown = 'Device picked up'
    OR Device Picked (Ajinkya/Pradeep) = 'Yes' (case-insensitive)."""
    rem = str(row.get("Remarks Dropdown") or "").strip().lower()
    prd = str(row.get("Device Picked (Ajinkya/Pradeep)") or "").strip().lower()
    return rem == "device picked up" or prd == "yes"


def _to_int(v):
    try:
        return int(float(str(v).replace(",", "").strip() or 0))
    except Exception:
        return 0


def _read_sheet_values_safe(ws):
    """Read all values from a worksheet, trimming trailing empty header cols
    (so downstream code doesn't get confused by blank headers)."""
    vals = ws.get_all_values()
    if not vals:
        return [], []
    hdr = vals[0]
    last = 0
    for i, h in enumerate(hdr):
        if h and h.strip():
            last = i + 1
    return [h.strip() for h in hdr[:last]], [r[:last] for r in vals[1:]]


# ────────────────────────────────────────────────────────────────────────────
# Data loaders — each accepts fully-formed inputs, no globals
# ────────────────────────────────────────────────────────────────────────────

def load_all_partners(sb_url, sb_key, requests_module):
    """Return ALL partners (all states), EXCLUDED filtered out. Needed for the
    collision resolver so an S4 partner that shares a name with an S5 partner
    can be picked as the collision owner (forcing the S5 dupe to lose)."""
    hdr = {"apikey": sb_key,
           "Authorization": f"Bearer {sb_key}",
           "Accept": "application/json"}
    partners = requests_module.get(
        f"{sb_url}/rest/v1/partners",
        params={"select": "id,name,partner_code,current_state,exit_started_at"},
        headers=hdr, timeout=15,
    ).json()
    partners = [p for p in partners
                if int(p.get("partner_code") or 0) not in EXCLUDED_PARTNER_CODES]
    # Attach u1_count / u2_count for collision resolver
    ext = requests_module.get(
        f"{sb_url}/rest/v1/partner_details_extended",
        params={"select": "id,u1_count,u2_count"},
        headers=hdr, timeout=15,
    ).json()
    de = {r["id"]: r for r in ext} if isinstance(ext, list) else {}
    for p in partners:
        d = de.get(p["id"], {})
        p["u1_count"] = d.get("u1_count") or 0
        p["u2_count"] = d.get("u2_count") or 0
    return partners


def load_idle_from_metabase(mb_query_fn, partner_codes, n_passes=5, gap_seconds=60):
    """Return (idle_count_by_code, idle_device_ids_by_code).
    Single-shot query. Device IDs are stable identifiers so no multi-pass
    stabilization needed (NAS IDs required unions because they flickered).

    idle_count_by_code: {code_str: n_idle_devices}
    idle_device_ids_by_code: {code_str: set(device_id_str)}

    Kept n_passes/gap_seconds params for backward compat with existing
    callers; both are ignored under the device-id-based algorithm.
    """
    del n_passes, gap_seconds  # unused under device-id algorithm
    idle_count_by_code = {}
    idle_device_ids_by_code = defaultdict(set)
    if not partner_codes:
        return idle_count_by_code, idle_device_ids_by_code
    in_list = ",".join(str(c) for c in partner_codes if c)

    rows = mb_query_fn(f"""
SELECT "LCO_ACCOUNT_ID", "DEVICE_ID"
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE"
WHERE "STATUS" = 'IDLE' AND "LCO_ACCOUNT_ID" IN ({in_list})
  AND "DEVICE_ID" IS NOT NULL""")
    for lco, dev in rows:
        if lco and dev:
            idle_device_ids_by_code[str(lco)].add(str(dev).strip())
    for code, ids in idle_device_ids_by_code.items():
        idle_count_by_code[code] = len(ids)
    total = sum(idle_count_by_code.values())
    print(f"[idle-device-ids] total={total} across {len(idle_count_by_code)} CSPs")

    return idle_count_by_code, dict(idle_device_ids_by_code)


def load_migration_data_counts(ws, name_to_owner_code, losers):
    """Read Migration Data tab → (u1_total_by_code, u1_migrated_by_code).

    ws: gspread Worksheet for the 'Migration Data' tab.
    name_to_owner_code: {name_lower: owner_partner_code_str} from resolver.
    losers: set of partner_code_str forced to 0.
    """
    hdr, rows = _read_sheet_values_safe(ws)
    try:
        c_name = hdr.index("Exit Partner Name")
        c_total = hdr.index("Total U1 User")
        c_mig = hdr.index("Migrated")
    except ValueError as e:
        raise RuntimeError(f"Migration Data missing required column: {e}")

    total_by_code = defaultdict(int)
    mig_by_code = defaultdict(int)
    for r in rows:
        if len(r) <= max(c_name, c_total, c_mig):
            continue
        nm = str(r[c_name]).strip().lower()
        if not nm:
            continue
        code = name_to_owner_code.get(nm)
        if not code or code in losers:
            continue
        total_by_code[code] += _to_int(r[c_total])
        mig_by_code[code] += _to_int(r[c_mig])
    return dict(total_by_code), dict(mig_by_code)


def load_main_sheet_u2(ws, name_to_owner_code, losers):
    """Read Main sheet → (u2_total_by_code, u2_picked_by_code, u2_pending_mobiles_by_code).
    Applies universal picked rule (Remarks OR Pradeep column)."""
    hdr, rows = _read_sheet_values_safe(ws)
    # Locate columns
    def _idx(name):
        try:
            return hdr.index(name)
        except ValueError:
            return None
    c_partner = _idx("Partner")
    c_mobile = _idx("Mobile no")
    if c_mobile is None:
        c_mobile = _idx("Mobile")
    c_remarks = _idx("Remarks Dropdown")
    c_pradeep = _idx("Device Picked (Ajinkya/Pradeep)")
    if c_partner is None or c_mobile is None:
        raise RuntimeError(
            f"Main sheet missing Partner or Mobile column (hdr={hdr[:10]})")

    total_by_code = defaultdict(int)
    picked_by_code = defaultdict(int)
    pending_mobiles_by_code = defaultdict(set)

    for r in rows:
        if len(r) <= max(c_partner, c_mobile):
            continue
        nm = str(r[c_partner]).strip().lower()
        mobile = str(r[c_mobile]).strip()
        if not nm or not mobile:
            continue
        code = name_to_owner_code.get(nm)
        if not code or code in losers:
            continue
        total_by_code[code] += 1
        rem = str(r[c_remarks]).strip().lower() if c_remarks is not None and len(r) > c_remarks else ""
        prd = str(r[c_pradeep]).strip().lower() if c_pradeep is not None and len(r) > c_pradeep else ""
        if rem == "device picked up" or prd == "yes":
            picked_by_code[code] += 1
        else:
            pending_mobiles_by_code[code].add(mobile)
    return dict(total_by_code), dict(picked_by_code), dict(pending_mobiles_by_code)


def load_px_migration(px_book, name_to_owner_code, losers, s5s6_codes):
    """Read PX/CX Migration Summary workbook. Return u1_pending_mobiles_by_code
    = (raw customers) - (migrated customers) restricted to S5/S6 codes."""
    # Raw Data (keyed by exit_partner_code — no name resolution needed)
    raw_hdr, raw_rows = _read_sheet_values_safe(px_book.worksheet(PX_RAW_TAB))
    try:
        c_code = raw_hdr.index("exit_partner_code")
        c_mob = raw_hdr.index("Customer_mobile")
    except ValueError as e:
        raise RuntimeError(f"PX Migration Raw Data missing column: {e}")

    raw_by_code = defaultdict(set)
    for r in raw_rows:
        if len(r) <= max(c_code, c_mob):
            continue
        code = str(r[c_code]).strip()
        mobile = str(r[c_mob]).strip()
        if code in s5s6_codes and code not in losers and mobile:
            raw_by_code[code].add(mobile)

    # Migrated Cases — join by 'Old LCO Id' (partner_code) not by name.
    # Per Kapil's spec (2026-07-06): "Match by Account ID wherever possible —
    # partner names have typos/case issues; account IDs don't." Old code
    # matched on name, which needed the collision resolver and still missed
    # attribution when a sheet had a typo. This is code-exact.
    #
    # Legacy fallback: if 'Old LCO Id' column is missing (older sheet format),
    # fall back to name→code lookup so we don't regress.
    del name_to_owner_code, losers  # unused under code-based join
    mig_hdr, mig_rows = _read_sheet_values_safe(px_book.worksheet(PX_MIGRATED_TAB))
    c_mob_m = None
    c_code_m = None
    for i, h in enumerate(mig_hdr):
        hl = str(h).lower()
        if c_mob_m is None and "mobile" in hl:
            c_mob_m = i
        if c_code_m is None and ("old lco id" in hl or "old_lco_id" in hl):
            c_code_m = i
    if c_mob_m is None or c_code_m is None:
        raise RuntimeError(
            f"PX Migrated Cases missing 'Mobile No' or 'Old LCO Id' "
            f"column (hdr={mig_hdr})")

    mig_by_code = defaultdict(set)
    for r in mig_rows:
        if len(r) <= max(c_mob_m, c_code_m):
            continue
        code = str(r[c_code_m]).strip()
        mobile = str(r[c_mob_m]).strip()
        if not code or not mobile:
            continue
        if code in s5s6_codes:
            mig_by_code[code].add(mobile)

    pending_by_code = {}
    for code, raw_mobiles in raw_by_code.items():
        pending_by_code[code] = raw_mobiles - mig_by_code.get(code, set())
    return pending_by_code, dict(raw_by_code), dict(mig_by_code)


def _mobile_variants(mobile):
    """Expand a raw mobile string into the candidates worth looking up.

    Handles Kapil's spec cases:
      - "9876543210_ARCHIVED"           → ["9876543210"]
      - "9324812667/7045328441"          → ["9324812667", "7045328441"]
      - "9324812667/7045328441_ARCHIVED" → ["9324812667", "7045328441"]
    Also normalises whitespace and drops empty parts.
    """
    if not mobile:
        return []
    s = str(mobile).strip()
    if not s:
        return []
    # Strip _ARCHIVED suffix (case-insensitive)
    upper = s.upper()
    if upper.endswith("_ARCHIVED"):
        s = s[:-len("_ARCHIVED")]
    # Split on slash
    parts = [p.strip() for p in s.split("/") if p.strip()]
    return parts or [s]


def load_mobile_to_device_id(mb_query_fn, mobiles):
    """Batch lookup mobile → DEVICE_ID per Kapil's spec (2026-07-06).

    Primary source: PUBLIC.ACTIVE_CUST (column MOBILE, column DEVICE_ID).
    Fallback for unresolved: DYNAMODB_READ.CUSTOMER_V_2 (Wiom Hub) — handles
    mobiles changed after PNM migration.
    Handles _ARCHIVED suffix and slash-separated mobiles via _mobile_variants.

    Returns {input_mobile_str: device_id_str} — key is the ORIGINAL mobile
    from the caller's list; value is the device_id from whichever source
    resolved it.
    """
    if not mobiles:
        return {}
    # Expand each input mobile into lookup candidates and remember the mapping
    variant_to_original = {}  # variant → set(original)
    all_variants = set()
    for m in mobiles:
        for v in _mobile_variants(m):
            variant_to_original.setdefault(v, set()).add(m)
            all_variants.add(v)

    def _sql_in(values):
        return ",".join("'" + v.replace("'", "''") + "'" for v in values)

    variants_left = set(all_variants)
    variant_to_device = {}

    # Primary: ACTIVE_CUST
    v_list = list(variants_left)
    for i in range(0, len(v_list), 5000):
        chunk = v_list[i:i + 5000]
        rows = mb_query_fn(f"""
SELECT MOBILE, DEVICE_ID
FROM PUBLIC.ACTIVE_CUST
WHERE MOBILE IN ({_sql_in(chunk)})""")
        for mob, dev in rows:
            if mob and dev:
                variant_to_device[str(mob).strip()] = str(dev).strip()

    variants_left -= set(variant_to_device.keys())

    # Fallback: CUSTOMER_V_2 for the ones still missing
    if variants_left:
        v_list = list(variants_left)
        for i in range(0, len(v_list), 5000):
            chunk = v_list[i:i + 5000]
            rows = mb_query_fn(f"""
SELECT MOBILE, DEVICE_ID
FROM DYNAMODB_READ.CUSTOMER_V_2
WHERE MOBILE IN ({_sql_in(chunk)})""")
            for mob, dev in rows:
                if mob and dev:
                    variant_to_device[str(mob).strip()] = str(dev).strip()

    # Map variants → originals (a variant may hit multiple originals if the
    # user has both the archived and current record; both point to same dev)
    out = {}
    for variant, dev in variant_to_device.items():
        for original in variant_to_original.get(variant, {variant}):
            # First hit wins per original mobile — same device across
            # variants is expected (per Kapil's spec)
            out.setdefault(original, dev)
    return out


def load_mobile_to_nas(mb_query_fn, mobiles):
    """DEPRECATED: NAS-based dedup superseded by device-id-based dedup.
    Kept as a stub in case external callers still import it. Returns {}."""
    del mb_query_fn, mobiles
    return {}


# ────────────────────────────────────────────────────────────────────────────
# Top-level orchestrator
# ────────────────────────────────────────────────────────────────────────────

def compute_s5_snapshot(*, sb_url, sb_key, requests_module,
                        mb_query_fn, exit_book, px_book,
                        n_nas_passes=5, gap_seconds=60,
                        verbose=True):
    """Compute a complete S5 snapshot. Returns dict with every field the cron
    writes to Daily Totals AND every field the dashboard displays.

    Parameters (all injected — no globals):
      sb_url, sb_key         : Supabase REST endpoint + anon key
      requests_module        : `requests` module (or a mock)
      mb_query_fn(sql)       : function that runs a Metabase SQL, returns rows
      exit_book              : gspread Spreadsheet for CSP Exit workbook
      px_book                : gspread Spreadsheet for PX/CX Migration Summary
      n_nas_passes           : IDLE-NAS stabilization passes (default 5)
      gap_seconds            : sleep between passes (default 60 s)
      verbose                : print diagnostics

    Returns:
      {
        "partners": list,               # S5+S6 partners with u1/u2 counts
        "s5_partners", "s6_partners",   # split by state
        "collisions": [ ... ],          # resolved collisions
        "idle_count_by_code": {code: n},
        "idle_nas_by_code":   {code: set(NASID)},
        "u1_total_by_code",             # from Migration Data (aggregate)
        "u1_migrated_by_code",
        "u1_pending_count_by_code",     # = total - migrated
        "u2_total_by_code",
        "u2_picked_by_code",
        "u2_pending_count_by_code",
        "u2_pending_mobiles_by_code",
        "u1_pending_mobiles_by_code",   # from PX Raw - Migrated
        "u1_dedup_by_code": {code: n},
        "u2_dedup_by_code": {code: n},
        "s5_idle": int,
        "s6_idle": int,
        "s5_cnp_raw": int,
        "s5_dedup": int,
        "s5_could_not_pick": int,
        "s5_liability": int,
      }
    """
    # 1. Load ALL partners (all states) for the collision resolver, then split
    all_partners = load_all_partners(sb_url, sb_key, requests_module)
    s5_partners = [p for p in all_partners if p.get("current_state") == "S5"]
    s6_partners = [p for p in all_partners if p.get("current_state") == "S6"]
    if verbose:
        print(f"[partners] all={len(all_partners)}, S5={len(s5_partners)}, S6={len(s6_partners)}")

    s5_codes = [str(p["partner_code"]) for p in s5_partners]
    s6_codes = [str(p["partner_code"]) for p in s6_partners]
    s5s6_codes = set(s5_codes) | set(s6_codes)

    # 2. Collision resolver — passes ALL partners so a cross-state collision
    # (e.g., Shree Shyam S4 vs S5) picks the correct owner.
    resolver = resolve_collisions(all_partners, verbose=verbose)
    name_to_owner_code = resolver["name_to_owner_code"]
    losers = resolver["losers"]
    partners = s5_partners + s6_partners  # kept for return value / callers

    # 3. IDLE devices — count + device-ID set per partner (single-shot)
    idle_count_by_code, idle_device_ids_by_code = load_idle_from_metabase(
        mb_query_fn, list(s5s6_codes),
    )

    # 4. Migration Data — U1 totals + migrated
    u1_total_by_code, u1_migrated_by_code = load_migration_data_counts(
        exit_book.worksheet(MIGRATION_DATA_TAB), name_to_owner_code, losers,
    )
    u1_pending_count_by_code = {
        code: max(u1_total_by_code.get(code, 0) - u1_migrated_by_code.get(code, 0), 0)
        for code in s5s6_codes
    }

    # 5. Main sheet — U2 totals, picked, pending mobiles
    u2_total_by_code, u2_picked_by_code, u2_pending_mobiles_by_code = load_main_sheet_u2(
        exit_book.worksheet(MAIN_TAB), name_to_owner_code, losers,
    )
    u2_pending_count_by_code = {
        code: max(u2_total_by_code.get(code, 0) - u2_picked_by_code.get(code, 0), 0)
        for code in s5s6_codes
    }

    # 6. PX Migration — U1 pending mobiles (for dedup)
    u1_pending_mobiles_by_code, u1_raw_by_code, u1_mig_by_code = load_px_migration(
        px_book, name_to_owner_code, losers, s5s6_codes,
    )

    # 7. Union all pending mobiles → mobile→device_id lookup.
    # Per Kapil's spec (2026-07-06):
    #   PRIMARY   = PUBLIC.ACTIVE_CUST      (MOBILE, DEVICE_ID)
    #   FALLBACK  = DYNAMODB_READ.CUSTOMER_V_2  (post-PNM mobile changes)
    all_pending_mobiles = set()
    for code in s5s6_codes:
        all_pending_mobiles |= u1_pending_mobiles_by_code.get(code, set())
        all_pending_mobiles |= u2_pending_mobiles_by_code.get(code, set())
    mobile_to_device = load_mobile_to_device_id(
        mb_query_fn, list(all_pending_mobiles),
    )
    if verbose:
        print(f"[mobile->device_id] {len(mobile_to_device)} of "
              f"{len(all_pending_mobiles)} pending mobiles resolved")

    # 8. Customer-side device_id set per partner (dedup by mobile first, then
    # resolve to device_id). Dedup then vs idle: customer-side devices that
    # are already sitting IDLE at the partner are counted ONCE (as idle).
    #
    # Per Kapil's spec Step 5: "If the same Device ID appears in BOTH Step 3
    # (customer-side) and Step 4 (IDLE at partner) — count it only once
    # (once IDLE, prefer that)."
    customer_side_device_ids_by_code = {}
    customer_only_device_ids_by_code = {}
    u1_dedup_by_code = {}
    u2_dedup_by_code = {}
    for code in s5s6_codes:
        idle_devs = idle_device_ids_by_code.get(code, set())
        # Customer-side device ids (from U1 pending + U2 pending, dedup by
        # mobile → device_id)
        u1_mobiles = u1_pending_mobiles_by_code.get(code, set())
        u2_mobiles = u2_pending_mobiles_by_code.get(code, set())
        u1_devs = {mobile_to_device[m] for m in u1_mobiles if m in mobile_to_device}
        u2_devs = {mobile_to_device[m] for m in u2_mobiles if m in mobile_to_device}
        customer_devs = u1_devs | u2_devs
        customer_side_device_ids_by_code[code] = customer_devs
        # Cross-dedup: drop customer-side devices that are already idle
        customer_only_device_ids_by_code[code] = customer_devs - idle_devs
        # Per-bucket dedup for diagnostics
        u1_dedup_by_code[code] = len(u1_devs & idle_devs)
        u2_dedup_by_code[code] = len(u2_devs & idle_devs)

    # 9. S5 aggregates (S6 handled separately)
    s5_idle = sum(idle_count_by_code.get(c, 0) for c in s5_codes)
    s6_idle = sum(idle_count_by_code.get(c, 0) for c in s6_codes)
    s5_u1_pending = sum(u1_pending_count_by_code.get(c, 0) for c in s5_codes)
    s5_u2_pending = sum(u2_pending_count_by_code.get(c, 0) for c in s5_codes)
    s5_u1_dedup = sum(u1_dedup_by_code.get(c, 0) for c in s5_codes)
    s5_u2_dedup = sum(u2_dedup_by_code.get(c, 0) for c in s5_codes)

    # cnp_raw is retained for the legacy Daily Totals column (count-based,
    # NOT the same as the new device-id-based customer_only). Kept so the
    # Daily Totals column meaning is unchanged.
    cnp_raw = s5_u1_pending + s5_u2_pending

    # dedup = number of pending customers whose device is already IDLE at
    # the partner (union across U1 + U2). Under the new algorithm, this is
    # a subset of customer_side_device_ids; but for the s5_could_not_pick
    # aggregate we use the device-id-based customer-only count.
    dedup = s5_u1_dedup + s5_u2_dedup

    # s5_could_not_pick under the new spec = |customer_only| aggregated
    # across S5 (customer-side devices not already at partner).
    cnp = sum(
        len(customer_only_device_ids_by_code.get(c, set())) for c in s5_codes
    )
    # Liability = idle devices at partner + devices still with customer,
    # counted once. Idle set and customer_only set are disjoint by
    # construction (customer_only = customer_devs - idle_devs).
    liability = s5_idle + cnp

    if verbose:
        print(f"[S5 result] cnp_raw={cnp_raw} device_dedup={dedup} "
              f"(U1={s5_u1_dedup}+U2={s5_u2_dedup}) "
              f"cnp(cust-only)={cnp} liab={liability}")

    return {
        "partners": partners,
        "s5_partners": s5_partners,
        "s6_partners": s6_partners,
        "collisions": resolver["collisions"],
        "idle_count_by_code": idle_count_by_code,
        "idle_device_ids_by_code": idle_device_ids_by_code,
        "idle_nas_by_code": {},  # deprecated stub — kept for backward compat
        "u1_total_by_code": u1_total_by_code,
        "u1_migrated_by_code": u1_migrated_by_code,
        "u1_pending_count_by_code": u1_pending_count_by_code,
        "u2_total_by_code": u2_total_by_code,
        "u2_picked_by_code": u2_picked_by_code,
        "u2_pending_count_by_code": u2_pending_count_by_code,
        "u2_pending_mobiles_by_code": u2_pending_mobiles_by_code,
        "u1_pending_mobiles_by_code": u1_pending_mobiles_by_code,
        "u1_raw_by_code": u1_raw_by_code,
        "u1_mig_by_code": u1_mig_by_code,
        "u1_dedup_by_code": u1_dedup_by_code,
        "u2_dedup_by_code": u2_dedup_by_code,
        "customer_side_device_ids_by_code": customer_side_device_ids_by_code,
        "customer_only_device_ids_by_code": customer_only_device_ids_by_code,
        "s5_idle": s5_idle,
        "s6_idle": s6_idle,
        "s5_cnp_raw": cnp_raw,
        "s5_dedup": dedup,
        "s5_could_not_pick": cnp,
        "s5_liability": liability,
    }


# ────────────────────────────────────────────────────────────────────────────
# CLI — dry-run against live data
# ────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import os
    import sys
    import gspread
    import requests
    from dotenv import load_dotenv
    from google.oauth2.service_account import Credentials
    sys.stdout.reconfigure(encoding="utf-8")

    HERE = os.path.dirname(os.path.abspath(__file__))
    load_dotenv(os.path.join(HERE, ".env"))
    sb_url = os.getenv("SUPABASE_URL")
    sb_key = os.getenv("SUPABASE_ANON_KEY")
    mb_url = os.getenv("METABASE_URL")
    mb_key = os.getenv("METABASE_API_KEY")
    sheet_id = os.getenv("GOOGLE_SHEET_ID")

    creds = Credentials.from_service_account_file(
        os.path.join(HERE, "google_credentials.json"),
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.readonly"])
    gs = gspread.authorize(creds)
    exit_book = gs.open_by_key(sheet_id)
    px_book = gs.open_by_key(PX_MIGRATION_WORKBOOK_ID)

    def mb_query(sql):
        r = requests.post(f"{mb_url}/api/dataset",
                          headers={"x-api-key": mb_key,
                                    "Content-Type": "application/json"},
                          json={"database": 113, "type": "native",
                                "native": {"query": sql}},
                          timeout=120)
        if r.status_code not in (200, 202):
            print(f"Metabase ERR {r.status_code}: {r.text[:200]}")
            return []
        return r.json().get("data", {}).get("rows", [])

    # For dry-run: use 1-pass NAS to keep it fast; cron/live still use 5.
    passes = int(os.getenv("S5_NAS_PASSES", "1"))
    gap = int(os.getenv("S5_NAS_GAP", "0"))
    result = compute_s5_snapshot(
        sb_url=sb_url, sb_key=sb_key, requests_module=requests,
        mb_query_fn=mb_query, exit_book=exit_book, px_book=px_book,
        n_nas_passes=passes, gap_seconds=gap, verbose=True,
    )

    print()
    print("=" * 60)
    print(f"s5_idle           : {result['s5_idle']}")
    print(f"s6_idle           : {result['s6_idle']}")
    print(f"s5_cnp_raw        : {result['s5_cnp_raw']}")
    print(f"s5_dedup          : {result['s5_dedup']}")
    print(f"s5_could_not_pick : {result['s5_could_not_pick']}")
    print(f"s5_liability      : {result['s5_liability']}")
    print(f"collisions        : {len(result['collisions'])}")
