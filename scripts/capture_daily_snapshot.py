"""Standalone daily snapshot capture — runs from GitHub Actions at 00:00 IST.

Mirrors dashboard.py's compute_today_metrics + write_today_totals logic so the
'Daily Totals' Google Sheet tab gets exactly one new row per day at midnight,
regardless of whether anyone has the Streamlit app open.

Inputs (all via env vars set by the GH Actions workflow):
  SUPABASE_URL, SUPABASE_ANON_KEY
  METABASE_URL, METABASE_API_KEY
  GOOGLE_SHEET_ID
  GCP_CREDS_JSON   (full google_credentials.json contents as a single string)
"""

import json
import os
import re
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials


IST = ZoneInfo("Asia/Kolkata")

U2_TAB = "Main sheet"
U1_TAB = "Migration Data"
NETBOX_COLLECTION_TAB = "S5 Netbox Collection"
TOTALS_TAB = "Daily Totals"

# PX/CX Migration Summary workbook — row-level U1 tabs for U1 dedup.
# Added 2026-07-02 to mirror dashboard.py; keeps snapshot in sync with the
# extended dedup that now covers both U1 and U2 pending customers.
PX_MIGRATION_WORKBOOK_ID = "1hmT50leXZUAibzd2zzfO4FVj-B3m675CCFUbdwFuVS4"
PX_RAW_TAB = "PX Migration Raw Data"
PX_MIGRATED_TAB = "PX Migrated Cases"

TOTALS_HEADERS = [
    "date",
    "s1_csps", "s1_userbase", "s1_voluntary", "s1_b1", "s1_b2",
    "s2_csps", "s2_userbase",
    "s3_csps", "s3_userbase",
    "s4a_csps", "s4a_u1_total", "s4a_u1_mig", "s4a_u2_total", "s4a_u2_pick",
    "s4b_csps", "s4b_u1", "s4b_u1_mig", "s4b_u2", "s4b_u2_pick", "s4b_pending",
    "s5_csps", "s5_idle", "s5_could_not_pick", "s5_liability", "s5_collected",
    "s6_csps", "s6_idle", "s6_collected",
    # Added 2026-06-29 to power the dedup-based sanity floor:
    "s5_cnp_raw", "s5_dedup",
]

# Same as dashboard.py — keep in sync
EXCLUDED_PARTNER_CODES = {
    281749854653857,
    281749854733209,
    274877909399,
    281749854790153,  # exit stopped
    281749854637042,  # exit stopped
}

SHEET_NAME_ALIAS = {
    "network solutions": "Manisha Traders 1",
    "manisha traders 2": "Manisha Traders 1",
    "khan enterprises": "Manisha Traders 1",
}

# Partner-code attribution overrides — used when two Supabase partners share
# the same sheet name. Each entry pins the per-partner sheet values directly,
# bypassing the name-based lookup that would otherwise double-count.
# Keep in sync with dashboard.py ATTRIBUTION_OVERRIDE.
ATTRIBUTION_OVERRIDE = {
    # Mirror of dashboard.py ATTRIBUTION_OVERRIDE — must stay in sync.
    # "auto" sentinel = use the live sheet aggregate for that field.
    # Pattern: one collision-owner uses "auto" so it auto-syncs; the others
    # are forced to 0 to avoid double-counting.
    "274877952814":    {"u1_total": 8, "u1_migrated": 8,
                        "u1_not_migrated": 0, "u1_wip": 0,
                        "u2_total": "auto", "u2_picked": "auto"},
    "281749855023736": {"u1_total": 2, "u1_migrated": 0,
                        "u1_not_migrated": 2, "u1_wip": 0,
                        "u2_total": 0, "u2_picked": 0},
    "274877953157":    {"u1_total": "auto", "u1_migrated": "auto",
                        "u1_not_migrated": "auto", "u1_wip": "auto",
                        "u2_total": "auto", "u2_picked": "auto"},
    "281749854772211": {"u1_total": 0, "u1_migrated": 0,
                        "u1_not_migrated": 0, "u1_wip": 0,
                        "u2_total": 0, "u2_picked": 0},
    "281749854778714": {"u1_total": "auto", "u1_migrated": "auto",
                        "u1_not_migrated": "auto", "u1_wip": "auto",
                        "u2_total": "auto", "u2_picked": "auto"},
    "281749854868832": {"u1_total": 0, "u1_migrated": 0,
                        "u1_not_migrated": 0, "u1_wip": 0,
                        "u2_total": 0, "u2_picked": 0},
}


def _u1_for(p, u1_by):
    """U1 metrics: 'auto' override fields fall back to live sheet aggregate."""
    code = str(p.get("partner_code") or "")
    key = p["name"].lower()
    key = SHEET_NAME_ALIAS.get(key, p["name"]).lower() if key in SHEET_NAME_ALIAS else key
    sheet_val = u1_by.get(key, {"total": 0, "migrated": 0})
    if code not in ATTRIBUTION_OVERRIDE:
        return sheet_val
    ov = ATTRIBUTION_OVERRIDE[code]
    def pick(field, ov_field):
        v = ov.get(ov_field, 0)
        return sheet_val.get(field, 0) if v == "auto" else v
    return {"total": pick("total", "u1_total"),
            "migrated": pick("migrated", "u1_migrated")}


def _u2_total_for(p, u2_total):
    code = str(p.get("partner_code") or "")
    key = p["name"].lower()
    key = SHEET_NAME_ALIAS.get(key, p["name"]).lower() if key in SHEET_NAME_ALIAS else key
    if code in ATTRIBUTION_OVERRIDE:
        v = ATTRIBUTION_OVERRIDE[code]["u2_total"]
        if v == "auto":
            return u2_total.get(key, 0)
        return v
    return u2_total.get(key, 0)


def _u2_picked_for(p, u2_picked):
    code = str(p.get("partner_code") or "")
    key = p["name"].lower()
    key = SHEET_NAME_ALIAS.get(key, p["name"]).lower() if key in SHEET_NAME_ALIAS else key
    if code in ATTRIBUTION_OVERRIDE:
        v = ATTRIBUTION_OVERRIDE[code]["u2_picked"]
        if v == "auto":
            return u2_picked.get(key, 0)
        return v
    return u2_picked.get(key, 0)


def _read_records_safe(ws):
    """Read a worksheet to list-of-dicts without tripping gspread's strict
    header validation. Trims trailing empty header columns so empty extras
    don't get flagged as 'duplicate header'. Mirror of dashboard.py helper."""
    all_values = ws.get_all_values()
    if not all_values:
        return []
    header_row = all_values[0]
    last_filled = 0
    for i, h in enumerate(header_row):
        if h and h.strip():
            last_filled = i + 1
    headers = [h.strip() for h in header_row[:last_filled]]
    out = []
    for row in all_values[1:]:
        padded = row[:last_filled] + [""] * max(0, last_filled - len(row))
        out.append(dict(zip(headers, padded)))
    return out


def _to_int(v):
    try:
        return int(str(v).replace(",", "").strip() or 0)
    except Exception:
        return 0


def _metabase_id_in_list(codes):
    return ",".join(str(int(c)) for c in codes if c)


def _metabase_safe_in_list(values):
    safe = []
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        safe.append("'" + s.replace("'", "''") + "'")
    return ",".join(safe)


def main():
    # ── 1. Validate env ─────────────────────────────────────────────────────
    required = ["SUPABASE_URL", "SUPABASE_ANON_KEY",
                "METABASE_URL", "METABASE_API_KEY",
                "GOOGLE_SHEET_ID", "GCP_CREDS_JSON"]
    missing = [k for k in required if not os.getenv(k)]
    if missing:
        print(f"ERROR: missing/empty env vars: {missing}")
        print("These should be set as GitHub Actions secrets. Check that each "
              "secret has a non-empty value at "
              "github.com/<owner>/<repo>/settings/secrets/actions")
        sys.exit(1)
    # Extra check: GCP_CREDS_JSON must parse as JSON, not just be non-empty
    raw_gcp = os.getenv("GCP_CREDS_JSON", "").strip()
    if not raw_gcp.startswith("{"):
        print("ERROR: GCP_CREDS_JSON does not look like JSON "
              f"(starts with: {raw_gcp[:30]!r}). "
              "Re-paste the contents of google_credentials.json into the secret.")
        sys.exit(1)

    supabase_url = os.getenv("SUPABASE_URL").rstrip("/")
    supabase_key = os.getenv("SUPABASE_ANON_KEY")
    mb_url = os.getenv("METABASE_URL").rstrip("/")
    mb_key = os.getenv("METABASE_API_KEY")
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    try:
        gcp_creds = json.loads(raw_gcp)
    except json.JSONDecodeError as e:
        print(f"ERROR: GCP_CREDS_JSON is set but not valid JSON: {e}")
        print(f"First 200 chars of value: {raw_gcp[:200]!r}")
        sys.exit(1)

    # Late-night capture pattern: the cron runs at 23:55 IST, just before
    # midnight. The snapshot represents "end of TODAY", which is the D-1
    # baseline for TOMORROW's dashboard view. So we label the row with
    # tomorrow's date so dashboard's "today's row" lookup finds it.
    now_ist = datetime.now(IST)
    if now_ist.hour >= 23:
        target_date = (now_ist + timedelta(days=1)).strftime("%Y-%m-%d")
        print(f"Late-night capture at {now_ist.strftime('%H:%M IST')} - "
              f"labeling row as next day: {target_date}")
    else:
        target_date = now_ist.strftime("%Y-%m-%d")
        print(f"Capture at {now_ist.strftime('%H:%M IST')} - "
              f"labeling row as today: {target_date}")
    today_str = target_date  # rest of script uses today_str as the row date

    # ── 2. Google Sheet client + idempotency check ─────────────────────────
    creds = Credentials.from_service_account_info(
        gcp_creds,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.readonly"],
    )
    book = gspread.authorize(creds).open_by_key(sheet_id)

    try:
        ws_totals = book.worksheet(TOTALS_TAB)
    except gspread.WorksheetNotFound:
        ws_totals = book.add_worksheet(
            title=TOTALS_TAB, rows=2000, cols=len(TOTALS_HEADERS))
        ws_totals.append_row(TOTALS_HEADERS)
        print(f"Created '{TOTALS_TAB}' tab.")

    existing_dates = set(ws_totals.col_values(1)[1:])
    if today_str in existing_dates:
        print(f"Row for {today_str} already exists — nothing to do.")
        return

    # ── 3. Supabase: partners ──────────────────────────────────────────────
    sb_hdr = {"apikey": supabase_key,
              "Authorization": f"Bearer {supabase_key}",
              "Accept": "application/json"}
    r = requests.get(
        f"{supabase_url}/rest/v1/partners",
        params={"select": "id,name,partner_code,current_state,exit_type"},
        headers=sb_hdr, timeout=30,
    )
    r.raise_for_status()
    partners = [p for p in r.json()
                if int(p.get("partner_code") or 0) not in EXCLUDED_PARTNER_CODES]
    print(f"Partners fetched: {len(partners)}")

    # ── 4. Google Sheet: U1 (Migration Data) + U2 (Main sheet) ─────────────
    # _read_records_safe tolerates trailing empty columns (otherwise gspread's
    # get_all_records() raises "duplicate header: ['']" when col_count > filled headers).
    u2_rows = _read_records_safe(book.worksheet(U2_TAB))
    u1_rows = _read_records_safe(book.worksheet(U1_TAB))

    u1_by = {}
    for row in u1_rows:
        name = str(row.get("Exit Partner Name") or "").strip()
        if not name:
            continue
        key = name.lower()
        key = SHEET_NAME_ALIAS.get(key, name).lower() if key in SHEET_NAME_ALIAS else key
        ex = u1_by.get(key, {"total": 0, "migrated": 0})
        u1_by[key] = {
            "total": ex["total"] + _to_int(row.get("Total U1 User")),
            "migrated": ex["migrated"] + _to_int(row.get("Migrated")),
        }

    u2_total = defaultdict(int)
    u2_picked = defaultdict(int)
    u2_pending_by_name = defaultdict(list)
    for row in u2_rows:
        name = str(row.get("Partner") or "").strip()
        mobile = row.get("Mobile no") or row.get("Mobile")
        if not name or not mobile:
            continue
        key = name.lower()
        key = SHEET_NAME_ALIAS.get(key, name).lower() if key in SHEET_NAME_ALIAS else key
        u2_total[key] += 1
        # PICKED RULE: customer counts as picked if EITHER column says so —
        # Remarks Dropdown == 'Device picked up'  OR
        # Device Picked (Ajinkya/Pradeep) == 'Yes' (both case-insensitive).
        # Mirrors dashboard.py's _u2_row_is_picked helper. Keep in sync.
        remark = str(row.get("Remarks Dropdown") or "").strip().lower()
        pradeep = str(row.get("Device Picked (Ajinkya/Pradeep)") or "").strip().lower()
        if remark == "device picked up" or pradeep == "yes":
            u2_picked[key] += 1
        else:
            u2_pending_by_name[key].append(str(mobile).strip())

    # ── 5. Metabase helper ─────────────────────────────────────────────────
    mb_hdr = {"x-api-key": mb_key, "Content-Type": "application/json"}

    def mb_run(sql):
        res = requests.post(
            f"{mb_url}/api/dataset",
            json={"database": 113, "type": "native",
                  "native": {"query": sql, "template-tags": {}}},
            headers=mb_hdr, timeout=120,
        )
        if res.status_code not in (200, 202):
            print(f"Metabase ERR {res.status_code}: {res.text[:300]}")
            return []
        return res.json().get("data", {}).get("rows", [])

    # ── 6. Per-stage partner sets ──────────────────────────────────────────
    by_state = defaultdict(list)
    for p in partners:
        by_state[p["current_state"]].append(p)

    s5_partners = by_state.get("S5", [])
    s6_partners = by_state.get("S6", [])
    s5_codes = [str(p["partner_code"]) for p in s5_partners]
    s6_codes = [str(p["partner_code"]) for p in s6_partners]
    all_codes = [str(p["partner_code"]) for p in partners]

    # ── 7. Idle devices: total for S5, total for S6, NAS sets for S5 ───────
    # IMPORTANT — STABILIZE THE NAS SET (V2):
    # The "IDLE" status in inventory flickers per NAS for windows >1 min
    # (V1 used 3 measurements over 60s and STILL caught a +7 phantom drop on
    # 21-Jun and 22-Jun). V2: 5 measurements spread over ~4 minutes; UNION
    # all results. Any NAS that appeared IDLE in ANY pass stays in the dedup
    # pool. Sanity floor: if union total is more than 10% below yesterday's
    # union total, fall back to yesterday's NAS set (read from Daily Totals).
    idle_total = 0
    idle_total_s6 = 0
    idle_nas_by_code = defaultdict(set)
    idle_count_by_code = {}   # per-CSP IDLE counts — written to "S5 Daily IDLE by CSP" tab
    if s5_codes:
        in_s5 = _metabase_id_in_list(s5_codes)
        rows = mb_run(f"""SELECT COUNT(*) FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_s5})""")
        idle_total = int(rows[0][0]) if rows else 0

        # Per-CSP IDLE count (one extra cheap query, written to its own tab below
        # so we can diff CSP-level day-over-day and explain aggregate IDLE swings)
        rows = mb_run(f"""SELECT td."LCO_ACCOUNT_ID", COUNT(*)
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_s5})
GROUP BY 1""")
        for code, n in rows:
            idle_count_by_code[str(code)] = int(n)

        N_MEASUREMENTS = 5
        GAP_SECONDS = 60
        measurement_sizes = []
        for attempt in range(N_MEASUREMENTS):
            if attempt > 0:
                time.sleep(GAP_SECONDS)
            rows = mb_run(f"""SELECT td."LCO_ACCOUNT_ID", td."NASID"
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_s5})
  AND td."NASID" IS NOT NULL""")
            this_pass = 0
            for code, nas in rows:
                if nas:
                    idle_nas_by_code[str(code)].add(str(nas))
                    this_pass += 1
            measurement_sizes.append(this_pass)
        total_nas = sum(len(v) for v in idle_nas_by_code.values())
        print(f"Idle NAS measurements ({N_MEASUREMENTS} passes, {GAP_SECONDS}s gap): "
              f"{measurement_sizes}  union total: {total_nas}")
    if s6_codes:
        in_s6 = _metabase_id_in_list(s6_codes)
        rows = mb_run(f"""SELECT COUNT(*) FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_s6})""")
        idle_total_s6 = int(rows[0][0]) if rows else 0
    print(f"Idle (S5): {idle_total}  Idle (S6): {idle_total_s6}")

    # ── 8. R15 active by code (for S4 pending fallback) ────────────────────
    r15_by_code = {}
    if all_codes:
        in_all = _metabase_id_in_list(all_codes)
        sql = f"""WITH partner_nas AS (
  SELECT DISTINCT NASID, MOBILE, LCO_ACCOUNT_ID
  FROM PUBLIC.T_WG_CUSTOMER
  WHERE LCO_ACCOUNT_ID IN ({in_all})
),
live_latest AS (
  SELECT t.ROUTER_NAS_ID, MAX(t.OTP_EXPIRY_TIME) AS LATEST_EXPIRY
  FROM PUBLIC.T_ROUTER_USER_MAPPING t
  WHERE t.AUTH_STATE = 1
    AND t.OTP NOT IN ('FREE','PAY_ONLINE','CASH','ROAM')
    AND t.MOBILE > '5999999999' AND t.DEVICE_LIMIT = 10
  GROUP BY t.ROUTER_NAS_ID
),
per_mobile AS (
  SELECT pn.LCO_ACCOUNT_ID, pn.MOBILE,
         MAX(l.LATEST_EXPIRY) AS LATEST_EXPIRY
  FROM partner_nas pn
  LEFT JOIN live_latest l ON pn.NASID = l.ROUTER_NAS_ID
  GROUP BY pn.LCO_ACCOUNT_ID, pn.MOBILE
)
SELECT LCO_ACCOUNT_ID,
       COUNT_IF(LATEST_EXPIRY >= DATEADD(day, -15, CURRENT_DATE)) AS R15
FROM per_mobile
GROUP BY 1"""
        for row in mb_run(sql):
            r15_by_code[str(row[0])] = int(row[1] or 0)
    print(f"R15 active fetched for {len(r15_by_code)} partners")

    # ── 9. S5 dedup — pending U2 customer's NASID already in IDLE set ─────
    s5_partner_keys = set()
    s5_name_to_code = {}
    for p in s5_partners:
        n = p["name"].lower()
        n = SHEET_NAME_ALIAS.get(n, p["name"]).lower() if n in SHEET_NAME_ALIAS else n
        s5_partner_keys.add(n)
        s5_name_to_code[n] = str(p["partner_code"])

    # ── U1 pending mobiles from PX/CX Migration Summary workbook ───────────
    # Added 2026-07-02: apply the same NAS-vs-IDLE dedup we do for U2 to U1
    # customers as well. Pending U1 = every raw-data row that is NOT in the
    # migrated-cases tab, per exit_partner_code.
    u1_pending_mobiles_by_code = defaultdict(set)
    try:
        px_book = client.open_by_key(PX_MIGRATION_WORKBOOK_ID)
        # Raw Data - all U1 customers of exiting CSPs, keyed by exit_partner_code
        u1_raw_by_code = defaultdict(set)
        raw = px_book.worksheet(PX_RAW_TAB).get_all_values()
        if raw:
            h = raw[0]
            try:
                c_code = h.index("exit_partner_code")
                c_mob = h.index("Customer_mobile")
            except ValueError:
                c_code = c_mob = None
            if c_code is not None and c_mob is not None:
                for r in raw[1:]:
                    if len(r) <= max(c_code, c_mob):
                        continue
                    code = str(r[c_code]).strip()
                    mobile = str(r[c_mob]).strip()
                    if code and mobile:
                        u1_raw_by_code[code].add(mobile)
        # Migrated Cases - keyed by Old Partner Name -> resolve to code
        u1_mig_by_code = defaultdict(set)
        mig = px_book.worksheet(PX_MIGRATED_TAB).get_all_values()
        if mig:
            h = mig[0]
            c_mob_m = c_name_m = None
            for i, hh in enumerate(h):
                hl = str(hh).strip().lower()
                if c_mob_m is None and "mobile" in hl:
                    c_mob_m = i
                if c_name_m is None and ("old partner" in hl or "exit partner" in hl):
                    c_name_m = i
            if c_mob_m is not None and c_name_m is not None:
                for r in mig[1:]:
                    if len(r) <= max(c_mob_m, c_name_m):
                        continue
                    nm = str(r[c_name_m]).strip().lower()
                    if nm in SHEET_NAME_ALIAS:
                        nm = SHEET_NAME_ALIAS[nm].lower()
                    mobile = str(r[c_mob_m]).strip()
                    if nm and mobile:
                        code = s5_name_to_code.get(nm)
                        if code:
                            u1_mig_by_code[code].add(mobile)
        # Pending = raw - migrated per code, only for S5 partners
        for code, raw_set in u1_raw_by_code.items():
            if code in {s5_name_to_code[k] for k in s5_partner_keys}:
                u1_pending_mobiles_by_code[code] = raw_set - u1_mig_by_code.get(code, set())
    except Exception as e:
        print(f"PX workbook read failed (U1 dedup disabled): {e}")

    # Pool ALL pending mobiles (U2 by partner name + U1 by partner code)
    all_pending_mobiles = []
    for key in s5_partner_keys:
        all_pending_mobiles.extend(u2_pending_by_name.get(key, []))
    for code, mobs in u1_pending_mobiles_by_code.items():
        all_pending_mobiles.extend(mobs)
    all_pending_mobiles = list(set(all_pending_mobiles))

    mobile_to_nas = {}
    if all_pending_mobiles:
        m_in = _metabase_safe_in_list(all_pending_mobiles)
        if m_in:
            rows = mb_run(f"""SELECT MOBILE, NASID FROM PUBLIC.T_WG_CUSTOMER
WHERE MOBILE IN ({m_in})""")
            for mobile, nas in rows:
                if mobile and nas:
                    mobile_to_nas[str(mobile)] = str(nas)

    s5_dup = 0
    # U2 dedup pass (existing behavior)
    for key in s5_partner_keys:
        partner_code = s5_name_to_code.get(key)
        idle_nas = idle_nas_by_code.get(partner_code, set()) if partner_code else set()
        for mobile in u2_pending_by_name.get(key, []):
            nas = mobile_to_nas.get(mobile)
            if nas and nas in idle_nas:
                s5_dup += 1
    # U1 dedup pass (new)
    for code, mobs in u1_pending_mobiles_by_code.items():
        idle_nas = idle_nas_by_code.get(code, set())
        for mobile in mobs:
            nas = mobile_to_nas.get(mobile)
            if nas and nas in idle_nas:
                s5_dup += 1
    print(f"S5 dedup count (U1 + U2): {s5_dup}")

    # ── 10. Netbox Collection (S5 Netbox Collection tab) ───────────────────
    netbox_collected_by_code = defaultdict(int)
    try:
        nc_rows = _read_records_safe(book.worksheet(NETBOX_COLLECTION_TAB))
        for row in nc_rows:
            code = str(row.get("CSP ID") or "").strip()
            if not code:
                continue
            netbox_collected_by_code[code] += _to_int(row.get("Devices collected from CSP"))
    except gspread.WorksheetNotFound:
        pass

    # ── 11. compute_today_metrics (same formulas as dashboard.py) ──────────
    def r15_of(p):
        return r15_by_code.get(str(p.get("partner_code") or ""), 0)

    def sheet_ub(p):
        key = p["name"].lower()
        key = SHEET_NAME_ALIAS.get(key, p["name"]).lower() if key in SHEET_NAME_ALIAS else key
        return (u1_by.get(key, {}).get("total", 0) or 0) + (u2_total.get(key, 0) or 0)

    def ub_of(p):
        s = sheet_ub(p)
        return s if s > 0 else r15_of(p)

    in_pipeline = [p for p in partners if p.get("current_state") in
                   ("S1", "S2", "S3", "S4", "S5", "S6")]
    current_s2 = by_state.get("S2", [])
    past_s3 = [p for p in in_pipeline if p["current_state"] in
               ("S3", "S4", "S5", "S6")]
    s4_partners = by_state.get("S4", [])
    completed = s5_partners + s6_partners

    s1_csps = len(in_pipeline)
    s1_userbase = sum(ub_of(p) for p in in_pipeline)
    s1_voluntary = sum(1 for p in in_pipeline
                       if str(p.get("exit_type") or "").strip() == "Voluntary")
    s1_b1 = sum(1 for p in in_pipeline
                if str(p.get("exit_type") or "").strip() == "B1")
    s1_b2 = sum(1 for p in in_pipeline
                if str(p.get("exit_type") or "").strip() == "B2")

    s2_csps = len(current_s2)
    s2_userbase = sum(ub_of(p) for p in current_s2)
    s3_csps = len(past_s3)
    s3_userbase = sum(ub_of(p) for p in past_s3)

    def key_of(p):
        n = p["name"].lower()
        return SHEET_NAME_ALIAS.get(n, p["name"]).lower() if n in SHEET_NAME_ALIAS else n

    s4a_csps = len(completed)
    s4a_u1_total = sum(_u1_for(p, u1_by)["total"] for p in completed)
    s4a_u1_mig = sum(_u1_for(p, u1_by)["migrated"] for p in completed)
    s4a_u2_total = sum(_u2_total_for(p, u2_total) for p in completed)
    s4a_u2_pick = sum(_u2_picked_for(p, u2_picked) for p in completed)

    s4b_csps = len(s4_partners)
    s4b_u1 = s4b_u1_mig = s4b_u2 = s4b_u2_pick = s4b_pending = 0
    for p in s4_partners:
        u1d = _u1_for(p, u1_by)
        u2t = _u2_total_for(p, u2_total)
        u2p = _u2_picked_for(p, u2_picked)
        s4b_u1 += u1d["total"]
        s4b_u1_mig += u1d["migrated"]
        s4b_u2 += u2t
        s4b_u2_pick += u2p
        if u1d["total"] == 0 and u2t == 0:
            s4b_pending += r15_of(p)

    # S5 reconciliation
    s5_u1_total = s5_u1_mig = s5_u2_total = s5_u2_picked = 0
    for p in s5_partners:
        u1d = _u1_for(p, u1_by)
        s5_u1_total += u1d["total"]; s5_u1_mig += u1d["migrated"]
        s5_u2_total += _u2_total_for(p, u2_total)
        s5_u2_picked += _u2_picked_for(p, u2_picked)
    raw_cnp = (s5_u1_total - s5_u1_mig) + (s5_u2_total - s5_u2_picked)
    s5_could_not_pick = max(raw_cnp - s5_dup, 0)
    s5_liability = idle_total + s5_could_not_pick

    # ── SANITY FLOOR v3 (2026-06-29) — dedup-based ───────────────────────
    # v2 floor required idle to be stable (delta <=5). That guard failed on
    # 28-Jun when idle legitimately dropped by 37 devices (CSPs returned
    # them) — the floor stepped back and the phantom +7 dedup drop slipped
    # through, requiring another manual correction.
    #
    # v3 fix: store dedup itself in Daily Totals, then check yesterday's
    # stored dedup directly. If today's measured dedup is meaningfully
    # lower than yesterday's stored dedup AND raw_cnp (sheet-driven, not
    # subject to flicker) hasn't moved much, substitute yesterday's dedup.
    # Completely independent of idle changes.
    try:
        target_dt = datetime.strptime(today_str, "%Y-%m-%d").date()
        yest_dt = (target_dt - timedelta(days=1)).strftime("%Y-%m-%d")
        prior_rows = _read_records_safe(ws_totals)
        prior = next((r for r in prior_rows if str(r.get("date")) == yest_dt), None)
        if prior:
            yest_dedup = int(float(prior.get("s5_dedup") or 0))
            yest_cnp_raw = int(float(prior.get("s5_cnp_raw") or 0))
            yest_cnp = int(float(prior.get("s5_could_not_pick") or 0))
            # Prefer dedup-based floor when yesterday's dedup is on file.
            if yest_dedup > 0 and yest_cnp_raw > 0:
                if (s5_dup < yest_dedup - 2
                        and abs(raw_cnp - yest_cnp_raw) <= 5):
                    print(f"SANITY FLOOR v3 triggered (dedup-based): "
                          f"measured dedup={s5_dup}, yesterday dedup={yest_dedup}, "
                          f"raw_cnp stable ({raw_cnp} vs {yest_cnp_raw}). "
                          f"Using yesterday's dedup to filter phantom flicker.")
                    s5_dup = yest_dedup
                    s5_could_not_pick = max(raw_cnp - s5_dup, 0)
                    s5_liability = idle_total + s5_could_not_pick
            # Fallback: if yesterday's dedup is not on file (first runs
            # after this deploy), use the old cnp-based check WITHOUT the
            # idle-stable guard so phantom on a real-idle-shift day still
            # gets caught.
            elif yest_cnp > 0 and s5_could_not_pick > yest_cnp + 5:
                print(f"SANITY FLOOR v3 triggered (fallback): "
                      f"measured cnp={s5_could_not_pick}, yesterday cnp={yest_cnp}. "
                      f"Using yesterday's cnp.")
                s5_could_not_pick = yest_cnp
                s5_liability = idle_total + s5_could_not_pick
    except Exception as e:
        print(f"Sanity floor check skipped (will not block snapshot): {e}")

    s5_collected = sum(
        netbox_collected_by_code.get(str(p.get("partner_code") or ""), 0)
        for p in s5_partners)

    s6_collected = sum(
        netbox_collected_by_code.get(str(p.get("partner_code") or ""), 0)
        for p in s6_partners)

    totals = {
        "s1_csps": s1_csps, "s1_userbase": s1_userbase,
        "s1_voluntary": s1_voluntary, "s1_b1": s1_b1, "s1_b2": s1_b2,
        "s2_csps": s2_csps, "s2_userbase": s2_userbase,
        "s3_csps": s3_csps, "s3_userbase": s3_userbase,
        "s4a_csps": s4a_csps,
        "s4a_u1_total": s4a_u1_total, "s4a_u1_mig": s4a_u1_mig,
        "s4a_u2_total": s4a_u2_total, "s4a_u2_pick": s4a_u2_pick,
        "s4b_csps": s4b_csps,
        "s4b_u1": s4b_u1, "s4b_u1_mig": s4b_u1_mig,
        "s4b_u2": s4b_u2, "s4b_u2_pick": s4b_u2_pick,
        "s4b_pending": s4b_pending,
        "s5_csps": len(s5_partners), "s5_idle": idle_total,
        "s5_could_not_pick": s5_could_not_pick, "s5_liability": s5_liability,
        "s5_collected": s5_collected,
        "s6_csps": len(s6_partners), "s6_idle": idle_total_s6,
        "s6_collected": s6_collected,
        # v3 sanity-floor inputs
        "s5_cnp_raw": raw_cnp,
        "s5_dedup": s5_dup,
    }

    # ── 12. Append row ──────────────────────────────────────────────────────
    # Make sure the sheet header row has every column we are about to write.
    # Auto-extend if TOTALS_HEADERS has grown since the sheet was created.
    current_header = ws_totals.row_values(1)
    if current_header != TOTALS_HEADERS:
        missing = [h for h in TOTALS_HEADERS if h not in current_header]
        if missing:
            print(f"Daily Totals header missing {missing} — extending sheet header.")
            # Build the new header preserving existing column positions, then
            # appending the missing ones to the end.
            new_header = list(current_header) + missing
            # Resize sheet if needed
            if ws_totals.col_count < len(new_header):
                ws_totals.add_cols(len(new_header) - ws_totals.col_count)
            ws_totals.update("A1", [new_header], value_input_option="USER_ENTERED")
            current_header = new_header

    # Build the row in the SHEET's column order (not TOTALS_HEADERS'),
    # so older columns stay where they were and new columns land at the end.
    row = [today_str if k == "date" else int(totals.get(k, 0) or 0)
           for k in current_header]
    ws_totals.append_row(row, value_input_option="USER_ENTERED")
    print(f"Wrote row for {today_str}: cnp_raw={raw_cnp}, dedup={s5_dup}, cnp={s5_could_not_pick}")

    # ── 13. Per-CSP IDLE snapshot ─────────────────────────────────────────
    # Writes one row per S5 partner per day to "S5 Daily IDLE by CSP" tab.
    # Lets us diff CSP-level day-over-day to explain aggregate IDLE swings
    # (e.g., the 37-device drop on 27-Jun where 21 turned out to be
    # LCO_ACCOUNT_ID reassignments to non-S5 partners).
    IDLE_TAB = "S5 Daily IDLE by CSP"
    IDLE_HEADERS = ["Date", "Partner Code", "Partner Name", "IDLE Count"]
    try:
        ws_idle = book.worksheet(IDLE_TAB)
    except gspread.WorksheetNotFound:
        ws_idle = book.add_worksheet(title=IDLE_TAB, rows=50000, cols=len(IDLE_HEADERS))
        ws_idle.append_row(IDLE_HEADERS)
        print(f"Created '{IDLE_TAB}' tab.")

    # Skip if today's per-CSP rows already exist (idempotent)
    existing_keys = {f"{r[0]}|{r[1]}" for r in ws_idle.get_all_values()[1:] if len(r) >= 2}
    code_to_name_local = {str(p["partner_code"]): p["name"] for p in s5_partners}
    new_rows = []
    for code, name in code_to_name_local.items():
        if f"{today_str}|{code}" in existing_keys:
            continue
        n = int(idle_count_by_code.get(code, 0))
        new_rows.append([today_str, code, name, n])
    if new_rows:
        ws_idle.append_rows(new_rows, value_input_option="USER_ENTERED")
        print(f"Wrote {len(new_rows)} per-CSP IDLE rows for {today_str}.")
    else:
        print(f"Per-CSP IDLE rows for {today_str} already present — skipped.")


if __name__ == "__main__":
    main()
