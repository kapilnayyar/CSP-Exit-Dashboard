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

TOTALS_HEADERS = [
    "date",
    "s1_csps", "s1_userbase", "s1_voluntary", "s1_b1", "s1_b2",
    "s2_csps", "s2_userbase",
    "s3_csps", "s3_userbase",
    "s4a_csps", "s4a_u1_total", "s4a_u1_mig", "s4a_u2_total", "s4a_u2_pick",
    "s4b_csps", "s4b_u1", "s4b_u1_mig", "s4b_u2", "s4b_u2_pick", "s4b_pending",
    "s5_csps", "s5_idle", "s5_could_not_pick", "s5_liability", "s5_collected",
    "s6_csps", "s6_idle", "s6_collected",
]

# Same as dashboard.py — keep in sync
EXCLUDED_PARTNER_CODES = {
    281749854653857,
    281749854733209,
    274877909399,
}

SHEET_NAME_ALIAS = {
    "network solutions": "Manisha Traders 1",
    "manisha traders 2": "Manisha Traders 1",
    "khan enterprises": "Manisha Traders 1",
}


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
    u2_rows = book.worksheet(U2_TAB).get_all_records()
    u1_rows = book.worksheet(U1_TAB).get_all_records()

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
        remark = str(row.get("Remarks Dropdown") or "").strip().lower()
        if remark == "device picked up":
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
    idle_total = 0
    idle_total_s6 = 0
    idle_nas_by_code = defaultdict(set)
    if s5_codes:
        in_s5 = _metabase_id_in_list(s5_codes)
        rows = mb_run(f"""SELECT COUNT(*) FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_s5})""")
        idle_total = int(rows[0][0]) if rows else 0

        rows = mb_run(f"""SELECT td."LCO_ACCOUNT_ID", td."NASID"
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_s5})
  AND td."NASID" IS NOT NULL""")
        for code, nas in rows:
            if nas:
                idle_nas_by_code[str(code)].add(str(nas))
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

    all_pending_mobiles = []
    for key in s5_partner_keys:
        all_pending_mobiles.extend(u2_pending_by_name.get(key, []))
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
    for key in s5_partner_keys:
        partner_code = s5_name_to_code.get(key)
        idle_nas = idle_nas_by_code.get(partner_code, set()) if partner_code else set()
        for mobile in u2_pending_by_name.get(key, []):
            nas = mobile_to_nas.get(mobile)
            if nas and nas in idle_nas:
                s5_dup += 1
    print(f"S5 dedup count: {s5_dup}")

    # ── 10. Netbox Collection (S5 Netbox Collection tab) ───────────────────
    netbox_collected_by_code = defaultdict(int)
    try:
        nc_rows = book.worksheet(NETBOX_COLLECTION_TAB).get_all_records()
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
    s4a_u1_total = sum(u1_by.get(key_of(p), {}).get("total", 0) for p in completed)
    s4a_u1_mig = sum(u1_by.get(key_of(p), {}).get("migrated", 0) for p in completed)
    s4a_u2_total = sum(u2_total.get(key_of(p), 0) for p in completed)
    s4a_u2_pick = sum(u2_picked.get(key_of(p), 0) for p in completed)

    s4b_csps = len(s4_partners)
    s4b_u1 = s4b_u1_mig = s4b_u2 = s4b_u2_pick = s4b_pending = 0
    for p in s4_partners:
        k = key_of(p)
        u1d = u1_by.get(k, {"total": 0, "migrated": 0})
        s4b_u1 += u1d["total"]
        s4b_u1_mig += u1d["migrated"]
        s4b_u2 += u2_total.get(k, 0)
        s4b_u2_pick += u2_picked.get(k, 0)
        if u1d["total"] == 0 and u2_total.get(k, 0) == 0:
            s4b_pending += r15_of(p)

    # S5 reconciliation
    s5_u1_total = s5_u1_mig = s5_u2_total = s5_u2_picked = 0
    for p in s5_partners:
        k = key_of(p)
        u1d = u1_by.get(k, {"total": 0, "migrated": 0})
        s5_u1_total += u1d["total"]; s5_u1_mig += u1d["migrated"]
        s5_u2_total += u2_total.get(k, 0); s5_u2_picked += u2_picked.get(k, 0)
    raw_cnp = (s5_u1_total - s5_u1_mig) + (s5_u2_total - s5_u2_picked)
    s5_could_not_pick = max(raw_cnp - s5_dup, 0)
    s5_liability = idle_total + s5_could_not_pick
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
    }

    # ── 12. Append row ──────────────────────────────────────────────────────
    row = [today_str] + [int(totals.get(k, 0) or 0) for k in TOTALS_HEADERS[1:]]
    ws_totals.append_row(row, value_input_option="USER_ENTERED")
    print(f"Wrote row for {today_str}: {row}")


if __name__ == "__main__":
    main()
