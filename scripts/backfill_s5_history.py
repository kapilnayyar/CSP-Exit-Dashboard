"""One-shot backfill for corrupted historical s5_dedup / s5_could_not_pick /
s5_liability rows in Daily Totals.

Why we need this
----------------
Between 2026-06-25 and 2026-07-02, the sanity floor v3 propagated a
phantom-low `s5_dedup = 7` across every day. cnp / liability stored in each
row was consequently ~14 too high. Tab 5's Delta view will spike +14 (once
corrected) then -14 (next day) unless we back-fill.

What this does
--------------
1. Runs the new compute_s5_snapshot() ONCE (5-pass NAS union) to get today's
   correct S5 numbers.
2. Overwrites the last N days of Daily Totals rows with a proxy value: the
   same cnp_raw / dedup / cnp / liability today's snapshot computed. Idle
   count and s5_collected are LEFT ALONE because those are stable per-day
   sheet-driven values that are already correct.
3. Prints "back-filled N historical rows".

Rationale for using today's number as a proxy
---------------------------------------------
We can't recompute historical NAS-IDLE state — Metabase T_DEVICE doesn't
retain history. Using today's live NAS union is a small directional error
compared to leaving visible +14/-14 delta spikes in Tab 5. Since actual
day-over-day cnp movement is usually < 5, back-fill error is comparable.

Usage
-----
Run once, then delete after 3 stable days:

    python -m scripts.backfill_s5_history --days 14

Environment variables required (same as capture_daily_snapshot.py):
    SUPABASE_URL, SUPABASE_ANON_KEY, METABASE_URL, METABASE_API_KEY,
    GOOGLE_SHEET_ID, plus google_credentials.json alongside .env.
"""
import argparse
import json
import os
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)
from s5_reconciliation import compute_s5_snapshot, PX_MIGRATION_WORKBOOK_ID  # noqa: E402

IST = ZoneInfo("Asia/Kolkata")
TOTALS_TAB = "Daily Totals"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--days", type=int, default=14,
                        help="Number of past days to backfill (default 14)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Compute and show, but do not write to sheet")
    args = parser.parse_args()

    # Load env (same pattern as snapshot script)
    from dotenv import load_dotenv
    env_path = os.path.join(_PARENT, ".env")
    if os.path.exists(env_path):
        load_dotenv(env_path)

    sb_url = os.getenv("SUPABASE_URL")
    sb_key = os.getenv("SUPABASE_ANON_KEY")
    mb_url = os.getenv("METABASE_URL")
    mb_key = os.getenv("METABASE_API_KEY")
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    gcp_json = os.getenv("GCP_CREDS_JSON")

    if gcp_json:
        gcp_creds = json.loads(gcp_json)
        creds = Credentials.from_service_account_info(
            gcp_creds,
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive.readonly"])
    else:
        cred_file = os.path.join(_PARENT, "google_credentials.json")
        creds = Credentials.from_service_account_file(
            cred_file,
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive.readonly"])

    gs = gspread.authorize(creds)
    book = gs.open_by_key(sheet_id)
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

    print(f"Computing today's correct S5 snapshot (5-pass NAS union, ~5 min)...")
    _s5 = compute_s5_snapshot(
        sb_url=sb_url, sb_key=sb_key, requests_module=requests,
        mb_query_fn=mb_query, exit_book=book, px_book=px_book,
        n_nas_passes=5, gap_seconds=60, verbose=True,
    )
    proxy_cnp_raw = _s5["s5_cnp_raw"]
    proxy_dedup = _s5["s5_dedup"]
    proxy_cnp = _s5["s5_could_not_pick"]

    # Read Daily Totals
    ws = book.worksheet(TOTALS_TAB)
    vals = ws.get_all_values()
    hdr = vals[0]
    idx = {name: i for i, name in enumerate(hdr)}
    if "s5_cnp_raw" not in idx or "s5_dedup" not in idx:
        print("Daily Totals missing s5_cnp_raw or s5_dedup columns — cron will add them next run.")
        return

    # Determine target rows (last N days)
    today = datetime.now(IST).date()
    targets = set()
    for i in range(args.days + 1):
        targets.add((today - timedelta(days=i)).strftime("%Y-%m-%d"))

    backfilled = 0
    for row_num, row in enumerate(vals[1:], start=2):
        if not row or not row[0]:
            continue
        row_date = row[0].strip()
        if row_date not in targets:
            continue
        # This row is in scope. Compute new liability = its own idle + proxy_cnp.
        try:
            row_idle = int(row[idx.get("s5_idle")] or 0) if "s5_idle" in idx else 0
        except (ValueError, IndexError):
            row_idle = 0
        new_liability = row_idle + proxy_cnp

        # Get column letters
        def _col_letter(i):
            return chr(ord("A") + i) if i < 26 else "A" + chr(ord("A") + i - 26)

        cells_to_update = [
            (idx["s5_cnp_raw"], proxy_cnp_raw),
            (idx["s5_dedup"], proxy_dedup),
            (idx["s5_could_not_pick"], proxy_cnp),
            (idx["s5_liability"], new_liability),
        ]
        if args.dry_run:
            print(f"  DRY: row {row_num} ({row_date}): "
                  f"idle={row_idle} -> cnp_raw={proxy_cnp_raw}, dedup={proxy_dedup}, "
                  f"cnp={proxy_cnp}, liab={new_liability}")
        else:
            for col_idx, val in cells_to_update:
                ws.update_cell(row_num, col_idx + 1, val)
            print(f"  UPDATED row {row_num} ({row_date}): "
                  f"cnp_raw={proxy_cnp_raw}, dedup={proxy_dedup}, "
                  f"cnp={proxy_cnp}, liab={new_liability}")
        backfilled += 1

    action = "would have back-filled" if args.dry_run else "back-filled"
    print(f"\n{action} {backfilled} historical rows.")


if __name__ == "__main__":
    main()
