"""
pull_all_data.py — one-shot dump of ALL data behind the CSP Exit Dashboard.

Reads (view-only) every source the live dashboard uses and writes them to
a dated folder so we can build a new "leaders' tracker" on top of clean data.

Sources pulled:
  1. Google Sheet workbook (GOOGLE_SHEET_ID) — EVERY tab (Main sheet=U2,
     Migration Data=U1, S5 Netbox Collection, Daily Totals, + any others).
  2. PX/CX Migration Summary workbook — Raw Data + Migrated Cases tabs.
  3. Supabase `partners` table — exit lifecycle (state/risk/type) + u1/u2 counts.

Nothing is written back anywhere. Pure read.
"""

import os
import json
import csv
from datetime import datetime

import gspread
import requests
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(HERE, ".env"))

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
PX_MIGRATION_WORKBOOK_ID = "1hmT50leXZUAibzd2zzfO4FVj-B3m675CCFUbdwFuVS4"
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY")

# Output folder (dated) under Claude\
STAMP = datetime.now().strftime("%Y-%m-%d")
OUT = os.path.join("C:\\Users\\Kapil Nayyar\\Claude", f"CSP_Exit_FullData_{STAMP}")
os.makedirs(OUT, exist_ok=True)


def gclient():
    creds = Credentials.from_service_account_file(
        os.path.join(HERE, "google_credentials.json"),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets.readonly",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    return gspread.authorize(creds)


def dump_rows(name, rows):
    """Write list-of-dicts to both JSON and CSV. Returns row count."""
    safe = name.replace(" ", "_")
    for ch in '<>:"/\\|?*':          # strip characters Windows forbids in filenames
        safe = safe.replace(ch, "-")
    with open(os.path.join(OUT, f"{safe}.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    if rows:
        # union of all keys, preserving first-seen order
        cols = []
        for r in rows:
            for k in r.keys():
                if k not in cols:
                    cols.append(k)
        with open(os.path.join(OUT, f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            for r in rows:
                w.writerow({c: r.get(c, "") for c in cols})
    return len(rows)


def ws_to_records(ws):
    """Read a worksheet -> list-of-dicts, tolerant of trailing empty headers."""
    vals = ws.get_all_values()
    if not vals:
        return []
    header = vals[0]
    last = 0
    for i, h in enumerate(header):
        if h and h.strip():
            last = i + 1
    headers = [h.strip() for h in header[:last]]
    # de-dupe blank/duplicate headers so dict keys stay unique
    seen = {}
    clean = []
    for i, h in enumerate(headers):
        h = h or f"col_{i}"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        clean.append(h)
    out = []
    for row in vals[1:]:
        padded = row[:last] + [""] * max(0, last - len(row))
        out.append(dict(zip(clean, padded)))
    return out


def pull_workbook(client, sheet_id, label):
    book = client.open_by_key(sheet_id)
    summary = []
    for ws in book.worksheets():
        recs = ws_to_records(ws)
        n = dump_rows(f"{label}__{ws.title}", recs)
        summary.append((ws.title, n))
        print(f"  [{label}] {ws.title}: {n} rows")
    return summary


def pull_supabase():
    if not (SUPABASE_URL and SUPABASE_ANON_KEY):
        print("  Supabase not configured — skipped")
        return []
    headers = {
        "apikey": SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Accept": "application/json",
    }
    # partners (full row so the leaders' tracker has every field available)
    all_p = []
    offset = 0
    while True:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/partners",
            params={"select": "*", "limit": 1000, "offset": offset},
            headers=headers, timeout=30,
        )
        if r.status_code != 200:
            print(f"  Supabase partners HTTP {r.status_code}: {r.text[:200]}")
            break
        batch = r.json()
        all_p.extend(batch)
        if len(batch) < 1000:
            break
        offset += 1000
    dump_rows("supabase__partners", all_p)
    print(f"  [supabase] partners: {len(all_p)} rows")
    return all_p


def main():
    client = gclient()
    manifest = {"generated_at": datetime.now().isoformat(), "output_dir": OUT, "sheets": {}}

    print("Main workbook (CSP Exit Tracker sheet):")
    manifest["sheets"]["main_workbook"] = pull_workbook(client, GOOGLE_SHEET_ID, "MAIN")

    print("PX/CX Migration Summary workbook:")
    try:
        manifest["sheets"]["px_workbook"] = pull_workbook(client, PX_MIGRATION_WORKBOOK_ID, "PX")
    except Exception as e:
        print(f"  PX workbook error: {e}")
        manifest["sheets"]["px_workbook"] = f"error: {e}"

    print("Supabase:")
    parts = pull_supabase()
    manifest["sheets"]["supabase_partners"] = len(parts)

    with open(os.path.join(OUT, "_manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"\nDONE. All files in: {OUT}")


if __name__ == "__main__":
    main()
