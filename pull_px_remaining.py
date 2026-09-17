"""Pull only the PX workbook tabs not already dumped (quota-friendly)."""
import os, json, csv, time
from pull_all_data import gclient, ws_to_records, dump_rows, OUT, PX_MIGRATION_WORKBOOK_ID

client = gclient()
book = client.open_by_key(PX_MIGRATION_WORKBOOK_ID)
for ws in book.worksheets():
    safe = "PX__" + ws.title.replace(" ", "_")
    for ch in '<>:"/\\|?*':
        safe = safe.replace(ch, "-")
    if os.path.exists(os.path.join(OUT, safe + ".json")):
        continue
    recs = ws_to_records(ws)
    dump_rows("PX__" + ws.title, recs)
    print(f"  [PX] {ws.title}: {len(recs)} rows")
    time.sleep(1.2)  # stay under the per-minute read quota
print("PX remaining done.")
