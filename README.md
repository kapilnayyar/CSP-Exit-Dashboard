# CSP Exit Tracker — Live Dashboard

Live web dashboard for the CSP exit funnel — same 5 tables as `CSP_Exit_Tracker.xlsx`, auto-refreshed every 30 seconds from the source Google Sheet.

## Local run

1. Create `.env` (copy from `.env.example`):
   ```
   APP_PASSWORD=<your-password>
   GOOGLE_SHEET_ID=<sheet-id-from-url>
   ```
2. Place `google_credentials.json` (Google service account key) in this folder.
3. Run:
   ```
   run_dashboard.bat
   ```
   Opens at http://localhost:8501.

## Deploy (Streamlit Cloud)

1. Push this folder to `github.com/kapilnayyar/CSP-Exit-Dashboard` (private/public).
2. On share.streamlit.io → New app → point at `dashboard.py`.
3. App Settings → Secrets → paste:
   ```toml
   APP_PASSWORD = "..."
   GOOGLE_SHEET_ID = "..."

   [gcp_service_account]
   type = "service_account"
   project_id = "..."
   # (full service account JSON fields)
   ```
4. Share the source Google Sheet with the service account email (read access).

## Source data expected

- **Sheet1** = U2 device pickup data (one row per customer). Required columns:
  `Mobile`, `Calling Status`, `Device Picked (Ajinkya/Pradeep)`, `NQT Ground Remarks`.
- **Sheet2** = U1 migration aggregates (one row per partner). Required columns:
  `Exit Partner Name`, `Total U1 User`, `Migrated`, `Not Migrated`, `Migration in process`, `Major Reason`.

## Whitelist

Any `@wiom.in` email + correct `APP_PASSWORD` can log in. To restrict further, change `ALLOWED_DOMAINS` / `_email_allowed()` in `dashboard.py`.
