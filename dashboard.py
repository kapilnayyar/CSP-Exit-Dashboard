"""
CSP Exit Tracker — Live Streamlit Dashboard

Live web view of the CSP exit funnel, sourced from the Google Sheet your team
updates daily. Mirrors the 5-table format of CSP_Exit_Tracker.xlsx exactly.

Auth: email + APP_PASSWORD (only @wiom.in emails can log in).
Data: pulled from Google Sheets every 5 minutes via service account.
"""

import streamlit as st
import gspread
import requests
from google.oauth2.service_account import Credentials
from streamlit_autorefresh import st_autorefresh
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import hashlib
import json
import os
import pandas as pd

IST = ZoneInfo("Asia/Kolkata")

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

ALLOWED_DOMAINS = ("@wiom.in",)

# Partner codes excluded from the dashboard (per-CSP exceptions).
# These are dropped at the data-fetch boundary so they don't appear in any tab.
EXCLUDED_PARTNER_CODES = {
    281749854653857,
    281749854733209,
    274877909399,
    274877951823,
    281749854674788,
    281749854632442,
}

# Sheet partner name -> canonical Supabase partner name (lowercased keys).
# Use this when a CSP appears in the Main sheet under a different name than in Supabase.
SHEET_NAME_ALIAS = {
    "network solutions": "MANISHA TRADERS 1",
    "manisha traders 2": "MANISHA TRADERS 1",
    "khan enterprises": "MANISHA TRADERS 1",
}

_TOKEN_SALT = "csp-exit-wiom-dashboard-2026"


def _email_allowed(email: str) -> bool:
    e = (email or "").strip().lower()
    return any(e.endswith(d) for d in ALLOWED_DOMAINS)

# ─────────────────────────────────────────────────────────────────────────────
# AUTH HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _load_app_password():
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
        return os.getenv("APP_PASSWORD", "")


def _make_token(email, correct_pw):
    raw = f"{email}|{correct_pw}|{_TOKEN_SALT}"
    return hashlib.sha256(raw.encode()).hexdigest()


def get_secrets():
    try:
        return {
            "sheet_id": st.secrets["GOOGLE_SHEET_ID"],
            "gcp_creds": dict(st.secrets["gcp_service_account"]),
            "supabase_url": st.secrets.get("SUPABASE_URL", ""),
            "supabase_key": st.secrets.get("SUPABASE_ANON_KEY", ""),
            "metabase_url": st.secrets.get("METABASE_URL", ""),
            "metabase_key": st.secrets.get("METABASE_API_KEY", ""),
        }
    except Exception:
        from dotenv import load_dotenv
        BASE = os.path.dirname(os.path.abspath(__file__))
        load_dotenv(os.path.join(BASE, ".env"))
        with open(os.path.join(BASE, "google_credentials.json")) as f:
            gcp = json.load(f)
        return {
            "sheet_id": os.getenv("GOOGLE_SHEET_ID"),
            "gcp_creds": gcp,
            "supabase_url": os.getenv("SUPABASE_URL", ""),
            "supabase_key": os.getenv("SUPABASE_ANON_KEY", ""),
            "metabase_url": os.getenv("METABASE_URL", ""),
            "metabase_key": os.getenv("METABASE_API_KEY", ""),
        }


# ─────────────────────────────────────────────────────────────────────────────
# DATA FETCH (cached 30s — matches autorefresh interval)
# ─────────────────────────────────────────────────────────────────────────────

U2_TAB = "Main sheet"
U1_TAB = "Migration Data"
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


@st.cache_data(ttl=300)
def fetch_sheets(sheet_id, gcp_creds):
    """Pull the two relevant tabs by name and return rows as list-of-dicts."""
    creds = Credentials.from_service_account_info(
        gcp_creds,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    client = gspread.authorize(creds)
    book = client.open_by_key(sheet_id)

    # U2 — one row per customer (device pickup tracking)
    u2_rows = book.worksheet(U2_TAB).get_all_records()
    # U1 — one row per partner (migration aggregates)
    u1_rows = book.worksheet(U1_TAB).get_all_records()

    return u2_rows, u1_rows


@st.cache_data(ttl=300)
def fetch_netbox_collection(sheet_id, gcp_creds):
    """Pull 'S5 Netbox Collection' tab — {partner_code: devices_collected}."""
    creds = Credentials.from_service_account_info(
        gcp_creds,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    client = gspread.authorize(creds)
    book = client.open_by_key(sheet_id)
    try:
        ws = book.worksheet("S5 Netbox Collection")
    except Exception:
        return {}
    rows = ws.get_all_records()
    out = {}
    for r in rows:
        csp_id = str(r.get("CSP ID") or "").strip()
        if not csp_id: continue
        try:
            # Sheet's actual column header is "Devices collected from CSP" — keep this as-is
            out[csp_id] = int(float(r.get("Devices collected from CSP") or 0))
        except (TypeError, ValueError):
            out[csp_id] = 0
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SUPABASE — partner-exit-tracker state (read-only via anon key)
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def fetch_partners(supabase_url, supabase_key):
    """Pull all partners with state, partner_code, u1/u2 counts. SELECT-only."""
    if not supabase_url or not supabase_key:
        return []
    headers = {"apikey": supabase_key, "Authorization": f"Bearer {supabase_key}", "Accept": "application/json"}
    # partners table → get state + code + name
    r = requests.get(
        f"{supabase_url}/rest/v1/partners",
        params={"select": "id,name,partner_code,current_state,risk_state,exit_type"},
        headers=headers, timeout=15,
    )
    partners = r.json() if r.status_code == 200 else []
    # partner_details_extended for u1/u2 counts
    r2 = requests.get(
        f"{supabase_url}/rest/v1/partner_details_extended",
        params={"select": "id,u1_count,u2_count"},
        headers=headers, timeout=15,
    )
    de = {row["id"]: row for row in (r2.json() if r2.status_code == 200 else [])}
    for p in partners:
        d = de.get(p["id"], {})
        p["u1_count"] = d.get("u1_count") or 0
        p["u2_count"] = d.get("u2_count") or 0
    # Drop excluded partners (per-CSP exceptions list)
    partners = [p for p in partners if int(p.get("partner_code") or 0) not in EXCLUDED_PARTNER_CODES]
    return partners


@st.cache_data(ttl=300)
def fetch_u1_customers(supabase_url, supabase_key):
    """Pull U1 customer-level records — used for dedup against Metabase IDLE devices."""
    if not supabase_url or not supabase_key:
        return []
    headers = {"apikey": supabase_key, "Authorization": f"Bearer {supabase_key}", "Accept": "application/json"}
    # Paginate over u1_customers (1631+ rows, default REST limit 1000)
    all_rows = []
    offset = 0
    while True:
        r = requests.get(
            f"{supabase_url}/rest/v1/u1_customers",
            params={"select": "customer_mobile,partner_id,installation_completed", "limit": 1000, "offset": offset},
            headers=headers, timeout=15,
        )
        if r.status_code != 200:
            break
        batch = r.json()
        all_rows.extend(batch)
        if len(batch) < 1000:
            break
        offset += 1000
    return all_rows


# ─────────────────────────────────────────────────────────────────────────────
# METABASE — read-only queries via API key
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def metabase_query(metabase_url, api_key, sql, database_id=113):
    """Run a Snowflake query via Metabase API. Cached 5 min (heavier query)."""
    if not metabase_url or not api_key:
        return {"rows": [], "cols": [], "error": "Metabase not configured"}
    try:
        r = requests.post(
            f"{metabase_url}/api/dataset",
            headers={"x-api-key": api_key, "Content-Type": "application/json"},
            json={"database": database_id, "type": "native", "native": {"query": sql}},
            timeout=60,
        )
        if r.status_code not in (200, 202):
            return {"rows": [], "cols": [], "error": f"HTTP {r.status_code}"}
        d = r.json().get("data", {})
        cols = [c.get("display_name") or c.get("name") for c in d.get("cols", [])]
        return {"rows": d.get("rows", []), "cols": cols, "error": None}
    except Exception as e:
        return {"rows": [], "cols": [], "error": str(e)}


def metabase_safe_in_list(names):
    """Build a SQL IN-clause string from a name list, single-quote escaped."""
    return ", ".join("'" + str(n).replace("'", "''") + "'" for n in names)


def metabase_id_in_list(codes):
    """Build a SQL IN-clause string from numeric partner_account_ids (no quotes)."""
    return ", ".join(str(c) for c in codes if c)


@st.cache_data(ttl=300)
def fetch_r15_active_by_code(metabase_url, api_key, partner_codes):
    """Return {partner_code: active_r15_count} using the canonical Wiom R15 SQL.
    Attribution = CURRENT serving partner via T_WG_CUSTOMER.LCO_ACCOUNT_ID.
    Dedup: latest expiry per NAS, then latest expiry per MOBILE."""
    if not partner_codes:
        return {}
    in_list = metabase_id_in_list(partner_codes)
    if not in_list:
        return {}
    sql = f"""WITH partner_nas AS (
  SELECT DISTINCT NASID, MOBILE, LCO_ACCOUNT_ID
  FROM PUBLIC.T_WG_CUSTOMER
  WHERE LCO_ACCOUNT_ID IN ({in_list})
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
       COUNT_IF(LATEST_EXPIRY >= DATEADD(day, -15, CURRENT_DATE)) AS R15_ACTIVE
FROM per_mobile
GROUP BY 1"""
    res = metabase_query(metabase_url, api_key, sql)
    return {str(row[0]): row[1] for row in res["rows"]}


@st.cache_data(ttl=300)
def fetch_idle_devices_total(metabase_url, api_key, partner_codes):
    """Return total IDLE devices summed across given partner_codes (= Netbox at CSPs)."""
    if not partner_codes:
        return 0
    in_list = metabase_id_in_list(partner_codes)
    if not in_list:
        return 0
    sql = f"""SELECT COUNT(*) AS total
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_list})"""
    res = metabase_query(metabase_url, api_key, sql)
    return res["rows"][0][0] if res["rows"] else 0


@st.cache_data(ttl=300)
def fetch_idle_devices_by_code(metabase_url, api_key, partner_codes):
    """Return {partner_code: count_idle_devices}. Used for per-partner snapshot row."""
    if not partner_codes:
        return {}
    in_list = metabase_id_in_list(partner_codes)
    if not in_list:
        return {}
    sql = f"""SELECT td."LCO_ACCOUNT_ID", COUNT(*) AS idle_cnt
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_list})
GROUP BY 1"""
    res = metabase_query(metabase_url, api_key, sql)
    return {str(row[0]): row[1] for row in res["rows"]}


@st.cache_data(ttl=300)
def search_devices_at_partner(metabase_url, api_key, partner_code):
    """Tab 4 search 1 — IDLE devices for a partner.
    Mirrors Metabase Q11266 'At Partner Device IDLE' which filters by PARTNER_ID
    (= LCO_ACCOUNT_ID) — immune to the supply_model.partner_name whitespace bug
    that affects the older partner_name-based card Q10398."""
    sql = f"""SELECT
  td."DEVICE_ID",
  td."MAC",
  td."SERIAL",
  td."MODEL",
  td."STATUS",
  td."LCO_ACCOUNT_ID",
  td."ADDED_TIME",
  td."STATUS_UPDATED_AT"
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE'
  AND td."LCO_ACCOUNT_ID" = {int(partner_code)}
ORDER BY td."STATUS_UPDATED_AT" DESC NULLS LAST, td."ADDED_TIME" DESC"""
    return metabase_query(metabase_url, api_key, sql)


@st.cache_data(ttl=300)
def fetch_idle_nas_by_code(metabase_url, api_key, partner_codes):
    """For each partner_code, return a set of NAS_IDs currently in IDLE state."""
    if not partner_codes:
        return {}
    in_list = metabase_id_in_list(partner_codes)
    if not in_list:
        return {}
    sql = f"""SELECT td."LCO_ACCOUNT_ID", td."NASID"
FROM "PROD_DB"."POSTGRES_RDS_INVENTORY_INVENTORY"."T_DEVICE" td
WHERE td."STATUS" = 'IDLE' AND td."LCO_ACCOUNT_ID" IN ({in_list})
  AND td."NASID" IS NOT NULL"""
    res = metabase_query(metabase_url, api_key, sql)
    out = defaultdict(set)
    for row in res["rows"]:
        out[str(row[0])].add(str(row[1]))
    return out


@st.cache_data(ttl=300)
def fetch_mobile_to_nas(metabase_url, api_key, mobiles):
    """Return {mobile: router_nas_id} for the given mobile list (most recent assignment)."""
    if not mobiles:
        return {}
    mobs = [str(m).strip() for m in mobiles if m]
    # Quote each mobile (they're stored as strings in t_router_user_mapping)
    in_list = ", ".join("'" + m.replace("'", "''") + "'" for m in mobs)
    sql = f"""SELECT mobile, router_nas_id
FROM prod_db.public.t_router_user_mapping
WHERE mobile IN ({in_list})
  AND router_nas_id IS NOT NULL
QUALIFY ROW_NUMBER() OVER (PARTITION BY mobile ORDER BY created_on DESC NULLS LAST) = 1"""
    res = metabase_query(metabase_url, api_key, sql)
    return {str(row[0]): str(row[1]) for row in res["rows"]}


@st.cache_data(ttl=300)
def search_active_userbase(metabase_url, api_key, partner_code):
    """Tab 4 search 2 — R15 active customers (mobile + netbox + latest expiry) for a partner.
    Uses canonical Wiom R15 SQL: current serving partner via T_WG_CUSTOMER.LCO_ACCOUNT_ID,
    dedup latest expiry per NAS, dedup latest expiry per MOBILE."""
    sql = f"""WITH partner_nas AS (
  SELECT DISTINCT NASID, MOBILE, DEVICE_ID
  FROM PUBLIC.T_WG_CUSTOMER
  WHERE LCO_ACCOUNT_ID = {int(partner_code)}
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
  SELECT pn.MOBILE, MAX(pn.DEVICE_ID) AS NETBOX_ID,
         MAX(l.LATEST_EXPIRY) AS LATEST_EXPIRY
  FROM partner_nas pn
  LEFT JOIN live_latest l ON pn.NASID = l.ROUTER_NAS_ID
  GROUP BY pn.MOBILE
)
SELECT MOBILE, NETBOX_ID, CAST(LATEST_EXPIRY AS DATE) AS PLAN_EXPIRY
FROM per_mobile
WHERE LATEST_EXPIRY >= DATEADD(day, -15, CURRENT_DATE)
ORDER BY LATEST_EXPIRY DESC"""
    return metabase_query(metabase_url, api_key, sql)


# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION (ported from build_csp_exit_tracker.py)
# ─────────────────────────────────────────────────────────────────────────────

def classify_u2(rows):
    """
    Classify each U2 row using column Q "Remarks Dropdown" as the single source of truth.
      Device Picked Up               -> Picked by Team
      Partner Collected the Device   -> Picked by Partner
      Anything else (incl. blank)    -> Still with User; cohort = the dropdown value
    """
    team = partner = swu = total = 0
    swu_cohorts = Counter()

    for row in rows:
        # Sheet column was renamed from "Mobile" to "Mobile no" — accept either
        mobile = row.get("Mobile no") or row.get("Mobile")
        if not mobile or str(mobile).strip() == "":
            continue
        total += 1
        remark = str(row.get("Remarks Dropdown") or "").strip()
        rl = remark.lower()

        if rl == "device picked up":
            team += 1
        elif rl == "partner collected the device":
            partner += 1
        else:
            swu += 1
            cohort = remark if remark else "(Blank / pending)"
            swu_cohorts[cohort] += 1

    return {
        "total": total,
        "team": team,
        "partner": partner,
        "swu": swu,
        "swu_cohorts": swu_cohorts,
    }


def classify_u1(rows):
    """Aggregate U1 counts and reason cohorts. Merges Not Feasible + Area not feasible."""
    total = mig = notmig = inproc = 0
    reasons = Counter()

    for row in rows:
        name = row.get("Exit Partner Name")
        if not name or str(name).strip() == "":
            continue
        tu = _to_int(row.get("Total U1 User"))
        m = _to_int(row.get("Migrated"))
        nm = _to_int(row.get("Not Migrated"))
        # WIP column was renamed from "Migration in process" — accept either.
        ip = _to_int(row.get("WIP") or row.get("Migration in process"))
        rsn = row.get("Major Reason")
        total += tu
        mig += m
        notmig += nm
        inproc += ip
        if rsn and str(rsn).strip():
            reasons[str(rsn).strip()] += (nm + ip)

    # Merge case-variant keys (e.g., "Not Feasible" / "Not feasible")
    canon = {}
    for k in reasons:
        canon.setdefault(k.lower(), k)
    merged = Counter()
    for k, v in reasons.items():
        merged[canon[k.lower()]] += v

    # Merge "Not Feasible" + "Area not feasible"
    _merge_keys(
        merged,
        ["not feasible", "area not feasible"],
        "Not Feasible / Area not feasible",
        case_insensitive=True,
    )

    return {
        "total": total,
        "migrated": mig,
        "not_migrated": notmig,
        "in_process": inproc,
        "reasons": merged,
    }


def _to_int(v):
    try:
        if v is None or v == "":
            return 0
        return int(float(v))
    except (ValueError, TypeError):
        return 0


def _merge_keys(counter, keys, new_label, case_insensitive=False):
    total = 0
    targets = [k.lower() for k in keys] if case_insensitive else keys
    for k in list(counter.keys()):
        check = k.lower() if case_insensitive else k
        if check in targets or k in keys:
            total += counter.pop(k)
    if total:
        counter[new_label] = counter.get(new_label, 0) + total


# ─────────────────────────────────────────────────────────────────────────────
# RENDERING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def pct(n, denom):
    return f"{(n / denom * 100):.1f}%" if denom else "0.0%"


def table_header(title, color="#2E75B6"):
    return (
        f'<div style="background:{color};color:#ffffff;padding:10px 14px;'
        f'font-weight:bold;font-size:14px;margin-top:18px;border-radius:6px 6px 0 0">'
        f'{title}</div>'
    )


def open_table():
    return (
        '<table class="csp-table">'
        '<tr>'
        '<th style="background:#2E75B6;color:#ffffff;width:60px">S.No</th>'
        '<th style="background:#2E75B6;color:#ffffff">Category / Reason</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:90px;text-align:right">Count</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:110px;text-align:right">%</th>'
        '</tr>'
    )


def row(sno, label, count, percent, status_color=None):
    label_bg = status_color if status_color else "#ffffff"
    return (
        f'<tr>'
        f'<td style="background:#ffffff;color:#000000;text-align:center">{sno}</td>'
        f'<td style="background:{label_bg};color:#000000">{label}</td>'
        f'<td style="background:#ffffff;color:#000000;text-align:right;font-weight:bold">{count:,}</td>'
        f'<td style="background:#ffffff;color:#000000;text-align:right">{percent}</td>'
        f'</tr>'
    )


def total_row(label, count, percent="100.0%"):
    return (
        f'<tr>'
        f'<td style="background:#D9E1F2;color:#000000;font-weight:bold;text-align:center">—</td>'
        f'<td style="background:#D9E1F2;color:#000000;font-weight:bold">{label}</td>'
        f'<td style="background:#D9E1F2;color:#000000;font-weight:bold;text-align:right">{count:,}</td>'
        f'<td style="background:#D9E1F2;color:#000000;font-weight:bold;text-align:right">{percent}</td>'
        f'</tr>'
    )


def close_table():
    return "</table>"


# Status colors (match the Excel)
GREEN = "#C6EFCE"
AMBER = "#FFEB9C"
RED = "#F8CBAD"


def render_tab1_status(u1, u2):
    """Tab 1 — the original 5-table view (Status & Cohorts)."""
    grand_total = u1["total"] + u2["total"]
    u1_pending = u1["not_migrated"] + u1["in_process"]
    u2_not_team = u2["partner"] + u2["swu"]

    # ── Table 1: Total Userbase — U1 vs U2 Bifurcation ───────────────────────
    html = table_header("Total Userbase — U1 vs U2 Bifurcation")
    html += open_table()
    html += row(1, "U1 — Migration", u1["total"], pct(u1["total"], grand_total))
    html += row(2, "U2 — Device Pickup", u2["total"], pct(u2["total"], grand_total))
    html += total_row("Total Userbase", grand_total)
    html += close_table()
    st.markdown(html, unsafe_allow_html=True)

    # ── Table 2: U1 — Migration Status ────────────────────────────────────────
    html = table_header("U1 — Migration Status")
    html += open_table()
    html += row(1, "Migrated ✓", u1["migrated"], pct(u1["migrated"], u1["total"]), GREEN)
    html += row(2, "Not Migrated 🔴", u1["not_migrated"], pct(u1["not_migrated"], u1["total"]), RED)
    html += row(3, "Migration in Process ⚠", u1["in_process"], pct(u1["in_process"], u1["total"]), AMBER)
    html += total_row("Total U1", u1["total"])
    html += close_table()
    st.markdown(html, unsafe_allow_html=True)

    # ── Table 3: U2 — Device Pickup Status ────────────────────────────────────
    html = table_header("U2 — Device Pickup Status")
    html += open_table()
    html += row(1, "Picked by Team ✓", u2["team"], pct(u2["team"], u2["total"]), GREEN)
    html += row(2, "Picked by Partner", u2["partner"], pct(u2["partner"], u2["total"]), AMBER)
    html += row(3, "Still with User 🔴", u2["swu"], pct(u2["swu"], u2["total"]), RED)
    html += total_row("Total U2", u2["total"])
    html += close_table()
    st.markdown(html, unsafe_allow_html=True)

    # ── Table 4: U1 — Reason Cohorts ──────────────────────────────────────────
    html = table_header("U1 — Reason Cohorts (Customers where migration is NOT done)")
    html += open_table()
    s = 1
    reason_sum = 0
    for reason, cnt in u1["reasons"].most_common():
        html += row(s, reason, cnt, pct(cnt, u1_pending))
        reason_sum += cnt
        s += 1
    others_blank = u1_pending - reason_sum
    if others_blank > 0:
        html += row(s, "Others / Blank", others_blank, pct(others_blank, u1_pending))
    html += total_row("Total", u1_pending)
    html += close_table()
    st.markdown(html, unsafe_allow_html=True)

    # ── Table 5: U2 — Reason Cohorts ──────────────────────────────────────────
    html = table_header("U2 — Reason Cohorts (Customers where device is NOT picked by our team)")
    html += open_table()
    # Partner Collected the Device at top
    u2_cohort_list = [("Partner Collected the Device", u2["partner"])]
    u2_cohort_list += u2["swu_cohorts"].most_common()
    # Top 10 + Others
    top10 = u2_cohort_list[:10]
    others_total = sum(v for _, v in u2_cohort_list[10:])
    s = 1
    for reason, cnt in top10:
        html += row(s, reason, cnt, pct(cnt, u2_not_team))
        s += 1
    if others_total > 0:
        html += row(s, "Others", others_total, pct(others_total, u2_not_team))
    html += total_row("Total", u2_not_team)
    html += close_table()
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# FUNNEL HELPERS — sheet lookups by partner name
# ─────────────────────────────────────────────────────────────────────────────

def build_sheet_lookups(u1_rows, u2_rows):
    """Return per-partner U1/U2 metrics from Google Sheet, keyed by lowercased name."""
    u1_by = {}
    for r in u1_rows:
        name = str(r.get("Exit Partner Name") or "").strip()
        if not name:
            continue
        key = name.lower()
        key = SHEET_NAME_ALIAS.get(key, name).lower() if key in SHEET_NAME_ALIAS else key
        # Aggregate if alias collapses multiple sheet names into one canonical
        existing = u1_by.get(key, {"total": 0, "migrated": 0})
        u1_by[key] = {
            "total": existing["total"] + _to_int(r.get("Total U1 User")),
            "migrated": existing["migrated"] + _to_int(r.get("Migrated")),
        }
    u2_total = defaultdict(int)
    u2_picked = defaultdict(int)
    for r in u2_rows:
        name = str(r.get("Partner") or "").strip()
        mobile = r.get("Mobile no") or r.get("Mobile")
        if not name or not mobile:
            continue
        # Map sheet name to canonical Supabase name if an alias is set
        key = name.lower()
        key = SHEET_NAME_ALIAS.get(key, name).lower() if key in SHEET_NAME_ALIAS else key
        u2_total[key] += 1
        if str(r.get("Remarks Dropdown") or "").strip().lower() == "device picked up":
            u2_picked[key] += 1
    return u1_by, u2_total, u2_picked


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — CSP Exit Funnel
# ─────────────────────────────────────────────────────────────────────────────

# Funnel stage colors
STAGE_COLORS = {
    "S1": "#1F4E78",
    "S2": "#2E75B6",
    "S3": "#B7950B",
    "S4a": "#548235",
    "S4b": "#BF9000",
    "S4c": "#375623",
    "S5": "#9C0006",
    "S6": "#404040",
}


def stage_card(stage_label, color, metrics):
    """Render one funnel stage: colored header + 4-column table (S.No / Category / Count / %).
    metrics is a list of (label, value, percent_string). Pass percent_string="" or None to hide."""
    html = (
        f'<div style="background:{color};color:#ffffff;padding:10px 14px;'
        f'font-weight:bold;font-size:14px;margin-top:14px;border-radius:6px 6px 0 0">'
        f'{stage_label}</div>'
        '<table class="csp-table">'
        '<tr>'
        '<th style="background:#2E75B6;color:#ffffff;width:60px">S.No</th>'
        '<th style="background:#2E75B6;color:#ffffff">Category</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:120px;text-align:right">Count</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:80px;text-align:right">%</th>'
        '</tr>'
    )
    serial = 0
    for item in metrics:
        label, value, pct_str = item if len(item) == 3 else (item[0], item[1], "")
        # Sub-rows (indented with leading whitespace) don't get a serial number
        is_sub = isinstance(label, str) and label.startswith(" ")
        if not is_sub:
            serial += 1
            serial_str = str(serial)
        else:
            serial_str = ""
        v_str = f"{value:,}" if isinstance(value, int) else str(value)
        html += (
            f'<tr>'
            f'<td style="background:#ffffff;color:#000000;text-align:center">{serial_str}</td>'
            f'<td style="background:#ffffff;color:#000000">{label}</td>'
            f'<td style="background:#ffffff;color:#000000;text-align:right;font-weight:bold">{v_str}</td>'
            f'<td style="background:#ffffff;color:#000000;text-align:right">{pct_str or ""}</td>'
            f'</tr>'
        )
    html += "</table>"
    return html


def fmt_pct(n, denom):
    if not denom: return ""
    return f"{(n / denom * 100):.1f}%"


# ─────────────────────────────────────────────────────────────────────────────
# DAILY TOTALS + DELTA (Tab 5 support)
# ─────────────────────────────────────────────────────────────────────────────

def compute_today_metrics(partners, u1_by, u2_total, u2_picked, r15_by_code,
                          idle_total, s5_dedup, netbox_collected_by_code,
                          idle_total_s6):
    """Aggregate today's funnel metrics into the flat dict written to Daily Totals.

    Self-contained: computes S5 reconciliation (could_not_pick, liability,
    devices collected) and S6 collected internally — same formulas as Tab 2.
    """
    netbox_collected_by_code = netbox_collected_by_code or {}
    by_state = defaultdict(list)
    for p in partners:
        by_state[p["current_state"]].append(p)

    def r15_of(p):
        return r15_by_code.get(str(p.get("partner_code") or ""), 0)

    def sheet_ub(p):
        key = p["name"].lower()
        return (u1_by.get(key, {}).get("total", 0) or 0) + (u2_total.get(key, 0) or 0)

    def ub_of(p):
        s = sheet_ub(p)
        return s if s > 0 else r15_of(p)

    in_pipeline = [p for p in partners if p.get("current_state") in ("S1","S2","S3","S4","S5","S6")]
    current_s2 = by_state.get("S2", [])
    past_s3 = [p for p in in_pipeline if p["current_state"] in ("S3","S4","S5","S6")]
    s4_partners = by_state.get("S4", [])
    s5_partners = by_state.get("S5", [])
    s6_partners = by_state.get("S6", [])
    completed = s5_partners + s6_partners

    # S1 sums (with exit_type breakdown)
    s1_csps = len(in_pipeline)
    s1_userbase = sum(ub_of(p) for p in in_pipeline)
    s1_voluntary = sum(1 for p in in_pipeline if str(p.get("exit_type") or "").strip() == "Voluntary")
    s1_b1 = sum(1 for p in in_pipeline if str(p.get("exit_type") or "").strip() == "B1")
    s1_b2 = sum(1 for p in in_pipeline if str(p.get("exit_type") or "").strip() == "B2")

    # S2/S3
    s2_csps = len(current_s2)
    s2_userbase = sum(ub_of(p) for p in current_s2)
    s3_csps = len(past_s3)
    s3_userbase = sum(ub_of(p) for p in past_s3)

    # S4a — completed (in S5 or S6)
    s4a_csps = len(completed)
    s4a_u1_total = sum(u1_by.get(p["name"].lower(), {}).get("total", 0) for p in completed)
    s4a_u1_mig = sum(u1_by.get(p["name"].lower(), {}).get("migrated", 0) for p in completed)
    s4a_u2_total = sum(u2_total.get(p["name"].lower(), 0) for p in completed)
    s4a_u2_pick = sum(u2_picked.get(p["name"].lower(), 0) for p in completed)

    # S4b — currently in S4
    s4b_csps = len(s4_partners)
    s4b_u1 = s4b_u1_mig = s4b_u2 = s4b_u2_pick = s4b_pending = 0
    for p in s4_partners:
        key = p["name"].lower()
        u1d = u1_by.get(key, {"total": 0, "migrated": 0})
        s4b_u1 += u1d["total"]
        s4b_u1_mig += u1d["migrated"]
        s4b_u2 += u2_total.get(key, 0)
        s4b_u2_pick += u2_picked.get(key, 0)
        if u1d["total"] == 0 and u2_total.get(key, 0) == 0:
            s4b_pending += r15_of(p)

    # S5 reconciliation — same formulas as render_tab2_funnel
    s5_u1_total = s5_u1_mig = s5_u2_total = s5_u2_picked = 0
    for p in s5_partners:
        key = p["name"].lower()
        u1d = u1_by.get(key, {"total": 0, "migrated": 0})
        s5_u1_total += u1d["total"]; s5_u1_mig += u1d["migrated"]
        s5_u2_total += u2_total.get(key, 0); s5_u2_picked += u2_picked.get(key, 0)
    s5_could_not_pick_raw = (s5_u1_total - s5_u1_mig) + (s5_u2_total - s5_u2_picked)
    dup = s5_dedup.get("duplicates", 0) if s5_dedup else 0
    s5_could_not_pick = max(s5_could_not_pick_raw - dup, 0)
    s5_liability = idle_total + s5_could_not_pick
    s5_devices_collected = sum(
        netbox_collected_by_code.get(str(p.get("partner_code") or ""), 0)
        for p in s5_partners
    )

    # S6 — netbox collected
    s6_collected = sum(
        netbox_collected_by_code.get(str(p.get("partner_code") or ""), 0)
        for p in s6_partners
    )

    return {
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
        "s5_collected": s5_devices_collected,
        "s6_csps": len(s6_partners), "s6_idle": idle_total_s6, "s6_collected": s6_collected,
    }


def _get_or_create_totals_tab(book):
    """Return the Daily Totals worksheet, creating it (with header row) if missing."""
    try:
        return book.worksheet(TOTALS_TAB)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=TOTALS_TAB, rows=2000, cols=len(TOTALS_HEADERS))
        ws.append_row(TOTALS_HEADERS)
        return ws


def write_today_totals(book, totals):
    """Idempotently append today's totals row to Daily Totals tab."""
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    try:
        ws = _get_or_create_totals_tab(book)
    except Exception:
        return False
    try:
        existing_dates = set(ws.col_values(1)[1:])
    except Exception:
        existing_dates = set()
    if today_str in existing_dates:
        return False
    row = [today_str] + [int(totals.get(k, 0) or 0) for k in TOTALS_HEADERS[1:]]
    try:
        ws.append_row(row, value_input_option="USER_ENTERED")
    except Exception:
        return False
    return True


def read_yesterday_totals(book):
    """Return {column → int} for yesterday's row in Daily Totals; {} if missing."""
    try:
        ws = book.worksheet(TOTALS_TAB)
    except Exception:
        return {}
    try:
        rows = ws.get_all_records()
    except Exception:
        return {}
    yesterday = (datetime.now(IST) - timedelta(days=1)).strftime("%Y-%m-%d")
    for r in rows:
        if str(r.get("date") or "").strip() == yesterday:
            out = {}
            for k in TOTALS_HEADERS[1:]:
                try:
                    out[k] = int(float(r.get(k) or 0))
                except (TypeError, ValueError):
                    out[k] = 0
            return out
    return {}


def _delta_str(d0, dm1):
    """Format the Delta column. Show '—' if D-1 is missing/None."""
    if dm1 is None:
        return "—"
    d = (d0 or 0) - (dm1 or 0)
    if d > 0: return f"+{d:,}"
    if d < 0: return f"{d:,}"
    return "0"


def stage_card_with_delta(stage_label, color, rows):
    """6-column stage card matching Tab 2 stage_card() styling.
    Columns: S.No / Category / D0 Count / % / D-1 Count / Delta.
    Each row is a 4-tuple (label, value, pct, dm1) — value/dm1/label may be
    int or HTML string (HTML strings render as-is to support multi-line cells
    like Tab 2's S1 CSP+breakdown row). Optionally a 5th element overrides
    the Delta cell HTML (used when value/dm1 are multi-line).

    Every Delta cell has an orange background; the inner text color is still
    green/red/grey by sign so +/-/0 stays readable."""
    html = (
        f'<div style="background:{color};color:#ffffff;padding:10px 14px;'
        f'font-weight:bold;font-size:14px;margin-top:14px;border-radius:6px 6px 0 0">'
        f'{stage_label}</div>'
        '<table class="csp-table">'
        '<tr>'
        '<th style="background:#2E75B6;color:#ffffff;width:60px">S.No</th>'
        '<th style="background:#2E75B6;color:#ffffff">Category</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:90px;text-align:right">Delta</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:120px;text-align:right">D-1 Count</th>'
        '<th style="background:#2E75B6;color:#ffffff;width:170px;text-align:right">D0 Count</th>'
        '</tr>'
    )
    serial = 0
    for item in rows:
        if len(item) == 5:
            label, value, pct_str, dm1, delta_html = item
        else:
            label, value, pct_str, dm1 = item
            delta_html = None

        is_sub = isinstance(label, str) and label.startswith(" ")
        if not is_sub:
            serial += 1
            serial_str = str(serial)
        else:
            serial_str = ""

        if isinstance(value, int):
            v_str = f"{value:,}"
            if pct_str:
                v_cell = (
                    f'<span style="display:inline-block;width:60px;text-align:right;font-weight:bold">{v_str}</span>'
                    f'<span style="display:inline-block;width:70px;text-align:right;color:#666">({pct_str})</span>'
                )
            else:
                v_cell = f'<b>{v_str}</b>'
        else:
            # value is already HTML (e.g. multi-line S1 cell) — pct_str ignored
            v_cell = str(value)
        if isinstance(dm1, int):
            dm1_str = f"{dm1:,}"
        elif isinstance(dm1, str) and dm1:
            dm1_str = dm1
        else:
            dm1_str = "—"

        if delta_html is not None:
            delta_cell = delta_html
        else:
            delta = _delta_str(value if isinstance(value, int) else None,
                               dm1 if isinstance(dm1, int) else None)
            if delta.startswith("+"):
                dc = "#107C10"
            elif delta.startswith("-"):
                dc = "#C42B1C"
            else:
                dc = "#333333"
            delta_cell = f'<span style="color:{dc};font-weight:bold">{delta}</span>'

        html += (
            f'<tr>'
            f'<td style="background:#ffffff;color:#000000;text-align:center">{serial_str}</td>'
            f'<td style="background:#ffffff;color:#000000">{label}</td>'
            f'<td style="background:#FFD8A8;color:#000000;text-align:right">{delta_cell}</td>'
            f'<td style="background:#ffffff;color:#000000;text-align:right">{dm1_str}</td>'
            f'<td style="background:#ffffff;color:#000000;text-align:right">{v_cell}</td>'
            f'</tr>'
        )
    html += "</table>"
    return html


def render_tab5_funnel_with_delta(m, y):
    """Tab 5 — same 7 stages as Tab 2, with D0 / D-1 / Delta columns.
    m = today_metrics dict; y = yesterday_totals dict (empty if missing)."""

    today_str = datetime.now(IST).strftime("%d-%b-%Y")
    has_y = bool(y)

    st.markdown(
        f'<div style="background:#1F4E78;color:#ffffff;padding:14px;border-radius:8px;'
        f'font-size:17px;font-weight:bold;text-align:center;margin-top:6px;margin-bottom:6px">'
        f'CSP EXIT FUNNEL — DAILY DELTA — {today_str}</div>',
        unsafe_allow_html=True,
    )
    if not has_y:
        st.info("ℹ No data for yesterday yet in **Daily Totals** tab. Seed yesterday's row manually to enable D-1 and Delta columns. Today's row is auto-captured.")

    def yd(k):
        return y.get(k) if has_y else None

    # ── S1 — CSP + Voluntary/B1/B2 stacked in ONE row (matches Tab 2 styling)
    lh = "line-height:1.9"

    def _line_delta(d0, dm1, bold=False):
        """Inline delta span for one multi-line entry. Returns '' if dm1 missing."""
        if dm1 is None:
            return "—"
        ds = _delta_str(d0, dm1)
        if ds.startswith("+"):
            dc = "#107C10"
        elif ds.startswith("-"):
            dc = "#C42B1C"
        else:
            dc = "#666666"
        weight = ";font-weight:bold" if bold else ""
        return f'<span style="color:{dc}{weight}">{ds}</span>'

    def _line_dm1(dm1):
        return f"{dm1:,}" if isinstance(dm1, int) else "—"

    yd_csps = yd("s1_csps")
    yd_vol = yd("s1_voluntary")
    yd_b1 = yd("s1_b1")
    yd_b2 = yd("s1_b2")

    s1_cat_html = (
        f'<div style="{lh}"><b>CSP</b><br>'
        f'<span style="color:#666">&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;Voluntary</span><br>'
        f'<span style="color:#666">&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;B1</span><br>'
        f'<span style="color:#666">&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;B2</span></div>'
    )
    def _pair(num, pct, bold=False):
        """Render 'NN,NNN  (XX.X%)' with the number and pct each in fixed-width
        inline-blocks so multiple lines align as two vertical columns."""
        nspan = (
            f'<span style="display:inline-block;width:60px;text-align:right'
            + (";font-weight:bold" if bold else ";color:#666")
            + f'">{num:,}</span>'
        )
        pspan = (
            f'<span style="display:inline-block;width:70px;text-align:right;color:#666">'
            f'({pct})</span>'
        )
        return nspan + pspan

    s1_cnt_html = (
        f'<div style="{lh}">'
        f'{_pair(m["s1_csps"], "100.0%", bold=True)}<br>'
        f'{_pair(m["s1_voluntary"], fmt_pct(m["s1_voluntary"], m["s1_csps"]))}<br>'
        f'{_pair(m["s1_b1"], fmt_pct(m["s1_b1"], m["s1_csps"]))}<br>'
        f'{_pair(m["s1_b2"], fmt_pct(m["s1_b2"], m["s1_csps"]))}</div>'
    )
    s1_pct_html = ""  # % is now baked into s1_cnt_html
    s1_dm1_html = (
        f'<div style="{lh}"><b>{_line_dm1(yd_csps)}</b><br>'
        f'<span style="color:#666">{_line_dm1(yd_vol)}</span><br>'
        f'<span style="color:#666">{_line_dm1(yd_b1)}</span><br>'
        f'<span style="color:#666">{_line_dm1(yd_b2)}</span></div>'
    )
    s1_delta_html = (
        f'<div style="{lh}">{_line_delta(m["s1_csps"], yd_csps, bold=True)}<br>'
        f'{_line_delta(m["s1_voluntary"], yd_vol)}<br>'
        f'{_line_delta(m["s1_b1"], yd_b1)}<br>'
        f'{_line_delta(m["s1_b2"], yd_b2)}</div>'
    )

    st.markdown(stage_card_with_delta(
        "STAGE 1  —  EXIT DECLARED (total in exit pipeline)", STAGE_COLORS["S1"],
        [
            (s1_cat_html, s1_cnt_html, s1_pct_html, s1_dm1_html, s1_delta_html),
            ("Userbase", m['s1_userbase'], "100.0%", yd("s1_userbase")),
        ]
    ), unsafe_allow_html=True)

    # ── S2 ───────────────────────────────────────────────────────────────────
    st.markdown(stage_card_with_delta(
        "STAGE 2  —  NOTICE PERIOD (currently serving)", STAGE_COLORS["S2"],
        [
            ("CSPs", m['s2_csps'], fmt_pct(m['s2_csps'], m['s1_csps']), yd("s2_csps")),
            ("Userbase", m['s2_userbase'], fmt_pct(m['s2_userbase'], m['s1_userbase']), yd("s2_userbase")),
        ]
    ), unsafe_allow_html=True)

    # ── S3 ───────────────────────────────────────────────────────────────────
    st.markdown(stage_card_with_delta(
        "STAGE 3  —  BLOCKING", STAGE_COLORS["S3"],
        [
            ("CSPs", m['s3_csps'], fmt_pct(m['s3_csps'], m['s1_csps']), yd("s3_csps")),
            ("Userbase", m['s3_userbase'], fmt_pct(m['s3_userbase'], m['s1_userbase']), yd("s3_userbase")),
        ]
    ), unsafe_allow_html=True)

    # ── S4a ──────────────────────────────────────────────────────────────────
    u1_conv = fmt_pct(m['s4a_u1_mig'], m['s4a_u1_total'])
    u2_conv = fmt_pct(m['s4a_u2_pick'], m['s4a_u2_total'])
    st.markdown(stage_card_with_delta(
        "STAGE 4a  —  EXECUTION COMPLETED (currently in S5 or S6)", STAGE_COLORS["S4c"],
        [
            ("CSPs", m['s4a_csps'], fmt_pct(m['s4a_csps'], m['s1_csps']), yd("s4a_csps")),
            ("U1 Migration Completed", m['s4a_u1_mig'], u1_conv, yd("s4a_u1_mig")),
            ("U2 Netbox Picked by Wiom", m['s4a_u2_pick'], u2_conv, yd("s4a_u2_pick")),
        ]
    ), unsafe_allow_html=True)

    # ── S4b ──────────────────────────────────────────────────────────────────
    s4b_total = m['s4b_u1'] + m['s4b_u2'] + m['s4b_pending']
    st.markdown(stage_card_with_delta(
        "STAGE 4b  —  EXECUTION IN PROCESS (currently in S4)", STAGE_COLORS["S4a"],
        [
            ("CSPs", m['s4b_csps'], fmt_pct(m['s4b_csps'], m['s1_csps']), yd("s4b_csps")),
            ("U1 Userbase", m['s4b_u1'], fmt_pct(m['s4b_u1'], s4b_total), yd("s4b_u1")),
            ("Migration Done", m['s4b_u1_mig'], fmt_pct(m['s4b_u1_mig'], m['s4b_u1']), yd("s4b_u1_mig")),
            ("U2 Userbase", m['s4b_u2'], fmt_pct(m['s4b_u2'], s4b_total), yd("s4b_u2")),
            ("Netbox Pickup Done", m['s4b_u2_pick'], fmt_pct(m['s4b_u2_pick'], m['s4b_u2']), yd("s4b_u2_pick")),
            ("Userbase Pending to Add", m['s4b_pending'], fmt_pct(m['s4b_pending'], s4b_total), yd("s4b_pending")),
        ]
    ), unsafe_allow_html=True)

    # ── S5 ───────────────────────────────────────────────────────────────────
    st.markdown(stage_card_with_delta(
        "STAGE 5  —  RECONCILIATION (FNF process)", STAGE_COLORS["S5"],
        [
            ("CSPs", m['s5_csps'], fmt_pct(m['s5_csps'], m['s1_csps']), yd("s5_csps")),
            ("Netbox at CSPs", m['s5_idle'], fmt_pct(m['s5_idle'], m['s5_liability']), yd("s5_idle")),
            ("Could not pick (deduped)", m['s5_could_not_pick'], fmt_pct(m['s5_could_not_pick'], m['s5_liability']), yd("s5_could_not_pick")),
            ("Total Netbox Liability", m['s5_liability'], "100.0%", yd("s5_liability")),
            ("Total Netbox Collected from CSP", m['s5_collected'], fmt_pct(m['s5_collected'], m['s5_liability']), yd("s5_collected")),
        ]
    ), unsafe_allow_html=True)

    # ── S6 ───────────────────────────────────────────────────────────────────
    s6_total_dev = m['s6_idle'] + m['s6_collected']
    st.markdown(stage_card_with_delta(
        "STAGE 6  —  COMPLETE", STAGE_COLORS["S6"],
        [
            ("CSPs", m['s6_csps'], fmt_pct(m['s6_csps'], m['s1_csps']), yd("s6_csps")),
            ("Netbox at CSP", m['s6_idle'], fmt_pct(m['s6_idle'], s6_total_dev) if s6_total_dev else "0.0%", yd("s6_idle")),
            ("Total Netbox Collected from CSP", m['s6_collected'], fmt_pct(m['s6_collected'], s6_total_dev) if s6_total_dev else "0.0%", yd("s6_collected")),
        ]
    ), unsafe_allow_html=True)



def render_tab2_funnel(partners, u1_by, u2_total, u2_picked, r15_by_code, idle_total, s5_dedup, idle_total_s6=0, netbox_collected_by_code=None):
    """Tab 2 — funnel. S1/S2/S3 are cumulative; S4a/S4b/S5/S6 are current snapshots.
    % computed against S1 totals."""
    by_state = defaultdict(list)
    for p in partners:
        by_state[p["current_state"]].append(p)

    def r15_of(p):
        return r15_by_code.get(str(p.get("partner_code") or ""), 0)

    def sheet_userbase_of(p):
        """U1+U2 total for this partner from Google Sheet (0 if not in sheet)."""
        key = p["name"].lower()
        u1 = u1_by.get(key, {}).get("total", 0) or 0
        u2 = u2_total.get(key, 0) or 0
        return u1 + u2

    def userbase_of(p):
        """Sheet first; R15 fallback when CSP is not in sheet."""
        sheet = sheet_userbase_of(p)
        return sheet if sheet > 0 else r15_of(p)

    # ── Cumulative pools (S1, S3) + current-state pools (S2) ─────────────────
    in_pipeline = [p for p in partners if p.get("current_state") in ("S1","S2","S3","S4","S5","S6")]
    current_s2 = by_state.get("S2", [])  # only CSPs currently serving notice
    past_s3 = [p for p in in_pipeline if p["current_state"] in ("S3","S4","S5","S6")]

    s1_csps = len(in_pipeline)
    s1_userbase = sum(userbase_of(p) for p in in_pipeline)

    # ── S1 — Total in exit ───────────────────────────────────────────────────
    s1_voluntary = sum(1 for p in in_pipeline if str(p.get("exit_type") or "").strip() == "Voluntary")
    s1_b1 = sum(1 for p in in_pipeline if str(p.get("exit_type") or "").strip() == "B1")
    s1_b2 = sum(1 for p in in_pipeline if str(p.get("exit_type") or "").strip() == "B2")
    # CSP + its 3 breakdown rows in ONE row. Use bullets at same level so it's
    # clear all 3 are peers under CSP (not nested inside each other).
    lh = "line-height:1.9"
    cat_html = (
        f'<div style="{lh}"><b>CSP</b><br>'
        f'<span style="color:#666">&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;Voluntary</span><br>'
        f'<span style="color:#666">&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;B1</span><br>'
        f'<span style="color:#666">&nbsp;&nbsp;&nbsp;&nbsp;&bull;&nbsp;B2</span></div>'
    )
    cnt_html = (
        f'<div style="{lh}"><b>{s1_csps:,}</b><br>'
        f'<span style="color:#666">{s1_voluntary:,}</span><br>'
        f'<span style="color:#666">{s1_b1:,}</span><br>'
        f'<span style="color:#666">{s1_b2:,}</span></div>'
    )
    pct_html = (
        f'<div style="{lh}"><b>100.0%</b><br>'
        f'<span style="color:#666">{fmt_pct(s1_voluntary, s1_csps)}</span><br>'
        f'<span style="color:#666">{fmt_pct(s1_b1, s1_csps)}</span><br>'
        f'<span style="color:#666">{fmt_pct(s1_b2, s1_csps)}</span></div>'
    )
    st.markdown(stage_card("STAGE 1  —  EXIT DECLARED (total in exit pipeline)", STAGE_COLORS["S1"], [
        (cat_html, cnt_html, pct_html),
        ("Userbase", s1_userbase, "100.0%"),
    ]), unsafe_allow_html=True)

    # ── S2 — Currently serving notice period ─────────────────────────────────
    s2_csps = len(current_s2)
    s2_userbase = sum(userbase_of(p) for p in current_s2)
    st.markdown(stage_card("STAGE 2  —  NOTICE PERIOD (currently serving)", STAGE_COLORS["S2"], [
        ("CSPs", s2_csps, fmt_pct(s2_csps, s1_csps)),
        ("Userbase", s2_userbase, fmt_pct(s2_userbase, s1_userbase)),
    ]), unsafe_allow_html=True)

    # ── S3 — Got blocked ─────────────────────────────────────────────────────
    s3_csps = len(past_s3)
    s3_userbase = sum(userbase_of(p) for p in past_s3)
    st.markdown(stage_card("STAGE 3  —  BLOCKING", STAGE_COLORS["S3"], [
        ("CSPs", s3_csps, fmt_pct(s3_csps, s1_csps)),
        ("Userbase", s3_userbase, fmt_pct(s3_userbase, s1_userbase)),
    ]), unsafe_allow_html=True)

    # ── S4a — Execution Completed (currently in S5 or S6) ────────────────────
    s5_partners = by_state.get("S5", [])
    completed_partners = s5_partners + by_state.get("S6", [])
    s4a_u1_total = sum(u1_by.get(p["name"].lower(), {}).get("total", 0) for p in completed_partners)
    s4a_u1_mig = sum(u1_by.get(p["name"].lower(), {}).get("migrated", 0) for p in completed_partners)
    s4a_u2_total = sum(u2_total.get(p["name"].lower(), 0) for p in completed_partners)
    s4a_u2_pick = sum(u2_picked.get(p["name"].lower(), 0) for p in completed_partners)
    s4a_csps_completed = len(completed_partners)

    st.markdown(stage_card("STAGE 4a  —  EXECUTION COMPLETED (currently in S5 or S6)", STAGE_COLORS["S4c"], [
        ("CSPs", s4a_csps_completed, fmt_pct(s4a_csps_completed, s1_csps)),
        ("U1 Migration Completed", s4a_u1_mig, fmt_pct(s4a_u1_mig, s4a_u1_total)),
        ("U2 Netbox Picked by Wiom", s4a_u2_pick, fmt_pct(s4a_u2_pick, s4a_u2_total)),
    ]), unsafe_allow_html=True)

    # ── S4b — Execution In Process (currently in S4) ─────────────────────────
    s4_partners = by_state.get("S4", [])
    s4b_u1 = s4b_u1_mig = s4b_u2 = s4b_u2_pick = s4b_pending = 0
    for p in s4_partners:
        key = p["name"].lower()
        u1d = u1_by.get(key, {"total": 0, "migrated": 0})
        s4b_u1 += u1d["total"]
        s4b_u1_mig += u1d["migrated"]
        s4b_u2 += u2_total.get(key, 0)
        s4b_u2_pick += u2_picked.get(key, 0)
        # If CSP has no sheet data, count its R15 active as "pending to add"
        if u1d["total"] == 0 and u2_total.get(key, 0) == 0:
            s4b_pending += r15_of(p)
    s4b_csps = len(s4_partners)

    s4b_total_userbase = s4b_u1 + s4b_u2 + s4b_pending
    st.markdown(stage_card("STAGE 4b  —  EXECUTION IN PROCESS (currently in S4)", STAGE_COLORS["S4a"], [
        ("CSPs", s4b_csps, fmt_pct(s4b_csps, s1_csps)),
        ("U1 Userbase", s4b_u1, fmt_pct(s4b_u1, s4b_total_userbase)),
        ("Migration Done", s4b_u1_mig, fmt_pct(s4b_u1_mig, s4b_u1)),
        ("U2 Userbase", s4b_u2, fmt_pct(s4b_u2, s4b_total_userbase)),
        ("Netbox Pickup Done", s4b_u2_pick, fmt_pct(s4b_u2_pick, s4b_u2)),
        ("Userbase Pending to Add", s4b_pending, fmt_pct(s4b_pending, s4b_total_userbase)),
    ]), unsafe_allow_html=True)

    # ── S5 — Reconciliation (Netbox metrics) ─────────────────────────────────
    s5_u1_total = s5_u1_mig = s5_u2_total = s5_u2_picked = 0
    for p in s5_partners:
        key = p["name"].lower()
        u1d = u1_by.get(key, {"total": 0, "migrated": 0})
        s5_u1_total += u1d["total"]; s5_u1_mig += u1d["migrated"]
        s5_u2_total += u2_total.get(key, 0); s5_u2_picked += u2_picked.get(key, 0)
    s5_could_not_pick_raw = (s5_u1_total - s5_u1_mig) + (s5_u2_total - s5_u2_picked)
    dup = s5_dedup.get("duplicates", 0)
    s5_could_not_pick = max(s5_could_not_pick_raw - dup, 0)
    s5_liability = idle_total + s5_could_not_pick

    # Netbox collected from CSP for S5 partners — from "S5 Netbox Collection" tab
    netbox_collected_by_code = netbox_collected_by_code or {}
    s5_devices_collected = sum(
        netbox_collected_by_code.get(str(p.get("partner_code") or ""), 0)
        for p in s5_partners
    )
    st.markdown(stage_card("STAGE 5  —  RECONCILIATION (FNF process)", STAGE_COLORS["S5"], [
        ("CSPs", len(s5_partners), fmt_pct(len(s5_partners), s1_csps)),
        ("Netbox at CSPs", idle_total, fmt_pct(idle_total, s5_liability)),
        ("Could not pick (U1+U2 pending)", s5_could_not_pick_raw, fmt_pct(s5_could_not_pick_raw, s5_liability)),
        ("Duplicates U2 (pending customer's netbox already at CSP)", dup, fmt_pct(dup, s5_could_not_pick_raw)),
        ("Could not pick deduped", s5_could_not_pick, fmt_pct(s5_could_not_pick, s5_liability)),
        ("Total Netbox Liability", s5_liability, "100.0%"),
        ("Total Netbox Collected from CSP", s5_devices_collected, fmt_pct(s5_devices_collected, s5_liability)),
    ]), unsafe_allow_html=True)

    # ── S6 — Complete ────────────────────────────────────────────────────────
    s6_csps = len(by_state.get("S6", []))
    if s6_csps > 0:
        netbox_collected_by_code = netbox_collected_by_code or {}
        s6_collected = sum(
            netbox_collected_by_code.get(str(p.get("partner_code") or ""), 0)
            for p in by_state.get("S6", [])
        )
        s6_device_total = idle_total_s6 + s6_collected
        s6_at_pct = fmt_pct(idle_total_s6, s6_device_total) if s6_device_total else "0.0%"
        s6_col_pct = fmt_pct(s6_collected, s6_device_total) if s6_device_total else "0.0%"
        st.markdown(stage_card("STAGE 6  —  COMPLETE", STAGE_COLORS["S6"], [
            ("CSPs", s6_csps, fmt_pct(s6_csps, s1_csps)),
            ("Netbox at CSP", idle_total_s6, s6_at_pct),
            ("Total Netbox Collected from CSP", s6_collected, s6_col_pct),
        ]), unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — Data Quality (sheet vs Metabase reconciliation)
# ─────────────────────────────────────────────────────────────────────────────

def render_tab3_data_quality(partners, u1_by, u2_total, r15_by):
    """List CSPs in S4/S5 with 0 sheet data but R15>0 in Metabase."""
    st.markdown(
        '<div style="background:#9C0006;color:#ffffff;padding:10px 14px;'
        'font-weight:bold;font-size:14px;margin-top:6px;border-radius:6px">'
        '🚨 Sheet vs Metabase Reconciliation</div>',
        unsafe_allow_html=True,
    )
    st.caption("CSPs flagged here have **0 rows in the Google Sheet** but **active R15 customers in Metabase** — likely missing entries.")

    for stage in ["S4", "S5"]:
        stage_partners = [p for p in partners if p["current_state"] == stage]
        flagged = []
        for p in stage_partners:
            key = p["name"].lower()
            in_sheet = (u1_by.get(key, {}).get("total", 0) > 0) or (u2_total.get(key, 0) > 0)
            r15 = r15_by.get(p["name"], 0)
            if not in_sheet and r15 > 0:
                flagged.append({
                    "CSP Name": p["name"],
                    "Partner Code": p.get("partner_code", ""),
                    "Risk State": p.get("risk_state", ""),
                    "Active R15 customers": r15,
                })

        st.markdown(f"#### {stage} — {len(flagged)} CSPs missing from sheet")
        if flagged:
            df = pd.DataFrame(sorted(flagged, key=lambda x: -x["Active R15 customers"]))
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption(f"Hidden active customers in {stage}: **{sum(f['Active R15 customers'] for f in flagged):,}**")
        else:
            st.success(f"All {stage} CSPs accounted for in the sheet.")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 — Search (4 lookup boxes)
# ─────────────────────────────────────────────────────────────────────────────

def render_tab4_search(partners, u2_rows, secrets):
    """Four searchbars: devices, active userbase, U2 list, U1 list (placeholder)."""
    # Build select options: "Name (state · code)"
    options = sorted(
        [f"{p['name']}  —  {p['current_state']}  ·  {p.get('partner_code','')}" for p in partners]
    )
    name_to_partner = {f"{p['name']}  —  {p['current_state']}  ·  {p.get('partner_code','')}": p for p in partners}

    def partner_picker(key):
        choice = st.selectbox(
            "Type partner name to search:",
            options=[""] + options,
            key=key,
            help="Start typing — list filters as you type. Pick one to run the search.",
        )
        return name_to_partner.get(choice) if choice else None

    # ── Search 1: Devices at partner ─────────────────────────────────────────
    st.markdown("### 🔍 Search 1 — Devices at Partner (from Metabase)")
    p1 = partner_picker("s1_pick")
    if p1:
        st.write(f"**{p1['name']}** · Code: `{p1.get('partner_code','')}` · State: `{p1['current_state']}`")
        with st.spinner("Querying Metabase..."):
            res = search_devices_at_partner(secrets["metabase_url"], secrets["metabase_key"], p1["partner_code"])
        if res.get("error"):
            st.error(f"Metabase error: {res['error']}")
        elif res["rows"]:
            df = pd.DataFrame(res["rows"], columns=res["cols"])
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption(f"{len(df)} IDLE device(s).")
        else:
            st.info("No IDLE devices at this partner.")

    st.divider()

    # ── Search 2: Active userbase (R15) ──────────────────────────────────────
    st.markdown("### 🔍 Search 2 — Active Userbase (R15) at Partner (from Metabase)")
    p2 = partner_picker("s2_pick")
    if p2:
        st.write(f"**{p2['name']}** · Code: `{p2.get('partner_code','')}` · State: `{p2['current_state']}`")
        with st.spinner("Querying Metabase..."):
            res = search_active_userbase(secrets["metabase_url"], secrets["metabase_key"], p2["partner_code"])
        if res.get("error"):
            st.error(f"Metabase error: {res['error']}")
        elif res["rows"]:
            df = pd.DataFrame(res["rows"], columns=res["cols"])
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption(f"{len(df)} active R15 customer(s).")
        else:
            st.info("No active R15 customers at this partner.")

    st.divider()

    # ── Search 3: U2 customer list with collected filter ─────────────────────
    st.markdown("### 🔍 Search 3 — U2 Customers (from Google Sheet)")
    p3 = partner_picker("s3_pick")
    if p3:
        st.write(f"**{p3['name']}** · Code: `{p3.get('partner_code','')}` · State: `{p3['current_state']}`")
        f = st.radio("Filter:", ["All", "Device collected", "Device not collected"], horizontal=True, key="s3_filter")
        target = p3["name"].lower()
        rows = []
        for r in u2_rows:
            if str(r.get("Partner") or "").strip().lower() != target:
                continue
            if not (r.get("Mobile no") or r.get("Mobile")):
                continue
            remark = str(r.get("Remarks Dropdown") or "").strip()
            collected = remark.lower() == "device picked up"
            if f == "Device collected" and not collected: continue
            if f == "Device not collected" and collected: continue
            rows.append({
                "Cx Name": r.get("Cx Name", ""),
                "Mobile": r.get("Mobile no") or r.get("Mobile", ""),
                "Address": r.get("Address", ""),
                "Remarks Dropdown": remark or "(blank)",
            })
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            st.caption(f"{len(rows)} customer(s) matching filter.")
        else:
            st.info("No U2 customers for this partner under the current filter.")

    st.divider()

    # ── Search 4: U1 migration list (placeholder) ────────────────────────────
    st.markdown("### 🔍 Search 4 — U1 Migration Status (from new Google Sheet)")
    st.warning("📋 Data source not yet configured. Share the new Google Sheet tab name + columns and I'll wire it in.")
    st.caption("Will support partner selection + filter (Migrated / Not Migrated) once the source is added.")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN render() — composes all 4 tabs
# ─────────────────────────────────────────────────────────────────────────────

def render():
    secrets = get_secrets()

    with st.spinner("Fetching live data..."):
        u2_rows, u1_rows = fetch_sheets(secrets["sheet_id"], secrets["gcp_creds"])
        partners = fetch_partners(secrets["supabase_url"], secrets["supabase_key"])
        netbox_collected_by_code = fetch_netbox_collection(secrets["sheet_id"], secrets["gcp_creds"])

    u2 = classify_u2(u2_rows)
    u1 = classify_u1(u1_rows)
    u1_by, u2_total, u2_picked = build_sheet_lookups(u1_rows, u2_rows)

    # Use partner_code (= Metabase partner_account_id) as the unique join key.
    def code_of(p): return p.get("partner_code")

    all_codes = [code_of(p) for p in partners if code_of(p)]
    s5_all_codes = [code_of(p) for p in partners if p["current_state"] == "S5"]
    s5_all_lower = {p["name"].lower() for p in partners if p["current_state"] == "S5"}
    name_to_code = {p["name"]: code_of(p) for p in partners if code_of(p)}

    # Fetch R15 for ALL partners in exit pipeline (so S1 has a userbase denominator)
    r15_by_code = fetch_r15_active_by_code(
        secrets["metabase_url"], secrets["metabase_key"], all_codes,
    )
    # Also expose by name for data-quality flagging (Tab 3)
    code_to_name = {str(code_of(p)): p["name"] for p in partners if code_of(p)}
    r15_by = {code_to_name.get(code, code): cnt for code, cnt in r15_by_code.items()}
    idle_total = fetch_idle_devices_total(
        secrets["metabase_url"], secrets["metabase_key"], s5_all_codes,
    )
    # IDLE devices at CSPs currently in S6 (for the new S6 row)
    s6_codes = [code_of(p) for p in partners if p["current_state"] == "S6"]
    idle_total_s6 = fetch_idle_devices_total(
        secrets["metabase_url"], secrets["metabase_key"], s6_codes,
    ) if s6_codes else 0
    # Per-partner IDLE counts (for daily snapshot rows)
    idle_by_code = fetch_idle_devices_by_code(
        secrets["metabase_url"], secrets["metabase_key"], all_codes,
    ) if secrets.get("metabase_key") else {}

    # ── S5 dedup: U2 pending customers whose netbox is already IDLE at CSP ───
    s5_dedup = {"duplicates": 0}
    if s5_all_codes and secrets["metabase_key"]:
        # U2 pending mobiles (per partner) from Main sheet
        pending_pairs = []  # (mobile, partner_name)
        for r in u2_rows:
            partner = str(r.get("Partner") or "").strip()
            mobile = r.get("Mobile no") or r.get("Mobile")
            if not partner or not mobile: continue
            if partner.lower() not in s5_all_lower: continue
            if str(r.get("Remarks Dropdown") or "").strip().lower() == "device picked up": continue
            pending_pairs.append((str(mobile).strip(), partner))

        # Get NAS_ID per mobile + IDLE NAS sets per partner (keyed by partner_code)
        all_mobiles = list({m for m, _ in pending_pairs})
        mobile_to_nas = fetch_mobile_to_nas(
            secrets["metabase_url"], secrets["metabase_key"], all_mobiles,
        ) if all_mobiles else {}
        idle_nas_by_code = fetch_idle_nas_by_code(
            secrets["metabase_url"], secrets["metabase_key"], s5_all_codes,
        )

        # Count duplicates: pending customer's NAS_ID is in IDLE set at their partner
        dup = 0
        for mobile, partner in pending_pairs:
            nas = mobile_to_nas.get(mobile)
            code = name_to_code.get(partner)
            if not nas or not code: continue
            if nas in idle_nas_by_code.get(str(code), set()):
                dup += 1
        s5_dedup["duplicates"] = dup

    # Title bar
    st.markdown(
        '<div style="background:#1F4E78;color:#ffffff;padding:14px;border-radius:8px;'
        'font-size:17px;font-weight:bold;text-align:center;margin-bottom:6px">'
        'CSP EXIT TRACKER</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="updated">Live — auto-refreshes every 5 minutes</div>',
        unsafe_allow_html=True,
    )

    # ── Daily Totals capture + D-1 read (powers Tab 5) ────────────────────────
    today_metrics = compute_today_metrics(
        partners, u1_by, u2_total, u2_picked, r15_by_code,
        idle_total, s5_dedup, netbox_collected_by_code, idle_total_s6,
    )
    yest_totals = {}
    try:
        _creds = Credentials.from_service_account_info(
            secrets["gcp_creds"],
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive.readonly"],
        )
        book = gspread.authorize(_creds).open_by_key(secrets["sheet_id"])
        write_today_totals(book, today_metrics)
        yest_totals = read_yesterday_totals(book)
    except Exception:
        pass

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Status & Cohorts", "CSP Exit Funnel", "Data Quality", "Search",
        "Daily Funnel + Delta",
    ])
    with tab1:
        render_tab1_status(u1, u2)
    with tab2:
        render_tab2_funnel(
            partners, u1_by, u2_total, u2_picked, r15_by_code, idle_total,
            s5_dedup, idle_total_s6, netbox_collected_by_code,
        )
    with tab3:
        render_tab3_data_quality(partners, u1_by, u2_total, r15_by)
    with tab4:
        render_tab4_search(partners, u2_rows, secrets)
    with tab5:
        render_tab5_funnel_with_delta(today_metrics, yest_totals)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG + STYLES
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="CSP Exit Tracker", layout="centered")

st.markdown(
    """
<style>
[data-testid="stToolbar"] {visibility: hidden !important;}
[data-testid="stDecoration"] {display: none !important;}
[data-testid="stStatusWidget"] {display: none !important;}
[data-testid="stDeployButton"] {display: none !important;}
[data-testid="manage-app-button"] {display: none !important;}
.stDeployButton {display: none !important;}
.stAppDeployButton {display: none !important;}
a[href^="https://streamlit.io"] {display: none !important;}
footer {visibility: hidden !important;}
#MainMenu {visibility: hidden !important;}

.csp-table {
    width: 100%;
    border-collapse: collapse;
    font-family: Arial, sans-serif;
    font-size: 13px;
    margin-bottom: 6px;
}
.csp-table th {
    background: #2E75B6 !important;
    color: #ffffff !important;
    padding: 8px 12px;
    border: 1px solid #999;
    text-align: left;
    font-weight: bold;
}
.csp-table td {
    padding: 7px 12px;
    border: 1px solid #cccccc;
    color: #000000 !important;
    background: #ffffff;
}
.updated {
    font-size: 11px;
    color: #cccccc !important;
    text-align: right;
    margin-bottom: 6px;
}
/* Black page background; table cells keep their own explicit backgrounds */
.stApp { background: #000000 !important; }
.main .block-container { background: #000000 !important; }
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] div { color: inherit; }
</style>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# AUTH GATE
# ─────────────────────────────────────────────────────────────────────────────

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# Auto-restore session from URL token (survives reload / restart)
if not st.session_state.authenticated:
    _params = st.query_params
    _e = _params.get("e", "")
    _t = _params.get("t", "")
    if _e and _t:
        _correct_pw = _load_app_password()
        if _email_allowed(_e) and _t == _make_token(_e, _correct_pw):
            st.session_state.authenticated = True
            st.session_state.user_email = _e

if not st.session_state.authenticated:
    st.markdown(
        """
    <style>
    .login-title {
        background: #1F4E78; color: white; padding: 16px;
        border-radius: 8px; text-align: center;
        font-size: 18px; font-weight: bold; margin-bottom: 8px;
    }
    .login-sub { text-align: center; color: #cccccc; font-size: 13px; margin-bottom: 24px; }
    </style>
    <div class="login-title">CSP EXIT TRACKER</div>
    <div class="login-sub">Wiom Internal Dashboard — Restricted Access</div>
    """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
        email = st.text_input("Wiom Email", placeholder="name@wiom.in")
        password = st.text_input("Password", type="password")

        if st.button("Login", use_container_width=True):
            correct_pw = _load_app_password()
            clean_email = email.strip().lower()
            if not _email_allowed(clean_email):
                st.error("Access restricted to @wiom.in emails only.")
            elif password.strip() != correct_pw.strip():
                st.error("Incorrect password. Please try again.")
            else:
                st.session_state.authenticated = True
                st.session_state.user_email = clean_email
                st.query_params["e"] = clean_email
                st.query_params["t"] = _make_token(clean_email, correct_pw.strip())
                st.rerun()
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# TOP BAR (logged in)
# ─────────────────────────────────────────────────────────────────────────────

col_title, col_user, col_logout = st.columns([5, 2, 1])
with col_title:
    st.markdown(
        "<div style='font-size:20px;font-weight:bold;color:#5DADE2;padding-top:6px'>"
        "CSP Exit Tracker</div>",
        unsafe_allow_html=True,
    )
with col_user:
    st.markdown(
        f"<div style='font-size:12px;color:#cccccc;text-align:right;padding-top:10px'>"
        f"{st.session_state.get('user_email','')}</div>",
        unsafe_allow_html=True,
    )
with col_logout:
    if st.button("Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user_email = ""
        st.query_params.clear()
        st.rerun()

# Auto-refresh every 30 seconds
st_autorefresh(interval=300000)  # 5 minutes

render()
