"""
CSP Exit Tracker — Live Streamlit Dashboard

Live web view of the CSP exit funnel, sourced from the Google Sheet your team
updates daily. Mirrors the 5-table format of CSP_Exit_Tracker.xlsx exactly.

Auth: email + APP_PASSWORD (only @wiom.in emails can log in).
Data: pulled from Google Sheets every 30 seconds via service account.
"""

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from streamlit_autorefresh import st_autorefresh
from collections import Counter
from datetime import datetime
from zoneinfo import ZoneInfo
import hashlib
import json
import os

IST = ZoneInfo("Asia/Kolkata")

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

ALLOWED_DOMAINS = ("@wiom.in",)

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
        }


# ─────────────────────────────────────────────────────────────────────────────
# DATA FETCH (cached 30s — matches autorefresh interval)
# ─────────────────────────────────────────────────────────────────────────────

U2_TAB = "Main sheet"
U1_TAB = "Migration Data"


@st.cache_data(ttl=30)
def fetch_sheets(sheet_id, gcp_creds):
    """Pull the two relevant tabs by name and return rows as list-of-dicts."""
    creds = Credentials.from_service_account_info(
        gcp_creds,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets.readonly",
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
        mobile = row.get("Mobile")
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


def render():
    secrets = get_secrets()

    with st.spinner("Fetching live data from Google Sheet..."):
        u2_rows, u1_rows = fetch_sheets(secrets["sheet_id"], secrets["gcp_creds"])

    u2 = classify_u2(u2_rows)
    u1 = classify_u1(u1_rows)

    grand_total = u1["total"] + u2["total"]
    u1_pending = u1["not_migrated"] + u1["in_process"]
    u2_not_team = u2["partner"] + u2["swu"]

    updated = datetime.now(IST).strftime("%d-%b-%Y %H:%M")

    # ── Title bar ────────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:#1F4E78;color:#ffffff;padding:14px;border-radius:8px;'
        f'font-size:17px;font-weight:bold;text-align:center;margin-bottom:6px">'
        f'CSP EXIT STATUS &nbsp;|&nbsp; Snapshot as of {updated}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="updated">Live — auto-refreshes every 30 seconds</div>',
        unsafe_allow_html=True,
    )

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
st_autorefresh(interval=30000)

render()
