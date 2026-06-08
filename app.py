"""
NICHE 2026 — Self Check-In System  (Streamlit)
==============================================
Single-file Streamlit app for the International Halal Conference.

Flow:
  • 3 Tentative views   — Academic (with abstracts), Industry, Gala Dinner
  • Public registration — by email (or walk-in via Admin)
  • SELF check-in       — participant enters email, ticks own attendance,
                           door gift, dinner confirmation, sees assigned table
  • Admin               — oversight, walk-in registration, table assignment
                           (tables 3, 4, 8, 9, 10, 27 · 10 seats each)

Deploy on Streamlit Cloud:
  1. Push this repo to GitHub  (must include niche_data.xlsx + requirements.txt)
  2. Streamlit Cloud → New app  → main file: app.py
"""
from __future__ import annotations

import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================================
# CONFIG
# ============================================================================
BASE_DIR   = Path(__file__).parent
DB_PATH    = BASE_DIR / "niche.db"
EXCEL_PATH = BASE_DIR / "niche_data.xlsx"

PARTICIPANT_TABLES = [3, 4, 8, 9, 10, 27]
SEATS_PER_TABLE    = 10
ADMIN_PASSWORD     = "NICHE2026admin"

st.set_page_config(
    page_title="NICHE 2026 · Self Check-In",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={"About": "NICHE 2026 — International Halal Conference"},
)


# ============================================================================
# THEME — inject all custom CSS + Google Fonts
# ============================================================================
def inject_theme():
    st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800;900&family=Inter:wght@400;500;600;700&family=Dancing+Script:wght@600;700&display=swap" rel="stylesheet">

<style>
:root {
  --navy-900: #050a2e;
  --navy-800: #0a1140;
  --navy-700: #0d1654;
  --navy-600: #131c6b;
  --gold-700: #8b6914;
  --gold-600: #b8860b;
  --gold-500: #d4af37;
  --gold-400: #e6c64f;
  --gold-300: #f4d469;
  --gold-200: #ffe88a;
  --ink: #f2efe4;
  --ink-soft: #d9d5c2;
  --ink-muted: #9590a3;
}

/* === Hide default Streamlit chrome === */
[data-testid="stSidebar"], [data-testid="stSidebarNav"],
[data-testid="collapsedControl"],
header [data-testid="stToolbar"], header [data-testid="stDecoration"],
#MainMenu, footer, .stDeployButton, [data-testid="stStatusWidget"] {
  display: none !important;
}
[data-testid="stHeader"] { background: transparent !important; }

/* === Global background === */
.stApp {
  background:
    radial-gradient(ellipse at top right, rgba(212,175,55,0.12), transparent 55%),
    radial-gradient(ellipse at bottom left, rgba(40,53,147,0.55), transparent 60%),
    linear-gradient(180deg, var(--navy-900) 0%, var(--navy-800) 60%, var(--navy-900) 100%);
  background-attachment: fixed;
  color: var(--ink);
}

/* Star field overlay */
.stApp::before {
  content: '';
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background-image:
    radial-gradient(1px 1px at 10% 20%, rgba(255,255,255,0.6) 50%, transparent 50%),
    radial-gradient(1px 1px at 25% 70%, rgba(255,255,255,0.4) 50%, transparent 50%),
    radial-gradient(1px 1px at 60% 30%, rgba(212,175,55,0.6) 50%, transparent 50%),
    radial-gradient(1.5px 1.5px at 80% 80%, rgba(255,255,255,0.5) 50%, transparent 50%),
    radial-gradient(1px 1px at 45% 50%, rgba(212,175,55,0.4) 50%, transparent 50%),
    radial-gradient(1px 1px at 90% 15%, rgba(255,255,255,0.5) 50%, transparent 50%);
  background-size: 400px 400px;
  animation: twinkle 6s ease-in-out infinite alternate;
}
@keyframes twinkle { from {opacity:0.4;} to {opacity:1;} }

.main .block-container {
  padding-top: 1rem; max-width: 1400px;
  position: relative; z-index: 1;
}

/* === Typography === */
html, body, [class*="css"], .stApp, .stMarkdown, p, div, span, label {
  font-family: 'Inter', -apple-system, sans-serif !important;
}
h1, h2, h3 {
  font-family: 'Playfair Display', Georgia, serif !important;
  font-weight: 800 !important;
  letter-spacing: 0.5px;
  color: var(--gold-200) !important;
}

/* === Top brand bar === */
.brand-bar {
  display: flex; align-items: center; justify-content: space-between;
  padding: 14px 22px; margin-bottom: 18px;
  backdrop-filter: blur(20px);
  background: rgba(5,10,46,0.6);
  border: 1px solid rgba(212,175,55,0.25);
  border-radius: 14px;
}
.brand-mark {
  font-family: 'Playfair Display', serif;
  font-size: 28px; font-weight: 900;
  background: linear-gradient(135deg, var(--gold-300), var(--gold-500), var(--gold-700));
  -webkit-background-clip: text; background-clip: text; color: transparent;
  letter-spacing: 2px;
}
.brand-sub {
  font-size: 11px; letter-spacing: 3px; text-transform: uppercase;
  color: var(--ink-muted); padding-left: 14px;
  border-left: 1px solid rgba(212,175,55,0.35);
}

/* === Tabs (top navigation, replaces sidebar) === */
.stTabs [data-baseweb="tab-list"] {
  background: rgba(5,10,46,0.55);
  border: 1px solid rgba(212,175,55,0.22);
  border-radius: 14px;
  padding: 6px;
  gap: 4px;
  backdrop-filter: blur(12px);
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  border-radius: 10px !important;
  padding: 10px 22px !important;
  color: var(--ink-soft) !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  letter-spacing: 0.5px;
  border: none !important;
  transition: all 0.2s ease;
}
.stTabs [data-baseweb="tab"]:hover {
  background: rgba(212,175,55,0.08) !important;
  color: var(--gold-300) !important;
}
.stTabs [aria-selected="true"] {
  background: linear-gradient(135deg, var(--gold-500), var(--gold-700)) !important;
  color: var(--navy-900) !important;
  box-shadow: 0 6px 18px rgba(212,175,55,0.35);
}
.stTabs [data-baseweb="tab-highlight"] { background: transparent !important; }
.stTabs [data-baseweb="tab-border"]    { background: transparent !important; }

/* === Hero === */
.hero {
  text-align: center; padding: 40px 20px 30px;
}
.hero .script {
  font-family: 'Dancing Script', cursive;
  font-size: 28px; color: var(--gold-300);
  display: block;
}
.hero h1 {
  font-size: clamp(2.4rem, 5vw, 4rem) !important;
  margin: 4px 0 !important;
  background: linear-gradient(135deg, var(--gold-200) 0%, var(--gold-500) 60%, var(--gold-700) 100%);
  -webkit-background-clip: text; background-clip: text; color: transparent !important;
}
.divider {
  width: 120px; height: 2px; margin: 18px auto;
  background: linear-gradient(90deg, transparent, var(--gold-500), transparent);
  position: relative;
}
.divider::after {
  content: '✦';
  position: absolute; top: -10px; left: 50%; transform: translateX(-50%);
  color: var(--gold-400); background: var(--navy-900);
  padding: 0 10px; font-size: 14px;
}
.hero .lede {
  color: var(--ink-soft); max-width: 700px; margin: 16px auto 0;
  font-size: 16px; line-height: 1.6;
}
.hero .when {
  color: var(--gold-300); letter-spacing: 2px; font-size: 12.5px;
  margin-top: 10px;
}

/* === Cards (HTML) === */
.t-card {
  position: relative;
  padding: 32px 26px;
  background: linear-gradient(160deg, rgba(26,35,126,0.6), rgba(13,22,84,0.7));
  border: 1px solid rgba(212,175,55,0.3);
  border-radius: 18px;
  height: 100%;
  box-shadow: 0 20px 60px rgba(0,0,0,0.4);
  transition: transform 0.3s, box-shadow 0.3s, border-color 0.3s;
}
.t-card::before, .t-card::after {
  content: ''; position: absolute; width: 24px; height: 24px;
  border: 2px solid var(--gold-500); opacity: 0.7;
}
.t-card::before { top: 10px; left: 10px; border-right: 0; border-bottom: 0; border-top-left-radius: 6px; }
.t-card::after  { bottom: 10px; right: 10px; border-left: 0; border-top: 0; border-bottom-right-radius: 6px; }
.t-card:hover { transform: translateY(-6px); border-color: var(--gold-400);
  box-shadow: 0 24px 60px rgba(212,175,55,0.25); }

.t-card .icon {
  width: 60px; height: 60px;
  display: grid; place-items: center;
  border-radius: 50%;
  background: linear-gradient(135deg, var(--gold-500), var(--gold-700));
  font-size: 26px; margin-bottom: 18px;
  box-shadow: 0 10px 28px rgba(212,175,55,0.35);
}
.t-card .script { font-family: 'Dancing Script', cursive; color: var(--gold-300); font-size: 20px; }
.t-card h3 {
  font-size: 24px !important; margin: 4px 0 6px !important;
  background: linear-gradient(135deg, var(--gold-200), var(--gold-500));
  -webkit-background-clip: text; background-clip: text; color: transparent !important;
}
.t-card .when {
  display: inline-block; font-size: 11px;
  letter-spacing: 1.5px; text-transform: uppercase;
  padding: 4px 10px; border-radius: 999px;
  background: rgba(212,175,55,0.15); color: var(--gold-200);
  margin-bottom: 16px;
}
.t-card ul { list-style: none; padding: 0; margin: 0 0 18px; }
.t-card li {
  padding: 5px 0; font-size: 14px; color: var(--ink-soft);
  display: flex; gap: 8px;
}
.t-card li::before { content: '✓'; color: var(--gold-400); font-weight: 700; }

/* === Stat pills === */
.stat-pill {
  text-align: center; padding: 22px;
  background: rgba(255,255,255,0.03);
  border: 1px solid rgba(212,175,55,0.22);
  border-radius: 14px;
}
.stat-pill .num {
  font-family: 'Playfair Display', serif;
  font-size: 40px; font-weight: 900; line-height: 1;
  background: linear-gradient(135deg, var(--gold-300), var(--gold-600));
  -webkit-background-clip: text; background-clip: text; color: transparent;
}
.stat-pill .lbl {
  font-size: 11px; letter-spacing: 2px; text-transform: uppercase;
  color: var(--ink-muted); margin-top: 6px;
}

/* === Programme rows === */
.prog-row {
  display: grid; grid-template-columns: 160px 1fr;
  gap: 22px; padding: 18px 20px;
  background: rgba(5,10,46,0.4);
  border: 1px solid rgba(212,175,55,0.18);
  border-radius: 12px; margin-bottom: 10px;
}
.prog-row:hover { border-color: rgba(212,175,55,0.4); background: rgba(5,10,46,0.6); }
.prog-time {
  font-family: 'Playfair Display', serif; font-weight: 700;
  font-size: 14px; color: var(--gold-300);
}
.prog-label {
  font-size: 10.5px; letter-spacing: 2px; text-transform: uppercase;
  color: var(--ink-muted); margin-bottom: 4px;
}
.prog-title { font-size: 14.5px; font-weight: 600; color: var(--ink); line-height: 1.4; }
.prog-meta  { font-size: 12.5px; color: var(--ink-muted); margin-top: 6px; }

.session-h {
  font-family: 'Playfair Display', serif;
  font-size: 22px; color: var(--gold-300) !important;
  border-bottom: 1px solid rgba(212,175,55,0.3);
  padding-bottom: 10px; margin: 28px 0 16px;
  display: flex; align-items: center; gap: 12px;
}
.session-h .badge {
  font-family: 'Inter', sans-serif; font-size: 10px;
  background: linear-gradient(135deg, var(--gold-500), var(--gold-700));
  color: var(--navy-900); padding: 3px 10px; border-radius: 999px;
  letter-spacing: 1px; font-weight: 700;
}

/* === Abstract panel === */
.abs-panel {
  background: linear-gradient(160deg, rgba(26,35,126,0.55), rgba(13,22,84,0.7));
  border: 1px solid var(--gold-500); border-radius: 16px;
  padding: 24px 28px; margin: 14px 0;
}
.abs-panel .pid {
  display: inline-block; font-size: 11px; letter-spacing: 2px;
  background: linear-gradient(135deg, var(--gold-400), var(--gold-600));
  color: var(--navy-900); padding: 4px 10px; border-radius: 999px;
  font-weight: 700; margin-bottom: 10px;
}
.abs-panel h4 {
  font-family: 'Playfair Display', serif !important;
  color: var(--gold-200) !important; margin: 4px 0 16px !important;
  font-size: 18px !important;
}
.abs-field-lbl {
  font-size: 11px; letter-spacing: 2px; text-transform: uppercase;
  color: var(--gold-400); font-weight: 600; margin: 14px 0 6px;
}
.abs-field-val { color: var(--ink-soft); font-size: 14px; line-height: 1.7; }
.abs-body {
  background: rgba(5,10,46,0.5); padding: 16px;
  border-radius: 8px; border-left: 3px solid var(--gold-500);
  text-align: justify; white-space: pre-line;
}
.kw {
  display: inline-block;
  background: rgba(212,175,55,0.12);
  border: 1px solid rgba(212,175,55,0.3);
  color: var(--gold-200);
  padding: 3px 11px; border-radius: 999px;
  font-size: 12px; margin: 3px 4px 3px 0;
}

/* === Status badges === */
.status-badge {
  display: inline-block; padding: 5px 12px; border-radius: 999px;
  font-size: 11px; font-weight: 700; letter-spacing: 1px;
  text-transform: uppercase;
}
.status-on  { background: rgba(46,204,113,0.18); color: #7ee2a5;
              border: 1px solid rgba(46,204,113,0.4); }
.status-off { background: rgba(255,255,255,0.04); color: var(--ink-muted);
              border: 1px solid rgba(255,255,255,0.12); }

/* === Forms & inputs === */
.stTextInput input, .stSelectbox > div > div, .stTextArea textarea {
  background: rgba(5,10,46,0.6) !important;
  border: 1px solid rgba(212,175,55,0.25) !important;
  color: var(--ink) !important;
  border-radius: 10px !important;
}
.stTextInput input:focus, .stSelectbox > div > div:focus-within {
  border-color: var(--gold-400) !important;
  box-shadow: 0 0 0 3px rgba(212,175,55,0.18) !important;
}
.stTextInput label, .stSelectbox label, .stCheckbox label,
.stRadio label, .stTextArea label, .stNumberInput label {
  color: var(--gold-300) !important;
  font-size: 12px !important; font-weight: 600 !important;
  letter-spacing: 1.5px !important; text-transform: uppercase !important;
}

/* === Buttons === */
.stButton > button {
  background: linear-gradient(135deg, var(--gold-400), var(--gold-600)) !important;
  color: var(--navy-900) !important;
  border: 0 !important;
  font-weight: 700 !important;
  letter-spacing: 0.5px;
  padding: 10px 22px !important;
  border-radius: 10px !important;
  box-shadow: 0 8px 24px rgba(212,175,55,0.25);
  transition: all 0.2s ease;
}
.stButton > button:hover {
  transform: translateY(-2px);
  box-shadow: 0 12px 32px rgba(212,175,55,0.4);
  color: var(--navy-900) !important;
}
.stButton > button:focus { color: var(--navy-900) !important; }
.stButton > button[kind="secondary"] {
  background: transparent !important;
  color: var(--gold-300) !important;
  border: 1px solid rgba(212,175,55,0.4) !important;
  box-shadow: none;
}
.stFormSubmitButton > button {
  background: linear-gradient(135deg, var(--gold-400), var(--gold-600)) !important;
  color: var(--navy-900) !important; border: 0 !important;
  font-weight: 700 !important;
  padding: 12px 26px !important; border-radius: 10px !important;
}

/* === Alerts === */
.stAlert {
  border-radius: 12px !important;
  backdrop-filter: blur(8px);
}

/* === Tables (data editor / dataframe) === */
[data-testid="stDataFrame"], .stDataFrame {
  border: 1px solid rgba(212,175,55,0.22);
  border-radius: 12px; overflow: hidden;
}

/* === Expanders === */
[data-testid="stExpander"] details {
  background: rgba(5,10,46,0.5) !important;
  border: 1px solid rgba(212,175,55,0.2) !important;
  border-radius: 12px !important;
}
[data-testid="stExpander"] summary {
  font-weight: 600; color: var(--ink) !important;
  padding: 14px 18px !important;
}
[data-testid="stExpander"] summary:hover { color: var(--gold-300) !important; }

/* === Metrics === */
[data-testid="stMetric"] {
  background: rgba(255,255,255,0.03);
  border: 1px solid rgba(212,175,55,0.22);
  border-radius: 14px; padding: 18px;
}
[data-testid="stMetricValue"] {
  font-family: 'Playfair Display', serif !important;
  color: var(--gold-300) !important;
}
[data-testid="stMetricLabel"] {
  color: var(--ink-muted) !important;
  font-size: 11px !important; letter-spacing: 2px;
  text-transform: uppercase;
}

/* === Table chip strip === */
.tchip {
  text-align: center; padding: 14px 8px;
  background: rgba(5,10,46,0.6);
  border: 1px solid rgba(212,175,55,0.25);
  border-radius: 12px;
}
.tchip.full { border-color: #e74c3c; }
.tchip .num {
  font-family: 'Playfair Display', serif;
  font-size: 26px; font-weight: 800; color: var(--gold-300);
}
.tchip .lbl {
  font-size: 9px; letter-spacing: 1.5px; text-transform: uppercase;
  color: var(--ink-muted);
}
.tchip .cap {
  margin-top: 6px; height: 5px;
  background: rgba(255,255,255,0.06); border-radius: 3px; overflow: hidden;
}
.tchip .cap-fill {
  height: 100%;
  background: linear-gradient(90deg, var(--gold-500), var(--gold-300));
}
.tchip.full .cap-fill { background: #e74c3c; }

/* My-record card on Check-In page */
.me-card {
  background: linear-gradient(160deg, rgba(26,35,126,0.7), rgba(13,22,84,0.85));
  border: 1px solid var(--gold-500); border-radius: 18px;
  padding: 28px 32px; margin: 14px 0;
  box-shadow: 0 20px 60px rgba(0,0,0,0.4);
}
.me-card .name {
  font-family: 'Playfair Display', serif;
  font-size: 28px; color: var(--gold-200);
  margin-bottom: 4px;
}
.me-card .email {
  color: var(--ink-muted); font-family: ui-monospace, monospace;
  font-size: 13px;
}
.me-card .org {
  color: var(--ink-soft); font-size: 14px; margin-top: 6px;
}
.me-card .table-banner {
  margin-top: 18px; padding: 18px;
  background: linear-gradient(135deg, var(--gold-500), var(--gold-700));
  color: var(--navy-900); border-radius: 12px;
  text-align: center;
}
.me-card .table-banner .lbl {
  font-size: 11px; letter-spacing: 2px; font-weight: 700;
  text-transform: uppercase; opacity: 0.85;
}
.me-card .table-banner .num {
  font-family: 'Playfair Display', serif;
  font-size: 42px; font-weight: 900;
}
</style>
""", unsafe_allow_html=True)


# ============================================================================
# DATABASE
# ============================================================================
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _clean(v):
    if v is None: return None
    if isinstance(v, float) and pd.isna(v): return None
    s = str(v).strip()
    return None if s.lower() in ("nan", "nat", "") else s


def _yes(v):
    return 1 if str(v).strip().lower() in ("yes", "y", "true", "1") else 0


def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS participants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        full_name TEXT NOT NULL,
        phone TEXT, organisation TEXT, category TEXT,
        academic INTEGER DEFAULT 0, industry INTEGER DEFAULT 0,
        conference_checkin INTEGER DEFAULT 0,
        doorgift_collected INTEGER DEFAULT 0,
        attend_dinner      INTEGER DEFAULT 0,
        dinner_checkin     INTEGER DEFAULT 0,
        table_number       INTEGER,
        registration_source TEXT DEFAULT 'Preloaded',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS academic_programme (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        venue TEXT, time TEXT, session TEXT, moderator TEXT, theme TEXT,
        paper_id TEXT, title TEXT, presenter TEXT, email TEXT, sort_order INTEGER
    );
    CREATE TABLE IF NOT EXISTS abstracts (
        paper_id TEXT PRIMARY KEY, title TEXT, presenter TEXT, email TEXT,
        venue TEXT, time TEXT, session TEXT, keywords TEXT,
        abstract_text TEXT, authors TEXT
    );
    CREATE TABLE IF NOT EXISTS industry_programme (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        day TEXT, time TEXT, venue TEXT, session TEXT,
        speaker TEXT, organisation TEXT, details TEXT, sort_order INTEGER
    );
    CREATE TABLE IF NOT EXISTS dinner_programme (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        time TEXT, event TEXT, sort_order INTEGER
    );
    """)
    conn.commit()

    if c.execute("SELECT COUNT(*) FROM participants").fetchone()[0] == 0 \
       and EXCEL_PATH.exists():
        seed_from_excel(conn)
    conn.close()


def seed_from_excel(conn):
    xl = pd.ExcelFile(EXCEL_PATH)
    c = conn.cursor()

    if "Participants" in xl.sheet_names:
        for _, r in pd.read_excel(xl, "Participants").iterrows():
            email = _clean(r.get("Email"))
            if not email: continue
            c.execute("""INSERT OR IGNORE INTO participants
                (email, full_name, phone, organisation, category,
                 academic, industry, conference_checkin, doorgift_collected,
                 attend_dinner, dinner_checkin, table_number, registration_source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (email, _clean(r.get("Full_Name")) or "—",
                 _clean(r.get("Phone")), _clean(r.get("Organisation")),
                 _clean(r.get("Category")),
                 _yes(r.get("Academic")), _yes(r.get("Industry")),
                 _yes(r.get("Conference_CheckIn")),
                 _yes(r.get("DoorGift_Collected")),
                 _yes(r.get("Attend_Dinner")),
                 _yes(r.get("Dinner_CheckIn")),
                 int(r["Table_Number"]) if pd.notna(r.get("Table_Number")) else None,
                 _clean(r.get("Registration_Source")) or "Preloaded"))

    if "Academic_Programme" in xl.sheet_names:
        for i, r in pd.read_excel(xl, "Academic_Programme").iterrows():
            c.execute("""INSERT INTO academic_programme
                (venue, time, session, moderator, theme, paper_id, title,
                 presenter, email, sort_order) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (_clean(r.get("Venue")), _clean(r.get("Time")),
                 _clean(r.get("Session")), _clean(r.get("Moderator")),
                 _clean(r.get("Theme")), _clean(r.get("Paper_ID")),
                 _clean(r.get("Title")), _clean(r.get("Presenter")),
                 _clean(r.get("Email_From_Abstract")), int(i)))

    if "Abstracts" in xl.sheet_names:
        for _, r in pd.read_excel(xl, "Abstracts").iterrows():
            pid = _clean(r.get("Paper_ID"))
            if not pid: continue
            c.execute("""INSERT OR REPLACE INTO abstracts
                (paper_id, title, presenter, email, venue, time, session,
                 keywords, abstract_text, authors) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (pid, _clean(r.get("Title")), _clean(r.get("Presenter")),
                 _clean(r.get("Email")), _clean(r.get("Venue")),
                 _clean(r.get("Time")), _clean(r.get("Session")),
                 _clean(r.get("Keywords")), _clean(r.get("Abstract_Text")),
                 _clean(r.get("Authors_Affiliation"))))

    if "Industry_Programme" in xl.sheet_names:
        for i, r in pd.read_excel(xl, "Industry_Programme").iterrows():
            c.execute("""INSERT INTO industry_programme
                (day, time, venue, session, speaker, organisation, details, sort_order)
                VALUES (?,?,?,?,?,?,?,?)""",
                (_clean(r.get("Day")), _clean(r.get("Time")),
                 _clean(r.get("Venue")), _clean(r.get("Session")),
                 _clean(r.get("Speaker")), _clean(r.get("Organisation")),
                 _clean(r.get("Details")), int(i)))

    if "Gala_Dinner_Programme" in xl.sheet_names:
        for i, r in pd.read_excel(xl, "Gala_Dinner_Programme").iterrows():
            c.execute("INSERT INTO dinner_programme (time, event, sort_order) VALUES (?,?,?)",
                (_clean(r.get("Time")), _clean(r.get("Event")), int(i)))
    conn.commit()


# ============================================================================
# DATA HELPERS
# ============================================================================
def fetch_df(query, params=()):
    conn = get_conn()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def s(v, default="—"):
    """Safe string: convert NaN/None to a default."""
    if v is None: return default
    if isinstance(v, float) and pd.isna(v): return default
    txt = str(v).strip()
    return txt if txt and txt.lower() not in ("nan", "nat", "none") else default


def fetch_one(query, params=()):
    conn = get_conn()
    r = conn.execute(query, params).fetchone()
    conn.close()
    return dict(r) if r else None


def execute(query, params=()):
    conn = get_conn()
    conn.execute(query, params)
    conn.commit()
    conn.close()


def table_occupancy():
    df = fetch_df("SELECT table_number, COUNT(*) c FROM participants "
                  "WHERE table_number IS NOT NULL GROUP BY table_number")
    occ = {t: 0 for t in PARTICIPANT_TABLES}
    for _, r in df.iterrows():
        if int(r["table_number"]) in occ:
            occ[int(r["table_number"])] = int(r["c"])
    return occ


# ============================================================================
# UI COMPONENTS
# ============================================================================
def brand_bar():
    st.markdown("""
    <div class="brand-bar">
      <div style="display:flex; align-items:center;">
        <span class="brand-mark">NICHE</span>
        <span class="brand-sub">2026 · Royale Chulan Seremban</span>
      </div>
      <div style="font-size:11px; letter-spacing:2px; color:var(--ink-muted);">
        9 – 10 JUNE 2026
      </div>
    </div>
    """, unsafe_allow_html=True)


def hero(script, title, lede, when=None):
    when_html = f'<div class="when">{when}</div>' if when else ""
    st.markdown(f"""
    <div class="hero">
      <span class="script">{script}</span>
      <h1>{title}</h1>
      <div class="divider"></div>
      <p class="lede">{lede}</p>
      {when_html}
    </div>
    """, unsafe_allow_html=True)


def programme_row(time, label, title, meta=""):
    time_s  = s(time, "—")
    title_s = s(title, "—")
    label_s = s(label, "")
    meta_s  = s(meta, "")
    st.markdown(f"""
    <div class="prog-row">
      <div class="prog-time">{time_s}</div>
      <div>
        {f'<div class="prog-label">{label_s}</div>' if label_s and label_s != "—" else ''}
        <div class="prog-title">{title_s}</div>
        {f'<div class="prog-meta">{meta_s}</div>' if meta_s and meta_s != "—" else ''}
      </div>
    </div>
    """, unsafe_allow_html=True)


# ============================================================================
# PAGE: HOME
# ============================================================================
def page_home():
    hero("Welcome to", "NICHE 2026",
         "The International Halal Conference bringing together "
         "industry leaders, academics, policymakers and global partners. "
         "Two days of insights, networking, and breakthrough conversations.",
         "9 – 10 JUNE 2026 · ROYALE CHULAN SEREMBAN")

    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        st.markdown("""
        <div class="t-card">
          <div class="icon">🏛</div>
          <span class="script">Industry</span>
          <h3>Industrial Conference</h3>
          <div class="when">9 – 10 June · Grand Ballroom</div>
          <ul>
            <li>Industry-focused sessions</li>
            <li>Keynotes, panels &amp; forums</li>
            <li>Business matching &amp; networking</li>
            <li>For owners, manufacturers, exporters</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="t-card">
          <div class="icon">🎓</div>
          <span class="script">Academic</span>
          <h3>Academic Conference</h3>
          <div class="when">10 June · 1st Floor Meeting Room</div>
          <ul>
            <li>Research paper presentations</li>
            <li>Full abstracts &amp; detailed schedule</li>
            <li>Academic discussions &amp; insights</li>
            <li>For researchers &amp; academicians</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="t-card">
          <div class="icon">✨</div>
          <span class="script">Exclusive</span>
          <h3>Gala Dinner</h3>
          <div class="when">9 June · 6:30 PM Onwards</div>
          <ul>
            <li>Official Launching Ceremony</li>
            <li>Cultural Performance &amp; Lucky Draw</li>
            <li>Networking with VVIPs &amp; Leaders</li>
            <li>Limited seats — by invitation</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

    # Stats
    st.markdown("<div style='height: 38px;'></div>", unsafe_allow_html=True)
    n_part = fetch_one("SELECT COUNT(*) c FROM participants")["c"]
    n_paper = fetch_one("SELECT COUNT(*) c FROM abstracts")["c"]
    n_sess = fetch_one("SELECT COUNT(*) c FROM industry_programme "
                       "WHERE speaker IS NOT NULL")["c"]
    s1, s2, s3 = st.columns(3, gap="medium")
    for col, num, lbl in [(s1, n_part, "Registered Participants"),
                          (s2, n_paper, "Research Papers"),
                          (s3, n_sess, "Industry Sessions")]:
        with col:
            st.markdown(f"""
            <div class="stat-pill">
              <div class="num">{num}</div>
              <div class="lbl">{lbl}</div>
            </div>
            """, unsafe_allow_html=True)


# ============================================================================
# PAGE: ACADEMIC
# ============================================================================
def page_academic():
    hero("Academic", "Academic Tentative",
         "Parallel sessions of refereed research presentations. "
         "Expand any paper to read the full abstract, keywords, and authors.",
         "10 JUNE 2026 · 1ST FLOOR MEETING ROOM · PARALLEL SESSIONS")

    prog = fetch_df("SELECT * FROM academic_programme ORDER BY sort_order")
    abs_df = fetch_df("SELECT * FROM abstracts")
    abs_map = {r["paper_id"]: r for _, r in abs_df.iterrows()}

    for venue in prog["venue"].dropna().unique():
        rows = prog[prog["venue"] == venue]
        st.markdown(
            f'<div class="session-h">{venue}'
            f'<span class="badge">{len(rows)} ITEMS</span></div>',
            unsafe_allow_html=True
        )

        for _, r in rows.iterrows():
            label = " · ".join([str(x) for x in [r.get("session"), r.get("theme")]
                                if pd.notna(x) and str(x).strip()])
            meta_parts = []
            if pd.notna(r.get("presenter")): meta_parts.append(f"👤 {r['presenter']}")
            if pd.notna(r.get("moderator")) and pd.isna(r.get("presenter")):
                meta_parts.append(f"🎤 Mod: {r['moderator']}")
            if pd.notna(r.get("paper_id")):  meta_parts.append(f"📄 ID {r['paper_id']}")
            meta = " &nbsp;·&nbsp; ".join(meta_parts)

            programme_row(r.get("time"), label, r.get("title"), meta)

            pid = r.get("paper_id")
            if pid and pid in abs_map:
                a = abs_map[pid]
                with st.expander(f"📖  Read Full Abstract — {pid}"):
                    kw_html = ""
                    if pd.notna(a["keywords"]):
                        kws = str(a["keywords"]).replace("[","").replace("]","")
                        for k in [x.strip() for x in kws.replace(";",",").split(",") if x.strip()]:
                            kw_html += f'<span class="kw">{k}</span>'

                    authors_html = (
                        f'<div class="abs-field-lbl">Authors &amp; Affiliations</div>'
                        f'<div class="abs-field-val" style="white-space:pre-line;">{s(a["authors"])}</div>'
                        if pd.notna(a["authors"]) else ""
                    )

                    st.markdown(f"""
                    <div class="abs-panel">
                      <span class="pid">Paper {pid}</span>
                      <h4>{s(a['title'], '—')}</h4>
                      <div style="color:var(--ink-muted); font-size:13px;">
                        {s(a['presenter'], '')} · {s(a['session'], '')}<br>
                        {s(a['venue'], '')} · {s(a['time'], '')}
                      </div>
                      {f'<div class="abs-field-lbl">Keywords</div><div>{kw_html}</div>' if kw_html else ''}
                      <div class="abs-field-lbl">Abstract</div>
                      <div class="abs-field-val abs-body">{s(a['abstract_text'], '—')}</div>
                      {authors_html}
                    </div>
                    """, unsafe_allow_html=True)


# ============================================================================
# PAGE: INDUSTRY
# ============================================================================
def page_industry():
    hero("Industry", "Industrial Tentative",
         "Two days of keynotes, special addresses, panel discussions and industry "
         "sessions bringing together regulators, manufacturers, financiers and "
         "global thought leaders.",
         "9 – 10 JUNE 2026 · GRAND BALLROOM")

    prog = fetch_df("SELECT * FROM industry_programme ORDER BY sort_order")
    for day in prog["day"].dropna().unique():
        rows = prog[prog["day"] == day]
        st.markdown(
            f'<div class="session-h">{day}'
            f'<span class="badge">{len(rows)} SESSIONS</span></div>',
            unsafe_allow_html=True
        )
        for _, r in rows.iterrows():
            meta_parts = []
            if pd.notna(r.get("speaker")):  meta_parts.append(f"🎤 {r['speaker']}")
            if pd.notna(r.get("details")):  meta_parts.append(f"📌 {r['details']}")
            org = s(r.get("organisation"), "")
            programme_row(r.get("time"), org if org != "—" else "",
                          r.get("session"),
                          " &nbsp;·&nbsp; ".join(meta_parts))


# ============================================================================
# PAGE: DINNER
# ============================================================================
def page_dinner():
    hero("Exclusive", "Gala Dinner",
         "An evening bringing together VVIPs, policymakers, industry leaders, "
         "delegates and global partners. Official launching, cultural performance, "
         "networking &amp; more.",
         "9 JUNE 2026 · 4:30 PM ONWARDS · LIMITED SEATS")

    prog = fetch_df("SELECT * FROM dinner_programme ORDER BY sort_order")
    st.markdown(
        f'<div class="session-h">Evening Programme'
        f'<span class="badge">{len(prog)} ITEMS</span></div>',
        unsafe_allow_html=True
    )
    for _, r in prog.iterrows():
        programme_row(r.get("time"), "", r.get("event"))

    st.markdown("""
    <div class="abs-panel" style="text-align:center; margin-top:30px;">
      <h4 style="font-size:20px !important;">✨ Confirm Your Seat at the Gala Dinner</h4>
      <p style="color:var(--ink-soft); max-width:600px; margin:8px auto 0;">
        Head to the <strong style="color:var(--gold-300);">Check-In</strong> tab,
        enter your email, and tick <em>"Confirm Dinner Attendance"</em>.
        The admin will assign your table — participant tables:
        <strong style="color:var(--gold-300);">3, 4, 8, 9, 10, 27</strong>
        (10 seats each).
      </p>
    </div>
    """, unsafe_allow_html=True)


# ============================================================================
# PAGE: SELF CHECK-IN  (★ the main flow)
# ============================================================================
def page_checkin():
    hero("Self Service", "Check-In",
         "Sila masukkan email anda untuk semak status registrasi dan tick "
         "kehadiran sendiri. Walk-in? Sila ke kaunter admin dahulu untuk daftar.")

    # Quick-find form
    st.markdown("<div style='max-width: 600px; margin: 0 auto;'>", unsafe_allow_html=True)
    with st.form("find_me", clear_on_submit=False):
        email = st.text_input(
            "Email", placeholder="you@example.com",
            value=st.session_state.get("ci_email", ""),
            key="ci_email_input"
        ).strip().lower()
        submitted = st.form_submit_button("🔍  Find My Registration",
                                          use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        st.session_state["ci_email"] = email

    email = st.session_state.get("ci_email", "")
    if not email:
        st.markdown("""
        <div style="text-align:center; padding:40px 20px; color:var(--ink-muted);">
          <div style="font-size:48px; opacity:0.5;">🎫</div>
          <p>Masukkan email anda di atas untuk mula check-in.</p>
        </div>
        """, unsafe_allow_html=True)
        return

    me = fetch_one("SELECT * FROM participants WHERE LOWER(email) = ?", (email,))
    if not me:
        st.error(f"❌ Email **{email}** tidak dijumpai dalam sistem. "
                 "Sila pergi ke kaunter admin untuk daftar walk-in.")
        return

    # --- Display participant card ---
    table_banner = ""
    if me["attend_dinner"] and me["table_number"]:
        table_banner = f"""
        <div class="table-banner">
          <div class="lbl">Your Dinner Table</div>
          <div class="num">{me['table_number']}</div>
        </div>
        """
    elif me["attend_dinner"] and not me["table_number"]:
        table_banner = """
        <div class="table-banner" style="background:rgba(243,156,18,0.15);
             color:#f5b860; border:1px solid rgba(243,156,18,0.35);">
          <div class="lbl">Dinner Confirmed</div>
          <div style="font-size:14px; margin-top:6px;">
            Sila tunggu admin assign meja anda
          </div>
        </div>
        """

    st.markdown(f"""
    <div class="me-card">
      <div class="name">{s(me['full_name'])}</div>
      <div class="email">{s(me['email'])}</div>
      <div class="org">🏢 {s(me['organisation'])} · {s(me['category'], 'Participant')}</div>
      {table_banner}
    </div>
    """, unsafe_allow_html=True)

    # --- Tick yourself ---
    st.markdown("### ✓ Tick Your Status")
    st.caption("Tick sendiri. Pastikan tunjuk skrin ini kepada admin di kaunter.")

    c1, c2 = st.columns(2)
    with c1:
        ci_new = st.checkbox(
            "✓ Saya dah sampai (Conference Check-In)",
            value=bool(me["conference_checkin"]),
            key=f"ci_check_{me['id']}"
        )
        dg_new = st.checkbox(
            "🎁  Saya dah ambil door gift",
            value=bool(me["doorgift_collected"]),
            key=f"ci_gift_{me['id']}"
        )
    with c2:
        dn_new = st.checkbox(
            "🍽  Saya akan hadir Gala Dinner",
            value=bool(me["attend_dinner"]),
            key=f"ci_dnr_{me['id']}"
        )
        dnci_new = st.checkbox(
            "✨  Saya dah sampai Gala Dinner (Dinner Check-In)",
            value=bool(me["dinner_checkin"]),
            disabled=not (me["attend_dinner"] and me["table_number"]),
            key=f"ci_dnci_{me['id']}",
            help="Boleh tick selepas admin assign meja"
        )

    # Save if any changed
    changes = (ci_new   != bool(me["conference_checkin"]) or
               dg_new   != bool(me["doorgift_collected"]) or
               dn_new   != bool(me["attend_dinner"])      or
               dnci_new != bool(me["dinner_checkin"]))

    if changes:
        if st.button("💾  Save My Status", use_container_width=True):
            # If turning off dinner → clear table & dinner_checkin
            if not dn_new:
                execute("""UPDATE participants
                           SET conference_checkin = ?, doorgift_collected = ?,
                               attend_dinner = 0, dinner_checkin = 0,
                               table_number = NULL
                           WHERE id = ?""",
                        (int(ci_new), int(dg_new), me["id"]))
            else:
                execute("""UPDATE participants
                           SET conference_checkin = ?, doorgift_collected = ?,
                               attend_dinner = ?, dinner_checkin = ?
                           WHERE id = ?""",
                        (int(ci_new), int(dg_new),
                         int(dn_new), int(dnci_new), me["id"]))
            st.success("✓ Status updated. Terima kasih!")
            st.balloons()
            st.rerun()

    # Status summary
    st.markdown("### Status Summary")
    me = fetch_one("SELECT * FROM participants WHERE id = ?", (me["id"],))
    cols = st.columns(4)
    items = [
        ("Conference",  me["conference_checkin"]),
        ("Door Gift",   me["doorgift_collected"]),
        ("Dinner",      me["attend_dinner"]),
        ("Dinner In",   me["dinner_checkin"]),
    ]
    for col, (lbl, val) in zip(cols, items):
        with col:
            cls = "status-on" if val else "status-off"
            mark = "✓" if val else "○"
            st.markdown(
                f'<div style="text-align:center; padding:14px;">'
                f'<div style="font-size:11px; letter-spacing:1.5px; '
                f'color:var(--ink-muted); margin-bottom:6px;">{lbl}</div>'
                f'<span class="status-badge {cls}">{mark} {"DONE" if val else "PENDING"}</span>'
                f'</div>',
                unsafe_allow_html=True
            )

    # Logout / switch user
    st.markdown("<div style='margin-top:30px;'></div>", unsafe_allow_html=True)
    if st.button("🚪 Check-In someone else", type="secondary"):
        st.session_state["ci_email"] = ""
        st.rerun()


# ============================================================================
# PAGE: REGISTER (public self-register)
# ============================================================================
def page_register():
    hero("Join Us", "New Registration",
         "Fill in your details to register for NICHE 2026. After submitting, "
         "head to the Check-In tab to manage your attendance.")

    st.markdown("<div style='max-width: 700px; margin: 0 auto;'>", unsafe_allow_html=True)
    with st.form("register", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            email = st.text_input("Email *", placeholder="you@example.com")
            name  = st.text_input("Full Name *")
            phone = st.text_input("Phone", placeholder="+60 12-345 6789")
        with c2:
            org = st.text_input("Organisation / Institution")
            cat = st.selectbox("Category",
                ["Participant","Academic Presenter","Industry Delegate",
                 "Sponsor","Media","VIP","Student"])
            dnr = st.checkbox("✨  I would like to attend the Gala Dinner")

        submitted = st.form_submit_button("Submit Registration →",
                                          use_container_width=True)

        if submitted:
            email = email.strip().lower()
            if not email or not name.strip():
                st.error("Email dan Full Name wajib diisi.")
            elif fetch_one("SELECT id FROM participants WHERE LOWER(email) = ?", (email,)):
                st.warning(f"⚠ Email **{email}** sudah ada dalam sistem. "
                           "Sila gi ke tab Check-In.")
            else:
                execute("""INSERT INTO participants
                    (email, full_name, phone, organisation, category,
                     attend_dinner, registration_source)
                    VALUES (?,?,?,?,?,?,?)""",
                    (email, name.strip(), phone.strip(), org.strip(),
                     cat, int(dnr), "Self-Register"))
                st.success(f"✓ Registration berjaya untuk **{name}**! "
                           "Pergi ke tab **Check-In** untuk semak status.")
                st.balloons()
    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================================
# PAGE: ADMIN
# ============================================================================
def page_admin():
    # ---- Login gate ----
    if not st.session_state.get("is_admin", False):
        hero("Restricted", "Admin Access",
             "Untuk organisers sahaja. Masukkan password admin.")
        st.markdown("<div style='max-width: 400px; margin: 0 auto;'>", unsafe_allow_html=True)
        with st.form("login"):
            pw = st.text_input("Password", type="password",
                               placeholder="••••••••")
            if st.form_submit_button("Enter Dashboard →",
                                     use_container_width=True):
                if pw == ADMIN_PASSWORD:
                    st.session_state["is_admin"] = True
                    st.rerun()
                else:
                    st.error("Password salah.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    # ---- Logged in ----
    cols = st.columns([6, 1])
    with cols[0]:
        st.markdown('<h1 style="margin:0;">Admin Dashboard</h1>',
                    unsafe_allow_html=True)
        st.caption("Manage participants · Walk-ins · Table assignments")
    with cols[1]:
        if st.button("Logout", type="secondary"):
            st.session_state["is_admin"] = False
            st.rerun()

    parts_df = fetch_df("SELECT * FROM participants ORDER BY id DESC")

    # ----- Stats -----
    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
    m = st.columns(5)
    metrics = [
        ("Total Registered",  len(parts_df)),
        ("Conference Check-In", int(parts_df["conference_checkin"].sum())),
        ("Door Gift Collected", int(parts_df["doorgift_collected"].sum())),
        ("Attending Dinner",    int(parts_df["attend_dinner"].sum())),
        ("Dinner Check-In",     int(parts_df["dinner_checkin"].sum())),
    ]
    for col, (lbl, v) in zip(m, metrics):
        with col:
            st.metric(lbl, v)

    # ----- Table occupancy -----
    st.markdown('<h3 style="margin-top:30px; font-size:18px;">Participant Tables · Gala Dinner</h3>',
                unsafe_allow_html=True)
    occ = table_occupancy()
    chips = st.columns(len(PARTICIPANT_TABLES))
    for col, t in zip(chips, PARTICIPANT_TABLES):
        with col:
            n = occ[t]
            cls = "tchip full" if n >= SEATS_PER_TABLE else "tchip"
            pct = round(n / SEATS_PER_TABLE * 100)
            st.markdown(f"""
            <div class="{cls}">
              <div class="num">{t}</div>
              <div class="lbl">Table</div>
              <div style="font-size:11px; margin-top:6px; color:var(--ink-soft);">
                <strong style="color:var(--gold-300);">{n}</strong>
                <span style="opacity:0.6;">/ {SEATS_PER_TABLE}</span>
              </div>
              <div class="cap"><div class="cap-fill" style="width:{pct}%"></div></div>
            </div>
            """, unsafe_allow_html=True)

    # ----- Tabs inside admin -----
    st.markdown("<div style='height:30px;'></div>", unsafe_allow_html=True)
    a_tab1, a_tab2, a_tab3 = st.tabs(["📋  All Participants", "➕  Walk-in Registration",
                                       "🍽  Dinner & Tables"])

    # ===== Tab 1: All participants editor =====
    with a_tab1:
        st.caption("Edit secara inline. Klik **Save Changes** untuk simpan. "
                   "Untuk hapus peserta, kosongkan checkbox 'Keep'.")

        # Search filter
        q = st.text_input("Search", placeholder="Cari nama, email, organisasi…",
                          label_visibility="collapsed").strip().lower()
        view_df = parts_df.copy()
        if q:
            mask = (view_df["full_name"].fillna("").str.lower().str.contains(q) |
                    view_df["email"].fillna("").str.lower().str.contains(q) |
                    view_df["organisation"].fillna("").str.lower().str.contains(q))
            view_df = view_df[mask]

        # Prepare for data_editor
        editor_df = view_df[[
            "id", "full_name", "email", "organisation", "category",
            "conference_checkin", "doorgift_collected", "attend_dinner",
            "table_number", "dinner_checkin", "registration_source"
        ]].copy()
        editor_df["conference_checkin"] = editor_df["conference_checkin"].astype(bool)
        editor_df["doorgift_collected"] = editor_df["doorgift_collected"].astype(bool)
        editor_df["attend_dinner"]      = editor_df["attend_dinner"].astype(bool)
        editor_df["dinner_checkin"]     = editor_df["dinner_checkin"].astype(bool)
        editor_df["Keep"] = True
        editor_df["table_number"] = editor_df["table_number"].apply(
            lambda v: int(v) if pd.notna(v) else None
        )

        edited = st.data_editor(
            editor_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            disabled=["id","full_name","email","organisation","category","registration_source"],
            column_config={
                "id":              None,  # hide
                "full_name":       st.column_config.TextColumn("Name", width="medium"),
                "email":           st.column_config.TextColumn("Email", width="medium"),
                "organisation":    st.column_config.TextColumn("Organisation", width="medium"),
                "category":        st.column_config.TextColumn("Category", width="small"),
                "registration_source": st.column_config.TextColumn("Source", width="small"),
                "conference_checkin": st.column_config.CheckboxColumn("✓ Conf"),
                "doorgift_collected": st.column_config.CheckboxColumn("🎁 Gift"),
                "attend_dinner":      st.column_config.CheckboxColumn("🍽 Dnr"),
                "table_number":       st.column_config.SelectboxColumn(
                                          "Table",
                                          options=[None] + PARTICIPANT_TABLES,
                                          width="small"),
                "dinner_checkin":     st.column_config.CheckboxColumn("✨ Dnr-In"),
                "Keep":               st.column_config.CheckboxColumn("Keep",
                                          help="Uncheck = delete on save", default=True),
            },
            key="parts_editor",
        )

        if st.button("💾  Save Changes", use_container_width=True):
            n_upd, n_del, errors = 0, 0, []

            # Build occupancy from EDITED values (not DB) for validation
            future_occ = {t: 0 for t in PARTICIPANT_TABLES}
            for _, row in edited.iterrows():
                t = row["table_number"]
                if row["Keep"] and pd.notna(t) and int(t) in future_occ:
                    future_occ[int(t)] += 1
            for t, n in future_occ.items():
                if n > SEATS_PER_TABLE:
                    errors.append(f"Meja {t} akan over-booked ({n}/{SEATS_PER_TABLE}).")

            if errors:
                for e in errors: st.error(e)
            else:
                # Apply changes vs original
                orig = view_df.set_index("id")
                for _, row in edited.iterrows():
                    pid = int(row["id"])
                    if not row["Keep"]:
                        execute("DELETE FROM participants WHERE id = ?", (pid,))
                        n_del += 1
                        continue
                    o = orig.loc[pid]
                    o_table = int(o["table_number"]) if pd.notna(o["table_number"]) else None
                    n_table = int(row["table_number"]) if pd.notna(row["table_number"]) else None

                    if (bool(o["conference_checkin"]) != bool(row["conference_checkin"]) or
                        bool(o["doorgift_collected"]) != bool(row["doorgift_collected"]) or
                        bool(o["attend_dinner"])      != bool(row["attend_dinner"])      or
                        bool(o["dinner_checkin"])     != bool(row["dinner_checkin"])     or
                        o_table != n_table):

                        attend = int(row["attend_dinner"])
                        if not attend:
                            n_table = None
                            row_dnci = 0
                        else:
                            row_dnci = int(row["dinner_checkin"])

                        # If table assigned, force attend_dinner ON
                        if n_table is not None:
                            attend = 1

                        execute("""UPDATE participants
                            SET conference_checkin = ?, doorgift_collected = ?,
                                attend_dinner = ?, dinner_checkin = ?, table_number = ?
                            WHERE id = ?""",
                            (int(row["conference_checkin"]),
                             int(row["doorgift_collected"]),
                             attend, row_dnci, n_table, pid))
                        n_upd += 1

                msg_bits = []
                if n_upd: msg_bits.append(f"{n_upd} updated")
                if n_del: msg_bits.append(f"{n_del} deleted")
                if msg_bits:
                    st.success("✓ " + ", ".join(msg_bits) + ".")
                    st.rerun()
                else:
                    st.info("Tiada perubahan.")

    # ===== Tab 2: Walk-in =====
    with a_tab2:
        st.caption("Daftar peserta yang datang on-the-spot. Auto check-in dilakukan.")
        with st.form("walkin", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                w_email = st.text_input("Email *")
                w_name  = st.text_input("Full Name *")
                w_phone = st.text_input("Phone")
            with c2:
                w_org = st.text_input("Organisation")
                w_cat = st.selectbox("Category",
                    ["Walk-In","Academic Presenter","Industry Delegate",
                     "Sponsor","Media","VIP","Student"])
                w_dnr = st.checkbox("✨  Attending Gala Dinner")

            if st.form_submit_button("Register & Auto Check-In →",
                                     use_container_width=True):
                w_email = w_email.strip().lower()
                if not w_email or not w_name.strip():
                    st.error("Email dan Name wajib.")
                elif fetch_one("SELECT id FROM participants WHERE LOWER(email)=?", (w_email,)):
                    st.warning(f"Email {w_email} sudah ada.")
                else:
                    execute("""INSERT INTO participants
                        (email, full_name, phone, organisation, category,
                         conference_checkin, attend_dinner, registration_source)
                        VALUES (?,?,?,?,?, 1, ?, 'Walk-In')""",
                        (w_email, w_name.strip(), w_phone.strip(),
                         w_org.strip(), w_cat, int(w_dnr)))
                    st.success(f"✓ Walk-in {w_name} didaftar & check-in.")
                    st.balloons()

    # ===== Tab 3: Dinner & table assignment =====
    with a_tab3:
        st.caption("Quick table assignment untuk participants yang attending dinner.")
        d_df = parts_df[parts_df["attend_dinner"] == 1].copy()

        if len(d_df) == 0:
            st.info("Belum ada participants confirmed untuk Gala Dinner.")
        else:
            st.markdown(f"**{len(d_df)} participants** attending. "
                        f"**{int(d_df['table_number'].notna().sum())}** assigned, "
                        f"**{int(d_df['table_number'].isna().sum())}** unassigned.")

            for _, p in d_df.iterrows():
                with st.container():
                    c1, c2, c3 = st.columns([3, 2, 2])
                    with c1:
                        st.markdown(
                            f"**{p['full_name']}**  \n"
                            f"<span style='color:var(--ink-muted); "
                            f"font-family:ui-monospace,monospace; "
                            f"font-size:12px;'>{p['email']}</span>  \n"
                            f"<span style='color:var(--ink-muted); font-size:12px;'>"
                            f"{p['organisation'] or '—'}</span>",
                            unsafe_allow_html=True
                        )
                    with c2:
                        cur = int(p["table_number"]) if pd.notna(p["table_number"]) else None
                        new_table = st.selectbox(
                            "Table", [None] + PARTICIPANT_TABLES,
                            index=([None]+PARTICIPANT_TABLES).index(cur),
                            format_func=lambda x: "— None —" if x is None else f"Table {x}",
                            key=f"tbl_{p['id']}", label_visibility="collapsed"
                        )
                    with c3:
                        if new_table != cur:
                            if st.button("Save", key=f"save_{p['id']}"):
                                if new_table is None:
                                    execute("UPDATE participants SET table_number=NULL WHERE id=?",
                                            (int(p["id"]),))
                                    st.toast(f"Cleared table for {p['full_name']}")
                                else:
                                    occ_now = table_occupancy()
                                    if new_table != cur and occ_now[new_table] >= SEATS_PER_TABLE:
                                        st.error(f"Table {new_table} penuh!")
                                    else:
                                        execute("UPDATE participants SET table_number=? WHERE id=?",
                                                (int(new_table), int(p["id"])))
                                        st.toast(f"Assigned {p['full_name']} → Table {new_table}")
                                st.rerun()


# ============================================================================
# MAIN
# ============================================================================
def main():
    inject_theme()
    init_db()
    brand_bar()

    tab_home, tab_ac, tab_in, tab_dn, tab_ci, tab_rg, tab_ad = st.tabs([
        "🏠  Home",
        "🎓  Academic",
        "🏛  Industry",
        "✨  Gala Dinner",
        "✓  Check-In",
        "📝  Register",
        "🔐  Admin",
    ])

    with tab_home: page_home()
    with tab_ac:   page_academic()
    with tab_in:   page_industry()
    with tab_dn:   page_dinner()
    with tab_ci:   page_checkin()
    with tab_rg:   page_register()
    with tab_ad:   page_admin()

    st.markdown("""
    <div style="text-align:center; padding:30px 20px; margin-top:30px;
                border-top:1px solid rgba(212,175,55,0.15);
                color:var(--ink-muted); font-size:12px;">
      NICHE 2026 · International Halal Conference · Royale Chulan Seremban<br>
      Organised by IFEA · UiTM · Evolusi Dekad 7
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
