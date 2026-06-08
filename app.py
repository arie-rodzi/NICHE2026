"""
NICHE 2026 — Streamlit Conference Registration & Check-In System
================================================================
Deploy on Streamlit Cloud.

Required files:
    app.py
    requirements.txt
Optional:
    niche_data.xlsx

requirements.txt:
    streamlit
    pandas
    openpyxl
"""

import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="NICHE 2026 Registration System",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "niche.db"
EXCEL_PATH = BASE_DIR / "niche_data.xlsx"

EVENT_NAME = "NICHE 2026"
EVENT_DATES = "9 – 10 June 2026"
EVENT_VENUE = "Royale Chulan Seremban, Negeri Sembilan"

ADMIN_PASSWORD = "NICHE2026admin"

PARTICIPANT_TABLES = [3, 4, 8, 9, 10, 27]
SEATS_PER_TABLE = 10


# =========================================================
# STYLE
# =========================================================
st.markdown(
    """
<style>
[data-testid="stSidebar"] {display:none;}
[data-testid="collapsedControl"] {display:none;}

html, body, [class*="css"] {
    font-family: Inter, Segoe UI, Arial, sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(255,214,102,.22), transparent 32%),
        radial-gradient(circle at top right, rgba(0,51,102,.24), transparent 34%),
        linear-gradient(135deg, #07111f 0%, #0d1b2a 36%, #112a46 100%);
    color: #f8fafc;
}

.block-container {
    padding-top: 1.2rem;
    max-width: 1400px;
}

.hero {
    padding: 26px 30px;
    border-radius: 28px;
    background: linear-gradient(135deg, rgba(255,255,255,.13), rgba(255,255,255,.05));
    border: 1px solid rgba(255,255,255,.18);
    box-shadow: 0 28px 80px rgba(0,0,0,.32);
    margin-bottom: 18px;
}

.hero h1 {
    margin: 0;
    font-size: 2.05rem;
    letter-spacing: -.03em;
    color: #ffffff;
}

.hero p {
    margin-top: 8px;
    color: #dbeafe;
    font-size: 1rem;
}

.gold {
    color: #ffd166;
    font-weight: 800;
}

.card {
    padding: 18px 20px;
    border-radius: 22px;
    background: rgba(255,255,255,.10);
    border: 1px solid rgba(255,255,255,.16);
    box-shadow: 0 18px 48px rgba(0,0,0,.22);
    margin-bottom: 14px;
}

.metric-card {
    padding: 18px;
    border-radius: 22px;
    background: linear-gradient(145deg, rgba(255,255,255,.16), rgba(255,255,255,.06));
    border: 1px solid rgba(255,255,255,.17);
}

.metric-card .label {
    color: #bfdbfe;
    font-size: .85rem;
    font-weight: 700;
    text-transform: uppercase;
}

.metric-card .value {
    color: #ffffff;
    font-size: 2rem;
    font-weight: 900;
    margin-top: 4px;
}

.nav-btn button {
    border-radius: 18px !important;
    min-height: 50px;
    font-weight: 800 !important;
}

div.stButton > button {
    border-radius: 14px;
    border: 1px solid rgba(255,255,255,.20);
    background: linear-gradient(135deg, #f6c453 0%, #d99b20 100%);
    color: #111827;
    font-weight: 800;
}

div.stButton > button:hover {
    border: 1px solid rgba(255,255,255,.45);
    filter: brightness(1.08);
}

.stTextInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea {
    border-radius: 12px;
}

[data-testid="stDataFrame"] {
    background: white;
    border-radius: 18px;
    overflow: hidden;
}

hr {
    border: none;
    border-top: 1px solid rgba(255,255,255,.15);
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# SESSION
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "Home"
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False


# =========================================================
# DATABASE HELPERS
# =========================================================
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _clean(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    s = str(v).strip()
    if s.lower() in ("nan", "nat", "", "none"):
        return None
    return s


def _yes(v):
    return 1 if str(v).strip().lower() in ("yes", "y", "true", "1", "ya", "hadir") else 0


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript(
        """
        CREATE TABLE IF NOT EXISTS participants (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            email               TEXT UNIQUE NOT NULL,
            full_name           TEXT NOT NULL,
            phone               TEXT,
            organisation        TEXT,
            category            TEXT,
            academic            INTEGER DEFAULT 0,
            industry            INTEGER DEFAULT 0,
            conference_checkin  INTEGER DEFAULT 0,
            doorgift_collected  INTEGER DEFAULT 0,
            attend_dinner       INTEGER DEFAULT 0,
            dinner_checkin      INTEGER DEFAULT 0,
            table_number        INTEGER,
            seat_number         INTEGER,
            registration_source TEXT DEFAULT 'Preloaded',
            created_at          TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS academic_programme (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            venue     TEXT, time TEXT, session TEXT, moderator TEXT,
            theme     TEXT, paper_id TEXT, title TEXT, presenter TEXT, email TEXT,
            sort_order INTEGER
        );

        CREATE TABLE IF NOT EXISTS abstracts (
            paper_id        TEXT PRIMARY KEY,
            title           TEXT, presenter TEXT, email TEXT,
            venue           TEXT, time TEXT, session TEXT,
            keywords        TEXT, abstract_text TEXT, authors TEXT
        );

        CREATE TABLE IF NOT EXISTS industry_programme (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            day       TEXT, time TEXT, venue TEXT, session TEXT,
            speaker   TEXT, organisation TEXT, details TEXT,
            sort_order INTEGER
        );

        CREATE TABLE IF NOT EXISTS dinner_programme (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            time      TEXT, event TEXT, sort_order INTEGER
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            ts        TEXT DEFAULT CURRENT_TIMESTAMP,
            action    TEXT, target_email TEXT, detail TEXT
        );
        """
    )
    conn.commit()

    n = conn.execute("SELECT COUNT(*) AS n FROM participants").fetchone()["n"]
    if n == 0 and EXCEL_PATH.exists():
        seed_from_excel(EXCEL_PATH, clear_existing=False)

    conn.close()


def log_action(action, email=None, detail=None):
    conn = get_db()
    conn.execute(
        "INSERT INTO audit_log (action, target_email, detail) VALUES (?,?,?)",
        (action, email, detail),
    )
    conn.commit()
    conn.close()


def clear_all_data():
    conn = get_db()
    c = conn.cursor()
    for table in [
        "participants",
        "academic_programme",
        "abstracts",
        "industry_programme",
        "dinner_programme",
        "audit_log",
    ]:
        c.execute(f"DELETE FROM {table}")
    conn.commit()
    conn.close()


def seed_from_excel(path, clear_existing=True):
    if clear_existing:
        clear_all_data()

    xl = pd.ExcelFile(path)
    conn = get_db()
    c = conn.cursor()

    if "Participants" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Participants")
        for _, r in df.iterrows():
            email = _clean(r.get("Email"))
            if not email:
                continue
            table_val = None
            if "Table_Number" in df.columns and pd.notna(r.get("Table_Number")):
                try:
                    table_val = int(r.get("Table_Number"))
                except Exception:
                    table_val = None

            c.execute(
                """
                INSERT OR IGNORE INTO participants
                (email, full_name, phone, organisation, category,
                 academic, industry, conference_checkin, doorgift_collected,
                 attend_dinner, dinner_checkin, table_number, registration_source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    email.strip().lower(),
                    _clean(r.get("Full_Name")) or "—",
                    _clean(r.get("Phone")),
                    _clean(r.get("Organisation")),
                    _clean(r.get("Category")),
                    _yes(r.get("Academic")),
                    _yes(r.get("Industry")),
                    _yes(r.get("Conference_CheckIn")),
                    _yes(r.get("DoorGift_Collected")),
                    _yes(r.get("Attend_Dinner")),
                    _yes(r.get("Dinner_CheckIn")),
                    table_val,
                    _clean(r.get("Registration_Source")) or "Preloaded",
                ),
            )

    if "Academic_Programme" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Academic_Programme")
        for i, r in df.iterrows():
            c.execute(
                """
                INSERT INTO academic_programme
                (venue, time, session, moderator, theme, paper_id, title,
                 presenter, email, sort_order)
                VALUES (?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    _clean(r.get("Venue")),
                    _clean(r.get("Time")),
                    _clean(r.get("Session")),
                    _clean(r.get("Moderator")),
                    _clean(r.get("Theme")),
                    _clean(r.get("Paper_ID")),
                    _clean(r.get("Title")),
                    _clean(r.get("Presenter")),
                    _clean(r.get("Email_From_Abstract")) or _clean(r.get("Email")),
                    i,
                ),
            )

    if "Abstracts" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Abstracts")
        for _, r in df.iterrows():
            pid = _clean(r.get("Paper_ID"))
            if not pid:
                continue
            c.execute(
                """
                INSERT OR REPLACE INTO abstracts
                (paper_id, title, presenter, email, venue, time, session,
                 keywords, abstract_text, authors)
                VALUES (?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    pid,
                    _clean(r.get("Title")),
                    _clean(r.get("Presenter")),
                    _clean(r.get("Email")),
                    _clean(r.get("Venue")),
                    _clean(r.get("Time")),
                    _clean(r.get("Session")),
                    _clean(r.get("Keywords")),
                    _clean(r.get("Abstract_Text")),
                    _clean(r.get("Authors_Affiliation")) or _clean(r.get("Authors")),
                ),
            )

    if "Industry_Programme" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Industry_Programme")
        for i, r in df.iterrows():
            c.execute(
                """
                INSERT INTO industry_programme
                (day, time, venue, session, speaker, organisation, details, sort_order)
                VALUES (?,?,?,?,?,?,?,?)
                """,
                (
                    _clean(r.get("Day")),
                    _clean(r.get("Time")),
                    _clean(r.get("Venue")),
                    _clean(r.get("Session")),
                    _clean(r.get("Speaker")),
                    _clean(r.get("Organisation")),
                    _clean(r.get("Details")),
                    i,
                ),
            )

    if "Gala_Dinner_Programme" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Gala_Dinner_Programme")
        for i, r in df.iterrows():
            c.execute(
                "INSERT INTO dinner_programme (time, event, sort_order) VALUES (?,?,?)",
                (_clean(r.get("Time")), _clean(r.get("Event")), i),
            )

    conn.commit()
    conn.close()
    log_action("seed_from_excel", detail=f"file={Path(path).name}")


def query_df(sql, params=()):
    conn = get_db()
    df = pd.read_sql_query(sql, conn, params=params)
    conn.close()
    return df


def exec_sql(sql, params=()):
    conn = get_db()
    conn.execute(sql, params)
    conn.commit()
    conn.close()


def get_stats():
    conn = get_db()
    stats = {
        "participants": conn.execute("SELECT COUNT(*) c FROM participants").fetchone()["c"],
        "checked": conn.execute("SELECT COUNT(*) c FROM participants WHERE conference_checkin=1").fetchone()["c"],
        "gift": conn.execute("SELECT COUNT(*) c FROM participants WHERE doorgift_collected=1").fetchone()["c"],
        "dinner": conn.execute("SELECT COUNT(*) c FROM participants WHERE attend_dinner=1").fetchone()["c"],
        "dinner_in": conn.execute("SELECT COUNT(*) c FROM participants WHERE dinner_checkin=1").fetchone()["c"],
        "papers": conn.execute("SELECT COUNT(*) c FROM abstracts").fetchone()["c"],
        "speakers": conn.execute("SELECT COUNT(*) c FROM industry_programme WHERE speaker IS NOT NULL AND speaker != ''").fetchone()["c"],
    }
    conn.close()
    return stats


# =========================================================
# UI HELPERS
# =========================================================
def hero():
    st.markdown(
        f"""
        <div class="hero">
            <h1>{EVENT_NAME} <span class="gold">Registration & Check-In System</span></h1>
            <p>{EVENT_DATES} • {EVENT_VENUE}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label, value):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def nav():
    pages = ["Home", "Academic", "Industry", "Gala Dinner", "Register", "Admin"]
    cols = st.columns(len(pages))
    for col, p in zip(cols, pages):
        with col:
            if st.button(p, use_container_width=True, key=f"nav_{p}"):
                st.session_state.page = p
                st.rerun()


def login_box():
    st.markdown("### 🔐 Admin Login")
    pw = st.text_input("Password", type="password")
    if st.button("Login Admin", use_container_width=True):
        if pw == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            log_action("admin_login")
            st.success("Login berjaya.")
            st.rerun()
        else:
            st.error("Password salah.")


def status_badge(v, yes="Ya", no="Belum"):
    return yes if int(v or 0) == 1 else no


# =========================================================
# PAGES
# =========================================================
def page_home():
    hero()
    stats = get_stats()

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: metric_card("Participants", stats["participants"])
    with c2: metric_card("Check-In", stats["checked"])
    with c3: metric_card("Door Gift", stats["gift"])
    with c4: metric_card("Dinner", stats["dinner"])
    with c5: metric_card("Papers", stats["papers"])

    st.markdown("### 🎯 System Menu")
    a, b, c = st.columns(3)
    with a:
        st.markdown('<div class="card"><h3>Public Registration</h3><p>Peserta boleh daftar menggunakan email, nama, telefon, organisasi dan pilihan dinner.</p></div>', unsafe_allow_html=True)
        if st.button("Go to Registration", use_container_width=True):
            st.session_state.page = "Register"
            st.rerun()
    with b:
        st.markdown('<div class="card"><h3>Programme Tentative</h3><p>Papar academic programme, industry programme dan gala dinner programme.</p></div>', unsafe_allow_html=True)
        if st.button("View Academic Programme", use_container_width=True):
            st.session_state.page = "Academic"
            st.rerun()
    with c:
        st.markdown('<div class="card"><h3>Admin Counter</h3><p>Check-in, door gift, walk-in registration, dinner attendance dan table assignment.</p></div>', unsafe_allow_html=True)
        if st.button("Open Admin", use_container_width=True):
            st.session_state.page = "Admin"
            st.rerun()


def page_academic():
    hero()
    st.markdown("## 📚 Academic Conference Tentative")

    programme = query_df("SELECT * FROM academic_programme ORDER BY sort_order")
    abstracts = query_df("SELECT * FROM abstracts")

    if programme.empty:
        st.warning("Tiada data Academic_Programme. Admin perlu upload niche_data.xlsx.")
        return

    venues = ["All"] + sorted([v for v in programme["venue"].dropna().unique()])
    selected_venue = st.selectbox("Filter Venue", venues)

    df = programme.copy()
    if selected_venue != "All":
        df = df[df["venue"] == selected_venue]

    for venue, g in df.groupby(df["venue"].fillna("—"), sort=False):
        st.markdown(f"### 📍 {venue}")
        show_cols = ["time", "session", "moderator", "theme", "paper_id", "title", "presenter"]
        st.dataframe(g[show_cols], use_container_width=True, hide_index=True)

    if not abstracts.empty:
        st.markdown("## 📝 Abstracts")
        search = st.text_input("Search abstract / presenter / paper ID")
        adf = abstracts.copy()
        if search:
            s = search.lower()
            adf = adf[
                adf.fillna("").astype(str).apply(
                    lambda row: row.str.lower().str.contains(s).any(), axis=1
                )
            ]

        for _, r in adf.iterrows():
            with st.expander(f"{r.get('paper_id','')} — {r.get('title','Untitled')}"):
                st.write(f"**Presenter:** {r.get('presenter','—')}")
                st.write(f"**Authors/Affiliation:** {r.get('authors','—')}")
                st.write(f"**Keywords:** {r.get('keywords','—')}")
                st.write(r.get("abstract_text", "—"))


def page_industry():
    hero()
    st.markdown("## 🏭 Industry Conference Tentative")

    df = query_df("SELECT * FROM industry_programme ORDER BY sort_order")
    if df.empty:
        st.warning("Tiada data Industry_Programme.")
        return

    for day, g in df.groupby(df["day"].fillna("—"), sort=False):
        st.markdown(f"### {day}")
        st.dataframe(
            g[["time", "venue", "session", "speaker", "organisation", "details"]],
            use_container_width=True,
            hide_index=True,
        )


def page_dinner():
    hero()
    st.markdown("## 🍽️ Gala Dinner Programme")

    df = query_df("SELECT * FROM dinner_programme ORDER BY sort_order")
    if df.empty:
        st.warning("Tiada data Gala_Dinner_Programme.")
        return

    st.dataframe(df[["time", "event"]], use_container_width=True, hide_index=True)

    occ = query_df(
        """
        SELECT table_number, COUNT(*) AS occupied
        FROM participants
        WHERE table_number IS NOT NULL
        GROUP BY table_number
        """
    )
    occ_map = dict(zip(occ["table_number"], occ["occupied"])) if not occ.empty else {}

    st.markdown("### 🪑 Participant Table Occupancy")
    cols = st.columns(len(PARTICIPANT_TABLES))
    for col, t in zip(cols, PARTICIPANT_TABLES):
        with col:
            metric_card(f"Table {t}", f"{occ_map.get(t,0)}/{SEATS_PER_TABLE}")


def page_register():
    hero()
    st.markdown("## 📝 Public Self-Registration")

    with st.form("register_form"):
        email = st.text_input("Email *").strip().lower()
        name = st.text_input("Full Name *").strip()
        phone = st.text_input("Phone")
        organisation = st.text_input("Organisation")
        category = st.selectbox("Category", ["Participant", "Academic Presenter", "Industry Participant", "Speaker", "Secretariat", "Guest"])
        attend_dinner = st.radio("Attend Gala Dinner?", ["Yes", "No"], horizontal=True)
        submitted = st.form_submit_button("Submit Registration", use_container_width=True)

    if submitted:
        if not email or not name:
            st.error("Email dan nama wajib diisi.")
            return

        conn = get_db()
        exists = conn.execute("SELECT full_name FROM participants WHERE email=?", (email,)).fetchone()
        if exists:
            conn.close()
            st.info(f"Email ini sudah ada dalam sistem: {exists['full_name']}. Sila tunjuk kepada admin di kaunter.")
            return

        conn.execute(
            """
            INSERT INTO participants
            (email, full_name, phone, organisation, category, attend_dinner, registration_source)
            VALUES (?,?,?,?,?,?,?)
            """,
            (
                email,
                name,
                phone.strip(),
                organisation.strip(),
                category,
                1 if attend_dinner == "Yes" else 0,
                "Self-Register",
            ),
        )
        conn.commit()
        conn.close()
        log_action("self_register", email, name)
        st.success("Registration berjaya. Sila tunjuk email kepada admin untuk check-in dan door gift.")


def page_admin():
    hero()
    st.markdown("## 🛡️ Admin Dashboard")

    if not st.session_state.is_admin:
        login_box()
        return

    if st.button("Logout Admin"):
        st.session_state.is_admin = False
        st.rerun()

    tabs = st.tabs(["Dashboard", "Counter Check-In", "Walk-In", "Upload / Reset Data", "Audit Log"])

    with tabs[0]:
        stats = get_stats()
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: metric_card("Total", stats["participants"])
        with c2: metric_card("Checked-In", stats["checked"])
        with c3: metric_card("Door Gift", stats["gift"])
        with c4: metric_card("Dinner List", stats["dinner"])
        with c5: metric_card("Dinner In", stats["dinner_in"])

        df = query_df("SELECT * FROM participants ORDER BY created_at DESC, id DESC")
        if df.empty:
            st.warning("Tiada participant lagi.")
        else:
            show = df.copy()
            for col in ["conference_checkin", "doorgift_collected", "attend_dinner", "dinner_checkin"]:
                show[col] = show[col].apply(lambda x: "✅" if int(x or 0) == 1 else "—")
            st.dataframe(
                show[
                    [
                        "id", "full_name", "email", "phone", "organisation", "category",
                        "conference_checkin", "doorgift_collected", "attend_dinner",
                        "dinner_checkin", "table_number", "registration_source"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

    with tabs[1]:
        st.markdown("### ✅ Counter Check-In / Door Gift / Dinner / Table")
        df = query_df("SELECT * FROM participants ORDER BY full_name")
        search = st.text_input("Search by name / email / organisation", key="admin_search")

        if search:
            s = search.lower()
            df = df[
                df.fillna("").astype(str).apply(
                    lambda row: row.str.lower().str.contains(s).any(), axis=1
                )
            ]

        if df.empty:
            st.warning("Tiada rekod dijumpai.")
        else:
            for _, p in df.iterrows():
                with st.container(border=True):
                    top1, top2 = st.columns([3, 1])
                    with top1:
                        st.markdown(f"### {p['full_name']}")
                        st.write(f"**Email:** {p['email']}  \n**Org:** {p.get('organisation') or '—'}  \n**Category:** {p.get('category') or '—'}")
                    with top2:
                        st.write(f"**Table:** {p.get('table_number') or '—'}")
                        st.write(f"**Source:** {p.get('registration_source') or '—'}")

                    c1, c2, c3, c4, c5 = st.columns(5)

                    def update_bool(field, current):
                        new_val = 0 if int(current or 0) == 1 else 1
                        exec_sql(f"UPDATE participants SET {field}=? WHERE id=?", (new_val, int(p["id"])))
                        if field == "attend_dinner" and new_val == 0:
                            exec_sql("UPDATE participants SET table_number=NULL, dinner_checkin=0 WHERE id=?", (int(p["id"]),))
                        log_action(f"toggle:{field}={new_val}", p["email"])
                        st.rerun()

                    with c1:
                        if st.button(f"Check-In: {status_badge(p['conference_checkin'])}", key=f"ci_{p['id']}", use_container_width=True):
                            update_bool("conference_checkin", p["conference_checkin"])
                    with c2:
                        if st.button(f"Door Gift: {status_badge(p['doorgift_collected'])}", key=f"dg_{p['id']}", use_container_width=True):
                            update_bool("doorgift_collected", p["doorgift_collected"])
                    with c3:
                        if st.button(f"Dinner: {status_badge(p['attend_dinner'])}", key=f"ad_{p['id']}", use_container_width=True):
                            update_bool("attend_dinner", p["attend_dinner"])
                    with c4:
                        if st.button(f"Dinner In: {status_badge(p['dinner_checkin'])}", key=f"di_{p['id']}", use_container_width=True):
                            update_bool("dinner_checkin", p["dinner_checkin"])
                    with c5:
                        table_options = [0] + PARTICIPANT_TABLES
                        current_table = int(p["table_number"]) if pd.notna(p["table_number"]) and p["table_number"] else 0
                        selected_table = st.selectbox(
                            "Table",
                            table_options,
                            index=table_options.index(current_table) if current_table in table_options else 0,
                            key=f"table_{p['id']}",
                            format_func=lambda x: "No Table" if x == 0 else f"Table {x}",
                        )
                        if st.button("Save Table", key=f"save_table_{p['id']}", use_container_width=True):
                            if selected_table == 0:
                                exec_sql("UPDATE participants SET table_number=NULL WHERE id=?", (int(p["id"]),))
                                log_action("table_unassign", p["email"])
                                st.success("Table cleared.")
                                st.rerun()

                            n = query_df(
                                "SELECT COUNT(*) AS c FROM participants WHERE table_number=? AND id!=?",
                                (selected_table, int(p["id"])),
                            )["c"].iloc[0]
                            if n >= SEATS_PER_TABLE:
                                st.error(f"Table {selected_table} sudah penuh ({n}/{SEATS_PER_TABLE}).")
                            else:
                                exec_sql(
                                    "UPDATE participants SET table_number=?, attend_dinner=1 WHERE id=?",
                                    (selected_table, int(p["id"])),
                                )
                                log_action("table_assign", p["email"], f"table={selected_table}")
                                st.success("Table saved.")
                                st.rerun()

                    if st.button("Delete Participant", key=f"del_{p['id']}"):
                        exec_sql("DELETE FROM participants WHERE id=?", (int(p["id"]),))
                        log_action("delete", p["email"])
                        st.warning("Deleted.")
                        st.rerun()

    with tabs[2]:
        st.markdown("### ➕ Walk-In Registration")
        with st.form("walkin_form"):
            email = st.text_input("Email *", key="walkin_email").strip().lower()
            name = st.text_input("Full Name *", key="walkin_name").strip()
            phone = st.text_input("Phone", key="walkin_phone")
            organisation = st.text_input("Organisation", key="walkin_org")
            category = st.selectbox("Category", ["Walk-In", "Participant", "Academic Presenter", "Industry Participant", "Speaker", "Secretariat", "Guest"])
            attend_dinner = st.radio("Attend Dinner?", ["Yes", "No"], horizontal=True, key="walkin_dinner")
            ok = st.form_submit_button("Register Walk-In & Auto Check-In", use_container_width=True)

        if ok:
            if not email or not name:
                st.error("Email dan nama wajib.")
            else:
                conn = get_db()
                try:
                    conn.execute(
                        """
                        INSERT INTO participants
                        (email, full_name, phone, organisation, category,
                         attend_dinner, conference_checkin, registration_source)
                        VALUES (?,?,?,?,?,?,?,?)
                        """,
                        (
                            email, name, phone.strip(), organisation.strip(), category,
                            1 if attend_dinner == "Yes" else 0,
                            1,
                            "Walk-In",
                        ),
                    )
                    conn.commit()
                    log_action("walk_in", email, name)
                    st.success(f"Walk-in {name} berjaya didaftar dan check-in.")
                except sqlite3.IntegrityError:
                    st.error(f"Email {email} sudah ada dalam sistem.")
                finally:
                    conn.close()

    with tabs[3]:
        st.markdown("### 📤 Upload / Reset Master Excel")
        st.info("Sheet yang disokong: Participants, Academic_Programme, Abstracts, Industry_Programme, Gala_Dinner_Programme.")

        uploaded = st.file_uploader("Upload niche_data.xlsx", type=["xlsx"])
        if uploaded is not None:
            with open(EXCEL_PATH, "wb") as f:
                f.write(uploaded.getbuffer())
            if st.button("Import Excel & Replace Existing Data", use_container_width=True):
                try:
                    seed_from_excel(EXCEL_PATH, clear_existing=True)
                    st.success("Data berjaya diimport.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Excel gagal dibaca: {e}")

        st.markdown("---")
        st.warning("Reset akan kosongkan semua data participant, tentative, abstract, dinner dan audit log.")
        confirm = st.checkbox("Saya faham dan mahu reset semua data.")
        if confirm and st.button("RESET ALL DATA", use_container_width=True):
            clear_all_data()
            st.success("Semua data telah dikosongkan.")
            st.rerun()

    with tabs[4]:
        st.markdown("### 🧾 Audit Log")
        audit = query_df("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500")
        st.dataframe(audit, use_container_width=True, hide_index=True)


# =========================================================
# MAIN
# =========================================================
init_db()
hero()
nav()
st.markdown("<hr>", unsafe_allow_html=True)

page = st.session_state.page
if page == "Home":
    page_home()
elif page == "Academic":
    page_academic()
elif page == "Industry":
    page_industry()
elif page == "Gala Dinner":
    page_dinner()
elif page == "Register":
    page_register()
elif page == "Admin":
    page_admin()
else:
    page_home()
