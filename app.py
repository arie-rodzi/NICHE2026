import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

APP_TITLE = "NICHE 2026"
ADMIN_PASSWORD = "NICHE2026admin"
DB_PATH = "niche2026.db"
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

st.set_page_config(
    page_title="NICHE 2026",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================
# CSS PREMIUM NO SIDEBAR
# =========================
st.markdown("""
<style>
[data-testid="stSidebar"] {display:none;}
[data-testid="collapsedControl"] {display:none;}

.stApp {
    background:
    radial-gradient(circle at top left, rgba(255,205,90,0.22), transparent 35%),
    radial-gradient(circle at top right, rgba(0,70,160,0.28), transparent 35%),
    linear-gradient(135deg, #050816 0%, #0B1229 45%, #111827 100%);
    color: #ffffff;
}

.block-container {
    padding-top: 1.2rem;
    max-width: 1400px;
}

h1, h2, h3 {
    color: #ffffff;
    font-weight: 800;
}

.nav-card, .glass, .metric-card {
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.18);
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 18px 45px rgba(0,0,0,0.35);
    backdrop-filter: blur(16px);
}

.hero {
    padding: 2px;
    border-radius: 32px;
    background:
    linear-gradient(135deg, rgba(255,215,128,0.20), rgba(0,90,180,0.22)),
    rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.20);
    box-shadow: 0 25px 60px rgba(0,0,0,0.45);
}

.gold {
    color: #F7C948;
}

.small {
    color: rgba(255,255,255,0.72);
    font-size: 0.95rem;
}

.stButton > button {
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.18);
    background: linear-gradient(135deg, #F7C948, #E89B2C);
    color: #111827;
    font-weight: 800;
    padding: 0.65rem 1.1rem;
}

.stButton > button:hover {
    transform: translateY(-1px);
    border-color: #ffffff;
}

input, textarea, select {
    border-radius: 14px !important;
}

[data-testid="stMetricValue"] {
    color: #F7C948;
}

hr {
    border-color: rgba(255,255,255,0.15);
}
</style>
""", unsafe_allow_html=True)


# =========================
# DATABASE
# =========================
def conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    c = conn()
    cur = c.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS participants (
        email TEXT PRIMARY KEY,
        full_name TEXT,
        phone TEXT,
        organisation TEXT,
        category TEXT,
        academic TEXT,
        industry TEXT,
        attend_dinner TEXT,
        conference_checkin TEXT DEFAULT 'No',
        conference_checkin_time TEXT,
        door_gift_collected TEXT DEFAULT 'No',
        door_gift_time TEXT,
        dinner_confirmed TEXT DEFAULT 'No',
        dinner_checkin TEXT DEFAULT 'No',
        dinner_checkin_time TEXT,
        table_number TEXT,
        seat_number TEXT,
        registration_source TEXT DEFAULT 'Online',
        created_at TEXT,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS academic_programme (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        venue TEXT,
        time TEXT,
        session TEXT,
        theme TEXT,
        paper_id TEXT,
        title TEXT,
        presenter TEXT,
        email TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS abstracts (
        paper_id TEXT PRIMARY KEY,
        title TEXT,
        presenter TEXT,
        email TEXT,
        keywords TEXT,
        abstract_text TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS industry_programme (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        day TEXT,
        time TEXT,
        venue TEXT,
        session TEXT,
        speaker TEXT,
        organisation TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS industry_speakers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        speaker_name TEXT,
        email TEXT,
        phone TEXT,
        organisation TEXT,
        designation TEXT,
        session TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS gala_dinner_programme (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        time TEXT,
        event TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS dinner_tables (
        table_number TEXT PRIMARY KEY,
        capacity INTEGER,
        type TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS system_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date_time TEXT,
        action TEXT,
        admin TEXT,
        email TEXT
    )
    """)

    c.commit()
    c.close()


def log_action(action, email=""):
    c = conn()
    c.execute(
        "INSERT INTO system_log(date_time, action, admin, email) VALUES (?, ?, ?, ?)",
        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), action, "Admin", email)
    )
    c.commit()
    c.close()


def df_query(sql, params=()):
    c = conn()
    df = pd.read_sql_query(sql, c, params=params)
    c.close()
    return df


def execute(sql, params=()):
    c = conn()
    c.execute(sql, params)
    c.commit()
    c.close()


init_db()


# =========================
# DATA HELPERS
# =========================
def clean_col(x):
    return str(x).strip().replace(" ", "_")


def load_sheet_if_exists(xls, sheet):
    if sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [clean_col(c) for c in df.columns]
        return df.fillna("")
    return pd.DataFrame()


def upload_master_excel(file):
    xls = pd.ExcelFile(file)

    mapping = {
        "Academic_Programme": "academic_programme",
        "Abstracts": "abstracts",
        "Industry_Programme": "industry_programme",
        "Industry_Speakers": "industry_speakers",
        "Gala_Dinner_Programme": "gala_dinner_programme",
        "Dinner_Tables": "dinner_tables",
    }

    c = conn()

    for sheet, table in mapping.items():
        df = load_sheet_if_exists(xls, sheet)
        if not df.empty:
            df.to_sql(table, c, if_exists="replace", index=False)

    participants = load_sheet_if_exists(xls, "Participants")
    if not participants.empty:
        for _, r in participants.iterrows():
            email = str(r.get("Email", "")).strip().lower()
            if email:
                c.execute("""
                INSERT OR REPLACE INTO participants
                (email, full_name, phone, organisation, category, academic, industry,
                 attend_dinner, conference_checkin, door_gift_collected, dinner_confirmed,
                 dinner_checkin, table_number, seat_number, registration_source, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    email,
                    r.get("Full_Name", ""),
                    r.get("Phone", ""),
                    r.get("Organisation", ""),
                    r.get("Category", ""),
                    r.get("Academic", ""),
                    r.get("Industry", ""),
                    r.get("Attend_Dinner", ""),
                    r.get("Conference_CheckIn", "No"),
                    r.get("DoorGift_Collected", "No"),
                    r.get("Dinner_Confirmed", "No"),
                    r.get("Dinner_CheckIn", "No"),
                    r.get("Table_Number", ""),
                    r.get("Seat_Number", ""),
                    r.get("Registration_Source", "Preloaded"),
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ))

    c.commit()
    c.close()
    log_action("Uploaded master Excel")


def get_participant(email):
    email = email.strip().lower()
    df = df_query("SELECT * FROM participants WHERE email=?", (email,))
    if df.empty:
        return None
    return df.iloc[0].to_dict()


def register_participant(data):
    execute("""
    INSERT OR REPLACE INTO participants
    (email, full_name, phone, organisation, category, academic, industry,
     attend_dinner, dinner_confirmed, conference_checkin, door_gift_collected,
     dinner_checkin, registration_source, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'No', 'No', 'No', ?, ?, ?)
    """, (
        data["email"].strip().lower(),
        data["full_name"],
        data["phone"],
        data["organisation"],
        data["category"],
        data["academic"],
        data["industry"],
        data["attend_dinner"],
        data["attend_dinner"],
        data["source"],
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ))
    log_action("Registered participant", data["email"])


def auto_assign_table(email):
    p = get_participant(email)
    if not p:
        return

    if p.get("table_number"):
        return

    category = str(p.get("category", "")).lower()

    if "media" in category:
        table_range = ["2"]
    else:
        table_range = [str(i) for i in range(5, 29)]

    for t in table_range:
        current = df_query("SELECT COUNT(*) AS n FROM participants WHERE table_number=?", (t,))
        n = int(current.iloc[0]["n"])
        if n < 10:
            execute(
                "UPDATE participants SET table_number=?, seat_number=?, updated_at=? WHERE email=?",
                (t, str(n + 1), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), email)
            )
            log_action(f"Assigned dinner table {t}", email)
            return


# =========================
# UI HELPERS
# =========================
def nav():
    pages = [
        "Home", "Register", "Academic", "Industry",
        "Gala Dinner", "Abstract Book", "My Status", "Admin"
    ]
    cols = st.columns(len(pages))
    if "page" not in st.session_state:
        st.session_state.page = "Home"

    for i, p in enumerate(pages):
        with cols[i]:
            if st.button(p, use_container_width=True):
                st.session_state.page = p


def hero():
    st.markdown("""
    <div class="hero">
        <h1>NICHE 2026</h1>
        <h2 class="gold">International Conference Registration & Event Management System</h2>
        <p class="small">
        Academic Conference • Industrial Conference • Gala Dinner • Abstract Viewer • Check-In • Door Gift • Dinner Table
        </p>
    </div>
    """, unsafe_allow_html=True)


def kpi(label, value):
    st.markdown(f"""
    <div class="metric-card">
        <div class="small">{label}</div>
        <h2 class="gold">{value}</h2>
    </div>
    """, unsafe_allow_html=True)


def save_upload(uploaded, name):
    if uploaded:
        path = UPLOAD_DIR / name
        path.write_bytes(uploaded.getbuffer())
        return str(path)
    return None


# =========================
# PAGES
# =========================
def page_home():
    hero()
    st.write("")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.subheader("Academic & Industrial Conference")
        poster = UPLOAD_DIR / "conference_poster"
        found = None
        for ext in [".jpg", ".jpeg", ".png"]:
            p = UPLOAD_DIR / f"conference_poster{ext}"
            if p.exists():
                found = p
        if found:
            st.image(str(found), use_container_width=True)
        else:
            st.info("Conference poster belum diupload.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.subheader("Gala Dinner")
        found = None
        for ext in [".jpg", ".jpeg", ".png"]:
            p = UPLOAD_DIR / f"dinner_poster{ext}"
            if p.exists():
                found = p
        if found:
            st.image(str(found), use_container_width=True)
        else:
            st.info("Dinner poster belum diupload.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.write("")
    total = df_query("SELECT COUNT(*) AS n FROM participants").iloc[0]["n"]
    checked = df_query("SELECT COUNT(*) AS n FROM participants WHERE conference_checkin='Yes'").iloc[0]["n"]
    gifts = df_query("SELECT COUNT(*) AS n FROM participants WHERE door_gift_collected='Yes'").iloc[0]["n"]
    dinner = df_query("SELECT COUNT(*) AS n FROM participants WHERE attend_dinner='Yes' OR dinner_confirmed='Yes'").iloc[0]["n"]

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi("Registered", total)
    with c2: kpi("Conference Check-In", checked)
    with c3: kpi("Door Gift Collected", gifts)
    with c4: kpi("Dinner Registered", dinner)


def page_register():
    st.header("Registration")
    st.markdown('<div class="glass">', unsafe_allow_html=True)

    email = st.text_input("Enter Email").strip().lower()

    if st.button("Continue", use_container_width=True):
        if not email:
            st.warning("Masukkan email dahulu.")
        else:
            st.session_state.reg_email = email
            st.session_state.reg_checked = True

    if st.session_state.get("reg_checked"):
        email = st.session_state.reg_email
        p = get_participant(email)

        if p:
            st.success("Email dijumpai dalam sistem.")
            st.write(f"**Name:** {p.get('full_name','')}")
            st.write(f"**Organisation:** {p.get('organisation','')}")
            st.write(f"**Category:** {p.get('category','')}")
            st.write(f"**Attend Dinner:** {p.get('attend_dinner','')}")
        else:
            st.warning("Email belum wujud. Sila lengkapkan pendaftaran baru.")

            full_name = st.text_input("Full Name *")
            phone = st.text_input("Phone")
            organisation = st.text_input("Organisation")
            category = st.selectbox(
                "Category",
                ["Academic Presenter", "Academic Participant", "Industry Speaker",
                 "Industry Participant", "Sponsor", "Exhibitor", "Government Agency", "Media", "Committee"]
            )

            c1, c2 = st.columns(2)
            with c1:
                academic = st.selectbox("Attend Academic Conference?", ["No", "Yes"])
            with c2:
                industry = st.selectbox("Attend Industrial Conference?", ["No", "Yes"])

            attend_dinner = st.selectbox("Attend Gala Dinner?", ["No", "Yes"])

            if st.button("Register Now", use_container_width=True):
                if not full_name:
                    st.error("Nama wajib diisi.")
                else:
                    register_participant({
                        "email": email,
                        "full_name": full_name,
                        "phone": phone,
                        "organisation": organisation,
                        "category": category,
                        "academic": academic,
                        "industry": industry,
                        "attend_dinner": attend_dinner,
                        "source": "Online"
                    })
                    if attend_dinner == "Yes":
                        auto_assign_table(email)
                    st.success("Registration successful.")
                    st.balloons()

    st.markdown("</div>", unsafe_allow_html=True)


def page_academic():
    st.header("Academic Conference")
    df = df_query("SELECT * FROM academic_programme")

    if df.empty:
        st.info("Academic programme belum ada. Upload Excel master di Admin.")
        return

    venue = st.selectbox("Venue", ["All"] + sorted(df["venue"].dropna().astype(str).unique().tolist()))
    q = st.text_input("Search paper / presenter / title")

    view = df.copy()
    if venue != "All":
        view = view[view["venue"].astype(str) == venue]

    if q:
        ql = q.lower()
        mask = view.astype(str).apply(lambda col: col.str.lower().str.contains(ql, na=False)).any(axis=1)
        view = view[mask]

    for _, r in view.iterrows():
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.subheader(f"{r.get('paper_id','')} — {r.get('title','')}")
        st.write(f"**Venue:** {r.get('venue','')}")
        st.write(f"**Time:** {r.get('time','')}")
        st.write(f"**Session:** {r.get('session','')}")
        st.write(f"**Theme:** {r.get('theme','')}")
        st.write(f"**Presenter:** {r.get('presenter','')}")
        pid = str(r.get("paper_id", ""))
        absdf = df_query("SELECT * FROM abstracts WHERE paper_id=?", (pid,))
        if not absdf.empty:
            with st.expander("View Abstract"):
                a = absdf.iloc[0]
                st.write(a.get("abstract_text", ""))
                st.caption(f"Keywords: {a.get('keywords','')}")
        st.markdown("</div>", unsafe_allow_html=True)


def page_industry():
    st.header("Industrial Conference")
    df = df_query("SELECT * FROM industry_programme")

    if df.empty:
        st.info("Industry programme belum ada. Upload Excel master di Admin.")
        return

    day = st.selectbox("Day", ["All"] + sorted(df["day"].dropna().astype(str).unique().tolist()))
    q = st.text_input("Search session / speaker / organisation")

    view = df.copy()
    if day != "All":
        view = view[view["day"].astype(str) == day]

    if q:
        ql = q.lower()
        mask = view.astype(str).apply(lambda col: col.str.lower().str.contains(ql, na=False)).any(axis=1)
        view = view[mask]

    for _, r in view.iterrows():
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.subheader(r.get("session", ""))
        st.write(f"**Day:** {r.get('day','')}")
        st.write(f"**Time:** {r.get('time','')}")
        st.write(f"**Venue:** {r.get('venue','')}")
        st.write(f"**Speaker:** {r.get('speaker','')}")
        st.write(f"**Organisation:** {r.get('organisation','')}")
        st.markdown("</div>", unsafe_allow_html=True)


def page_dinner():
    st.header("Gala Dinner")
    df = df_query("SELECT * FROM gala_dinner_programme")

    if not df.empty:
        st.subheader("Programme")
        for _, r in df.iterrows():
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.write(f"**{r.get('time','')}**")
            st.write(r.get("event", ""))
            st.markdown("</div>", unsafe_allow_html=True)

    st.subheader("Check Dinner Table")
    email = st.text_input("Enter your email for dinner table")
    if st.button("Check Table", use_container_width=True):
        p = get_participant(email)
        if not p:
            st.error("Email not found.")
        elif str(p.get("attend_dinner", "")) != "Yes" and str(p.get("dinner_confirmed", "")) != "Yes":
            st.warning("You are not registered for dinner.")
        else:
            if not p.get("table_number"):
                auto_assign_table(email)
                p = get_participant(email)
            st.success(f"Table {p.get('table_number')} — Seat {p.get('seat_number')}")


def page_abstract():
    st.header("Abstract Book")
    df = df_query("SELECT * FROM abstracts")

    if df.empty:
        st.info("Abstract belum ada. Upload Excel master di Admin.")
        return

    q = st.text_input("Search title / presenter / keyword / paper id")

    view = df.copy()
    if q:
        ql = q.lower()
        mask = view.astype(str).apply(lambda col: col.str.lower().str.contains(ql, na=False)).any(axis=1)
        view = view[mask]

    st.write(f"Total abstract: {len(view)}")

    for _, r in view.iterrows():
        with st.expander(f"{r.get('paper_id','')} — {r.get('title','')}"):
            st.write(f"**Presenter:** {r.get('presenter','')}")
            st.write(f"**Email:** {r.get('email','')}")
            st.write(r.get("abstract_text", ""))
            st.caption(f"Keywords: {r.get('keywords','')}")


def page_status():
    st.header("My Registration Status")
    email = st.text_input("Enter Email")
    if st.button("Check My Status", use_container_width=True):
        p = get_participant(email)
        if not p:
            st.error("Email not found.")
        else:
            c1, c2, c3 = st.columns(3)
            with c1:
                kpi("Conference Check-In", p.get("conference_checkin", "No"))
            with c2:
                kpi("Door Gift", p.get("door_gift_collected", "No"))
            with c3:
                kpi("Dinner", p.get("attend_dinner", "No"))

            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.write(f"**Name:** {p.get('full_name','')}")
            st.write(f"**Email:** {p.get('email','')}")
            st.write(f"**Organisation:** {p.get('organisation','')}")
            st.write(f"**Category:** {p.get('category','')}")
            st.write(f"**Table:** {p.get('table_number','Not Assigned')}")
            st.write(f"**Seat:** {p.get('seat_number','Not Assigned')}")
            st.markdown("</div>", unsafe_allow_html=True)


def page_admin():
    st.header("Admin Panel")

    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False

    if not st.session_state.admin_ok:
        pwd = st.text_input("Admin Password", type="password")
        if st.button("Login", use_container_width=True):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_ok = True
                st.success("Admin access granted.")
                st.rerun()
            else:
                st.error("Wrong password.")
        return

    st.success("Admin Mode Active")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Upload", "Check-In", "Door Gift", "Dinner", "Data Control"
    ])

    with tab1:
        st.subheader("Upload Master Excel")
        master = st.file_uploader("Upload NICHE2026_MASTER.xlsx", type=["xlsx"])
        if master and st.button("Process Excel", use_container_width=True):
            upload_master_excel(master)
            st.success("Master Excel uploaded and processed.")

        st.subheader("Upload Posters")
        cp = st.file_uploader("Conference Poster", type=["jpg", "jpeg", "png"])
        if cp:
            ext = "." + cp.name.split(".")[-1].lower()
            save_upload(cp, f"conference_poster{ext}")
            st.success("Conference poster uploaded.")

        dp = st.file_uploader("Dinner Poster", type=["jpg", "jpeg", "png"])
        if dp:
            ext = "." + dp.name.split(".")[-1].lower()
            save_upload(dp, f"dinner_poster{ext}")
            st.success("Dinner poster uploaded.")

    with tab2:
        st.subheader("Conference Check-In")
        email = st.text_input("Search Email for Check-In", key="checkin_email")
        if st.button("Find Participant", key="find_checkin"):
            p = get_participant(email)
            if not p:
                st.error("Not found.")
            else:
                st.session_state.checkin_target = p

        if "checkin_target" in st.session_state:
            p = st.session_state.checkin_target
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.write(f"**Name:** {p.get('full_name','')}")
            st.write(f"**Email:** {p.get('email','')}")
            st.write(f"**Check-In:** {p.get('conference_checkin','No')}")
            if st.button("Confirm Conference Check-In", use_container_width=True):
                execute(
                    "UPDATE participants SET conference_checkin='Yes', conference_checkin_time=?, updated_at=? WHERE email=?",
                    (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     p["email"])
                )
                log_action("Conference check-in", p["email"])
                st.success("Checked in.")
                del st.session_state.checkin_target
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    with tab3:
        st.subheader("Door Gift Collection")
        email = st.text_input("Search Email for Door Gift", key="gift_email")
        if st.button("Find Participant", key="find_gift"):
            p = get_participant(email)
            if not p:
                st.error("Not found.")
            else:
                st.session_state.gift_target = p

        if "gift_target" in st.session_state:
            p = st.session_state.gift_target
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.write(f"**Name:** {p.get('full_name','')}")
            st.write(f"**Door Gift:** {p.get('door_gift_collected','No')}")

            if p.get("door_gift_collected") == "Yes":
                st.error(f"Already collected at {p.get('door_gift_time','')}")
            else:
                if st.button("Mark Door Gift Collected", use_container_width=True):
                    execute(
                        "UPDATE participants SET door_gift_collected='Yes', door_gift_time=?, updated_at=? WHERE email=?",
                        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                         p["email"])
                    )
                    log_action("Door gift collected", p["email"])
                    st.success("Door gift marked as collected.")
                    del st.session_state.gift_target
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    with tab4:
        st.subheader("Dinner Check-In & Table Assignment")
        email = st.text_input("Search Email for Dinner", key="dinner_email")
        if st.button("Find Dinner Guest", key="find_dinner"):
            p = get_participant(email)
            if not p:
                st.error("Not found.")
            else:
                st.session_state.dinner_target = p

        if "dinner_target" in st.session_state:
            p = st.session_state.dinner_target
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.write(f"**Name:** {p.get('full_name','')}")
            st.write(f"**Dinner:** {p.get('attend_dinner','No')}")
            st.write(f"**Table:** {p.get('table_number','')}")
            st.write(f"**Seat:** {p.get('seat_number','')}")

            if st.button("Auto Assign Table", use_container_width=True):
                auto_assign_table(p["email"])
                st.success("Table assigned.")
                del st.session_state.dinner_target
                st.rerun()

            if st.button("Dinner Check-In", use_container_width=True):
                execute(
                    "UPDATE participants SET dinner_checkin='Yes', dinner_checkin_time=?, updated_at=? WHERE email=?",
                    (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     p["email"])
                )
                log_action("Dinner check-in", p["email"])
                st.success("Dinner checked in.")
                del st.session_state.dinner_target
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        st.subheader("Dinner Table Overview")
        table_df = df_query("""
            SELECT table_number, COUNT(*) AS occupied
            FROM participants
            WHERE table_number IS NOT NULL AND table_number != ''
            GROUP BY table_number
            ORDER BY CAST(table_number AS INTEGER)
        """)
        st.dataframe(table_df, use_container_width=True)

    with tab5:
        st.subheader("Data Control")

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            if st.button("Clear Check-In", use_container_width=True):
                execute("UPDATE participants SET conference_checkin='No', conference_checkin_time=NULL")
                log_action("Clear check-in")
                st.success("Check-in cleared.")

        with c2:
            if st.button("Clear Door Gift", use_container_width=True):
                execute("UPDATE participants SET door_gift_collected='No', door_gift_time=NULL")
                log_action("Clear door gift")
                st.success("Door gift cleared.")

        with c3:
            if st.button("Clear Dinner Check-In", use_container_width=True):
                execute("UPDATE participants SET dinner_checkin='No', dinner_checkin_time=NULL")
                log_action("Clear dinner check-in")
                st.success("Dinner check-in cleared.")

        with c4:
            if st.button("Factory Reset Participants", use_container_width=True):
                execute("DELETE FROM participants")
                log_action("Factory reset participants")
                st.success("All participants deleted.")

        st.subheader("Download Participants Backup")
        df = df_query("SELECT * FROM participants")
        st.download_button(
            "Download Participants CSV",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name="participants_backup.csv",
            mime="text/csv",
            use_container_width=True
        )

        st.subheader("Participants")
        st.dataframe(df, use_container_width=True)


# =========================
# ROUTER
# =========================
nav()
st.write("")

page = st.session_state.page

if page == "Home":
    page_home()
elif page == "Register":
    page_register()
elif page == "Academic":
    page_academic()
elif page == "Industry":
    page_industry()
elif page == "Gala Dinner":
    page_dinner()
elif page == "Abstract Book":
    page_abstract()
elif page == "My Status":
    page_status()
elif page == "Admin":
    page_admin()
