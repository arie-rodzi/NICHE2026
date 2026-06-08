"""
NICHE 2026 — Streamlit Self Check-In System
===========================================
Flow sebenar:
1. Participant masukkan email.
2. Sistem cari rekod dalam database.
3. Jika jumpa, keluar details participant.
4. Participant pilih sama ada hadir Gala Dinner atau tidak.
5. Jika hadir dinner, sistem auto-assign table participant yang masih kosong.
6. Staff/admin di kaunter tick:
   - Conference Check-In
   - Door Gift Collected
   - Dinner Check-In
7. Public boleh lihat tentative Academic, Industry, Gala Dinner dan Abstract.
8. Admin boleh upload Excel, reset/import data, dan lihat participant list.

requirements.txt:
streamlit
pandas
openpyxl
"""

import sqlite3
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st


# =========================================================
# BASIC CONFIG
# =========================================================
st.set_page_config(
    page_title="NICHE 2026 Self Check-In",
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

.stApp {
    background:
        radial-gradient(circle at top left, rgba(255, 209, 102, .22), transparent 30%),
        radial-gradient(circle at top right, rgba(59, 130, 246, .20), transparent 34%),
        linear-gradient(135deg, #07111f 0%, #0d1b2a 45%, #102a43 100%);
    color: #f8fafc;
}

.block-container {
    padding-top: 1.2rem;
    max-width: 1450px;
}

.hero {
    padding: 24px 28px;
    border-radius: 26px;
    background: linear-gradient(135deg, rgba(255,255,255,.15), rgba(255,255,255,.055));
    border: 1px solid rgba(255,255,255,.18);
    box-shadow: 0 25px 70px rgba(0,0,0,.30);
    margin-bottom: 16px;
}

.hero h1 {
    margin: 0;
    font-size: 2rem;
    color: #ffffff;
    letter-spacing: -.03em;
}

.hero p {
    margin-top: 8px;
    color: #dbeafe;
}

.gold {
    color: #ffd166;
    font-weight: 900;
}

.card {
    padding: 18px 20px;
    border-radius: 22px;
    background: rgba(255,255,255,.10);
    border: 1px solid rgba(255,255,255,.15);
    box-shadow: 0 16px 45px rgba(0,0,0,.22);
    margin-bottom: 14px;
}

.success-card {
    padding: 18px 20px;
    border-radius: 22px;
    background: linear-gradient(135deg, rgba(34,197,94,.22), rgba(255,255,255,.08));
    border: 1px solid rgba(134,239,172,.35);
}

.warning-card {
    padding: 18px 20px;
    border-radius: 22px;
    background: linear-gradient(135deg, rgba(245,158,11,.24), rgba(255,255,255,.08));
    border: 1px solid rgba(253,230,138,.35);
}

.metric-card {
    padding: 16px;
    border-radius: 20px;
    background: linear-gradient(145deg, rgba(255,255,255,.16), rgba(255,255,255,.06));
    border: 1px solid rgba(255,255,255,.17);
}

.metric-card .label {
    color: #bfdbfe;
    font-size: .82rem;
    font-weight: 800;
    text-transform: uppercase;
}

.metric-card .value {
    color: #ffffff;
    font-size: 1.8rem;
    font-weight: 950;
    margin-top: 3px;
}

div.stButton > button {
    border-radius: 14px;
    font-weight: 850;
    border: 1px solid rgba(255,255,255,.2);
    background: linear-gradient(135deg, #ffd166 0%, #d99b20 100%);
    color: #111827;
}

div.stButton > button:hover {
    filter: brightness(1.08);
    border: 1px solid rgba(255,255,255,.45);
}

[data-testid="stDataFrame"] {
    background: white;
    border-radius: 18px;
    overflow: hidden;
}

hr {
    border: none;
    border-top: 1px solid rgba(255,255,255,.16);
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# SESSION
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "Self Check-In"
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "current_email" not in st.session_state:
    st.session_state.current_email = ""


# =========================================================
# DATABASE
# =========================================================
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def clean(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    s = str(v).strip()
    if s.lower() in ["", "nan", "nat", "none"]:
        return None
    return s


def yes(v):
    return 1 if str(v).strip().lower() in ["yes", "y", "true", "1", "ya", "hadir"] else 0


def init_db():
    conn = get_db()
    conn.executescript(
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
            created_at          TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at          TEXT
        );

        CREATE TABLE IF NOT EXISTS academic_programme (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            venue TEXT, time TEXT, session TEXT, moderator TEXT,
            theme TEXT, paper_id TEXT, title TEXT, presenter TEXT, email TEXT,
            sort_order INTEGER
        );

        CREATE TABLE IF NOT EXISTS abstracts (
            paper_id TEXT PRIMARY KEY,
            title TEXT, presenter TEXT, email TEXT,
            venue TEXT, time TEXT, session TEXT,
            keywords TEXT, abstract_text TEXT, authors TEXT
        );

        CREATE TABLE IF NOT EXISTS industry_programme (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            day TEXT, time TEXT, venue TEXT, session TEXT,
            speaker TEXT, organisation TEXT, details TEXT,
            sort_order INTEGER
        );

        CREATE TABLE IF NOT EXISTS dinner_programme (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            time TEXT, event TEXT, sort_order INTEGER
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT DEFAULT CURRENT_TIMESTAMP,
            action TEXT, target_email TEXT, detail TEXT
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


def clear_all():
    conn = get_db()
    for t in [
        "participants",
        "academic_programme",
        "abstracts",
        "industry_programme",
        "dinner_programme",
        "audit_log",
    ]:
        conn.execute(f"DELETE FROM {t}")
    conn.commit()
    conn.close()


def seed_from_excel(path, clear_existing=True):
    if clear_existing:
        clear_all()

    xl = pd.ExcelFile(path)
    conn = get_db()

    if "Participants" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Participants")
        for _, r in df.iterrows():
            email = clean(r.get("Email"))
            if not email:
                continue
            table_no = None
            if "Table_Number" in df.columns and clean(r.get("Table_Number")) is not None:
                try:
                    table_no = int(float(r.get("Table_Number")))
                except Exception:
                    table_no = None

            conn.execute(
                """
                INSERT OR IGNORE INTO participants
                (email, full_name, phone, organisation, category,
                 academic, industry, conference_checkin, doorgift_collected,
                 attend_dinner, dinner_checkin, table_number, registration_source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    email.lower(),
                    clean(r.get("Full_Name")) or clean(r.get("Name")) or "—",
                    clean(r.get("Phone")),
                    clean(r.get("Organisation")) or clean(r.get("Organization")),
                    clean(r.get("Category")) or clean(r.get("Type")),
                    yes(r.get("Academic")),
                    yes(r.get("Industry")),
                    yes(r.get("Conference_CheckIn")),
                    yes(r.get("DoorGift_Collected")),
                    yes(r.get("Attend_Dinner")),
                    yes(r.get("Dinner_CheckIn")),
                    table_no,
                    clean(r.get("Registration_Source")) or "Preloaded",
                ),
            )

    if "Academic_Programme" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Academic_Programme")
        for i, r in df.iterrows():
            conn.execute(
                """
                INSERT INTO academic_programme
                (venue, time, session, moderator, theme, paper_id, title, presenter, email, sort_order)
                VALUES (?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    clean(r.get("Venue")),
                    clean(r.get("Time")),
                    clean(r.get("Session")),
                    clean(r.get("Moderator")),
                    clean(r.get("Theme")),
                    clean(r.get("Paper_ID")),
                    clean(r.get("Title")),
                    clean(r.get("Presenter")),
                    clean(r.get("Email_From_Abstract")) or clean(r.get("Email")),
                    i,
                ),
            )

    if "Abstracts" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Abstracts")
        for _, r in df.iterrows():
            pid = clean(r.get("Paper_ID"))
            if not pid:
                continue
            conn.execute(
                """
                INSERT OR REPLACE INTO abstracts
                (paper_id, title, presenter, email, venue, time, session, keywords, abstract_text, authors)
                VALUES (?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    pid,
                    clean(r.get("Title")),
                    clean(r.get("Presenter")),
                    clean(r.get("Email")),
                    clean(r.get("Venue")),
                    clean(r.get("Time")),
                    clean(r.get("Session")),
                    clean(r.get("Keywords")),
                    clean(r.get("Abstract_Text")) or clean(r.get("Abstract")),
                    clean(r.get("Authors_Affiliation")) or clean(r.get("Authors")),
                ),
            )

    if "Industry_Programme" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Industry_Programme")
        for i, r in df.iterrows():
            conn.execute(
                """
                INSERT INTO industry_programme
                (day, time, venue, session, speaker, organisation, details, sort_order)
                VALUES (?,?,?,?,?,?,?,?)
                """,
                (
                    clean(r.get("Day")),
                    clean(r.get("Time")),
                    clean(r.get("Venue")),
                    clean(r.get("Session")),
                    clean(r.get("Speaker")),
                    clean(r.get("Organisation")) or clean(r.get("Organization")),
                    clean(r.get("Details")),
                    i,
                ),
            )

    if "Gala_Dinner_Programme" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="Gala_Dinner_Programme")
        for i, r in df.iterrows():
            conn.execute(
                "INSERT INTO dinner_programme (time, event, sort_order) VALUES (?,?,?)",
                (clean(r.get("Time")), clean(r.get("Event")), i),
            )

    conn.commit()
    conn.close()
    log_action("import_excel", detail=Path(path).name)


# =========================================================
# CORE CHECK-IN LOGIC
# =========================================================
def get_participant_by_email(email):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM participants WHERE lower(email)=?",
        (email.strip().lower(),),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def table_occupancy():
    df = query_df(
        """
        SELECT table_number, COUNT(*) AS occupied
        FROM participants
        WHERE table_number IS NOT NULL
        GROUP BY table_number
        """
    )
    if df.empty:
        return {}
    return {int(r["table_number"]): int(r["occupied"]) for _, r in df.iterrows()}


def auto_assign_table(email):
    p = get_participant_by_email(email)
    if not p:
        return None, "Participant not found."

    if p.get("table_number"):
        return int(p["table_number"]), "Existing table retained."

    occ = table_occupancy()

    for t in PARTICIPANT_TABLES:
        used = occ.get(t, 0)
        if used < SEATS_PER_TABLE:
            exec_sql(
                """
                UPDATE participants
                SET attend_dinner=1, table_number=?, updated_at=?
                WHERE lower(email)=?
                """,
                (t, datetime.now().isoformat(timespec="seconds"), email.lower()),
            )
            log_action("auto_assign_table", email, f"table={t}")
            return t, f"Assigned to Table {t}."

    return None, "All participant dinner tables are full."


def update_dinner_choice(email, attend):
    if attend:
        return auto_assign_table(email)

    exec_sql(
        """
        UPDATE participants
        SET attend_dinner=0, dinner_checkin=0, table_number=NULL, updated_at=?
        WHERE lower(email)=?
        """,
        (datetime.now().isoformat(timespec="seconds"), email.lower()),
    )
    log_action("dinner_declined", email)
    return None, "Dinner set to No. Table cleared."


def toggle_field(pid, email, field):
    allowed = ["conference_checkin", "doorgift_collected", "dinner_checkin"]
    if field not in allowed:
        return
    p = query_df("SELECT * FROM participants WHERE id=?", (pid,))
    if p.empty:
        return
    current = int(p.iloc[0][field] or 0)
    new_val = 0 if current == 1 else 1
    exec_sql(
        f"UPDATE participants SET {field}=?, updated_at=? WHERE id=?",
        (new_val, datetime.now().isoformat(timespec="seconds"), pid),
    )
    log_action(f"toggle_{field}_{new_val}", email)


# =========================================================
# UI HELPERS
# =========================================================
def hero():
    st.markdown(
        f"""
        <div class="hero">
            <h1>{EVENT_NAME} <span class="gold">Self Check-In System</span></h1>
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
    pages = ["Self Check-In", "Tentative", "Abstracts", "Admin"]
    cols = st.columns(len(pages))
    for col, p in zip(cols, pages):
        with col:
            if st.button(p, use_container_width=True, key=f"nav_{p}"):
                st.session_state.page = p
                st.rerun()


def stats():
    conn = get_db()
    out = {
        "total": conn.execute("SELECT COUNT(*) c FROM participants").fetchone()["c"],
        "checked": conn.execute("SELECT COUNT(*) c FROM participants WHERE conference_checkin=1").fetchone()["c"],
        "gift": conn.execute("SELECT COUNT(*) c FROM participants WHERE doorgift_collected=1").fetchone()["c"],
        "dinner": conn.execute("SELECT COUNT(*) c FROM participants WHERE attend_dinner=1").fetchone()["c"],
        "dinner_in": conn.execute("SELECT COUNT(*) c FROM participants WHERE dinner_checkin=1").fetchone()["c"],
        "papers": conn.execute("SELECT COUNT(*) c FROM abstracts").fetchone()["c"],
    }
    conn.close()
    return out


def yesno(v):
    return "✅ Yes" if int(v or 0) == 1 else "—"


# =========================================================
# PAGES
# =========================================================
def page_self_checkin():
    hero()

    s = stats()
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("Registered", s["total"])
    with c2: metric_card("Checked-In", s["checked"])
    with c3: metric_card("Door Gift", s["gift"])
    with c4: metric_card("Dinner", s["dinner"])

    st.markdown("## ✉️ Participant Self Check-In")
    st.markdown(
        """
        <div class="card">
        Masukkan email yang digunakan semasa pendaftaran. Sistem akan paparkan maklumat peserta.
        Selepas itu peserta pilih sama ada hadir Gala Dinner atau tidak.
        Jika pilih hadir, sistem akan assign meja participant secara automatik.
        </div>
        """,
        unsafe_allow_html=True,
    )

    email = st.text_input("Enter your registered email", value=st.session_state.current_email).strip().lower()

    if st.button("Search / Check My Details", use_container_width=True):
        st.session_state.current_email = email
        st.rerun()

    if st.session_state.current_email:
        p = get_participant_by_email(st.session_state.current_email)

        if not p:
            st.error("Email tidak dijumpai dalam senarai peserta. Sila rujuk admin kaunter.")
            return

        st.markdown("## 👤 Participant Details")
        st.markdown(
            f"""
            <div class="success-card">
            <h3>{p.get("full_name")}</h3>
            <p><b>Email:</b> {p.get("email")}<br>
            <b>Phone:</b> {p.get("phone") or "—"}<br>
            <b>Organisation:</b> {p.get("organisation") or "—"}<br>
            <b>Category:</b> {p.get("category") or "—"}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("### 🍽️ Gala Dinner Attendance")
        current = "Yes" if int(p.get("attend_dinner") or 0) == 1 else "No"
        choice = st.radio(
            "Will you attend the Gala Dinner?",
            ["Yes", "No"],
            index=0 if current == "Yes" else 1,
            horizontal=True,
        )

        if st.button("Confirm Dinner Choice", use_container_width=True):
            table, msg = update_dinner_choice(p["email"], choice == "Yes")
            if choice == "Yes" and table:
                st.success(f"Dinner confirmed. Your table is Table {table}.")
            elif choice == "Yes":
                st.warning(msg)
            else:
                st.info("Dinner choice saved as No.")
            st.rerun()

        p = get_participant_by_email(st.session_state.current_email)

        st.markdown("### ✅ Current Status")
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1: metric_card("Conference Check-In", yesno(p["conference_checkin"]))
        with sc2: metric_card("Door Gift", yesno(p["doorgift_collected"]))
        with sc3: metric_card("Dinner", yesno(p["attend_dinner"]))
        with sc4: metric_card("Table", p["table_number"] or "—")

        st.markdown(
            """
            <div class="warning-card">
            <b>Untuk staff/admin:</b> Selepas participant tunjuk skrin ini di kaunter,
            buka Admin → Staff Counter untuk tick check-in dan door gift menggunakan telefon staff.
            </div>
            """,
            unsafe_allow_html=True,
        )


def page_tentative():
    hero()
    st.markdown("## 🗓️ Conference Tentative")

    t1, t2, t3 = st.tabs(["Academic Conference", "Industry Conference", "Gala Dinner"])

    with t1:
        df = query_df("SELECT * FROM academic_programme ORDER BY sort_order")
        if df.empty:
            st.warning("Tiada data Academic_Programme.")
        else:
            venue_options = ["All"] + sorted([x for x in df["venue"].dropna().unique()])
            venue = st.selectbox("Venue", venue_options)
            if venue != "All":
                df = df[df["venue"] == venue]
            for v, g in df.groupby(df["venue"].fillna("—"), sort=False):
                st.markdown(f"### 📍 {v}")
                st.dataframe(
                    g[["time", "session", "moderator", "theme", "paper_id", "title", "presenter"]],
                    use_container_width=True,
                    hide_index=True,
                )

    with t2:
        df = query_df("SELECT * FROM industry_programme ORDER BY sort_order")
        if df.empty:
            st.warning("Tiada data Industry_Programme.")
        else:
            for d, g in df.groupby(df["day"].fillna("—"), sort=False):
                st.markdown(f"### {d}")
                st.dataframe(
                    g[["time", "venue", "session", "speaker", "organisation", "details"]],
                    use_container_width=True,
                    hide_index=True,
                )

    with t3:
        df = query_df("SELECT * FROM dinner_programme ORDER BY sort_order")
        if df.empty:
            st.warning("Tiada data Gala_Dinner_Programme.")
        else:
            st.dataframe(df[["time", "event"]], use_container_width=True, hide_index=True)

        st.markdown("### 🪑 Table Occupancy")
        occ = table_occupancy()
        cols = st.columns(len(PARTICIPANT_TABLES))
        for col, table in zip(cols, PARTICIPANT_TABLES):
            with col:
                metric_card(f"Table {table}", f"{occ.get(table, 0)}/{SEATS_PER_TABLE}")


def page_abstracts():
    hero()
    st.markdown("## 📝 Abstract Search")

    df = query_df("SELECT * FROM abstracts ORDER BY paper_id")
    if df.empty:
        st.warning("Tiada data Abstracts.")
        return

    q = st.text_input("Search by paper ID / title / presenter / keyword")
    if q:
        s = q.lower()
        df = df[
            df.fillna("").astype(str).apply(
                lambda row: row.str.lower().str.contains(s).any(), axis=1
            )
        ]

    st.write(f"Total abstracts: **{len(df)}**")

    for _, r in df.iterrows():
        with st.expander(f"{r.get('paper_id','')} — {r.get('title','Untitled')}"):
            st.write(f"**Presenter:** {r.get('presenter') or '—'}")
            st.write(f"**Email:** {r.get('email') or '—'}")
            st.write(f"**Venue / Time / Session:** {r.get('venue') or '—'} / {r.get('time') or '—'} / {r.get('session') or '—'}")
            st.write(f"**Authors/Affiliation:** {r.get('authors') or '—'}")
            st.write(f"**Keywords:** {r.get('keywords') or '—'}")
            st.write(r.get("abstract_text") or "—")


def page_admin():
    hero()
    st.markdown("## 🛡️ Admin")

    if not st.session_state.is_admin:
        pw = st.text_input("Admin Password", type="password")
        if st.button("Login", use_container_width=True):
            if pw == ADMIN_PASSWORD:
                st.session_state.is_admin = True
                log_action("admin_login")
                st.rerun()
            else:
                st.error("Password salah.")
        return

    if st.button("Logout"):
        st.session_state.is_admin = False
        st.rerun()

    tabs = st.tabs(["Staff Counter", "Participants", "Upload Data", "Audit Log"])

    with tabs[0]:
        st.markdown("### 📱 Staff Counter: Tick Check-In / Door Gift / Dinner Check-In")
        q = st.text_input("Search participant by email / name / organisation", key="counter_search")
        df = query_df("SELECT * FROM participants ORDER BY full_name")

        if q:
            s = q.lower()
            df = df[
                df.fillna("").astype(str).apply(
                    lambda row: row.str.lower().str.contains(s).any(), axis=1
                )
            ]

        if df.empty:
            st.warning("Tiada participant dijumpai.")
        else:
            for _, p in df.iterrows():
                with st.container(border=True):
                    left, right = st.columns([3, 1])
                    with left:
                        st.markdown(f"### {p['full_name']}")
                        st.write(f"**Email:** {p['email']}")
                        st.write(f"**Organisation:** {p.get('organisation') or '—'}")
                        st.write(f"**Category:** {p.get('category') or '—'}")
                    with right:
                        st.write(f"**Dinner:** {yesno(p['attend_dinner'])}")
                        st.write(f"**Table:** {p.get('table_number') or '—'}")

                    c1, c2, c3, c4 = st.columns(4)

                    with c1:
                        if st.button(f"Conference Check-In: {yesno(p['conference_checkin'])}", key=f"ci_{p['id']}", use_container_width=True):
                            toggle_field(int(p["id"]), p["email"], "conference_checkin")
                            st.rerun()

                    with c2:
                        if st.button(f"Door Gift: {yesno(p['doorgift_collected'])}", key=f"dg_{p['id']}", use_container_width=True):
                            toggle_field(int(p["id"]), p["email"], "doorgift_collected")
                            st.rerun()

                    with c3:
                        if int(p["attend_dinner"] or 0) == 1:
                            if st.button(f"Dinner Check-In: {yesno(p['dinner_checkin'])}", key=f"di_{p['id']}", use_container_width=True):
                                toggle_field(int(p["id"]), p["email"], "dinner_checkin")
                                st.rerun()
                        else:
                            st.info("Not attending dinner")

                    with c4:
                        table_opts = [0] + PARTICIPANT_TABLES
                        current = int(p["table_number"]) if pd.notna(p["table_number"]) and p["table_number"] else 0
                        table = st.selectbox(
                            "Table",
                            table_opts,
                            index=table_opts.index(current) if current in table_opts else 0,
                            key=f"table_{p['id']}",
                            format_func=lambda x: "No Table" if x == 0 else f"Table {x}",
                        )
                        if st.button("Save Table", key=f"save_table_{p['id']}", use_container_width=True):
                            if table == 0:
                                exec_sql(
                                    "UPDATE participants SET table_number=NULL, updated_at=? WHERE id=?",
                                    (datetime.now().isoformat(timespec="seconds"), int(p["id"])),
                                )
                            else:
                                occ = query_df(
                                    "SELECT COUNT(*) AS c FROM participants WHERE table_number=? AND id!=?",
                                    (table, int(p["id"])),
                                )["c"].iloc[0]
                                if occ >= SEATS_PER_TABLE:
                                    st.error(f"Table {table} penuh.")
                                    st.stop()
                                exec_sql(
                                    "UPDATE participants SET table_number=?, attend_dinner=1, updated_at=? WHERE id=?",
                                    (table, datetime.now().isoformat(timespec="seconds"), int(p["id"])),
                                )
                            log_action("manual_table_update", p["email"], f"table={table}")
                            st.rerun()

    with tabs[1]:
        st.markdown("### 👥 Participants List")
        df = query_df("SELECT * FROM participants ORDER BY created_at DESC, id DESC")
        if df.empty:
            st.warning("Tiada participants.")
        else:
            filter_col1, filter_col2, filter_col3 = st.columns(3)
            with filter_col1:
                category = st.selectbox("Category", ["All"] + sorted([x for x in df["category"].dropna().unique()]))
            with filter_col2:
                dinner = st.selectbox("Dinner", ["All", "Yes", "No"])
            with filter_col3:
                gift = st.selectbox("Door Gift", ["All", "Collected", "Not Yet"])

            fdf = df.copy()
            if category != "All":
                fdf = fdf[fdf["category"] == category]
            if dinner == "Yes":
                fdf = fdf[fdf["attend_dinner"] == 1]
            elif dinner == "No":
                fdf = fdf[fdf["attend_dinner"] == 0]
            if gift == "Collected":
                fdf = fdf[fdf["doorgift_collected"] == 1]
            elif gift == "Not Yet":
                fdf = fdf[fdf["doorgift_collected"] == 0]

            display = fdf.copy()
            for col in ["academic", "industry", "conference_checkin", "doorgift_collected", "attend_dinner", "dinner_checkin"]:
                display[col] = display[col].apply(yesno)

            st.dataframe(display, use_container_width=True, hide_index=True)

            csv = fdf.to_csv(index=False).encode("utf-8")
            st.download_button("Download Participants CSV", csv, "niche2026_participants.csv", "text/csv")

    with tabs[2]:
        st.markdown("### 📤 Upload Master Data")
        st.info("Upload Excel yang ada sheet: Participants, Academic_Programme, Abstracts, Industry_Programme, Gala_Dinner_Programme.")
        uploaded = st.file_uploader("Upload Excel", type=["xlsx"])

        if uploaded is not None:
            EXCEL_PATH.write_bytes(uploaded.getbuffer())
            if st.button("Import Excel and Replace Current Data", use_container_width=True):
                try:
                    seed_from_excel(EXCEL_PATH, clear_existing=True)
                    st.success("Data berjaya diimport.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal import Excel: {e}")

        st.markdown("---")
        st.warning("Reset akan padam semua data dalam sistem.")
        confirm = st.checkbox("Saya faham dan mahu reset semua data.")
        if confirm and st.button("RESET ALL DATA", use_container_width=True):
            clear_all()
            st.success("Semua data telah dipadam.")
            st.rerun()

    with tabs[3]:
        st.markdown("### 🧾 Audit Log")
        df = query_df("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500")
        st.dataframe(df, use_container_width=True, hide_index=True)


# =========================================================
# MAIN
# =========================================================
init_db()
nav()

page = st.session_state.page

if page == "Self Check-In":
    page_self_checkin()
elif page == "Tentative":
    page_tentative()
elif page == "Abstracts":
    page_abstracts()
elif page == "Admin":
    page_admin()
else:
    page_self_checkin()
