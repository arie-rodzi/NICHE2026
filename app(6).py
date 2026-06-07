
# NICHE 2026 Conference Registration System
# Mobile-first Streamlit app with:
# - Admin-only upload center
# - One Master Excel upload
# - Poster upload
# - No sidebar
# - No PDF embed; download only
# - Academic schedule cards + abstract in expander
# - Walk-in registration for non-presenters
# - Check-in + door gift
# - Gala dinner RSVP + table assignment
# - Built-in seating layout
# - Reset buttons

import base64
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

DB_PATH = Path("niche2026.db")
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

ADMIN_USER = "admin"
ADMIN_PASS = "NICHE2026admin"

st.set_page_config(
    page_title="NICHE 2026",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
[data-testid="stSidebar"], [data-testid="collapsedControl"] {display:none !important;}
#MainMenu, footer, header {visibility:hidden;}
html, body, [data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg,#020B2D 0%,#061A52 55%,#0D2D73 100%) !important;
    color: white !important;
    font-family: "Inter","Segoe UI",Arial,sans-serif;
}
.block-container {
    padding-top: .8rem !important;
    padding-left: .7rem !important;
    padding-right: .7rem !important;
    max-width: 1100px !important;
}
h1,h2,h3,h4,p,label,span,div { color: inherit; }
.poster-img img {
    border-radius: 24px;
    box-shadow: 0 18px 40px rgba(0,0,0,.45);
    border: 1px solid rgba(212,175,55,.45);
}
.niche-title {
    font-size: clamp(2.1rem, 8vw, 5rem);
    line-height: .95;
    font-weight: 900;
    color: #F7D774;
    letter-spacing: -1px;
    margin: .2rem 0;
    text-shadow: 0 6px 24px rgba(0,0,0,.55);
}
.niche-sub {
    color: #FFFFFF;
    font-size: clamp(1rem, 3.5vw, 1.35rem);
    font-weight: 700;
    line-height:1.35;
}
.gold-line{
    height:4px;width:90px;
    background:linear-gradient(90deg,#D4AF37,#F7D774);
    border-radius:99px;
    margin:.6rem 0 1rem 0;
}
.card {
    background:#FFFFFF;
    color:#061A52;
    border-radius:22px;
    padding:18px;
    margin:12px 0;
    box-shadow:0 14px 32px rgba(0,0,0,.25);
    border:1px solid rgba(212,175,55,.45);
}
.card * { color:#061A52 !important; }
.dark-card {
    background:linear-gradient(135deg,#071A4F,#0D2D73);
    border:1px solid rgba(247,215,116,.45);
    color:white;
    border-radius:24px;
    padding:18px;
    margin:12px 0;
    box-shadow:0 14px 34px rgba(0,0,0,.35);
}
.gold-card {
    background:linear-gradient(135deg,#D4AF37,#F7D774);
    color:#061A52;
    border-radius:24px;
    padding:18px;
    margin:12px 0;
    box-shadow:0 14px 34px rgba(0,0,0,.25);
    font-weight:800;
}
.gold-card * { color:#061A52 !important; }
.section-title {
    display:inline-block;
    color:#061A52;
    background:linear-gradient(135deg,#D4AF37,#F7D774);
    border-radius:16px;
    padding:10px 16px;
    margin:18px 0 10px 0;
    font-size:1.35rem;
    font-weight:900;
    box-shadow:0 8px 20px rgba(0,0,0,.25);
}
.session-header {
    background:linear-gradient(135deg,#D4AF37,#F7D774);
    color:#061A52;
    border-radius:20px;
    padding:16px;
    margin-top:18px;
    box-shadow:0 10px 24px rgba(0,0,0,.25);
}
.session-header * { color:#061A52 !important; }
.paper-card {
    background:#FFFFFF;
    color:#061A52;
    border-radius:20px;
    padding:16px;
    margin:12px 0;
    border-left:8px solid #D4AF37;
    box-shadow:0 10px 26px rgba(0,0,0,.22);
}
.paper-card * { color:#061A52 !important; }
.paper-title {
    font-size:1.02rem;
    font-weight:900;
    line-height:1.25;
    color:#061A52 !important;
}
.meta-pill {
    display:inline-block;
    padding:6px 10px;
    border-radius:999px;
    margin:3px 3px 3px 0;
    background:#EAF2FF;
    color:#0D2D73 !important;
    font-size:.82rem;
    font-weight:800;
}
.status-ok {background:#DCFCE7;color:#166534!important;}
.status-warn {background:#FEF3C7;color:#92400E!important;}
.nav-grid {
    display:grid;
    grid-template-columns:repeat(2,minmax(0,1fr));
    gap:12px;
    margin:18px 0;
}
.nav-card {
    background:linear-gradient(135deg,#FFFFFF,#FFF8E7);
    color:#061A52!important;
    border-radius:22px;
    padding:18px 14px;
    min-height:96px;
    box-shadow:0 12px 30px rgba(0,0,0,.28);
    border:1px solid rgba(212,175,55,.5);
}
.nav-card b { color:#061A52!important;font-size:1rem; }
.nav-card small { color:#334!important; }
.stButton>button, .stDownloadButton>button {
    width:100%;
    background:linear-gradient(135deg,#D4AF37,#F7D774)!important;
    color:#061A52!important;
    border:0!important;
    border-radius:999px!important;
    padding:.75rem 1rem!important;
    font-weight:900!important;
    box-shadow:0 10px 26px rgba(0,0,0,.25);
}
.table-badge {
    display:inline-flex;
    align-items:center;
    justify-content:center;
    width:56px;height:56px;
    border-radius:50%;
    margin:6px;
    background:#E5E7EB;
    color:#061A52;
    font-weight:900;
    border:3px solid #8B7A64;
}
.table-highlight{
    background:linear-gradient(135deg,#D4AF37,#F7D774)!important;
    border:4px solid #061A52!important;
    transform:scale(1.08);
}
.zone-box{
    background:#FFFFFF;
    color:#061A52;
    padding:12px;
    border-radius:20px;
    margin:12px 0;
    box-shadow:0 10px 26px rgba(0,0,0,.18);
}
.zone-title{
    display:inline-block;
    padding:8px 12px;
    border-radius:12px;
    color:white;
    font-weight:900;
    letter-spacing:.5px;
}
.bottom-space{height:80px;}
@media(max-width:640px){
  .block-container{padding-left:.55rem!important;padding-right:.55rem!important;}
  .card,.dark-card,.gold-card{padding:14px;border-radius:20px;}
  .nav-grid{gap:9px;}
  .nav-card{min-height:88px;padding:14px 12px;}
}
</style>
""", unsafe_allow_html=True)


def conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    with conn() as con:
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS kv (k TEXT PRIMARY KEY, v TEXT)")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS participants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            registration_id TEXT,
            full_name TEXT,
            email TEXT UNIQUE,
            category TEXT,
            institution TEXT,
            phone TEXT,
            faculty TEXT,
            participant_type TEXT,
            attend_dinner TEXT DEFAULT 'No',
            dinner_table TEXT,
            conference_table TEXT,
            checked_in INTEGER DEFAULT 0,
            door_gift INTEGER DEFAULT 0,
            dinner_checked_in INTEGER DEFAULT 0,
            is_walkin INTEGER DEFAULT 0,
            created_at TEXT
        )""")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS academic (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paper_id TEXT,
            abstract_no TEXT,
            title TEXT,
            presenter TEXT,
            presenter_email TEXT,
            institution TEXT,
            theme TEXT,
            session TEXT,
            moderator TEXT,
            date TEXT,
            time_start TEXT,
            time_end TEXT,
            venue TEXT,
            abstract TEXT,
            keywords TEXT,
            presentation_type TEXT
        )""")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS industry (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            speaker TEXT,
            speaker_email TEXT,
            role TEXT,
            organisation TEXT,
            session TEXT,
            date TEXT,
            time_start TEXT,
            time_end TEXT,
            venue TEXT,
            description TEXT
        )""")
        cur.execute("""
        CREATE TABLE IF NOT EXISTS programme (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_type TEXT,
            title TEXT,
            speaker TEXT,
            date TEXT,
            time_start TEXT,
            time_end TEXT,
            venue TEXT,
            description TEXT
        )""")
        con.commit()


def set_kv(k, v):
    with conn() as con:
        con.execute("INSERT OR REPLACE INTO kv(k,v) VALUES (?,?)", (k, v))
        con.commit()


def get_kv(k, default=None):
    with conn() as con:
        row = con.execute("SELECT v FROM kv WHERE k=?", (k,)).fetchone()
    return row[0] if row else default


def save_asset(key, uploaded):
    if uploaded is None:
        return
    ext = Path(uploaded.name).suffix.lower()
    file_path = UPLOAD_DIR / f"{key}{ext}"
    file_path.write_bytes(uploaded.getvalue())
    set_kv(f"asset_{key}", str(file_path))


def asset_path(key):
    p = get_kv(f"asset_{key}")
    if p and Path(p).exists():
        return Path(p)
    return None


def clear_table(name):
    with conn() as con:
        con.execute(f"DELETE FROM {name}")
        con.commit()


def table_to_df(name):
    with conn() as con:
        try:
            return pd.read_sql_query(f"SELECT * FROM {name}", con)
        except Exception:
            return pd.DataFrame()


def normalize_col(c):
    return str(c).strip().lower().replace(" ", "_").replace("-", "_").replace("/", "_")


def pick(row, names, default=""):
    for n in names:
        if n in row and pd.notna(row[n]):
            return str(row[n]).strip()
    return default


def bool_yes(x):
    return str(x).strip().lower() in ["yes", "y", "true", "1", "attend", "hadir", "ya"]


def make_reg_id(prefix, n):
    return f"NICHE2026-{prefix}-{n:04d}"


def import_master_excel(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    sheets = {s.lower().strip(): s for s in xl.sheet_names}

    def find_sheet(candidates):
        for key in candidates:
            if key in sheets:
                return sheets[key]
        return None

    participant_sheet = find_sheet(["participants", "participant", "presenter_participants", "registration", "keynote"])
    academic_sheet = find_sheet(["academic", "abstracts", "abstract", "schedule", "academic_schedule"])
    industry_sheet = find_sheet(["industry", "industry_keynote", "keynotes", "keynote"])
    programme_sheet = find_sheet(["programme", "program", "agenda"])

    if participant_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=participant_sheet)
        df.columns = [normalize_col(c) for c in df.columns]
        clear_table("participants")
        with conn() as con:
            for i, r in df.iterrows():
                email = pick(r, ["email", "e_mail", "emel"])
                name = pick(r, ["full_name", "name", "nama", "presenter", "participant_name"])
                if not email and not name:
                    continue
                category = pick(r, ["category", "kategori", "type", "role"], "Participant")
                prefix = "PAR"
                if "presenter" in category.lower():
                    prefix = "ACAD"
                elif "industry" in category.lower():
                    prefix = "IND"
                elif "keynote" in category.lower():
                    prefix = "KEY"
                elif "media" in category.lower():
                    prefix = "MED"
                registration_id = pick(r, ["registration_id", "reg_id", "id"], make_reg_id(prefix, i + 1))
                attend = pick(r, ["attend_dinner", "dinner", "gala_dinner"], "No")
                con.execute("""
                    INSERT OR REPLACE INTO participants
                    (registration_id, full_name, email, category, institution, phone, faculty,
                     participant_type, attend_dinner, dinner_table, conference_table,
                     checked_in, door_gift, dinner_checked_in, is_walkin, created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    registration_id, name, email.lower(), category,
                    pick(r, ["institution", "organisation", "organization", "company", "affiliation"]),
                    pick(r, ["phone", "mobile", "contact", "telephone"]),
                    pick(r, ["faculty", "sector"]),
                    pick(r, ["participant_type", "type"], ""),
                    "Yes" if bool_yes(attend) else "No",
                    pick(r, ["dinner_table", "table_dinner"]),
                    pick(r, ["conference_table", "table", "seat", "seat_no"]),
                    0, 0, 0, 0, datetime.now().isoformat(timespec="seconds")
                ))
            con.commit()

    if academic_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=academic_sheet)
        df.columns = [normalize_col(c) for c in df.columns]
        clear_table("academic")
        with conn() as con:
            for _, r in df.iterrows():
                title = pick(r, ["title", "paper_title", "presentation_title"])
                presenter = pick(r, ["presenter", "presenter_name", "name"])
                paper_id = pick(r, ["paper_id", "id", "abstract_id"])
                if not title and not presenter and not paper_id:
                    continue
                con.execute("""
                    INSERT INTO academic
                    (paper_id, abstract_no, title, presenter, presenter_email, institution, theme,
                     session, moderator, date, time_start, time_end, venue, abstract, keywords, presentation_type)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    paper_id,
                    pick(r, ["abstract_no", "no", "number"]),
                    title,
                    presenter,
                    pick(r, ["presenter_email", "email", "e_mail"]).lower(),
                    pick(r, ["institution", "affiliation", "organisation"]),
                    pick(r, ["theme", "track"]),
                    pick(r, ["session", "parallel_session"]),
                    pick(r, ["moderator"]),
                    pick(r, ["date"]),
                    pick(r, ["time_start", "start_time", "time"]),
                    pick(r, ["time_end", "end_time"]),
                    pick(r, ["venue", "room", "location"]),
                    pick(r, ["abstract", "abstrak"]),
                    pick(r, ["keywords", "keyword", "kata_kunci"]),
                    pick(r, ["presentation_type", "type"], "Oral")
                ))
            con.commit()

    if industry_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=industry_sheet)
        df.columns = [normalize_col(c) for c in df.columns]
        clear_table("industry")
        with conn() as con:
            for _, r in df.iterrows():
                title = pick(r, ["title", "topic", "event", "session_title"])
                speaker = pick(r, ["speaker", "name", "panelist", "moderator"])
                if not title and not speaker:
                    continue
                con.execute("""
                    INSERT INTO industry
                    (title, speaker, speaker_email, role, organisation, session, date, time_start, time_end, venue, description)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    title,
                    speaker,
                    pick(r, ["speaker_email", "email"]).lower(),
                    pick(r, ["role", "type"], "Speaker"),
                    pick(r, ["organisation", "organization", "institution", "company"]),
                    pick(r, ["session"]),
                    pick(r, ["date"]),
                    pick(r, ["time_start", "start_time", "time"]),
                    pick(r, ["time_end", "end_time"]),
                    pick(r, ["venue", "room", "location"], "Seri Negeri Ballroom III"),
                    pick(r, ["description", "details", "remarks"])
                ))
            con.commit()

    if programme_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=programme_sheet)
        df.columns = [normalize_col(c) for c in df.columns]
        clear_table("programme")
        with conn() as con:
            for _, r in df.iterrows():
                title = pick(r, ["title", "event", "activity"])
                if not title:
                    continue
                con.execute("""
                    INSERT INTO programme
                    (event_type, title, speaker, date, time_start, time_end, venue, description)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (
                    pick(r, ["event_type", "type"], "Conference"),
                    title,
                    pick(r, ["speaker", "name"]),
                    pick(r, ["date"]),
                    pick(r, ["time_start", "start_time", "time"]),
                    pick(r, ["time_end", "end_time"]),
                    pick(r, ["venue", "room", "location"]),
                    pick(r, ["description", "details", "remarks"])
                ))
            con.commit()

    return {"participants": participant_sheet, "academic": academic_sheet, "industry": industry_sheet, "programme": programme_sheet}


def seed_demo_if_empty():
    with conn() as con:
        n = con.execute("SELECT COUNT(*) FROM academic").fetchone()[0]
        if n > 0:
            return
        academic_rows = [
            ("018-015", "1", "ENHANCING HALAL GOVERNANCE LITERACY FOR A SUSTAINABLE HALAL ECOSYSTEM THROUGH MOOC-BASED EDUCATION", "Dr Hajah Makiah Tussaripah Jamil", "nurul2350@uitm.edu.my", "UiTM", "Halal Governance & Policy", "Parallel Session I", "Dr Hajah Makiah Tussaripah Jamil", "10 June 2026", "09:00 AM", "10:30 AM", "Dewan Ampangan 1", "The sustainability of Malaysia’s Halal ecosystem is linked to the competency of its human capital in navigating regulatory and governance frameworks. This study evaluates a MOOC-based multidisciplinary educational intervention in enhancing halal management and governance literacy among students.", "halal management; governance; sustainability", "Oral"),
            ("026-022", "2", "THE SHARIAH-HALAL NEXUS: A UNIFIED FRAMEWORK FOR INSTITUTIONAL AND OPERATIONAL INTEGRITY", "Dr Wan Nor Aisyah Wan Yussof", "wannoraisyah@uitm.edu.my", "UiTM", "Halal Governance & Policy", "Parallel Session I", "Dr Hajah Makiah Tussaripah Jamil", "10 June 2026", "09:00 AM", "10:30 AM", "Dewan Ampangan 1", "This study explores governance evolution within the global Islamic economy and addresses fragmentation between Shariah and Halal governance frameworks through a unified institutional integrity perspective.", "Shariah Governance; Halal Governance; Maqasid", "Oral"),
            ("013-009", "11", "CABARAN PELAKSANAAN PERHOTELAN PATUH SYARIAH DALAM KONTEKS KONTEMPORARI: KAJIAN KUALITATIF BERASASKAN TEMU BUAL INDUSTRI", "Siti Nur Husna Abd Rahman", "snhusna@uitm.edu.my", "UiTM", "Contemporary & Emerging Issues in Halal Ecosystem", "Parallel Session III", "Siti Nur Husna Abd Rahman", "10 June 2026", "09:00 AM", "10:30 AM", "Dewan Ampangan 2", "Kajian ini mengenal pasti cabaran utama pelaksanaan perhotelan patuh syariah berdasarkan perspektif penggiat industri hospitaliti di Malaysia melalui pendekatan kualitatif.", "Perhotelan patuh Syariah; hospitaliti Islam", "Oral"),
        ]
        con.executemany("""
            INSERT INTO academic
            (paper_id, abstract_no, title, presenter, presenter_email, institution, theme, session, moderator, date, time_start, time_end, venue, abstract, keywords, presentation_type)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, academic_rows)

        industry_rows = [
            ("Halal Integrity in a Global Market: Challenges & Solutions", "Mr Muhyidin bin Aziz @ Saari", "", "Keynote Speaker", "JAKIM", "Keynote 1", "9 June 2026", "09:30 AM", "10:30 AM", "Seri Negeri Ballroom III", ""),
            ("Understanding Halal OEM: Opportunities and Challenges", "Puan Ashley Amina Mohd Raihan", "", "Speaker", "HDC", "Industry Session 1", "9 June 2026", "01:30 PM", "02:30 PM", "Seri Negeri Ballroom III", ""),
            ("Digital Halal – Blockchain, Traceability & Anti-Fraud", "Ms Noorhaina Mohd Noor", "", "Speaker", "Interstream Sdn Bhd", "Industry Session 2", "9 June 2026", "02:30 PM", "03:30 PM", "Seri Negeri Ballroom III", ""),
        ]
        con.executemany("""
            INSERT INTO industry
            (title, speaker, speaker_email, role, organisation, session, date, time_start, time_end, venue, description)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """, industry_rows)

        programme_rows = [
            ("Gala Dinner", "Registration of Guests, Participants & Media", "", "9 June 2026", "06:00 PM", "07:15 PM", "Grand Ballroom", ""),
            ("Gala Dinner", "Arrival of VIPs and VVIPs", "", "9 June 2026", "07:40 PM", "08:20 PM", "Grand Ballroom", ""),
            ("Gala Dinner", "Grand Welcome Dinner and Networking Session", "", "9 June 2026", "08:50 PM", "10:00 PM", "Grand Ballroom", ""),
        ]
        con.executemany("""
            INSERT INTO programme
            (event_type, title, speaker, date, time_start, time_end, venue, description)
            VALUES (?,?,?,?,?,?,?,?)
        """, programme_rows)

        participants = [
            ("NICHE2026-ACAD-0001", "Dr Hajah Makiah Tussaripah Jamil", "nurul2350@uitm.edu.my", "Academic Presenter", "UiTM", "", "", "Presenter", "Yes", "", "", 0, 0, 0, 0, datetime.now().isoformat(timespec="seconds")),
            ("NICHE2026-ACAD-0002", "Dr Wan Nor Aisyah Wan Yussof", "wannoraisyah@uitm.edu.my", "Academic Presenter", "UiTM", "", "", "Presenter", "Yes", "", "", 0, 0, 0, 0, datetime.now().isoformat(timespec="seconds")),
            ("NICHE2026-ACAD-0003", "Dr Nabilah Huda Zaim", "nabilahhuda@uitm.edu.my", "Academic Presenter", "UiTM", "", "", "Presenter", "Yes", "", "", 0, 0, 0, 0, datetime.now().isoformat(timespec="seconds")),
            ("NICHE2026-ACAD-0004", "Ahmad Faiz bin Haji Ahmad Ubaidah", "faizubaidah@uitm.edu.my", "Academic Presenter", "UiTM", "", "", "Presenter", "Yes", "", "", 0, 0, 0, 0, datetime.now().isoformat(timespec="seconds")),
        ]
        con.executemany("""
            INSERT OR REPLACE INTO participants
            (registration_id, full_name, email, category, institution, phone, faculty, participant_type,
             attend_dinner, dinner_table, conference_table, checked_in, door_gift, dinner_checked_in, is_walkin, created_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, participants)
        con.commit()


def show_asset_image(key, caption=None):
    p = asset_path(key)
    if p:
        st.markdown('<div class="poster-img">', unsafe_allow_html=True)
        st.image(str(p), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    elif caption:
        st.markdown(f'<div class="dark-card"><h3>{caption}</h3><p>Image will appear after admin upload.</p></div>', unsafe_allow_html=True)


def download_asset_button(key, label):
    p = asset_path(key)
    if p and p.exists():
        st.download_button(label, data=p.read_bytes(), file_name=p.name, mime="application/octet-stream")
    else:
        st.info(f"{label} will be available after admin upload.")


def top_nav():
    pages = [
        ("Home", "home"), ("Programme", "programme"), ("Academic", "academic"),
        ("Industry", "industry"), ("Dinner", "dinner"), ("Registration", "registration"),
        ("Venue", "venue"), ("Contact", "contact"), ("Admin", "admin"),
    ]
    cols = st.columns(3)
    for i, (lab, key) in enumerate(pages):
        with cols[i % 3]:
            if st.button(lab, key=f"nav_{key}"):
                st.session_state.page = key
                st.rerun()


def get_participant_by_email(email):
    if not email:
        return None
    with conn() as con:
        cur = con.execute("SELECT * FROM participants WHERE LOWER(email)=LOWER(?)", (email.strip(),))
        row = cur.fetchone()
        cols = [d[0] for d in cur.description]
    return dict(zip(cols, row)) if row else None


def participant_status_card(p):
    complete = bool(p.get("checked_in")) and bool(p.get("door_gift"))
    st.markdown(f"""
    <div class="card">
      <h3>Registration Details</h3>
      <p><b>Name:</b> {p.get('full_name','')}</p>
      <p><b>Category:</b> {p.get('category','')}</p>
      <p><b>Institution:</b> {p.get('institution','')}</p>
      <p><b>Registration ID:</b> {p.get('registration_id','')}</p>
      <div>
        <span class="meta-pill status-ok">Pre-Registered</span>
        <span class="meta-pill {'status-ok' if p.get('checked_in') else 'status-warn'}">{'Checked-In' if p.get('checked_in') else 'Awaiting Check-In'}</span>
        <span class="meta-pill {'status-ok' if p.get('door_gift') else 'status-warn'}">{'Door Gift Collected' if p.get('door_gift') else 'Door Gift Pending'}</span>
      </div>
      <h3 style="margin-top:14px;">{'Registration Completed' if complete else 'Please proceed to the registration counter'}</h3>
    </div>
    """, unsafe_allow_html=True)

    if p.get("attend_dinner", "No") == "Yes":
        st.markdown(f"""
        <div class="gold-card">
          <h3>Gala Dinner</h3>
          <p><b>Status:</b> Registered</p>
          <p><b>Assigned Table:</b> {p.get('dinner_table') or 'To be assigned by organiser'}</p>
        </div>
        """, unsafe_allow_html=True)


def get_logged_participant():
    email = st.session_state.get("participant_email")
    return get_participant_by_email(email) if email else None


def page_home():
    p = get_logged_participant()
    show_asset_image("main_poster")
    if not asset_path("main_poster"):
        st.markdown("""
        <div class="dark-card">
          <div class="niche-title">NICHE 2026</div>
          <div class="niche-sub">Negeri Sembilan International Conference on Halal & Sustainability Ecosystems</div>
          <div class="gold-line"></div>
          <p>9–10 June 2026 · Royale Chulan Seremban</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="nav-grid">', unsafe_allow_html=True)
    st.markdown('<div class="nav-card"><b>Conference Programme</b><br><small>Academic, industry and dinner agenda</small></div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-card"><b>Academic Sessions</b><br><small>View papers and abstracts</small></div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-card"><b>Gala Dinner</b><br><small>Dinner details and table status</small></div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-card"><b>My Registration</b><br><small>Check registration and door gift status</small></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if p:
        participant_status_card(p)
    else:
        st.markdown('<div class="section-title">Participant Login</div>', unsafe_allow_html=True)
        email = st.text_input("Enter your registered email", placeholder="example@email.com")
        if st.button("View My Registration"):
            found = get_participant_by_email(email)
            if found:
                st.session_state.participant_email = email.strip().lower()
                st.session_state.page = "registration"
                st.rerun()
            else:
                st.warning("Registration record not found. Please proceed to the registration counter.")


def page_registration():
    st.markdown('<div class="section-title">My Registration</div>', unsafe_allow_html=True)
    p = get_logged_participant()
    if not p:
        email = st.text_input("Enter your registered email", placeholder="example@email.com")
        if st.button("View My Registration"):
            found = get_participant_by_email(email)
            if found:
                st.session_state.participant_email = email.strip().lower()
                st.rerun()
            else:
                st.warning("Registration record not found. Please proceed to the registration counter.")
        return

    participant_status_card(p)

    st.markdown('<div class="card"><h3>Gala Dinner RSVP</h3><p>You may update your dinner attendance before table assignment is generated.</p></div>', unsafe_allow_html=True)
    attend = st.radio("Will you attend NICHE 2026 Gala Dinner?", ["Yes", "No"], index=0 if p.get("attend_dinner") == "Yes" else 1)
    if st.button("Update Dinner RSVP"):
        with conn() as con:
            con.execute("UPDATE participants SET attend_dinner=?, dinner_table=NULL WHERE email=?", (attend, p["email"]))
            con.commit()
        st.success("Dinner RSVP updated.")
        st.rerun()


def page_academic():
    st.markdown('<div class="section-title">Academic Conference</div>', unsafe_allow_html=True)
    st.markdown('<div class="dark-card"><h3>Wednesday · 10 June 2026</h3><p>Abstracts are available directly under each paper.</p></div>', unsafe_allow_html=True)
    download_asset_button("abstract_book", "Download Abstract Book")
    df = table_to_df("academic").fillna("")
    if df.empty:
        st.info("Academic schedule will appear after admin uploads the master Excel.")
        return
    query = st.text_input("Search paper, presenter, theme or keyword")
    if query:
        q = query.lower()
        df = df[df.apply(lambda r: q in " ".join([str(x).lower() for x in r.values]), axis=1)]

    group_cols = ["venue", "session", "theme", "time_start", "time_end"]
    for keys, g in df.groupby(group_cols, dropna=False, sort=False):
        venue, session, theme, ts, te = keys
        st.markdown(f"""
        <div class="session-header">
          <h3>{session or 'Academic Session'}</h3>
          <p><b>{theme}</b></p>
          <p>{ts} – {te} · {venue}</p>
        </div>
        """, unsafe_allow_html=True)
        for _, r in g.iterrows():
            st.markdown(f"""
            <div class="paper-card">
              <div><span class="meta-pill">{r.get('paper_id','')}</span><span class="meta-pill">{r.get('presentation_type','Oral')}</span></div>
              <div class="paper-title">{r.get('title','')}</div>
              <p><b>Presenter:</b> {r.get('presenter','')}</p>
              <p><b>Institution:</b> {r.get('institution','')}</p>
              <p><b>Moderator:</b> {r.get('moderator','')}</p>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("View Abstract"):
                st.write(r.get("abstract", "Abstract will be updated by organiser."))
                if r.get("keywords"):
                    st.caption(f"Keywords: {r.get('keywords')}")


def page_industry():
    st.markdown('<div class="section-title">Industry Conference</div>', unsafe_allow_html=True)
    download_asset_button("programme_book", "Download Programme Book")
    df = table_to_df("industry").fillna("")
    if df.empty:
        st.info("Industry programme will appear after admin uploads the master Excel.")
        return
    query = st.text_input("Search speaker, session or topic")
    if query:
        q = query.lower()
        df = df[df.apply(lambda r: q in " ".join([str(x).lower() for x in r.values]), axis=1)]
    for _, r in df.iterrows():
        st.markdown(f"""
        <div class="card">
          <span class="meta-pill">{r.get('role','Speaker')}</span>
          <span class="meta-pill">{r.get('time_start','')} – {r.get('time_end','')}</span>
          <h3>{r.get('title','')}</h3>
          <p><b>{r.get('speaker','')}</b></p>
          <p>{r.get('organisation','')}</p>
          <p>📍 {r.get('venue','Seri Negeri Ballroom III')}</p>
        </div>
        """, unsafe_allow_html=True)


def page_programme():
    st.markdown('<div class="section-title">Conference Programme</div>', unsafe_allow_html=True)
    show_asset_image("conference_poster")
    download_asset_button("programme_book", "Download Programme Book")
    st.markdown('<div class="section-title">Programme Highlights</div>', unsafe_allow_html=True)
    df = table_to_df("programme").fillna("")
    if df.empty:
        st.markdown('<div class="card"><h3>9 June 2026</h3><p>Industry Conference, Launch Ceremony and Gala Dinner</p><h3>10 June 2026</h3><p>Academic Conference and parallel sessions</p></div>', unsafe_allow_html=True)
        return
    for _, r in df.iterrows():
        st.markdown(f"""
        <div class="card">
          <span class="meta-pill">{r.get('event_type','Programme')}</span>
          <span class="meta-pill">{r.get('time_start','')} – {r.get('time_end','')}</span>
          <h3>{r.get('title','')}</h3>
          <p>{r.get('date','')} · {r.get('venue','')}</p>
          <p>{r.get('description','')}</p>
        </div>
        """, unsafe_allow_html=True)


def page_dinner():
    st.markdown('<div class="section-title">NICHE 2026 Gala Dinner</div>', unsafe_allow_html=True)
    show_asset_image("gala_dinner_poster")
    st.markdown("""
    <div class="gold-card">
      <h3>9 June 2026 · 7.00 PM – 10.00 PM</h3>
      <p><b>Venue:</b> Royale Chulan Hotel, Seremban</p>
      <p><b>Dress Code:</b> Batik</p>
      <p><b>Participation Fee:</b> RM200</p>
    </div>
    <div class="card">
      <h3>Programme</h3>
      <p>Registration of guests, participants and media</p>
      <p>Arrival of VIPs and VVIPs</p>
      <p>Cultural performance by UiTM students</p>
      <p>Welcoming speech and official launching ceremony</p>
      <p>Grand welcome dinner and networking session</p>
      <p>Photo session and press conference</p>
    </div>
    """, unsafe_allow_html=True)
    p = get_logged_participant()
    if p:
        st.markdown(f"""
        <div class="card">
          <h3>My Dinner Status</h3>
          <p><b>Attendance:</b> {p.get('attend_dinner','No')}</p>
          <p><b>Assigned Table:</b> {p.get('dinner_table') or 'To be assigned by organiser'}</p>
        </div>
        """, unsafe_allow_html=True)


def render_zone(title, color, tables, highlight=None):
    st.markdown(f'<div class="zone-box"><div class="zone-title" style="background:{color};">{title}</div><br>', unsafe_allow_html=True)
    html = ""
    for t in tables:
        cls = "table-badge table-highlight" if str(t) == str(highlight) else "table-badge"
        html += f'<span class="{cls}">{t}</span>'
    st.markdown(html + "</div>", unsafe_allow_html=True)


def page_venue():
    st.markdown('<div class="section-title">Venue & Seating</div>', unsafe_allow_html=True)
    p = get_logged_participant()
    highlight = (p.get("conference_table") or p.get("dinner_table")) if p else None
    st.markdown('<div class="dark-card"><h3>Seri Negeri Ballroom III</h3><p>Complete zoned seating layout · 28 tables + 2 VIP tables</p></div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><h3 style="text-align:center;">MAIN STAGE</h3><p style="text-align:center;">Extension Stage · Screen · Console</p></div>', unsafe_allow_html=True)
    render_zone("ZONE A · VIP / FRONT", "#E91E63", ["4", "3", "VIP2", "VIP1", "1", "2"], highlight)
    render_zone("ZONE B · MIDDLE LEFT", "#F5B400", ["10", "9", "8", "16", "15", "14"], highlight)
    render_zone("ZONE C · MIDDLE RIGHT", "#2196F3", ["7", "6", "5", "13", "12", "11"], highlight)
    render_zone("ZONE D · BACK LEFT", "#E53935", ["22", "21", "20", "28", "27", "26"], highlight)
    render_zone("ZONE E · BACK RIGHT", "#43A047", ["19", "18", "17", "25", "24", "23"], highlight)


def page_contact():
    st.markdown('<div class="section-title">Contact Secretariat</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
      <h3>Registration Team</h3>
      <p><b>Dr. Nabilah Huda</b><br>+6013-2712856</p>
      <p><b>Gala Dinner Enquiries</b><br>Dr Huda: +6013-2712856<br>Hajar Sopia: +6014-7846225</p>
    </div>
    """, unsafe_allow_html=True)


def admin_login():
    if st.session_state.get("admin"):
        return True
    st.markdown('<div class="section-title">Admin Login</div>', unsafe_allow_html=True)
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == ADMIN_USER and p == ADMIN_PASS:
            st.session_state.admin = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    return False


def find_participants(q):
    df = table_to_df("participants")
    if df.empty:
        return df
    q = str(q).strip().lower()
    if not q:
        return df.head(20)
    return df[df.apply(lambda r: q in " ".join([str(x).lower() for x in r.values]), axis=1)]


def update_participant_flags(pid, checked_in=None, door_gift=None, dinner_checked_in=None):
    fields, vals = [], []
    if checked_in is not None:
        fields.append("checked_in=?")
        vals.append(int(checked_in))
    if door_gift is not None:
        fields.append("door_gift=?")
        vals.append(int(door_gift))
    if dinner_checked_in is not None:
        fields.append("dinner_checked_in=?")
        vals.append(int(dinner_checked_in))
    vals.append(pid)
    with conn() as con:
        con.execute(f"UPDATE participants SET {', '.join(fields)} WHERE id=?", vals)
        con.commit()


def generate_dinner_tables():
    df = table_to_df("participants")
    if df.empty:
        return
    attendees = df[df["attend_dinner"].fillna("No").str.lower().eq("yes")].copy()
    with conn() as con:
        con.execute("UPDATE participants SET dinner_table=NULL")
        media = attendees[attendees["category"].fillna("").str.lower().str.contains("media")]
        for _, r in media.iterrows():
            con.execute("UPDATE participants SET dinner_table=? WHERE id=?", ("4", int(r["id"])))
        others = attendees[~attendees["category"].fillna("").str.lower().str.contains("media")]
        table, count = 5, 0
        for _, r in others.iterrows():
            con.execute("UPDATE participants SET dinner_table=? WHERE id=?", (str(table), int(r["id"])))
            count += 1
            if count >= 10:
                table += 1
                count = 0
        con.commit()


def admin_upload_center():
    st.markdown('<div class="section-title">Upload Center</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
      <h3>Admin only needs 5 items</h3>
      <p>1. Main Poster JPG/PNG</p>
      <p>2. Conference Poster JPG/PNG</p>
      <p>3. Gala Dinner Poster JPG/PNG</p>
      <p>4. ONE Master Excel: Participants, Academic, Industry, Programme</p>
      <p>5. Programme Book + Abstract Book for download only</p>
    </div>
    """, unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        save_asset("main_poster", st.file_uploader("Upload Main Poster", type=["jpg", "jpeg", "png"], key="up_main"))
        save_asset("conference_poster", st.file_uploader("Upload Conference Poster", type=["jpg", "jpeg", "png"], key="up_conf"))
        save_asset("gala_dinner_poster", st.file_uploader("Upload Gala Dinner Poster", type=["jpg", "jpeg", "png"], key="up_dinner_img"))
    with c2:
        save_asset("programme_book", st.file_uploader("Upload Programme Book PDF", type=["pdf"], key="up_prog_pdf"))
        save_asset("abstract_book", st.file_uploader("Upload Abstract Book PDF/DOCX", type=["pdf", "docx"], key="up_abs_pdf"))
    master = st.file_uploader("Upload ONE Master Excel Workbook", type=["xlsx"], key="master_xlsx")
    if master and st.button("Import Master Excel"):
        result = import_master_excel(master)
        st.success(f"Imported. Sheets detected: {result}")


def admin_counter():
    st.markdown('<div class="section-title">Registration Counter</div>', unsafe_allow_html=True)
    q = st.text_input("Search name / email / registration ID")
    res = find_participants(q)
    if res.empty:
        st.info("No participants found.")
    else:
        for _, r in res.head(20).iterrows():
            st.markdown(f"""
            <div class="card">
              <h3>{r['full_name']}</h3>
              <p>{r['email']} · {r['category']} · {r['institution']}</p>
              <p><b>Registration ID:</b> {r['registration_id']}</p>
              <span class="meta-pill {'status-ok' if r['checked_in'] else 'status-warn'}">{'Checked-In' if r['checked_in'] else 'Not Checked-In'}</span>
              <span class="meta-pill {'status-ok' if r['door_gift'] else 'status-warn'}">{'Door Gift Collected' if r['door_gift'] else 'Door Gift Pending'}</span>
              <span class="meta-pill">Dinner: {r['attend_dinner']} {('· Table '+str(r['dinner_table'])) if str(r.get('dinner_table','')) not in ['', 'nan', 'None'] else ''}</span>
            </div>
            """, unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("Check-In", key=f"ci_{r['id']}"):
                    update_participant_flags(int(r["id"]), checked_in=1)
                    st.rerun()
            with c2:
                if st.button("Door Gift Collected", key=f"dg_{r['id']}"):
                    update_participant_flags(int(r["id"]), door_gift=1)
                    st.rerun()
            with c3:
                if st.button("Dinner Present", key=f"di_{r['id']}"):
                    update_participant_flags(int(r["id"]), dinner_checked_in=1)
                    st.rerun()


def admin_walkin():
    st.markdown('<div class="section-title">Walk-In Registration</div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><h3>Presenter walk-in is not allowed</h3><p>Presenter records must come from the Master Excel.</p></div>', unsafe_allow_html=True)
    name = st.text_input("Full Name")
    email = st.text_input("Email")
    phone = st.text_input("Phone")
    category = st.selectbox("Category", ["Academic Participant", "Industry Participant", "Media", "VIP", "Invited Guest", "Committee"])
    institution = st.text_input("Institution / Company")
    attend_dinner = st.radio("Attend Gala Dinner?", ["Yes", "No"], horizontal=True)
    if st.button("Register Walk-In"):
        if not name or not email:
            st.error("Name and email are required.")
            return
        with conn() as con:
            n = con.execute("SELECT COUNT(*) FROM participants WHERE is_walkin=1").fetchone()[0] + 1
            reg = make_reg_id("WI", n)
            dinner_table = "4" if category.lower() == "media" and attend_dinner == "Yes" else None
            con.execute("""
                INSERT OR REPLACE INTO participants
                (registration_id, full_name, email, category, institution, phone, participant_type,
                 attend_dinner, dinner_table, checked_in, door_gift, dinner_checked_in, is_walkin, created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (reg, name, email.lower(), category, institution, phone, "Walk-In", attend_dinner, dinner_table, 1, 0, 0, 1, datetime.now().isoformat(timespec="seconds")))
            con.commit()
        st.success(f"Walk-in registered: {reg}")


def admin_reports():
    st.markdown('<div class="section-title">Reports</div>', unsafe_allow_html=True)
    df = table_to_df("participants")
    if df.empty:
        st.info("No data.")
        return
    metrics = [
        ("Registered", len(df)),
        ("Checked-In", int(df["checked_in"].sum())),
        ("Door Gift", int(df["door_gift"].sum())),
        ("Dinner", int((df["attend_dinner"].fillna("No") == "Yes").sum())),
        ("Walk-In", int(df["is_walkin"].sum())),
    ]
    cols = st.columns(5)
    for c, (lab, val) in zip(cols, metrics):
        with c:
            st.markdown(f'<div class="gold-card"><h3>{val}</h3><p>{lab}</p></div>', unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True)
    st.download_button("Download Participants Report CSV", df.to_csv(index=False).encode(), "niche2026_participants_report.csv", "text/csv")


def admin_reset():
    st.markdown('<div class="section-title">System Maintenance</div>', unsafe_allow_html=True)
    st.warning("Use reset buttons carefully. These are useful after testing.")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Reset Check-In"):
            with conn() as con:
                con.execute("UPDATE participants SET checked_in=0")
                con.commit()
            st.success("Check-in reset.")
        if st.button("Reset Door Gift"):
            with conn() as con:
                con.execute("UPDATE participants SET door_gift=0")
                con.commit()
            st.success("Door gift reset.")
        if st.button("Reset Dinner Attendance"):
            with conn() as con:
                con.execute("UPDATE participants SET dinner_checked_in=0")
                con.commit()
            st.success("Dinner attendance reset.")
    with c2:
        if st.button("Reset Dinner Table Assignment"):
            with conn() as con:
                con.execute("UPDATE participants SET dinner_table=NULL")
                con.commit()
            st.success("Dinner table assignment reset.")
        if st.button("Delete Walk-In Records"):
            with conn() as con:
                con.execute("DELETE FROM participants WHERE is_walkin=1")
                con.commit()
            st.success("Walk-in records deleted.")
        if st.button("Generate Dinner Tables"):
            generate_dinner_tables()
            st.success("Dinner tables generated. Media assigned to Table 4. Participants start from Table 5, 10 pax per table.")


def page_admin():
    if not admin_login():
        return
    st.markdown('<div class="section-title">Admin Portal</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Upload Center", "Registration Counter", "Walk-In", "Reports", "Reset"])
    with tabs[0]:
        admin_upload_center()
    with tabs[1]:
        admin_counter()
    with tabs[2]:
        admin_walkin()
    with tabs[3]:
        admin_reports()
    with tabs[4]:
        admin_reset()


init_db()
seed_demo_if_empty()

if "page" not in st.session_state:
    st.session_state.page = "home"

top_nav()

page = st.session_state.page
if page == "home":
    page_home()
elif page == "programme":
    page_programme()
elif page == "academic":
    page_academic()
elif page == "industry":
    page_industry()
elif page == "dinner":
    page_dinner()
elif page == "registration":
    page_registration()
elif page == "venue":
    page_venue()
elif page == "contact":
    page_contact()
elif page == "admin":
    page_admin()

st.markdown('<div class="bottom-space"></div>', unsafe_allow_html=True)
