from __future__ import annotations

import base64
import hashlib
import io
import json
import os
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image

APP_TITLE = "NICHE 2026"
APP_SUBTITLE = "Negeri Sembilan International Conference on Halal & Sustainability Ecosystems"
DB_PATH = Path("data/niche2026.db")
UPLOAD_DIR = Path("data/uploads")
SAMPLE_MASTER = Path("data/samples/NICHE2026_MASTER_REAL_DATA.xlsx")
ADMIN_USERNAME = os.getenv("NICHE_ADMIN_USERNAME", "admin")
ADMIN_PASSWORD_HASH = hashlib.sha256(os.getenv("NICHE_ADMIN_PASSWORD", "NICHE2026admin").encode()).hexdigest()

ASSET_KEYS = {
    "main_poster": {"label": "Main Poster", "types": ["jpg", "jpeg", "png"]},
    "conference_poster": {"label": "Conference Poster", "types": ["jpg", "jpeg", "png"]},
    "programme_book": {"label": "Programme Book PDF", "types": ["pdf"]},
    "academic_tentative": {"label": "Academic Tentative PDF", "types": ["pdf"]},
    "industry_tentative": {"label": "Industry Tentative PDF", "types": ["pdf"]},
    "gala_dinner_tentative": {"label": "Gala Dinner Tentative PDF", "types": ["pdf"]},
    "abstract_book_pdf": {"label": "Abstract Book PDF", "types": ["pdf"]},
    "seating_layout": {"label": "Venue / Seating Layout", "types": ["jpg", "jpeg", "png"]},
    "venue_map": {"label": "Venue Map", "types": ["jpg", "jpeg", "png", "pdf"]},
    "logo_1": {"label": "Logo 1", "types": ["jpg", "jpeg", "png"]},
    "logo_2": {"label": "Logo 2", "types": ["jpg", "jpeg", "png"]},
}

SHEET_ALIASES = {
    "participants": ["participants", "participant", "presenter", "presenters", "registration", "keynote"],
    "abstracts": ["abstracts", "abstract", "schedule", "presentation", "paper"],
    "programme": ["programme", "program", "events", "tentative"],
    "dinner": ["dinner", "gala"],
    "seating": ["seating", "seat", "table"],
    "settings": ["settings", "setting", "info"],
}

REQUIRED_DATASETS = ["participants", "abstracts", "programme", "dinner", "seating", "settings"]

st.set_page_config(page_title="NICHE 2026 Registration", page_icon="🎟️", layout="wide")


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root{--navy:#071A3A;--blue:#0A2A5E;--gold:#D6AE4B;--soft:#F6F1E7;--ink:#102033;}
        .stApp{background:linear-gradient(135deg,#06122A 0%,#0A2A5E 42%,#08142E 100%);}
        section[data-testid="stSidebar"]{background:linear-gradient(180deg,#06142E,#0B2C64);border-right:1px solid rgba(214,174,75,.25);}
        .block-container{padding-top:1.4rem; padding-bottom:3rem; max-width:1180px;}
        h1,h2,h3{color:#FDF8E8!important; letter-spacing:-.02em;}
        p, label, .stMarkdown, .stTextInput label{color:#F5F7FB!important;}
        div[data-testid="stMetric"]{background:rgba(255,255,255,.08);border:1px solid rgba(214,174,75,.28);border-radius:20px;padding:18px;box-shadow:0 10px 30px rgba(0,0,0,.18);}
        div[data-testid="stMetric"] label, div[data-testid="stMetric"] div{color:#fff!important;}
        .hero{padding:28px 30px;border-radius:28px;background:linear-gradient(135deg,rgba(255,255,255,.10),rgba(255,255,255,.035));border:1px solid rgba(214,174,75,.35);box-shadow:0 18px 50px rgba(0,0,0,.25);margin-bottom:18px;}
        .hero .eyebrow{color:#D6AE4B;font-weight:800;letter-spacing:.14em;text-transform:uppercase;font-size:.82rem;}
        .hero h1{font-size:2.7rem;margin:.15rem 0 .2rem 0;color:#fff!important;}
        .hero .sub{font-size:1.05rem;color:#F7EED2!important;max-width:850px;}
        .card{background:rgba(255,255,255,.92);border-radius:22px;padding:22px;border:1px solid rgba(214,174,75,.40);box-shadow:0 14px 42px rgba(0,0,0,.20);color:#102033!important;margin:12px 0;}
        .card *{color:#102033!important;}
        .gold-card{background:linear-gradient(135deg,#F9E9B7,#D6AE4B);border-radius:22px;padding:22px;color:#071A3A!important;border:1px solid rgba(255,255,255,.5);}
        .gold-card *{color:#071A3A!important;}
        .pill{display:inline-block;padding:7px 12px;border-radius:999px;background:#F3E3B4;color:#071A3A!important;font-weight:700;margin:3px 4px 3px 0;}
        .status-ok{background:#DCFCE7;color:#166534!important;}
        .status-warn{background:#FEF3C7;color:#92400E!important;}
        .small-muted{font-size:.86rem;color:#CBD5E1!important;}
        .whitebox{background:rgba(255,255,255,.94);border-radius:18px;padding:16px;margin:8px 0;border:1px solid rgba(214,174,75,.35);}
        .whitebox *{color:#102033!important;}
        .stButton>button, .stDownloadButton>button{border-radius:14px!important;background:linear-gradient(135deg,#D6AE4B,#F6DA8A)!important;color:#071A3A!important;border:0!important;font-weight:800!important;box-shadow:0 8px 20px rgba(0,0,0,.2)!important;}
        .stTextInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea{border-radius:14px!important;}
        .dataframe{background:white!important;}
        @media(max-width:768px){.hero h1{font-size:2rem}.hero{padding:20px}.block-container{padding-left:1rem;padding-right:1rem}.card{padding:16px}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as con:
        con.execute("CREATE TABLE IF NOT EXISTS datasets (name TEXT PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL)")
        con.execute("CREATE TABLE IF NOT EXISTS assets (asset_key TEXT PRIMARY KEY, file_name TEXT, file_path TEXT, mime TEXT, updated_at TEXT)")
        con.execute("CREATE TABLE IF NOT EXISTS attendance (email TEXT PRIMARY KEY, checked_in INTEGER DEFAULT 0, door_gift INTEGER DEFAULT 0, dinner_attend TEXT DEFAULT '', dinner_table TEXT DEFAULT '', conference_seat TEXT DEFAULT '', updated_at TEXT)")
        con.execute("CREATE TABLE IF NOT EXISTS walkins (id TEXT PRIMARY KEY, data TEXT NOT NULL, created_at TEXT NOT NULL)")
        con.commit()


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def save_dataset(name: str, df: pd.DataFrame) -> None:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    data = df.fillna("").astype(str).to_json(orient="records", force_ascii=False)
    with sqlite3.connect(DB_PATH) as con:
        con.execute("REPLACE INTO datasets(name,data,updated_at) VALUES(?,?,?)", (name, data, now()))
        con.commit()


def load_dataset(name: str) -> pd.DataFrame:
    with sqlite3.connect(DB_PATH) as con:
        row = con.execute("SELECT data FROM datasets WHERE name=?", (name,)).fetchone()
    if not row:
        return pd.DataFrame()
    try:
        return pd.read_json(io.StringIO(row[0]), dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()


def get_assets() -> dict:
    with sqlite3.connect(DB_PATH) as con:
        rows = con.execute("SELECT asset_key,file_name,file_path,mime,updated_at FROM assets").fetchall()
    return {r[0]: {"file_name": r[1], "file_path": r[2], "mime": r[3], "updated_at": r[4]} for r in rows}


def save_asset(asset_key: str, uploaded_file) -> None:
    if uploaded_file is None:
        return
    ext = uploaded_file.name.split(".")[-1].lower()
    safe_name = f"{asset_key}_{uuid.uuid4().hex[:8]}.{ext}"
    path = UPLOAD_DIR / safe_name
    path.write_bytes(uploaded_file.getbuffer())
    mime = uploaded_file.type or ("application/pdf" if ext == "pdf" else "image")
    with sqlite3.connect(DB_PATH) as con:
        con.execute("REPLACE INTO assets(asset_key,file_name,file_path,mime,updated_at) VALUES(?,?,?,?,?)", (asset_key, uploaded_file.name, str(path), mime, now()))
        con.commit()


def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df.empty:
        return None
    lookup = {str(c).strip().lower().replace("_", " "): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().replace("_", " ")
        if key in lookup:
            return lookup[key]
    for c in df.columns:
        lc = str(c).lower()
        if any(cand.lower() in lc for cand in candidates):
            return c
    return None


def email_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["email", "e-mail", "emel", "mail", "presenter email", "participant email"])


def name_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["name", "full name", "nama", "presenter", "presenter name", "participant name"])


def category_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["category", "kategori", "type", "role", "status", "participant type"])


def title_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["title", "paper title", "tajuk", "paper"])


def paper_id_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["paper id", "id", "abstract id", "abstract no", "paper code"])


def abstract_col(df: pd.DataFrame) -> str | None:
    return find_col(df, ["abstract", "abstrak"])



def smart_read_excel_sheet(xl: pd.ExcelFile, sheet_name: str, canonical: str | None = None) -> pd.DataFrame:
    """Read a sheet even when organizer workbooks have title/notes rows before the real header."""
    raw = pd.read_excel(xl, sheet_name, header=None, dtype=str).fillna("")
    if raw.empty:
        return pd.DataFrame()
    keyword_map = {
        "participants": ["email", "name", "category", "institution", "participant"],
        "abstracts": ["paper", "abstract", "presenter", "title", "venue", "session"],
        "programme": ["event", "time", "date", "venue", "programme"],
        "dinner": ["dinner", "table", "email", "attend"],
        "seating": ["seat", "table", "email", "name"],
        "settings": ["item", "value", "setting"],
    }
    keys = keyword_map.get(canonical or "", [])
    best_idx, best_score = 0, -1
    for idx in range(min(12, len(raw))):
        vals = [str(v).strip().lower() for v in raw.iloc[idx].tolist() if str(v).strip()]
        nonempty = len(vals)
        score = nonempty
        score += sum(4 for v in vals for k in keys if k in v)
        if score > best_score:
            best_idx, best_score = idx, score
    headers = [str(v).strip() if str(v).strip() else f"Column_{i+1}" for i, v in enumerate(raw.iloc[best_idx].tolist())]
    df = raw.iloc[best_idx+1:].copy()
    df.columns = headers
    df = df.loc[:, ~pd.Index(df.columns).astype(str).str.startswith("Unnamed")]
    df = df.loc[:, ~pd.Index(df.columns).astype(str).str.startswith("Column_")] if df.shape[1] > 1 else df
    df = df.dropna(how="all").fillna("")
    # Remove rows that are fully empty after stripping
    if not df.empty:
        df = df[df.astype(str).apply(lambda r: any(x.strip() for x in r), axis=1)]
    return df.reset_index(drop=True)

def import_master_excel(file) -> dict:
    xl = pd.ExcelFile(file)
    imported = {}
    used_sheets = set()
    for canonical, aliases in SHEET_ALIASES.items():
        match = None
        for sheet in xl.sheet_names:
            s = sheet.lower().strip()
            if sheet in used_sheets:
                continue
            if any(a in s for a in aliases):
                match = sheet
                break
        if match:
            df = smart_read_excel_sheet(xl, match, canonical)
            save_dataset(canonical, df)
            imported[canonical] = {"sheet": match, "rows": len(df), "cols": len(df.columns)}
            used_sheets.add(match)
    # If only one or two sheets found, make best effort mapping first sheets.
    if "participants" not in imported and xl.sheet_names:
        df = smart_read_excel_sheet(xl, xl.sheet_names[0], "participants")
        save_dataset("participants", df)
        imported["participants"] = {"sheet": xl.sheet_names[0], "rows": len(df), "cols": len(df.columns)}
    return imported


def qr_image(data: str) -> Image.Image:
    img = qrcode.make(data)
    return img.convert("RGB")


def show_hero(title: str, subtitle: str | None = None, eyebrow: str = "Official Conference Portal") -> None:
    st.markdown(
        f"""
        <div class="hero">
          <div class="eyebrow">{eyebrow}</div>
          <h1>{title}</h1>
          <div class="sub">{subtitle or APP_SUBTITLE}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def file_bytes(path: str) -> bytes:
    return Path(path).read_bytes()


def display_asset(asset_key: str, title: str | None = None, compact: bool = False) -> None:
    assets = get_assets()
    asset = assets.get(asset_key)
    label = title or ASSET_KEYS.get(asset_key, {}).get("label", asset_key)
    if not asset or not asset.get("file_path") or not Path(asset["file_path"]).exists():
        st.info(f"{label} has not been uploaded by the administrator.")
        return
    path = Path(asset["file_path"])
    suffix = path.suffix.lower()
    if suffix in [".jpg", ".jpeg", ".png"]:
        if title:
            st.subheader(title)
        st.image(str(path), use_container_width=True)
    elif suffix == ".pdf":
        if title:
            st.subheader(title)
        with open(path, "rb") as f:
            data = f.read()
        st.download_button(f"Download {label}", data, file_name=asset["file_name"] or path.name, mime="application/pdf")
        b64 = base64.b64encode(data).decode("utf-8")
        height = 480 if compact else 760
        st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="{height}" style="border:1px solid rgba(214,174,75,.4);border-radius:16px;"></iframe>', unsafe_allow_html=True)


def get_participant_by_email(email: str) -> tuple[pd.Series | None, pd.DataFrame]:
    df = load_dataset("participants")
    ec = email_col(df)
    if df.empty or not ec or not email:
        return None, df
    matches = df[df[ec].astype(str).str.strip().str.lower() == email.strip().lower()]
    if matches.empty:
        return None, df
    return matches.iloc[0], df


def attendance_for(email: str) -> dict:
    if not email:
        return {}
    with sqlite3.connect(DB_PATH) as con:
        row = con.execute("SELECT checked_in,door_gift,dinner_attend,dinner_table,conference_seat FROM attendance WHERE email=?", (email.lower(),)).fetchone()
    if not row:
        return {"checked_in": 0, "door_gift": 0, "dinner_attend": "", "dinner_table": "", "conference_seat": ""}
    return {"checked_in": row[0], "door_gift": row[1], "dinner_attend": row[2], "dinner_table": row[3], "conference_seat": row[4]}


def update_attendance(email: str, **kwargs) -> None:
    current = attendance_for(email)
    current.update(kwargs)
    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            "REPLACE INTO attendance(email,checked_in,door_gift,dinner_attend,dinner_table,conference_seat,updated_at) VALUES(?,?,?,?,?,?,?)",
            (email.lower(), int(current.get("checked_in") or 0), int(current.get("door_gift") or 0), current.get("dinner_attend", ""), current.get("dinner_table", ""), current.get("conference_seat", ""), now()),
        )
        con.commit()


def render_public_home() -> None:
    show_hero(APP_TITLE, "Welcome to the official registration and conference information portal.")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        display_asset("main_poster", "Main Poster")
    with c2:
        st.markdown("""<div class='gold-card'><h3>Conference Access</h3><p>Participants may use their registered email address to view registration status, QR code, programme, abstract book, seating information and dinner information.</p></div>""", unsafe_allow_html=True)
        st.write("")
        display_asset("conference_poster", "Conference Poster", compact=True)


def render_registration() -> None:
    show_hero("My Registration", "Enter your registered email address to view your conference registration details.")
    email = st.text_input("Registered Email Address", value=st.session_state.get("participant_email", ""), placeholder="example@email.com")
    if st.button("View My Registration"):
        st.session_state["participant_email"] = email.strip().lower()
    email = st.session_state.get("participant_email", "").strip().lower()
    if not email:
        return
    row, df = get_participant_by_email(email)
    if row is None:
        st.warning("Registration record not found. Please proceed to the registration counter for assistance.")
        return
    nc, cc, ic = name_col(df), category_col(df), find_col(df, ["institution", "organisation", "organization", "company", "affiliation"])
    name = row.get(nc, "Participant") if nc else "Participant"
    category = row.get(cc, "Participant") if cc else "Participant"
    inst = row.get(ic, "") if ic else ""
    reg_id_col = find_col(df, ["registration id", "reg id", "id", "qr id"])
    reg_id = row.get(reg_id_col, "") if reg_id_col else ""
    if not reg_id:
        reg_id = "NICHE2026-" + hashlib.md5(email.encode()).hexdigest()[:6].upper()
    att = attendance_for(email)
    st.markdown(f"<div class='card'><h2>Welcome, {name}</h2><span class='pill'>{category}</span><span class='pill'>{inst}</span></div>", unsafe_allow_html=True)
    a, b, c = st.columns(3)
    a.metric("Check-In Status", "Checked In" if att.get("checked_in") else "Pending")
    b.metric("Door Gift", "Collected" if att.get("door_gift") else "Ready for Collection")
    c.metric("Registration ID", reg_id)
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown("<div class='whitebox'><h3>Participant Information</h3></div>", unsafe_allow_html=True)
        info = pd.DataFrame([{"Field": str(k), "Value": str(v)} for k, v in row.items() if str(v).strip()])
        st.dataframe(info, use_container_width=True, hide_index=True)
    with col2:
        st.markdown("<div class='whitebox'><h3>QR Code</h3><p>Please show this QR code at the registration counter.</p></div>", unsafe_allow_html=True)
        st.image(qr_image(reg_id), width=260)
    # Presenter information if available
    abstracts = load_dataset("abstracts")
    if not abstracts.empty:
        emc = email_col(abstracts)
        pc = find_col(abstracts, ["presenter email", "email"])
        if emc or pc:
            ac = emc or pc
            pmatch = abstracts[abstracts[ac].astype(str).str.strip().str.lower() == email]
            if not pmatch.empty:
                st.subheader("Presentation Information")
                render_abstract_cards(pmatch, limit=3)


def render_programme() -> None:
    show_hero("Conference Programme", "View or download the uploaded programme book and event timetable.")
    display_asset("programme_book", "Programme Book")
    df = load_dataset("programme")
    if not df.empty:
        st.subheader("Programme Table")
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Programme table has not been imported from the master Excel workbook.")


def filter_df(df: pd.DataFrame, query: str) -> pd.DataFrame:
    if df.empty or not query:
        return df
    mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False)).any(axis=1)
    return df[mask]


def render_academic() -> None:
    show_hero("Academic Conference", "Academic tentative, parallel sessions and presenter information.")
    display_asset("academic_tentative", "Academic Tentative")
    df = load_dataset("abstracts")
    if not df.empty:
        st.subheader("Academic Presentation List")
        q = st.text_input("Search academic presentation", placeholder="Paper ID / presenter / title / venue")
        st.dataframe(filter_df(df, q), use_container_width=True, hide_index=True)


def render_industry() -> None:
    show_hero("Industry Conference", "Industry keynote, speaker and moderator information.")
    display_asset("industry_tentative", "Industry Tentative")
    df = load_dataset("participants")
    cc = category_col(df)
    if not df.empty and cc:
        industry = df[df[cc].astype(str).str.contains("industry|keynote|moderator|speaker", case=False, na=False)]
        st.subheader("Industry / Keynote List")
        st.dataframe(industry, use_container_width=True, hide_index=True)


def render_abstract_cards(df: pd.DataFrame, limit: int | None = None) -> None:
    if df.empty:
        st.info("No abstract found.")
        return
    pid = paper_id_col(df)
    tc = title_col(df)
    ac = abstract_col(df)
    nc = name_col(df)
    ic = find_col(df, ["institution", "affiliation", "organisation", "organization"])
    vc = find_col(df, ["venue", "room", "dewan"])
    dc = find_col(df, ["date", "tarikh"])
    tm = find_col(df, ["time", "masa"])
    rows = df.head(limit) if limit else df
    for _, r in rows.iterrows():
        paper = r.get(pid, "") if pid else ""
        title = r.get(tc, "Untitled") if tc else "Untitled"
        presenter = r.get(nc, "") if nc else ""
        inst = r.get(ic, "") if ic else ""
        venue = r.get(vc, "") if vc else ""
        date = r.get(dc, "") if dc else ""
        time = r.get(tm, "") if tm else ""
        abst = r.get(ac, "") if ac else ""
        with st.expander(f"{paper} — {title}", expanded=False):
            st.markdown(f"**Presenter:** {presenter}")
            if inst: st.markdown(f"**Institution:** {inst}")
            if date or time or venue: st.markdown(f"**Schedule:** {date} {time} | {venue}")
            st.markdown("**Abstract**")
            st.write(abst if abst else "Abstract text not available.")


def render_abstract_book() -> None:
    show_hero("Abstract Book", "Search by paper ID, presenter name, title or keyword. Optimised for mobile viewing.")
    display_asset("abstract_book_pdf", "Compiled Abstract Book PDF", compact=True)
    df = load_dataset("abstracts")
    if df.empty:
        st.info("Abstract records have not been imported.")
        return
    q = st.text_input("Search Abstract", placeholder="Paper ID / presenter / keyword")
    filtered = filter_df(df, q)
    st.caption(f"Showing {len(filtered)} record(s).")
    render_abstract_cards(filtered)


def render_presentation_info() -> None:
    show_hero("Presentation Information", "This section is shown for presenters based on the email entered in My Registration.")
    email = st.text_input("Presenter Email", value=st.session_state.get("participant_email", ""))
    if st.button("Search Presentation"):
        st.session_state["participant_email"] = email.strip().lower()
    email = email.strip().lower()
    df = load_dataset("abstracts")
    if df.empty:
        st.info("Presentation records have not been imported.")
        return
    emc = email_col(df)
    if emc and email:
        matches = df[df[emc].astype(str).str.strip().str.lower() == email]
    else:
        matches = pd.DataFrame()
    if matches.empty and email:
        matches = filter_df(df, email)
    render_abstract_cards(matches if email else pd.DataFrame())


def render_gala_dinner() -> None:
    show_hero("Gala Dinner", "Dinner tentative and attendance response.")
    display_asset("gala_dinner_tentative", "Gala Dinner Tentative")
    email = st.text_input("Registered Email for Dinner Response", value=st.session_state.get("participant_email", ""))
    if email:
        choice = st.radio("Dinner Attendance", ["I will attend", "Unable to attend"], horizontal=True)
        if st.button("Submit Dinner Response"):
            update_attendance(email, dinner_attend=choice)
            st.success("Dinner response saved.")
        att = attendance_for(email)
        if att.get("dinner_attend"):
            st.info(f"Current dinner response: {att.get('dinner_attend')} | Table: {att.get('dinner_table') or 'Not assigned'}")


def render_venue() -> None:
    show_hero("Venue & Seating", "View uploaded seating layout, venue map and table assignment.")
    c1, c2 = st.columns(2)
    with c1:
        display_asset("seating_layout", "Seating Layout")
    with c2:
        display_asset("venue_map", "Venue Map")
    df = load_dataset("seating")
    if not df.empty:
        q = st.text_input("Search seating", placeholder="Name / email / table")
        st.dataframe(filter_df(df, q), use_container_width=True, hide_index=True)


def render_contact() -> None:
    show_hero("Contact Secretariat", "For registration, programme and conference assistance.")
    settings = load_dataset("settings")
    if not settings.empty:
        st.dataframe(settings, use_container_width=True, hide_index=True)
    else:
        st.markdown("""<div class='card'><h3>NICHE 2026 Secretariat</h3><p>Please proceed to the registration counter for assistance.</p></div>""", unsafe_allow_html=True)


def admin_login() -> bool:
    if st.session_state.get("admin_ok"):
        return True
    show_hero("Admin Portal", "Administrator login is required to upload conference data and manage registration.")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == ADMIN_USERNAME and hashlib.sha256(p.encode()).hexdigest() == ADMIN_PASSWORD_HASH:
            st.session_state["admin_ok"] = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    return False


def admin_upload_center() -> None:
    st.subheader("Master Excel Upload")
    st.caption("Upload one Excel workbook containing multiple sheets such as PARTICIPANTS, ABSTRACTS, PROGRAMME, DINNER, SEATING and SETTINGS.")
    uploaded = st.file_uploader("Upload NICHE Master Excel", type=["xlsx", "xls"], key="master_excel")
    if uploaded and st.button("Import Master Excel"):
        imported = import_master_excel(uploaded)
        st.success("Master Excel imported successfully.")
        st.json(imported)
    if SAMPLE_MASTER.exists():
        with open(SAMPLE_MASTER, "rb") as f:
            st.download_button("Download Sample Master Excel", f.read(), file_name="NICHE2026_MASTER_REAL_DATA_SAMPLE.xlsx")

    st.divider()
    st.subheader("Conference Assets Upload")
    st.caption("Upload or replace all posters, PDF tentatives, programme book, abstract book and seating layout here. Participant pages will update automatically.")
    cols = st.columns(2)
    for i, (key, meta) in enumerate(ASSET_KEYS.items()):
        with cols[i % 2]:
            f = st.file_uploader(meta["label"], type=meta["types"], key=f"asset_{key}")
            if f and st.button(f"Save {meta['label']}", key=f"save_{key}"):
                save_asset(key, f)
                st.success(f"{meta['label']} saved.")
    st.divider()
    st.subheader("Uploaded Assets")
    assets = get_assets()
    if assets:
        asset_rows = []
        for key, asset in assets.items():
            asset_rows.append({"Asset": ASSET_KEYS.get(key, {}).get("label", key), "File": asset.get("file_name"), "Updated": asset.get("updated_at")})
        st.dataframe(pd.DataFrame(asset_rows), use_container_width=True, hide_index=True)
    else:
        st.info("No assets uploaded yet.")


def admin_checkin() -> None:
    st.subheader("Check-In & Door Gift")
    df = load_dataset("participants")
    if df.empty:
        st.info("Please upload the Master Excel first.")
        return
    q = st.text_input("Search by name / email / registration ID")
    result = filter_df(df, q) if q else pd.DataFrame()
    if not result.empty:
        st.dataframe(result.head(20), use_container_width=True, hide_index=True)
        ec = email_col(result)
        if ec:
            selected = st.selectbox("Select participant email", result[ec].dropna().astype(str).unique())
            att = attendance_for(selected)
            c1, c2, c3 = st.columns(3)
            if c1.button("Check In"):
                update_attendance(selected, checked_in=1)
                st.success("Participant checked in.")
            if c2.button("Door Gift Collected"):
                update_attendance(selected, door_gift=1)
                st.success("Door gift marked as collected.")
            seat = c3.text_input("Conference Seat / Table", value=att.get("conference_seat", ""))
            if c3.button("Save Seat"):
                update_attendance(selected, conference_seat=seat)
                st.success("Seat saved.")
            st.write(attendance_for(selected))


def admin_walkin() -> None:
    st.subheader("Walk-In Registration")
    st.caption("Presenter walk-in is disabled. Presenter records must be added through the Master Excel and verified by the secretariat.")
    with st.form("walkin_form"):
        name = st.text_input("Full Name *")
        email = st.text_input("Email *")
        phone = st.text_input("Phone")
        org = st.text_input("Institution / Company")
        category = st.selectbox("Category", ["Academic Participant", "Industry Delegate", "Invited Guest", "VIP", "Student", "Sponsor", "Media"])
        seat = st.text_input("Conference Seat / Table")
        dinner = st.selectbox("Dinner Attendance", ["", "I will attend", "Unable to attend"])
        submitted = st.form_submit_button("Register Walk-In")
    if submitted:
        if not name or not email:
            st.error("Name and email are required.")
            return
        walkin_id = "NICHE-WI-" + uuid.uuid4().hex[:6].upper()
        data = {"Registration ID": walkin_id, "Full Name": name, "Email": email, "Phone": phone, "Institution": org, "Category": category, "Walk-In": "Yes"}
        with sqlite3.connect(DB_PATH) as con:
            con.execute("INSERT INTO walkins(id,data,created_at) VALUES(?,?,?)", (walkin_id, json.dumps(data), now()))
            con.commit()
        # Also append to participants dataset
        df = load_dataset("participants")
        new = pd.DataFrame([data])
        df = pd.concat([df, new], ignore_index=True) if not df.empty else new
        save_dataset("participants", df)
        update_attendance(email, checked_in=1, conference_seat=seat, dinner_attend=dinner)
        st.success(f"Walk-in registered: {walkin_id}")


def admin_reports() -> None:
    st.subheader("Reports & Export")
    participants = load_dataset("participants")
    abstracts = load_dataset("abstracts")
    programme = load_dataset("programme")
    with sqlite3.connect(DB_PATH) as con:
        att = pd.read_sql_query("SELECT * FROM attendance", con)
        walkins = pd.read_sql_query("SELECT * FROM walkins", con)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Participants", len(participants))
    c2.metric("Abstracts", len(abstracts))
    c3.metric("Checked In", int(att["checked_in"].sum()) if not att.empty else 0)
    c4.metric("Door Gifts", int(att["door_gift"].sum()) if not att.empty else 0)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        participants.to_excel(writer, sheet_name="participants", index=False)
        abstracts.to_excel(writer, sheet_name="abstracts", index=False)
        programme.to_excel(writer, sheet_name="programme", index=False)
        att.to_excel(writer, sheet_name="attendance", index=False)
        walkins.to_excel(writer, sheet_name="walkins", index=False)
    st.download_button("Download Full Report Excel", output.getvalue(), file_name="NICHE2026_full_report.xlsx")
    st.subheader("Attendance Records")
    st.dataframe(att, use_container_width=True, hide_index=True)


def render_admin() -> None:
    if not admin_login():
        return
    show_hero("Admin Portal", "Upload conference data/assets and manage registration operations.", eyebrow="Administrator")
    choice = st.radio("Admin Menu", ["Upload Center", "Check-In & Door Gift", "Walk-In Registration", "Reports & Export"], horizontal=True)
    if choice == "Upload Center":
        admin_upload_center()
    elif choice == "Check-In & Door Gift":
        admin_checkin()
    elif choice == "Walk-In Registration":
        admin_walkin()
    else:
        admin_reports()


def render_sidebar() -> str:
    st.sidebar.markdown(f"# {APP_TITLE}")
    st.sidebar.markdown("Registration & Conference Portal")
    menu = [
        "Home",
        "My Registration",
        "Conference Programme",
        "Academic Conference",
        "Industry Conference",
        "Abstract Book",
        "Presentation Information",
        "Gala Dinner",
        "Venue & Seating",
        "Contact Secretariat",
        "Admin Portal",
    ]
    return st.sidebar.radio("Navigation", menu)


def main() -> None:
    inject_css()
    init_db()
    page = render_sidebar()
    if page == "Home":
        render_public_home()
    elif page == "My Registration":
        render_registration()
    elif page == "Conference Programme":
        render_programme()
    elif page == "Academic Conference":
        render_academic()
    elif page == "Industry Conference":
        render_industry()
    elif page == "Abstract Book":
        render_abstract_book()
    elif page == "Presentation Information":
        render_presentation_info()
    elif page == "Gala Dinner":
        render_gala_dinner()
    elif page == "Venue & Seating":
        render_venue()
    elif page == "Contact Secretariat":
        render_contact()
    elif page == "Admin Portal":
        render_admin()


if __name__ == "__main__":
    main()
