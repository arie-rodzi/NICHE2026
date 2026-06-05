import base64
import io
import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import qrcode
import streamlit as st

APP_TITLE = "NICHE 2026"
DB = Path("niche2026.db")
UPLOAD_DIR = Path("user_uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

ASSET_KEYS = {
    "main_poster": ("Main Poster", ["jpg", "jpeg", "png"]),
    "conference_poster": ("Conference Poster", ["jpg", "jpeg", "png"]),
    "programme_pdf": ("Programme Book PDF", ["pdf"]),
    "academic_pdf": ("Academic Tentative PDF", ["pdf"]),
    "industry_pdf": ("Industry Tentative PDF", ["pdf"]),
    "dinner_pdf": ("Gala Dinner Tentative PDF", ["pdf"]),
    "seating_layout": ("Venue / Seating Layout", ["jpg", "jpeg", "png"]),
}

st.set_page_config(page_title="NICHE 2026", page_icon="✨", layout="wide", initial_sidebar_state="collapsed")

CSS = """
<style>
:root{
  --navy:#061B46; --navy2:#09245D; --gold:#D9B454; --gold2:#F6DE8B; --cream:#FFF8E7;
  --ink:#172033; --muted:#657185; --green:#16A34A; --red:#DC2626; --blue:#2563EB;
}
html, body, [class*="css"]{font-family: Inter, Segoe UI, Arial, sans-serif;}
.stApp{background:linear-gradient(180deg,#F7F8FB 0%,#EEF2F8 100%); color:var(--ink);} 
.block-container{padding-top:1.2rem; padding-bottom:3rem; max-width:1180px;}
[data-testid="stSidebar"]{background:#061B46;}
.hero{border-radius:28px; overflow:hidden; background:linear-gradient(135deg,#061B46 0%,#0A2E73 62%,#181341 100%); box-shadow:0 18px 50px rgba(4,18,51,.18); margin-bottom:22px;}
.hero-inner{display:grid; grid-template-columns:1.05fr .95fr; gap:22px; padding:30px; align-items:center;}
.hero-title{font-size:46px; line-height:1.02; font-weight:900; color:white; letter-spacing:-1px; margin:0 0 12px;}
.hero-sub{font-size:18px; line-height:1.45; color:#F9E6A7; margin-bottom:20px; font-weight:600;}
.hero-meta{display:flex; gap:10px; flex-wrap:wrap; margin-top:12px;}
.pill{display:inline-flex; align-items:center; gap:7px; border-radius:999px; padding:9px 14px; font-weight:800; background:rgba(255,255,255,.12); color:white; border:1px solid rgba(255,255,255,.22);}
.poster-frame{background:white; border-radius:22px; padding:10px; box-shadow:0 15px 38px rgba(0,0,0,.28);} 
.poster-frame img{width:100%; border-radius:16px; display:block;}
.section-title{font-size:30px; font-weight:900; color:#061B46; letter-spacing:-.4px; margin:18px 0 8px;}
.section-sub{font-size:16px; color:var(--muted); margin-top:-4px; margin-bottom:16px;}
.card-grid{display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin:16px 0 24px;}
.nav-card{border-radius:22px; padding:22px; min-height:130px; background:#fff; border:1px solid #E5E9F2; box-shadow:0 10px 28px rgba(12,30,66,.08);}
.nav-card h3{margin:0 0 8px; color:#061B46; font-size:20px; font-weight:900;}
.nav-card p{margin:0; color:#637083; font-size:14px; line-height:1.35;}
.gold-card{background:linear-gradient(135deg,#D9B454,#FFF1A8); color:#061B46;}
.blue-card{background:linear-gradient(135deg,#0B3B91,#061B46); color:white;}
.blue-card h3,.blue-card p{color:white;}
.green-card{background:linear-gradient(135deg,#ECFDF5,#FFFFFF); border-left:8px solid #16A34A;}
.status{display:inline-block; padding:7px 12px; border-radius:999px; font-weight:900; font-size:13px;}
.status-green{background:#DCFCE7; color:#166534;} .status-yellow{background:#FEF3C7;color:#92400E;} .status-blue{background:#DBEAFE;color:#1E40AF;} .status-red{background:#FEE2E2;color:#991B1B;}
.big-card{border-radius:26px; background:white; border:1px solid #E3E8F1; box-shadow:0 12px 30px rgba(12,30,66,.08); padding:24px; margin:14px 0;}
.session-head{border-radius:24px; padding:20px 22px; margin:18px 0 12px; color:white; background:linear-gradient(135deg,#061B46,#0E3C87); border-left:9px solid #D9B454; box-shadow:0 12px 28px rgba(6,27,70,.2);}
.session-head h2{margin:0 0 6px; font-size:25px; font-weight:950; color:white;}
.session-head .line{color:#F6DE8B; font-weight:800;}
.paper-card{border-radius:24px; background:#FFFFFF; padding:20px; margin:13px 0; border:1px solid #E5E9F2; box-shadow:0 8px 22px rgba(6,27,70,.07); border-left:7px solid #D9B454;}
.paper-id{font-weight:900; color:#0B3B91; font-size:13px; letter-spacing:.3px; text-transform:uppercase; margin-bottom:6px;}
.paper-title{font-size:20px; line-height:1.25; font-weight:950; color:#061B46; margin-bottom:12px;}
.meta-row{display:flex; gap:9px; flex-wrap:wrap; margin:8px 0;}
.meta-chip{border-radius:999px; background:#F1F5F9; color:#334155; padding:7px 11px; font-size:13px; font-weight:800;}
.meta-chip.gold{background:#FFF7D6;color:#7C5200;} .meta-chip.green{background:#DCFCE7;color:#166534;} .meta-chip.blue{background:#DBEAFE;color:#1E40AF;}
.abstract-box{background:#FAFBFE; border:1px solid #E4E9F2; border-radius:18px; padding:18px; color:#263244; line-height:1.55; font-size:15px;}
.speaker-card{border-radius:24px; background:#fff; padding:20px; margin:13px 0; border:1px solid #E5E9F2; box-shadow:0 8px 22px rgba(6,27,70,.07); border-left:7px solid #0B3B91;}
.speaker-card h3{font-size:21px; color:#061B46; margin:0 0 8px; font-weight:950;}
.btn-note{border-radius:20px; background:#FFF7D6; padding:16px; border:1px solid #E5C76A; color:#5F4107; font-weight:700;}
.login-box{max-width:520px; margin:auto;}
.admin-box{background:#061B46; color:white; padding:20px; border-radius:22px; margin-bottom:18px;}
.admin-box h2{color:#F6DE8B; margin:0;}
hr{border:none; border-top:1px solid #E3E8F1; margin:22px 0;}
.stButton>button, .stDownloadButton>button{border-radius:999px!important; border:none!important; background:linear-gradient(135deg,#D9B454,#F6DE8B)!important; color:#061B46!important; font-weight:900!important; padding:.62rem 1.1rem!important; box-shadow:0 8px 18px rgba(217,180,84,.28)!important;}
[data-testid="stDataFrame"]{border-radius:18px; overflow:hidden;}
@media(max-width:900px){
 .block-container{padding-left:1rem; padding-right:1rem; padding-top:.5rem;}
 .hero-inner{grid-template-columns:1fr; padding:18px;}
 .hero-title{font-size:33px;}
 .hero-sub{font-size:15px;}
 .card-grid{grid-template-columns:repeat(2,1fr); gap:12px;}
 .nav-card{padding:16px; min-height:112px;}
 .nav-card h3{font-size:17px;}
 .section-title{font-size:24px;}
 .paper-title{font-size:18px;}
 .session-head h2{font-size:21px;}
 .big-card{padding:18px; border-radius:22px;}
}
@media(max-width:520px){.card-grid{grid-template-columns:1fr 1fr;} .nav-card p{font-size:12px}.poster-frame{padding:7px}.hero-meta .pill{font-size:12px;padding:8px 10px}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# ---------- DB ----------
def db():
    con = sqlite3.connect(DB)
    con.execute("CREATE TABLE IF NOT EXISTS dataframes (name TEXT PRIMARY KEY, payload TEXT, updated_at TEXT)")
    con.execute("CREATE TABLE IF NOT EXISTS assets (key TEXT PRIMARY KEY, label TEXT, filename TEXT, path TEXT, mime TEXT, updated_at TEXT)")
    con.commit()
    return con

def save_df(name, df):
    con = db()
    con.execute("REPLACE INTO dataframes VALUES (?,?,?)", (name, df.to_json(orient="records", force_ascii=False), datetime.now().isoformat(timespec="seconds")))
    con.commit(); con.close()

def load_df(name):
    con = db()
    row = con.execute("SELECT payload FROM dataframes WHERE name=?", (name,)).fetchone()
    con.close()
    if not row:
        return pd.DataFrame()
    try:
        return pd.read_json(io.StringIO(row[0]))
    except Exception:
        return pd.DataFrame()

def save_asset(key, label, upload):
    ext = upload.name.split(".")[-1].lower()
    safe = f"{key}.{ext}"
    path = UPLOAD_DIR / safe
    path.write_bytes(upload.getvalue())
    con = db()
    con.execute("REPLACE INTO assets VALUES (?,?,?,?,?,?)", (key, label, upload.name, str(path), upload.type or "", datetime.now().isoformat(timespec="seconds")))
    con.commit(); con.close()

def get_asset(key):
    con = db()
    row = con.execute("SELECT label, filename, path, mime, updated_at FROM assets WHERE key=?", (key,)).fetchone()
    con.close()
    if not row:
        return None
    return {"label":row[0], "filename":row[1], "path":Path(row[2]), "mime":row[3], "updated_at":row[4]}

def image_b64(path):
    if path and Path(path).exists():
        data = Path(path).read_bytes()
        ext = Path(path).suffix.lower().replace('.', '') or 'jpeg'
        return f"data:image/{ext};base64," + base64.b64encode(data).decode()
    return None

def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df

def extract_sheet(file, sheet_name, required_any):
    try:
        raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    except Exception:
        return pd.DataFrame()
    header_idx = None
    for i in range(min(12, len(raw))):
        row_values = [str(x).strip() for x in raw.iloc[i].fillna("").tolist()]
        if any(req in row_values for req in required_any):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0
    headers = [str(x).strip() if str(x).strip() and str(x) != 'nan' else f"Column {j+1}" for j, x in enumerate(raw.iloc[header_idx].tolist())]
    df = raw.iloc[header_idx+1:].copy()
    df.columns = headers
    df = df.dropna(how="all")
    return normalize_columns(df)

def import_master_excel(uploaded_file):
    # Save a backup copy for admin reference
    backup_path = UPLOAD_DIR / "latest_master_excel.xlsx"
    backup_path.write_bytes(uploaded_file.getvalue())
    file_obj = io.BytesIO(uploaded_file.getvalue())
    xl = pd.ExcelFile(file_obj)
    sheets = xl.sheet_names

    def find_sheet(candidates):
        for cand in candidates:
            for s in sheets:
                if cand.lower() in s.lower():
                    return s
        return None

    p_sheet = find_sheet(["PARTICIPANTS", "PARTICIPANT", "MASTER"])
    a_sheet = find_sheet(["ABSTRACT", "SCHEDULE"])
    k_sheet = find_sheet(["INDUSTRY", "KEYNOTE"])
    pr_sheet = find_sheet(["PROGRAMME", "PROGRAM"])

    file_obj = io.BytesIO(uploaded_file.getvalue())
    participants = extract_sheet(file_obj, p_sheet, ["Registration ID", "Full Name", "Email"]) if p_sheet else pd.DataFrame()
    file_obj = io.BytesIO(uploaded_file.getvalue())
    abstracts = extract_sheet(file_obj, a_sheet, ["Paper ID", "Presenter Name", "Paper Title"]) if a_sheet else pd.DataFrame()
    file_obj = io.BytesIO(uploaded_file.getvalue())
    industry = extract_sheet(file_obj, k_sheet, ["Date", "Time", "Title", "Name"]) if k_sheet else pd.DataFrame()
    file_obj = io.BytesIO(uploaded_file.getvalue())
    programme = extract_sheet(file_obj, pr_sheet, ["Day", "Time", "Title"]) if pr_sheet else pd.DataFrame()

    # Clean key columns
    for df in [participants, abstracts, industry, programme]:
        if not df.empty:
            for c in df.columns:
                if df[c].dtype == object:
                    df[c] = df[c].astype(str).replace({"nan":"", "NaT":""}).str.strip()

    if not participants.empty:
        # ensure useful columns
        for col in ["Registration ID", "Full Name", "Email", "Category", "Institution", "Check-In Status", "Door Gift Status", "Dinner RSVP", "Table No", "Dinner Table"]:
            if col not in participants.columns:
                participants[col] = ""
        participants["Email"] = participants["Email"].str.lower()
        save_df("participants", participants)
    if not abstracts.empty:
        for col in ["Paper ID", "Abstract No", "Presenter Email", "Presenter Name", "Institution", "Paper Title", "Abstract", "Keywords", "Track", "Session", "Date", "Time", "Room", "Moderator", "Presentation Type"]:
            if col not in abstracts.columns:
                abstracts[col] = ""
        if "Presenter Email" in abstracts.columns:
            abstracts["Presenter Email"] = abstracts["Presenter Email"].astype(str).str.lower()
        save_df("abstracts", abstracts)
    if not industry.empty:
        for col in ["Date", "Time", "Title", "Name", "Designation/Details"]:
            if col not in industry.columns:
                industry[col] = ""
        save_df("industry", industry)
    if not programme.empty:
        for col in ["Day", "Time", "Title", "Speaker", "Venue", "Category", "Details"]:
            if col not in programme.columns:
                programme[col] = ""
        save_df("programme", programme)
    return participants.shape, abstracts.shape, industry.shape, programme.shape

# ---------- UTIL ----------
def safe(x, default=""):
    if pd.isna(x): return default
    return str(x)

def download_asset_button(key, label=None):
    asset = get_asset(key)
    if asset and asset["path"].exists():
        label = label or f"Download {asset['label']}"
        st.download_button(label, asset["path"].read_bytes(), file_name=asset["filename"], mime=asset["mime"] or "application/octet-stream", key=f"dl_{key}")
    else:
        st.info(f"{label or ASSET_KEYS.get(key,('File',))[0]} belum dimuat naik oleh admin.")

def qr_png(text):
    img = qrcode.make(text or "NICHE2026")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def get_participant_by_email(email):
    df = load_df("participants")
    if df.empty or "Email" not in df.columns:
        return None
    hit = df[df["Email"].astype(str).str.lower().str.strip() == email.lower().strip()]
    if hit.empty:
        return None
    return hit.iloc[0].to_dict()

def participant_papers(email, name=""):
    df = load_df("abstracts")
    if df.empty:
        return df
    email = str(email).lower().strip()
    mask = pd.Series([False]*len(df))
    if "Presenter Email" in df.columns and email:
        mask = mask | (df["Presenter Email"].astype(str).str.lower().str.strip() == email)
    if "Presenter Name" in df.columns and name:
        mask = mask | (df["Presenter Name"].astype(str).str.lower().str.contains(name.lower().strip(), na=False))
    return df[mask]

def card_grid():
    st.markdown("""
    <div class='card-grid'>
      <div class='nav-card gold-card'><h3>Academic Conference</h3><p>Parallel sessions, presenter details and abstracts.</p></div>
      <div class='nav-card blue-card'><h3>Industry Conference</h3><p>Keynote, panels and industry sharing sessions.</p></div>
      <div class='nav-card green-card'><h3>Gala Dinner</h3><p>Dinner programme and attendance information.</p></div>
      <div class='nav-card'><h3>My Registration</h3><p>Registration status, QR code, seating and door gift.</p></div>
    </div>
    """, unsafe_allow_html=True)

# ---------- PAGES ----------
def page_home():
    main = get_asset("main_poster")
    main_src = image_b64(main["path"]) if main else None
    if main_src:
        poster_html = f"<div class='poster-frame'><img src='{main_src}'></div>"
    else:
        poster_html = "<div class='poster-frame' style='min-height:360px;display:flex;align-items:center;justify-content:center;color:#061B46;font-weight:900'>Main poster will appear after admin upload.</div>"
    st.markdown(f"""
    <div class='hero'><div class='hero-inner'>
      <div>
        <div class='hero-title'>NICHE 2026</div>
        <div class='hero-sub'>Negeri Sembilan International Conference on Halal & Sustainability Ecosystems</div>
        <div class='hero-meta'>
          <span class='pill'>📅 9–10 June 2026</span>
          <span class='pill'>📍 Royale Chulan Seremban</span>
          <span class='pill'>Halal & Sustainability</span>
        </div>
      </div>
      {poster_html}
    </div></div>
    """, unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Quick Access</div>", unsafe_allow_html=True)
    card_grid()
    c1, c2 = st.columns(2)
    with c1:
        download_asset_button("programme_pdf", "Download Programme Book")
    with c2:
        download_asset_button("academic_pdf", "Download Academic Tentative / Abstract Reference")

def page_registration():
    st.markdown("<div class='section-title'>My Registration</div><div class='section-sub'>Enter your registered email to view your conference details.</div>", unsafe_allow_html=True)
    email = st.text_input("Registered email address", placeholder="name@example.com")
    if st.button("View My Registration"):
        st.session_state["participant_email"] = email
    email = st.session_state.get("participant_email", email)
    if not email:
        return
    p = get_participant_by_email(email)
    if not p:
        st.error("Registration record not found. Please proceed to the registration counter for assistance.")
        return
    rid = safe(p.get("Registration ID")) or f"NICHE2026-{safe(p.get('Email'))}"
    st.markdown(f"<div class='big-card'><h2 style='margin-top:0;color:#061B46'>Welcome, {safe(p.get('Full Name'))}</h2><p>Your registration details are shown below.</p></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([2,1])
    with c1:
        check = safe(p.get("Check-In Status")) or "Awaiting Check-In"
        gift = safe(p.get("Door Gift Status")) or "Pending Collection"
        dinner = safe(p.get("Dinner RSVP")) or "Not Confirmed"
        st.markdown(f"""
        <div class='big-card'>
          <div class='meta-row'><span class='status status-blue'>{safe(p.get('Category'),'Participant')}</span><span class='status status-yellow'>{rid}</span></div>
          <h3 style='color:#061B46'>Participant Information</h3>
          <p><b>Institution:</b> {safe(p.get('Institution'))}</p>
          <p><b>Conference Table:</b> {safe(p.get('Table No'),'To be assigned')}</p>
          <p><b>Dinner Table:</b> {safe(p.get('Dinner Table'),'To be assigned')}</p>
          <hr>
          <p><span class='status {'status-green' if check.lower() in ['checked in','yes','completed'] else 'status-yellow'}'>{check}</span></p>
          <p><span class='status {'status-green' if gift.lower() in ['collected','yes','given'] else 'status-yellow'}'>Door Gift: {gift}</span></p>
          <p><span class='status status-blue'>Gala Dinner: {dinner}</span></p>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.image(qr_png(rid), caption="NICHE QR Code", use_container_width=True)
    papers = participant_papers(safe(p.get("Email")), safe(p.get("Full Name")))
    if not papers.empty:
        st.markdown("<div class='section-title'>Your Presentation</div>", unsafe_allow_html=True)
        render_academic_cards(papers, compact=False)

def page_programme():
    st.markdown("<div class='section-title'>Conference Programme</div><div class='section-sub'>Programme is shown in mobile-friendly cards. PDF is available for download at the bottom.</div>", unsafe_allow_html=True)
    df = load_df("programme")
    if df.empty:
        st.info("Programme will appear after admin uploads the Master Excel workbook.")
    else:
        for day, g in df.groupby(df.get("Day", pd.Series([""]*len(df))).fillna(""), sort=False):
            st.markdown(f"<div class='session-head'><h2>{safe(day, 'Programme')}</h2><div class='line'>NICHE 2026 Programme</div></div>", unsafe_allow_html=True)
            for _, r in g.iterrows():
                st.markdown(f"""
                <div class='speaker-card'>
                  <div class='meta-row'><span class='meta-chip gold'>🕒 {safe(r.get('Time'))}</span><span class='meta-chip blue'>{safe(r.get('Category'))}</span><span class='meta-chip green'>📍 {safe(r.get('Venue'))}</span></div>
                  <h3>{safe(r.get('Title'))}</h3>
                  <p><b>{safe(r.get('Speaker'))}</b></p>
                  <p style='color:#657185'>{safe(r.get('Details'))}</p>
                </div>
                """, unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)
    download_asset_button("programme_pdf", "Download Programme Book PDF")

def render_academic_cards(df, compact=False):
    if df.empty:
        return
    group_cols = [c for c in ["Room", "Session", "Track", "Time"] if c in df.columns]
    if group_cols:
        grouped = df.groupby(group_cols, sort=False, dropna=False)
    else:
        grouped = [("Academic Presentations", df)]
    for key, g in grouped:
        if isinstance(key, tuple):
            parts = [safe(k) for k in key if safe(k)]
        else:
            parts = [safe(key)]
        heading = " · ".join(parts) or "Academic Session"
        st.markdown(f"<div class='session-head'><h2>{heading}</h2><div class='line'>Academic Conference</div></div>", unsafe_allow_html=True)
        for _, r in g.iterrows():
            pid = safe(r.get("Paper ID")) or safe(r.get("Abstract No"))
            title = safe(r.get("Paper Title"))
            presenter = safe(r.get("Presenter Name"))
            inst = safe(r.get("Institution"))
            time = safe(r.get("Time"))
            room = safe(r.get("Room"))
            track = safe(r.get("Track"))
            st.markdown(f"""
            <div class='paper-card'>
              <div class='paper-id'>{pid}</div>
              <div class='paper-title'>{title}</div>
              <div class='meta-row'>
                <span class='meta-chip blue'>Presenter: {presenter}</span>
                <span class='meta-chip green'>📍 {room}</span>
                <span class='meta-chip gold'>🕒 {time}</span>
              </div>
              <p style='color:#657185;margin-bottom:4px'><b>Institution:</b> {inst}</p>
              <p style='color:#657185;margin-top:0'><b>Track:</b> {track}</p>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("View Abstract"):
                st.markdown(f"<div class='abstract-box'>{safe(r.get('Abstract'), 'Abstract will be updated by the organizer.')}</div>", unsafe_allow_html=True)

def page_academic():
    st.markdown("<div class='section-title'>Academic Conference</div><div class='section-sub'>Presentation schedule with abstracts. Abstracts open directly here; no separate abstract-book page is required.</div>", unsafe_allow_html=True)
    df = load_df("abstracts")
    if df.empty:
        st.info("Academic schedule and abstracts will appear after admin uploads the Master Excel workbook.")
    else:
        q = st.text_input("Search paper, presenter, keyword or room", placeholder="Example: 018-015 / governance / Dewan Ampangan 1")
        show = df.copy()
        if q:
            ql = q.lower()
            show = show[show.apply(lambda row: row.astype(str).str.lower().str.contains(ql, na=False).any(), axis=1)]
        render_academic_cards(show)
    st.markdown("<hr>", unsafe_allow_html=True)
    download_asset_button("academic_pdf", "Download Academic Tentative / Abstract Reference PDF")

def page_industry():
    st.markdown("<div class='section-title'>Industry Conference</div><div class='section-sub'>Keynote, moderator and industry sharing sessions are displayed as cards for mobile viewing.</div>", unsafe_allow_html=True)
    df = load_df("industry")
    if df.empty:
        # fallback from programme
        pr = load_df("programme")
        if not pr.empty and "Category" in pr.columns:
            df = pr[pr["Category"].astype(str).str.contains("industry|keynote|panel", case=False, na=False)].copy()
            df["Name"] = df.get("Speaker", "")
            df["Designation/Details"] = df.get("Details", "")
    if df.empty:
        st.info("Industry programme will appear after admin uploads the Master Excel workbook.")
    else:
        for _, r in df.iterrows():
            title = safe(r.get("Title")) or safe(r.get("Event"))
            name = safe(r.get("Name")) or safe(r.get("Speaker"))
            st.markdown(f"""
            <div class='speaker-card'>
              <div class='meta-row'><span class='meta-chip gold'>🕒 {safe(r.get('Time'))}</span><span class='meta-chip green'>📅 {safe(r.get('Date'))}</span></div>
              <h3>{title}</h3>
              <p><b>{name}</b></p>
              <p style='color:#657185'>{safe(r.get('Designation/Details'))}</p>
            </div>
            """, unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)
    download_asset_button("industry_pdf", "Download Industry Tentative PDF")

def page_dinner():
    st.markdown("<div class='section-title'>Gala Dinner</div><div class='section-sub'>Dinner programme and attendance reference.</div>", unsafe_allow_html=True)
    pr = load_df("programme")
    show = pd.DataFrame()
    if not pr.empty and "Category" in pr.columns:
        show = pr[pr["Category"].astype(str).str.contains("dinner|gala", case=False, na=False)].copy()
    if show.empty:
        st.markdown("<div class='big-card'><h3 style='color:#061B46'>NICHE 2026 Gala Dinner</h3><p>The dinner programme will appear after admin uploads the Master Excel workbook or Gala Dinner PDF.</p></div>", unsafe_allow_html=True)
    else:
        for _, r in show.iterrows():
            st.markdown(f"""
            <div class='speaker-card'>
              <div class='meta-row'><span class='meta-chip gold'>🕒 {safe(r.get('Time'))}</span><span class='meta-chip green'>📍 {safe(r.get('Venue'))}</span></div>
              <h3>{safe(r.get('Title'))}</h3>
              <p style='color:#657185'>{safe(r.get('Details'))}</p>
            </div>
            """, unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)
    download_asset_button("dinner_pdf", "Download Gala Dinner Tentative PDF")

def page_seating():
    st.markdown("<div class='section-title'>Venue & Seating</div><div class='section-sub'>Seating layout is displayed as an image. Use PNG/JPG for best mobile viewing.</div>", unsafe_allow_html=True)
    asset = get_asset("seating_layout")
    if asset and asset["path"].exists():
        st.image(str(asset["path"]), use_container_width=True)
    else:
        st.info("Seating layout will appear after admin upload.")

def page_contact():
    st.markdown("<div class='section-title'>Contact Secretariat</div>", unsafe_allow_html=True)
    st.markdown("""
    <div class='big-card'>
      <h3 style='color:#061B46'>NICHE 2026 Secretariat</h3>
      <p>Please proceed to the registration counter for check-in, door gift collection, walk-in assistance and seating enquiries.</p>
      <p><b>Venue:</b> Royale Chulan Seremban</p>
      <p><b>Conference Date:</b> 9–10 June 2026</p>
    </div>
    """, unsafe_allow_html=True)

# ---------- ADMIN ----------
def admin_login():
    if st.session_state.get("admin_ok"):
        return True
    st.markdown("<div class='admin-box'><h2>Admin Portal</h2><p>Restricted access for NICHE 2026 organizer and registration staff.</p></div>", unsafe_allow_html=True)
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == "admin" and p == "NICHE2026admin":
            st.session_state["admin_ok"] = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    return False

def admin_upload_center():
    st.markdown("<div class='section-title'>Admin Upload Center</div><div class='section-sub'>Upload one Master Excel and all public conference assets here. Participant pages update automatically.</div>", unsafe_allow_html=True)
    master = st.file_uploader("Upload Master Excel Workbook", type=["xlsx", "xls"])
    if master and st.button("Import Master Excel"):
        try:
            shapes = import_master_excel(master)
            st.success(f"Master Excel imported. Participants {shapes[0]}, Abstracts {shapes[1]}, Industry {shapes[2]}, Programme {shapes[3]}.")
        except Exception as e:
            st.error(f"Import failed: {e}")
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### Conference Assets")
    cols = st.columns(2)
    for i, (key, (label, types)) in enumerate(ASSET_KEYS.items()):
        with cols[i % 2]:
            current = get_asset(key)
            if current:
                st.caption(f"Current: {current['filename']} | Updated: {current['updated_at']}")
            file = st.file_uploader(label, type=types, key=f"up_{key}")
            if file and st.button(f"Save {label}", key=f"save_{key}"):
                save_asset(key, label, file)
                st.success(f"{label} uploaded.")

def update_participant_rows(df):
    save_df("participants", df)

def admin_checkin():
    st.markdown("<div class='section-title'>Check-In & Door Gift</div>", unsafe_allow_html=True)
    df = load_df("participants")
    if df.empty:
        st.info("Upload Master Excel first.")
        return
    q = st.text_input("Search by name, email or registration ID")
    show = df.copy()
    if q:
        ql = q.lower()
        show = show[show.apply(lambda row: row.astype(str).str.lower().str.contains(ql, na=False).any(), axis=1)]
    st.write(f"{len(show)} record(s) found")
    for idx, r in show.head(30).iterrows():
        with st.container():
            st.markdown(f"""
            <div class='big-card'>
              <h3 style='color:#061B46;margin:0'>{safe(r.get('Full Name'))}</h3>
              <p>{safe(r.get('Registration ID'))} · {safe(r.get('Email'))} · {safe(r.get('Category'))}</p>
            </div>
            """, unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            if c1.button("Check In", key=f"check_{idx}"):
                df.loc[idx, "Check-In Status"] = "Checked In"
                update_participant_rows(df); st.rerun()
            if c2.button("Door Gift Given", key=f"gift_{idx}"):
                df.loc[idx, "Door Gift Status"] = "Collected"
                update_participant_rows(df); st.rerun()
            if c3.button("Dinner Confirmed", key=f"dinner_{idx}"):
                df.loc[idx, "Dinner RSVP"] = "Confirmed"
                update_participant_rows(df); st.rerun()

def admin_walkin():
    st.markdown("<div class='section-title'>Walk-In Registration</div><div class='section-sub'>Presenter walk-in is not allowed. Use this form for non-presenter walk-in participants only.</div>", unsafe_allow_html=True)
    df = load_df("participants")
    if df.empty:
        df = pd.DataFrame(columns=["Registration ID","Full Name","Email","Category","Institution","Phone","Check-In Status","Door Gift Status","Dinner RSVP","Table No","Dinner Table","Admin Notes"])
    with st.form("walkin_form"):
        c1, c2 = st.columns(2)
        name = c1.text_input("Full Name *")
        email = c2.text_input("Email *")
        phone = c1.text_input("Phone")
        inst = c2.text_input("Institution / Company")
        cat = st.selectbox("Category", ["Academic Participant", "Industry Delegate", "Invited Guest", "Student", "Sponsor", "Media", "Committee"])
        dinner = st.selectbox("Gala Dinner RSVP", ["Not Confirmed", "Confirmed", "Unable to Attend"])
        table = st.text_input("Conference Table / Seat (optional)")
        submitted = st.form_submit_button("Register Walk-In")
    if submitted:
        if not name or not email:
            st.error("Full Name and Email are required.")
            return
        n = len(df) + 1
        rid = f"NICHE2026-WALKIN-{n:04d}"
        row = {"Registration ID":rid,"Full Name":name,"Email":email.lower(),"Category":cat,"Institution":inst,"Phone":phone,"Check-In Status":"Checked In","Door Gift Status":"Pending Collection","Dinner RSVP":dinner,"Table No":table,"Dinner Table":"","Admin Notes":"Walk-in registered by admin"}
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        save_df("participants", df)
        st.success(f"Walk-in participant registered: {rid}")
        st.image(qr_png(rid), width=180)

def admin_reports():
    st.markdown("<div class='section-title'>Reports & Export</div>", unsafe_allow_html=True)
    df = load_df("participants")
    if df.empty:
        st.info("Upload Master Excel first.")
        return
    c1,c2,c3,c4 = st.columns(4)
    total = len(df)
    checked = df.get("Check-In Status", pd.Series(dtype=str)).astype(str).str.contains("checked|yes", case=False, na=False).sum()
    gift = df.get("Door Gift Status", pd.Series(dtype=str)).astype(str).str.contains("collected|given|yes", case=False, na=False).sum()
    dinner = df.get("Dinner RSVP", pd.Series(dtype=str)).astype(str).str.contains("confirmed|yes", case=False, na=False).sum()
    for col, title, num in [(c1,"Total",total),(c2,"Checked In",checked),(c3,"Door Gift",gift),(c4,"Dinner",dinner)]:
        col.markdown(f"<div class='nav-card'><h3>{num}</h3><p>{title}</p></div>", unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Download Participants Report CSV", df.to_csv(index=False).encode("utf-8-sig"), "niche2026_participants_report.csv", "text/csv")

def page_admin():
    if not admin_login(): return
    st.success("Admin access granted.")
    tabs = st.tabs(["Upload Center", "Check-In", "Walk-In", "Reports"])
    with tabs[0]: admin_upload_center()
    with tabs[1]: admin_checkin()
    with tabs[2]: admin_walkin()
    with tabs[3]: admin_reports()

# ---------- NAV ----------
def main():
    db()
    st.sidebar.title("NICHE 2026")
    page = st.sidebar.radio("Menu", ["Home", "My Registration", "Conference Programme", "Academic Conference", "Industry Conference", "Gala Dinner", "Venue & Seating", "Contact Secretariat", "Admin Portal"])
    if page == "Home": page_home()
    elif page == "My Registration": page_registration()
    elif page == "Conference Programme": page_programme()
    elif page == "Academic Conference": page_academic()
    elif page == "Industry Conference": page_industry()
    elif page == "Gala Dinner": page_dinner()
    elif page == "Venue & Seating": page_seating()
    elif page == "Contact Secretariat": page_contact()
    elif page == "Admin Portal": page_admin()

if __name__ == "__main__":
    main()
