import os, sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "niche.db"
EXCEL_PATH = BASE_DIR / "niche_data.xlsx"

POSTER_DIR = BASE_DIR / "posters"
POSTER_DIR.mkdir(exist_ok=True)

MAIN_POSTER = POSTER_DIR / "main_poster.jpg"
DINNER_POSTER = POSTER_DIR / "dinner_poster.jpg"

ADMIN_PASSWORD = "NICHE2026admin"
TABLES = [3, 4, 8, 9, 10, 27]
SEATS = 10

st.set_page_config(
    page_title="NICHE 2026 Self Check-In",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ================= CSS =================
st.markdown("""
<style>
[data-testid="stSidebar"], [data-testid="collapsedControl"], #MainMenu, footer {
    display:none!important;
}
.stApp {
    background: radial-gradient(circle at top left,#1b2364,#050829 50%,#020312);
    color:white;
}
.block-container {max-width:1350px; padding-top:1rem;}
h1,h2,h3 {color:#ffe88a!important;}
.card {
    background:rgba(5,10,46,.75);
    border:1px solid rgba(244,212,105,.25);
    border-radius:22px;
    padding:22px;
    margin:14px 0;
    box-shadow:0 20px 60px rgba(0,0,0,.35);
}
.brand {
    padding:18px 22px;
    border-radius:24px;
    border:1px solid rgba(244,212,105,.3);
    background:linear-gradient(135deg,rgba(5,10,46,.85),rgba(25,35,105,.5));
    margin-bottom:18px;
}
.brand-title {
    font-size:38px;
    font-weight:900;
    letter-spacing:2px;
    background:linear-gradient(135deg,#fff4aa,#d4af37,#8b6914);
    -webkit-background-clip:text;
    color:transparent;
}
.stButton button {
    background:linear-gradient(135deg,#fff0a3,#d4af37)!important;
    color:#07103e!important;
    font-weight:900!important;
    border-radius:14px!important;
    border:none!important;
}
.session {
    border:1px solid rgba(244,212,105,.24);
    background:rgba(5,10,46,.72);
    border-radius:22px;
    padding:18px;
    margin:16px 0;
}
.session-grid {
    display:grid;
    grid-template-columns:minmax(160px,220px) 1fr;
    gap:18px;
    align-items:start;
}
.time {
    color:#ffe88a;
    font-size:25px;
    font-weight:900;
    line-height:1.2;
    word-break:break-word;
}
.label {
    color:#aaa4bd;
    letter-spacing:5px;
    text-transform:uppercase;
    font-size:13px;
    margin-bottom:8px;
    word-break:break-word;
}
.title {
    color:white;
    font-size:23px;
    font-weight:900;
    line-height:1.28;
    word-break:break-word;
}
.paper {
    background:rgba(255,255,255,.055);
    border:1px solid rgba(255,255,255,.10);
    border-radius:18px;
    padding:15px;
    margin-top:12px;
    overflow:hidden;
}
.paper-top {
    display:flex;
    flex-wrap:wrap;
    gap:10px;
    align-items:center;
}
.pid {
    background:linear-gradient(135deg,#fff0a3,#d4af37);
    color:#07103e;
    padding:5px 10px;
    border-radius:999px;
    font-weight:900;
    font-size:12px;
}
.author {
    color:white;
    font-weight:800;
    overflow-wrap:anywhere;
}
.abstract-title {
    color:#fff0a3;
    font-weight:900;
    font-size:18px;
    line-height:1.35;
    margin-top:8px;
    overflow-wrap:anywhere;
}
.abstract {
    color:#cfc9df;
    font-size:14px;
    line-height:1.55;
    margin-top:6px;
    overflow-wrap:anywhere;
}
@media(max-width:750px){
    .session-grid{grid-template-columns:1fr;}
}
</style>
""", unsafe_allow_html=True)

# ================= DATABASE =================
def conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def exec_sql(sql, params=()):
    with conn() as c:
        cur = c.cursor()
        cur.execute(sql, params)
        c.commit()
        return cur

def df_sql(sql, params=()):
    with conn() as c:
        return pd.read_sql_query(sql, c, params=params)

def init_db():
    exec_sql("""
    CREATE TABLE IF NOT EXISTS participants(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT NOT NULL,
        email TEXT NOT NULL UNIQUE,
        organisation TEXT,
        phone TEXT,
        participant_type TEXT,
        dinner_join INTEGER DEFAULT 0,
        table_number INTEGER,
        checked_in INTEGER DEFAULT 0,
        checkin_time TEXT,
        door_gift_collected INTEGER DEFAULT 0,
        door_gift_time TEXT,
        created_at TEXT
    )
    """)
    exec_sql("""
    CREATE TABLE IF NOT EXISTS audit_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT,
        details TEXT,
        created_at TEXT
    )
    """)

def log(action, details=""):
    exec_sql(
        "INSERT INTO audit_log(action,details,created_at) VALUES(?,?,?)",
        (action, details, datetime.now().isoformat(timespec="seconds"))
    )

init_db()

# ================= HELPERS =================
def save_file(uploaded, path):
    if uploaded is None:
        return False
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    return True

def clean(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def read_sheets():
    if not EXCEL_PATH.exists():
        return {}
    try:
        return pd.read_excel(EXCEL_PATH, sheet_name=None)
    except Exception as e:
        st.error(f"Fail Excel tidak dapat dibaca: {e}")
        return {}

def find_exact_sheet(sheet_name):
    sheets = read_sheets()

    for name, df in sheets.items():
        if name.strip().lower() == sheet_name.strip().lower():
            return df

    return None

def import_participants_from_excel():
    imported = 0
    updated = 0
    skipped = 0

    df = find_exact_sheet("Participants")

    if df is None or df.empty:
        st.error("Sheet 'Participants' tidak dijumpai.")
        return 0, 0, 0

    for _, r in df.iterrows():
        email = clean(r.get("Email", "")).lower()

        if not email or "@" not in email:
            skipped += 1
            continue

        full_name = clean(r.get("Full_Name", "")) or email
        phone = clean(r.get("Phone", ""))
        organisation = clean(r.get("Organisation", ""))
        category = clean(r.get("Category", "Participant"))
        dinner = clean(r.get("Attend_Dinner", ""))

        dinner_flag = 1 if dinner.lower() in ["yes", "y", "1", "true"] else 0

        table_number = r.get("Table_Number", None)
        if pd.isna(table_number) or table_number == "":
            table_number = None
        else:
            try:
                table_number = int(table_number)
            except:
                table_number = None

        exists = df_sql(
            "SELECT id FROM participants WHERE lower(email)=lower(?)",
            (email,)
        )

        if exists.empty:
            exec_sql("""
                INSERT INTO participants
                (full_name,email,organisation,phone,participant_type,dinner_join,table_number,created_at)
                VALUES (?,?,?,?,?,?,?,?)
            """, (
                full_name, email, organisation, phone, category,
                dinner_flag, table_number,
                datetime.now().isoformat(timespec="seconds")
            ))
            imported += 1
        else:
            exec_sql("""
                UPDATE participants
                SET full_name=?,
                    organisation=?,
                    phone=?,
                    participant_type=?,
                    dinner_join=?,
                    table_number=?
                WHERE lower(email)=lower(?)
            """, (
                full_name, organisation, phone, category,
                dinner_flag, table_number, email
            ))
            updated += 1

    return imported, updated, skipped

def find_sheet(keywords):
    sheets = read_sheets()
    for name, df in sheets.items():
        low = name.lower()
        if any(k in low for k in keywords):
            return df
    return None

def find_col(df, candidates):
    cols = {str(c).lower().replace(" ","").replace("_",""): c for c in df.columns}
    for c in candidates:
        k = c.lower().replace(" ","").replace("_","")
        if k in cols:
            return cols[k]
    return None

def table_occupancy():
    df = df_sql("""
        SELECT table_number, COUNT(*) n
        FROM participants
        WHERE dinner_join=1 AND table_number IS NOT NULL
        GROUP BY table_number
    """)
    
    # Bina dictionary awal berdasarkan senarai TABLES rasmi
    occ = {t: 0 for t in TABLES}
    
    for _, r in df.iterrows():
        try:
            t_num = int(float(r["table_number"]))
            # Jika meja ada dalam senarai rasmi atau luar rasmi, ia tetap selamat digabungkan
            occ[t_num] = int(r["n"])
        except (ValueError, TypeError):
            # Mengelakkan crash jika ada data pelik seperti string kosong, "nan", atau None
            continue
            
    return occ
def next_table():
    occ = table_occupancy()
    for t in TABLES:
        if occ[t] < SEATS:
            return t
    return None

def show_poster(path):
    if path.exists():
        st.image(str(path), use_container_width=True)

def brand():
    st.markdown("""
    <div class="brand">
      <div class="brand-title">NICHE 2026</div>
      <div style="color:#cfc9df;letter-spacing:2px;text-transform:uppercase;font-size:12px;">
      Self Check-In System · Academic · Industry · Gala Dinner
      </div>
    </div>
    """, unsafe_allow_html=True)

# ================= PAGES =================
def page_home():
    show_poster(MAIN_POSTER)
    st.markdown("""
    <div class="card">
    <h2>Welcome to NICHE 2026</h2>
    <p>Participants register by email, check in using the same email, confirm dinner attendance, and view assigned dinner table.</p>
    <p>Door gift collection is controlled by staff/admin only.</p>
    </div>
    """, unsafe_allow_html=True)

def page_academic():
    st.header("🎓 Academic Conference Tentative & Abstracts")

    programme = find_exact_sheet("Academic_Programme")
    abstracts = find_exact_sheet("Abstracts")

    if programme is None or programme.empty:
        st.info("Academic_Programme belum ada dalam Excel.")
        return

    if abstracts is None:
        abstracts = pd.DataFrame()

    abs_map = {}
    if not abstracts.empty and "Paper_ID" in abstracts.columns:
        for _, r in abstracts.iterrows():
            pid = clean(r.get("Paper_ID", ""))
            if pid:
                abs_map[pid] = {
                    "title": clean(r.get("Title", "")),
                    "presenter": clean(r.get("Presenter", "")),
                    "email": clean(r.get("Email", "")),
                    "abstract": clean(r.get("Abstract_Text", "")),
                    "keywords": clean(r.get("Keywords", "")),
                    "affiliation": clean(r.get("Authors_Affiliation", ""))
                }

    non_paper_keywords = [
        "registration", "break", "lunch", "tea", "best award", 
        "best paper", "closing", "opening", "photo", "prayer", "certificate"
    ]

    # --- POPUP DIALOG MODAL ---
    @st.dialog("📝 Academic Research Paper Details", width="large")
    def show_abstract_popup(p_id, data):
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, rgba(255,240,163,0.1), rgba(212,175,55,0.05)); padding: 20px; border-radius: 14px; border-left: 5px solid #d4af37; margin-bottom: 20px;">
            <span style="background: linear-gradient(135deg,#fff0a3,#d4af37); color:#07103e; padding:4px 12px; border-radius:999px; font-weight:900; font-size:12px; letter-spacing:1px; text-transform:uppercase;">{p_id}</span>
            <h2 style="color: #ffe88a !important; font-size: 24px; font-weight: 900; margin-top: 12px; margin-bottom: 5px; line-height: 1.3;">{data['title']}</h2>
            <p style="color: #aaa4bd; font-size: 14px; margin: 0;">Presenter: <strong style="color: #fff;">{data['presenter']}</strong> ({data['email']})</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 👥 Authors & Affiliations")
        if data['affiliation']:
            formatted_aff = data['affiliation'].replace('\n', '<br>')
            st.markdown(f"""
            <div style="background: rgba(255,255,255,0.04); padding: 16px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08); font-size: 14px; color: #e2dfeb; line-height: 1.6;">
                {formatted_aff}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.caption("No affiliation metadata provided.")
            
        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("### 📖 Abstract")
        if data['abstract']:
            st.markdown(f"""
            <div style="background: rgba(5,10,46,0.6); padding: 20px; border-radius: 12px; border: 1px solid rgba(244,212,105,0.15); font-size: 15px; color: #cfc9df; line-height: 1.7; text-align: justify; max-height: 350px; overflow-y: auto;">
                {data['abstract']}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info("Abstract text not available for this presentation.")

        st.markdown("<br>", unsafe_allow_html=True)

        if data['keywords']:
            st.markdown(f"""
            <div style="font-size: 13px; color: #ffe88a; background: rgba(244,212,105,0.08); padding: 10px 15px; border-radius: 8px; display: inline-block;">
                🔑 <strong>Keywords:</strong> {data['keywords']}
            </div>
            """, unsafe_allow_html=True)

    # --- JADUAL AKADEMIK (MAIN) ---
    # Guna enumerate pada loop pertama (Sesi) untung mengelakkan sebarang pertembungan key global
    for s_idx, ((venue, time, session), g) in enumerate(programme.groupby(
        ["Venue", "Time", "Session"], dropna=False, sort=False
    )):
        venue = clean(venue)
        time = clean(time)
        session = clean(session)

        theme = clean(g.iloc[0].get("Theme", ""))
        moderator = clean(g.iloc[0].get("Moderator", ""))

        st.markdown(f"""
        <div class="session">
          <div class="session-grid">
            <div class="time">{time}</div>
            <div>
              <div class="label">{session} · {venue}</div>
              <div class="title">{theme if theme else 'Academic Session'}</div>
              <div style="color:#cfc9df;margin-top:8px;">Moderator: {moderator if moderator else '-'}</div>
            </div>
          </div>
        """, unsafe_allow_html=True)

        # Guna enumerate pada loop kedua (Paper) dengan indeks p_idx
        for p_idx, (_, r) in enumerate(g.iterrows()):
            paper_id = clean(r.get("Paper_ID", ""))
            title = clean(r.get("Title", ""))
            presenter = clean(r.get("Presenter", ""))
            email = clean(r.get("Email_From_Abstract", ""))

            title_low = title.lower()
            is_non_paper = (
                not paper_id or
                any(k in title_low for k in non_paper_keywords)
            )

            if is_non_paper:
                st.markdown(f"""
                <div class="paper">
                  <div class="abstract-title" style="font-size: 16px; color: #aaa4bd !important;">{title if title else session}</div>
                </div>
                """, unsafe_allow_html=True)
                continue

            st.markdown(f"""
            <div class="paper" style="border-left: 4px solid #d4af37;">
              <div class="paper-top">
                <span class="pid">{paper_id}</span>
                <span class="author">{presenter}</span>
              </div>
              <div class="abstract-title">{title}</div>
              <div style="color:#cfc9df;margin-top:5px;font-size:13px;margin-bottom:10px;">{email}</div>
            </div>
            """, unsafe_allow_html=True)
            
            paper_info = abs_map.get(paper_id, {
                "title": title, "presenter": presenter, "email": email,
                "abstract": "", "keywords": "", "affiliation": ""
            })
            
            # KEY DIBAIKI: Menggabungkan s_idx (sesi) dan p_idx (paper) untuk jaminan kelainan key yang unik
            if st.button(f"📄 View Abstract & Authors ({paper_id})", key=f"btn_{paper_id}_{s_idx}_{p_idx}"):
                show_abstract_popup(paper_id, paper_info)

        st.markdown("</div>", unsafe_allow_html=True)
def page_industry():
    st.header("🏛 Industrial Conference Tentative")

    df = find_exact_sheet("Industry_Programme")

    if df is None or df.empty:
        df = find_exact_sheet("Industrial_Programme")

    if df is None or df.empty:
        st.info("Industry_Programme belum ada dalam Excel.")
        return

    for _, r in df.iterrows():
        time_txt = clean(r.get("Time", ""))
        session = clean(r.get("Session", ""))
        title = clean(r.get("Title", ""))
        speaker = clean(r.get("Speaker", ""))
        venue = clean(r.get("Venue", ""))

        display_title = title if title else session

        st.markdown(f"""
        <div class="session">
          <div class="session-grid">
            <div class="time">{time_txt if time_txt else 'Time TBC'}</div>
            <div>
              <div class="label">Industrial Conference {('· ' + venue) if venue else ''}</div>
              <div class="title">{display_title if display_title else 'Industry Session'}</div>
              <div style="color:#cfc9df;margin-top:8px;">{speaker}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

def page_dinner():
    show_poster(DINNER_POSTER)
    st.header("✨ Gala Dinner Tentative")

    df = find_exact_sheet("Gala_Dinner_Programme")

    if df is None or df.empty:
        st.info("Gala_Dinner_Programme belum ada dalam Excel.")
        return

    for _, r in df.iterrows():
        time_txt = clean(r.get("Time", ""))
        event = clean(r.get("Event", ""))

        st.markdown(f"""
        <div class="session">
          <div class="session-grid">
            <div class="time">{time_txt}</div>
            <div>
              <div class="label">Gala Dinner Programme</div>
              <div class="title">{event}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

def page_register():
    st.header("📝 New Registration & Instant Check-In")

    # --- POP-UP MODAL MEWAH (KALIS DATA ROBOT & BOLEH TICK DOOR GIFT) ---
    @st.dialog("✨ Welcome to NICHE 2026!", width="large")
    def show_registration_success_popup(p_name, p_email, p_type, p_table, p_id):
        
        # --- LOGIK TAPISAN DATA ROBOT (BYTES TO INT) ---
        display_table = p_table
        if isinstance(p_table, bytes):
            try:
                display_table = int.from_bytes(p_table, byteorder='little')
            except:
                display_table = p_table

        st.markdown(f"""
        <div style="background: linear-gradient(135deg, rgba(255,240,163,0.1), rgba(212,175,55,0.05)); padding: 22px; border-radius: 18px; border: 1px solid rgba(244,212,105,0.3); margin-bottom: 20px; text-align: center;">
            <div style="font-size: 50px; margin-bottom: 10px;">🎉</div>
            <h2 style="color: #ffe88a !important; font-size: 26px; font-weight: 900; margin: 0 0 10px 0;">Registration & Check-In Complete!</h2>
            <p style="color: #fff; font-size: 18px; font-weight: 700; margin: 0;">{p_name}</p>
            <p style="color: #aaa4bd; font-size: 14px; margin: 4px 0 0 0;">{p_email} · <span style="color: #ffe88a;">{p_type}</span></p>
        </div>
        """, unsafe_allow_html=True)

        col_left, col_right = st.columns(2)
        
        # --- SEKSYEN PAPARAN MEJA DINNER ---
        with col_left:
            st.markdown("""<div style="text-align: center; margin-top: 10px;">""", unsafe_allow_html=True)
            st.subheader("🍽️ Gala Dinner Table")
            if display_table and str(display_table).strip() not in ["", "None", "nan"]:
                try:
                    table_num = int(float(display_table))
                    st.markdown(f"""
                    <div style="background: rgba(5,10,46,0.8); border: 2px solid #d4af37; border-radius: 14px; padding: 20px; font-size: 32px; font-weight: 900; color: #ffe88a; text-align: center; margin-bottom: 15px;">
                        TABLE {table_num}
                    </div>
                    """, unsafe_allow_html=True)
                except:
                    st.markdown(f"""
                    <div style="background: rgba(5,10,46,0.8); border: 2px solid #d4af37; border-radius: 14px; padding: 20px; font-size: 32px; font-weight: 900; color: #ffe88a; text-align: center; margin-bottom: 15px;">
                        TABLE {display_table}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="background: rgba(255,255,255,0.05); border: 1px dashed rgba(255,255,255,0.2); border-radius: 14px; padding: 20px; font-size: 20px; font-weight: 700; color: #aaa4bd; text-align: center; margin-bottom: 15px;">
                    ❌ Not Attending Dinner
                </div>
                """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # --- SEKSYEN TEKAN TICK DOOR GIFT ---
        with col_right:
            st.markdown("""<div style="text-align: center; margin-top: 10px;">""", unsafe_allow_html=True)
            st.subheader("🎁 Counter Door Gift")
            
            gift_key = f"reg_pop_gift_{p_id}"
            if gift_key not in st.session_state:
                st.session_state[gift_key] = False

            if st.session_state[gift_key]:
                st.markdown("""
                <div style="background: rgba(46,204,113,0.15); border: 1px solid #2ecc71; border-radius: 14px; padding: 24px; font-size: 18px; font-weight: 700; color: #2ecc71; text-align: center;">
                    ✅ Door Gift Taken
                </div>
                """, unsafe_allow_html=True)
            else:
                if st.button("🎁 Tick Taken / Ambil Door Gift", type="primary", use_container_width=True, key=f"btn_reg_pop_gift_{p_id}"):
                    exec_sql(
                        "UPDATE participants SET door_gift_collected=1, door_gift_time=? WHERE id=?",
                        (datetime.now().isoformat(timespec="seconds"), int(p_id))
                    )
                    st.session_state[gift_key] = True
                    st.success("Door gift collected successfully!")
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<br><hr style='border:0; border-top:1px solid rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
        if st.button("Finish & Clear Form", use_container_width=True, key="btn_close_reg_popup"):
            st.rerun()

    # --- BORANG INPUT PENDAFTARAN ---
    with st.form("reg_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Full Name *")
            email = st.text_input("Email *")
            org = st.text_input("Organisation")
        with c2:
            phone = st.text_input("Phone")
            ptype = st.selectbox("Participant Type", [
                "Academic Presenter",
                "Industry Participant",
                "Media",
                "Speaker",
                "Delegate",
                "Guest",
                "Walk-in"
            ])
            dinner = st.radio("Attend Gala Dinner?", ["Yes", "No"], horizontal=True)

        submit = st.form_submit_button("Register & Check-In Now")

    # --- PROSES AUTO-SAVE & AUTO-CHECK IN KE DATABASES ---
    if submit:
        if not name.strip() or not email.strip():
            st.error("Nama dan email wajib diisi.")
            return

        email2 = email.strip().lower()
        if "@" not in email2:
            st.error("Email tidak sah.")
            return

        exists = df_sql("SELECT id FROM participants WHERE lower(email)=lower(?)", (email2,))
        if not exists.empty:
            st.warning("Email ini sudah didaftarkan. Sila semak status di tab 'Check-In'.")
            return

        # Logik Agihan Meja: Media terus lock Table 28, lain-lain ikut baki kekosongan semeja max 8 orang
        if ptype == "Media":
            dinner_flag = 1
            table_no = 28
        else:
            dinner_flag = 1 if dinner == "Yes" else 0
            table_no = next_table() if dinner_flag else None

        # Set data terus kepada CHECKED-IN (checked_in=1 dan checkin_time dicop terus)
        now_str = datetime.now().isoformat(timespec="seconds")
        exec_sql("""
            INSERT INTO participants
            (full_name, email, organisation, phone, participant_type, dinner_join, table_number, checked_in, checkin_time, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            name.strip(), email2, org.strip(), phone.strip(), ptype,
            dinner_flag, table_no, 1, now_str, now_str
        ))

        # Dapatkan ID peserta baharu untuk tracking butang door gift di popup dialog
        new_id_df = df_sql("SELECT id FROM participants WHERE lower(email)=lower(?)", (email2,))
        new_id = int(new_id_df.iloc[0]["id"]) if not new_id_df.empty else 0

        log("REGISTER_AND_CHECKIN", f"Instant Check-In for {email2}, Assigned Table: {table_no}")
        
        # CETUSKAN POP-UP SERTA-MERTA SECARA KEKAL
        show_registration_success_popup(name.strip(), email2, ptype, table_no, new_id)
        
def page_checkin():
    st.header("✓ Self Check-In & Counter Services")

    email = st.text_input("Enter registered email").strip().lower()
    if not email:
        st.info("Masukkan email yang digunakan semasa register.")
        return

    df = df_sql("SELECT * FROM participants WHERE lower(email)=lower(?)", (email,))
    if df.empty:
        st.warning("Email tidak dijumpai. Sila register dahulu.")
        return

    p = df.iloc[0]
    p_id = int(p['id'])

    # --- DEKLARASI POP-UP MEJA DINNER YANG KEKAL (TIADA AUTO-RESET) ---
    @st.dialog("🍽️ Gala Dinner Seat Allocation", width="large")
    def show_dinner_table_popup(name, table_no):
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, rgba(255,240,163,0.1), rgba(212,175,55,0.05)); padding: 25px; border-radius: 18px; border: 1px solid rgba(244,212,105,0.3); text-align: center; margin-bottom: 15px;">
            <div style="font-size: 50px; margin-bottom: 10px;">🍽️</div>
            <h2 style="color: #ffe88a !important; font-size: 24px; font-weight: 900; margin: 0 0 10px 0;">Dinner Status Confirmed!</h2>
            <p style="color: #aaa4bd; font-size: 15px; margin: 0;">Participant Name:</p>
            <p style="color: #fff; font-size: 20px; font-weight: 800; margin: 4px 0 15px 0;">{name}</p>
        </div>
        """, unsafe_allow_html=True)
        
        if table_no:
            st.markdown(f"""
            <div style="background: #050a2e; border: 2px solid #d4af37; border-radius: 14px; padding: 25px; text-align: center; margin-bottom: 20px;">
                <span style="color: #aaa4bd; font-size: 14px; text-transform: uppercase; letter-spacing: 2px;">Your Assigned Seat</span>
                <h1 style="color: #ffe88a !important; font-size: 48px; font-weight: 900; margin: 10px 0 0 0;">TABLE {table_no}</h1>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background: rgba(255,255,255,0.05); border: 1px dashed rgba(255,255,255,0.2); border-radius: 14px; padding: 25px; text-align: center; margin-bottom: 20px; color: #aaa4bd; font-size: 18px; font-weight: 700;">
                ❌ Status: Not Attending Dinner
            </div>
            """, unsafe_allow_html=True)
            
        if st.button("OK, Close Window", use_container_width=True, key="btn_close_dinner_pop"):
            st.rerun()

    # --- KAD MAKLUMAT PESERTA ---
    st.markdown(f"""
    <div class="card">
    <h3>{clean(p['full_name'])}</h3>
    <p>Email: {clean(p['email'])}<br>
    Organisation: {clean(p['organisation']) or '-'}<br>
    Type: <span style="color: #ffe88a; font-weight: bold;">{clean(p['participant_type']) or '-'}</span></p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.info("Checked-in" if int(p["checked_in"]) else "Not checked-in")
    c2.info("Dinner: Yes" if int(p["dinner_join"]) else "Dinner: No")
    
    table_val = p["table_number"]
    try:
        if pd.notna(table_val) and str(table_val).strip() not in ["", "nan", "None"]:
            c3.success(f"Table {int(float(table_val))}")
        else:
            c3.warning("Table not assigned")
    except:
        c3.warning("Table not assigned")

    # ---------------- 1. BUTTON CONFIRM CHECK-IN ----------------
    if not int(p["checked_in"]):
        if st.button("Confirm Check-In", use_container_width=True, key="btn_main_checkin"):
            exec_sql(
                "UPDATE participants SET checked_in=1, checkin_time=? WHERE id=?",
                (datetime.now().isoformat(timespec="seconds"), p_id)
            )
            log("CHECKIN", email)
            st.success("Check-in successful.")
            st.rerun()
    else:
        st.success(f"Already checked in at: {clean(p['checkin_time'])}")

    st.markdown("---")

    # ---------------- 2. GALA DINNER & AUTOMATIC MEJA ----------------
    st.subheader("🍽️ Gala Dinner Confirmation")
    current = "Yes" if int(p["dinner_join"]) else "No"
    dinner_new = st.radio("Attend Gala Dinner?", ["Yes", "No"], index=0 if current=="Yes" else 1, horizontal=True, key="radio_dinner_confirm")

    if st.button("Update Dinner & Assign Table", use_container_width=True, key="btn_update_dinner_table"):
        flag = 1 if dinner_new == "Yes" else 0
        table_no = p["table_number"]
        p_type = clean(p['participant_type'])

        # Aturan agihan meja tersuai
        if flag and (pd.isna(table_no) or str(table_no).strip() in ["", "None", "nan"]):
            if p_type == "Media":
                table_no = 28
            else:
                table_no = next_table()
                if table_no is None:
                    st.error("Maaf, semua meja dinner (Maksimum 8 orang semeja) telah penuh!")
                    return
        elif not flag:
            table_no = None

        # Simpan ke pangkalan data
        exec_sql("UPDATE participants SET dinner_join=?, table_number=? WHERE id=?", (flag, table_no, p_id))
        log("DINNER_UPDATE", f"{email} updated dinner status to {dinner_new}")
        
        # CETUSKAN POPUP DAN BIARKAN DIA KEKAL (Rerun hanya berlaku bila user klik Close)
        show_dinner_table_popup(clean(p['full_name']), table_no)

    st.markdown("---")

    # ---------------- 3. COUNTER SELF DOOR GIFT ----------------
    st.subheader("🎁 Door Gift Collection")
    if int(p["door_gift_collected"]):
        st.success(f"✅ Door gift telah diambil pada: {clean(p['door_gift_time'])}")
    else:
        st.warning("Anda belum mengambil door gift.")
        if st.button("Ambil Door Gift Sekarang (Tick)", type="primary", use_container_width=True, key="btn_self_gift_tick"):
            exec_sql(
                "UPDATE participants SET door_gift_collected=1, door_gift_time=? WHERE id=?",
                (datetime.now().isoformat(timespec="seconds"), p_id)
            )
            log("DOOR_GIFT_SELF", email)
            st.success("Door gift berjaya ditandakan!")
            st.rerun()

def page_admin():
    st.header("🔐 Admin")

    if "admin_login" not in st.session_state:
        st.session_state.admin_login = False

    if not st.session_state.admin_login:
        pwd = st.text_input("Admin Password", type="password")
        if st.button("Login Admin"):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_login = True
                st.rerun()
            else:
                st.error("Wrong password.")
        return

    st.success("Admin login active.")

    st.markdown("## 🖼 Upload Poster")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Main Poster")
        main = st.file_uploader("Upload Industrial + Academic Poster", type=["jpg","jpeg","png"], key="mainposter")
        if st.button("Save Main Poster"):
            if save_file(main, MAIN_POSTER):
                st.success("Main poster saved.")
                st.rerun()
            else:
                st.warning("Pilih gambar dahulu.")
        show_poster(MAIN_POSTER)

    with col2:
        st.markdown("### Gala Dinner Poster")
        dinner = st.file_uploader("Upload Gala Dinner Poster", type=["jpg","jpeg","png"], key="dinnerposter")
        if st.button("Save Dinner Poster"):
            if save_file(dinner, DINNER_POSTER):
                st.success("Dinner poster saved.")
                st.rerun()
            else:
                st.warning("Pilih gambar dahulu.")
        show_poster(DINNER_POSTER)

    st.markdown("---")
    st.markdown("## 📘 Upload Excel Data")
    excel = st.file_uploader("Upload niche_data.xlsx", type=["xlsx"], key="exceldata")
    if st.button("Save Excel Data"):
        if save_file(excel, EXCEL_PATH):
            st.success("Excel saved as niche_data.xlsx.")
            st.rerun()
        else:
            st.warning("Pilih fail Excel dahulu.")

    if EXCEL_PATH.exists():
        sheets = read_sheets()
        st.info("Excel aktif: niche_data.xlsx")
        st.write("Sheets detected:", list(sheets.keys()))

        st.markdown("---")
    st.markdown("### 📥 Import Participants for Check-In")

    if st.button("Import Participants From Excel"):
        imported, updated, skipped = import_participants_from_excel()
        st.success(
            f"Import selesai. New: {imported} | Updated: {updated} | Skipped: {skipped}"
        )
        st.rerun()




    
    st.markdown("## 👥 Registered Participants")

    pdf = df_sql("SELECT * FROM participants ORDER BY id DESC")
    if pdf.empty:
        st.info("Belum ada peserta.")
    else:
        search = st.text_input("Search participant")
        if search:
            s = search.lower()
            pdf = pdf[
                pdf["full_name"].fillna("").str.lower().str.contains(s) |
                pdf["email"].fillna("").str.lower().str.contains(s) |
                pdf["organisation"].fillna("").str.lower().str.contains(s)
            ]

        st.dataframe(pdf, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("## 🍽 Table Assignment Summary")

    occ = table_occupancy()
    
    # Gabungkan senarai TABLES rasmi dengan meja 28 (Media) secara dinamik
    all_monitored_tables = TABLES + [28] if 28 not in TABLES else TABLES
    
    table_status_list = []
    for t in all_monitored_tables:
        occupied_count = occ.get(t, 0)
        
        # Bezakan kapasiti: Meja 28 khas untuk Media, meja lain terhad kepada SEATS (8 orang)
        if t == 28:
            capacity_label = "Khas Media"
            available_label = "-"
        else:
            capacity_label = SEATS
            available_label = max(0, SEATS - occupied_count)
            
        table_status_list.append({
            "Table Number": f"Table {t}",
            "Occupied (Pax)": occupied_count,
            "Max Capacity": capacity_label,
            "Available Seats": available_label
        })
        
    occ_df = pd.DataFrame(table_status_list)
    st.dataframe(occ_df, use_container_width=True, hide_index=True)

    dinner_df = df_sql("""
        SELECT id, full_name, email, organisation, table_number
        FROM participants
        WHERE dinner_join=1
        ORDER BY table_number IS NULL DESC, table_number, full_name
    """)

    dinner_df = df_sql("""
        SELECT id, full_name, email, organisation, table_number
        FROM participants
        WHERE dinner_join=1
        ORDER BY table_number IS NULL DESC, table_number, full_name
    """)

    for _, p in dinner_df.iterrows():
        c1, c2, c3 = st.columns([4, 1.5, 1])
        c1.write(f"**{p['full_name']}** \n{p['email']}")
        
        # --- FIXED CONVERSION START ---
        # Menukar data kepada integer secara selamat untuk mengelakkan ralat ValueError
        table_val = p["table_number"]
        cur = None
        try:
            if pd.notna(table_val) and str(table_val).strip() not in ["", "nan", "None"]:
                cur = int(float(table_val))
        except (ValueError, TypeError):
            cur = None
        # --- FIXED CONVERSION END ---

        options = [None] + TABLES
        
        # Memastikan indeks selectbox selamat jika cur tiada dalam TABLES rasmi
        try:
            sel_index = options.index(cur)
        except ValueError:
            sel_index = 0  # Default balik kepada None jika nombor meja di luar senarai rasmi

        new_table = c2.selectbox(
            "Table",
            options,
            index=sel_index,
            format_func=lambda x: "None" if x is None else f"Table {x}",
            key=f"table_{p['id']}",
            label_visibility="collapsed"
        )
        if c3.button("Save", key=f"save_table_{p['id']}"):
            if new_table is None:
                exec_sql("UPDATE participants SET table_number=NULL WHERE id=?", (int(p["id"]),))
            else:
                occ_now = table_occupancy()
                if new_table != cur and occ_now.get(new_table, 0) >= SEATS:
                    st.error(f"Table {new_table} penuh.")
                    continue
                exec_sql("UPDATE participants SET table_number=? WHERE id=?", (new_table, int(p["id"])))
            st.success("Table updated.")
            st.rerun()

    st.markdown("---")
    st.markdown("## 🎁 Door Gift")

    gift_df = df_sql("""
        SELECT id, full_name, email, door_gift_collected, door_gift_time
        FROM participants
        ORDER BY checked_in DESC, full_name
    """)

    gift_search = st.text_input("Search for door gift")
    if gift_search:
        s = gift_search.lower()
        gift_df = gift_df[
            gift_df["full_name"].fillna("").str.lower().str.contains(s) |
            gift_df["email"].fillna("").str.lower().str.contains(s)
        ]

    for _, p in gift_df.iterrows():
        c1, c2 = st.columns([4, 1])
        status = "✅ Collected" if int(p["door_gift_collected"]) else "Not collected"
        c1.write(f"**{p['full_name']}**  \n{p['email']}  \n{status}")

        if not int(p["door_gift_collected"]):
            if c2.button("Tick", key=f"gift_{p['id']}"):
                exec_sql(
                    "UPDATE participants SET door_gift_collected=1, door_gift_time=? WHERE id=?",
                    (datetime.now().isoformat(timespec="seconds"), int(p["id"]))
                )
                st.success("Door gift ticked.")
                st.rerun()
        else:
            if c2.button("Undo", key=f"undo_gift_{p['id']}"):
                exec_sql(
                    "UPDATE participants SET door_gift_collected=0, door_gift_time=NULL WHERE id=?",
                    (int(p["id"]),)
                )
                st.warning("Door gift undone.")
                st.rerun()

    st.markdown("---")
    st.markdown("---")
    st.markdown("## ➕ Walk-In Registration")

    # --- POP-UP MODAL EXCLUSIVE SELEPAS BERJAYA SAVE WALK-IN ---
    @st.dialog("✨ Registration & Check-In Success!", width="large")
    def show_walkin_success_popup(p_name, p_email, p_type, p_table, p_id):
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, rgba(255,240,163,0.1), rgba(212,175,55,0.05)); padding: 22px; border-radius: 18px; border: 1px solid rgba(244,212,105,0.3); margin-bottom: 20px; text-align: center;">
            <div style="font-size: 50px; margin-bottom: 10px;">✅</div>
            <h2 style="color: #ffe88a !important; font-size: 26px; font-weight: 900; margin: 0 0 10px 0;">Successfully Checked-In!</h2>
            <p style="color: #fff; font-size: 18px; font-weight: 700; margin: 0;">{p_name}</p>
            <p style="color: #aaa4bd; font-size: 14px; margin: 4px 0 0 0;">{p_email} · <span style="color: #ffe88a;">{p_type}</span></p>
        </div>
        """, unsafe_allow_html=True)

        col_left, col_right = st.columns(2)
        
        with col_left:
            st.markdown("""<div style="text-align: center; margin-top: 10px;">""", unsafe_allow_html=True)
            st.subheader("🍽️ Gala Dinner Table")
            if p_table:
                st.markdown(f"""
                <div style="background: rgba(5,10,46,0.8); border: 2px solid #d4af37; border-radius: 14px; padding: 20px; font-size: 32px; font-weight: 900; color: #ffe88a; text-align: center; margin-bottom: 15px;">
                    TABLE {p_table}
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="background: rgba(255,255,255,0.05); border: 1px dashed rgba(255,255,255,0.2); border-radius: 14px; padding: 20px; font-size: 20px; font-weight: 700; color: #aaa4bd; text-align: center; margin-bottom: 15px;">
                    ❌ Not Attending Dinner
                </div>
                """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with col_right:
            st.markdown("""<div style="text-align: center; margin-top: 10px;">""", unsafe_allow_html=True)
            st.subheader("🎁 Counter Door Gift")
            
            # Kita guna session state sementara supaya admin boleh tick direct dalam popup ini
            gift_key = f"popup_gift_{p_id}"
            if gift_key not in st.session_state:
                st.session_state[gift_key] = False

            if st.session_state[gift_key]:
                st.markdown("""
                <div style="background: rgba(46,204,113,0.15); border: 1px solid #2ecc71; border-radius: 14px; padding: 24px; font-size: 18px; font-weight: 700; color: #2ecc71; text-align: center;">
                    ✅ Door Gift Collected
                </div>
                """, unsafe_allow_html=True)
            else:
                if st.button("🎁 Tick Taken / Ambil Door Gift", type="primary", use_container_width=True, key=f"btn_pop_gift_{p_id}"):
                    exec_sql(
                        "UPDATE participants SET door_gift_collected=1, door_gift_time=? WHERE id=?",
                        (datetime.now().isoformat(timespec="seconds"), int(p_id))
                    )
                    st.session_state[gift_key] = True
                    st.success("Door gift marked as collected!")
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<br><hr style='border:0; border-top:1px solid rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
        if st.button("Close & Clear Form", use_container_width=True, key="btn_close_popup"):
            st.rerun()

    # --- BORANG WALK-IN INPUT ---
    with st.form("walkin", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            w_name = st.text_input("Walk-in Name")
            w_email = st.text_input("Walk-in Email")
            w_org = st.text_input("Walk-in Organisation")
        with c2:
            w_phone = st.text_input("Walk-in Phone")
            # Dropdown ditambah jenis "Media"
            w_type = st.selectbox("Walk-in Type", ["Walk-in", "Guest", "Media", "Industry Participant", "Academic Presenter"])
            w_dinner = st.radio("Dinner?", ["Yes", "No"], horizontal=True)

        w_submit = st.form_submit_button("Save & Check-In Participant")

    # --- PROSES SIMPAN DATA KE SQLITE ---
    if w_submit:
        if not w_name.strip() or not w_email.strip():
            st.error("Name and email required.")
        else:
            email2 = w_email.strip().lower()
            exists = df_sql("SELECT id FROM participants WHERE lower(email)=lower(?)", (email2,))
            
            if not exists.empty:
                st.warning("Email already exists. Please check under Registered Participants.")
            else:
                # Logik Auto-Assign Meja (Media dapat Table 28, lain-lain ikut kekosongan semeja max 8 orang)
                if w_type == "Media":
                    flag = 1
                    tbl = 28
                else:
                    flag = 1 if w_dinner == "Yes" else 0
                    tbl = next_table() if flag else None

                # Simpan terus sebagai data yang dah CHECKED-IN (checked_in=1)
                now_str = datetime.now().isoformat(timespec="seconds")
                exec_sql("""
                    INSERT INTO participants
                    (full_name, email, organisation, phone, participant_type, dinner_join, table_number, checked_in, checkin_time, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    w_name.strip(), email2, w_org.strip(), w_phone.strip(), w_type,
                    flag, tbl, 1, now_str, now_str
                ))
                
                # Ambil ID terakhir dimasukkan untuk tracking door gift di popup
                new_id_df = df_sql("SELECT id FROM participants WHERE lower(email)=lower(?)", (email2,))
                new_id = int(new_id_df.iloc[0]["id"]) if not new_id_df.empty else 0
                
                log("WALKIN_REGISTER", f"{email2} assigned to Table {tbl}")
                
                # CETUSKAN POP-UP MODAL YANG WOW
                show_walkin_success_popup(w_name.strip(), email2, w_type, tbl, new_id)

    st.markdown("---")
    st.markdown("## 🧹 Clear / Reset Data")

    st.warning("Bahagian ini untuk clear data peserta sahaja. Poster dan Excel tidak dibuang.")
    confirm = st.text_input("Type RESET to clear all participant data")

    if st.button("CLEAR ALL PARTICIPANTS DATA"):
        if confirm == "RESET":
            exec_sql("DELETE FROM participants")
            exec_sql("DELETE FROM sqlite_sequence WHERE name='participants'")
            log("RESET_PARTICIPANTS", "All participant data cleared")
            st.success("Semua data peserta telah dipadam.")
            st.rerun()
        else:
            st.error("Taip RESET dahulu.")

    st.warning("Bahagian bawah ini untuk clear poster dan Excel.")
    confirm2 = st.text_input("Type DELETEFILES to delete posters and Excel")

    if st.button("DELETE POSTERS AND EXCEL"):
        if confirm2 == "DELETEFILES":
            for p in [MAIN_POSTER, DINNER_POSTER, EXCEL_PATH]:
                if p.exists():
                    p.unlink()
            st.success("Poster dan Excel telah dipadam.")
            st.rerun()
        else:
            st.error("Taip DELETEFILES dahulu.")



# ================= MAIN =================
brand()

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🏠 Home",
    "🎓 Academic",
    "🏛 Industry",
    "✨ Gala Dinner",
    "✓ Check-In",
    "📝 Register",
    "🔐 Admin"
])

with tab1:
    page_home()
with tab2:
    page_academic()
with tab3:
    page_industry()
with tab4:
    page_dinner()
with tab5:
    page_checkin()
with tab6:
    page_register()
with tab7:
    page_admin()
