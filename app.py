from pathlib import Path
from datetime import datetime
import sqlite3, hashlib, base64, io, re, zipfile

import pandas as pd
import streamlit as st
import qrcode

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
UPLOADS = BASE / "uploads"
ASSETS = BASE / "assets"
DB = DATA / "niche2026.db"
DATA.mkdir(exist_ok=True); UPLOADS.mkdir(exist_ok=True)

st.set_page_config(page_title="NICHE 2026", page_icon="✨", layout="wide", initial_sidebar_state="collapsed")

ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin") if hasattr(st, "secrets") else "admin"
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "NICHE2026admin") if hasattr(st, "secrets") else "NICHE2026admin"

CSS = """
<style>
:root{--navy:#07114A;--navy2:#101A5C;--gold:#D4AF37;--gold2:#F6E08E;--white:#fff;}
.stApp{background: radial-gradient(circle at top left,#17237a 0,#07114A 35%,#030728 100%); color:white;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#050b33,#101A5C)!important;}
.block-container{padding-top:1.2rem; max-width:1120px;}
.hero{border:1px solid rgba(212,175,55,.38); border-radius:28px; padding:22px; background:linear-gradient(135deg,rgba(255,255,255,.10),rgba(255,255,255,.03)); box-shadow:0 18px 50px rgba(0,0,0,.35);}
.card{border:1px solid rgba(212,175,55,.25); border-radius:24px; padding:20px; background:rgba(255,255,255,.08); box-shadow:0 10px 30px rgba(0,0,0,.22); margin-bottom:16px;}
.gold-title{font-size:2rem; font-weight:900; color:#F6D35C; letter-spacing:.5px; margin:0;}
.sub{color:rgba(255,255,255,.82); font-size:1.02rem;}
.badge{display:inline-block; padding:.35rem .75rem; border:1px solid rgba(212,175,55,.5); border-radius:999px; color:#FFE79B; background:rgba(212,175,55,.12); margin:4px 4px 4px 0; font-weight:700;}
.ok{color:#9DFFC8;font-weight:800}.warn{color:#FFE79B;font-weight:800}.danger{color:#FFB3B3;font-weight:800}
.stButton>button,.stDownloadButton>button{border-radius:999px!important; border:1px solid rgba(212,175,55,.55)!important; background:linear-gradient(135deg,#7A5615,#D4AF37,#FFEAA2)!important; color:#07114A!important; font-weight:900!important;}
input, textarea, select{border-radius:14px!important;}
hr{border-color:rgba(212,175,55,.25)}
.small{font-size:.88rem;color:rgba(255,255,255,.72)}
@media(max-width:700px){.gold-title{font-size:1.45rem}.block-container{padding-left:.75rem;padding-right:.75rem}.hero,.card{border-radius:18px;padding:15px}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# ---------- Database ----------
def conn():
    return sqlite3.connect(DB, check_same_thread=False)

def init_db():
    c=conn(); cur=c.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS participants(
        registration_id TEXT PRIMARY KEY, full_name TEXT, email TEXT UNIQUE, category TEXT, institution TEXT, phone TEXT,
        participant_type TEXT DEFAULT 'Pre-Registered', table_no TEXT, seat_no TEXT,
        checked_in INTEGER DEFAULT 0, door_gift INTEGER DEFAULT 0, dinner_rsvp TEXT DEFAULT 'Not Submitted', dinner_table TEXT,
        qr_code TEXT, notes TEXT, created_at TEXT, updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS abstracts(
        paper_id TEXT PRIMARY KEY, abstract_no TEXT, presenter_email TEXT, presenter_name TEXT, institution TEXT,
        title TEXT, abstract TEXT, keywords TEXT, track TEXT, session TEXT, date TEXT, time TEXT, room TEXT, moderator TEXT,
        presentation_type TEXT DEFAULT 'Oral Presentation', updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS programme(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, time TEXT, title TEXT, speaker TEXT, venue TEXT, category TEXT, details TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS uploads(
        name TEXT PRIMARY KEY, filename TEXT, path TEXT, uploaded_at TEXT)""")
    c.commit(); c.close()
init_db()

def safe(x):
    if pd.isna(x): return ""
    return str(x).strip()

def norm_email(x): return safe(x).lower()

def make_reg_id(prefix="NICHE2026"):
    c=conn(); cur=c.cursor(); cur.execute("SELECT COUNT(*) FROM participants"); n=cur.fetchone()[0]+1; c.close()
    return f"{prefix}-{n:04d}"

def qr_img(text):
    img=qrcode.make(text)
    buf=io.BytesIO(); img.save(buf, format="PNG")
    return buf.getvalue()

def read_excel_any(file):
    name=file.name.lower()
    engine="xlrd" if name.endswith(".xls") else "openpyxl"
    return pd.read_excel(file, sheet_name=None, engine=engine)

def find_col(df, options):
    cols={re.sub(r"[^a-z0-9]","",str(c).lower()):c for c in df.columns}
    for opt in options:
        key=re.sub(r"[^a-z0-9]","",opt.lower())
        if key in cols: return cols[key]
    for key,c in cols.items():
        for opt in options:
            if re.sub(r"[^a-z0-9]","",opt.lower()) in key: return c
    return None

def import_participants(file):
    sheets=read_excel_any(file); rows=[]
    for s,df in sheets.items():
        df=df.dropna(how="all")
        if df.empty: continue
        name_col=find_col(df,["Full Name","Name","Nama","Presenter","Participant Name"])
        email_col=find_col(df,["Email","E-mail","Emel"])
        cat_col=find_col(df,["Category","Kategori","Role","Type","Status"])
        inst_col=find_col(df,["Institution","Organisation","Organization","Company","Affiliation","Universiti"])
        phone_col=find_col(df,["Phone","Mobile","No Telefon","Tel"])
        table_col=find_col(df,["Table","Table No","Meja"])
        if not name_col: continue
        for _,r in df.iterrows():
            full=safe(r.get(name_col,"")); email=norm_email(r.get(email_col,"")) if email_col else ""
            if not full: continue
            rid = make_reg_id()
            qr = rid
            cat = safe(r.get(cat_col,"")) if cat_col else ("Keynote / Moderator" if "keynote" in s.lower() else "Academic Participant")
            inst = safe(r.get(inst_col,"")) if inst_col else ""
            phone = safe(r.get(phone_col,"")) if phone_col else ""
            table = safe(r.get(table_col,"")) if table_col else ""
            rows.append((rid,full,email,cat,inst,phone,"Pre-Registered",table,"",0,0,"Not Submitted","",qr,"Imported from "+s,datetime.now().isoformat(),datetime.now().isoformat()))
    c=conn(); cur=c.cursor(); count=0
    for row in rows:
        try:
            if row[2]:
                cur.execute("""INSERT INTO participants VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(email) DO UPDATE SET full_name=excluded.full_name, category=excluded.category, institution=excluded.institution,
                phone=excluded.phone, table_no=excluded.table_no, updated_at=excluded.updated_at""", row)
            else:
                cur.execute("INSERT OR REPLACE INTO participants VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row)
            count+=1
        except Exception:
            pass
    c.commit(); c.close(); return count

def import_abstracts(file):
    sheets=read_excel_any(file); rows=[]
    for s,df in sheets.items():
        df=df.dropna(how="all")
        if df.empty: continue
        paper=find_col(df,["Paper ID","ID","Abstract ID"]); no=find_col(df,["Abstract No","No"])
        email=find_col(df,["Presenter Email","Email"]); presenter=find_col(df,["Presenter Name","Presenter"])
        inst=find_col(df,["Institution","Affiliation"]); title=find_col(df,["Paper Title","Title"])
        abst=find_col(df,["Abstract","Abstrak"]); keywords=find_col(df,["Keywords","Kata Kunci"])
        track=find_col(df,["Track","Theme"]); session=find_col(df,["Session"]); date=find_col(df,["Date"])
        time=find_col(df,["Time"]); room=find_col(df,["Room","Venue"]); mod=find_col(df,["Moderator"]); ptype=find_col(df,["Presentation Type","Type"])
        if not (paper or title or presenter): continue
        for idx,r in df.iterrows():
            titlev=safe(r.get(title,"")) if title else ""; pres=safe(r.get(presenter,"")) if presenter else ""
            if not titlev and not pres: continue
            pid=safe(r.get(paper,"")) if paper else f"ABS-{idx+1:03d}"
            rows.append((pid,safe(r.get(no,"")) if no else str(idx+1), norm_email(r.get(email,"")) if email else "", pres,
                         safe(r.get(inst,"")) if inst else "", titlev, safe(r.get(abst,"")) if abst else "",
                         safe(r.get(keywords,"")) if keywords else "", safe(r.get(track,"")) if track else "",
                         safe(r.get(session,"")) if session else s, safe(r.get(date,"")) if date else "10 June 2026",
                         safe(r.get(time,"")) if time else "", safe(r.get(room,"")) if room else "",
                         safe(r.get(mod,"")) if mod else "", safe(r.get(ptype,"")) if ptype else "Oral Presentation", datetime.now().isoformat()))
    c=conn(); cur=c.cursor(); count=0
    for row in rows:
        cur.execute("""INSERT INTO abstracts VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(paper_id) DO UPDATE SET abstract_no=excluded.abstract_no,presenter_email=excluded.presenter_email,presenter_name=excluded.presenter_name,
        institution=excluded.institution,title=excluded.title,abstract=excluded.abstract,keywords=excluded.keywords,track=excluded.track,session=excluded.session,
        date=excluded.date,time=excluded.time,room=excluded.room,moderator=excluded.moderator,presentation_type=excluded.presentation_type,updated_at=excluded.updated_at""", row)
        count+=1
    c.commit(); c.close(); return count

def df_sql(q, params=()):
    c=conn(); df=pd.read_sql_query(q,c,params=params); c.close(); return df

def execute(q, params=()):
    c=conn(); cur=c.cursor(); cur.execute(q,params); c.commit(); c.close()

def save_upload(file, key):
    dest=UPLOADS / file.name
    dest.write_bytes(file.getvalue())
    execute("INSERT OR REPLACE INTO uploads VALUES (?,?,?,?)", (key,file.name,str(dest),datetime.now().isoformat()))
    return dest

# ---------- UI ----------
def top():
    poster=ASSETS/"main_poster.jpeg"
    st.markdown('<div class="hero">', unsafe_allow_html=True)
    c1,c2=st.columns([1,1.4])
    with c1:
        if poster.exists(): st.image(str(poster), use_container_width=True)
    with c2:
        st.markdown('<p class="gold-title">NICHE 2026</p>', unsafe_allow_html=True)
        st.markdown('<p class="sub">Negeri Sembilan International Conference on Halal & Sustainability Ecosystems</p>', unsafe_allow_html=True)
        st.markdown('<span class="badge">9–10 June 2026</span><span class="badge">Royale Chulan Seremban</span><span class="badge">Official Registration Portal</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

def participant_lookup():
    email=norm_email(st.text_input("Enter your registered email", placeholder="name@example.com"))
    if st.button("View My Registration", use_container_width=True):
        st.session_state["participant_email"]=email
    if st.session_state.get("participant_email"):
        show_participant(st.session_state["participant_email"])

def show_participant(email):
    p=df_sql("SELECT * FROM participants WHERE lower(email)=?", (email,)) if email else pd.DataFrame()
    if p.empty:
        st.warning("Registration record not found. Please proceed to the registration counter for assistance.")
        return
    r=p.iloc[0]
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f"### Welcome, {r.full_name}")
    a,b=st.columns([1.2,.8])
    with a:
        st.write(f"**Category:** {r.category}")
        st.write(f"**Institution:** {r.institution}")
        st.write(f"**Registration ID:** {r.registration_id}")
        st.write(f"**Seating:** Table {r.table_no or 'To be assigned'} {r.seat_no or ''}")
        st.markdown(f"**Check-In:** {'<span class=ok>Checked In</span>' if r.checked_in else '<span class=warn>Registered – Please proceed to counter</span>'}", unsafe_allow_html=True)
        st.markdown(f"**Door Gift:** {'<span class=ok>Collected</span>' if r.door_gift else '<span class=warn>Ready for Collection</span>'}", unsafe_allow_html=True)
        st.write(f"**Gala Dinner:** {r.dinner_rsvp}")
    with b:
        st.image(qr_img(r.qr_code or r.registration_id), caption="Registration QR", width=180)
    st.markdown('</div>', unsafe_allow_html=True)
    prs=df_sql("SELECT * FROM abstracts WHERE lower(presenter_email)=? OR lower(presenter_name) LIKE ?", (email, f"%{str(r.full_name).lower()}%"))
    if not prs.empty:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Presentation Information")
        for _,x in prs.iterrows():
            st.markdown(f"**{x.paper_id} — {x.title}**")
            st.write(f"{x.date} | {x.time} | {x.room}")
            st.write(f"Session: {x.session} | Moderator: {x.moderator}")
            with st.expander("View Abstract"):
                st.write(x.abstract)
        st.markdown('</div>', unsafe_allow_html=True)

def public_pages():
    menu=st.sidebar.radio("Menu", ["Home","My Registration","Conference Programme","Academic Conference","Industry Conference","Abstract Book","Gala Dinner","Venue & Seating","Contact Secretariat","Admin Portal"])
    if menu=="Home": top()
    elif menu=="My Registration": top(); participant_lookup()
    elif menu=="Conference Programme": show_programme("All")
    elif menu=="Academic Conference": show_programme("Academic")
    elif menu=="Industry Conference": show_programme("Industry")
    elif menu=="Abstract Book": show_abstracts()
    elif menu=="Gala Dinner": show_dinner()
    elif menu=="Venue & Seating": show_seating()
    elif menu=="Contact Secretariat": st.markdown('<div class="card"><h3>Contact Secretariat</h3><p>Please proceed to the registration counter for assistance.</p></div>', unsafe_allow_html=True)
    else: admin_login()

def show_programme(kind="All"):
    st.markdown('<p class="gold-title">Conference Programme</p>', unsafe_allow_html=True)
    up=df_sql("SELECT * FROM uploads WHERE name='programme_pdf'")
    if not up.empty and Path(up.iloc[0].path).exists():
        data=Path(up.iloc[0].path).read_bytes()
        st.download_button("Download Programme Book (PDF)", data, file_name=up.iloc[0].filename, mime="application/pdf", use_container_width=True)
    st.info("The programme book will be available here after it is uploaded by the organizer.")
    if kind=="Academic":
        df=df_sql("SELECT paper_id,title,presenter_name,session,date,time,room,moderator FROM abstracts ORDER BY room,time,paper_id")
        st.dataframe(df, use_container_width=True, hide_index=True)
    elif kind=="Industry":
        df=df_sql("SELECT * FROM programme WHERE category LIKE '%Industry%' ORDER BY day,time")
        st.dataframe(df, use_container_width=True, hide_index=True) if not df.empty else st.info("Industry programme will appear after admin upload/update.")
    else:
        df=df_sql("SELECT * FROM programme ORDER BY day,time")
        if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)

def show_abstracts():
    st.markdown('<p class="gold-title">Abstract Book</p>', unsafe_allow_html=True)
    df=df_sql("SELECT * FROM abstracts ORDER BY CAST(abstract_no AS INTEGER), paper_id")
    if df.empty: st.info("Abstracts will appear after admin upload."); return
    q=st.text_input("Search by Paper ID, Presenter, Title or Keyword")
    if q:
        ql=q.lower(); df=df[df.apply(lambda r: ql in " ".join(map(str,r.values)).lower(), axis=1)]
    for _,r in df.iterrows():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f"**Abstract {r.abstract_no or r.paper_id}**  ")
        st.markdown(f"### {r.title}")
        st.write(f"Presenter: {r.presenter_name}")
        st.write(f"{r.date} | {r.time} | {r.room}")
        with st.expander("View Abstract"):
            st.write(r.abstract)
            if r.keywords: st.write("Keywords:", r.keywords)
        st.markdown('</div>', unsafe_allow_html=True)

def show_dinner():
    st.markdown('<p class="gold-title">Gala Dinner</p>', unsafe_allow_html=True)
    st.write("9 June 2026 | Grand Ballroom | Royale Chulan Seremban")
    email=norm_email(st.text_input("Enter your registered email for dinner RSVP"))
    choice=st.radio("Attendance", ["I Will Attend", "Unable to Attend"], horizontal=True)
    if st.button("Submit Gala Dinner RSVP", use_container_width=True):
        execute("UPDATE participants SET dinner_rsvp=?, updated_at=? WHERE lower(email)=?", (choice, datetime.now().isoformat(), email))
        st.success("Dinner RSVP updated.")

def show_seating():
    st.markdown('<p class="gold-title">Venue & Seating</p>', unsafe_allow_html=True)
    img=ASSETS/"seating_layout.png"
    if img.exists(): st.image(str(img), use_container_width=True)
    else: st.info("Seating layout will be available after admin upload.")

def admin_login():
    st.markdown('<p class="gold-title">Admin Portal</p>', unsafe_allow_html=True)
    if st.session_state.get("admin_ok"):
        admin_panel(); return
    u=st.text_input("Username")
    p=st.text_input("Password", type="password")
    if st.button("Login", use_container_width=True):
        if u==ADMIN_USERNAME and p==ADMIN_PASSWORD:
            st.session_state["admin_ok"]=True; st.rerun()
        else: st.error("Invalid admin credentials.")

def admin_panel():
    st.success("Admin access granted.")
    tab=st.tabs(["Upload Data","Participants","Check-In","Walk-In","Abstracts","Programme/PDF","Reports"])
    with tab[0]:
        st.subheader("Upload Participant / Presenter / Keynote Excel")
        f=st.file_uploader("Excel file", type=["xlsx","xls"], key="pfile")
        if f and st.button("Import Participants", use_container_width=True): st.success(f"Imported/updated {import_participants(f)} records.")
        st.subheader("Upload Abstract + Academic Schedule Excel")
        a=st.file_uploader("Abstract Excel", type=["xlsx","xls"], key="afile")
        if a and st.button("Import Abstracts", use_container_width=True): st.success(f"Imported/updated {import_abstracts(a)} abstracts.")
    with tab[1]:
        df=df_sql("SELECT * FROM participants ORDER BY full_name")
        st.dataframe(df, use_container_width=True, hide_index=True)
    with tab[2]:
        q=st.text_input("Search name / email / registration ID / QR")
        if q:
            df=df_sql("SELECT * FROM participants WHERE full_name LIKE ? OR email LIKE ? OR registration_id LIKE ? OR qr_code LIKE ?", tuple([f"%{q}%"]*4))
            for _,r in df.iterrows():
                st.markdown(f"### {r.full_name}")
                st.write(r.email, r.category, r.registration_id)
                c1,c2,c3=st.columns(3)
                if c1.button("Check In", key="ci"+r.registration_id): execute("UPDATE participants SET checked_in=1 WHERE registration_id=?",(r.registration_id,)); st.rerun()
                if c2.button("Door Gift Collected", key="dg"+r.registration_id): execute("UPDATE participants SET door_gift=1 WHERE registration_id=?",(r.registration_id,)); st.rerun()
                tbl=c3.text_input("Table", value=r.table_no or "", key="tb"+r.registration_id)
                if c3.button("Save Table", key="sv"+r.registration_id): execute("UPDATE participants SET table_no=? WHERE registration_id=?",(tbl,r.registration_id)); st.rerun()
    with tab[3]:
        st.subheader("Walk-In Registration")
        with st.form("walkin"):
            name=st.text_input("Full Name *"); email=norm_email(st.text_input("Email *")); phone=st.text_input("Phone")
            inst=st.text_input("Institution / Company")
            cat=st.selectbox("Category", ["Academic Participant","Industry Delegate","Invited Guest","VIP","Student Visitor","Sponsor","Media"])
            table=st.text_input("Table No")
            submitted=st.form_submit_button("Register Walk-In")
        if submitted and name and email:
            rid=make_reg_id("NICHE-WALKIN"); now=datetime.now().isoformat()
            execute("INSERT OR REPLACE INTO participants VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (rid,name,email,cat,inst,phone,"Walk-In",table,"",1,0,"Not Submitted","",rid,"Walk-in registration",now,now))
            st.success(f"Walk-in registered: {rid}")
        st.caption("Presenter walk-in is not enabled. Presenter records must be verified by the secretariat.")
    with tab[4]:
        show_abstracts()
    with tab[5]:
        pdf=st.file_uploader("Upload Programme Book PDF", type=["pdf"])
        if pdf and st.button("Save Programme PDF", use_container_width=True):
            save_upload(pdf,"programme_pdf"); st.success("Programme book saved.")
        seat=st.file_uploader("Upload Seating Layout Image", type=["png","jpg","jpeg"])
        if seat and st.button("Save Seating Layout", use_container_width=True):
            (ASSETS/"seating_layout.png").write_bytes(seat.getvalue()); st.success("Seating layout saved.")
    with tab[6]:
        df=df_sql("SELECT * FROM participants")
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Total", len(df)); c2.metric("Checked In", int(df.checked_in.sum()) if not df.empty else 0); c3.metric("Door Gift", int(df.door_gift.sum()) if not df.empty else 0); c4.metric("Dinner Attend", int((df.dinner_rsvp=='I Will Attend').sum()) if not df.empty else 0)
        st.download_button("Export Participants CSV", df.to_csv(index=False).encode(), "niche2026_participants_export.csv", "text/csv", use_container_width=True)
        if DB.exists(): st.download_button("Download Database Backup", DB.read_bytes(), "niche2026.db", "application/octet-stream", use_container_width=True)

public_pages()
