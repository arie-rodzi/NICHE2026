from pathlib import Path
import io, sqlite3, base64
import pandas as pd
import streamlit as st
from PIL import Image
import qrcode

BASE = Path(__file__).resolve().parent
DATA = BASE / 'data'
ASSETS = BASE / 'assets'
DB = DATA / 'niche2026.db'
MASTER_DEFAULT = DATA / 'NICHE2026_MASTER.xlsx'
ADMIN_USER = 'admin'
ADMIN_PASS = 'NICHE2026admin'

st.set_page_config(page_title='NICHE 2026', page_icon='✨', layout='wide', initial_sidebar_state='collapsed')

CSS = '''
<style>
.stApp{background:linear-gradient(135deg,#02073a 0%,#07185a 55%,#06113f 100%); color:#fff;}
.block-container{padding-top:1rem;max-width:1180px;}
[data-testid="stHeader"]{background:rgba(0,0,0,0)}
.card{background:rgba(255,255,255,.08); border:1px solid rgba(255,214,92,.22); border-radius:22px; padding:22px; box-shadow:0 18px 55px rgba(0,0,0,.25); margin:12px 0;}
.hero{background:linear-gradient(135deg,rgba(255,215,105,.18),rgba(255,255,255,.06)); border:1px solid rgba(255,217,102,.35); border-radius:28px; padding:28px; margin-bottom:16px;}
.gold{color:#ffd86b!important;font-weight:800}.muted{color:#cfd7ff}.small{font-size:.9rem}.status{padding:8px 13px;border-radius:999px;background:rgba(255,216,107,.16);display:inline-block;margin:4px 4px 4px 0}
.stTabs [data-baseweb="tab-list"]{gap:8px; flex-wrap:wrap}.stTabs [data-baseweb="tab"]{background:rgba(255,255,255,.08);border-radius:999px;padding:8px 14px;color:white}.stTabs [aria-selected="true"]{background:#d7a72d;color:#07124a}
button[kind="primary"], .stDownloadButton button{background:linear-gradient(90deg,#b8860b,#ffdf75)!important;color:#07124a!important;border:0!important;border-radius:14px!important;font-weight:800!important}
.stTextInput input, .stSelectbox div[data-baseweb="select"]>div, .stTextArea textarea{background:rgba(255,255,255,.95)!important;color:#111!important;border-radius:14px!important}
hr{border-color:rgba(255,255,255,.15)}
</style>
'''
st.markdown(CSS, unsafe_allow_html=True)

def clean_df(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all')
    return df

def read_sheet(file, sheet):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=2)
        return clean_df(df)
    except Exception:
        return pd.DataFrame()

def init_db():
    DATA.mkdir(exist_ok=True)
    con=sqlite3.connect(DB)
    cur=con.cursor()
    cur.execute('CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)')
    for t in ['participants','abstracts','programme','industry']:
        cur.execute(f'CREATE TABLE IF NOT EXISTS {t} (data TEXT)')
    cur.execute('CREATE TABLE IF NOT EXISTS status (email TEXT PRIMARY KEY, checkin TEXT, doorgift TEXT, dinner TEXT, table_no TEXT, seat_no TEXT, dinner_table TEXT)')
    con.commit(); con.close()

def save_df(name, df):
    con=sqlite3.connect(DB); cur=con.cursor(); cur.execute(f'DELETE FROM {name}')
    if not df.empty:
        cur.execute(f'INSERT INTO {name}(data) VALUES (?)', (df.to_json(orient='records'),))
    con.commit(); con.close()

def load_df(name):
    con=sqlite3.connect(DB); cur=con.cursor(); cur.execute(f'SELECT data FROM {name} LIMIT 1'); row=cur.fetchone(); con.close()
    if row and row[0]:
        try: return pd.read_json(io.StringIO(row[0]))
        except Exception: return pd.DataFrame()
    return pd.DataFrame()

def import_master(file):
    parts = read_sheet(file,'01_PARTICIPANTS_MASTER')
    abstracts = read_sheet(file,'02_ABSTRACT_SCHEDULE')
    industry = read_sheet(file,'03_INDUSTRY_KEYNOTE')
    programme = read_sheet(file,'04_PROGRAMME_EVENTS')
    save_df('participants', parts); save_df('abstracts', abstracts); save_df('industry', industry); save_df('programme', programme)
    return parts, abstracts, industry, programme

def get_status(email):
    con=sqlite3.connect(DB); cur=con.cursor(); cur.execute('SELECT * FROM status WHERE lower(email)=?', (email.lower(),)); r=cur.fetchone(); con.close()
    if r: return {'email':r[0],'checkin':r[1],'doorgift':r[2],'dinner':r[3],'table_no':r[4],'seat_no':r[5],'dinner_table':r[6]}
    return {'email':email,'checkin':'Not Checked In','doorgift':'Not Collected','dinner':'Not Submitted','table_no':'','seat_no':'','dinner_table':''}

def update_status(email, **kw):
    s=get_status(email); s.update({k:v for k,v in kw.items() if v is not None})
    con=sqlite3.connect(DB); cur=con.cursor(); cur.execute('REPLACE INTO status VALUES (?,?,?,?,?,?,?)',(email,s['checkin'],s['doorgift'],s['dinner'],s['table_no'],s['seat_no'],s['dinner_table'])); con.commit(); con.close()

def qr_img(text):
    img=qrcode.make(text); buf=io.BytesIO(); img.save(buf, format='PNG'); return buf.getvalue()

def img_tag(path, width='100%'):
    if Path(path).exists():
        st.image(str(path), use_container_width=True)
    else:
        st.warning(f'Image not found: {path.name}')

def pdf_download(path, label):
    if Path(path).exists():
        st.download_button(label, data=Path(path).read_bytes(), file_name=Path(path).name, mime='application/pdf')
    else:
        st.info('File belum dimuat naik oleh admin.')

def card(title, body=''):
    st.markdown(f"<div class='card'><h3 class='gold'>{title}</h3>{body}</div>", unsafe_allow_html=True)

def ensure_data_loaded():
    if load_df('participants').empty and MASTER_DEFAULT.exists():
        import_master(MASTER_DEFAULT)

init_db(); ensure_data_loaded()

parts=load_df('participants'); abstracts=load_df('abstracts'); programme=load_df('programme'); industry=load_df('industry')

st.markdown("<div class='hero'><h1 class='gold'>NICHE 2026</h1><p class='muted'>Negeri Sembilan International Conference on Halal & Sustainability Ecosystems<br>9–10 June 2026 · Royale Chulan Seremban</p></div>", unsafe_allow_html=True)

mode = st.sidebar.radio('Portal', ['Participant Portal','Admin Portal'])

if mode=='Participant Portal':
    tabs=st.tabs(['Home','My Registration','Conference Programme','Academic Conference','Industry Conference','Abstract Book','Gala Dinner','Venue & Seating','Contact'])
    with tabs[0]:
        st.subheader('Official Event Poster')
        img_tag(ASSETS/'main_poster.jpeg')
        with st.expander('Conference Poster / Call for Paper'):
            img_tag(ASSETS/'conference_poster.jpeg')
    with tabs[1]:
        email=st.text_input('Enter your registered email').strip()
        if email:
            if 'Email' in parts.columns:
                rec=parts[parts['Email'].astype(str).str.lower().str.strip()==email.lower()]
            else: rec=pd.DataFrame()
            if rec.empty:
                st.error('Registration record not found. Please proceed to the registration counter for assistance.')
            else:
                r=rec.iloc[0].to_dict(); s=get_status(email); rid=str(r.get('Registration ID', r.get('QR Code','NICHE2026')))
                st.markdown(f"<div class='card'><h2 class='gold'>Welcome, {r.get('Full Name','Participant')}</h2><p>{r.get('Category','')}</p><p>{r.get('Institution','')}</p></div>", unsafe_allow_html=True)
                c1,c2=st.columns([2,1])
                with c1:
                    st.write('**Registration ID:**', rid); st.write('**Email:**', email); st.write('**Table:**', s.get('table_no') or r.get('Table No','To be assigned'))
                    st.markdown(f"<span class='status'>Check-In: {s['checkin']}</span><span class='status'>Door Gift: {s['doorgift']}</span><span class='status'>Dinner: {s['dinner']}</span>", unsafe_allow_html=True)
                with c2: st.image(qr_img(rid), caption='Registration QR')
                if not abstracts.empty and 'Presenter Email' in abstracts.columns:
                    mine=abstracts[abstracts['Presenter Email'].astype(str).str.lower().str.strip()==email.lower()]
                    if not mine.empty:
                        st.subheader('Presentation Information')
                        st.dataframe(mine[[c for c in ['Paper ID','Abstract No','Paper Title','Session','Date','Time','Room','Moderator'] if c in mine.columns]], use_container_width=True)
        else: st.info('Please enter your registered email to view your registration details.')
    with tabs[2]:
        st.subheader('Conference Programme')
        pdf_download(DATA/'industry_programme.pdf','Download Industry Programme PDF')
        pdf_download(DATA/'academic_programme.pdf','Download Academic Conference PDF')
        pdf_download(DATA/'gala_dinner.pdf','Download Gala Dinner PDF')
        if not programme.empty:
            q=st.text_input('Search programme', key='prog_search')
            df=programme.copy()
            if q: df=df[df.astype(str).apply(lambda row: row.str.contains(q, case=False, na=False).any(), axis=1)]
            st.dataframe(df, use_container_width=True, hide_index=True)
    with tabs[3]:
        st.subheader('Academic Conference')
        pdf_download(DATA/'academic_programme.pdf','Download Academic Conference Timetable')
        if not abstracts.empty:
            room=st.selectbox('Filter venue', ['All']+sorted([str(x) for x in abstracts.get('Room',pd.Series(dtype=str)).dropna().unique()]))
            df=abstracts if room=='All' else abstracts[abstracts['Room'].astype(str)==room]
            st.dataframe(df[[c for c in ['Paper ID','Abstract No','Paper Title','Presenter Name','Theme','Session','Date','Time','Room','Moderator'] if c in df.columns]], use_container_width=True, hide_index=True)
    with tabs[4]:
        st.subheader('Industry Conference')
        pdf_download(DATA/'industry_programme.pdf','Download Industry Conference Timetable')
        if not industry.empty: st.dataframe(industry, use_container_width=True, hide_index=True)
    with tabs[5]:
        st.subheader('Abstract Book')
        if abstracts.empty: st.info('Abstracts will appear after admin upload.')
        else:
            q=st.text_input('Search by Paper ID / presenter / keyword / title')
            df=abstracts.copy()
            if q: df=df[df.astype(str).apply(lambda row: row.str.contains(q, case=False, na=False).any(), axis=1)]
            choices=[]
            for _,r in df.iterrows(): choices.append(f"{r.get('Abstract No','')} | {r.get('Paper ID','')} | {str(r.get('Paper Title',''))[:70]}")
            sel=st.selectbox('Select abstract', choices) if choices else None
            if sel:
                idx=choices.index(sel); r=df.iloc[idx]
                st.markdown(f"<div class='card'><h3 class='gold'>{r.get('Paper Title','')}</h3><p><b>Paper ID:</b> {r.get('Paper ID','')}<br><b>Presenter:</b> {r.get('Presenter Name','')}<br><b>Venue:</b> {r.get('Room','')} · {r.get('Time','')}</p><hr><p>{r.get('Abstract','')}</p><p><b>Keywords:</b> {r.get('Keywords','')}</p></div>", unsafe_allow_html=True)
    with tabs[6]:
        st.subheader('Gala Dinner')
        pdf_download(DATA/'gala_dinner.pdf','Download Gala Dinner Tentative')
        email=st.text_input('Registered email for dinner RSVP')
        if email:
            choice=st.radio('Dinner attendance', ['I will attend','Unable to attend'])
            if st.button('Submit Dinner RSVP', type='primary'):
                update_status(email, dinner=choice); st.success('Dinner RSVP updated.')
    with tabs[7]:
        st.subheader('Venue & Seating')
        img_tag(ASSETS/'seating_layout.png')
    with tabs[8]:
        st.subheader('Contact Secretariat')
        card('Registration Counter','<p>Please proceed to the registration counter for assistance, check-in, door gift collection, walk-in registration, or presenter verification.</p>')

else:
    if 'admin' not in st.session_state: st.session_state.admin=False
    if not st.session_state.admin:
        u=st.text_input('Username'); p=st.text_input('Password', type='password')
        if st.button('Login', type='primary'):
            if u==ADMIN_USER and p==ADMIN_PASS: st.session_state.admin=True; st.rerun()
            else: st.error('Invalid admin login.')
    else:
        st.success('Admin access granted.')
        tabs=st.tabs(['Upload Data','Participants','Check-In','Walk-In','Abstracts','Programme/PDF','Posters & Seating','Reports'])
        with tabs[0]:
            st.subheader('Upload One Master Excel')
            up=st.file_uploader('Upload NICHE2026_MASTER.xlsx', type=['xlsx','xls'])
            if up and st.button('Import Master Excel', type='primary'):
                p,a,i,pr=import_master(up); st.success(f'Imported: {len(p)} participants, {len(a)} abstracts, {len(i)} industry rows, {len(pr)} programme rows.')
            st.download_button('Download Current Master Excel Sample', MASTER_DEFAULT.read_bytes(), file_name='NICHE2026_MASTER.xlsx')
        with tabs[1]:
            st.subheader('Participants Database')
            st.dataframe(parts, use_container_width=True, hide_index=True)
        with tabs[2]:
            st.subheader('Check-In / Door Gift Counter')
            q=st.text_input('Search name/email/registration ID')
            df=parts.copy()
            if q: df=df[df.astype(str).apply(lambda row: row.str.contains(q, case=False, na=False).any(), axis=1)]
            st.dataframe(df[[c for c in ['Registration ID','Full Name','Email','Category','Institution'] if c in df.columns]], use_container_width=True, hide_index=True)
            email=st.text_input('Email to update')
            c1,c2,c3=st.columns(3)
            with c1:
                if st.button('Mark Checked In', type='primary') and email: update_status(email, checkin='Checked In'); st.success('Updated')
            with c2:
                if st.button('Mark Door Gift Collected') and email: update_status(email, doorgift='Collected'); st.success('Updated')
            with c3:
                table=st.text_input('Assign Table')
                if st.button('Save Table') and email: update_status(email, table_no=table); st.success('Updated')
        with tabs[3]:
            st.subheader('Walk-In Registration')
            st.info('Walk-in presenter is not allowed. Presenter record must be verified by secretariat.')
            with st.form('walkin'):
                name=st.text_input('Full Name'); email=st.text_input('Email'); org=st.text_input('Institution / Company')
                cat=st.selectbox('Category', ['Academic Participant','Industry Delegate','Invited Guest','Student Visitor','Sponsor','Media'])
                submit=st.form_submit_button('Register Walk-In')
            if submit and name and email:
                new={'Registration ID':f'NICHE-WALKIN-{len(parts)+1:03d}','Full Name':name,'Email':email,'Category':cat,'Institution':org,'Participant Type':'Walk-In'}
                parts2=pd.concat([parts,pd.DataFrame([new])], ignore_index=True); save_df('participants', parts2); update_status(email, checkin='Checked In'); st.success('Walk-in participant registered.')
        with tabs[4]:
            st.subheader('Abstract Book')
            st.dataframe(abstracts, use_container_width=True, hide_index=True)
        with tabs[5]:
            st.subheader('Programme / PDF')
            st.dataframe(programme, use_container_width=True, hide_index=True)
            pdf_download(DATA/'industry_programme.pdf','Download Industry Programme PDF')
            pdf_download(DATA/'academic_programme.pdf','Download Academic Programme PDF')
            pdf_download(DATA/'gala_dinner.pdf','Download Gala Dinner PDF')
        with tabs[6]:
            st.subheader('Posters & Seating')
            st.write('Main Poster'); img_tag(ASSETS/'main_poster.jpeg')
            st.write('Conference Poster'); img_tag(ASSETS/'conference_poster.jpeg')
            st.write('Seating Layout'); img_tag(ASSETS/'seating_layout.png')
        with tabs[7]:
            st.subheader('Reports')
            st.metric('Participants', len(parts)); st.metric('Abstracts', len(abstracts)); st.metric('Programme Events', len(programme))
            st.download_button('Export Participants CSV', parts.to_csv(index=False).encode('utf-8'), 'participants_export.csv')
            st.download_button('Export Abstracts CSV', abstracts.to_csv(index=False).encode('utf-8'), 'abstracts_export.csv')
