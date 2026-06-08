# NICHE 2026 — Self Check-In System

Streamlit app untuk International Halal Conference 2026.

## Features

- **3 Tentative views** — Academic (with expandable abstracts), Industry, Gala Dinner
- **Public registration** by email
- **Self check-in** — peserta enter email, tick sendiri (Conference, Door Gift, Dinner, Dinner Check-In)
- **Admin** — walk-in registration, full participants editor, table assignment
- **Tables 3, 4, 8, 9, 10, 27** untuk participants (10 seats each, validated)
- **NO sidebar** — top tabs only
- **Royal Navy + Gold** theme matching event posters

## Local Development

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open <http://localhost:8501>

## Deploy to Streamlit Cloud

1. Push these files to a **public GitHub repo**:
   - `app.py`
   - `requirements.txt`
   - `niche_data.xlsx`
   - `.streamlit/config.toml`
   - `README.md`

2. Go to <https://share.streamlit.io>, click **New app**
3. Pick your repo, set:
   - **Main file path**: `app.py`
   - **Python version**: 3.11+
4. Click **Deploy**

## Admin Access

- Tab **🔐 Admin**
- Password: `NICHE2026admin` (edit `ADMIN_PASSWORD` in `app.py`)

## Data Persistence Notes

⚠ Streamlit Cloud's filesystem is **ephemeral** — `niche.db` will be wiped on
every restart / redeploy. For an active 2-day event, this is usually fine
because:
- The DB auto-seeds from `niche_data.xlsx` on each fresh start
- All changes are saved to `niche.db` during the session

For long-term persistence beyond a single event, swap SQLite for:
- Streamlit's built-in **st.connection** with a hosted Postgres (e.g. Supabase, Neon)
- Or push periodic backups of `niche.db` to S3/Google Drive

## File Structure

```
niche2026/
├── app.py                  # Main Streamlit app
├── requirements.txt        # streamlit, pandas, openpyxl
├── niche_data.xlsx         # Master data (seed source)
├── .streamlit/
│   └── config.toml         # Theme + hide sidebar
└── README.md
```
