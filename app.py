"""
NICHE 2026 — Conference Registration & Check-In System
========================================================
A single-file Flask backend serving:
  • Public landing  + 3 tentative pages (Academic w/ abstracts, Industry, Gala Dinner)
  • Public self-registration by email
  • Admin dashboard: walk-in, check-in, door gift, dinner attendance, table assignment

Run:
    pip install flask pandas openpyxl
    python app.py
Then open  http://127.0.0.1:5000
Admin login   →   password: NICHE2026admin   (path: /admin)
"""
import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DB_PATH    = os.path.join(BASE_DIR, "niche.db")
EXCEL_PATH = os.path.join(BASE_DIR, "niche_data.xlsx")

# Tables reserved for participants at the Gala Dinner
PARTICIPANT_TABLES = [3, 4, 8, 9, 10, 27]
SEATS_PER_TABLE    = 10

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    """Create tables and seed once from the Excel master file."""
    conn = get_db()
    c    = conn.cursor()

    c.executescript("""
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
    """)
    conn.commit()

    # Seed only if empty
    c.execute("SELECT COUNT(*) AS n FROM participants")
    if c.fetchone()["n"] == 0 and os.path.exists(EXCEL_PATH):
        seed_from_excel(conn)

    conn.close()


def _clean(v):
    """Convert NaN / NaT / 'nan' to None so SQLite stores NULL."""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    if s.lower() in ("nan", "nat", ""):
        return None
    return s


def _yes(v):
    return 1 if str(v).strip().lower() in ("yes", "y", "true", "1") else 0


def seed_from_excel(conn):
    print("→ Seeding database from Excel master file...")
    xl = pd.ExcelFile(EXCEL_PATH)
    c  = conn.cursor()

    # Participants
    if "Participants" in xl.sheet_names:
        df = pd.read_excel(xl, "Participants")
        for _, r in df.iterrows():
            email = _clean(r.get("Email"))
            if not email:
                continue
            c.execute("""INSERT OR IGNORE INTO participants
                (email, full_name, phone, organisation, category,
                 academic, industry, conference_checkin, doorgift_collected,
                 attend_dinner, dinner_checkin, table_number, registration_source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (email,
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
                 int(r["Table_Number"]) if pd.notna(r.get("Table_Number")) else None,
                 _clean(r.get("Registration_Source")) or "Preloaded"))

    # Academic programme
    if "Academic_Programme" in xl.sheet_names:
        df = pd.read_excel(xl, "Academic_Programme")
        for i, r in df.iterrows():
            c.execute("""INSERT INTO academic_programme
                (venue, time, session, moderator, theme, paper_id, title,
                 presenter, email, sort_order) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (_clean(r.get("Venue")), _clean(r.get("Time")),
                 _clean(r.get("Session")), _clean(r.get("Moderator")),
                 _clean(r.get("Theme")), _clean(r.get("Paper_ID")),
                 _clean(r.get("Title")), _clean(r.get("Presenter")),
                 _clean(r.get("Email_From_Abstract")), i))

    # Abstracts (keyed by paper_id)
    if "Abstracts" in xl.sheet_names:
        df = pd.read_excel(xl, "Abstracts")
        for _, r in df.iterrows():
            pid = _clean(r.get("Paper_ID"))
            if not pid:
                continue
            c.execute("""INSERT OR REPLACE INTO abstracts
                (paper_id, title, presenter, email, venue, time, session,
                 keywords, abstract_text, authors) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (pid,
                 _clean(r.get("Title")), _clean(r.get("Presenter")),
                 _clean(r.get("Email")), _clean(r.get("Venue")),
                 _clean(r.get("Time")), _clean(r.get("Session")),
                 _clean(r.get("Keywords")), _clean(r.get("Abstract_Text")),
                 _clean(r.get("Authors_Affiliation"))))

    # Industry
    if "Industry_Programme" in xl.sheet_names:
        df = pd.read_excel(xl, "Industry_Programme")
        for i, r in df.iterrows():
            c.execute("""INSERT INTO industry_programme
                (day, time, venue, session, speaker, organisation, details, sort_order)
                VALUES (?,?,?,?,?,?,?,?)""",
                (_clean(r.get("Day")), _clean(r.get("Time")),
                 _clean(r.get("Venue")), _clean(r.get("Session")),
                 _clean(r.get("Speaker")), _clean(r.get("Organisation")),
                 _clean(r.get("Details")), i))

    # Dinner
    if "Gala_Dinner_Programme" in xl.sheet_names:
        df = pd.read_excel(xl, "Gala_Dinner_Programme")
        for i, r in df.iterrows():
            c.execute("""INSERT INTO dinner_programme (time, event, sort_order)
                         VALUES (?,?,?)""",
                (_clean(r.get("Time")), _clean(r.get("Event")), i))

    conn.commit()
    print("✓ Seed complete.")


def log_action(action, email=None, detail=None):
    conn = get_db()
    conn.execute("INSERT INTO audit_log (action, target_email, detail) VALUES (?,?,?)",
                 (action, email, detail))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
ADMIN_PASSWORD = "NICHE2026admin"   # from Settings sheet

def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login", next=request.path))
        return fn(*a, **kw)
    return wrapper


# ---------------------------------------------------------------------------
# Public routes
# ---------------------------------------------------------------------------
@app.route("/")
def home():
    conn = get_db()
    stats = {
        "participants": conn.execute("SELECT COUNT(*) c FROM participants").fetchone()["c"],
        "papers"      : conn.execute("SELECT COUNT(*) c FROM abstracts").fetchone()["c"],
        "speakers"    : conn.execute(
            "SELECT COUNT(*) c FROM industry_programme WHERE speaker IS NOT NULL"
        ).fetchone()["c"],
    }
    conn.close()
    return render_template("home.html", stats=stats)


@app.route("/academic")
def academic():
    conn = get_db()
    programme = conn.execute(
        "SELECT * FROM academic_programme ORDER BY sort_order"
    ).fetchall()
    abstracts = {a["paper_id"]: dict(a)
                 for a in conn.execute("SELECT * FROM abstracts").fetchall()}
    conn.close()

    # Group by Venue → Session
    grouped = {}
    for row in programme:
        v = row["venue"] or "—"
        grouped.setdefault(v, []).append(dict(row))
    return render_template("academic.html",
                           grouped=grouped, abstracts=abstracts)


@app.route("/industry")
def industry():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM industry_programme ORDER BY sort_order"
    ).fetchall()
    conn.close()
    grouped = {}
    for r in rows:
        grouped.setdefault(r["day"] or "—", []).append(dict(r))
    return render_template("industry.html", grouped=grouped)


@app.route("/dinner")
def dinner():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM dinner_programme ORDER BY sort_order"
    ).fetchall()
    conn.close()
    return render_template("dinner.html", programme=rows)


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        name  = request.form.get("full_name", "").strip()
        if not email or not name:
            flash("Email dan nama wajib diisi.", "error")
            return redirect(url_for("register"))

        conn = get_db()
        exists = conn.execute(
            "SELECT id, full_name FROM participants WHERE email = ?", (email,)
        ).fetchone()
        if exists:
            conn.close()
            flash(f"Email ini dah ada dalam sistem ({exists['full_name']}). Sila tunjuk ke admin di kaunter.", "info")
            return redirect(url_for("register"))

        conn.execute("""INSERT INTO participants
            (email, full_name, phone, organisation, category,
             attend_dinner, registration_source)
            VALUES (?,?,?,?,?,?,?)""",
            (email, name,
             request.form.get("phone", "").strip(),
             request.form.get("organisation", "").strip(),
             request.form.get("category", "Participant").strip(),
             1 if request.form.get("attend_dinner") == "yes" else 0,
             "Self-Register"))
        conn.commit()
        conn.close()
        log_action("self_register", email, name)
        flash("✓ Registration berjaya! Sila tunjuk email ke admin untuk ambil door gift.", "success")
        return redirect(url_for("register_success", email=email))

    return render_template("register.html")


@app.route("/register/success")
def register_success():
    email = request.args.get("email", "")
    return render_template("register_success.html", email=email)


# ---------------------------------------------------------------------------
# Admin routes
# ---------------------------------------------------------------------------
@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    if session.get("is_admin"):
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        pw = request.form.get("password", "")
        if pw == ADMIN_PASSWORD:
            session["is_admin"] = True
            log_action("admin_login")
            return redirect(request.args.get("next") or url_for("admin_dashboard"))
        flash("Password salah. Cuba lagi.", "error")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("home"))


@app.route("/admin/dashboard")
@admin_required
def admin_dashboard():
    conn = get_db()
    parts = conn.execute("""SELECT * FROM participants
                            ORDER BY created_at DESC, id DESC""").fetchall()

    stats = {
        "total"   : len(parts),
        "checked" : sum(1 for p in parts if p["conference_checkin"]),
        "gift"    : sum(1 for p in parts if p["doorgift_collected"]),
        "dinner"  : sum(1 for p in parts if p["attend_dinner"]),
        "dinner_in": sum(1 for p in parts if p["dinner_checkin"]),
    }

    # Table occupancy ( only participant tables )
    occupancy = {t: 0 for t in PARTICIPANT_TABLES}
    for p in parts:
        if p["table_number"] in occupancy:
            occupancy[p["table_number"]] += 1

    conn.close()
    return render_template("admin_dashboard.html",
                           parts=[dict(p) for p in parts],
                           stats=stats,
                           occupancy=occupancy,
                           participant_tables=PARTICIPANT_TABLES,
                           seats_per_table=SEATS_PER_TABLE)


@app.route("/admin/walkin", methods=["POST"])
@admin_required
def admin_walkin():
    email = request.form.get("email", "").strip().lower()
    name  = request.form.get("full_name", "").strip()
    if not email or not name:
        flash("Email dan nama wajib.", "error")
        return redirect(url_for("admin_dashboard"))

    conn = get_db()
    try:
        conn.execute("""INSERT INTO participants
            (email, full_name, phone, organisation, category,
             attend_dinner, conference_checkin, registration_source)
            VALUES (?,?,?,?,?,?,?,?)""",
            (email, name,
             request.form.get("phone", "").strip(),
             request.form.get("organisation", "").strip(),
             request.form.get("category", "Walk-In").strip(),
             1 if request.form.get("attend_dinner") == "yes" else 0,
             1,             # walk-in is auto checked-in
             "Walk-In"))
        conn.commit()
        log_action("walk_in", email, name)
        flash(f"✓ Walk-in {name} berjaya didaftar & check-in.", "success")
    except sqlite3.IntegrityError:
        flash(f"⚠ Email {email} sudah ada dalam sistem.", "error")
    finally:
        conn.close()
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/toggle/<int:pid>/<field>", methods=["POST"])
@admin_required
def admin_toggle(pid, field):
    """AJAX toggle for the boolean fields."""
    allowed = {"conference_checkin", "doorgift_collected",
               "attend_dinner",      "dinner_checkin"}
    if field not in allowed:
        abort(400)

    conn = get_db()
    row = conn.execute("SELECT * FROM participants WHERE id = ?", (pid,)).fetchone()
    if not row:
        conn.close()
        abort(404)

    new_val = 0 if row[field] else 1
    conn.execute(f"UPDATE participants SET {field} = ? WHERE id = ?", (new_val, pid))

    # If turning OFF attend_dinner → unassign table & clear dinner check-in
    if field == "attend_dinner" and new_val == 0:
        conn.execute("""UPDATE participants
                        SET table_number = NULL, dinner_checkin = 0
                        WHERE id = ?""", (pid,))
    conn.commit()
    conn.close()
    log_action(f"toggle:{field}={new_val}", row["email"])
    return jsonify(ok=True, new_value=new_val)


@app.route("/admin/assign_table/<int:pid>", methods=["POST"])
@admin_required
def admin_assign_table(pid):
    raw = request.form.get("table_number", "").strip()
    conn = get_db()

    if raw in ("", "0", "none", "None"):
        conn.execute("UPDATE participants SET table_number = NULL WHERE id = ?", (pid,))
        conn.commit()
        conn.close()
        log_action("table_unassign", detail=f"pid={pid}")
        return jsonify(ok=True, table=None)

    try:
        t = int(raw)
    except ValueError:
        conn.close()
        return jsonify(ok=False, error="Nombor meja tidak sah."), 400

    if t not in PARTICIPANT_TABLES:
        conn.close()
        return jsonify(ok=False,
                       error=f"Meja {t} bukan untuk participants. "
                             f"Pilih: {PARTICIPANT_TABLES}"), 400

    # capacity check
    n = conn.execute("""SELECT COUNT(*) c FROM participants
                        WHERE table_number = ? AND id != ?""",
                     (t, pid)).fetchone()["c"]
    if n >= SEATS_PER_TABLE:
        conn.close()
        return jsonify(ok=False,
                       error=f"Meja {t} sudah penuh ({n}/{SEATS_PER_TABLE})."), 400

    # ensure dinner attendance is ON when assigning a table
    conn.execute("""UPDATE participants
                    SET table_number = ?, attend_dinner = 1
                    WHERE id = ?""", (t, pid))
    conn.commit()
    conn.close()
    log_action("table_assign", detail=f"pid={pid} → table {t}")
    return jsonify(ok=True, table=t)


@app.route("/admin/delete/<int:pid>", methods=["POST"])
@admin_required
def admin_delete(pid):
    conn = get_db()
    row = conn.execute("SELECT email FROM participants WHERE id = ?", (pid,)).fetchone()
    conn.execute("DELETE FROM participants WHERE id = ?", (pid,))
    conn.commit()
    conn.close()
    if row:
        log_action("delete", row["email"])
    return jsonify(ok=True)


# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------
@app.context_processor
def inject_globals():
    return {
        "EVENT_NAME": "NICHE 2026",
        "EVENT_DATES": "9 – 10 June 2026",
        "EVENT_VENUE": "Royale Chulan Seremban, Negeri Sembilan",
        "is_admin": session.get("is_admin", False),
        "current_year": datetime.now().year,
    }


# ---------------------------------------------------------------------------
# Bootstrap
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    init_db()
    print("=" * 60)
    print("  NICHE 2026 — Registration System")
    print("=" * 60)
    print(f"  Public      : http://127.0.0.1:5000")
    print(f"  Admin login : http://127.0.0.1:5000/admin")
    print(f"  Password    : {ADMIN_PASSWORD}")
    print("=" * 60)
    app.run(debug=True, host="0.0.0.0", port=5000)
