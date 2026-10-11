"""
APPS - Army Personnel Psychological Support (Flask backend)

Cleaned-up version. Main fixes compared to the old app.py:
  * pick_officer() was defined twice (nested) and always returned None -> fixed
  * Passwords are hashed (old plain-text passwords still work once and are
    upgraded to a hash automatically at the next successful login)
  * Secret key / access codes are no longer hard-coded; debug is off by default
  * Survey input is validated, one check-in per person per day (re-submitting
    updates today's entry instead of adding duplicates)
  * "Going Quiet" page now really lists personnel who stopped checking in
  * Same-origin check on every POST (basic CSRF protection), login rate limit,
    basic security headers, escaped text in PDFs, safe Excel/CSV cells
  * Repeated admin summary code moved into one helper, role checks into one decorator
"""

import csv
import os
import secrets
import shutil
import sqlite3
import time
from datetime import date, datetime, timedelta
from functools import wraps
from io import BytesIO, StringIO
from urllib.parse import urlparse
from xml.sax.saxutils import escape as xml_escape

import joblib
import pandas as pd
from flask import (Flask, Response, abort, redirect, render_template, request,
                   session, url_for)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if os.environ.get("VERCEL"):
    TMP_DIR = "/tmp"
    DB_PATH = os.path.join(TMP_DIR, "apps.db")
    SECRET_FILE = os.path.join(TMP_DIR, ".secret_key")
    orig_db = os.path.join(BASE_DIR, "apps.db")
    if not os.path.exists(DB_PATH) and os.path.exists(orig_db):
        shutil.copyfile(orig_db, DB_PATH)
else:
    DB_PATH = os.path.join(BASE_DIR, "apps.db")
    SECRET_FILE = os.path.join(BASE_DIR, ".secret_key")
MODEL_PATH = os.path.join(BASE_DIR, "stress_model.pkl")


# ----------------------------------------------------------------------
# App setup
# ----------------------------------------------------------------------
def _load_secret_key():
    """Secret key from env var, else from a local file (created once)."""
    key = os.environ.get("APPS_SECRET_KEY")
    if key:
        return key
    if os.path.exists(SECRET_FILE):
        with open(SECRET_FILE, "r") as f:
            return f.read().strip()
    key = secrets.token_hex(32)
    with open(SECRET_FILE, "w") as f:
        f.write(key)
    return key


app = Flask(__name__)
app.secret_key = _load_secret_key()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
)

# Trained Random Forest (made by train_model.py)
model = joblib.load(MODEL_PATH)
FEATURE_COLUMNS = ["sleep_hours", "duty_hours", "heart_rate", "mood_score"]


# ----------------------------------------------------------------------
# Database
# ----------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def pick_officer(conn):
    """Officer with the fewest assigned personnel (None if no officers exist)."""
    row = conn.execute("""
        SELECT o.id FROM Users o
        WHERE o.role = 'officer'
        ORDER BY (SELECT COUNT(*) FROM Users p WHERE p.assigned_officer_id = o.id), o.id
        LIMIT 1
    """).fetchone()
    return row["id"] if row else None


def _seed_codes(c, table, prefix, count):
    """Create random one-time access codes the first time the table is empty."""
    existing = c.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
    if existing:
        return
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    print(f"\n[APPS] First run - new {prefix} access codes (shown only once):")
    for _ in range(count):
        code = f"{prefix}-" + "".join(secrets.choice(alphabet) for _ in range(6))
        c.execute(f"INSERT INTO {table} (code, is_used) VALUES (?, 0)", (code,))
        print("   ", code)


def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS Users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            unit TEXT,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'personnel'
        )
    """)
    user_cols = [r["name"] for r in c.execute("PRAGMA table_info(Users)").fetchall()]
    if "assigned_officer_id" not in user_cols:
        c.execute("ALTER TABLE Users ADD COLUMN assigned_officer_id INTEGER")
        # One-time backfill: spread existing personnel across existing officers
        for u in c.execute("SELECT id FROM Users WHERE role = 'personnel'").fetchall():
            oid = pick_officer(conn)
            if oid:
                c.execute("UPDATE Users SET assigned_officer_id = ? WHERE id = ?", (oid, u["id"]))

    c.execute("""
        CREATE TABLE IF NOT EXISTS DailyData (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            sleep_hours REAL,
            duty_hours REAL,
            heart_rate INTEGER,
            mood_score INTEGER,
            stress_level TEXT,
            recommendation TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS OfficerCodes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            is_used INTEGER NOT NULL DEFAULT 0
        )
    """)
    _seed_codes(c, "OfficerCodes", "OFC", 5)

    c.execute("""
        CREATE TABLE IF NOT EXISTS AdminCodes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            is_used INTEGER NOT NULL DEFAULT 0
        )
    """)
    _seed_codes(c, "AdminCodes", "ADM", 3)

    c.execute("""
        CREATE TABLE IF NOT EXISTS StressRequests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            message TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Pending',
            stage TEXT NOT NULL DEFAULT 'New',
            officer_id INTEGER,
            remedy TEXT,
            created_date TEXT NOT NULL,
            resolved_date TEXT,
            stage_updated_at TEXT
        )
    """)
    request_cols = [r["name"] for r in c.execute("PRAGMA table_info(StressRequests)").fetchall()]
    if "stage" not in request_cols:
        c.execute("ALTER TABLE StressRequests ADD COLUMN stage TEXT NOT NULL DEFAULT 'New'")
    if "stage_updated_at" not in request_cols:
        c.execute("ALTER TABLE StressRequests ADD COLUMN stage_updated_at TEXT")
    c.execute("UPDATE StressRequests SET stage = 'Resolved' WHERE status = 'Resolved' AND stage = 'New'")

    c.execute("""
        CREATE TABLE IF NOT EXISTS CounselingNotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            officer_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            notes TEXT,
            action_taken TEXT,
            follow_up_date TEXT
        )
    """)
    note_cols = [r["name"] for r in c.execute("PRAGMA table_info(CounselingNotes)").fetchall()]
    if "follow_up_date" not in note_cols:
        c.execute("ALTER TABLE CounselingNotes ADD COLUMN follow_up_date TEXT")

    c.execute("""
        CREATE TABLE IF NOT EXISTS ActivityLog (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor_id INTEGER NOT NULL,
            action TEXT NOT NULL,
            target_user_id INTEGER,
            details TEXT,
            created_at TEXT NOT NULL
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS PersonnelTags (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            tag TEXT NOT NULL COLLATE NOCASE,
            officer_id INTEGER NOT NULL,
            created_date TEXT NOT NULL,
            UNIQUE(user_id, tag)
        )
    """)

    conn.commit()
    conn.close()


# ----------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------
# Stress request pipeline: a request moves through these stages IN ORDER.
STAGE_FLOW = ["New", "Acknowledged", "In Progress", "Resolved", "Closed"]
STAGE_INFO = {
    "New":          {"label": "New",          "next_label": "Acknowledge Request"},
    "Acknowledged": {"label": "Acknowledged", "next_label": "Start Working On This"},
    "In Progress":  {"label": "In Progress",  "next_label": "Mark Resolved"},
    "Resolved":     {"label": "Resolved",     "next_label": "Close Case"},
    "Closed":       {"label": "Closed",       "next_label": None},
}
PRESET_TAGS = ["Needs monitoring", "Improving", "Escalated"]
QUIET_AFTER_DAYS = 3  # no check-in for this many days -> "going quiet"


def get_next_stage(current_stage):
    """The single next stage in the pipeline, or None at the end."""
    if current_stage not in STAGE_FLOW:
        return STAGE_FLOW[0]
    idx = STAGE_FLOW.index(current_stage)
    return STAGE_FLOW[idx + 1] if idx + 1 < len(STAGE_FLOW) else None


def log_activity(conn, actor_id, action, target_user_id=None, details=""):
    """Audit-trail entry. The caller still does conn.commit()."""
    conn.execute(
        "INSERT INTO ActivityLog (actor_id, action, target_user_id, details, created_at) "
        "VALUES (?, ?, ?, ?, ?)",
        (actor_id, action, target_user_id, details, datetime.now().isoformat(timespec="seconds"))
    )


def claim_or_check(conn, person_id, officer_id):
    """True if this officer may act on this person.
    - already theirs -> True
    - unassigned -> assign to this officer, True
    - belongs to another officer (or isn't personnel) -> False"""
    row = conn.execute(
        "SELECT assigned_officer_id FROM Users WHERE id = ? AND role = 'personnel'",
        (person_id,)
    ).fetchone()
    if not row:
        return False
    owner = row["assigned_officer_id"]
    if owner is None:
        conn.execute(
            "UPDATE Users SET assigned_officer_id = ? WHERE id = ? AND assigned_officer_id IS NULL",
            (officer_id, person_id)
        )
        log_activity(conn, officer_id, "Auto-claimed case", person_id)
        conn.commit()
        return True
    return owner == officer_id


def to_int(value):
    """int(value) or None if it isn't a whole number."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def parse_number(raw, low, high, as_int=False):
    """Convert form text to a number inside [low, high]; None if invalid."""
    try:
        num = float(str(raw).strip())
    except (TypeError, ValueError):
        return None
    if num != num or num < low or num > high:  # NaN or out of range
        return None
    if as_int:
        if num != int(num):
            return None
        return int(num)
    return round(num, 1)


def valid_iso_date(text):
    try:
        date.fromisoformat(text)
        return True
    except (TypeError, ValueError):
        return False


def safe_cell(value):
    """Stop spreadsheet formula injection (a name like =HYPERLINK(...))."""
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        return "'" + value
    return value


def role_required(*roles):
    """Redirect to login unless the user is logged in with one of the roles."""
    def decorator(view):
        @wraps(view)
        def wrapper(*args, **kwargs):
            if "user_id" not in session or session.get("role") not in roles:
                if "officer" in roles and len(roles) == 1:
                    return redirect(url_for("officer_login"))
                elif "admin" in roles and len(roles) == 1:
                    return redirect(url_for("admin_login"))
                return redirect(url_for("login"))
            return view(*args, **kwargs)
        return wrapper
    return decorator


def home_for_role(role):
    if role == "officer":
        return url_for("officer_dashboard")
    if role == "admin":
        return url_for("admin_dashboard")
    return url_for("dashboard")


# ----------------------------------------------------------------------
# Security: same-origin check on POST, headers, login rate limit
# ----------------------------------------------------------------------
@app.before_request
def block_cross_site_posts():
    """Basic CSRF protection: a POST must come from this same site."""
    if request.method != "POST":
        return
    source = request.headers.get("Origin") or request.headers.get("Referer")
    if source:
        if urlparse(source).netloc != request.host:
            abort(403)


@app.after_request
def add_security_headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    resp.headers.setdefault("Referrer-Policy", "same-origin")
    if "user_id" in session:
        resp.headers["Cache-Control"] = "no-store"
    return resp


_LOGIN_FAILS = {}      # key -> (fail_count, first_fail_time)
MAX_LOGIN_FAILS = 5
LOCK_SECONDS = 300


def _login_key(name):
    return f"{request.remote_addr}|{(name or '').lower()}"


def login_locked(key):
    entry = _LOGIN_FAILS.get(key)
    if not entry:
        return False
    count, first = entry
    if time.time() - first > LOCK_SECONDS:
        _LOGIN_FAILS.pop(key, None)
        return False
    return count >= MAX_LOGIN_FAILS


def record_login_fail(key):
    count, first = _LOGIN_FAILS.get(key, (0, time.time()))
    _LOGIN_FAILS[key] = (count + 1, first)


@app.context_processor
def inject_nav_counts():
    """Badge count for the officer's request pipeline link."""
    if session.get("role") == "officer":
        conn = get_db()
        try:
            n = conn.execute(
                """SELECT COUNT(*) FROM StressRequests
                   JOIN Users ON StressRequests.user_id = Users.id
                   WHERE StressRequests.stage != 'Closed'
                     AND (Users.assigned_officer_id = ? OR Users.assigned_officer_id IS NULL)""",
                (session["user_id"],)
            ).fetchone()[0]
        finally:
            conn.close()
        return {"pending_count": n}
    return {}


# ----------------------------------------------------------------------
# ML prediction
# ----------------------------------------------------------------------
RECOMMENDATIONS = {
    "Low": "You're within a healthy range today. Keep up the balance of rest and duty.",
    "Medium": "Some strain showing up today. Try to get extra rest tonight if possible.",
    "High": "Your readings suggest significant strain. Please reach out to your welfare officer or medical staff.",
}


def calculate_stress(sleep_hours, duty_hours, heart_rate, mood_score):
    """Predict Low / Medium / High with the trained Random Forest."""
    features = pd.DataFrame([[sleep_hours, duty_hours, heart_rate, mood_score]],
                            columns=FEATURE_COLUMNS)
    level = str(model.predict(features)[0])
    return level, RECOMMENDATIONS.get(level, RECOMMENDATIONS["Medium"])


# ----------------------------------------------------------------------
# Login / signup
# ----------------------------------------------------------------------
@app.route("/")
@app.route("/portal")
def home():
    if "user_id" in session:
        return redirect(home_for_role(session.get("role")))
    return redirect(url_for("login"))


def _name_taken(conn, name):
    return conn.execute(
        "SELECT 1 FROM Users WHERE LOWER(TRIM(name)) = LOWER(?)", (name,)
    ).fetchone() is not None


def _read_signup_form():
    name = request.form.get("name", "").strip()
    unit = request.form.get("unit", "").strip()
    password = request.form.get("password", "")
    error = None
    if not name or len(name) > 50:
        error = "Please enter a name (up to 50 characters)."
    elif len(unit) > 50:
        error = "Unit name is too long (max 50 characters)."
    elif len(password) < 6:
        error = "Password must be at least 6 characters."
    return name, unit, password, error


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if "user_id" in session:
        return redirect(home_for_role(session.get("role")))
    error = None
    role = request.args.get("role", "personnel").lower()
    if role not in ("personnel", "officer", "admin"):
        role = "personnel"

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        unit = request.form.get("unit", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "personnel").lower()
        if role not in ("personnel", "officer", "admin"):
            role = "personnel"

        if not name or len(name) > 50:
            error = "Please enter a name (up to 50 characters)."
        elif len(unit) > 50:
            error = "Unit/Department name is too long (max 50 characters)."
        elif len(password) < 6:
            error = "Password must be at least 6 characters."
        else:
            conn = get_db()
            if _name_taken(conn, name):
                error = "That name is already registered. Please log in or use a different name."
                conn.close()
            else:
                assigned_officer = pick_officer(conn) if role == "personnel" else None
                cur = conn.execute(
                    "INSERT INTO Users (name, unit, password, role, assigned_officer_id) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (name, unit, generate_password_hash(password), role, assigned_officer)
                )
                conn.commit()
                new_id = cur.lastrowid
                conn.close()

                # Automatically log in the user upon registration
                session.clear()
                session.permanent = True
                session["user_id"] = new_id
                session["user_name"] = name
                session["role"] = role
                return redirect(home_for_role(role))

    return render_template("signup.html", error=error, active_role=role)


@app.route("/officer_signup", methods=["GET", "POST"])
def officer_signup():
    if request.method == "POST":
        return signup()
    return redirect(url_for("signup", role="officer"))


@app.route("/admin_signup", methods=["GET", "POST"])
def admin_signup():
    if request.method == "POST":
        return signup()
    return redirect(url_for("signup", role="admin"))


def _password_ok(conn, user, entered):
    """Check a password. Old plain-text passwords still work and are upgraded."""
    stored = user["password"] or ""
    if stored.startswith(("pbkdf2:", "scrypt:")):
        return check_password_hash(stored, entered)
    if secrets.compare_digest(stored.encode(), entered.encode()):
        conn.execute("UPDATE Users SET password = ? WHERE id = ?",
                     (generate_password_hash(entered), user["id"]))
        conn.commit()
        return True
    return False


@app.route("/login", methods=["GET", "POST"])
@app.route("/personnel_login", methods=["GET", "POST"])
@app.route("/officer_login", methods=["GET", "POST"])
@app.route("/admin_login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(home_for_role(session.get("role")))

    error = None
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        password = request.form.get("password", "")
        key = _login_key(name)

        if login_locked(key):
            error = "Too many failed attempts. Please wait a few minutes and try again."
        elif not name or not password:
            error = "Please enter both your name and password."
        else:
            conn = get_db()
            candidates = conn.execute(
                "SELECT * FROM Users WHERE LOWER(TRIM(name)) = LOWER(?) ORDER BY id", (name,)
            ).fetchall()
            valid_users = [u for u in candidates if _password_ok(conn, u, password)]
            conn.close()

            if valid_users:
                _LOGIN_FAILS.pop(key, None)
                user = valid_users[0]
                session.clear()
                session.permanent = True
                session["user_id"] = user["id"]
                session["user_name"] = user["name"]
                session["role"] = user["role"]
                return redirect(home_for_role(user["role"]))
            else:
                record_login_fail(key)
                error = "Invalid name or password. Please try again."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ----------------------------------------------------------------------
# Personnel: dashboard, comparison, check-in, requests
# ----------------------------------------------------------------------
# key, label, suffix, higher_is_better, "about the same" tolerance, bar max
COMPARE_METRICS = [
    ("sleep", "Sleep",      "hrs", True,  0.5, 12),
    ("duty",  "Duty",       "hrs", False, 0.5, 24),
    ("hr",    "Heart rate", "bpm", False, 3,   150),
    ("mood",  "Mood",       "/5",  True,  0.5, 5),
]
COMPARE_DAYS = 7          # look-back window
COMPARE_MIN_MEMBERS = 3   # privacy: never show a unit average built from fewer people


def get_unit_comparison(conn, user_id, unit):
    """Compare a person's last-7-day averages with their unit's.
    Each person is averaged first so someone who checks in a lot doesn't
    outweigh the others. Only aggregates leave this function."""
    if not unit:
        return {"status": "no_unit"}

    since = (date.today() - timedelta(days=COMPARE_DAYS - 1)).isoformat()
    rows = conn.execute("""
        SELECT DailyData.user_id AS uid,
               AVG(DailyData.sleep_hours) AS sleep,
               AVG(DailyData.duty_hours)  AS duty,
               AVG(DailyData.heart_rate)  AS hr,
               AVG(DailyData.mood_score)  AS mood
        FROM DailyData
        JOIN Users ON DailyData.user_id = Users.id
        WHERE Users.role = 'personnel' AND Users.unit = ? AND DailyData.date >= ?
        GROUP BY DailyData.user_id
    """, (unit, since)).fetchall()

    mine = next((r for r in rows if r["uid"] == user_id), None)
    if mine is None:
        return {"status": "no_checkins", "unit": unit}
    if len(rows) < COMPARE_MIN_MEMBERS:
        return {"status": "not_enough", "unit": unit, "min_members": COMPARE_MIN_MEMBERS}

    metrics = []
    watch_count = 0
    for key, label, suffix, higher_better, tol, bar_max in COMPARE_METRICS:
        you = mine[key]
        unit_avg = sum(r[key] for r in rows) / len(rows)
        diff = you - unit_avg

        if abs(diff) <= tol:
            verdict = "same"
        elif (diff > 0) == higher_better:
            verdict = "good"
        else:
            verdict = "watch"
            watch_count += 1

        metrics.append({
            "label": label, "suffix": suffix,
            "you": round(you, 1), "unit": round(unit_avg, 1), "diff": round(diff, 1),
            "verdict": verdict,
            "you_pct": min(100, round(you / bar_max * 100)),
            "unit_pct": min(100, round(unit_avg / bar_max * 100)),
        })

    if watch_count == 0:
        summary = "You're in line with, or doing better than, your unit this week."
    elif watch_count == 1:
        summary = "One indicator is worth a look this week - see the amber row below."
    else:
        summary = f"{watch_count} indicators are worth a look this week - see the amber rows below."

    return {
        "status": "ok", "unit": unit, "members": len(rows), "days": COMPARE_DAYS,
        "metrics": metrics, "summary": summary, "watch_count": watch_count,
    }


def build_tips(latest):
    """Up to 3 quick tips based on the latest check-in."""
    tips = []
    sleep, duty = latest["sleep_hours"], latest["duty_hours"]
    hr, mood = latest["heart_rate"], latest["mood_score"]

    if mood <= 2:
        tips.append("You're feeling low today. Consider using 'Talk to Welfare Officer' to share what's on your mind.")
    elif mood == 3:
        tips.append("Mood is so-so. A short walk or a chat with a colleague can lift it.")

    if sleep < 5:
        tips.append("You slept under 5 hours. Try to get to bed 45-60 minutes earlier tonight.")
    elif sleep < 6.5:
        tips.append("Your sleep was a little short. Aim for 7+ hours tonight and keep screens away for 30 minutes before bed.")

    if duty > 12:
        tips.append("Very long duty day. Take a proper meal and a short rest break, and rest well before your next shift.")
    elif duty > 9:
        tips.append("Heavy duty day. Short breaks between tasks help keep fatigue down.")

    if hr > 90:
        tips.append("Your resting heart rate is high. Try 5 minutes of slow breathing (4 seconds in, 6 out) and drink some water.")
    elif hr > 80:
        tips.append("Resting heart rate is slightly raised. A short walk or some stretching can help.")

    if not tips:
        tips.append("All your indicators look healthy today. Keep up the balance of rest and duty.")
    return tips[:3]


@app.route("/dashboard")
@role_required("personnel")
def dashboard():
    uid = session["user_id"]
    conn = get_db()
    history = conn.execute(
        "SELECT * FROM DailyData WHERE user_id = ? ORDER BY date DESC, id DESC", (uid,)
    ).fetchall()

    me = conn.execute("SELECT unit FROM Users WHERE id = ?", (uid,)).fetchone()
    compare = get_unit_comparison(conn, uid, me["unit"] if me else None)

    pending_request = conn.execute(
        "SELECT * FROM StressRequests WHERE user_id = ? AND stage NOT IN ('Resolved', 'Closed') "
        "ORDER BY id DESC LIMIT 1", (uid,)
    ).fetchone()
    resolved_request = None
    if not pending_request:
        resolved_request = conn.execute(
            "SELECT * FROM StressRequests WHERE user_id = ? AND stage IN ('Resolved', 'Closed') "
            "ORDER BY id DESC LIMIT 1", (uid,)
        ).fetchone()
    conn.close()

    latest = history[0] if history else None
    tips = build_tips(latest) if latest else []

    # Consecutive-day check-in streak
    streak = 0
    dates = {row["date"] for row in history}
    expected = date.today()
    while expected.isoformat() in dates:
        streak += 1
        expected -= timedelta(days=1)
    milestone = streak if streak in (7, 14, 30, 60, 100) else None

    today_str = date.today().isoformat()
    checked_in_today = today_str in dates

    hour = datetime.now().hour
    greeting = "Good morning" if hour < 12 else "Good afternoon" if hour < 17 else "Good evening"

    # Last 7 check-ins, oldest -> newest, for the personal trend chart
    recent = list(history[:7])
    recent.reverse()

    return render_template(
        "dashboard.html",
        user_name=session["user_name"],
        history=history,
        latest=latest,
        streak=streak,
        milestone=milestone,
        checked_in_today=checked_in_today,
        pending_request=pending_request,
        resolved_request=resolved_request,
        tips=tips,
        greeting=greeting,
        today=today_str,
        trend_labels=[r["date"] for r in recent],
        trend_sleep=[r["sleep_hours"] for r in recent],
        trend_duty=[r["duty_hours"] for r in recent],
        compare=compare,
    )


@app.route("/snapshot")
@role_required("personnel")
def snapshot():
    uid = session["user_id"]
    conn = get_db()
    history = conn.execute(
        "SELECT * FROM DailyData WHERE user_id = ? ORDER BY date DESC, id DESC", (uid,)
    ).fetchall()
    conn.close()

    latest = history[0] if history else None
    tips = build_tips(latest) if latest else []

    # Consecutive-day check-in streak
    streak = 0
    dates = {row["date"] for row in history}
    expected = date.today()
    while expected.isoformat() in dates:
        streak += 1
        expected -= timedelta(days=1)

    recent = list(history[:7])
    recent.reverse()

    return render_template(
        "snapshot.html",
        user_name=session["user_name"],
        history=history,
        latest=latest,
        streak=streak,
        tips=tips,
        trend_labels=[r["date"] for r in recent],
        trend_sleep=[r["sleep_hours"] for r in recent],
        trend_duty=[r["duty_hours"] for r in recent],
    )


@app.route("/results")
@role_required("personnel")
def results():
    conn = get_db()
    history = conn.execute(
        "SELECT * FROM DailyData WHERE user_id = ? ORDER BY date DESC, id DESC",
        (session["user_id"],)
    ).fetchall()
    conn.close()
    return render_template("results.html", user_name=session["user_name"], history=history)


@app.route("/submit_survey", methods=["POST"])
@role_required("personnel")
def submit_survey():
    sleep_hours = parse_number(request.form.get("sleep_hours"), 0, 12)
    duty_hours = parse_number(request.form.get("duty_hours"), 0, 24)
    heart_rate = parse_number(request.form.get("heart_rate"), 40, 150, as_int=True)
    mood_score = parse_number(request.form.get("mood_score"), 1, 5, as_int=True)

    # Bad / missing values: go back without saving anything
    if None in (sleep_hours, duty_hours, heart_rate, mood_score):
        return redirect(url_for("dashboard"))

    level, recommendation = calculate_stress(sleep_hours, duty_hours, heart_rate, mood_score)
    today = date.today().isoformat()

    conn = get_db()
    existing = conn.execute(
        "SELECT id FROM DailyData WHERE user_id = ? AND date = ? ORDER BY id DESC LIMIT 1",
        (session["user_id"], today)
    ).fetchone()
    if existing:  # one entry per day: update today's check-in
        conn.execute(
            """UPDATE DailyData SET sleep_hours = ?, duty_hours = ?, heart_rate = ?,
                      mood_score = ?, stress_level = ?, recommendation = ?
               WHERE id = ?""",
            (sleep_hours, duty_hours, heart_rate, mood_score, level, recommendation, existing["id"])
        )
    else:
        conn.execute(
            """INSERT INTO DailyData (user_id, date, sleep_hours, duty_hours, heart_rate,
                                      mood_score, stress_level, recommendation)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (session["user_id"], today, sleep_hours, duty_hours, heart_rate,
             mood_score, level, recommendation)
        )
    conn.commit()
    conn.close()
    return redirect(url_for("results", just_submitted=1))


@app.route("/talk_to_officer")
@role_required("personnel")
def talk_to_officer():
    """Separate 'Talk to Welfare Officer' page (sidebar menu)."""
    uid = session["user_id"]
    conn = get_db()
    requests_list = conn.execute(
        "SELECT * FROM StressRequests WHERE user_id = ? ORDER BY id DESC", (uid,)
    ).fetchall()
    conn.close()
    pending_request = next((r for r in requests_list if r["stage"] not in ("Resolved", "Closed")), None)
    resolved_request = None
    if not pending_request:
        resolved_request = next((r for r in requests_list if r["stage"] in ("Resolved", "Closed")), None)
    return render_template("talk_to_officer.html", requests=requests_list,
   pending_request=pending_request, resolved_request=resolved_request) 


@app.route("/submit_stress_request", methods=["POST"])
@role_required("personnel")
def submit_stress_request():
    message = request.form.get("message", "").strip()[:1000]
    back = "talk_to_officer" if request.form.get("from") == "talk" else "dashboard"
    if not message:
        return redirect(url_for(back))

    conn = get_db()
    # Only one active request at a time (anything short of Closed is still active)
    existing = conn.execute(
        "SELECT id FROM StressRequests WHERE user_id = ? AND stage != 'Closed'",
        (session["user_id"],)
    ).fetchone()
    if not existing:
        conn.execute(
            """INSERT INTO StressRequests (user_id, message, status, stage, created_date, stage_updated_at)
               VALUES (?, ?, 'Pending', 'New', ?, ?)""",
            (session["user_id"], message, date.today().isoformat(), datetime.now().isoformat())
        )
        conn.commit()
    conn.close()
    return redirect(url_for(back))

# ----------------------------------------------------------------------
# Officer: dashboard, quiet list, activity log, personnel detail
# ----------------------------------------------------------------------
@app.route("/officer_dashboard")
@role_required("officer")
def officer_dashboard():
    me = session["user_id"]
    conn = get_db()

    # Every personnel's most recent check-in (my caseload + unassigned)
    personnel = conn.execute("""
        SELECT Users.id, Users.name, Users.unit, DailyData.date, DailyData.sleep_hours,
               DailyData.duty_hours, DailyData.stress_level, Users.assigned_officer_id
        FROM DailyData
        JOIN Users ON DailyData.user_id = Users.id
        WHERE DailyData.id IN (SELECT MAX(id) FROM DailyData GROUP BY user_id)
          AND Users.role = 'personnel'
          AND (Users.assigned_officer_id = ? OR Users.assigned_officer_id IS NULL)
        ORDER BY DailyData.stress_level = 'High' DESC, DailyData.stress_level = 'Medium' DESC,
                 Users.name
    """, (me,)).fetchall()

    tag_rows = conn.execute("SELECT user_id, tag FROM PersonnelTags ORDER BY id").fetchall()
    tags_by_user = {}
    for r in tag_rows:
        tags_by_user.setdefault(r["user_id"], []).append(r["tag"])
    all_tags = sorted({r["tag"] for r in tag_rows}, key=str.lower)

    # Engagement scoring: who checks in regularly vs has gone quiet
    freq_rows = conn.execute("""
        SELECT user_id, COUNT(*) AS cnt FROM DailyData WHERE date >= ? GROUP BY user_id
    """, ((date.today() - timedelta(days=13)).isoformat(),)).fetchall()
    freq_by_user = {r["user_id"]: r["cnt"] for r in freq_rows}
    conn.close()

    scored = []
    for row in personnel:
        p = dict(row)
        days_since = (date.today() - date.fromisoformat(p["date"])).days
        recency_score = max(0, 100 - days_since * 20)
        frequency_score = min(100, round(freq_by_user.get(p["id"], 0) / 14 * 100))
        score = round((recency_score + frequency_score) / 2)
        p["engagement_score"] = score
        p["engagement_label"] = "Engaged" if score >= 70 else "Fading" if score >= 40 else "Quiet"
        p["days_since_checkin"] = days_since
        scored.append(p)
    personnel = scored

    total = len(personnel)
    high_count = sum(1 for p in personnel if p["stress_level"] == "High")
    medium_count = sum(1 for p in personnel if p["stress_level"] == "Medium")
    low_count = sum(1 for p in personnel if p["stress_level"] == "Low")

    return render_template(
        "officer_dashboard.html",
        user_name=session["user_name"],
        personnel=personnel,
        tags_by_user=tags_by_user,
        all_tags=all_tags,
        total=total,
        high_count=high_count,
        medium_count=medium_count,
        low_count=low_count,
        today=date.today().isoformat(),
        quiet_list=[p for p in personnel if p["engagement_label"] == "Quiet"],
    )


@app.route("/officer/personnel")
@role_required("officer")
def officer_personnel_list():
    """Separate 'Personnel List' page (sidebar menu)."""
    me = session["user_id"]
    conn = get_db()
    rows = conn.execute("""
        SELECT Users.id, Users.name, Users.unit, DailyData.date, DailyData.sleep_hours,
               DailyData.duty_hours, DailyData.stress_level, Users.assigned_officer_id
        FROM DailyData
        JOIN Users ON DailyData.user_id = Users.id
        WHERE DailyData.id IN (SELECT MAX(id) FROM DailyData GROUP BY user_id)
          AND Users.role = 'personnel'
          AND (Users.assigned_officer_id = ? OR Users.assigned_officer_id IS NULL)
        ORDER BY DailyData.stress_level = 'High' DESC, DailyData.stress_level = 'Medium' DESC,
                 Users.name
    """, (me,)).fetchall()

    tag_rows = conn.execute("SELECT user_id, tag FROM PersonnelTags ORDER BY id").fetchall()
    tags_by_user = {}
    for r in tag_rows:
        tags_by_user.setdefault(r["user_id"], []).append(r["tag"])
    all_tags = sorted({r["tag"] for r in tag_rows}, key=str.lower)

    freq_rows = conn.execute(
        "SELECT user_id, COUNT(*) AS cnt FROM DailyData WHERE date >= ? GROUP BY user_id",
        ((date.today() - timedelta(days=13)).isoformat(),)).fetchall()
    freq_by_user = {r["user_id"]: r["cnt"] for r in freq_rows}
    conn.close()

    personnel = []
    for row in rows:
        p = dict(row)
        days_since = (date.today() - date.fromisoformat(p["date"])).days
        score = round((max(0, 100 - days_since * 20)
                       + min(100, round(freq_by_user.get(p["id"], 0) / 14 * 100))) / 2)
        p["engagement_label"] = "Engaged" if score >= 70 else "Fading" if score >= 40 else "Quiet"
        p["days_since_checkin"] = days_since
        personnel.append(p)

    return render_template("officer_personnel.html", personnel=personnel,
                           tags_by_user=tags_by_user, all_tags=all_tags,
                           total=len(personnel))


@app.route("/officer/going_quiet")
@role_required("officer")
def officer_going_quiet():
    """Personnel in my caseload who haven't checked in for a few days (or never)."""
    conn = get_db()
    rows = conn.execute("""
        SELECT Users.id, Users.name, MAX(DailyData.date) AS last_date
        FROM Users
        LEFT JOIN DailyData ON DailyData.user_id = Users.id
        WHERE Users.role = 'personnel'
          AND (Users.assigned_officer_id = ? OR Users.assigned_officer_id IS NULL)
        GROUP BY Users.id
    """, (session["user_id"],)).fetchall()
    conn.close()

    followups = []
    for r in rows:
        if r["last_date"] is None:
            days = None  # never checked in
        else:
            days = (date.today() - date.fromisoformat(r["last_date"])).days
            if days < QUIET_AFTER_DAYS:
                continue
        followups.append({"id": r["id"], "name": r["name"], "days": days})

    # Never-checked-in first, then the longest silence first
    followups.sort(key=lambda f: (f["days"] is not None, -(f["days"] or 0)))
    return render_template("officer_going_quiet.html", followups=followups)


@app.route("/officer/activity_log")
@role_required("officer")
def officer_activity_log():
    conn = get_db()
    entries = conn.execute("""
        SELECT ActivityLog.*, Users.name AS target_name
        FROM ActivityLog
        LEFT JOIN Users ON ActivityLog.target_user_id = Users.id
        WHERE ActivityLog.actor_id = ?
        ORDER BY ActivityLog.id DESC
        LIMIT 100
    """, (session["user_id"],)).fetchall()
    conn.close()
    return render_template("officer_activity_log.html",
                           officer_name=session["user_name"], entries=entries)


@app.route("/officer/personnel/<int:user_id>")
@role_required("officer")
def personnel_detail(user_id):
    conn = get_db()
    person = conn.execute("SELECT * FROM Users WHERE id = ?", (user_id,)).fetchone()
    if (not person or person["role"] != "personnel"
            or person["assigned_officer_id"] not in (None, session["user_id"])):
        conn.close()
        return redirect(url_for("officer_dashboard"))

    history = conn.execute(
        "SELECT * FROM DailyData WHERE user_id = ? ORDER BY date DESC, id DESC", (user_id,)
    ).fetchall()
    notes = conn.execute(
        """SELECT CounselingNotes.*, Users.name AS officer_name
           FROM CounselingNotes
           JOIN Users ON CounselingNotes.officer_id = Users.id
           WHERE CounselingNotes.user_id = ?
           ORDER BY CounselingNotes.id DESC""", (user_id,)
    ).fetchall()
    requests_rows = conn.execute(
        """SELECT StressRequests.*, Users.name AS officer_name
           FROM StressRequests
           LEFT JOIN Users ON StressRequests.officer_id = Users.id
           WHERE StressRequests.user_id = ?""", (user_id,)
    ).fetchall()
    tags = conn.execute(
        "SELECT id, tag FROM PersonnelTags WHERE user_id = ? ORDER BY id", (user_id,)
    ).fetchall()
    db_tags = [r["tag"] for r in conn.execute(
        "SELECT DISTINCT tag FROM PersonnelTags ORDER BY tag").fetchall()]
    conn.close()

    preset_lower = [t.lower() for t in PRESET_TAGS]
    tag_suggestions = PRESET_TAGS + [t for t in db_tags if t.lower() not in preset_lower]

    # Unified timeline (newest first). "order" sequences events within one day.
    timeline = []
    for e in history:
        timeline.append({"kind": "checkin", "date": e["date"][:10], "order": 0,
                         "id": e["id"], "data": e})
    for r in requests_rows:
        timeline.append({"kind": "request_raised", "date": r["created_date"][:10],
                         "order": 1, "id": r["id"], "data": r})
        if r["resolved_date"]:
            timeline.append({"kind": "request_resolved", "date": r["resolved_date"][:10],
                             "order": 3, "id": r["id"], "data": r})
    for n in notes:
        timeline.append({"kind": "note", "date": n["date"][:10], "order": 2,
                         "id": n["id"], "data": n})
    timeline.sort(key=lambda x: (x["date"], x["order"], x["id"]), reverse=True)

    # Last 10 check-ins, oldest -> newest, for the trend chart
    recent = list(history[:10])
    recent.reverse()

    return render_template(
        "personnel_detail.html",
        officer_name=session["user_name"],
        person=person,
        history=history,
        notes=notes,
        timeline=timeline,
        tags=tags,
        tag_suggestions=tag_suggestions,
        trend_labels=[r["date"] for r in recent],
        trend_sleep=[r["sleep_hours"] for r in recent],
        trend_duty=[r["duty_hours"] for r in recent],
    )


@app.route("/add_note", methods=["POST"])
@role_required("officer")
def add_note():
    target = to_int(request.form.get("user_id"))
    note_text = request.form.get("notes", "").strip()[:2000]
    action_taken = request.form.get("action_taken", "").strip()[:500]
    follow_up_date = request.form.get("follow_up_date") or None
    if follow_up_date and not valid_iso_date(follow_up_date):
        follow_up_date = None

    conn = get_db()
    if target is None or not claim_or_check(conn, target, session["user_id"]):
        conn.close()
        return redirect(url_for("officer_dashboard"))
    if not note_text:
        conn.close()
        return redirect(url_for("personnel_detail", user_id=target))

    conn.execute(
        """INSERT INTO CounselingNotes (user_id, officer_id, date, notes, action_taken, follow_up_date)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (target, session["user_id"], date.today().isoformat(), note_text, action_taken, follow_up_date)
    )
    log_activity(conn, session["user_id"], "Added counseling note", target,
                 f"Follow-up set: {follow_up_date}" if follow_up_date else "")
    conn.commit()
    conn.close()
    return redirect(url_for("personnel_detail", user_id=target, saved=1))


@app.route("/add_tag", methods=["POST"])
@role_required("officer")
def add_tag():
    target = to_int(request.form.get("user_id"))
    tag = request.form.get("tag", "").strip()[:30]
    if target is None:
        return redirect(url_for("officer_dashboard"))

    if tag:
        conn = get_db()
        try:
            if not claim_or_check(conn, target, session["user_id"]):
                return redirect(url_for("officer_dashboard"))
            try:
                conn.execute(
                    "INSERT INTO PersonnelTags (user_id, tag, officer_id, created_date) VALUES (?, ?, ?, ?)",
                    (target, tag, session["user_id"], date.today().isoformat())
                )
                log_activity(conn, session["user_id"], f"Added tag '{tag}'", target)
                conn.commit()
            except sqlite3.IntegrityError:
                pass  # tag already on this person
        finally:
            conn.close()
    return redirect(url_for("personnel_detail", user_id=target))


@app.route("/remove_tag/<int:tag_id>", methods=["POST"])
@role_required("officer")
def remove_tag(tag_id):
    conn = get_db()
    row = conn.execute("SELECT user_id, tag FROM PersonnelTags WHERE id = ?", (tag_id,)).fetchone()
    if not row:
        conn.close()
        return redirect(url_for("officer_dashboard"))

    if not claim_or_check(conn, row["user_id"], session["user_id"]):
        conn.close()
        return redirect(url_for("officer_dashboard"))

    conn.execute("DELETE FROM PersonnelTags WHERE id = ?", (tag_id,))
    log_activity(conn, session["user_id"], f"Removed tag '{row['tag']}'", row["user_id"])
    conn.commit()
    conn.close()
    return redirect(url_for("personnel_detail", user_id=row["user_id"]))


@app.route("/release_personnel", methods=["POST"])
@role_required("officer")
def release_personnel():
    target = to_int(request.form.get("user_id"))
    if target is not None:
        conn = get_db()
        cur = conn.execute(
            "UPDATE Users SET assigned_officer_id = NULL WHERE id = ? AND assigned_officer_id = ?",
            (target, session["user_id"])
        )
        if cur.rowcount:  # only log if this officer actually owned the case
            log_activity(conn, session["user_id"], "Released case", target)
        conn.commit()
        conn.close()
    return redirect(url_for("officer_dashboard"))


# ----------------------------------------------------------------------
# Officer: stress-request pipeline
# ----------------------------------------------------------------------
@app.route("/officer/requests")
@role_required("officer")
def view_stress_requests():
    conn = get_db()
    active_requests = conn.execute(
        """SELECT StressRequests.*, Users.name AS personnel_name
           FROM StressRequests
           JOIN Users ON StressRequests.user_id = Users.id
           WHERE StressRequests.stage != 'Closed'
             AND (Users.assigned_officer_id = ? OR Users.assigned_officer_id IS NULL)
           ORDER BY StressRequests.id DESC""",
        (session["user_id"],)
    ).fetchall()
    conn.close()

    pending_requests = []
    for req in active_requests:
        row = dict(req)
        info = STAGE_INFO.get(row["stage"], STAGE_INFO["New"])
        row["current_label"] = info["label"]
        row["next_stage"] = get_next_stage(row["stage"])
        row["next_label"] = info["next_label"]
        pending_requests.append(row)

    return render_template(
        "officer_requests.html",
        officer_name=session["user_name"],
        pending_requests=pending_requests,
        stage_flow=STAGE_FLOW,
    )


@app.route("/officer/requests/<int:request_id>/advance", methods=["POST"])
@role_required("officer")
def advance_stress_request(request_id):
    conn = get_db()
    req = conn.execute("SELECT * FROM StressRequests WHERE id = ?", (request_id,)).fetchone()
    if not req or not claim_or_check(conn, req["user_id"], session["user_id"]):
        conn.close()
        return redirect(url_for("view_stress_requests"))

    current_stage = req["stage"] or "New"
    expected_next = get_next_stage(current_stage)

    # Guard rail: only move to the exact next stage
    if not expected_next or request.form.get("target_stage") != expected_next:
        conn.close()
        return redirect(url_for("view_stress_requests"))

    remedy = req["remedy"]
    resolved_date = req["resolved_date"]
    if expected_next == "Resolved":
        remedy = request.form.get("remedy", "").strip()[:2000]
        if not remedy:
            conn.close()
            return redirect(url_for("view_stress_requests"))
        resolved_date = date.today().isoformat()

    status = "Resolved" if expected_next in ("Resolved", "Closed") else "Pending"
    conn.execute(
        """UPDATE StressRequests
           SET stage = ?, stage_updated_at = ?, officer_id = ?, remedy = ?,
               resolved_date = ?, status = ?
           WHERE id = ?""",
        (expected_next, datetime.now().isoformat(), session["user_id"],
         remedy, resolved_date, status, request_id)
    )
    log_activity(conn, session["user_id"], f"Moved request to '{expected_next}'", req["user_id"])
    conn.commit()
    conn.close()
    return redirect(url_for("view_stress_requests"))


# ----------------------------------------------------------------------
# Admin: overview, units
# ----------------------------------------------------------------------
def get_unit_summary(conn):
    """Latest check-in per personnel, grouped by unit.
    Returns (unit_summary dict, list of latest stress levels)."""
    rows = conn.execute("""
        SELECT Users.unit, DailyData.stress_level
        FROM DailyData
        JOIN Users ON DailyData.user_id = Users.id
        WHERE DailyData.id IN (SELECT MAX(id) FROM DailyData GROUP BY user_id)
          AND Users.role = 'personnel'
    """).fetchall()

    unit_summary = {}
    levels = []
    for row in rows:
        unit = row["unit"] or "Unassigned"
        level = row["stress_level"]
        levels.append(level)
        bucket = unit_summary.setdefault(unit, {"total": 0, "High": 0, "Medium": 0, "Low": 0})
        bucket["total"] += 1
        if level in bucket:
            bucket[level] += 1
    return unit_summary, levels


def count_levels(levels):
    return (len(levels), levels.count("High"), levels.count("Medium"), levels.count("Low"))


@app.route("/admin_dashboard")
@role_required("admin")
def admin_dashboard():
    conn = get_db()
    unit_summary, levels = get_unit_summary(conn)
    conn.close()
    total, high, medium, low = count_levels(levels)
    return render_template(
        "admin_dashboard.html",
        admin_name=session["user_name"],
        unit_summary=unit_summary,
        total_personnel=total, total_high=high, total_medium=medium, total_low=low,
    )


@app.route("/admin/generate_code", methods=["POST"])
@role_required("admin")
def admin_generate_code():
    code_type = request.form.get("type", "officer")
    table = "OfficerCodes" if code_type == "officer" else "AdminCodes"
    prefix = "OFC" if code_type == "officer" else "ADM"
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    new_code = f"{prefix}-" + "".join(secrets.choice(alphabet) for _ in range(6))
    conn = get_db()
    conn.execute(f"INSERT INTO {table} (code, is_used) VALUES (?, 0)", (new_code,))
    conn.commit()
    conn.close()
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/units")
@role_required("admin")
def admin_units():
    conn = get_db()
    unit_summary, _ = get_unit_summary(conn)
    conn.close()
    return render_template("admin_units.html", unit_summary=unit_summary)


@app.route("/admin/unit/<path:unit_name>")
@role_required("admin")
def admin_unit_detail(unit_name):
    conn = get_db()
    unit_filter = "(Users.unit IS NULL OR Users.unit = '')" if unit_name == "Unassigned" \
        else "Users.unit = ?"
    params = () if unit_name == "Unassigned" else (unit_name,)
    personnel = conn.execute(f"""
        SELECT Users.id, Users.name, Users.unit, DailyData.date, DailyData.sleep_hours,
               DailyData.duty_hours, DailyData.heart_rate, DailyData.mood_score,
               DailyData.stress_level
        FROM DailyData
        JOIN Users ON DailyData.user_id = Users.id
        WHERE DailyData.id IN (SELECT MAX(id) FROM DailyData GROUP BY user_id)
          AND Users.role = 'personnel'
          AND {unit_filter}
        ORDER BY DailyData.stress_level = 'High' DESC, DailyData.stress_level = 'Medium' DESC
    """, params).fetchall()
    conn.close()
    return render_template("admin_unit_detail.html", admin_name=session["user_name"],
                           unit_name=unit_name, personnel=personnel)


# ----------------------------------------------------------------------
# Reports (PDF / Excel / CSV)
# ----------------------------------------------------------------------
HEADER_BG = colors.HexColor("#2c3e50")


def _pdf_table(data, font_size=None, banded=False):
    table = Table(data, repeatRows=1, hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    if font_size:
        style.append(("FONTSIZE", (0, 0), (-1, -1), font_size))
    if banded:
        style.append(("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]))
    else:
        style.append(("ALIGN", (0, 0), (-1, -1), "CENTER"))
    table.setStyle(TableStyle(style))
    return table


def _pdf_response(elements, filename, doc):
    doc.build(elements)
    return Response(
        doc.filename.getvalue(),
        mimetype="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def _new_pdf(title, extra_lines=()):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm)
    styles = getSampleStyleSheet()
    elements = [Paragraph(xml_escape(title), styles["Title"])]
    for line in extra_lines:
        elements.append(Paragraph(xml_escape(line), styles["Normal"]))
    elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%d-%m-%Y %H:%M')}",
                              styles["Normal"]))
    elements.append(Spacer(1, 16))
    return doc, elements, styles


def _excel_response(wb, filename):
    buffer = BytesIO()
    wb.save(buffer)
    return Response(
        buffer.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def _style_header(ws, row_number):
    fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[row_number]:
        cell.fill = fill
        cell.font = font


def _autosize(ws):
    for col in ws.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(longest, 60) + 4


def _officer_roster(conn, officer_id):
    return conn.execute("""
        SELECT Users.name, Users.unit, DailyData.date, DailyData.sleep_hours,
               DailyData.duty_hours, DailyData.stress_level
        FROM DailyData
        JOIN Users ON DailyData.user_id = Users.id
        WHERE DailyData.id IN (SELECT MAX(id) FROM DailyData GROUP BY user_id)
          AND Users.role = 'personnel'
          AND (Users.assigned_officer_id = ? OR Users.assigned_officer_id IS NULL)
        ORDER BY DailyData.stress_level = 'High' DESC, DailyData.stress_level = 'Medium' DESC,
                 Users.name
    """, (officer_id,)).fetchall()


@app.route("/admin/export_report")
@role_required("admin")
def export_report():
    conn = get_db()
    unit_summary, levels = get_unit_summary(conn)
    conn.close()
    total, high, medium, low = count_levels(levels)

    doc, elements, styles = _new_pdf("APPS - Organization Stress Report")
    elements.append(_pdf_table([
        ["Total Personnel", "High Risk", "Medium Risk", "Low Risk"],
        [str(total), str(high), str(medium), str(low)],
    ]))
    elements += [Spacer(1, 20), Paragraph("Unit-Wise Breakdown", styles["Heading2"]), Spacer(1, 8)]

    data = [["Unit", "Total", "High", "Medium", "Low"]]
    for unit, c in unit_summary.items():
        data.append([unit, c["total"], c["High"], c["Medium"], c["Low"]])
    elements.append(_pdf_table(data, font_size=9, banded=True))
    return _pdf_response(elements, "admin_report.pdf", doc)


@app.route("/admin/export_excel")
@role_required("admin")
def export_excel():
    conn = get_db()
    unit_summary, levels = get_unit_summary(conn)
    conn.close()
    total, high, medium, low = count_levels(levels)

    wb = Workbook()
    ws = wb.active
    ws.title = "Organization Summary"
    ws["A1"] = "APPS - Organization Stress Report"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Generated on: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
    ws.append([])
    ws.append(["Total Personnel", "High Risk", "Medium Risk", "Low Risk"])
    _style_header(ws, ws.max_row)
    ws.append([total, high, medium, low])
    ws.append([])
    ws.append(["Unit", "Total", "High", "Medium", "Low"])
    _style_header(ws, ws.max_row)
    for unit, c in unit_summary.items():
        ws.append([safe_cell(unit), c["total"], c["High"], c["Medium"], c["Low"]])
    _autosize(ws)
    return _excel_response(wb, "admin_report.xlsx")


@app.route("/officer/export_report")
@role_required("officer")
def officer_export_report():
    conn = get_db()
    personnel = _officer_roster(conn, session["user_id"])
    conn.close()
    levels = [p["stress_level"] for p in personnel]
    total, high, medium, low = count_levels(levels)

    doc, elements, styles = _new_pdf("APPS - Personnel Stress Report")
    elements.append(_pdf_table([
        ["Total Personnel", "High Risk", "Medium Risk", "Low Risk"],
        [str(total), str(high), str(medium), str(low)],
    ]))
    elements += [Spacer(1, 20), Paragraph("Personnel List", styles["Heading2"]), Spacer(1, 8)]

    data = [["Name", "Unit", "Date", "Sleep Hrs", "Duty Hrs", "Stress Level"]]
    for p in personnel:
        data.append([p["name"], p["unit"] or "-", p["date"] or "-",
                     p["sleep_hours"], p["duty_hours"], p["stress_level"]])
    elements.append(_pdf_table(data, font_size=9, banded=True))
    return _pdf_response(elements, "officer_report.pdf", doc)


@app.route("/officer/export_excel")
@role_required("officer")
def officer_export_excel():
    conn = get_db()
    personnel = _officer_roster(conn, session["user_id"])
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Personnel Report"
    ws["A1"] = "APPS - Personnel Stress Report"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Generated on: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
    ws.append([])
    ws.append(["Name", "Unit", "Date", "Sleep Hrs", "Duty Hrs", "Stress Level"])
    _style_header(ws, ws.max_row)
    for p in personnel:
        ws.append([safe_cell(p["name"]), safe_cell(p["unit"] or "-"), p["date"] or "-",
                   p["sleep_hours"], p["duty_hours"], p["stress_level"]])
    _autosize(ws)
    return _excel_response(wb, "officer_report.xlsx")


def _my_history(user_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT date, sleep_hours, duty_hours, heart_rate, mood_score, stress_level, recommendation "
        "FROM DailyData WHERE user_id = ? ORDER BY date DESC, id DESC", (user_id,)
    ).fetchall()
    conn.close()
    return rows


@app.route("/my/export_csv")
@role_required("personnel")
def my_export_csv():
    rows = _my_history(session["user_id"])
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow(["Date", "Sleep Hours", "Duty Hours", "Heart Rate", "Mood (1-5)",
                     "Stress Level", "Recommendation"])
    for r in rows:
        writer.writerow([r["date"], r["sleep_hours"], r["duty_hours"], r["heart_rate"],
                         r["mood_score"], r["stress_level"], safe_cell(r["recommendation"])])
    return Response(
        buffer.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=my_checkins.csv"},
    )


@app.route("/my/export_pdf")
@role_required("personnel")
def my_export_pdf():
    rows = _my_history(session["user_id"])
    doc, elements, styles = _new_pdf("APPS - My Check-in History",
    extra_lines=[f"Name: {session['user_name']}"])

    if rows:
        counts = {"Low": 0, "Medium": 0, "High": 0}
        for r in rows:
            counts[r["stress_level"]] = counts.get(r["stress_level"], 0) + 1
        elements.append(_pdf_table([
            ["Total Check-ins", "Low", "Medium", "High"],
            [str(len(rows)), str(counts["Low"]), str(counts["Medium"]), str(counts["High"])],
        ]))
        elements.append(Spacer(1, 20))

        data = [["Date", "Sleep", "Duty", "Heart Rate", "Mood", "Stress"]]
        for r in rows:
            data.append([r["date"], r["sleep_hours"], r["duty_hours"],
           r["heart_rate"], r["mood_score"], r["stress_level"]])
        elements.append(_pdf_table(data, font_size=9, banded=True))
    else:
        elements.append(Paragraph("No check-ins recorded yet.", styles["Normal"]))
    return _pdf_response(elements, "my_checkins.pdf", doc)


# ----------------------------------------------------------------------
# Start-up
# ----------------------------------------------------------------------
init_db()  # also runs when started with `flask run`, not only `python app.py`

if __name__ == "__main__":
    print("Server starting...")
    # Debug mode is OFF by default. For local development only:
    #   set APPS_DEBUG=1   (Windows)   /   export APPS_DEBUG=1   (Linux/Mac)
    app.run(debug=True)