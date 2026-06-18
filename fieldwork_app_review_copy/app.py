
from flask import Flask, render_template, request, redirect, url_for, jsonify, Response, session
import sqlite3, io, zipfile, json, calendar, csv, os, urllib.request, urllib.error
from xml.etree import ElementTree as ET
from datetime import datetime, date, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def load_local_env():
    env_path = os.path.join(BASE_DIR, ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as env_file:
        for line in env_file:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            os.environ.setdefault(key, value)

load_local_env()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "local-dev-only-change-me")

ITEM_TYPES = ["Borehole", "CPTU", "Test Pit", "Geophysics"]
ITEM_STATUSES = ["Planned", "In Progress", "Completed", "Skipped", "Delayed"]
TYPE_PREFIX_MAP = {
    "SCPTU": "CPTU", "SCPT": "CPTU", "CPTU": "CPTU", "CPT": "CPTU",
    "TP": "Test Pit", "MASW": "Geophysics", "ERT": "Geophysics", "ERG": "Geophysics", "BH": "Borehole",
}
STATUS_COLORS = {"Planned": "yellow", "In Progress": "orange", "Completed": "green", "Delayed": "red", "Skipped": "gray"}
NS = {"kml": "http://www.opengis.net/kml/2.2"}

PROJECT_EXPORT_FIELDS = [
    "name", "project_number", "task_code", "client", "site_location",
    "borehole_start_date", "borehole_end_date", "borehole_include_saturday",
    "cptu_start_date", "cptu_end_date", "cptu_include_saturday",
    "test_pit_start_date", "test_pit_end_date", "test_pit_include_saturday",
    "geophysics_start_date", "geophysics_end_date", "geophysics_include_saturday",
    "custom_methods_json", "use_borehole", "use_cptu", "use_test_pit", "use_geophysics",
    "borehole_budget_meters", "cptu_budget_meters", "geophysics_budget_meters",
]

MAP_ITEM_EXPORT_FIELDS = [
    "item_type", "item_id", "geometry_type", "coords_json", "location_plan",
    "planned_amount", "status", "work_start_date", "work_end_date", "notes",
    "depth_m", "exclude_from_history",
]

def env_flag(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}

def int_flag(value):
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, str) and value.strip().lower() in {"true", "yes", "on"}:
        return 1
    return 1 if int(safe_float(value)) else 0

def database_url():
    return os.environ.get("DATABASE_URL", "").strip()

def use_postgres():
    return bool(database_url())

def supabase_url():
    return os.environ.get("SUPABASE_URL", "").strip().rstrip("/")

def supabase_anon_key():
    return os.environ.get("SUPABASE_ANON_KEY", "").strip()

def auth_enabled():
    return bool(supabase_url() and supabase_anon_key()) and not env_flag("LOCAL_AUTH_BYPASS", False)

def setup_status():
    status = {
        "ok": True,
        "database_mode": "supabase_postgres" if use_postgres() else "local_sqlite",
        "database_connected": False,
        "auth_configured": auth_enabled(),
        "supabase_url_configured": bool(supabase_url()),
        "supabase_anon_key_configured": bool(supabase_anon_key()),
        "local_auth_bypass": env_flag("LOCAL_AUTH_BYPASS", False),
        "project_count": None,
        "errors": [],
    }
    try:
        conn = get_db()
        row = conn.execute("SELECT COUNT(*) AS count FROM projects").fetchone()
        conn.close()
        status["database_connected"] = True
        status["project_count"] = row["count"]
    except Exception as exc:
        status["ok"] = False
        status["errors"].append(f"Database check failed: {exc}")
    if use_postgres() and not status["auth_configured"]:
        status["ok"] = False
        status["errors"].append("Cloud database is configured, but Supabase Auth is not fully configured.")
    return status

class PostgresConnection:
    def __init__(self):
        import psycopg
        from psycopg.rows import dict_row
        self.conn = psycopg.connect(database_url(), row_factory=dict_row)

    def execute(self, sql, params=None):
        translated_sql = sql.replace("?", "%s")
        return self.conn.execute(translated_sql, tuple(params or ()))

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.close()

def get_db():
    if use_postgres():
        return PostgresConnection()
    conn = sqlite3.connect(os.path.join(BASE_DIR, "data.db"))
    conn.row_factory = sqlite3.Row
    return conn

def supabase_password_login(email, password):
    login_url = f"{supabase_url()}/auth/v1/token?grant_type=password"
    body = json.dumps({"email": email, "password": password}).encode("utf-8")
    req = urllib.request.Request(
        login_url,
        data=body,
        method="POST",
        headers={
            "apikey": supabase_anon_key(),
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return True, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        try:
            payload = json.loads(exc.read().decode("utf-8"))
            message = payload.get("msg") or payload.get("error_description") or payload.get("error") or "Login failed."
        except Exception:
            message = "Login failed."
        return False, message
    except Exception as exc:
        return False, f"Could not contact Supabase Auth: {exc}"

@app.context_processor
def inject_auth_state():
    return {
        "auth_enabled": auth_enabled(),
        "current_user_email": session.get("user_email", ""),
        "using_cloud_database": use_postgres(),
    }

@app.before_request
def require_login_for_hosted_app():
    if not auth_enabled():
        return None
    if request.endpoint in {"login", "logout", "healthz", "setup_status_json", "static"}:
        return None
    if request.path.startswith("/static/"):
        return None
    if not session.get("supabase_access_token"):
        next_url = request.full_path if request.query_string else request.path
        return redirect(url_for("login", next=next_url))
    return None

@app.route("/login", methods=["GET", "POST"])
def login():
    if not auth_enabled():
        return redirect(url_for("index"))
    error = ""
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        ok, result = supabase_password_login(email, password)
        if ok:
            user = result.get("user") or {}
            session["supabase_access_token"] = result.get("access_token", "")
            session["user_email"] = user.get("email") or email
            next_url = request.args.get("next") or url_for("index")
            if not next_url.startswith("/"):
                next_url = url_for("index")
            return redirect(next_url)
        error = result
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login" if auth_enabled() else "index"))

@app.route("/healthz")
def healthz():
    status = setup_status()
    return jsonify({
        "ok": status["ok"],
        "database_connected": status["database_connected"],
        "database_mode": status["database_mode"],
        "auth_configured": status["auth_configured"],
    }), 200 if status["ok"] else 503

@app.route("/setup-status")
def setup_status_json():
    status = setup_status()
    public_status = dict(status)
    if auth_enabled() and not session.get("supabase_access_token"):
        public_status.pop("project_count", None)
    return jsonify(public_status), 200 if status["ok"] else 503

def safe_float(v):
    try:
        return float(v) if str(v).strip() else 0.0
    except Exception:
        return 0.0

def parse_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date() if s else None
    except Exception:
        return None

def fmt_date(d):
    return d.strftime("%Y-%m-%d") if d else ""

def parse_int(value, default, min_value=None, max_value=None):
    try:
        out = int(value)
    except Exception:
        out = default
    if min_value is not None:
        out = max(out, min_value)
    if max_value is not None:
        out = min(out, max_value)
    return out

def clean_status(value, default="Planned"):
    value = (value or "").strip()
    return value if value in ITEM_STATUSES else default

def parse_custom_methods(raw):
    try:
        data = json.loads(raw or "[]")
        if isinstance(data, list):
            out = []
            for x in data:
                if isinstance(x, dict):
                    name = str(x.get("name","")).strip()
                    if name:
                        out.append({
                            "name": name,
                            "color": str(x.get("color","#0f766e")).strip() or "#0f766e",
                            "symbol": str(x.get("symbol","circle")).strip() or "circle",
                        })
            return out
    except Exception:
        pass
    return []


def method_enabled(project, field):
    val = project.get(field, 1)
    if val is None or str(val).strip() == "":
        return True
    try:
        return bool(int(val))
    except Exception:
        return True

def build_method_definitions(project):
    defs = []
    if method_enabled(project, "use_borehole"):
        defs.append({"name":"Borehole","color":"#2563eb","symbol":"circle","start_field":"borehole_start_date","end_field":"borehole_end_date","sat_field":"borehole_include_saturday","meter_field":"borehole_budget_meters"})
    if method_enabled(project, "use_cptu"):
        defs.append({"name":"CPTU","color":"#7c3aed","symbol":"triangle","start_field":"cptu_start_date","end_field":"cptu_end_date","sat_field":"cptu_include_saturday","meter_field":"cptu_budget_meters"})
    if method_enabled(project, "use_test_pit"):
        defs.append({"name":"Test Pit","color":"#ea580c","symbol":"square","start_field":"test_pit_start_date","end_field":"test_pit_end_date","sat_field":"test_pit_include_saturday"})
    if method_enabled(project, "use_geophysics"):
        defs.append({"name":"Geophysics","color":"#16a34a","symbol":"line","start_field":"geophysics_start_date","end_field":"geophysics_end_date","sat_field":"geophysics_include_saturday","meter_field":"geophysics_budget_meters"})
    defs.extend(parse_custom_methods(project.get("custom_methods_json", "[]")))
    return defs


def working_days_between(start_s, end_s, include_saturday):
    start = parse_date(start_s)
    end = parse_date(end_s)
    if not start or not end or end < start:
        return 0
    total = 0
    d = start
    while d <= end:
        wd = d.weekday()
        if wd < 5 or (include_saturday and wd == 5):
            total += 1
        d += timedelta(days=1)
    return total

def columns(conn, table):
    return [r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]

def init_db():
    if use_postgres():
        return
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            project_number TEXT DEFAULT '',
            task_code TEXT DEFAULT '',
            client TEXT DEFAULT '',
            site_location TEXT DEFAULT '',
            borehole_start_date TEXT DEFAULT '',
            borehole_end_date TEXT DEFAULT '',
            borehole_include_saturday INTEGER DEFAULT 0,
            cptu_start_date TEXT DEFAULT '',
            cptu_end_date TEXT DEFAULT '',
            cptu_include_saturday INTEGER DEFAULT 0,
            test_pit_start_date TEXT DEFAULT '',
            test_pit_end_date TEXT DEFAULT '',
            test_pit_include_saturday INTEGER DEFAULT 0,
            geophysics_start_date TEXT DEFAULT '',
            geophysics_end_date TEXT DEFAULT '',
            geophysics_include_saturday INTEGER DEFAULT 0,
            borehole_budget_meters REAL DEFAULT 0,
            cptu_budget_meters REAL DEFAULT 0,
            geophysics_budget_meters REAL DEFAULT 0
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS map_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            item_type TEXT NOT NULL,
            item_id TEXT NOT NULL,
            geometry_type TEXT NOT NULL,
            coords_json TEXT NOT NULL,
            location_plan TEXT DEFAULT '',
            planned_amount REAL DEFAULT 0,
            status TEXT DEFAULT 'Planned',
            work_start_date TEXT DEFAULT '',
            work_end_date TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            depth_m REAL DEFAULT 0
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS import_backups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            item_count INTEGER NOT NULL,
            backup_json TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS historical_rates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_item_id INTEGER UNIQUE,
            project_id INTEGER NOT NULL,
            project_name TEXT DEFAULT '',
            item_type TEXT NOT NULL,
            item_id TEXT NOT NULL,
            work_start_date TEXT DEFAULT '',
            work_end_date TEXT DEFAULT '',
            completion_month TEXT DEFAULT '',
            work_days INTEGER DEFAULT 0,
            depth_m REAL DEFAULT 0,
            items_per_day REAL DEFAULT 0,
            meters_per_day REAL,
            recorded_at TEXT NOT NULL
        )
    """)
    pcols = columns(conn, "projects")
    needed = {
        "borehole_start_date": "TEXT DEFAULT ''", "borehole_end_date": "TEXT DEFAULT ''", "borehole_include_saturday": "INTEGER DEFAULT 0",
        "cptu_start_date": "TEXT DEFAULT ''", "cptu_end_date": "TEXT DEFAULT ''", "cptu_include_saturday": "INTEGER DEFAULT 0",
        "test_pit_start_date": "TEXT DEFAULT ''", "test_pit_end_date": "TEXT DEFAULT ''", "test_pit_include_saturday": "INTEGER DEFAULT 0",
        "geophysics_start_date": "TEXT DEFAULT ''", "geophysics_end_date": "TEXT DEFAULT ''", "geophysics_include_saturday": "INTEGER DEFAULT 0",
        "custom_methods_json": "TEXT DEFAULT '[]'",
        "use_borehole": "INTEGER DEFAULT 1",
        "use_cptu": "INTEGER DEFAULT 1",
        "use_test_pit": "INTEGER DEFAULT 1",
        "use_geophysics": "INTEGER DEFAULT 1",
        "borehole_budget_meters": "REAL DEFAULT 0",
        "cptu_budget_meters": "REAL DEFAULT 0",
        "geophysics_budget_meters": "REAL DEFAULT 0",
    }
    for col, ddl in needed.items():
        if col not in pcols:
            conn.execute(f"ALTER TABLE projects ADD COLUMN {col} {ddl}")
    icols = columns(conn, "map_items")
    ineeded = {
        "planned_amount": "REAL DEFAULT 0", "status": "TEXT DEFAULT 'Planned'",
        "work_start_date": "TEXT DEFAULT ''", "work_end_date": "TEXT DEFAULT ''",
        "notes": "TEXT DEFAULT ''", "location_plan": "TEXT DEFAULT ''", "depth_m": "REAL DEFAULT 0",
        "exclude_from_history": "INTEGER DEFAULT 0",
    }
    for col, ddl in ineeded.items():
        if col not in icols:
            conn.execute(f"ALTER TABLE map_items ADD COLUMN {col} {ddl}")
    conn.commit()
    sync_historical_rates(conn)
    conn.close()

def detect_type(name):
    upper = (name or "").strip().upper()
    for prefix, item_type in TYPE_PREFIX_MAP.items():
        if upper.startswith(prefix):
            return item_type
    if "ERT" in upper or "MASW" in upper or "ERG" in upper or "LINE" in upper:
        return "Geophysics"
    return "Borehole"

def parse_coord_text(text):
    coords = []
    for token in text.strip().split():
        parts = token.split(",")
        if len(parts) >= 2:
            try:
                lon = float(parts[0]); lat = float(parts[1])
            except ValueError:
                continue
            coords.append([lat, lon])
    return coords

def parse_kml_bytes(file_bytes):
    root = ET.fromstring(file_bytes)
    items = []
    for placemark in root.findall(".//kml:Placemark", NS):
        name_el = placemark.find("kml:name", NS)
        desc_el = placemark.find("kml:description", NS)
        name = name_el.text.strip() if name_el is not None and name_el.text else "Untitled"
        desc = desc_el.text.strip() if desc_el is not None and desc_el.text else ""
        point_el = placemark.find(".//kml:Point/kml:coordinates", NS)
        line_el = placemark.find(".//kml:LineString/kml:coordinates", NS)
        if point_el is not None and point_el.text:
            coords = parse_coord_text(point_el.text)
            if coords:
                items.append({"item_type": detect_type(name), "item_id": name, "geometry_type": "Point", "coords_json": json.dumps(coords[0]), "location_plan": "", "planned_amount": 1.0, "status": "Planned", "work_start_date": "", "work_end_date": "", "notes": desc})
        elif line_el is not None and line_el.text:
            coords = parse_coord_text(line_el.text)
            if coords:
                items.append({"item_type": detect_type(name), "item_id": name, "geometry_type": "LineString", "coords_json": json.dumps(coords), "location_plan": "", "planned_amount": 1.0, "status": "Planned", "work_start_date": "", "work_end_date": "", "notes": desc})
    return items

def extract_kml_bytes(upload_file):
    filename = (upload_file.filename or "").lower()
    raw = upload_file.read()
    if filename.endswith(".kml"):
        return raw
    if filename.endswith(".kmz"):
        with zipfile.ZipFile(io.BytesIO(raw), "r") as zf:
            kml_name = None
            for n in zf.namelist():
                if n.lower() == "doc.kml":
                    kml_name = n
                    break
            if not kml_name:
                for n in zf.namelist():
                    if n.lower().endswith(".kml"):
                        kml_name = n
                        break
            if not kml_name:
                raise ValueError("No KML found in KMZ.")
            return zf.read(kml_name)
    raise ValueError("Only KML and KMZ supported.")

def normalize_item_dates(item):
    if item["item_type"] == "Test Pit" and item.get("work_start_date") and not item.get("work_end_date"):
        item["work_end_date"] = item["work_start_date"]
    return item

def item_dict(row):
    d = dict(row)
    d["planned_amount"] = safe_float(d.get("planned_amount"))
    d["depth_m"] = safe_float(d.get("depth_m"))
    d["exclude_from_history"] = int(safe_float(d.get("exclude_from_history")))
    d["status_color"] = STATUS_COLORS.get(d["status"], "gray")
    try:
        d["coords"] = json.loads(d["coords_json"])
    except Exception:
        d["coords"] = None
    return normalize_item_dates(d)

def project_items(conn, pid):
    return [item_dict(r) for r in conn.execute("SELECT * FROM map_items WHERE project_id=? ORDER BY item_type, item_id", (pid,)).fetchall()]

def is_skipped(item):
    return item.get("status") == "Skipped"

def calculable_items(items):
    return [item for item in items if not is_skipped(item)]

def filter_items_by_methods(items, method_defs):
    method_names = {m["name"] for m in method_defs}
    return [item for item in items if item["item_type"] in method_names]

def method_meta_map(project):
    return {m["name"]: m for m in build_method_definitions(project)}

def geometry_type_for_item_type(project, item_type):
    meta = method_meta_map(project).get(item_type, {})
    return "LineString" if item_type == "Geophysics" or meta.get("symbol") == "line" else "Point"

def is_single_day_type(item_type):
    return item_type == "Test Pit"

def grouped_items(items):
    names = sorted(set([x["item_type"] for x in items] + ITEM_TYPES))
    out = {k: [] for k in names}
    for item in items:
        out.setdefault(item["item_type"], []).append(item)
    return out





def project_summary(items, project):
    active_items = calculable_items(items)
    summary = {
        "total": len(active_items),
        "raw_total": len(items),
        "skipped": len(items) - len(active_items),
        "detail": {},
        "meters": {}
    }
    mapping = {}
    for m in build_method_definitions(project):
        mapping[m["name"]] = m

    def rate_class_from_ratio(current_rate, target_rate):
        if target_rate <= 0:
            return "green"
        ratio = current_rate / target_rate if target_rate else 0
        if ratio < 0.50:
            return "dark-red"
        if ratio < 0.75:
            return "red"
        if ratio < 0.90:
            return "light-red"
        if ratio < 1.00:
            return "orange"
        if ratio < 1.10:
            return "yellow"
        if ratio < 1.25:
            return "light-green"
        return "green"

    scores = []
    for typ, meta in mapping.items():
        start_f = meta.get("start_field")
        end_f = meta.get("end_field")
        sat_f = meta.get("sat_field")
        meter_f = meta.get("meter_field")
        include_sat = bool(project.get(sat_f)) if sat_f else False
        budget_days = working_days_between(project.get(start_f), project.get(end_f), include_sat) if start_f and end_f else 0

        rows = [x for x in active_items if x["item_type"] == typ]
        total_items = len(rows)
        completed_rows = [x for x in rows if x["status"] == "Completed"]
        completed = len(completed_rows)

        completed_days = 0
        item_rates = []
        for r in completed_rows:
            end_date = r.get("work_end_date") or r.get("work_start_date")
            item_days = working_days_between(r.get("work_start_date"), end_date, include_sat)
            completed_days += item_days
            if item_days > 0:
                item_rates.append(1.0 / item_days)

        current_rate = round(sum(item_rates) / len(item_rates), 3) if item_rates else 0.0
        remaining_items = max(total_items - completed, 0)
        days_left = max(budget_days - completed_days, 0)
        target_rate = round(remaining_items / days_left, 3) if days_left > 0 and remaining_items > 0 else 0.0
        projected_days_needed = round((remaining_items / current_rate), 1) if current_rate > 0 else None
        rate_class = rate_class_from_ratio(current_rate, target_rate)
        pct_done = round((completed / total_items) * 100, 1) if total_items > 0 else 0.0

        budget_meters = safe_float(project.get(meter_f)) if meter_f else 0.0
        logged_meters = round(sum(safe_float(x.get("depth_m")) for x in rows), 2)
        completed_meters = round(sum(safe_float(x.get("depth_m")) for x in completed_rows), 2)

        summary["detail"][typ] = {
            "budget_days": budget_days, "total_items": total_items, "completed": completed,
            "remaining": remaining_items, "completed_days": completed_days, "days_left": days_left,
            "current_rate": current_rate, "target_rate": target_rate,
            "projected_days_needed": projected_days_needed, "pct_done": pct_done,
            "status_class": rate_class
        }
        summary["meters"][typ] = {
            "budget_meters": budget_meters,
            "logged_meters": logged_meters,
            "completed_meters": completed_meters,
        }
        scores.append(pct_done)

    avg = round(sum(scores) / len(scores), 1) if scores else 0.0
    overall_code = "GREEN" if avg >= 90 else ("ORANGE" if avg >= 50 else "RED")
    summary["overall_code"] = overall_code
    summary["overall_class"] = overall_code.lower()
    summary["overall_avg"] = avg
    summary["overall_text"] = "Overall project view based on completed items versus planned date windows."
    return summary


def include_saturday_for_type(project, item_type):
    for meta in build_method_definitions(project):
        if meta["name"] == item_type and meta.get("sat_field"):
            return bool(project.get(meta["sat_field"]))
    return False

def item_work_days(item, project):
    end_date = item.get("work_end_date") or item.get("work_start_date")
    return working_days_between(item.get("work_start_date"), end_date, include_saturday_for_type(project, item.get("item_type")))

def avg(values):
    values = [v for v in values if v is not None]
    return round(sum(values) / len(values), 3) if values else 0.0

def percentile(values, pct):
    values = sorted(v for v in values if v is not None)
    if not values:
        return None
    if len(values) == 1:
        return values[0]
    pos = (len(values) - 1) * pct
    lower = int(pos)
    upper = min(lower + 1, len(values) - 1)
    weight = pos - lower
    return round(values[lower] + (values[upper] - values[lower]) * weight, 2)

def build_historical_rate_rows(conn):
    projects = {r["id"]: dict(r) for r in conn.execute("SELECT * FROM projects").fetchall()}
    rows = []
    for row in conn.execute("SELECT * FROM map_items WHERE status='Completed' AND COALESCE(exclude_from_history, 0)=0 ORDER BY item_type, item_id").fetchall():
        item = item_dict(row)
        project = projects.get(item["project_id"], {})
        days = item_work_days(item, project)
        if days <= 0:
            continue
        depth = safe_float(item.get("depth_m"))
        end_date = item.get("work_end_date") or item.get("work_start_date") or ""
        rows.append({
            "source_item_id": item["id"],
            "project_id": item["project_id"],
            "project_name": project.get("name", ""),
            "item_id": item["item_id"],
            "item_type": item["item_type"],
            "work_start_date": item.get("work_start_date", ""),
            "work_end_date": end_date,
            "completion_month": end_date[:7] if len(end_date) >= 7 else "",
            "work_days": days,
            "depth_m": depth,
            "items_per_day": round(1.0 / days, 3),
            "meters_per_day": round(depth / days, 3) if depth > 0 else None,
        })
    return rows

def sync_historical_rates(conn):
    rows = build_historical_rate_rows(conn)
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute("DELETE FROM historical_rates")
    for row in rows:
        conn.execute("""
            INSERT INTO historical_rates
            (source_item_id, project_id, project_name, item_type, item_id, work_start_date,
             work_end_date, completion_month, work_days, depth_m, items_per_day, meters_per_day, recorded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            row["source_item_id"], row["project_id"], row["project_name"], row["item_type"], row["item_id"],
            row["work_start_date"], row["work_end_date"], row["completion_month"], row["work_days"],
            row["depth_m"], row["items_per_day"], row["meters_per_day"], now
        ))
    conn.commit()
    return rows

def historical_rate_dataset(conn):
    rows = []
    for row in conn.execute("SELECT * FROM historical_rates ORDER BY item_type, completion_month, item_id").fetchall():
        d = dict(row)
        d.pop("id", None)
        d.pop("recorded_at", None)
        rows.append(d)
    return rows

def rate_trend_context(history):
    by_type_month = {}
    for row in history:
        month = row.get("completion_month") or "Undated"
        by_type_month.setdefault(row["item_type"], {}).setdefault(month, []).append(row)
    series = []
    max_rate = 1
    for typ, months in sorted(by_type_month.items()):
        points = []
        for month, rows in sorted(months.items()):
            items_rate = avg([r["items_per_day"] for r in rows])
            meter_rate = avg([r["meters_per_day"] for r in rows if r["meters_per_day"] is not None])
            max_rate = max(max_rate, items_rate)
            points.append({"month": month, "items_per_day": items_rate, "meters_per_day": meter_rate, "count": len(rows)})
        series.append({"type": typ, "points": points})
    return {"series": series, "max_rate": max_rate}

def budgeting_context(conn, project, items):
    history = historical_rate_dataset(conn)
    active_items = calculable_items(items)
    by_type = {}
    for row in history:
        by_type.setdefault(row["item_type"], []).append(row)

    methods = build_method_definitions(project)
    cards = []
    for meta in methods:
        typ = meta["name"]
        rows = by_type.get(typ, [])
        current_items = [x for x in active_items if x["item_type"] == typ]
        include_sat = bool(project.get(meta.get("sat_field"))) if meta.get("sat_field") else False
        planned_days = working_days_between(project.get(meta.get("start_field")), project.get(meta.get("end_field")), include_sat) if meta.get("start_field") and meta.get("end_field") else 0
        budget_meters = safe_float(project.get(meta.get("meter_field"))) if meta.get("meter_field") else 0.0
        logged_meters = round(sum(safe_float(x.get("depth_m")) for x in current_items), 2)
        meter_basis = budget_meters or logged_meters
        avg_items_rate = avg([r["items_per_day"] for r in rows])
        avg_meter_rate = avg([r["meters_per_day"] for r in rows if r["meters_per_day"] is not None])
        estimated_days_by_items = round(len(current_items) / avg_items_rate, 1) if avg_items_rate > 0 and current_items else None
        estimated_days_by_meters = round(meter_basis / avg_meter_rate, 1) if avg_meter_rate > 0 and meter_basis > 0 else None
        estimates = [x for x in [estimated_days_by_items, estimated_days_by_meters] if x is not None]
        recommended_days = max(estimates) if estimates else None
        variance = round(planned_days - recommended_days, 1) if planned_days and recommended_days is not None else None
        if recommended_days is None:
            status = "gray"
        elif planned_days >= recommended_days:
            status = "green"
        elif planned_days >= recommended_days * 0.85:
            status = "orange"
        else:
            status = "red"
        cards.append({
            "type": typ,
            "sample_count": len(rows),
            "current_items": len(current_items),
            "planned_days": planned_days,
            "budget_meters": budget_meters,
            "logged_meters": logged_meters,
            "avg_items_per_day": avg_items_rate,
            "avg_meters_per_day": avg_meter_rate,
            "median_days_per_item": percentile([r["work_days"] for r in rows], 0.5),
            "p80_days_per_item": percentile([r["work_days"] for r in rows], 0.8),
            "estimated_days_by_items": estimated_days_by_items,
            "estimated_days_by_meters": estimated_days_by_meters,
            "recommended_days": recommended_days,
            "variance_days": variance,
            "status": status,
        })

    return {
        "history": history,
        "trends": rate_trend_context(history),
        "cards": cards,
        "max_sample_count": max([c["sample_count"] for c in cards] + [1]),
        "max_rate": max([c["avg_items_per_day"] for c in cards] + [1]),
        "max_days": max([c["recommended_days"] or 0 for c in cards] + [1]),
    }

def dashboard_context(summary, items, budget):
    active_items = calculable_items(items)
    completed = sum(1 for item in active_items if item["status"] == "Completed")
    remaining = max(len(active_items) - completed, 0)
    completion_pct = round((completed / len(active_items)) * 100, 1) if active_items else 0.0
    at_risk = [c for c in budget["cards"] if c["status"] in {"red", "orange"}]
    recommended_total = round(sum(c["recommended_days"] or 0 for c in budget["cards"]), 1)
    planned_total = sum(c["planned_days"] for c in budget["cards"])
    variance = round(planned_total - recommended_total, 1) if recommended_total else None
    return {
        "completion_pct": completion_pct,
        "completed": completed,
        "remaining": remaining,
        "skipped": summary["skipped"],
        "active_total": len(active_items),
        "raw_total": len(items),
        "at_risk_count": len(at_risk),
        "at_risk_methods": ", ".join(c["type"] for c in at_risk) if at_risk else "None",
        "planned_total_days": planned_total,
        "recommended_total_days": recommended_total,
        "variance_days": variance,
    }




def month_context(project, items, year, month):
    cal = calendar.Calendar(firstweekday=0)
    weeks = cal.monthdatescalendar(year, month)
    month_start = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    def split_segments(start_d, end_d, include_saturday):
        segments = []
        current_start = None
        previous_day = None
        d = start_d
        while d <= end_d:
            workday = (d.weekday() < 5) or (include_saturday and d.weekday() == 5)
            if workday:
                if current_start is None:
                    current_start = d
                previous_day = d
            else:
                if current_start is not None and previous_day is not None:
                    segments.append((current_start, previous_day))
                    current_start = None
                    previous_day = None
            d += timedelta(days=1)
        if current_start is not None and previous_day is not None:
            segments.append((current_start, previous_day))
        return segments

    plans = []
    for m in build_method_definitions(project):
        if m.get("start_field") and m.get("end_field"):
            plans.append((m["name"], project.get(m["start_field"]), project.get(m["end_field"]), bool(project.get(m.get("sat_field"))), m.get("color","#0f766e")))

    actual_items = []
    for item in items:
        if is_skipped(item):
            continue
        s = parse_date(item.get("work_start_date"))
        e = parse_date(item.get("work_end_date") or item.get("work_start_date"))
        if s and e and not (e < month_start or s > month_end):
            color = next((m.get("color","#0f766e") for m in build_method_definitions(project) if m["name"] == item["item_type"]), "#0f766e")
            actual_items.append({"id": item["id"], "label": item["item_id"], "type": item["item_type"], "status": item["status"], "start": max(s, month_start), "end": min(e, month_end), "color": color})

    week_rows = []
    for week in weeks:
        week_start = week[0]
        week_end = week[-1]
        planned_rows = []
        actual_rows = []

        for label, start_s, end_s, include_saturday, color in plans:
            s = parse_date(start_s)
            e = parse_date(end_s)
            if not s or not e:
                continue
            visible_start = max(s, month_start)
            visible_end = min(e, month_end)
            for seg_start, seg_end in split_segments(visible_start, visible_end, include_saturday):
                if seg_end < week_start or seg_start > week_end:
                    continue
                display_start = max(seg_start, week_start)
                display_end = min(seg_end, week_end)
                start_col = week.index(display_start) + 1
                end_col = week.index(display_end) + 1
                planned_rows.append({
                    "label": label,
                    "type": label,
                    "start_col": start_col,
                    "end_col": end_col,
                    "color": color,
                    "is_true_start": display_start == s,
                    "is_true_end": display_end == e
                })

        for item in actual_items:
            if item["end"] < week_start or item["start"] > week_end:
                continue
            display_start = max(item["start"], week_start)
            display_end = min(item["end"], week_end)
            start_col = week.index(display_start) + 1
            end_col = week.index(display_end) + 1
            actual_rows.append({
                "label": item["label"], "type": item["type"], "status": item["status"], "id": item["id"],
                "start_col": start_col, "end_col": end_col, "color": item["color"]
            })

        week_rows.append({"days": week, "planned_rows": planned_rows, "actual_rows": actual_rows})

    return {"weeks": weeks, "week_rows": week_rows, "actual_items": actual_items}


def insert_project_record(conn, project_data):
    values = []
    for field in PROJECT_EXPORT_FIELDS:
        if field == "custom_methods_json":
            value = project_data.get(field, "[]")
            if isinstance(value, list):
                value = json.dumps(value)
        elif field.startswith("use_") or field.endswith("_include_saturday"):
            value = int_flag(project_data.get(field, 0))
        elif field.endswith("_budget_meters"):
            value = safe_float(project_data.get(field))
        else:
            value = project_data.get(field, "") or ""
        values.append(value)
    placeholders = ", ".join("?" for _ in PROJECT_EXPORT_FIELDS)
    columns_sql = ", ".join(PROJECT_EXPORT_FIELDS)
    if use_postgres():
        row = conn.execute(
            f"INSERT INTO projects ({columns_sql}) VALUES ({placeholders}) RETURNING id",
            values,
        ).fetchone()
        return row["id"]
    cursor = conn.execute(
        f"INSERT INTO projects ({columns_sql}) VALUES ({placeholders})",
        values,
    )
    return cursor.lastrowid


def insert_project_item_record(conn, project_id, item_data):
    values = []
    for field in MAP_ITEM_EXPORT_FIELDS:
        if field in {"planned_amount", "depth_m"}:
            value = safe_float(item_data.get(field))
        elif field == "exclude_from_history":
            value = int_flag(item_data.get(field, 0))
        elif field == "status":
            value = clean_status(item_data.get(field))
        elif field == "geometry_type":
            value = item_data.get(field) or "Point"
        elif field == "coords_json":
            value = item_data.get(field) or json.dumps([0.0, 0.0])
        else:
            value = item_data.get(field, "") or ""
        values.append(value)
    columns_sql = ", ".join(["project_id"] + MAP_ITEM_EXPORT_FIELDS)
    placeholders = ", ".join("?" for _ in ["project_id"] + MAP_ITEM_EXPORT_FIELDS)
    conn.execute(
        f"INSERT INTO map_items ({columns_sql}) VALUES ({placeholders})",
        [project_id] + values,
    )


def safe_project_filename(project):
    base = f"{project.get('project_number') or project.get('name') or 'fieldwork_project'}"
    cleaned = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in base)
    return cleaned.strip("_") or "fieldwork_project"


@app.route("/")
def index():
    conn = get_db()
    projects = [dict(r) for r in conn.execute("SELECT * FROM projects ORDER BY id DESC").fetchall()]
    conn.close()
    return render_template("index.html", projects=projects, notice=request.args.get("notice", ""), error=request.args.get("error", ""))


@app.route("/add_project", methods=["POST"])
def add_project():
    custom_methods = []
    extra_names = request.form.getlist("custom_method_name")
    extra_colors = request.form.getlist("custom_method_color")
    extra_symbols = request.form.getlist("custom_method_symbol")
    for i, name in enumerate(extra_names):
        name = (name or "").strip()
        if name:
            custom_methods.append({"name": name, "color": extra_colors[i] if i < len(extra_colors) else "#0f766e", "symbol": extra_symbols[i] if i < len(extra_symbols) else "circle"})
    conn = get_db()
    conn.execute("""INSERT INTO projects
    (name, project_number, task_code, client, site_location,
    borehole_start_date, borehole_end_date, borehole_include_saturday,
    cptu_start_date, cptu_end_date, cptu_include_saturday,
    test_pit_start_date, test_pit_end_date, test_pit_include_saturday,
    geophysics_start_date, geophysics_end_date, geophysics_include_saturday,
    custom_methods_json, use_borehole, use_cptu, use_test_pit, use_geophysics, borehole_budget_meters, cptu_budget_meters, geophysics_budget_meters)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
    (
        request.form.get("name","").strip(), request.form.get("project_number","").strip(), request.form.get("task_code","").strip(),
        request.form.get("client","").strip(), request.form.get("site_location","").strip(),
        request.form.get("borehole_start_date","").strip(), request.form.get("borehole_end_date","").strip(), 1 if request.form.get("borehole_include_saturday") else 0,
        request.form.get("cptu_start_date","").strip(), request.form.get("cptu_end_date","").strip(), 1 if request.form.get("cptu_include_saturday") else 0,
        request.form.get("test_pit_start_date","").strip(), request.form.get("test_pit_end_date","").strip(), 1 if request.form.get("test_pit_include_saturday") else 0,
        request.form.get("geophysics_start_date","").strip(), request.form.get("geophysics_end_date","").strip(), 1 if request.form.get("geophysics_include_saturday") else 0,
        json.dumps(custom_methods),
        1 if request.form.get("use_borehole") else 0,
        1 if request.form.get("use_cptu") else 0,
        1 if request.form.get("use_test_pit") else 0,
        1 if request.form.get("use_geophysics") else 0,
        safe_float(request.form.get("borehole_budget_meters")),
        safe_float(request.form.get("cptu_budget_meters")),
        safe_float(request.form.get("geophysics_budget_meters")),
    ))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("index"))


@app.route("/import_project_json", methods=["POST"])
def import_project_json():
    upload = request.files.get("project_json")
    if not upload or not upload.filename:
        return redirect(url_for("index", error="Choose a project JSON export first."))
    try:
        payload = json.loads(upload.read().decode("utf-8"))
        project_payload = payload.get("project") or {}
        items_payload = payload.get("items") or []
        if not isinstance(project_payload, dict) or not isinstance(items_payload, list):
            raise ValueError("Invalid project export format.")
    except Exception as exc:
        return redirect(url_for("index", error=f"Import failed: {str(exc)[:160]}"))

    project_payload = dict(project_payload)
    original_name = project_payload.get("name") or "Imported Project"
    project_payload["name"] = f"{original_name} (Imported)"
    conn = get_db()
    try:
        new_pid = insert_project_record(conn, project_payload)
        imported_count = 0
        for item in items_payload:
            if isinstance(item, dict):
                insert_project_item_record(conn, new_pid, item)
                imported_count += 1
        conn.commit()
        sync_historical_rates(conn)
    finally:
        conn.close()
    return redirect(url_for("project", pid=new_pid, tab="overview", notice=f"Imported project with {imported_count} items."))


@app.route("/project/<int:pid>")
def project(pid):
    tab = request.args.get("tab", "overview")
    year = parse_int(request.args.get("year"), date.today().year, 1970, 2100)
    month = parse_int(request.args.get("month"), date.today().month, 1, 12)
    conn = get_db()
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not prow:
        conn.close(); return redirect(url_for("index"))
    project = dict(prow)
    items = project_items(conn, pid)
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    method_defs = build_method_definitions(project)
    visible_items = filter_items_by_methods(items, method_defs)
    group_keys = {typ: f"group_{i}" for i, typ in enumerate(sorted(set([x["item_type"] for x in visible_items] + [m["name"] for m in method_defs])), start=1)}
    calctx = month_context(project, visible_items, year, month)
    prev_month = (date(year, month, 1) - timedelta(days=1))
    next_month = (date(year, month, calendar.monthrange(year, month)[1]) + timedelta(days=1))
    active_names = sorted(set([x["item_type"] for x in visible_items] + [m["name"] for m in method_defs]))
    grouped = {name: [x for x in visible_items if x["item_type"] == name] for name in active_names}
    backups = [dict(r) for r in conn.execute(
        "SELECT id, created_at, item_count FROM import_backups WHERE project_id=? ORDER BY id DESC LIMIT 8",
        (pid,)
    ).fetchall()]
    summary = project_summary(visible_items, project)
    budget = budgeting_context(conn, project, visible_items)
    dashboard = dashboard_context(summary, visible_items, budget)
    conn.close()
    return render_template(
        "project.html",
        project=project, items=visible_items, grouped=grouped, summary=summary,
        tab=tab, statuses=ITEM_STATUSES, types=sorted(set([x["item_type"] for x in visible_items] + [m["name"] for m in method_defs])), group_keys=group_keys, method_defs=method_defs,
        calctx=calctx, cal_year=year, cal_month=month,
        prev_year=prev_month.year, prev_month=prev_month.month, next_year=next_month.year, next_month=next_month.month,
        month_name=calendar.month_name[month], budget=budget, dashboard=dashboard, backups=backups,
        notice=request.args.get("notice", ""), error=request.args.get("error", "")
    )


@app.route("/project/<int:pid>/budget_data")
def budget_data(pid):
    conn = get_db()
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not prow:
        conn.close()
        return jsonify({"error": "Project not found"}), 404
    project = dict(prow)
    items = project_items(conn, pid)
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    data = budgeting_context(conn, project, filter_items_by_methods(items, build_method_definitions(project)))
    conn.close()
    return jsonify(data)

@app.route("/project/<int:pid>/budget_history.csv")
def budget_history_csv(pid):
    conn = get_db()
    rows = historical_rate_dataset(conn)
    conn.close()
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=[
        "source_item_id", "project_id", "project_name", "item_type", "item_id", "work_start_date",
        "work_end_date", "completion_month", "work_days", "depth_m", "items_per_day", "meters_per_day"
    ], extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=project_{pid}_historical_rates.csv"}
    )


@app.route("/project/<int:pid>/export_json")
def export_project_json(pid):
    conn = get_db()
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not prow:
        conn.close()
        return redirect(url_for("index"))
    project = dict(prow)
    items = [dict(r) for r in conn.execute("SELECT * FROM map_items WHERE project_id=? ORDER BY item_type, item_id", (pid,)).fetchall()]
    conn.close()
    project_export = {field: project.get(field, "") for field in PROJECT_EXPORT_FIELDS}
    project_export["original_id"] = project.get("id")
    item_exports = []
    for item in items:
        item_export = {field: item.get(field, "") for field in MAP_ITEM_EXPORT_FIELDS}
        item_export["original_id"] = item.get("id")
        item_exports.append(item_export)
    payload = {
        "schema_version": 1,
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "project": project_export,
        "items": item_exports,
    }
    filename = f"{safe_project_filename(project)}_fieldwork_export.json"
    return Response(
        json.dumps(payload, indent=2),
        mimetype="application/json",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/project/<int:pid>/update_plan_bar", methods=["POST"])
def update_plan_bar(pid):
    method_name = request.form.get("method_name", "").strip()
    start_date = request.form.get("start_date", "").strip()
    end_date = request.form.get("end_date", "").strip()
    include_saturday = 1 if request.form.get("include_saturday") else 0

    conn = get_db()
    row = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not row:
        conn.close()
        return redirect(url_for("index"))
    project = dict(row)
    method_defs = build_method_definitions(project)
    match = next((m for m in method_defs if m["name"] == method_name), None)
    if match and match.get("start_field") and match.get("end_field"):
        updates = [f'{match["start_field"]}=?', f'{match["end_field"]}=?']
        params = [start_date, end_date]
        if match.get("sat_field"):
            updates.append(f'{match["sat_field"]}=?')
            params.append(include_saturday)
        params.append(pid)
        conn.execute(f'UPDATE projects SET {", ".join(updates)} WHERE id=?', params)
        conn.commit()
        sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab="calendar"))
@app.route("/project/<int:pid>/map_data")
def map_data(pid):
    conn = get_db()
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not prow:
        conn.close()
        return jsonify({"items": []})
    project = dict(prow)
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    items = filter_items_by_methods(project_items(conn, pid), build_method_definitions(project))
    conn.close()
    return jsonify({"items": items})

@app.route("/project/<int:pid>/import_kml", methods=["POST"])
def import_kml(pid):
    upload = request.files.get("kml_file")
    if not upload or not upload.filename:
        return redirect(url_for("project", pid=pid, tab="map"))
    try:
        parsed_items = parse_kml_bytes(extract_kml_bytes(upload))
    except (ValueError, ET.ParseError, zipfile.BadZipFile) as exc:
        return redirect(url_for("project", pid=pid, tab="map", error=f"Import failed: {str(exc)[:180]}"))
    if not parsed_items:
        return redirect(url_for("project", pid=pid, tab="map", error="Import found no point or line features, so existing map items were kept."))
    conn = get_db()
    existing = [dict(r) for r in conn.execute("SELECT * FROM map_items WHERE project_id=?", (pid,)).fetchall()]
    conn.execute(
        "INSERT INTO import_backups (project_id, created_at, item_count, backup_json) VALUES (?, ?, ?, ?)",
        (pid, datetime.now().isoformat(timespec="seconds"), len(existing), json.dumps(existing))
    )
    conn.execute("DELETE FROM map_items WHERE project_id=?", (pid,))
    for item in parsed_items:
        conn.execute("""INSERT INTO map_items
        (project_id,item_type,item_id,geometry_type,coords_json,location_plan,planned_amount,status,work_start_date,work_end_date,notes,depth_m)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (pid, item["item_type"], item["item_id"], item["geometry_type"], item["coords_json"], item["location_plan"], item["planned_amount"], item["status"], item["work_start_date"], item["work_end_date"], item["notes"], 0.0))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab="map", notice=f"Imported {len(parsed_items)} map items. Previous map data was backed up."))

@app.route("/project/<int:pid>/restore_import_backup/<int:backup_id>", methods=["POST"])
def restore_import_backup(pid, backup_id):
    conn = get_db()
    backup = conn.execute(
        "SELECT * FROM import_backups WHERE id=? AND project_id=?",
        (backup_id, pid)
    ).fetchone()
    if not backup:
        conn.close()
        return redirect(url_for("project", pid=pid, tab="map", error="Backup not found."))
    current = [dict(r) for r in conn.execute("SELECT * FROM map_items WHERE project_id=?", (pid,)).fetchall()]
    conn.execute(
        "INSERT INTO import_backups (project_id, created_at, item_count, backup_json) VALUES (?, ?, ?, ?)",
        (pid, datetime.now().isoformat(timespec="seconds"), len(current), json.dumps(current))
    )
    try:
        rows = json.loads(backup["backup_json"])
        if not isinstance(rows, list):
            rows = []
    except Exception:
        rows = []
    conn.execute("DELETE FROM map_items WHERE project_id=?", (pid,))
    for row in rows:
        conn.execute("""INSERT INTO map_items
        (project_id,item_type,item_id,geometry_type,coords_json,location_plan,planned_amount,status,work_start_date,work_end_date,notes,depth_m,exclude_from_history)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
            pid,
            row.get("item_type", "Borehole"),
            row.get("item_id", ""),
            row.get("geometry_type", "Point"),
            row.get("coords_json", json.dumps([0.0, 0.0])),
            row.get("location_plan", ""),
            safe_float(row.get("planned_amount")),
            clean_status(row.get("status")),
            row.get("work_start_date", ""),
            row.get("work_end_date", ""),
            row.get("notes", ""),
            safe_float(row.get("depth_m")),
            int(safe_float(row.get("exclude_from_history"))),
        ))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab="map", notice=f"Restored backup with {len(rows)} items. Current map data was backed up first."))

@app.route("/project/<int:pid>/add_item", methods=["POST"])
def add_item(pid):
    item_type = request.form.get("item_type","").strip() or "Borehole"
    conn = get_db()
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not prow:
        conn.close()
        return redirect(url_for("index"))
    project = dict(prow)
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    geometry_type = geometry_type_for_item_type(project, item_type)
    coords_json = "[]" if geometry_type == "LineString" else json.dumps([0.0, 0.0])
    work_start = request.form.get("work_start_date","").strip()
    work_end = request.form.get("work_end_date","").strip()
    if is_single_day_type(item_type) and work_start and not work_end:
        work_end = work_start
    conn.execute("""INSERT INTO map_items
    (project_id,item_type,item_id,geometry_type,coords_json,location_plan,planned_amount,status,work_start_date,work_end_date,notes,depth_m)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
    (pid, item_type, request.form.get("item_id","").strip(), geometry_type, coords_json, request.form.get("location_plan","").strip(), safe_float(request.form.get("planned_amount")), clean_status(request.form.get("status")), work_start, work_end, request.form.get("notes","").strip(), safe_float(request.form.get("depth_m"))))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab="data"))

@app.route("/project/<int:pid>/mass_add", methods=["POST"])
def mass_add(pid):
    item_type = request.form.get("item_type", "Borehole").strip()
    prefix = request.form.get("prefix","").strip()
    start_no = int(safe_float(request.form.get("start_no")))
    end_no = int(safe_float(request.form.get("end_no")))
    default_planned = safe_float(request.form.get("default_planned_amount"))
    status = clean_status(request.form.get("status"))
    conn = get_db()
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not prow:
        conn.close()
        return redirect(url_for("index"))
    project = dict(prow)
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    geometry_type = geometry_type_for_item_type(project, item_type)
    coords_json = "[]" if geometry_type == "LineString" else json.dumps([0.0, 0.0])
    if prefix and start_no > 0 and end_no >= start_no and (end_no - start_no) <= 1000:
        for i in range(start_no, end_no + 1):
            conn.execute("""INSERT INTO map_items
            (project_id,item_type,item_id,geometry_type,coords_json,location_plan,planned_amount,status,work_start_date,work_end_date,notes,depth_m)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (pid, item_type, f"{prefix}{i:02d}", geometry_type, coords_json, "", default_planned, status, "", "", "", 0.0))
        conn.commit()
        sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab="data"))

@app.route("/item/<int:item_id>/update", methods=["POST"])
def update_item(item_id):
    conn = get_db()
    row = conn.execute("SELECT project_id, item_type FROM map_items WHERE id=?", (item_id,)).fetchone()
    if not row:
        conn.close(); return redirect(url_for("index"))
    pid = row["project_id"]
    prow = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    project = dict(prow) if prow else {}
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    item_type = request.form.get("item_type","").strip() or None
    work_start = request.form.get("work_start_date","").strip()
    work_end = request.form.get("work_end_date","").strip()
    if is_single_day_type(item_type or row["item_type"]) and work_start and not work_end:
        work_end = work_start
    if item_type:
        geometry_type = geometry_type_for_item_type(project, item_type)
        conn.execute("""UPDATE map_items SET
        item_type=?, geometry_type=?, item_id=?, location_plan=?, planned_amount=?, status=?, work_start_date=?, work_end_date=?, notes=?, depth_m=? WHERE id=?""",
        (item_type, geometry_type, request.form.get("item_id","").strip(), request.form.get("location_plan","").strip(), safe_float(request.form.get("planned_amount")), clean_status(request.form.get("status")), work_start, work_end, request.form.get("notes","").strip(), safe_float(request.form.get("depth_m")), item_id))
    else:
        conn.execute("""UPDATE map_items SET
        item_id=?, location_plan=?, planned_amount=?, status=?, work_start_date=?, work_end_date=?, notes=?, depth_m=? WHERE id=?""",
        (request.form.get("item_id","").strip(), request.form.get("location_plan","").strip(), safe_float(request.form.get("planned_amount")), clean_status(request.form.get("status")), work_start, work_end, request.form.get("notes","").strip(), safe_float(request.form.get("depth_m")), item_id))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab=request.form.get("next_tab","map"), year=request.form.get("year"), month=request.form.get("month")))

@app.route("/item/<int:item_id>/quick_update", methods=["POST"])
def quick_update(item_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM map_items WHERE id=?", (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({"ok": False}), 404
    field = request.form.get("field", "").strip()
    value = request.form.get("value", "").strip()
    allowed = {"item_type","item_id","location_plan","status","work_start_date","work_end_date","planned_amount","depth_m"}
    if field not in allowed:
        conn.close()
        return jsonify({"ok": False, "error": "Invalid field"}), 400
    if field in {"planned_amount","depth_m"}:
        value = safe_float(value)
    if field == "status":
        value = clean_status(value)
    if field == "item_type":
        prow = conn.execute("SELECT * FROM projects WHERE id=?", (row["project_id"],)).fetchone()
        project = dict(prow) if prow else {}
        project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
        geometry_type = geometry_type_for_item_type(project, value)
        conn.execute("UPDATE map_items SET item_type=?, geometry_type=? WHERE id=?", (value, geometry_type, item_id))
    else:
        conn.execute(f"UPDATE map_items SET {field}=? WHERE id=?", (value, item_id))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return jsonify({"ok": True})


@app.route("/item/<int:item_id>/adjust_dates", methods=["POST"])
def adjust_dates(item_id):
    conn = get_db()
    row = conn.execute("SELECT id, project_id, item_type, work_start_date, work_end_date FROM map_items WHERE id=?", (item_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({"ok": False, "error": "Not found"}), 404
    start = parse_date(row["work_start_date"])
    end = parse_date(row["work_end_date"] or row["work_start_date"])
    if not start:
        conn.close()
        return jsonify({"ok": False, "error": "No start date"}), 400
    mode = request.form.get("mode", "move").strip()
    delta = int(request.form.get("delta", "0") or 0)
    if mode == "move":
        start = start + timedelta(days=delta)
        end = (end or start) + timedelta(days=delta)
    elif mode == "resize_start":
        start = start + timedelta(days=delta)
        if end and start > end:
            start = end
    elif mode == "resize_end":
        end = (end or start) + timedelta(days=delta)
        if end < start:
            end = start
    if row["item_type"] == "Test Pit":
        end = start
    conn.execute("UPDATE map_items SET work_start_date=?, work_end_date=? WHERE id=?", (fmt_date(start), fmt_date(end), item_id))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return jsonify({"ok": True, "work_start_date": fmt_date(start), "work_end_date": fmt_date(end)})

@app.route("/item/<int:item_id>/toggle_history_exclusion", methods=["POST"])
def toggle_history_exclusion(item_id):
    conn = get_db()
    row = conn.execute("SELECT project_id, exclude_from_history FROM map_items WHERE id=?", (item_id,)).fetchone()
    if not row:
        conn.close()
        return redirect(url_for("index"))
    new_value = 0 if int(safe_float(row["exclude_from_history"])) else 1
    conn.execute("UPDATE map_items SET exclude_from_history=? WHERE id=?", (new_value, item_id))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=row["project_id"], tab="data"))


@app.route("/project/<int:pid>/bulk_update", methods=["POST"])
def bulk_update(pid):
    selected = request.form.getlist("selected_ids")
    action = request.form.get("bulk_action","").strip()
    new_status = clean_status(request.form.get("bulk_status"), default="")
    conn = get_db()
    if selected:
        placeholders = ",".join("?" for _ in selected)
        if action == "delete":
            conn.execute(f"DELETE FROM map_items WHERE project_id=? AND id IN ({placeholders})", [pid] + selected)
        elif action == "status" and new_status:
            conn.execute(f"UPDATE map_items SET status=? WHERE project_id=? AND id IN ({placeholders})", [new_status, pid] + selected)
        conn.commit()
        sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("project", pid=pid, tab="data"))


@app.route("/project/<int:pid>/bulk_map_status", methods=["POST"])
def bulk_map_status(pid):
    ids = request.form.get("selected_ids_json", "[]")
    try:
        selected = json.loads(ids)
    except Exception:
        selected = []
    status = clean_status(request.form.get("status"), default="")
    work_start = request.form.get("work_start_date", "").strip()
    work_end = request.form.get("work_end_date", "").strip()

    if selected:
        conn = get_db()
        placeholders = ",".join("?" for _ in selected)
        rows = conn.execute(
            f"SELECT id, item_type FROM map_items WHERE project_id=? AND id IN ({placeholders})",
            [pid] + selected
        ).fetchall()

        for row in rows:
            this_end = work_end
            if row["item_type"] == "Test Pit" and work_start and not this_end:
                this_end = work_start

            if status in ITEM_STATUSES:
                conn.execute(
                    "UPDATE map_items SET status=?, work_start_date=?, work_end_date=? WHERE project_id=? AND id=?",
                    (status, work_start, this_end, pid, row["id"])
                )
            else:
                conn.execute(
                    "UPDATE map_items SET work_start_date=?, work_end_date=? WHERE project_id=? AND id=?",
                    (work_start, this_end, pid, row["id"])
                )
        conn.commit()
        sync_historical_rates(conn)
        conn.close()

    return redirect(url_for("project", pid=pid, tab="map"))



@app.route("/edit_project/<int:pid>", methods=["GET","POST"])
def edit_project(pid):
    conn = get_db()
    row = conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not row:
        conn.close(); return redirect(url_for("index"))
    if request.method == "POST":
        custom_methods = []
        extra_names = request.form.getlist("custom_method_name")
        extra_colors = request.form.getlist("custom_method_color")
        extra_symbols = request.form.getlist("custom_method_symbol")
        for i, name in enumerate(extra_names):
            name = (name or "").strip()
            if name:
                custom_methods.append({"name": name, "color": extra_colors[i] if i < len(extra_colors) else "#0f766e", "symbol": extra_symbols[i] if i < len(extra_symbols) else "circle"})
        conn.execute("""UPDATE projects SET
        name=?, project_number=?, task_code=?, client=?, site_location=?,
        borehole_start_date=?, borehole_end_date=?, borehole_include_saturday=?,
        cptu_start_date=?, cptu_end_date=?, cptu_include_saturday=?,
        test_pit_start_date=?, test_pit_end_date=?, test_pit_include_saturday=?,
        geophysics_start_date=?, geophysics_end_date=?, geophysics_include_saturday=?,
        custom_methods_json=?, use_borehole=?, use_cptu=?, use_test_pit=?, use_geophysics=?, borehole_budget_meters=?, cptu_budget_meters=?, geophysics_budget_meters=?
        WHERE id=?""",
        (
            request.form.get("name","").strip(), request.form.get("project_number","").strip(), request.form.get("task_code","").strip(),
            request.form.get("client","").strip(), request.form.get("site_location","").strip(),
            request.form.get("borehole_start_date","").strip(), request.form.get("borehole_end_date","").strip(), 1 if request.form.get("borehole_include_saturday") else 0,
            request.form.get("cptu_start_date","").strip(), request.form.get("cptu_end_date","").strip(), 1 if request.form.get("cptu_include_saturday") else 0,
            request.form.get("test_pit_start_date","").strip(), request.form.get("test_pit_end_date","").strip(), 1 if request.form.get("test_pit_include_saturday") else 0,
            request.form.get("geophysics_start_date","").strip(), request.form.get("geophysics_end_date","").strip(), 1 if request.form.get("geophysics_include_saturday") else 0,
            json.dumps(custom_methods),
            1 if request.form.get("use_borehole") else 0,
            1 if request.form.get("use_cptu") else 0,
            1 if request.form.get("use_test_pit") else 0,
            1 if request.form.get("use_geophysics") else 0,
            safe_float(request.form.get("borehole_budget_meters")),
            safe_float(request.form.get("cptu_budget_meters")),
            safe_float(request.form.get("geophysics_budget_meters")),
            pid
        ))
        conn.commit()
        sync_historical_rates(conn)
        conn.close()
        return redirect(url_for("project", pid=pid, tab="overview"))
    project = dict(row)
    project["custom_methods"] = parse_custom_methods(project.get("custom_methods_json", "[]"))
    conn.close()
    return render_template("edit_project.html", project=project)


@app.route("/delete_project/<int:pid>", methods=["POST"])
def delete_project(pid):
    conn = get_db()
    conn.execute("DELETE FROM map_items WHERE project_id=?", (pid,))
    conn.execute("DELETE FROM import_backups WHERE project_id=?", (pid,))
    conn.execute("DELETE FROM historical_rates WHERE project_id=?", (pid,))
    conn.execute("DELETE FROM projects WHERE id=?", (pid,))
    conn.commit()
    sync_historical_rates(conn)
    conn.close()
    return redirect(url_for("index"))

if not use_postgres():
    init_db()

if __name__ == "__main__":
    app.run(debug=False, use_reloader=False)
