"""Albrost Accounts online server: logins, roles, activity log and saving to Excel.

Roles
  owner    everything, including editing, deleting, timestamps and Logs.
  manager  sees everything except Logs and timestamps, and can only ADD new
           daily sales and supplier entries. The rule is checked here on the
           server by comparing each upload with the current file, not just by
           hiding buttons on the page.

Login is off until ALBROST_LOGIN=on is set (see START ONLINE.bat). While it is
off, anyone who opens the link gets the owner role, so the link must stay private.

The Excel file, backups, logs and password hashes are local only and are never
pushed to GitHub (see .gitignore). Add or change a login with set_password.py.
"""
import hashlib
import hmac
import io
import json
import os
import re
import secrets
import shutil
import subprocess
import threading
import time
import zipfile
from collections import Counter
from datetime import date, datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

import openpyxl

CODE = Path(__file__).resolve().parent
APP = CODE / "app"
WEB = CODE / "web"
# ALBROST_ROOT / ALBROST_PORT let a test run against a copy of the data.
ROOT = Path(os.environ.get("ALBROST_ROOT", CODE))
PORT = int(os.environ.get("ALBROST_PORT", "8766"))
WORKBOOK = ROOT / "Al_Brost_Finance_Supplier_Balances.xlsx"
BACKUPS = ROOT / "Backups"
DAILY_BACKUPS = BACKUPS / "Daily"
DATA = ROOT / "data"
USERS_FILE = DATA / "users.json"
AUDIT_FILE = DATA / "audit.jsonl"
BASELINE = DATA / "last_known.xlsx"

LOGIN_REQUIRED = os.environ.get("ALBROST_LOGIN", "off").lower() == "on"
OPEN_SESSION = {"username": "", "name": "No login", "role": "owner"}

MAX_UPLOAD = 10 * 1024 * 1024
KEEP_BACKUPS = 60
SESSION_SECONDS = 7 * 24 * 3600
MAX_FAILED_LOGINS = 10
LOCKOUT_SECONDS = 15 * 60
GIT_SYNC_SECONDS = 60

FILE_LOCK = threading.RLock()
AUDIT_LOCK = threading.Lock()
AUTH_LOCK = threading.Lock()
SESSIONS = {}
FAILED_LOGINS = {}

STATIC = {
    "/exceljs.min.js": (APP / "exceljs.min.js", "application/javascript"),
    "/chart.umd.js": (APP / "chart.umd.js", "application/javascript"),
}

# Sheet1 layout: see the comment at the top of the script in app/index.html.
DAILY_INPUTS = {"online": 2, "cash": 3, "oexp": 5, "cexp": 6, "saddam": 7, "arshad": 8}
DAILY_LABELS = {"online": "online", "cash": "cash", "oexp": "online expense",
                "cexp": "cash expense", "saddam": "Saddam", "arshad": "Arshad"}
SUPPLIER_INPUT_COLS = (17, 18, 19, 20)  # Q bill, R paid, S supplier, T item
OVERVIEW_FORMULA = re.compile(r"^=SUM\(([A-Z])2:\1\d+\)$")


# ---------- passwords and sessions ----------

def hash_password(password, salt=None):
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=2**14, r=8, p=1)
    return salt.hex(), digest.hex()


def load_users():
    return json.loads(USERS_FILE.read_text(encoding="utf-8"))


def check_password(username, password):
    """Returns the user record, or None. Takes the same time for unknown names."""
    users = load_users()
    user = users.get(username)
    salt = bytes.fromhex(user["salt"]) if user else b"0" * 16
    _, digest = hash_password(password, salt)
    if user and hmac.compare_digest(digest, user["hash"]):
        return {"username": username, "name": user["name"], "role": user["role"]}
    return None


def too_many_failures(ip):
    with AUTH_LOCK:
        fails, since = FAILED_LOGINS.get(ip, (0, 0))
        if time.time() - since > LOCKOUT_SECONDS:
            FAILED_LOGINS.pop(ip, None)
            return False
        return fails >= MAX_FAILED_LOGINS


def note_failure(ip):
    with AUTH_LOCK:
        fails, since = FAILED_LOGINS.get(ip, (0, time.time()))
        FAILED_LOGINS[ip] = (fails + 1, since)


def new_session(user):
    token = secrets.token_urlsafe(32)
    with AUTH_LOCK:
        SESSIONS[token] = {**user, "expires": time.time() + SESSION_SECONDS}
    return token


def get_session(token):
    with AUTH_LOCK:
        session = SESSIONS.get(token)
        if session and session["expires"] < time.time():
            del SESSIONS[token]
            return None
        return session


# ---------- activity log ----------

def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def audit(user, action, details="", sales=(), supp=()):
    entry = {"at": now_iso(), "user": user, "action": action, "details": details,
             "sales": list(sales), "supp": list(supp)}
    with AUDIT_LOCK:
        DATA.mkdir(parents=True, exist_ok=True)
        with open(AUDIT_FILE, "a", encoding="utf-8") as log:
            log.write(json.dumps(entry, ensure_ascii=False) + "\n")


def read_audit():
    if not AUDIT_FILE.exists():
        return []
    with AUDIT_LOCK:
        lines = AUDIT_FILE.read_text(encoding="utf-8").splitlines()
    return [json.loads(line) for line in lines if line.strip()]


def last_updated():
    """When each daily sale and supplier entry last changed, replayed from the log."""
    sales, supp = {}, {}
    for entry in read_audit():
        stamp = {"at": entry["at"], "by": entry["user"]}
        for key in entry.get("sales", []):
            sales[key] = stamp
        for key in entry.get("supp", []):
            supp[key] = stamp
    return {"sales": sales, "supp": supp}


# ---------- reading the workbook ----------

def open_book(data):
    return openpyxl.load_workbook(io.BytesIO(data), data_only=False)


def is_workbook(data):
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if "xl/workbook.xml" not in archive.namelist():
                return False
        return "Sheet1" in open_book(data).sheetnames
    except Exception:
        return False


def plain(value):
    """Comparable form of a cell value, so a re-saved file doesn't look changed."""
    value = getattr(value, "text", value)  # array formulas
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, str):
        return value.strip() or None
    return value


def to_number(value):
    value = plain(value)
    if isinstance(value, float):
        return value
    try:
        return float(str(value).replace(",", "").replace("₹", "").strip())
    except ValueError:
        return 0.0


def to_ymd(value):
    value = plain(value)
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return ""


def number_text(n):
    return str(int(n)) if float(n).is_integer() else str(n)


def pretty(ymd):
    if not ymd:
        return "no date"
    return datetime.strptime(ymd, "%Y-%m-%d").strftime("%d %b %Y").lstrip("0")


def read_records(book):
    sheet = book["Sheet1"]
    daily, supplier = {}, Counter()
    for r in range(2, sheet.max_row + 1):
        day = to_ymd(sheet.cell(r, 1).value)
        if day and any(plain(sheet.cell(r, c).value) is not None for c in DAILY_INPUTS.values()):
            daily[day] = {k: to_number(sheet.cell(r, c).value) for k, c in DAILY_INPUTS.items()}
        name = plain(sheet.cell(r, 19).value)
        bill, paid = plain(sheet.cell(r, 17).value), plain(sheet.cell(r, 18).value)
        if name is not None or bill is not None or paid is not None:
            item = plain(sheet.cell(r, 20).value)
            key = "|".join([to_ymd(sheet.cell(r, 16).value), str(name or "(no name)"), str(item or ""),
                            number_text(to_number(bill)), number_text(to_number(paid))])
            supplier[key] += 1
    return daily, supplier


def describe_supplier(key):
    day, name, item, bill, paid = key.split("|")
    item_text = f" ({item})" if item else ""
    return f"{name}{item_text}, {pretty(day)}: bill ₹{bill}, paid ₹{paid}"


def describe_sale(values):
    return ", ".join(f"{DAILY_LABELS[k]} ₹{number_text(v)}" for k, v in values.items())


def log_differences(user, old_book, new_book):
    old_daily, old_supp = read_records(old_book)
    new_daily, new_supp = read_records(new_book)
    changed = False
    for day in sorted(set(old_daily) | set(new_daily)):
        before, after = old_daily.get(day), new_daily.get(day)
        if before == after:
            continue
        changed = True
        if before is None:
            audit(user, "Added daily sale", f"{pretty(day)}: {describe_sale(after)}", sales=[day])
        elif after is None:
            audit(user, "Deleted daily sale", f"{pretty(day)}: {describe_sale(before)}", sales=[day])
        else:
            audit(user, "Edited daily sale",
                  f"{pretty(day)}: was {describe_sale(before)}; now {describe_sale(after)}", sales=[day])
    removed = list((old_supp - new_supp).elements())
    added = list((new_supp - old_supp).elements())
    for before, after in zip(removed, added):
        changed = True
        audit(user, "Edited supplier entry",
              f"was {describe_supplier(before)}; now {describe_supplier(after)}", supp=[after])
    for key in removed[len(added):]:
        changed = True
        audit(user, "Deleted supplier entry", describe_supplier(key), supp=[key])
    for key in added[len(removed):]:
        changed = True
        audit(user, "Added supplier entry", describe_supplier(key), supp=[key])
    return changed


def manager_may_save(old_book, new_book):
    """A manager may only fill rows that are still empty, never change existing data."""
    if old_book.sheetnames != new_book.sheetnames:
        return False
    for name in old_book.sheetnames:
        old_sheet, new_sheet = old_book[name], new_book[name]
        rows = max(old_sheet.max_row, new_sheet.max_row)
        cols = max(old_sheet.max_column, new_sheet.max_column)
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                before = plain(old_sheet.cell(r, c).value)
                after = plain(new_sheet.cell(r, c).value)
                if before == after:
                    continue
                if name != "Sheet1" or not manager_cell_allowed(old_sheet, r, c, before, after):
                    return False
    return True


def manager_cell_allowed(sheet, r, c, before, after):
    if r == 1:
        return False
    if 1 <= c <= 10:  # daily sale columns
        if any(plain(sheet.cell(r, col).value) is not None for col in DAILY_INPUTS.values()):
            return False
        return c != 1 or before is None
    if 16 <= c <= 21:  # supplier columns
        return all(plain(sheet.cell(r, col).value) is None for col in SUPPLIER_INPUT_COLS)
    if c == 14 and 2 <= r <= 6:  # overview totals grow as new rows are added
        old_match = OVERVIEW_FORMULA.match(str(before))
        new_match = OVERVIEW_FORMULA.match(str(after))
        return bool(old_match and new_match and old_match.group(1) == new_match.group(1))
    return False


# ---------- saving the workbook ----------

def back_up():
    DAILY_BACKUPS.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    shutil.copy2(WORKBOOK, BACKUPS / f"Backup_{now:%Y-%m-%d_%H-%M-%S}.xlsx")
    for old in sorted(BACKUPS.glob("Backup_*.xlsx"), reverse=True)[KEEP_BACKUPS:]:
        old.unlink(missing_ok=True)
    # The day's first version is never pruned, so a burst of bad saves can't
    # push every good copy out of the rolling backups above.
    daily = DAILY_BACKUPS / f"Daily_{now:%Y-%m-%d}.xlsx"
    if not daily.exists():
        shutil.copy2(WORKBOOK, daily)


def note_outside_changes():
    """Logs edits made in Excel or the shop PC app, which skip this server."""
    with FILE_LOCK:
        current = WORKBOOK.read_bytes()
        if not BASELINE.exists():
            DATA.mkdir(parents=True, exist_ok=True)
            BASELINE.write_bytes(current)
            return current
        baseline = BASELINE.read_bytes()
        if hashlib.sha256(current).digest() != hashlib.sha256(baseline).digest():
            try:
                log_differences("Shop PC / Excel", open_book(baseline), open_book(current))
                BASELINE.write_bytes(current)
            except Exception as error:  # e.g. Excel is partway through saving
                print("Could not compare outside change:", error, flush=True)
        return current


def save_workbook(session, data):
    """Returns an HTTP status code for the save."""
    with FILE_LOCK:
        current = note_outside_changes()
        old_book, new_book = open_book(current), open_book(data)
        if session["role"] != "owner" and not manager_may_save(old_book, new_book):
            audit(session["name"], "Blocked save", "Tried to change existing data")
            return 403
        temp = ROOT / "~saving.tmp"
        try:
            # Fails while Excel has the file open; the page explains this on 423.
            with open(WORKBOOK, "r+b"):
                pass
            back_up()
            temp.write_bytes(data)
            os.replace(temp, WORKBOOK)
        except PermissionError:
            return 423
        finally:
            temp.unlink(missing_ok=True)
        BASELINE.write_bytes(data)
        if not log_differences(session["name"], old_book, new_book):
            audit(session["name"], "Saved", "No entries changed")
    return 200


# ---------- pages ----------

def main_page(session):
    html = (APP / "index.html").read_text(encoding="utf-8")
    me = json.dumps({"name": session["name"], "role": session["role"], "login": LOGIN_REQUIRED}).replace("<", "\\u003c")
    css = (WEB / "roles.css").read_text(encoding="utf-8")
    script = (WEB / "roles.js").read_text(encoding="utf-8")
    html = html.replace("</head>", f"<style>{css}</style></head>", 1)
    # After the app's own script, so it can wrap the app's functions before the first render.
    extra = f"<script>window.ALBROST_USER={me};</script><script>{script}</script></body>"
    return html.replace("</body>", extra, 1).encode("utf-8")


class Handler(BaseHTTPRequestHandler):
    server_version = "Albrost"
    sys_version = ""
    protocol_version = "HTTP/1.1"

    # ----- helpers -----
    def send_body(self, code, body, content_type, headers=()):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Robots-Tag", "noindex, nofollow")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        for name, value in headers:
            self.send_header(name, value)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def send_json(self, code, obj, headers=()):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_body(code, body, "application/json; charset=utf-8", headers)

    def redirect(self, where):
        self.send_body(303, b"", "text/plain", [("Location", where)])

    def client_ip(self):
        return self.headers.get("CF-Connecting-IP") or self.client_address[0]

    def token(self):
        for part in (self.headers.get("Cookie") or "").split(";"):
            name, _, value = part.strip().partition("=")
            if name == "albrost_session":
                return value
        return ""

    def session(self):
        return get_session(self.token()) if LOGIN_REQUIRED else OPEN_SESSION

    def cookie(self, value, max_age):
        secure = "; Secure" if self.headers.get("X-Forwarded-Proto") == "https" else ""
        return ("Set-Cookie", f"albrost_session={value}; Path=/; Max-Age={max_age}; "
                              f"HttpOnly; SameSite=Lax{secure}")

    def same_origin(self):
        origin = self.headers.get("Origin")
        return not origin or urlsplit(origin).netloc == self.headers.get("Host")

    def read_body(self, limit):
        try:
            length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            length = -1
        if not 0 < length <= limit:
            self.close_connection = True
            return None
        return self.rfile.read(length)

    # ----- GET -----
    def do_GET(self):
        path = urlsplit(self.path).path
        if path == "/logo.jpg":
            self.send_body(200, (APP / "logo.jpg").read_bytes(), "image/jpeg")
            return
        session = self.session()
        if path == "/login":
            if session:
                self.redirect("/")
            else:
                self.send_body(200, (WEB / "login.html").read_bytes(), "text/html; charset=utf-8")
            return
        if not session:
            if path.startswith("/api/"):
                self.send_json(401, {"error": "login"})
            else:
                self.redirect("/login")
            return
        if path in ("/", "/index.html"):
            self.send_body(200, main_page(session), "text/html; charset=utf-8")
        elif path in STATIC:
            file, content_type = STATIC[path]
            self.send_body(200, file.read_bytes(), content_type)
        elif path == "/api/file":
            self.send_body(200, note_outside_changes(), "application/octet-stream")
        elif path == "/api/ping":
            self.send_json(200, {"ok": True})
        elif path in ("/api/logs", "/api/times"):
            if session["role"] != "owner":
                self.send_json(403, {"error": "owner only"})
                return
            note_outside_changes()
            if path == "/api/logs":
                entries = [{k: e.get(k, "") for k in ("at", "user", "action", "details")}
                           for e in read_audit()]
                self.send_json(200, entries[::-1])
            else:
                self.send_json(200, last_updated())
        else:
            self.send_json(404, {"error": "not found"})

    do_HEAD = do_GET

    # ----- POST -----
    def do_POST(self):
        path = urlsplit(self.path).path
        if not self.same_origin():
            self.close_connection = True
            self.send_json(403, {"error": "origin"})
        elif LOGIN_REQUIRED and path == "/api/login":
            self.handle_login()
        elif LOGIN_REQUIRED and path == "/api/logout":
            self.handle_logout()
        elif path == "/api/save":
            self.handle_save()
        else:
            self.close_connection = True
            self.send_json(403, {"error": "not allowed"})

    def handle_login(self):
        ip = self.client_ip()
        if too_many_failures(ip):
            self.close_connection = True
            self.send_json(429, {"error": "too many tries"})
            return
        raw = self.read_body(4096)
        try:
            form = json.loads(raw or b"{}")
            username, password = str(form["username"]).strip(), str(form["password"])
        except (ValueError, KeyError, TypeError):
            self.send_json(400, {"error": "bad request"})
            return
        user = check_password(username, password)
        if not user:
            note_failure(ip)
            shown = username if username in load_users() else "unknown name"
            audit(shown, "Failed login")
            self.send_json(401, {"error": "wrong"})
            return
        FAILED_LOGINS.pop(ip, None)
        audit(user["name"], "Logged in")
        cookie = self.cookie(new_session(user), SESSION_SECONDS)
        self.send_json(200, {"ok": True}, [cookie])

    def handle_logout(self):
        session = self.session()
        if session:
            with AUTH_LOCK:
                SESSIONS.pop(self.token(), None)
            audit(session["name"], "Logged out")
        self.close_connection = True
        self.send_json(200, {"ok": True}, [self.cookie("", 0)])

    def handle_save(self):
        session = self.session()
        if not session:
            self.close_connection = True
            self.send_json(401, {"error": "login"})
            return
        data = self.read_body(MAX_UPLOAD)
        if data is None:
            self.send_json(413, {"error": "bad size"})
            return
        if not is_workbook(data):
            self.send_json(400, {"error": "not a workbook"})
            return
        code = save_workbook(session, data)
        self.send_json(code, {"ok": code == 200})

    def refuse(self):
        self.close_connection = True
        self.send_json(403, {"error": "not allowed"})

    do_PUT = do_PATCH = do_DELETE = refuse

    def log_message(self, *args):
        pass


# ---------- GitHub auto-update ----------

def git(*args):
    return subprocess.run(["git", *args], cwd=CODE, capture_output=True, text=True, timeout=120)


def sync_code_to_github():
    """Commits and pushes code changes. .gitignore only lets listed code files through."""
    while True:
        time.sleep(GIT_SYNC_SECONDS)
        try:
            git("add", "-A")
            if git("diff", "--cached", "--quiet").returncode != 0:
                git("commit", "-m", f"Auto-update {datetime.now():%Y-%m-%d %H:%M}")
            status = git("status", "-sb").stdout.splitlines()
            if status and "ahead" in status[0]:
                result = git("push", "origin", "HEAD")
                if result.returncode != 0:
                    print("GitHub push failed:", result.stderr, flush=True)
        except Exception as error:
            print("GitHub update failed:", error, flush=True)


if __name__ == "__main__":
    if LOGIN_REQUIRED and not USERS_FILE.exists():
        raise SystemExit("No logins yet. Create one first: python set_password.py --help")
    if "ALBROST_ROOT" not in os.environ and (CODE / ".git").exists():
        threading.Thread(target=sync_code_to_github, daemon=True).start()
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
