#!/usr/bin/env python3
"""
Vaani CRM v1.0
==============
A personal CRM for the leads Vaani scrapes - just for you, on your own
computer. Excel was never built for follow-ups, remarks and daily call
lists; this is.

It uses the SAME database as Vaani (vaani_master.db, kept next to this
file), so every lead Vaani has ever scraped is already here - nothing
to import by hand. Your tracker outcomes, notes and callback dates are
pulled in automatically the first time it runs.

    py vaani_crm.py

Opens in your browser at http://127.0.0.1:8765 - nothing leaves your
computer, no account, no internet needed, no cost.
"""
import html
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import urllib.parse
import urllib.request
import webbrowser
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "vaani_master.db")
TRACKER_PATH = os.path.join(HERE, "vaani-call-tracker.xlsx")
BACKUP_PATH = os.path.join(HERE, "vaani_master.db.crm-backup")

VERSION = "1.0.0"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/vaani_crm.py")
PORTS = (8765, 8766, 8767, 8768, 8769)

STATUSES = ["Not called", "No answer", "Staff reached",
            "Decision-maker reached", "Conversation", "Interested",
            "Demo booked", "Demo done", "Deck sent", "Callback",
            "Not interested", "Wrong number", "Do not contact",
            "Won", "Lost"]
SUPPRESS_STATUSES = {"Not interested", "Wrong number", "Do not contact"}
TERMINAL_STATUSES = {"Won", "Lost", "Not interested", "Wrong number",
                     "Do not contact"}


# ---------------------------------------------------------------- auto-update

def _version_tuple(text):
    try:
        return tuple(int(x) for x in text.strip().split("."))
    except Exception:
        return (0,)


def self_update():
    """Same deal as Vaani: check online for a newer vaani_crm.py, swap
    it in and restart. Silent when offline or up to date. Only the
    program file is downloaded - your data never leaves this PC."""
    if os.environ.get("VAANI_CRM_NO_UPDATE") == "1" \
            or "--no-update" in sys.argv:
        return
    try:
        req = urllib.request.Request(UPDATE_URL,
                                     headers={"User-Agent": "VaaniCRM"})
        with urllib.request.urlopen(req, timeout=5) as r:
            content = r.read().decode("utf-8", "replace")
    except Exception:
        return
    m = re.search(r'^VERSION\s*=\s*"([0-9.]+)"', content, re.M)
    if not m or _version_tuple(m.group(1)) <= _version_tuple(VERSION):
        return
    if "Vaani CRM" not in content or len(content) < 10000:
        return
    try:
        compile(content, "vaani_crm.py", "exec")
    except Exception:
        return
    me = os.path.abspath(__file__)
    try:
        shutil.copyfile(me, me + ".bak")
        with open(me, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        return
    print(f"\n  Updated Vaani CRM to v{m.group(1)} - restarting...\n")
    env = dict(os.environ)
    env["VAANI_CRM_NO_UPDATE"] = "1"
    subprocess.Popen([sys.executable, me] + sys.argv[1:], env=env)
    sys.exit(0)


# ---------------------------------------------------------------- database

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("""CREATE TABLE IF NOT EXISTS leads (
        phone TEXT PRIMARY KEY, name TEXT, address TEXT, website TEXT,
        email TEXT, rating REAL, reviews INTEGER, score INTEGER,
        tier TEXT, angle TEXT, first_seen TEXT, last_seen TEXT,
        times_seen INTEGER DEFAULT 1, suppressed INTEGER DEFAULT 0,
        suppress_reason TEXT, suppressed_on TEXT)""")
    cols = [r[1] for r in conn.execute("PRAGMA table_info(leads)")]
    if "suppressed_on" not in cols:
        conn.execute("ALTER TABLE leads ADD COLUMN suppressed_on TEXT")
    conn.execute("""CREATE TABLE IF NOT EXISTS crm_state (
        phone TEXT PRIMARY KEY, status TEXT DEFAULT 'Not called',
        followup TEXT DEFAULT '', updated_at TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS crm_remarks (
        id INTEGER PRIMARY KEY AUTOINCREMENT, phone TEXT, ts TEXT,
        text TEXT, source TEXT DEFAULT 'app')""")
    conn.execute("""CREATE TABLE IF NOT EXISTS crm_meta (
        key TEXT PRIMARY KEY, value TEXT)""")
    conn.commit()
    return conn


def backup_db():
    """One rolling backup of the whole lead database, refreshed daily."""
    try:
        if not os.path.exists(DB_PATH):
            return
        if os.path.exists(BACKUP_PATH):
            age = datetime.now() - datetime.fromtimestamp(
                os.path.getmtime(BACKUP_PATH))
            if age < timedelta(hours=20):
                return
        shutil.copyfile(DB_PATH, BACKUP_PATH)
    except Exception:
        pass


def _phone_keys(phone_field):
    keys = set()
    for part in re.split(r"[,;/|]", str(phone_field or "")):
        d = re.sub(r"\D", "", part)
        if len(d) >= 10:
            keys.add(d[-10:])
    return keys


def _parse_date(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d %b %Y",
                "%d-%b-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except Exception:
            continue
    return ""


# ---------------------------------------------------------------- tracker import

def import_tracker(conn):
    """First run only: pull the user's existing tracker outcomes,
    notes and callback dates into the CRM, so work already done in
    Excel is not lost. The newest sheet's word wins per phone."""
    row = conn.execute(
        "SELECT value FROM crm_meta WHERE key = 'tracker_import_v1'"
    ).fetchone()
    if row:
        return 0
    if not os.path.exists(TRACKER_PATH):
        conn.execute(
            "INSERT OR REPLACE INTO crm_meta VALUES"
            " ('tracker_import_v1', 'no-tracker')")
        conn.commit()
        return 0
    try:
        import openpyxl
    except Exception:
        print("  (tracker import skipped - openpyxl not available)")
        return 0
    latest = {}
    try:
        wb = openpyxl.load_workbook(TRACKER_PATH, read_only=True)
        for ws in wb.worksheets:
            if ws.title == "How to use":
                continue
            headers = {}
            for c in range(1, (ws.max_column or 0) + 1):
                v = ws.cell(row=1, column=c).value
                if v in ("Phone", "Outcome", "Restaurant name",
                         "Location", "Email", "Notes",
                         "Callback date") and v not in headers:
                    headers[v] = c  # first match wins (the summary
                    # block reuses the word "Outcome" further right)
            if "Phone" not in headers or "Outcome" not in headers:
                continue
            for r in range(2, (ws.max_row or 0) + 1):
                outcome = ws.cell(row=r,
                                  column=headers["Outcome"]).value
                keys = _phone_keys(
                    ws.cell(row=r, column=headers["Phone"]).value)
                if not keys:
                    continue
                info = {"outcome": outcome or "Not called"}
                for field, col in (("name", "Restaurant name"),
                                   ("location", "Location"),
                                   ("email", "Email"),
                                   ("notes", "Notes")):
                    if col in headers:
                        info[field] = ws.cell(
                            row=r, column=headers[col]).value or ""
                info["callback"] = ""
                if "Callback date" in headers:
                    info["callback"] = _parse_date(ws.cell(
                        row=r,
                        column=headers["Callback date"]).value)
                for k in keys:
                    latest[k] = info
        wb.close()
    except Exception:
        return 0
    now = datetime.now().isoformat(timespec="seconds")
    today = date.today().isoformat()
    imported = 0
    for phone, info in latest.items():
        exists = conn.execute(
            "SELECT 1 FROM leads WHERE phone = ?", (phone,)).fetchone()
        if not exists:
            conn.execute(
                "INSERT OR IGNORE INTO leads (phone, name, address,"
                " email, first_seen, last_seen, times_seen) VALUES"
                " (?, ?, ?, ?, ?, ?, 1)",
                (phone, info.get("name") or "", info.get("location")
                 or "", info.get("email") or "", today, today))
        status = info["outcome"]
        meaningful = status != "Not called" or info.get("notes") \
            or info.get("callback")
        state = conn.execute(
            "SELECT 1 FROM crm_state WHERE phone = ?",
            (phone,)).fetchone()
        if meaningful and not state:
            conn.execute(
                "INSERT INTO crm_state (phone, status, followup,"
                " updated_at) VALUES (?, ?, ?, ?)",
                (phone, status if status in STATUSES else "Not called",
                 info.get("callback") or "", now))
            imported += 1
        if info.get("notes"):
            already = conn.execute(
                "SELECT 1 FROM crm_remarks WHERE phone = ?"
                " AND source = 'tracker'", (phone,)).fetchone()
            if not already:
                conn.execute(
                    "INSERT INTO crm_remarks (phone, ts, text, source)"
                    " VALUES (?, ?, ?, 'tracker')",
                    (phone, now, str(info["notes"])))
        if status in SUPPRESS_STATUSES:
            conn.execute(
                "UPDATE leads SET suppressed = 1, suppress_reason = ?,"
                " suppressed_on = COALESCE(suppressed_on, ?) WHERE"
                " phone = ?", (status, today, phone))
    conn.execute(
        "INSERT OR REPLACE INTO crm_meta VALUES"
        " ('tracker_import_v1', ?)", (now,))
    conn.commit()
    return imported


# ---------------------------------------------------------------- queries & actions

LEAD_SELECT = """
    SELECT l.*, COALESCE(s.status, 'Not called') AS status,
           COALESCE(s.followup, '') AS followup,
           (SELECT COUNT(*) FROM crm_remarks r
             WHERE r.phone = l.phone) AS n_remarks,
           (SELECT r.text FROM crm_remarks r WHERE r.phone = l.phone
             ORDER BY r.id DESC LIMIT 1) AS last_remark
    FROM leads l LEFT JOIN crm_state s ON s.phone = l.phone
"""


def fetch_leads(conn, q="", tier="", status="", sort="score"):
    sql = LEAD_SELECT + " WHERE 1=1"
    args = []
    if q:
        sql += (" AND (l.name LIKE ? OR l.phone LIKE ?"
                " OR l.address LIKE ? OR l.email LIKE ?)")
        like = f"%{q}%"
        args += [like, like, like, like]
    if tier in ("A", "B", "C"):
        sql += " AND l.tier = ?"
        args.append(tier)
    if status in STATUSES:
        sql += " AND COALESCE(s.status, 'Not called') = ?"
        args.append(status)
    if sort == "followup":
        sql += (" ORDER BY (COALESCE(s.followup, '') = ''),"
                " s.followup, (l.score IS NULL), l.score DESC")
    elif sort == "recent":
        sql += " ORDER BY l.last_seen DESC, l.score DESC"
    elif sort == "name":
        sql += " ORDER BY l.name"
    else:
        sql += (" ORDER BY (l.score IS NULL), l.score DESC,"
                " l.reviews DESC")
    return conn.execute(sql, args).fetchall()


def fetch_lead(conn, phone):
    return conn.execute(LEAD_SELECT + " WHERE l.phone = ?",
                        (phone,)).fetchone()


def due_leads(conn):
    today = date.today().isoformat()
    rows = conn.execute(
        LEAD_SELECT +
        " WHERE COALESCE(s.followup, '') != ''"
        " AND COALESCE(s.status, 'Not called') NOT IN"
        " ('Won', 'Lost', 'Not interested', 'Wrong number',"
        "  'Do not contact')"
        " ORDER BY s.followup, (l.score IS NULL), l.score DESC"
    ).fetchall()
    overdue = [r for r in rows if r["followup"] < today]
    due_today = [r for r in rows if r["followup"] == today]
    upcoming = [r for r in rows if r["followup"] > today]
    return overdue, due_today, upcoming


def set_status(conn, phone, status, followup):
    if status not in STATUSES:
        return False
    followup = _parse_date(followup) if followup else ""
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        "INSERT INTO crm_state (phone, status, followup, updated_at)"
        " VALUES (?, ?, ?, ?) ON CONFLICT(phone) DO UPDATE SET"
        " status = excluded.status, followup = excluded.followup,"
        " updated_at = excluded.updated_at",
        (phone, status, followup, now))
    if status in SUPPRESS_STATUSES:
        conn.execute(
            "UPDATE leads SET suppressed = 1, suppress_reason = ?,"
            " suppressed_on = COALESCE(suppressed_on, ?) WHERE"
            " phone = ?",
            (status, date.today().isoformat(), phone))
    else:
        conn.execute(
            "UPDATE leads SET suppressed = 0, suppress_reason = NULL,"
            " suppressed_on = NULL WHERE phone = ? AND suppressed = 1"
            " AND suppress_reason IN ('Not interested',"
            " 'Wrong number', 'Do not contact', 'marked in tracker')",
            (phone,))
    conn.commit()
    return True


def add_remark(conn, phone, text):
    text = (text or "").strip()
    if not text:
        return False
    conn.execute(
        "INSERT INTO crm_remarks (phone, ts, text, source) VALUES"
        " (?, ?, ?, 'app')",
        (phone, datetime.now().isoformat(timespec="seconds"), text))
    conn.commit()
    return True


def fetch_remarks(conn, phone):
    return conn.execute(
        "SELECT * FROM crm_remarks WHERE phone = ? ORDER BY id DESC",
        (phone,)).fetchall()


# ---------------------------------------------------------------- html helpers

def esc(value):
    return html.escape(str(value if value is not None else ""),
                       quote=True)


def fmt_date(iso_text):
    if not iso_text:
        return ""
    try:
        return date.fromisoformat(str(iso_text)[:10]).strftime(
            "%d %b %Y")
    except Exception:
        return str(iso_text)


def fmt_phone(phone):
    p = str(phone or "")
    if len(p) == 10:
        return f"{p[:5]} {p[5:]}"
    return p


def tier_badge(tier):
    if not tier:
        return '<span class="badge bC">-</span>'
    return f'<span class="badge b{tier}">{esc(tier)}</span>'


def status_pill(status):
    hot = {"Interested", "Demo booked", "Demo done", "Conversation"}
    good = {"Won"}
    bad = {"Not interested", "Lost"}
    dead = {"Wrong number", "Do not contact"}
    if status in good:
        cls = "st-good"
    elif status in hot:
        cls = "st-hot"
    elif status in bad:
        cls = "st-bad"
    elif status in dead:
        cls = "st-dead"
    elif status == "Not called":
        cls = "st-new"
    else:
        cls = "st-mid"
    return f'<span class="pill {cls}">{esc(status)}</span>'


CSS = """
* { box-sizing: border-box; }
body { margin: 0; font-family: 'Segoe UI', Arial, sans-serif;
  background: #F4F6F7; color: #1c2b30; }
header { background: #1F4E5F; color: #fff; padding: 12px 22px;
  display: flex; align-items: center; gap: 22px; flex-wrap: wrap; }
header .brand { font-size: 20px; font-weight: 700; letter-spacing: .3px; }
header nav a { color: #CFE3E9; text-decoration: none; margin-right: 16px;
  font-size: 14.5px; }
header nav a:hover { color: #fff; text-decoration: underline; }
.alertbar { background: #B3261E; color: #fff; padding: 9px 22px;
  font-size: 14.5px; }
.alertbar.ok { background: #2E7D32; }
main { max-width: 1240px; margin: 0 auto; padding: 20px 22px 42px; }
h1 { font-size: 22px; margin: 4px 0 14px; }
h2 { font-size: 16.5px; margin: 26px 0 10px; color: #1F4E5F; }
.cards { display: flex; gap: 12px; flex-wrap: wrap; }
.card { background: #fff; border: 1px solid #dde3e6; border-radius: 10px;
  padding: 14px 18px; min-width: 138px; }
.card .num { font-size: 26px; font-weight: 700; color: #1F4E5F; }
.card .lbl { font-size: 12.5px; color: #5a6a70; margin-top: 2px; }
.card.red .num { color: #B3261E; }
table { width: 100%; border-collapse: collapse; background: #fff;
  border: 1px solid #dde3e6; border-radius: 10px; overflow: hidden; }
th { background: #EAF1F3; text-align: left; font-size: 12.5px;
  padding: 9px 10px; color: #33525d; }
td { padding: 8px 10px; font-size: 14px; border-top: 1px solid #edf1f2;
  vertical-align: top; }
tr:hover td { background: #f7fafb; }
a { color: #1668b4; }
.badge { display: inline-block; min-width: 22px; text-align: center;
  padding: 1px 6px; border-radius: 6px; font-weight: 700; font-size: 13px;
  color: #fff; }
.bA { background: #2E7D32; } .bB { background: #B26A00; }
.bC { background: #8a979c; }
.pill { display: inline-block; padding: 1px 9px; border-radius: 20px;
  font-size: 12.5px; white-space: nowrap; }
.st-new { background: #E3ECF7; color: #14487a; }
.st-mid { background: #EDEFF1; color: #444; }
.st-hot { background: #FFE9C7; color: #7a4a00; }
.st-good { background: #D8F0DA; color: #1d5c26; }
.st-bad { background: #F6DAD8; color: #8c1d18; }
.st-dead { background: #3d3d3d; color: #fff; }
.filters { background: #fff; border: 1px solid #dde3e6; border-radius: 10px;
  padding: 12px 14px; display: flex; gap: 10px; flex-wrap: wrap;
  align-items: end; margin-bottom: 14px; }
.filters label { font-size: 12px; color: #5a6a70; display: block;
  margin-bottom: 3px; }
input, select, textarea { font: inherit; font-size: 14px; padding: 7px 9px;
  border: 1px solid #bfccd1; border-radius: 7px; background: #fff; }
button, .btn { background: #1F4E5F; color: #fff; border: 0;
  padding: 8px 16px; border-radius: 7px; cursor: pointer; font-size: 14px;
  text-decoration: none; display: inline-block; }
button:hover, .btn:hover { background: #173d4a; }
button.ghost, .btn.ghost { background: #EAF1F3; color: #1F4E5F; }
.small { font-size: 12.5px; color: #5a6a70; }
.remark { background: #fff; border: 1px solid #dde3e6; border-radius: 10px;
  padding: 10px 14px; margin-bottom: 9px; }
.remark .when { font-size: 12px; color: #7b8a90; }
.tag { font-size: 11px; background: #EDEFF1; border-radius: 4px;
  padding: 0 5px; color: #555; }
.detail { background: #fff; border: 1px solid #dde3e6; border-radius: 10px;
  padding: 18px 20px; }
.detail .grid { display: grid; grid-template-columns: repeat(auto-fit,
  minmax(210px, 1fr)); gap: 10px 26px; margin-top: 8px; }
.detail .k { font-size: 12px; color: #5a6a70; }
.detail .v { font-size: 15px; margin-bottom: 8px; }
.overdue td { background: #fdf1f0; }
.empty { background: #fff; border: 1px dashed #b9c6cb; border-radius: 10px;
  padding: 26px; text-align: center; color: #5a6a70; }
footer { text-align: center; color: #8a979c; font-size: 12.5px;
  padding: 0 0 26px; }
"""


def page(title, body, alert=None):
    bar = ""
    if alert:
        cls, text = alert
        bar = f'<div class="alertbar {cls}">{text}</div>'
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} - Vaani CRM</title><style>{CSS}</style></head>
<body>
<header><span class="brand">Vaani CRM</span>
<nav><a href="/">Dashboard</a><a href="/leads">All leads</a>
<a href="/followups">Follow-ups</a><a href="/export.csv">Export CSV</a>
</nav></header>
{bar}
<main>{body}</main>
<footer>Vaani CRM v{VERSION} - runs only on this computer. Your data
never leaves this PC.</footer>
</body></html>"""


# ---------------------------------------------------------------- pages

def status_options(selected):
    out = []
    for s in STATUSES:
        sel = " selected" if s == selected else ""
        out.append(f'<option value="{esc(s)}"{sel}>{esc(s)}</option>')
    return "".join(out)


def alert_for(conn):
    overdue, due_today, _ = due_leads(conn)
    if overdue:
        return ("", f"&#9888; {len(overdue)} follow-up(s) OVERDUE and "
                    f"{len(due_today)} due today - open Follow-ups and "
                    "clear them first.")
    if due_today:
        return ("", f"{len(due_today)} follow-up(s) due today - open "
                    "Follow-ups to see who to call.")
    return None


def due_row(r, show_date=True):
    late = ""
    if r["followup"]:
        d = (date.today() - date.fromisoformat(r["followup"])).days
        if d > 0:
            late = f' <span class="pill st-bad">{d}d late</span>'
    datecell = fmt_date(r["followup"]) + late if show_date else ""
    return (
        f'<tr><td><a href="/lead?phone={esc(r["phone"])}">'
        f'<b>{esc(r["name"])}</b></a><div class="small">'
        f'{esc(r["angle"] or "")}</div></td>'
        f'<td>{fmt_phone(r["phone"])}</td>'
        f'<td>{datecell}</td>'
        f'<td>{tier_badge(r["tier"])} {esc(r["score"] or "")}</td>'
        f'<td>{status_pill(r["status"])}</td></tr>')


def lead_row(r, next_url):
    remark = esc((r["last_remark"] or "")[:64])
    if r["n_remarks"]:
        remark += f' <span class="small">({r["n_remarks"]})</span>'
    return (
        f'<tr><td><a href="/lead?phone={esc(r["phone"])}">'
        f'<b>{esc(r["name"])}</b></a><div class="small">'
        f'{esc(r["angle"] or "")}</div></td>'
        f'<td>{fmt_phone(r["phone"])}</td>'
        f'<td class="small">{esc((r["address"] or "")[:58])}</td>'
        f'<td>{tier_badge(r["tier"])} {esc(r["score"] or "")}</td>'
        f'<td><form method="post" action="/update" style="margin:0">'
        f'<input type="hidden" name="phone" value="{esc(r["phone"])}">'
        f'<input type="hidden" name="followup" '
        f'value="{esc(r["followup"])}">'
        f'<input type="hidden" name="next" value="{esc(next_url)}">'
        f'<select name="status" onchange="this.form.submit()">'
        f'{status_options(r["status"])}</select></form></td>'
        f'<td>{fmt_date(r["followup"])}</td>'
        f'<td class="small">{remark}</td>'
        f'<td class="small">{esc(r["times_seen"] or 1)}x</td></tr>')


def dashboard_html(conn):
    total = conn.execute("SELECT COUNT(*) c FROM leads").fetchone()["c"]
    if total == 0:
        return ("<h1>Dashboard</h1><div class='empty'>No leads yet.<br>"
                "Run Vaani once (double-click <b>run-vaani.bat</b>) and "
                "every restaurant it scrapes will appear here "
                "automatically.</div>")
    tier_a = conn.execute(
        "SELECT COUNT(*) c FROM leads WHERE tier = 'A'").fetchone()["c"]
    counts = {}
    for row in conn.execute(
            "SELECT COALESCE(s.status, 'Not called') st, COUNT(*) c"
            " FROM leads l LEFT JOIN crm_state s ON s.phone = l.phone"
            " GROUP BY st"):
        counts[row["st"]] = row["c"]
    demos = counts.get("Demo booked", 0) + counts.get("Demo done", 0)
    won = counts.get("Won", 0)
    overdue, due_today, _ = due_leads(conn)
    cards = (
        f'<div class="cards">'
        f'<div class="card"><div class="num">{total}</div>'
        f'<div class="lbl">Total leads</div></div>'
        f'<div class="card"><div class="num">{tier_a}</div>'
        f'<div class="lbl">A-tier (call today)</div></div>'
        f'<div class="card"><div class="num">{len(due_today)}</div>'
        f'<div class="lbl">Follow-ups due today</div></div>'
        f'<div class="card red"><div class="num">{len(overdue)}</div>'
        f'<div class="lbl">Overdue follow-ups</div></div>'
        f'<div class="card"><div class="num">{demos}</div>'
        f'<div class="lbl">Demos booked / done</div></div>'
        f'<div class="card"><div class="num">{won}</div>'
        f'<div class="lbl">Won</div></div></div>')
    body = "<h1>Dashboard</h1>" + cards
    if overdue:
        rows = "".join(due_row(r) for r in overdue)
        body += (f"<h2>Overdue follow-ups ({len(overdue)})</h2>"
                 f"<table><tr><th>Restaurant</th><th>Phone</th>"
                 f"<th>Follow-up</th><th>Tier / Score</th>"
                 f"<th>Status</th></tr>{rows}</table>")
    if due_today:
        rows = "".join(due_row(r) for r in due_today)
        body += (f"<h2>Due today ({len(due_today)})</h2>"
                 f"<table><tr><th>Restaurant</th><th>Phone</th>"
                 f"<th>Follow-up</th><th>Tier / Score</th>"
                 f"<th>Status</th></tr>{rows}</table>")
    if not overdue and not due_today:
        body += ('<div class="empty" style="margin-top:18px">No '
                 'follow-ups due. Set a follow-up date on any lead '
                 'and it will appear here on the day.</div>')
    funnel = "".join(
        f"<tr><td>{status_pill(s)}</td><td>{counts.get(s, 0)}</td></tr>"
        for s in STATUSES if counts.get(s, 0))
    body += (f"<h2>Your funnel</h2><table style='max-width:420px'>"
             f"<tr><th>Status</th><th>Leads</th></tr>{funnel}</table>")
    recents = conn.execute(
        "SELECT r.*, l.name FROM crm_remarks r JOIN leads l"
        " ON l.phone = r.phone ORDER BY r.id DESC LIMIT 8").fetchall()
    if recents:
        items = "".join(
            f'<div class="remark"><div class="when">'
            f'{esc(r["ts"][:10])} - <a href="/lead?phone='
            f'{esc(r["phone"])}">{esc(r["name"])}</a></div>'
            f'{esc(r["text"][:160])}</div>' for r in recents)
        body += f"<h2>Latest remarks</h2>{items}"
    return body


def leads_html(conn, params):
    q = params.get("q", "")
    tier = params.get("tier", "")
    status = params.get("status", "")
    sort = params.get("sort", "score")
    rows = fetch_leads(conn, q, tier, status, sort)
    query = urllib.parse.urlencode(
        {k: v for k, v in
         (("q", q), ("tier", tier), ("status", status), ("sort", sort))
         if v})
    next_url = "/leads" + (f"?{query}" if query else "")
    tier_opts = "".join(
        f'<option value="{t}"{" selected" if tier == t else ""}>'
        f'{t or "All tiers"}</option>' for t in ("", "A", "B", "C"))
    status_opts = '<option value="">All statuses</option>' + "".join(
        f'<option value="{esc(s)}"{" selected" if status == s else ""}>'
        f'{esc(s)}</option>' for s in STATUSES)
    sort_opts = "".join(
        f'<option value="{v}"{" selected" if sort == v else ""}>'
        f'{lbl}</option>'
        for v, lbl in (("score", "Best score first"),
                       ("followup", "Follow-up date"),
                       ("recent", "Recently seen"),
                       ("name", "Name")))
    body = f"""<h1>All leads ({len(rows)})</h1>
<form class="filters" method="get" action="/leads">
<div><label>Search name / phone / area / email</label>
<input name="q" value="{esc(q)}" size="30"></div>
<div><label>Tier</label><select name="tier">{tier_opts}</select></div>
<div><label>Status</label><select name="status">{status_opts}</select></div>
<div><label>Sort</label><select name="sort">{sort_opts}</select></div>
<div><button type="submit">Filter</button>
<a class="btn ghost" href="/leads">Clear</a></div></form>"""
    if not rows:
        body += ('<div class="empty">No leads match. Clear the filters, '
                 'or run Vaani to scrape more.</div>')
        return body
    trs = "".join(lead_row(r, next_url) for r in rows)
    body += (f"<table><tr><th>Restaurant</th><th>Phone</th>"
             f"<th>Location</th><th>Score</th><th>Status</th>"
             f"<th>Follow-up</th><th>Last remark</th><th>Seen</th></tr>"
             f"{trs}</table>")
    return body


def lead_html(conn, lead):
    phone = lead["phone"]
    remarks = fetch_remarks(conn, phone)
    wa = f"https://wa.me/91{phone}" if len(str(phone)) == 10 else "#"
    grid = (
        f'<div><div class="k">Phone</div><div class="v">'
        f'{fmt_phone(phone)} &nbsp;<a class="btn ghost" '
        f'href="tel:+91{esc(phone)}">Call</a> '
        f'<a class="btn ghost" href="{wa}" target="_blank">WhatsApp</a>'
        f'</div></div>'
        f'<div><div class="k">Email</div><div class="v">'
        f'{esc(lead["email"] or "-")}</div></div>'
        f'<div><div class="k">Location</div><div class="v">'
        f'{esc(lead["address"] or "-")}</div></div>'
        f'<div><div class="k">Website</div><div class="v">'
        + (f'<a href="{esc(lead["website"])}" target="_blank">'
           f'{esc(lead["website"])}</a>' if lead["website"] else "-")
        + f'</div></div>'
        f'<div><div class="k">Rating / Reviews</div><div class="v">'
        f'{esc(lead["rating"] or "-")} / {esc(lead["reviews"] or "-")}'
        f'</div></div>'
        f'<div><div class="k">Score / Tier</div><div class="v">'
        f'{tier_badge(lead["tier"])} {esc(lead["score"] or "-")}</div>'
        f'</div>'
        f'<div><div class="k">Pitch angle</div><div class="v">'
        f'{esc(lead["angle"] or "-")}</div></div>'
        f'<div><div class="k">First seen / Last seen / Times seen</div>'
        f'<div class="v">{fmt_date(lead["first_seen"])} / '
        f'{fmt_date(lead["last_seen"])} / {esc(lead["times_seen"] or 1)}'
        f'</div></div>')
    thread = ""
    for r in remarks:
        when = r["ts"]
        try:
            when = datetime.fromisoformat(when).strftime(
                "%d %b %Y, %I:%M %p")
        except Exception:
            pass
        tag = (' <span class="tag">from tracker</span>'
               if r["source"] == "tracker" else "")
        thread += (f'<div class="remark"><div class="when">{esc(when)}'
                   f'{tag}</div>{esc(r["text"]).replace(chr(10), "<br>")}'
                   f'</div>')
    if not thread:
        thread = '<div class="empty">No remarks yet.</div>'
    body = f"""<h1>{esc(lead['name'])}</h1>
<div class="detail"><div class="grid">{grid}</div>
<h2>Status &amp; follow-up</h2>
<form method="post" action="/update">
<input type="hidden" name="phone" value="{esc(phone)}">
<input type="hidden" name="next" value="/lead?phone={esc(phone)}">
<select name="status">{status_options(lead['status'])}</select>
<input type="date" id="fu" name="followup" value="{esc(lead['followup'])}">
<button type="button" class="ghost" onclick="setFU(1)">Tomorrow</button>
<button type="button" class="ghost" onclick="setFU(3)">+3 days</button>
<button type="button" class="ghost" onclick="setFU(7)">+1 week</button>
<button type="button" class="ghost" onclick="clearFU()">No date</button>
<button type="submit">Save</button></form>
<div class="small" style="margin-top:6px">Current status:
{status_pill(lead['status'])} &nbsp; Follow-up:
{fmt_date(lead['followup']) or 'not set'}</div></div>
<h2>Remarks ({len(remarks)})</h2>
<form method="post" action="/remark" style="margin-bottom:14px">
<input type="hidden" name="phone" value="{esc(phone)}">
<textarea name="text" rows="3" style="width:100%"
placeholder="What happened on the call? Who did you speak to? What did they say?"></textarea>
<br><button type="submit">Add remark</button></form>
{thread}
<script>
function setFU(days){{var d=new Date(); d.setDate(d.getDate()+days);
 var m=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
 document.getElementById('fu').value=d.getFullYear()+'-'+m+'-'+dd;}}
function clearFU(){{document.getElementById('fu').value='';}}
</script>"""
    return body


def followups_html(conn):
    overdue, due_today, upcoming = due_leads(conn)
    week = [r for r in upcoming
            if (date.fromisoformat(r["followup"])
                - date.today()).days <= 7]
    later = [r for r in upcoming
             if (date.fromisoformat(r["followup"])
                 - date.today()).days > 7]

    def section(title, rows, cls=""):
        if not rows:
            return ""
        trs = "".join(due_row(r) for r in rows)
        return (f"<h2>{title} ({len(rows)})</h2><table class='{cls}'>"
                f"<tr><th>Restaurant</th><th>Phone</th><th>Follow-up</th>"
                f"<th>Tier / Score</th><th>Status</th></tr>{trs}</table>")

    body = "<h1>Follow-ups</h1>"
    body += section("Overdue - call these first", overdue, "overdue")
    body += section("Due today", due_today)
    body += section("Coming this week", week)
    body += section("Later", later)
    if not (overdue or due_today or upcoming):
        body += ('<div class="empty">No follow-ups set. Open any lead '
                 'and set a follow-up date - it will show up here on '
                 'the day.</div>')
    return body


def export_csv_text(conn):
    import csv
    import io
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Restaurant Name", "Phone", "Location", "Email",
                "Website", "Rating", "Reviews", "Score", "Tier",
                "Pitch Angle", "Status", "Follow-up Date", "Remarks",
                "First Seen", "Last Seen", "Times Seen"])
    for r in fetch_leads(conn):
        w.writerow([r["name"], r["phone"], r["address"], r["email"],
                    r["website"], r["rating"], r["reviews"], r["score"],
                    r["tier"], r["angle"], r["status"], r["followup"],
                    r["n_remarks"], r["first_seen"], r["last_seen"],
                    r["times_seen"]])
    return buf.getvalue()


# ---------------------------------------------------------------- server

class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def _send(self, code, body, ctype="text/html; charset=utf-8",
              extra=None):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _redirect(self, to):
        if not to or not to.startswith("/") or to.startswith("//"):
            to = "/"
        self.send_response(303)
        self.send_header("Location", to)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = {k: v[0] for k, v in
                  urllib.parse.parse_qs(parsed.query).items()}
        conn = get_conn()
        try:
            if parsed.path == "/":
                self._send(200, page("Dashboard", dashboard_html(conn),
                                     alert=alert_for(conn)))
            elif parsed.path == "/leads":
                self._send(200, page("All leads",
                                     leads_html(conn, params)))
            elif parsed.path == "/lead":
                lead = fetch_lead(conn, params.get("phone", ""))
                if not lead:
                    self._send(404, page(
                        "Not found",
                        "<h1>Lead not found</h1><div class='empty'>"
                        "That lead is not in the database. "
                        "<a href='/leads'>Back to all leads</a></div>"))
                else:
                    self._send(200, page(lead["name"] or "Lead",
                                         lead_html(conn, lead)))
            elif parsed.path == "/followups":
                self._send(200, page("Follow-ups",
                                     followups_html(conn),
                                     alert=alert_for(conn)))
            elif parsed.path == "/export.csv":
                self._send(200, export_csv_text(conn),
                           ctype="text/csv; charset=utf-8",
                           extra={"Content-Disposition":
                                  'attachment; filename='
                                  '"vaani-crm-export.csv"'})
            else:
                self._send(404, page("Not found",
                                     "<h1>Page not found</h1>"))
        finally:
            conn.close()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        try:
            length = int(self.headers.get("Content-Length", 0) or 0)
        except Exception:
            length = 0
        if length > 65536:
            self._send(413, "Too large")
            return
        raw = self.rfile.read(length).decode("utf-8", "replace")
        form = {k: v[0] for k, v in
                urllib.parse.parse_qs(raw).items()}
        conn = get_conn()
        try:
            if parsed.path == "/update":
                phone = form.get("phone", "")
                if fetch_lead(conn, phone):
                    set_status(conn, phone, form.get("status", ""),
                               form.get("followup", ""))
                self._redirect(form.get("next") or
                               f"/lead?phone={phone}")
            elif parsed.path == "/remark":
                phone = form.get("phone", "")
                if fetch_lead(conn, phone):
                    add_remark(conn, phone, form.get("text", ""))
                self._redirect(f"/lead?phone={phone}")
            else:
                self._send(404, page("Not found",
                                     "<h1>Page not found</h1>"))
        finally:
            conn.close()


def main():
    self_update()
    conn = get_conn()
    backup_db()
    imported = import_tracker(conn)
    conn.close()
    server = None
    for port in PORTS:
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if server is None:
        print("  Could not start Vaani CRM - ports 8765-8769 are busy.")
        print("  Close the other copy of Vaani CRM and try again.")
        return
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print()
    print(f"  VAANI CRM v{VERSION}")
    print("  ----------------------")
    if imported:
        print(f"  Imported {imported} lead(s) from your Excel tracker"
              " (statuses, notes and callback dates).")
    print(f"  Open in your browser: {url}")
    print("  Keep this window open while you work in the CRM.")
    print("  To stop: close this window or press Ctrl+C.")
    print()
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
