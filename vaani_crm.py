#!/usr/bin/env python3
"""
Vaani CRM v1.4
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

VERSION = "1.4.1"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/vaani_crm.py")
PORTS = (8765, 8766, 8767, 8768, 8769)
USER_NAME = "Yash"

# WhatsApp kit (v1.3): the report card generator and brochure live in
# this same folder; BMF_WEBSITE is added to kit messages when set.
REPORT_CARD_SCRIPT = os.path.join(HERE, "report_card.py")
REPORT_CARDS_DIR = os.path.join(HERE, "report-cards")
BROCHURE_PATH = os.path.join(HERE, "book-my-food-brochure.pdf")
BMF_WEBSITE = "https://bookmyfood.si/"
KIT_SEND_SCRIPT = os.path.join(HERE, "kit_send.py")
KIT_SEND_URL = ("https://raw.githubusercontent.com/yash6969-io/"
                "vaani-lead-engine/main/kit_send.py")

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


# ---------------------------------------------------------------- browser launch

def _chrome_paths():
    cands = []
    if os.name == "nt":
        for env in ("ProgramFiles", "ProgramFiles(x86)",
                    "LocalAppData"):
            base = os.environ.get(env)
            if base:
                cands.append(os.path.join(
                    base, "Google", "Chrome", "Application",
                    "chrome.exe"))
    else:
        for name in ("google-chrome", "google-chrome-stable",
                     "chromium", "chromium-browser"):
            p = shutil.which(name)
            if p:
                cands.append(p)
    return [c for c in cands if os.path.exists(c)]


def open_browser(url):
    """Open the CRM in Google Chrome as its own app window when
    Chrome is installed - not in whatever the Windows default
    browser happens to be. Falls back to the default browser.
    Force the default with --browser=default (or the environment
    variable VAANI_CRM_BROWSER=default)."""
    pref = os.environ.get("VAANI_CRM_BROWSER", "chrome").lower()
    for a in sys.argv[1:]:
        if a.startswith("--browser="):
            pref = a.split("=", 1)[1].lower()
    if pref != "default":
        for p in _chrome_paths():
            try:
                subprocess.Popen(
                    [p, f"--app={url}"], stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL)
                return
            except Exception:
                continue
        try:
            if webbrowser.get("chrome").open(url):
                return
        except Exception:
            pass
    try:
        webbrowser.open(url)
    except Exception:
        pass


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
    conn.execute("""CREATE TABLE IF NOT EXISTS kit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, phone TEXT, ts TEXT,
        channel TEXT DEFAULT 'whatsapp')""")
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
             WHERE r.phone = l.phone AND r.source != 'status')
             AS n_remarks,
           (SELECT r.text FROM crm_remarks r WHERE r.phone = l.phone
             AND r.source != 'status'
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


def call_queue(conn):
    """Today's calling queue, in the order a rep should work it:
    overdue follow-ups first, then today's follow-ups, then fresh
    A-tier and B-tier leads that have never been called."""
    overdue, due_today, _ = due_leads(conn)
    seen = {r["phone"] for r in overdue + due_today}
    fresh = [r for r in fetch_leads(conn, status="Not called")
             if r["phone"] not in seen and r["tier"] in ("A", "B")]
    fresh.sort(key=lambda r: (-(r["score"] or 0)))
    a_first = [r for r in fresh if r["tier"] == "A"] + \
        [r for r in fresh if r["tier"] == "B"]
    return overdue + due_today + a_first


def set_status(conn, phone, status, followup=""):
    """followup=None keeps the existing follow-up date untouched."""
    if status not in STATUSES:
        return False
    cur = conn.execute(
        "SELECT status, followup FROM crm_state WHERE phone = ?",
        (phone,)).fetchone()
    old_status = cur["status"] if cur else "Not called"
    old_followup = cur["followup"] if cur else ""
    if followup is None:
        followup = old_followup or ""
    else:
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
    events = []
    if status != old_status:
        events.append(f"Status: {old_status} \u2192 {status}")
    if followup != (old_followup or ""):
        events.append("Follow-up: " + (fmt_date(followup)
                                       if followup else "cleared"))
    if events:
        conn.execute(
            "INSERT INTO crm_remarks (phone, ts, text, source)"
            " VALUES (?, ?, ?, 'status')",
            (phone, now, " \u00b7 ".join(events)))
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


# ---------------------------------------------------------------- whatsapp kit

KIT_CHAINS = ("mcdonald", "domino", "pizza hut", "kfc", "burger king",
              "subway", "starbucks", "haldiram", "barbeque nation",
              "wow momo", "wow! momo", "la pino", "baskin",
              "cafe coffee", "jumbo king", "behrouz", "faasos",
              "freshmenu", "oven story", "chaayos", "third wave",
              "burger singh")


def fmt_ts(ts):
    try:
        return datetime.fromisoformat(ts).strftime("%d %b %Y, %I:%M %p")
    except Exception:
        return ts or ""


def kit_area(lead):
    segs = [s.strip() for s in str(lead["address"] or "").split(",")
            if s.strip()]
    segs = [s for s in segs if not re.fullmatch(r"[\d\s\-]+", s)]
    if len(segs) > 1:
        segs = segs[:-1]
    return segs[-1] if segs else ""


def kit_skip_reason(lead):
    """The Report Card generator's don't-pitch screen, plus CRM
    terminal statuses. '' means: pitch this lead."""
    name_low = (lead["name"] or "").lower()
    if any(c in name_low for c in KIT_CHAINS):
        return "National chain - decisions are centralised"
    if lead["suppressed"]:
        return "Marked " + (lead["suppress_reason"] or "suppressed")
    if (lead["status"] or "") in TERMINAL_STATUSES:
        return "Already closed: " + lead["status"]
    if ((lead["angle"] or "").strip() == "Operations upgrade"
            and (lead["rating"] or 0) >= 4.2
            and (lead["reviews"] or 0) >= 200):
        return "Already has direct ordering and a solid setup"
    return ""


def kit_card_file(phone):
    if not phone or len(str(phone)) < 10:
        return None
    try:
        for f in os.listdir(REPORT_CARDS_DIR):
            if f.startswith(str(phone) + "-") and f.endswith(".png"):
                return os.path.join(REPORT_CARDS_DIR, f)
    except Exception:
        pass
    return None


def generate_card(phone):
    """Run the Report Card generator for one lead; True when a card
    PNG exists afterwards."""
    if not os.path.exists(REPORT_CARD_SCRIPT):
        return False
    try:
        subprocess.run([sys.executable, REPORT_CARD_SCRIPT,
                        str(phone), "--no-update"], cwd=HERE,
                       timeout=120, capture_output=True)
    except Exception:
        pass
    return kit_card_file(phone) is not None


def kit_sent_map(conn):
    rows = conn.execute(
        "SELECT phone, MAX(ts) ts FROM kit_log GROUP BY phone"
    ).fetchall()
    return {r["phone"]: r["ts"] for r in rows if r["ts"]}


def kit_sent_at(conn, phone):
    row = conn.execute(
        "SELECT MAX(ts) ts FROM kit_log WHERE phone = ?",
        (phone,)).fetchone()
    return row["ts"] if row and row["ts"] else None


def kit_pending_count(conn):
    sent = kit_sent_map(conn)
    rows = conn.execute(
        LEAD_SELECT + " WHERE l.suppressed = 0").fetchall()
    n = 0
    for r in rows:
        if r["phone"] in sent or kit_skip_reason(r):
            continue
        n += 1
    return n


def kit_observation(lead):
    angle = (lead["angle"] or "").strip()
    reviews = lead["reviews"] or 0
    if angle.startswith("Aggregator-dependent") or \
            angle.startswith("No direct ordering"):
        return ("One thing I noticed: customers who find you still "
                "have to order through apps that keep 25-35% of "
                "every order.")
    if not (lead["website"] or "").strip():
        return ("One thing I noticed: your Google listing has no "
                "website, so a customer who wants to reach you "
                "directly has nowhere to land.")
    if reviews >= 300:
        return (f"One thing I noticed: {reviews:,} people have "
                "reviewed you, but their names and numbers sit with "
                "the apps - you cannot reach a single regular "
                "directly.")
    if lead["rating"]:
        return (f"One thing I noticed: at {lead['rating']} stars, "
                "your Google presence is doing more selling than you "
                "get credit for - the report shows where it leaks.")
    return ("The report shows exactly where your Google presence "
            "is quietly costing you orders.")


def kit_message(lead):
    name = lead["name"] or "your restaurant"
    lines = [
        f"Namaste {name} team, this is {USER_NAME} Tanna from "
        "Book My Food.",
        "",
        "A little about us: Book My Food is an operations "
        "platform for restaurants - your own online ordering "
        "with no aggregator commission on your regulars, plus "
        "billing and payments, kitchen order flow and your own "
        "customer list, all in one place.",
        "",
        kit_observation(lead),
        "",
        f"I am sharing two things right after this message - a "
        f"free one-page Report Card on how {name} looks on "
        "Google today, and our brochure with the platform "
        "details. Please do check them out - no charge, no "
        "obligation.",
        "",
        f"If they are useful, I would love 10 minutes to show "
        f"you how this would work for {name}.",
    ]
    if BMF_WEBSITE:
        lines += ["", f"More about us: {BMF_WEBSITE}"]
    lines += ["", f"- {USER_NAME} Tanna, Book My Food"]
    return "\n".join(lines)


def kit_brochure_message(lead):
    lines = [
        "And this is our short brochure - what Book My Food does, "
        "how your own ordering page works, and the 7-day trial we "
        "start with.",
        "",
        "Happy to walk you through it on a short call whenever "
        "suits you.",
    ]
    if BMF_WEBSITE:
        lines += ["", f"Website: {BMF_WEBSITE}"]
    return "\n".join(lines)


def wa_link(phone, text=None):
    base = f"https://wa.me/91{phone}"
    if text:
        base += "?text=" + urllib.parse.quote(text)
    return base


def mark_kit_sent(conn, phone):
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        "INSERT INTO kit_log (phone, ts, channel) VALUES"
        " (?, ?, 'whatsapp')", (phone, now))
    bits = "intro message, report card"
    if os.path.exists(BROCHURE_PATH):
        bits += ", brochure"
    conn.execute(
        "INSERT INTO crm_remarks (phone, ts, text, source) VALUES"
        " (?, ?, ?, 'kit')",
        (phone, now, f"WhatsApp kit sent - in order: {bits}"))
    conn.commit()
    lead = fetch_lead(conn, phone)
    if lead and not lead["followup"]:
        tomorrow = (date.today() + timedelta(days=1)).isoformat()
        set_status(conn, phone, lead["status"] or "Not called",
                   tomorrow)
    return now


def ensure_kit_sender():
    """The Auto-send helper (kit_send.py) lives next to the CRM. If
    it is not there yet, fetch it from the same GitHub repo the CRM
    updates itself from - same sanity checks as the updater."""
    if os.path.exists(KIT_SEND_SCRIPT):
        return True
    try:
        req = urllib.request.Request(
            KIT_SEND_URL, headers={"User-Agent": "VaaniCRM"})
        with urllib.request.urlopen(req, timeout=15) as r:
            content = r.read().decode("utf-8", "replace")
    except Exception:
        return False
    if "Vaani Kit Sender" not in content or len(content) < 5000:
        return False
    try:
        compile(content, "kit_send.py", "exec")
        with open(KIT_SEND_SCRIPT, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        return False
    return True


def autosend_kit(conn, phone):
    """Send the whole kit - intro + report card + brochure, together -
    through WhatsApp Web via kit_send.py. Returns (ok, message).
    Both files or nothing: if either is missing and cannot be made,
    nothing is sent."""
    lead = fetch_lead(conn, phone)
    if not lead:
        return False, "Lead not found."
    reason = kit_skip_reason(lead)
    if reason:
        return False, "On the don't-pitch list: " + reason
    if not ensure_kit_sender():
        return False, ("The sender tool (kit_send.py) is not in "
                       "this folder and could not be downloaded - "
                       "check the internet connection.")
    card = kit_card_file(phone)
    if not card:
        generate_card(phone)
        card = kit_card_file(phone)
    if not card:
        return False, ("The report card could not be generated for "
                       "this lead, so nothing was sent.")
    if not os.path.exists(BROCHURE_PATH):
        return False, ("book-my-food-brochure.pdf is not in this "
                       "folder - the kit is only sent complete, so "
                       "nothing was sent.")
    caption_path = os.path.join(HERE, "kit_caption.txt")
    try:
        with open(caption_path, "w", encoding="utf-8") as f:
            f.write(kit_message(lead))
    except Exception:
        return False, "Could not prepare the intro message."
    try:
        proc = subprocess.run(
            [sys.executable, KIT_SEND_SCRIPT, str(phone), card,
             BROCHURE_PATH, caption_path],
            cwd=HERE, capture_output=True, text=True, timeout=420)
        out = (proc.stdout or "") + (proc.stderr or "")
    except subprocess.TimeoutExpired:
        return False, ("Sending took too long and was stopped. If "
                       "the WhatsApp window is still open, check "
                       "whether the kit went through before trying "
                       "again.")
    except Exception as e:
        return False, f"The sender could not start: {e}"
    lines = [ln.strip() for ln in out.splitlines() if ln.strip()]
    if lines and lines[-1] == "SENT":
        return True, "Sent"
    for ln in reversed(lines):
        if ln.startswith("ERROR:"):
            return False, ln[len("ERROR:"):].strip()
    return False, ("The sender finished without confirming the "
                   "send. Check the WhatsApp window - if the kit "
                   "is in the chat, use Mark sent to log it.")


AUTOSEND_JS = """<script>
function autosend(btn, phone) {
  var old = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Sending - watch the WhatsApp window...';
  fetch('/api/kit-autosend', {method: 'POST',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body: 'phone=' + encodeURIComponent(phone)})
  .then(r => r.json()).then(j => {
    if (j.ok) {
      btn.textContent = 'Kit sent';
      setTimeout(function() { location.reload(); }, 900);
    } else {
      btn.disabled = false;
      btn.textContent = old;
      alert(j.error || 'Auto-send failed.');
    }
  }).catch(function() {
    btn.disabled = false;
    btn.textContent = old;
    alert('Auto-send could not reach the CRM.');
  });
}
</script>"""


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


MONO_COLORS = [("#E0F2F1", "#0F766E"), ("#E0E7FF", "#3730A3"),
               ("#FEF3C7", "#92400E"), ("#FCE7F3", "#9D174D"),
               ("#E0F2FE", "#075985"), ("#ECFCCB", "#3F6212"),
               ("#F3E8FF", "#6B21A8"), ("#FFE4E6", "#9F1239")]


def monogram(name, phone):
    words = [w for w in re.split(r"[^A-Za-z]+", str(name or "?"))
             if w]
    initials = "".join(w[0].upper() for w in words[:2]) or "?"
    idx = sum(ord(c) for c in str(phone)) % len(MONO_COLORS)
    bg, fg = MONO_COLORS[idx]
    return (f'<span class="mono" style="background:{bg};color:{fg}">'
            f'{esc(initials)}</span>')


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


def followup_chip(lead):
    fu = lead["followup"]
    if not fu or lead["status"] in TERMINAL_STATUSES:
        return ""
    today = date.today().isoformat()
    if fu < today:
        d = (date.today() - date.fromisoformat(fu)).days
        return (f'<span class="chip chip-red">&#9200; {d}d overdue</span>')
    if fu == today:
        return '<span class="chip chip-amber">&#9200; Due today</span>'
    return (f'<span class="chip chip-gray">&#9200; '
            f'{fmt_date(fu)}</span>')


# ---------------------------------------------------------------- design system

ICONS = {
    "grid": '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
    "phone": '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.13.96.36 1.9.7 2.8a2 2 0 0 1-.45 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.45c.9.34 1.84.57 2.8.7A2 2 0 0 1 22 16.9z"/>',
    "board": '<rect x="3" y="3" width="5.5" height="18" rx="1.5"/><rect x="9.75" y="3" width="5.5" height="12" rx="1.5"/><rect x="16.5" y="3" width="5.5" height="8" rx="1.5"/>',
    "users": '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "download": '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/>',
    "search": '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
    "flame": '<path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"/>',
    "check": '<path d="M20 6 9 17l-5-5"/>',
    "calendar": '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
    "send": '<path d="m22 2-7 20-4-9-9-4Z"/><path d="M22 2 11 13"/>',
}


def icon(name, size=18):
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" '
            f'fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round">'
            f'{ICONS.get(name, "")}</svg>')


CSS = """
:root { --bg:#F4F6F8; --panel:#FFFFFF; --ink:#101828; --muted:#667085;
  --faint:#98A2B3; --line:#E4E7EC; --accent:#0F766E; --accent2:#115E59;
  --sidebar:#0B1D26; --red:#D92D20; --amber:#B54708; --radius:14px;
  --shadow:0 1px 2px rgba(16,24,40,.06),0 1px 3px rgba(16,24,40,.05); }
* { box-sizing:border-box; }
html { scrollbar-color:#C4CED4 transparent; }
body { margin:0; font-family:-apple-system,'Segoe UI',Inter,Roboto,
  'Helvetica Neue',Arial,sans-serif; background:var(--bg);
  color:var(--ink); font-size:14.5px; line-height:1.45; }
::selection { background:#99F6E4; }
a { color:#1668B4; text-decoration:none; }
a:hover { text-decoration:underline; }
.sidebar { position:fixed; top:0; left:0; bottom:0; width:234px;
  background:linear-gradient(180deg,#0B1D26,#0E2836); color:#D7E3E8;
  display:flex; flex-direction:column; z-index:30; }
.brand { display:flex; align-items:center; gap:11px;
  padding:20px 18px 16px; }
.logo { width:38px; height:38px; border-radius:11px; flex:none;
  background:linear-gradient(135deg,#14B8A6,#0F766E); color:#fff;
  display:flex; align-items:center; justify-content:center;
  font-size:20px; font-weight:800; box-shadow:0 4px 14px rgba(20,184,166,.35); }
.brand .t1 { font-size:16.5px; font-weight:700; color:#fff;
  letter-spacing:.2px; }
.brand .t2 { font-size:11.5px; color:#8FA6B0; }
.sidebar nav { padding:6px 12px; flex:1; }
.sidebar nav a { display:flex; align-items:center; gap:11px;
  padding:9.5px 12px; margin-bottom:3px; border-radius:9px;
  color:#B9CBD3; font-size:14.5px; position:relative; }
.sidebar nav a:hover { background:rgba(255,255,255,.06); color:#fff;
  text-decoration:none; }
.sidebar nav a.active { background:rgba(20,184,166,.17); color:#fff; }
.sidebar nav a.active::before { content:''; position:absolute; left:-12px;
  top:8px; bottom:8px; width:3.5px; border-radius:3px;
  background:#2DD4BF; }
.navbadge { margin-left:auto; background:#E5484D; color:#fff;
  font-size:11.5px; font-weight:700; border-radius:20px;
  padding:0.5px 8px; }
.side-foot { padding:14px 18px; border-top:1px solid rgba(255,255,255,.08);
  font-size:12px; color:#8FA6B0; }
.side-foot b { color:#D7E3E8; }
.shell { margin-left:234px; min-height:100vh; display:flex;
  flex-direction:column; }
.topbar { position:sticky; top:0; z-index:20; height:62px;
  background:rgba(255,255,255,.88); backdrop-filter:blur(8px);
  border-bottom:1px solid var(--line); display:flex; align-items:center;
  gap:16px; padding:0 26px; }
.topbar .ptitle { font-size:16px; font-weight:650; }
.searchbox { margin-left:auto; position:relative; }
.searchbox svg { position:absolute; left:10px; top:50%;
  transform:translateY(-50%); color:var(--faint); }
.searchbox input { width:330px; padding:8.5px 12px 8.5px 34px;
  border:1px solid var(--line); border-radius:10px; font:inherit;
  font-size:14px; background:#F9FAFB; }
.searchbox input:focus { outline:2px solid #99F6E4; background:#fff;
  border-color:#5EEAD4; }
.searchbox .kbd { position:absolute; right:9px; top:50%;
  transform:translateY(-50%); font-size:11px; color:var(--faint);
  border:1px solid var(--line); border-radius:5px; padding:0 5px;
  background:#fff; }
main { padding:24px 26px 40px; max-width:1360px; width:100%;
  margin:0 auto; flex:1; }
footer { text-align:center; color:var(--faint); font-size:12.5px;
  padding:0 0 24px; }
h1 { font-size:23px; margin:2px 0 4px; letter-spacing:-.01em; }
h2 { font-size:16px; margin:28px 0 12px; letter-spacing:-.005em; }
.sub { color:var(--muted); font-size:14px; }
.btn { display:inline-flex; align-items:center; gap:7px;
  background:var(--accent); color:#fff; border:0; padding:8.5px 16px;
  border-radius:9px; cursor:pointer; font:inherit; font-size:14px;
  font-weight:600; box-shadow:var(--shadow); }
.btn:hover { background:var(--accent2); text-decoration:none; }
.btn.ghost { background:#fff; color:var(--ink);
  border:1px solid var(--line); box-shadow:none; }
.btn.ghost:hover { background:#F9FAFB; }
.btn.big { padding:12px 22px; font-size:15.5px; border-radius:11px; }
.btn.red { background:var(--red); }
button.btn { font-family:inherit; }
input, select, textarea { font:inherit; font-size:14.5px;
  padding:8px 10px; border:1px solid #CFD6DD; border-radius:9px;
  background:#fff; color:var(--ink); }
input:focus, select:focus, textarea:focus { outline:2px solid #99F6E4;
  border-color:#5EEAD4; }
.cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr));
  gap:13px; margin-top:16px; }
.kpi { background:var(--panel); border:1px solid var(--line);
  border-radius:var(--radius); padding:16px 18px; box-shadow:var(--shadow);
  display:flex; gap:13px; align-items:flex-start; }
.kpi .ic { width:40px; height:40px; border-radius:11px; flex:none;
  display:flex; align-items:center; justify-content:center; }
.kpi .num { font-size:26px; font-weight:750; line-height:1.05; }
.kpi .lbl { font-size:12.5px; color:var(--muted); margin-top:2px; }
.kpi.red .num { color:var(--red); }
.panel { background:var(--panel); border:1px solid var(--line);
  border-radius:var(--radius); box-shadow:var(--shadow);
  padding:18px 20px; }
table { width:100%; border-collapse:collapse; background:var(--panel);
  border:1px solid var(--line); border-radius:var(--radius);
  overflow:hidden; box-shadow:var(--shadow); }
th { text-align:left; font-size:11px; text-transform:uppercase;
  letter-spacing:.07em; color:var(--faint); padding:11px 12px;
  background:#FBFCFD; border-bottom:1px solid var(--line);
  font-weight:650; }
td { padding:10.5px 12px; border-bottom:1px solid #EEF1F4;
  vertical-align:middle; }
tr:last-child td { border-bottom:0; }
tbody tr:hover td, table tr:hover td { background:#F8FAFC; }
.name-cell { display:flex; align-items:center; gap:11px; }
.name-cell .nm { font-weight:620; color:var(--ink); }
.name-cell .an { font-size:12px; color:var(--muted); }
.mono { width:36px; height:36px; border-radius:50%; flex:none;
  display:inline-flex; align-items:center; justify-content:center;
  font-size:13px; font-weight:700; }
.badge { display:inline-block; min-width:23px; text-align:center;
  padding:1.5px 7px; border-radius:20px; font-weight:700;
  font-size:12.5px; }
.bA { background:#DCFCE7; color:#166534; }
.bB { background:#FEF3C7; color:#92400E; }
.bC { background:#F1F5F9; color:#64748B; }
.pill { display:inline-block; padding:2px 10px; border-radius:20px;
  font-size:12.5px; font-weight:560; white-space:nowrap; }
.st-new { background:#E0E9F7; color:#14487A; }
.st-mid { background:#EEF1F4; color:#475467; }
.st-hot { background:#FFEED3; color:#8A4B00; }
.st-good { background:#D9F2E0; color:#166534; }
.st-bad { background:#FBE3E1; color:#8C1D18; }
.st-dead { background:#3A3A3A; color:#fff; }
.chip { display:inline-block; padding:1.5px 8px; border-radius:20px;
  font-size:12px; font-weight:600; white-space:nowrap; }
.chip-red { background:#FDECEA; color:#B42318; }
.chip-amber { background:#FEF3C7; color:#92400E; }
.chip-gray { background:#EEF1F4; color:#475467; }
.filters { background:var(--panel); border:1px solid var(--line);
  border-radius:var(--radius); box-shadow:var(--shadow);
  padding:13px 15px; display:flex; gap:11px; flex-wrap:wrap;
  align-items:end; margin:14px 0; }
.filters label { font-size:11.5px; color:var(--muted); display:block;
  margin-bottom:4px; font-weight:600; text-transform:uppercase;
  letter-spacing:.05em; }
.funnel-row { display:flex; align-items:center; gap:12px;
  padding:5.5px 0; }
.funnel-row .fl { width:178px; flex:none; }
.funnel-row .ft { flex:1; background:#EEF1F4; border-radius:8px;
  height:15px; overflow:hidden; }
.funnel-row .fb { height:100%; border-radius:8px;
  background:linear-gradient(90deg,#14B8A6,#0F766E); min-width:3px; }
.funnel-row .fb.good { background:linear-gradient(90deg,#4ADE80,#16A34A); }
.funnel-row .fb.bad { background:#D0D5DD; }
.funnel-row .fc { width:38px; text-align:right; font-weight:650; }
.taskcard { display:flex; align-items:center; gap:13px;
  background:var(--panel); border:1px solid var(--line);
  border-left:4px solid var(--red); border-radius:12px;
  padding:11px 15px; margin-bottom:9px; box-shadow:var(--shadow); }
.taskcard.amber { border-left-color:#F79009; }
.taskcard .grow { flex:1; min-width:0; }
.empty { background:var(--panel); border:1.5px dashed #CBD5DC;
  border-radius:var(--radius); padding:30px; text-align:center;
  color:var(--muted); }
.remark { background:var(--panel); border:1px solid var(--line);
  border-radius:12px; padding:12px 15px; margin-bottom:9px;
  box-shadow:var(--shadow); }
.remark .when { font-size:12px; color:var(--faint); margin-bottom:3px; }
.event { text-align:center; color:var(--muted); font-size:12.5px;
  padding:5px 0 9px; }
.tag { font-size:11px; background:#EEF1F4; border-radius:5px;
  padding:0.5px 6px; color:#475467; font-weight:600; }
.small { font-size:12.5px; color:var(--muted); }
.board { display:flex; gap:12px; overflow-x:auto; padding:4px 2px 16px;
  align-items:flex-start; }
.col { width:252px; min-width:252px; background:#EBEEF1;
  border-radius:var(--radius); padding:10px; }
.col.drop { outline:2.5px dashed var(--accent); outline-offset:-2px; }
.col h3 { margin:2px 4px 10px; font-size:13px; display:flex;
  align-items:center; gap:8px; color:#344054; font-weight:650; }
.col .cnt { background:#fff; border-radius:12px; padding:0.5px 8px;
  font-size:12px; color:var(--muted); }
.dcard { background:#fff; border-radius:11px; padding:11px 12px;
  margin-bottom:9px; box-shadow:var(--shadow); cursor:grab;
  border:1px solid #E7EBEE; }
.dcard:active { cursor:grabbing; }
.dcard .nm { font-weight:630; display:flex; gap:8px;
  align-items:center; }
.dcard .meta { display:flex; gap:6px; flex-wrap:wrap;
  margin-top:8px; align-items:center; }
.progress { height:7px; background:#E4E7EC; border-radius:6px;
  overflow:hidden; }
.progress div { height:100%; background:linear-gradient(90deg,#14B8A6,#0F766E); }
.callwrap { max-width:660px; margin:0 auto; }
.callcard { background:var(--panel); border:1px solid var(--line);
  border-radius:18px; box-shadow:0 8px 28px rgba(16,24,40,.09);
  padding:26px 28px; }
.outcomes { display:grid; grid-template-columns:1fr 1fr; gap:9px;
  margin-top:14px; }
.outcomes form { margin:0; }
.outcomes button { width:100%; }
.detail-grid { display:grid; grid-template-columns:305px 1fr 300px;
  gap:16px; align-items:start; }
.detail-grid .k { font-size:11.5px; color:var(--muted);
  text-transform:uppercase; letter-spacing:.05em; font-weight:600; }
.detail-grid .v { font-size:15px; margin:2px 0 13px; }
.crumbs { font-size:13px; color:var(--muted); margin-bottom:8px; }
.hero { position:relative; overflow:hidden; border-radius:18px;
  padding:27px 30px; color:#EAF6F3; margin-bottom:18px;
  background:linear-gradient(118deg,#0A1B24 0%,#0D3040 52%,#0F766E 135%);
  box-shadow:0 10px 30px rgba(11,29,38,.25); }
.hero::before { content:''; position:absolute; width:380px;
  height:380px; border-radius:50%; top:-190px; right:-70px;
  background:radial-gradient(circle,rgba(45,212,191,.38),
  rgba(45,212,191,0) 65%); }
.hero::after { content:''; position:absolute; width:260px;
  height:260px; border-radius:50%; bottom:-170px; left:34%;
  background:radial-gradient(circle,rgba(20,184,166,.22),
  rgba(20,184,166,0) 65%); }
.hero h1 { color:#fff; font-size:27px; margin:0 0 5px;
  letter-spacing:-.02em; position:relative; z-index:1; }
.hero .hsub { color:#A9C6BE; font-size:14.5px; position:relative;
  z-index:1; }
.hero .hstats { display:flex; gap:26px; margin-top:17px;
  position:relative; z-index:1; flex-wrap:wrap; }
.hero .hstat .n { font-size:21px; font-weight:780; color:#fff;
  font-variant-numeric:tabular-nums; }
.hero .hstat .l { font-size:12px; color:#A9C6BE; }
.hero .hcta { margin-top:19px; display:flex; gap:10px;
  position:relative; z-index:1; }
.hero .btn { background:#2DD4BF; color:#07332C;
  box-shadow:0 4px 16px rgba(45,212,191,.4); }
.hero .btn:hover { background:#5EEAD4; }
.hero .btn.ghost { background:rgba(255,255,255,.09); color:#EAF6F3;
  border:1px solid rgba(255,255,255,.25); box-shadow:none; }
.hero .btn.ghost:hover { background:rgba(255,255,255,.17); }
.cards2 { display:grid; gap:13px;
  grid-template-columns:repeat(auto-fit,minmax(228px,1fr)); }
.kpi2 { background:var(--panel); border:1px solid var(--line);
  border-radius:16px; padding:17px 19px 14px; box-shadow:var(--shadow);
  transition:transform .15s ease, box-shadow .15s ease; }
.kpi2:hover { transform:translateY(-2px);
  box-shadow:0 8px 22px rgba(16,24,40,.09); }
.kpi2 .top { display:flex; justify-content:space-between;
  align-items:center; }
.kpi2 .lbl { font-size:11.5px; font-weight:680; color:var(--muted);
  text-transform:uppercase; letter-spacing:.07em; }
.kpi2 .ic2 { width:33px; height:33px; border-radius:10px;
  background:#F1F5F9; color:#475467; display:flex;
  align-items:center; justify-content:center; }
.kpi2 .ic2.red { background:#FDECEA; color:#B42318; }
.kpi2 .ic2.green { background:#DCFCE7; color:#166534; }
.kpi2 .num { font-size:32px; font-weight:800; letter-spacing:-.025em;
  font-variant-numeric:tabular-nums; line-height:1.12; margin-top:7px; }
.kpi2 .num.red { color:var(--red); }
.kpi2 .foot { display:flex; align-items:flex-end;
  justify-content:space-between; gap:10px; margin-top:9px;
  min-height:30px; }
.kpi2 .note { font-size:12.5px; color:var(--muted); }
.kpi2 .note.red { color:#B42318; font-weight:620; }
.delta { font-size:12px; font-weight:720; padding:2px 8px;
  border-radius:20px; white-space:nowrap; }
.delta.up { background:#DCFCE7; color:#166534; }
.delta.down { background:#FDECEA; color:#B42318; }
.delta.flat { background:#EEF1F4; color:#667085; }
.spark { display:inline-flex; align-items:flex-end; gap:2.5px;
  height:30px; }
.spark i { width:6px; background:#CBD8DE; border-radius:2px;
  display:block; }
.spark i:last-child { background:var(--accent); }
.segbar { display:flex; height:9px; border-radius:6px;
  overflow:hidden; background:#EEF1F4; width:100%; }
.segbar i { display:block; height:100%; }
.fpct { width:46px; flex:none; color:var(--muted); font-size:12.5px; }
.ringwrap { display:flex; align-items:center; gap:13px; }
@media (max-width:1080px) { .detail-grid { grid-template-columns:1fr; }
  .searchbox input { width:200px; } }
.kitmsg { background:#F7F5FC; border:1px solid #E9E4F5;
  border-radius:12px; padding:13px 15px; font-size:14px;
  line-height:1.55; white-space:pre-wrap; color:#2A2440; }
.kitsteps { list-style:none; margin:12px 0 0; padding:0; }
.kitsteps li { display:flex; gap:11px; padding:10px 0;
  align-items:flex-start; }
.kitsteps li + li { border-top:1px solid #EEF1F4; }
.kitsteps .n { flex:none; width:24px; height:24px; border-radius:50%;
  background:var(--accent); color:#fff; font-weight:700; font-size:13px;
  display:flex; align-items:center; justify-content:center;
  margin-top:1px; }
.kitthumb { max-height:130px; border-radius:9px;
  border:1px solid var(--line); display:block; margin-top:7px; }
.kitrow.sent { opacity:.5; }
.kittag { background:#DCFCE7; color:#15803D; border-radius:20px;
  padding:1px 8px; font-size:11.5px; font-weight:700;
  white-space:nowrap; }
"""


def page(title, active, body, conn=None):
    badge = ""
    if conn is not None:
        overdue, due_today, _ = due_leads(conn)
        n = len(overdue) + len(due_today)
        if n:
            badge = f'<span class="navbadge">{n}</span>'
    total = ""
    if conn is not None:
        c = conn.execute("SELECT COUNT(*) c FROM leads").fetchone()["c"]
        total = f"<b>{c}</b> leads in the database"
    kitbadge = ""
    if conn is not None:
        nk = kit_pending_count(conn)
        if nk:
            kitbadge = f'<span class="navbadge">{nk}</span>'

    def nav(key, href, ic, label, extra=""):
        cls = ' class="active"' if active == key else ""
        return (f'<a href="{href}"{cls}>{icon(ic)}'
                f'<span>{label}</span>{extra}</a>')

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} - Vaani CRM</title><style>{CSS}</style></head>
<body>
<aside class="sidebar">
  <div class="brand"><div class="logo">V</div>
    <div><div class="t1">Vaani CRM</div>
    <div class="t2">Lead manager</div></div></div>
  <nav>
    {nav("dashboard", "/", "grid", "Dashboard")}
    {nav("call", "/call", "phone", "Call Mode")}
    {nav("kit", "/kit", "send", "Send Kit", kitbadge)}
    {nav("board", "/board", "board", "Pipeline")}
    {nav("leads", "/leads", "users", "All Leads")}
    {nav("followups", "/followups", "clock", "Follow-ups", badge)}
    {nav("export", "/export.csv", "download", "Export CSV")}
  </nav>
  <div class="side-foot">{total}<br>Vaani CRM v{VERSION}<br>
  Data stays on this PC</div>
</aside>
<div class="shell">
  <div class="topbar"><span class="ptitle">{esc(title)}</span>
    <form class="searchbox" method="get" action="/leads">
      {icon("search", 15)}
      <input id="globalsearch" name="q" placeholder="Search leads...">
      <span class="kbd">Ctrl K</span>
    </form>
  </div>
  <main>{body}</main>
  <footer>Vaani CRM v{VERSION} - runs only on this computer. Your data
  never leaves this PC.</footer>
</div>
<script>
document.addEventListener('keydown', function(e) {{
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {{
    e.preventDefault();
    document.getElementById('globalsearch').focus();
  }}
}});
</script>
</body></html>"""


# ---------------------------------------------------------------- pages

def status_options(selected):
    out = []
    for s in STATUSES:
        sel = " selected" if s == selected else ""
        out.append(f'<option value="{esc(s)}"{sel}>{esc(s)}</option>')
    return "".join(out)


def greeting():
    h = datetime.now().hour
    if h < 12:
        return "Good morning"
    if h < 17:
        return "Good afternoon"
    return "Good evening"


def task_card(r, kind="overdue"):
    style = ""
    cls = "taskcard"
    if kind == "today":
        cls += " amber"
    elif kind == "later":
        style = ' style="border-left-color:#D0D5DD"'
    return (
        f'<div class="{cls}"{style}>{monogram(r["name"], r["phone"])}'
        f'<div class="grow"><a href="/lead?phone={esc(r["phone"])}">'
        f'<b>{esc(r["name"])}</b></a> '
        f'<span class="small">{fmt_phone(r["phone"])}</span><br>'
        f'{followup_chip(r)} {tier_badge(r["tier"])} '
        f'<span class="small">Score {esc(r["score"] or "-")} - '
        f'{esc(r["angle"] or "")}</span></div>'
        f'{status_pill(r["status"])}'
        f'<a class="btn ghost" href="/lead?phone={esc(r["phone"])}">'
        f'Open</a></div>')


def delta_chip(cur, prev):
    if prev == 0:
        if cur > 0:
            return f'<span class="delta up">&#9650; {cur} new</span>'
        return '<span class="delta flat">- no change</span>'
    pct = round((cur - prev) * 100 / prev)
    if pct > 0:
        return (f'<span class="delta up">&#9650; {pct}%</span>'
                f' <span class="note">vs last week</span>')
    if pct < 0:
        return (f'<span class="delta down">&#9660; {abs(pct)}%</span>'
                f' <span class="note">vs last week</span>')
    return '<span class="delta flat">- same as last week</span>'


def spark_bars(values):
    mx = max(values) if values else 0
    bars = ""
    for v in values:
        h = 3 if mx == 0 else max(3, round(v * 27 / mx))
        bars += f'<i style="height:{h}px"></i>'
    return f'<span class="spark">{bars}</span>'


def dashboard_html(conn):
    total = conn.execute("SELECT COUNT(*) c FROM leads").fetchone()["c"]
    today = date.today()
    date_line = datetime.now().strftime("%A, %d %B %Y")
    if total == 0:
        return (f"<h1>{greeting()}, {esc(USER_NAME)}</h1>"
                f'<div class="sub">{date_line}</div>'
                '<div class="empty" style="margin-top:18px">No leads '
                'yet.<br>Run Vaani once (double-click <b>run-vaani.bat'
                '</b>) and every restaurant it scrapes will appear '
                'here automatically.</div>')
    counts = {}
    for row in conn.execute(
            "SELECT COALESCE(s.status, 'Not called') st, COUNT(*) c"
            " FROM leads l LEFT JOIN crm_state s ON s.phone = l.phone"
            " GROUP BY st"):
        counts[row["st"]] = row["c"]
    tiers = {}
    for row in conn.execute(
            "SELECT tier, COUNT(*) c FROM leads GROUP BY tier"):
        if row["tier"]:
            tiers[row["tier"]] = row["c"]
    tier_a = tiers.get("A", 0)
    demos = counts.get("Demo booked", 0) + counts.get("Demo done", 0)
    won = counts.get("Won", 0)
    overdue, due_today, _ = due_leads(conn)

    # new-lead trend: sparkline (14 days) + week-on-week delta
    daily = {}
    for row in conn.execute(
            "SELECT substr(first_seen, 1, 10) d, COUNT(*) c FROM leads"
            " WHERE first_seen >= ? GROUP BY d",
            ((today - timedelta(days=13)).isoformat(),)):
        daily[row["d"]] = row["c"]
    vals = [daily.get((today - timedelta(days=i)).isoformat(), 0)
            for i in range(13, -1, -1)]
    new_cur = sum(vals[7:])
    new_prev = sum(vals[:7])

    # demos won this week vs last week (from the status timeline)
    def demo_events(since, until):
        return conn.execute(
            "SELECT COUNT(*) c FROM crm_remarks WHERE source ="
            " 'status' AND (text LIKE '%\u2192 Demo booked%' OR"
            " text LIKE '%\u2192 Demo done%') AND ts >= ? AND ts < ?",
            (since, until)).fetchone()["c"]
    wk = (today - timedelta(days=6)).isoformat()
    pw = (today - timedelta(days=13)).isoformat()
    tm = (today + timedelta(days=1)).isoformat()
    demos_cur = demo_events(wk, tm)
    demos_prev = demo_events(pw, wk)

    closed = sum(counts.get(s, 0) for s in
                 ("Won", "Lost", "Not interested", "Wrong number",
                  "Do not contact"))
    won_pct = round(won * 100 / closed) if closed else 0
    contacted = total - counts.get("Not called", 0)
    contacted_pct = round(contacted * 100 / total) if total else 0

    need = len(overdue) + len(due_today)
    hstats = (
        f'<div class="hstats">'
        f'<div class="hstat"><div class="n">{need}</div>'
        f'<div class="l">follow-ups need you today</div></div>'
        f'<div class="hstat"><div class="n">'
        f'{counts.get("Not called", 0)}</div>'
        f'<div class="l">leads never called yet</div></div>'
        f'<div class="hstat"><div class="n">{tier_a}</div>'
        f'<div class="l">A-tier leads in the bank</div></div>'
        f'<div class="hstat"><div class="n">{demos}</div>'
        f'<div class="l">demos in motion</div></div></div>')
    hero = (
        f'<div class="hero"><h1>{greeting()}, {esc(USER_NAME)}</h1>'
        f'<div class="hsub">{date_line} - here is where your leads '
        f'stand.</div>{hstats}'
        f'<div class="hcta"><a class="btn" href="/call">'
        f'{icon("phone", 16)} Start calling</a>'
        f'<a class="btn ghost" href="/board">View pipeline</a>'
        f'</div></div>')

    def kpi2(lbl, ic, num, foot, ic_cls="", num_cls=""):
        return (f'<div class="kpi2"><div class="top">'
                f'<span class="lbl">{lbl}</span>'
                f'<span class="ic2 {ic_cls}">{icon(ic, 17)}</span></div>'
                f'<div class="num {num_cls}">{num}</div>'
                f'<div class="foot">{foot}</div></div>')

    seg = ""
    tot_t = sum(tiers.values()) or 1
    for t, col in (("A", "#22C55E"), ("B", "#F59E0B"),
                   ("C", "#CBD5DC")):
        w = round(tiers.get(t, 0) * 100 / tot_t)
        if w:
            seg += f'<i style="width:{w}%;background:{col}"></i>'
    a_pct = round(tier_a * 100 / tot_t) if tot_t else 0
    card_total = kpi2(
        "Total leads", "users", total,
        f'<span>{delta_chip(new_cur, new_prev)}</span>'
        f'{spark_bars(vals)}')
    card_a = kpi2(
        "A-tier - call today", "flame", tier_a,
        f'<span class="segbar">{seg}</span>'
        f'<span class="note">{a_pct}% A</span>')
    if overdue:
        due_note = (f'<span class="note red">{len(overdue)} also '
                    f'overdue</span>')
    else:
        due_note = '<span class="note">Nothing overdue</span>'
    card_due = kpi2("Due today", "calendar", len(due_today),
                    due_note + '<span class="note">follow-ups</span>')
    if overdue:
        oldest = min(r["followup"] for r in overdue)
        od = (today - date.fromisoformat(oldest)).days
        od_note = f'<span class="note red">Oldest: {od}d overdue</span>'
    else:
        od_note = '<span class="note">All clear</span>'
    card_od = kpi2("Overdue", "clock", len(overdue),
                   od_note + '<span class="note">follow-ups</span>',
                   ic_cls="red" if overdue else "",
                   num_cls="red" if overdue else "")
    if demos_cur or demos_prev:
        demo_foot = f'<span>{delta_chip(demos_cur, demos_prev)}</span>'
    else:
        demo_foot = '<span class="note">Book one from Call Mode</span>'
    card_demo = kpi2("Demos booked / done", "phone", demos, demo_foot)
    if closed:
        won_foot = (f'<span class="note">{won_pct}% of closed leads '
                    f'won</span>')
    else:
        won_foot = '<span class="note">No closed leads yet</span>'
    card_won = kpi2("Won", "check", won, won_foot,
                    ic_cls="green" if won else "")
    body = hero + ('<div class="cards2">' + card_total + card_a
                   + card_due + card_od + card_demo + card_won
                   + '</div>')
    if overdue:
        body += f"<h2>Overdue - clear these first ({len(overdue)})</h2>"
        body += "".join(task_card(r, "overdue") for r in overdue)
    if due_today:
        body += f"<h2>Due today ({len(due_today)})</h2>"
        body += "".join(task_card(r, "today") for r in due_today)
    if not overdue and not due_today:
        body += ('<div class="empty" style="margin-top:20px">No '
                 'follow-ups due. Set a follow-up date on any lead '
                 'and it will appear here on the day.</div>')
    maxc = max(counts.values()) if counts else 1
    rows = ""
    for s in STATUSES:
        c = counts.get(s, 0)
        if not c:
            continue
        bar_cls = "fb"
        if s == "Won":
            bar_cls += " good"
        elif s in ("Lost", "Not interested", "Wrong number",
                   "Do not contact"):
            bar_cls += " bad"
        pct = round(c * 100 / total) if total else 0
        rows += (f'<div class="funnel-row"><div class="fl">'
                 f'{status_pill(s)}</div><div class="ft">'
                 f'<div class="{bar_cls}" style="width:'
                 f'{max(3, round(c * 100 / maxc))}%"></div></div>'
                 f'<div class="fc">{c}</div>'
                 f'<div class="fpct">{pct}%</div></div>')
    circ = 157.1
    off = circ * (1 - contacted_pct / 100)
    ring = (
        f'<div class="ringwrap"><svg width="66" height="66" '
        f'viewBox="0 0 66 66"><circle cx="33" cy="33" r="25" '
        f'fill="none" stroke="#EEF1F4" stroke-width="8"/>'
        f'<circle cx="33" cy="33" r="25" fill="none" '
        f'stroke="#0F766E" stroke-width="8" stroke-linecap="round" '
        f'stroke-dasharray="{circ}" stroke-dashoffset="{off:.1f}" '
        f'transform="rotate(-90 33 33)"/>'
        f'<text x="33" y="34" text-anchor="middle" '
        f'dominant-baseline="central" font-size="14.5" '
        f'font-weight="800" fill="#101828">{contacted_pct}%</text>'
        f'</svg><div><b>Leads contacted</b><div class="small">'
        f'{contacted} of {total} leads have had a first call or '
        f'beyond</div></div></div>')
    body += (f'<h2>Your funnel</h2><div class="panel">'
             f'<div style="margin-bottom:12px">{ring}</div>'
             f'{rows}</div>')
    recents = conn.execute(
        "SELECT r.*, l.name FROM crm_remarks r JOIN leads l"
        " ON l.phone = r.phone WHERE r.source != 'status'"
        " ORDER BY r.id DESC LIMIT 6").fetchall()
    if recents:
        items = "".join(
            f'<div class="remark"><div class="when">'
            f'{esc(r["ts"][:10])} - <a href="/lead?phone='
            f'{esc(r["phone"])}">{esc(r["name"])}</a></div>'
            f'{esc(r["text"][:170])}</div>' for r in recents)
        body += f"<h2>Latest remarks</h2>{items}"
    return body


def lead_row(r, next_url, kit_ts=None):
    kittag = (f' <span class="kittag">Kit sent '
              f'{esc(fmt_date(str(kit_ts)[:10]))}</span>'
              if kit_ts else "")
    remark = esc((r["last_remark"] or "")[:60])
    if r["n_remarks"]:
        remark += f' <span class="small">({r["n_remarks"]})</span>'
    fu = followup_chip(r) or fmt_date(r["followup"])
    return (
        f'<tr><td><div class="name-cell">'
        f'{monogram(r["name"], r["phone"])}<div>'
        f'<a class="nm" href="/lead?phone={esc(r["phone"])}">'
        f'{esc(r["name"])}</a><div class="an">'
        f'{esc(r["angle"] or "")}{kittag}</div></div></div></td>'
        f'<td>{fmt_phone(r["phone"])}</td>'
        f'<td class="small">{esc((r["address"] or "")[:54])}</td>'
        f'<td>{tier_badge(r["tier"])} <b>{esc(r["score"] or "-")}</b></td>'
        f'<td><form method="post" action="/update" style="margin:0">'
        f'<input type="hidden" name="phone" value="{esc(r["phone"])}">'
        f'<input type="hidden" name="followup" '
        f'value="{esc(r["followup"])}">'
        f'<input type="hidden" name="next" value="{esc(next_url)}">'
        f'<select name="status" onchange="this.form.submit()" '
        f'style="padding:5.5px 7px;font-size:13.5px">'
        f'{status_options(r["status"])}</select></form></td>'
        f'<td>{fu}</td>'
        f'<td class="small">{remark}</td>'
        f'<td class="small">{esc(r["times_seen"] or 1)}x</td></tr>')


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
        f'<option value="{v}"{" selected" if sort == v else ""}">'
        f'{lbl}</option>'
        for v, lbl in (("score", "Best score first"),
                       ("followup", "Follow-up date"),
                       ("recent", "Recently seen"),
                       ("name", "Name")))
    body = f"""<h1>All leads</h1>
<div class="sub">{len(rows)} lead(s) shown - the pipeline board and
this list always agree.</div>
<form class="filters" method="get" action="/leads">
<div><label>Search</label>
<input name="q" value="{esc(q)}" size="28"
placeholder="Name, phone, area, email"></div>
<div><label>Tier</label><select name="tier">{tier_opts}</select></div>
<div><label>Status</label><select name="status">{status_opts}</select></div>
<div><label>Sort</label><select name="sort">{sort_opts}</select></div>
<div><button class="btn" type="submit">Filter</button>
<a class="btn ghost" href="/leads">Clear</a></div></form>"""
    if not rows:
        body += ('<div class="empty">No leads match. Clear the '
                 'filters, or run Vaani to scrape more.</div>')
        return body
    sent_map = kit_sent_map(conn)
    trs = "".join(lead_row(r, next_url, sent_map.get(r["phone"]))
                  for r in rows)
    body += (f"<table><tr><th>Restaurant</th><th>Phone</th>"
             f"<th>Location</th><th>Score</th><th>Status</th>"
             f"<th>Follow-up</th><th>Last remark</th><th>Seen</th></tr>"
             f"{trs}</table>")
    return body


def kit_panel(conn, lead):
    phone = lead["phone"]
    head = '<h2 style="margin-top:0">WhatsApp kit</h2>'
    reason = kit_skip_reason(lead)
    if reason:
        return (f'<div class="panel" style="margin-top:16px">{head}'
                '<div class="small">On the don\'t-pitch list: '
                f'{esc(reason)}. No kit needed for this one.</div>'
                '</div>')
    sent = kit_sent_at(conn, phone)
    sent_line = ""
    if sent:
        sent_line = (f'<div style="color:#15803D;font-weight:600;'
                     f'margin-bottom:9px">Kit sent on '
                     f'{esc(fmt_ts(sent))}. You can re-send any '
                     'time below.</div>')
    msg = kit_message(lead)
    card = kit_card_file(phone)
    if card:
        card_html = (f'<img class="kitthumb" src="/card?phone='
                     f'{esc(phone)}"><div class="small">'
                     f'{esc(os.path.basename(card))} - in the '
                     'report-cards folder next to the CRM.</div>')
    elif os.path.exists(REPORT_CARD_SCRIPT):
        card_html = (f'<form method="post" action="/kit-generate" '
                     f'style="margin:0"><input type="hidden" '
                     f'name="phone" value="{esc(phone)}">'
                     f'<input type="hidden" name="next" '
                     f'value="/lead?phone={esc(phone)}">'
                     f'<button class="btn ghost" type="submit">'
                     f'Generate the report card</button></form>'
                     f'<div class="small">Takes a few seconds - it '
                     f'runs the Report Card tool for this lead.</div>')
    else:
        card_html = ('<div class="small">report_card.py is not in '
                     'this folder, so no card yet - the intro and '
                     'brochure still work.</div>')
    if os.path.exists(BROCHURE_PATH):
        bro_html = (f'Attach <b>book-my-food-brochure.pdf</b> (in '
                    f'this folder, or <a href="/brochure">download '
                    f'it</a>), then send the follow-up text: '
                    f'<a class="btn ghost" target="_blank" '
                    f'href="{esc(wa_link(phone, kit_brochure_message(lead)))}">'
                    f'Brochure message</a>')
    else:
        bro_html = ('<div class="small">The brochure PDF is not in '
                    'this folder - skip this step.</div>')
    if os.path.exists(BROCHURE_PATH):
        auto_block = f"""<div style="background:#ECFDF5;border:1px solid #A6F4C5;border-radius:12px;padding:14px 16px;margin-bottom:14px">
<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
<button class="btn" type="button" onclick="autosend(this,'{esc(phone)}')">Send kit automatically</button>
<span class="small">Sends in order: the intro message, then the report card, then the brochure.</span></div>
<div class="small" style="margin-top:7px">Opens WhatsApp Web in a Chrome window and sends it for you, then logs it here and sets tomorrow's follow-up. The first time only, scan the QR code in that window with your phone (WhatsApp &gt; Linked devices). Keep the window open until it finishes.</div></div>"""
    else:
        auto_block = ('<div class="small" style="margin-bottom:12px">'
                      'Auto-send is off: the brochure PDF is not in '
                      'this folder, and the kit only goes out '
                      'complete.</div>')
    return f"""<div class="panel" style="margin-top:16px">{head}
{sent_line}{auto_block}<div class="kitmsg">{esc(msg)}</div>
<div class="small" style="margin:14px 0 0"><b>Or send it yourself, step by step:</b></div>
<ol class="kitsteps">
<li><span class="n">1</span><div>
<a class="btn" target="_blank" href="{esc(wa_link(phone, msg))}">Open WhatsApp with this intro</a>
<div class="small">Send this message first. Then attach the report card (step 2) and the brochure (step 3) in the same chat, in that order.</div></div></li>
<li><span class="n">2</span><div><b>Attach the report card</b>{card_html}</div></li>
<li><span class="n">3</span><div><b>Send the brochure</b><br>{bro_html}</div></li>
<li><span class="n">4</span><div><form method="post" action="/kit-sent" style="margin:0">
<input type="hidden" name="phone" value="{esc(phone)}">
<input type="hidden" name="next" value="/lead?phone={esc(phone)}">
<button class="btn" type="submit">Mark kit as sent</button>
<span class="small">Logs it on this lead and sets a call follow-up for tomorrow.</span>
</form></div></li>
</ol></div>""" + AUTOSEND_JS


def kit_html(conn):
    sent_map = kit_sent_map(conn)
    today = date.today().isoformat()
    sent_today = [(p, ts) for p, ts in sent_map.items()
                  if str(ts).startswith(today)]
    rows = conn.execute(
        LEAD_SELECT + " WHERE l.suppressed = 0 ORDER BY"
        " (l.score IS NULL), l.score DESC, l.reviews DESC").fetchall()
    batch = []
    for r in rows:
        if kit_skip_reason(r) or r["phone"] in sent_map:
            continue
        batch.append(r)
        if len(batch) >= 30:
            break
    ready = sum(1 for r in batch if kit_card_file(r["phone"]))
    body = f"""<h1>Send the WhatsApp kit</h1>
<div class="sub">Your top {len(batch)} pitchable leads who have not
received the kit yet - best scores first. <b>Auto-send kit</b> opens
WhatsApp Web and sends, in order, the intro message, his report
card and the brochure, then logs it and sets tomorrow's call
follow-up. The
WhatsApp button is the manual way (message prefilled, you attach the
files). Pace yourself like a person - a few at a time, not all of
them in one minute.</div>
<div class="panel" style="margin:16px 0">
<b><span id="sentcount">{len(sent_today)}</span> kit(s) sent today</b>
<span class="small"> - {ready} of these {len(batch)} leads already
have a report card ready; generate the rest as you go.</span></div>"""
    if not batch:
        body += ('<div class="empty">Nothing waiting - every '
                 'pitchable lead already has the kit. Run Vaani to '
                 'scrape more leads.</div>')
    for r in batch:
        card = kit_card_file(r["phone"])
        if card:
            card_bit = '<span class="kittag">Card ready</span>'
            gen_bit = ""
        else:
            card_bit = '<span class="small">No card yet</span>'
            gen_bit = (f'<form method="post" action="/kit-generate" '
                       f'style="display:inline;margin:0">'
                       f'<input type="hidden" name="phone" '
                       f'value="{esc(r["phone"])}">'
                       f'<input type="hidden" name="next" '
                       f'value="/kit"><button class="btn ghost" '
                       f'type="submit">Generate card</button></form>')
        body += (
            f'<div class="taskcard kitrow">'
            f'{monogram(r["name"], r["phone"])}'
            f'<div class="grow"><a href="/lead?phone='
            f'{esc(r["phone"])}"><b>{esc(r["name"])}</b></a> '
            f'<span class="small">{fmt_phone(r["phone"])}</span><br>'
            f'<span class="small">{esc(kit_area(r))} - '
            f'{esc(r["angle"] or "")}</span> {tier_badge(r["tier"])} '
            f'<span class="small">Score {esc(r["score"] or "-")}'
            f'</span> {card_bit}</div>'
            f'<button class="btn" onclick="autosend(this,'
            f'\'{esc(r["phone"])}\')">Auto-send kit</button>'
            f'<a class="btn ghost" target="_blank" '
            f'href="{esc(wa_link(r["phone"], kit_message(r)))}">'
            f'WhatsApp</a>{gen_bit}'
            f'<button class="btn ghost" onclick="markSent(this,'
            f'\'{esc(r["phone"])}\')">Mark sent</button></div>')
    if sent_today:
        items = ""
        for p, ts in sorted(sent_today, key=lambda x: x[1]):
            lead = fetch_lead(conn, p)
            nm = esc(lead["name"]) if lead else esc(p)
            items += (f'<div class="small" style="padding:3px 0">'
                      f'<a href="/lead?phone={esc(p)}">{nm}</a> - '
                      f'{esc(fmt_ts(ts))}</div>')
        body += (f'<div class="panel" style="margin-top:16px">'
                 f'<h2 style="margin-top:0">Sent today</h2>'
                 f'{items}</div>')
    body += """<script>
function markSent(btn, phone) {
  fetch('/api/kit-sent', {method: 'POST',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body: 'phone=' + encodeURIComponent(phone)})
  .then(r => r.json()).then(j => {
    if (j.ok) {
      btn.closest('.kitrow').classList.add('sent');
      btn.outerHTML = '<span class="kittag">Sent</span>';
      var c = document.getElementById('sentcount');
      c.textContent = parseInt(c.textContent) + 1;
    }
  });
}
</script>"""
    body += AUTOSEND_JS
    return body


def lead_html(conn, lead):
    phone = lead["phone"]
    remarks = fetch_remarks(conn, phone)
    kitp = kit_panel(conn, lead)
    wa = f"https://wa.me/91{phone}" if len(str(phone)) == 10 else "#"
    timeline = ""
    n_notes = 0
    for r in remarks:
        when = r["ts"]
        try:
            when = datetime.fromisoformat(when).strftime(
                "%d %b %Y, %I:%M %p")
        except Exception:
            pass
        if r["source"] == "status":
            timeline += (f'<div class="event">&#9679; '
                         f'{esc(r["text"])}<br>'
                         f'<span class="small">{esc(when)}</span></div>')
            continue
        n_notes += 1
        tag = (' <span class="tag">from tracker</span>'
               if r["source"] == "tracker" else
               ' <span class="tag">whatsapp kit</span>'
               if r["source"] == "kit" else "")
        timeline += (
            f'<div class="remark"><div class="when">{esc(when)}{tag}'
            f'</div>{esc(r["text"]).replace(chr(10), "<br>")}</div>')
    if not remarks:
        timeline = '<div class="empty">Nothing here yet. Add the ' \
                   'first remark after your call.</div>'
    supp = ""
    if lead["suppressed"]:
        supp = (f'<div class="v" style="color:#B42318">Kept out of '
                f'Vaani results ({esc(lead["suppress_reason"] or "")})'
                f'</div>')
    body = f"""<div class="crumbs"><a href="/leads">All Leads</a> /
{esc(lead['name'])}</div>
<h1>{esc(lead['name'])}</h1>
<div class="sub">{status_pill(lead['status'])} &nbsp;{followup_chip(lead)}</div>
<div class="detail-grid" style="margin-top:16px">
<div class="panel">
  <div class="k">Phone</div><div class="v">{fmt_phone(phone)}</div>
  <div class="k">Email</div><div class="v">{esc(lead['email'] or '-')}</div>
  <div class="k">Location</div><div class="v">{esc(lead['address'] or '-')}</div>
  <div class="k">Website</div><div class="v">"""
    body += (f'<a href="{esc(lead["website"])}" target="_blank">'
             f'{esc(lead["website"])}</a>' if lead["website"] else "-")
    body += f"""</div>
  <div class="k">Rating / Reviews</div>
  <div class="v">{esc(lead['rating'] or '-')} / {esc(lead['reviews'] or '-')}</div>
  <div class="k">Score / Tier</div>
  <div class="v">{tier_badge(lead['tier'])} <b>{esc(lead['score'] or '-')}</b></div>
  <div class="k">Pitch angle</div><div class="v">{esc(lead['angle'] or '-')}</div>
  <div class="k">First seen / Last seen / Times seen</div>
  <div class="v">{fmt_date(lead['first_seen'])} / {fmt_date(lead['last_seen'])} / {esc(lead['times_seen'] or 1)}</div>
  {supp}
</div>
<div>
  <form method="post" action="/remark" style="margin-bottom:13px">
    <input type="hidden" name="phone" value="{esc(phone)}">
    <textarea name="text" rows="3" style="width:100%"
    placeholder="Add a remark - what happened on the call? Who did you speak to?"></textarea>
    <div style="margin-top:7px"><button class="btn" type="submit">Add remark</button></div>
  </form>
  {timeline}
</div>
<div class="panel">
  <h2 style="margin-top:0">Next step</h2>
  <div style="display:flex;gap:8px;margin-bottom:14px">
    <a class="btn big" style="flex:1;justify-content:center" href="tel:+91{esc(phone)}">Call</a>
    <a class="btn ghost big" style="flex:1;justify-content:center" href="{wa}" target="_blank">WhatsApp</a>
  </div>
  <form method="post" action="/update">
    <input type="hidden" name="phone" value="{esc(phone)}">
    <input type="hidden" name="next" value="/lead?phone={esc(phone)}">
    <div class="k">Status</div>
    <select name="status" style="width:100%;margin:4px 0 12px">{status_options(lead['status'])}</select>
    <div class="k">Follow-up date</div>
    <input type="date" id="fu" name="followup" value="{esc(lead['followup'])}" style="width:100%;margin:4px 0 9px">
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:13px">
      <button type="button" class="btn ghost" onclick="setFU(1)">Tomorrow</button>
      <button type="button" class="btn ghost" onclick="setFU(3)">+3 days</button>
      <button type="button" class="btn ghost" onclick="setFU(7)">+1 week</button>
      <button type="button" class="btn ghost" onclick="clearFU()">No date</button>
    </div>
    <button class="btn" type="submit" style="width:100%;justify-content:center">Save</button>
  </form>
</div>
</div>
{kitp}
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
    body = ("<h1>Follow-ups</h1><div class='sub'>Every lead with a "
            "date, in the order to work them.</div>")
    if overdue:
        body += f"<h2>Overdue ({len(overdue)})</h2>" + "".join(
            task_card(r, "overdue") for r in overdue)
    if due_today:
        body += f"<h2>Due today ({len(due_today)})</h2>" + "".join(
            task_card(r, "today") for r in due_today)
    if week:
        body += f"<h2>Coming this week ({len(week)})</h2>" + "".join(
            task_card(r, "later") for r in week)
    if later:
        body += f"<h2>Later ({len(later)})</h2>" + "".join(
            task_card(r, "later") for r in later)
    if not (overdue or due_today or upcoming):
        body += ('<div class="empty" style="margin-top:16px">No '
                 'follow-ups set. Open any lead and set a follow-up '
                 'date - it shows up here on the day.</div>')
    return body


def board_html(conn):
    rows = fetch_leads(conn)
    by_status = {s: [] for s in STATUSES}
    for r in rows:
        by_status[r["status"]].append(r)
    cols = ""
    for s in STATUSES:
        cards = ""
        for r in by_status[s]:
            cards += (
                f'<div class="dcard" draggable="true" '
                f'data-phone="{esc(r["phone"])}">'
                f'<div class="nm">{esc(r["name"])}</div>'
                f'<div class="meta">{tier_badge(r["tier"])}'
                f'<span class="chip chip-gray">Score '
                f'{esc(r["score"] or "-")}</span>'
                f'{followup_chip(r)}</div>'
                f'<div class="small" style="margin-top:6px">'
                f'{fmt_phone(r["phone"])} - '
                f'{esc((r["address"] or "")[:34])}</div></div>')
        if not cards:
            cards = '<div class="small" style="padding:2px 4px">' \
                    'No leads here</div>'
        cols += (f'<div class="col" data-status="{esc(s)}">'
                 f'<h3>{esc(s)}<span class="cnt">'
                 f'{len(by_status[s])}</span></h3>{cards}</div>')
    body = f"""<h1>Pipeline</h1>
<div class="sub">Your whole funnel on one board - drag a card to a new
stage and it saves instantly. This board and All Leads always agree.</div>
<div class="board" style="margin-top:14px">{cols}</div>
<script>
document.querySelectorAll('.dcard').forEach(function(c) {{
  c.addEventListener('dragstart', function(e) {{
    e.dataTransfer.setData('text/plain', c.dataset.phone);
  }});
}});
document.querySelectorAll('.col').forEach(function(col) {{
  col.addEventListener('dragover', function(e) {{
    e.preventDefault(); col.classList.add('drop');
  }});
  col.addEventListener('dragleave', function() {{
    col.classList.remove('drop');
  }});
  col.addEventListener('drop', function(e) {{
    e.preventDefault(); col.classList.remove('drop');
    var phone = e.dataTransfer.getData('text/plain');
    var card = document.querySelector(
      '.dcard[data-phone="' + phone + '"]');
    if (!card || card.parentElement === col) return;
    var body = new URLSearchParams();
    body.set('phone', phone);
    body.set('status', col.dataset.status);
    fetch('/api/update', {{method: 'POST', body: body}})
      .then(function(r) {{ return r.json(); }})
      .then(function(j) {{
        if (j.ok) {{
          col.appendChild(card);
          document.querySelectorAll('.col').forEach(function(c2) {{
            c2.querySelector('.cnt').textContent =
              c2.querySelectorAll('.dcard').length;
          }});
        }}
      }});
  }});
}});
</script>"""
    return body


def call_html(conn, pos):
    queue = call_queue(conn)
    total = len(queue)
    head = "<h1>Call Mode</h1>"
    if total == 0:
        return head + (
            '<div class="empty" style="margin-top:16px">Queue '
            'clear - nothing overdue, nothing due today, and no '
            'fresh A or B leads left uncalled.<br>Run Vaani to '
            'scrape a new area and this fills up again.</div>')
    if pos >= total:
        return head + (
            f'<div class="empty" style="margin-top:16px">You '
            f'reached the end of today\'s queue ({total} lead(s)).'
            f'<br><br><a class="btn" href="/">Back to Dashboard</a> '
            f'<a class="btn ghost" href="/call">Start again</a></div>')
    pos = max(0, pos)
    lead = queue[pos]
    phone = lead["phone"]
    today = date.today()
    if lead["followup"] and lead["followup"] < today.isoformat():
        why = "Overdue follow-up"
    elif lead["followup"] == today.isoformat():
        why = "Follow-up due today"
    else:
        why = f'Fresh {esc(lead["tier"])}-tier lead'
    pct = round(pos * 100 / total)
    recent = [r for r in fetch_remarks(conn, phone)
              if r["source"] != "status"][:3]
    notes = "".join(
        f'<div class="remark"><div class="when">{esc(r["ts"][:10])}'
        f'</div>{esc(r["text"][:220])}</div>' for r in recent)
    if not notes:
        notes = '<div class="small">No remarks yet on this lead.</div>'
    wa = f"https://wa.me/91{phone}" if len(str(phone)) == 10 else "#"
    nxt = f"/call?pos={pos}"
    d1 = (today + timedelta(days=1)).isoformat()
    d2 = (today + timedelta(days=2)).isoformat()

    def outcome(label, status, fu, cls="btn ghost", style=""):
        return (
            f'<form method="post" action="/update">'
            f'<input type="hidden" name="phone" value="{esc(phone)}">'
            f'<input type="hidden" name="status" value="{esc(status)}">'
            f'<input type="hidden" name="followup" value="{fu}">'
            f'<input type="hidden" name="next" value="{nxt}">'
            f'<button class="{cls}" type="submit" style="{style}">'
            f'{label}</button></form>')

    body = f"""{head}
<div class="sub">Working today's queue - {why.lower()} first, then
fresh A and B leads.</div>
<div class="callwrap" style="margin-top:16px">
<div class="small" style="display:flex;justify-content:space-between">
<span><b>Lead {pos + 1} of {total}</b></span>
<a href="/call?pos={pos + 1}">Skip this lead &#8594;</a></div>
<div class="progress" style="margin:7px 0 14px"><div style="width:{pct}%"></div></div>
<div class="callcard">
  <div style="display:flex;gap:14px;align-items:center">
    {monogram(lead['name'], phone)}
    <div><div style="font-size:21px;font-weight:720">{esc(lead['name'])}</div>
    <div class="small">{esc(lead['address'] or '')}</div></div>
  </div>
  <div style="margin:13px 0;display:flex;gap:7px;flex-wrap:wrap">
    {tier_badge(lead['tier'])}
    <span class="chip chip-gray">Score {esc(lead['score'] or '-')}</span>
    <span class="chip chip-gray">{esc(lead['angle'] or '')}</span>
    {status_pill(lead['status'])} {followup_chip(lead)}
    <span class="chip chip-amber">{why}</span>
  </div>
  <div style="font-size:19px;font-weight:650;letter-spacing:.02em">{fmt_phone(phone)}</div>
  <div style="display:flex;gap:9px;margin:14px 0">
    <a class="btn big" style="flex:1;justify-content:center" href="tel:+91{esc(phone)}">Call now</a>
    <a class="btn ghost big" style="flex:1;justify-content:center" href="{wa}" target="_blank">WhatsApp</a>
  </div>
  <div class="small" style="margin-bottom:5px"><b>Last remarks</b></div>
  {notes}
  <form method="post" action="/remark" style="margin-top:10px">
    <input type="hidden" name="phone" value="{esc(phone)}">
    <input type="hidden" name="next" value="{nxt}">
    <textarea name="text" rows="2" style="width:100%" placeholder="Quick remark about this call..."></textarea>
    <div style="margin-top:6px"><button class="btn ghost" type="submit">Add remark</button>
    <a class="small" href="/lead?phone={esc(phone)}">Open full lead page</a></div>
  </form>
  <h2 style="margin:18px 0 4px">How did it go?</h2>
  <div class="outcomes">
    {outcome('No answer - retry tomorrow', 'No answer', d1)}
    {outcome('Callback - in 2 days', 'Callback', d2)}
    {outcome('Interested', 'Interested', '', 'btn')}
    {outcome('Demo booked', 'Demo booked', '', 'btn')}
    {outcome('Not interested', 'Not interested', '', 'btn ghost', 'color:#B42318')}
    {outcome('Wrong number', 'Wrong number', '', 'btn ghost', 'color:#B42318')}
  </div>
</div></div>"""
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

    def _json(self, obj):
        import json
        self._send(200, json.dumps(obj),
                   ctype="application/json; charset=utf-8")

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
                self._send(200, page("Dashboard", "dashboard",
                                     dashboard_html(conn), conn))
            elif parsed.path == "/call":
                try:
                    pos = int(params.get("pos", 0))
                except Exception:
                    pos = 0
                self._send(200, page("Call Mode", "call",
                                     call_html(conn, pos), conn))
            elif parsed.path == "/board":
                self._send(200, page("Pipeline", "board",
                                     board_html(conn), conn))
            elif parsed.path == "/leads":
                self._send(200, page("All Leads", "leads",
                                     leads_html(conn, params), conn))
            elif parsed.path == "/lead":
                lead = fetch_lead(conn, params.get("phone", ""))
                if not lead:
                    self._send(404, page(
                        "Not found", "",
                        "<h1>Lead not found</h1><div class='empty'>"
                        "That lead is not in the database. "
                        "<a href='/leads'>Back to all leads</a></div>",
                        conn))
                else:
                    self._send(200, page(lead["name"] or "Lead", "leads",
                                         lead_html(conn, lead), conn))
            elif parsed.path == "/kit":
                self._send(200, page("Send Kit", "kit",
                                     kit_html(conn), conn))
            elif parsed.path == "/card":
                path = kit_card_file(params.get("phone", ""))
                if path:
                    with open(path, "rb") as fh:
                        self._send(200, fh.read(),
                                   ctype="image/png")
                else:
                    self._send(404, page(
                        "Not found", "",
                        "<h1>No report card yet</h1>", conn))
            elif parsed.path == "/brochure":
                if os.path.exists(BROCHURE_PATH):
                    with open(BROCHURE_PATH, "rb") as fh:
                        self._send(
                            200, fh.read(), ctype="application/pdf",
                            extra={"Content-Disposition":
                                   'attachment; filename='
                                   '"book-my-food-brochure.pdf"'})
                else:
                    self._send(404, page(
                        "Not found", "",
                        "<h1>Brochure not in this folder</h1>", conn))
            elif parsed.path == "/followups":
                self._send(200, page("Follow-ups", "followups",
                                     followups_html(conn), conn))
            elif parsed.path == "/export.csv":
                self._send(200, export_csv_text(conn),
                           ctype="text/csv; charset=utf-8",
                           extra={"Content-Disposition":
                                  'attachment; filename='
                                  '"vaani-crm-export.csv"'})
            else:
                self._send(404, page("Not found", "",
                                     "<h1>Page not found</h1>", conn))
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
            if parsed.path == "/api/update":
                phone = form.get("phone", "")
                ok = False
                if fetch_lead(conn, phone):
                    ok = set_status(conn, phone,
                                    form.get("status", ""),
                                    form.get("followup"))
                self._json({"ok": bool(ok)})
            elif parsed.path == "/update":
                phone = form.get("phone", "")
                if fetch_lead(conn, phone):
                    set_status(conn, phone, form.get("status", ""),
                               form.get("followup", ""))
                self._redirect(form.get("next") or
                               f"/lead?phone={phone}")
            elif parsed.path == "/api/kit-autosend":
                phone = form.get("phone", "")
                ok, msg = autosend_kit(conn, phone)
                if ok:
                    mark_kit_sent(conn, phone)
                self._json({"ok": ok, "error": "" if ok else msg})
            elif parsed.path == "/api/kit-sent":
                phone = form.get("phone", "")
                ok = False
                if fetch_lead(conn, phone):
                    mark_kit_sent(conn, phone)
                    ok = True
                self._json({"ok": ok})
            elif parsed.path == "/kit-sent":
                phone = form.get("phone", "")
                if fetch_lead(conn, phone):
                    mark_kit_sent(conn, phone)
                self._redirect(form.get("next") or
                               f"/lead?phone={phone}")
            elif parsed.path == "/kit-generate":
                phone = form.get("phone", "")
                if fetch_lead(conn, phone):
                    generate_card(phone)
                self._redirect(form.get("next") or
                               f"/lead?phone={phone}")
            elif parsed.path == "/remark":
                phone = form.get("phone", "")
                if fetch_lead(conn, phone):
                    add_remark(conn, phone, form.get("text", ""))
                self._redirect(form.get("next") or
                               f"/lead?phone={phone}")
            else:
                self._send(404, page("Not found", "",
                                     "<h1>Page not found</h1>", conn))
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
    open_browser(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
