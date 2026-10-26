#!/usr/bin/env python3
"""Vaani CRM launcher and database-backed web UI.

The original CRM implementation is kept in ``vaani_crm_legacy.py`` so the
existing SQLite schema, Excel import, WhatsApp kit and healthcare section stay
available. This launcher adds the new UI, a small same-origin JSON API, and a
safe allow-listed self-updater that never touches local data files.
"""

import hashlib
import csv
import io
import json
import mimetypes
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from http.server import ThreadingHTTPServer

import vaani_crm_legacy as legacy


HERE = Path(__file__).resolve().parent
UI_ROOT = HERE / "vaani_ui"
VERSION = "2.2.0"
UPDATE_MANIFEST_URL = (
    "https://raw.githubusercontent.com/yash6969-io/"
    "vaani-lead-engine/main/update-manifest.json"
)

# Only code and UI files may be updated. Databases, Excel trackers, report
# output, WhatsApp profiles and generated assets are deliberately excluded.
UPDATE_ALLOWLIST = {
    "vaani_crm.py",
    "vaani_crm_legacy.py",
    "vaani.py",
    "report_card.py",
    "clinic_report.py",
    "kit_send.py",
    "crm_notify.py",
    "requirements.txt",
    "run-crm.bat",
    "run-crm-healthcare.bat",
    "vaani_ui/index.html",
    "vaani_ui/src/main.js",
    "vaani_ui/src/core.js",
    "vaani_ui/src/styles.css",
    "update-manifest.json",
    "CRM-README.md",
    "CLINIC-README.md",
}

CRM_PORTS = (8765, 8766, 8767, 8768, 8769)


def _workspace_label_from_slug(slug):
    return " ".join(part.capitalize() for part in slug.replace("_", "-").split("-"))


def _ensure_workspace_schema(conn, label):
    conn.execute("""CREATE TABLE IF NOT EXISTS leads (
        phone TEXT PRIMARY KEY, name TEXT, address TEXT, website TEXT,
        email TEXT, rating REAL, reviews INTEGER, score INTEGER,
        tier TEXT, angle TEXT, first_seen TEXT, last_seen TEXT,
        times_seen INTEGER DEFAULT 1, suppressed INTEGER DEFAULT 0,
        suppress_reason TEXT, suppressed_on TEXT)""")
    columns = {row[1] for row in conn.execute("PRAGMA table_info(leads)")}
    if "suppressed_on" not in columns:
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
    conn.execute(
        "INSERT OR IGNORE INTO crm_meta (key, value) VALUES ('workspace_label', ?)",
        (label,),
    )
    conn.commit()


def discover_workspaces():
    """Find built-in and scraper-created category databases."""
    records = []
    paths = [("restaurants", "Restaurants", HERE / "vaani_master.db"),
             ("healthcare", "Healthcare", HERE / "vaani_healthcare.db")]
    known = {path.resolve() for _, _, path in paths}
    for path in sorted(HERE.glob("vaani_*.db")):
        if path.name.endswith(".crm-backup") or path.resolve() in known:
            continue
        slug = path.stem[len("vaani_"):]
        if slug:
            paths.append((slug, _workspace_label_from_slug(slug), path))
    for key, fallback_label, path in paths:
        if not path.exists() and key not in ("restaurants", "healthcare"):
            continue
        try:
            conn = sqlite3.connect(path)
            conn.row_factory = sqlite3.Row
            _ensure_workspace_schema(conn, fallback_label)
            meta = conn.execute(
                "SELECT value FROM crm_meta WHERE key='workspace_label'"
            ).fetchone()
            label = (meta[0] if meta and meta[0] else fallback_label)
            conn.close()
            records.append({"key": key, "label": label, "path": path})
        except Exception:
            continue
    return records


def resolve_workspace(value=None):
    records = discover_workspaces()
    if not records:
        raise RuntimeError("No CRM workspace database is available")
    wanted = str(value or "").strip().lower()
    for record in records:
        if wanted in (record["key"].lower(), record["label"].lower()):
            return record
    return records[0]


def workspace_conn(record):
    conn = sqlite3.connect(record["path"])
    conn.row_factory = sqlite3.Row
    _ensure_workspace_schema(conn, record["label"])
    return conn


def workspace_rows(conn):
    return conn.execute(legacy.LEAD_SELECT +
                        " ORDER BY (l.score IS NULL), l.score DESC, l.reviews DESC").fetchall()


def _configure_legacy_for_workspace(record):
    """Point the legacy report/kit helpers at one selected workspace."""
    key = record["key"]
    if key not in ("restaurants", "healthcare"):
        return False
    healthcare = key == "healthcare"
    legacy.HEALTHCARE = healthcare
    legacy.SECTION = "healthcare" if healthcare else "food"
    legacy.SECTION_LABEL = record["label"]
    legacy.DB_PATH = str(record["path"])
    legacy.BACKUP_PATH = str(record["path"]) + ".crm-backup"
    legacy.TRACKER_PATH = str(HERE / ("vaani-healthcare-tracker.xlsx" if healthcare else "vaani-call-tracker.xlsx"))
    legacy.REPORT_CARD_SCRIPT = str(HERE / ("clinic_report.py" if healthcare else "report_card.py"))
    legacy.REPORT_CARDS_DIR = str(HERE / ("clinic-reports" if healthcare else "report-cards"))
    legacy.BROCHURE_PATH = str(HERE / "book-my-food-brochure.pdf")
    legacy.KIT_SEND_SCRIPT = str(HERE / "kit_send.py")
    return True


def _find_workspace_lead(conn, phone):
    phone = _phone_key(phone)
    return conn.execute(legacy.LEAD_SELECT + " WHERE l.phone = ?", (phone,)).fetchone()


def _kit_metadata(record, conn, phone):
    lead = _find_workspace_lead(conn, phone)
    if not lead:
        return {"available": False, "error": "Lead not found in this workspace."}
    if not _configure_legacy_for_workspace(record):
        return {
            "available": False,
            "supported": False,
            "message": "Report cards and kits are currently available for Restaurants and Healthcare workspaces.",
        }
    card = legacy.kit_card_file(lead["phone"])
    sent_at = legacy.kit_sent_at(conn, lead["phone"])
    skip_reason = legacy.kit_skip_reason(lead)
    return {
        "available": not bool(skip_reason),
        "supported": True,
        "cardReady": bool(card),
        "cardUrl": ("/api/report-card?workspace=" + urllib.parse.quote(record["key"]) +
                    "&phone=" + urllib.parse.quote(lead["phone"])) if card else "",
        "sentAt": sent_at or "",
        "skipReason": skip_reason or "",
        "message": legacy.kit_message(lead),
        "brochureAvailable": not legacy.HEALTHCARE and Path(legacy.BROCHURE_PATH).exists(),
        "brochureUrl": "/brochure?workspace=" + urllib.parse.quote(record["key"]),
    }


def _generate_report(record, phone):
    if not _configure_legacy_for_workspace(record):
        return False, "This workspace does not have a report generator yet."
    normalized = _phone_key(phone)
    script = Path(legacy.REPORT_CARD_SCRIPT)
    if not script.exists():
        return False, f"The report generator is missing: {script.name}."
    try:
        result = subprocess.run(
            [sys.executable, str(script), normalized, "--no-update"],
            cwd=HERE, timeout=180, capture_output=True, text=True,
        )
    except Exception as exc:
        return False, f"The report generator could not start: {exc}"
    card = legacy.kit_card_file(normalized)
    if card:
        return True, "Report generated successfully."
    detail = (result.stderr or result.stdout or "").strip().splitlines()
    return False, detail[-1] if detail else "The report generator finished without creating a report card."


def _kit_action(record, phone, action):
    conn = workspace_conn(record)
    try:
        lead = _find_workspace_lead(conn, phone)
        if not lead:
            return False, "Lead not found in this workspace."
        if not _configure_legacy_for_workspace(record):
            return False, "Kits are currently available for Restaurants and Healthcare only."
        if action == "generate":
            return _generate_report(record, lead["phone"])
        if action == "mark":
            legacy.mark_kit_sent(conn, lead["phone"])
            return True, "Kit marked as sent."
        if action == "autosend":
            ok, message = legacy.autosend_kit(conn, lead["phone"])
            if ok:
                legacy.mark_kit_sent(conn, lead["phone"])
            return ok, message
        return False, "Unknown kit action."
    finally:
        conn.close()


def _version_tuple(value):
    try:
        return tuple(int(part) for part in str(value).strip().split("."))
    except Exception:
        return (0,)


def _raw_url(relative_path):
    base = UPDATE_MANIFEST_URL.rsplit("/", 1)[0]
    return f"{base}/{urllib.parse.quote(relative_path, safe='/')}"


def _read_url(url, timeout=8):
    request = urllib.request.Request(url, headers={"User-Agent": "VaaniCRM"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def _safe_target(relative_path):
    relative_path = str(relative_path).replace("\\", "/").lstrip("/")
    if relative_path not in UPDATE_ALLOWLIST:
        return None
    target = (HERE / relative_path).resolve()
    if target != HERE and not str(target).startswith(f"{HERE}{os.sep}"):
        return None
    return target


def self_update():
    """Update all CRM code/UI files, then restart once.

    The manifest is fetched from the public repository. Each target is
    allow-listed, downloaded to a temporary sibling, optionally hash-checked,
    and syntax-checked before replacement. Local databases and trackers are
    never part of the update set.
    """
    if os.environ.get("VAANI_CRM_NO_UPDATE") == "1" or "--no-update" in sys.argv:
        return
    try:
        manifest = json.loads(_read_url(UPDATE_MANIFEST_URL).decode("utf-8"))
        remote_version = str(manifest.get("version", "0.0.0"))
        if _version_tuple(remote_version) <= _version_tuple(VERSION):
            return
        files = manifest.get("files", [])
        if not isinstance(files, list) or not files:
            return
    except Exception:
        return

    staged = []
    try:
        for item in files:
            if not isinstance(item, dict):
                return
            relative_path = item.get("path", "")
            target = _safe_target(relative_path)
            if target is None:
                return
            content = _read_url(_raw_url(relative_path), timeout=15)
            expected_hash = item.get("sha256")
            if expected_hash:
                actual_hash = hashlib.sha256(content).hexdigest()
                if actual_hash.lower() != str(expected_hash).lower():
                    return
            if target.suffix == ".py":
                compile(content.decode("utf-8"), str(target), "exec")
            staged_path = target.with_name(f".{target.name}.update")
            staged_path.write_bytes(content)
            staged.append((target, staged_path))
    except Exception:
        for _, staged_path in staged:
            staged_path.unlink(missing_ok=True)
        return

    try:
        for target, staged_path in staged:
            backup = target.with_name(f"{target.name}.pre-update")
            if target.exists():
                shutil.copyfile(target, backup)
            os.replace(staged_path, target)
    except Exception:
        return

    print(f"\n  Updated Vaani CRM to v{remote_version} - restarting...\n")
    env = dict(os.environ)
    env["VAANI_CRM_NO_UPDATE"] = "1"
    subprocess.Popen([sys.executable, str(Path(__file__).resolve()), *sys.argv[1:]], env=env)
    raise SystemExit(0)


def _json_safe_date(value):
    return str(value or "")[:10]


def _lead_id(phone):
    digits = re.sub(r"\D", "", str(phone or ""))
    return f"lead-{digits[-10:] or 'unknown'}"


def _phone_key(value):
    keys = legacy._phone_keys(value)
    return sorted(keys)[0] if keys else re.sub(r"\D", "", str(value or ""))[-10:]


def _signals(row):
    signals = []
    if row["rating"]:
        signals.append(f"Rated {row['rating']}")
    if row["reviews"]:
        signals.append(f"{row['reviews']:,} reviews")
    if row["website"]:
        signals.append("Website found")
    if row["email"]:
        signals.append("Email found")
    if row["angle"]:
        signals.append(str(row["angle"]))
    return signals[:5]


def _serialize_lead(conn, row, workspace_label, record=None):
    remarks = legacy.fetch_remarks(conn, row["phone"])
    kit = _kit_metadata(record, conn, row["phone"]) if record else {}
    return {
        "id": _lead_id(row["phone"]),
        "vertical": workspace_label,
        "workspaceKey": "",
        "name": row["name"] or "Unnamed lead",
        "contact": "",
        "location": row["address"] or "",
        "phone": row["phone"] or "",
        "email": row["email"] or "",
        "website": row["website"] or "",
        "rating": row["rating"] or 0,
        "reviews": row["reviews"] or 0,
        "score": row["score"] or 0,
        "tier": row["tier"] or "C",
        "angle": row["angle"] or "New opportunity",
        "status": row["status"] or "Not called",
        "followup": row["followup"] or "",
        "source": f"Google Maps · {_json_safe_date(row['first_seen'])}",
        "firstSeen": _json_safe_date(row["first_seen"]),
        "lastSeen": _json_safe_date(row["last_seen"]),
        "timesSeen": row["times_seen"] or 1,
        "signals": _signals(row),
        "notes": [
            {
                "id": f"note-{remark['id']}",
                "text": remark["text"] or "",
                "source": remark["source"] or "note",
                "createdAt": remark["ts"] or "",
            }
            for remark in remarks
            if remark["text"]
        ],
        "kit": kit,
    }


def _run_summary(rows, workspace_label):
    grouped = defaultdict(list)
    for row in rows:
        grouped[_json_safe_date(row["first_seen"])].append(row)
    runs = []
    for assigned, group in sorted(grouped.items(), reverse=True):
        if not assigned:
            continue
        locations = [str(row["address"] or "").strip() for row in group if row["address"]]
        location = locations[0].split(",")[-2].strip() if locations else "Local database"
        contacts = sum(1 for row in group if row["phone"])
        runs.append({
            "id": f"run-{assigned}",
            "vertical": workspace_label,
            "location": location or "Local database",
            "category": "Scraped leads",
            "createdAt": f"{assigned}T12:00:00",
            "found": len(group),
            "saved": len(group),
            "contacts": contacts,
            "status": "Complete",
        })
    return runs[:90]


def api_bootstrap(selected=None):
    records = discover_workspaces()
    active = resolve_workspace(selected)
    all_leads = []
    all_runs = []
    for record in records:
        conn = workspace_conn(record)
        try:
            rows = workspace_rows(conn)
            all_leads.extend(_serialize_lead(conn, row, record["label"], record)
                             for row in rows)
            all_runs.extend(_run_summary(rows, record["label"]))
        finally:
            conn.close()
    return {
        "version": VERSION,
        "workspace": active["label"],
        "workspaces": [{"key": r["key"], "label": r["label"]} for r in records],
        "leads": all_leads,
        "runs": all_runs,
        "lastSavedAt": datetime.now().isoformat(timespec="seconds"),
    }


def export_workspace_csv(record):
    conn = workspace_conn(record)
    try:
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(["Business", "Phone", "Location", "Email", "Website",
                         "Rating", "Reviews", "Score", "Tier", "Pitch Angle",
                         "Status", "Follow-up Date", "Remarks", "First Seen",
                         "Last Seen", "Times Seen"])
        for row in workspace_rows(conn):
            writer.writerow([row["name"], row["phone"], row["address"],
                             row["email"], row["website"], row["rating"],
                             row["reviews"], row["score"], row["tier"],
                             row["angle"], row["status"], row["followup"],
                             row["n_remarks"], row["first_seen"],
                             row["last_seen"], row["times_seen"]])
        return buf.getvalue()
    finally:
        conn.close()


def sync_client_state(conn, payload):
    leads = payload.get("leads") if isinstance(payload, dict) else None
    if not isinstance(leads, list):
        return 0
    now = datetime.now().isoformat(timespec="seconds")
    changed = 0
    for item in leads:
        if not isinstance(item, dict):
            continue
        phone = _phone_key(item.get("phone"))
        if len(phone) < 10:
            continue
        existing = conn.execute("SELECT phone FROM leads WHERE phone = ?", (phone,)).fetchone()
        values = (
            phone,
            str(item.get("name") or ""),
            str(item.get("location") or item.get("address") or ""),
            str(item.get("website") or ""),
            str(item.get("email") or ""),
            float(item.get("rating") or 0),
            int(item.get("reviews") or 0),
            int(item.get("score") or 0),
            str(item.get("tier") or "C"),
            str(item.get("angle") or ""),
            str(item.get("firstSeen") or datetime.now().date().isoformat()),
            str(item.get("lastSeen") or datetime.now().date().isoformat()),
        )
        if existing:
            conn.execute(
                "UPDATE leads SET name=?, address=?, website=?, email=?, rating=?, reviews=?, score=?, tier=?, angle=?, first_seen=?, last_seen=? WHERE phone=?",
                (*values[1:], phone),
            )
        else:
            conn.execute(
                "INSERT INTO leads (phone,name,address,website,email,rating,reviews,score,tier,angle,first_seen,last_seen,times_seen,suppressed) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,1,0)",
                values,
            )
        status = str(item.get("status") or "Not called")
        if status not in legacy.STATUSES:
            status = "Not called"
        followup = str(item.get("followup") or "")
        conn.execute(
            "INSERT INTO crm_state (phone,status,followup,updated_at) VALUES (?,?,?,?) ON CONFLICT(phone) DO UPDATE SET status=excluded.status, followup=excluded.followup, updated_at=excluded.updated_at",
            (phone, status, followup, now),
        )
        if status in legacy.SUPPRESS_STATUSES:
            conn.execute("UPDATE leads SET suppressed=1, suppress_reason=? WHERE phone=?", (status, phone))
        else:
            conn.execute("UPDATE leads SET suppressed=0, suppress_reason=NULL WHERE phone=?", (phone,))
        for item_note in item.get("notes") or []:
            text = str(item_note.get("text") or "").strip()
            source = str(item_note.get("source") or "app")
            if not text:
                continue
            exists_note = conn.execute(
                "SELECT 1 FROM crm_remarks WHERE phone=? AND text=? AND source=? LIMIT 1",
                (phone, text, source),
            ).fetchone()
            if not exists_note:
                conn.execute(
                    "INSERT INTO crm_remarks (phone,ts,text,source) VALUES (?,?,?,?)",
                    (phone, str(item_note.get("createdAt") or now), text, source),
                )
        changed += 1
    conn.commit()
    return changed


class VaaniHandler(legacy.Handler):
    def _serve_ui(self, relative_path):
        relative_path = relative_path.replace("\\", "/").lstrip("/")
        if relative_path in ("", "index.html"):
            target = UI_ROOT / "index.html"
            content = target.read_text(encoding="utf-8")
            self._send(200, content)
            return True
        if relative_path.startswith("src/"):
            target = (UI_ROOT / relative_path).resolve()
            if target != UI_ROOT and str(target).startswith(f"{UI_ROOT}{os.sep}") and target.is_file():
                self._send(200, target.read_bytes(), mimetypes.guess_type(str(target))[0] or "application/octet-stream")
                return True
        return False

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/", "/index.html") or parsed.path.startswith("/src/"):
            if self._serve_ui(parsed.path.lstrip("/")):
                return
        params = urllib.parse.parse_qs(parsed.query)
        if parsed.path == "/api/bootstrap":
            selected = params.get("workspace", [None])[0]
            self._send(200, json.dumps(api_bootstrap(selected)),
                       "application/json; charset=utf-8")
            return
        if parsed.path == "/api/workspaces":
            self._send(200, json.dumps({
                "workspaces": [{"key": r["key"], "label": r["label"]}
                                for r in discover_workspaces()]
            }), "application/json; charset=utf-8")
            return
        if parsed.path == "/api/export.csv":
            record = resolve_workspace(params.get("workspace", [None])[0])
            self._send(200, export_workspace_csv(record),
                       "text/csv; charset=utf-8", {
                           "Content-Disposition": 'attachment; filename="vaani-crm-export.csv"'
                       })
            return
        if parsed.path == "/api/kit":
            record = resolve_workspace(params.get("workspace", [None])[0])
            conn = workspace_conn(record)
            try:
                self._send(200, json.dumps(_kit_metadata(
                    record, conn, params.get("phone", [""])[0]
                )), "application/json; charset=utf-8")
            finally:
                conn.close()
            return
        if parsed.path == "/api/report-card":
            record = resolve_workspace(params.get("workspace", [None])[0])
            if not _configure_legacy_for_workspace(record):
                self._send(404, "Report cards are not available for this workspace.")
                return
            path = legacy.kit_card_file(params.get("phone", [""])[0])
            if not path:
                self._send(404, "Report card not generated yet.")
                return
            with open(path, "rb") as handle:
                self._send(200, handle.read(), "image/png")
            return
        if parsed.path in ("/api/brochure", "/brochure"):
            record = resolve_workspace(params.get("workspace", [None])[0])
            if record["key"] != "restaurants":
                self._send(404, "No brochure is configured for this workspace.")
                return
            brochure = HERE / "book-my-food-brochure.pdf"
            if brochure.exists():
                self._send(200, brochure.read_bytes(), "application/pdf", {
                    "Content-Disposition": 'attachment; filename="book-my-food-brochure.pdf"'
                })
            else:
                self._send(404, "Brochure not found.")
            return
        legacy.Handler.do_GET(self)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/api/kit-generate", "/api/kit-autosend", "/api/kit-sent"):
            try:
                length = int(self.headers.get("Content-Length", 0) or 0)
                raw = self.rfile.read(length).decode("utf-8")
                if self.headers.get("Content-Type", "").startswith("application/json"):
                    payload = json.loads(raw or "{}")
                else:
                    payload = {key: values[0] for key, values in
                               urllib.parse.parse_qs(raw).items()}
                record = resolve_workspace(payload.get("workspace"))
                action = "generate" if parsed.path.endswith("generate") else "autosend" if parsed.path.endswith("autosend") else "mark"
                ok, message = _kit_action(record, payload.get("phone", ""), action)
                refreshed = workspace_conn(record)
                try:
                    kit = _kit_metadata(record, refreshed, payload.get("phone", ""))
                finally:
                    refreshed.close()
                self._send(200, json.dumps({"ok": ok, "message": message,
                                           "kit": kit}),
                           "application/json; charset=utf-8")
            except Exception as exc:
                self._send(400, json.dumps({"ok": False, "message": str(exc)}),
                           "application/json; charset=utf-8")
            return
        if parsed.path == "/api/sync":
            try:
                length = int(self.headers.get("Content-Length", 0) or 0)
                if length > 2_000_000:
                    self._send(413, "Too large")
                    return
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                record = resolve_workspace(payload.get("workspace"))
                conn = workspace_conn(record)
                try:
                    changed = sync_client_state(conn, payload)
                finally:
                    conn.close()
                self._send(200, json.dumps({"ok": True, "changed": changed}), "application/json; charset=utf-8")
            except Exception as exc:
                self._send(400, json.dumps({"ok": False, "error": str(exc)}), "application/json; charset=utf-8")
            return
        legacy.Handler.do_POST(self)


def main():
    self_update()
    server = None
    for port in CRM_PORTS:
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), VaaniHandler)
            break
        except OSError:
            continue
    if server is None:
        print(f"  Could not start Vaani CRM - ports {CRM_PORTS[0]}-{CRM_PORTS[-1]} are busy.")
        return
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print()
    print(f"  VAANI CRM v{VERSION}  (all workspaces)")
    print("  ----------------------------")
    print("  Restaurants, Healthcare and scraper-created categories are available")
    print("  from the Workspace switcher in this one browser window.")
    print(f"  Open in your browser: {url}")
    print("  Your local SQLite database and trackers stay in this folder.")
    print("  To stop: close this window or press Ctrl+C.")
    print()
    legacy.open_browser(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
