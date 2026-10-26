#!/usr/bin/env python3
"""Vaani CRM launcher and database-backed web UI.

The original CRM implementation is kept in ``vaani_crm_legacy.py`` so the
existing SQLite schema, Excel import, WhatsApp kit and healthcare section stay
available. This launcher adds the new UI, a small same-origin JSON API, and a
safe allow-listed self-updater that never touches local data files.
"""

import hashlib
import json
import mimetypes
import os
import re
import shutil
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
VERSION = "2.0.0"
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


def _serialize_lead(conn, row):
    remarks = legacy.fetch_remarks(conn, row["phone"])
    return {
        "id": _lead_id(row["phone"]),
        "vertical": legacy.SECTION_LABEL,
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
    }


def _run_summary(rows):
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
            "vertical": legacy.SECTION_LABEL,
            "location": location or "Local database",
            "category": "Scraped leads",
            "createdAt": f"{assigned}T12:00:00",
            "found": len(group),
            "saved": len(group),
            "contacts": contacts,
            "status": "Complete",
        })
    return runs[:90]


def api_bootstrap(conn):
    rows = legacy.fetch_leads(conn, sort="score")
    return {
        "version": VERSION,
        "workspace": legacy.SECTION_LABEL,
        "leads": [_serialize_lead(conn, row) for row in rows],
        "runs": _run_summary(rows),
        "lastSavedAt": datetime.now().isoformat(timespec="seconds"),
    }


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
        conn = legacy.get_conn()
        try:
            if parsed.path == "/api/bootstrap":
                self._send(200, json.dumps(api_bootstrap(conn)), "application/json; charset=utf-8")
                return
            if parsed.path == "/api/export.csv":
                self._send(200, legacy.export_csv_text(conn), "text/csv; charset=utf-8", {"Content-Disposition": 'attachment; filename="vaani-crm-export.csv"'})
                return
        finally:
            conn.close()
        legacy.Handler.do_GET(self)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/sync":
            try:
                length = int(self.headers.get("Content-Length", 0) or 0)
                if length > 2_000_000:
                    self._send(413, "Too large")
                    return
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                conn = legacy.get_conn()
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
    conn = legacy.get_conn()
    try:
        legacy.backup_db()
        imported = legacy.import_tracker(conn)
    finally:
        conn.close()
    server = None
    for port in legacy.PORTS:
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), VaaniHandler)
            break
        except OSError:
            continue
    if server is None:
        print(f"  Could not start Vaani CRM - ports {legacy.PORTS[0]}-{legacy.PORTS[-1]} are busy.")
        return
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print()
    print(f"  VAANI CRM v{VERSION}  ({legacy.SECTION_LABEL} section)")
    print("  ----------------------")
    if imported:
        print(f"  Imported {imported} lead(s) from your Excel tracker.")
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
