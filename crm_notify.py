#!/usr/bin/env python3
"""
Vaani CRM - follow-up reminder popup.

Run by Windows at every login (installed by install-crm-reminder.bat).
Looks at vaani_master.db next to this file: if any lead has a
follow-up due today or overdue, it shows ONE popup summarising them.
If nothing is due, it stays completely silent.

Run it by hand with --test to see the popup even when nothing is due.
"""
import os
import sqlite3
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "vaani_master.db")
TERMINAL = {"Won", "Lost", "Not interested", "Wrong number",
            "Do not contact"}


def due_summary():
    """(overdue_names, today_names) for leads with a follow-up date
    whose status is not finished. Empty lists when nothing is due or
    the CRM has never been used."""
    if not os.path.exists(DB_PATH):
        return [], []
    try:
        db = sqlite3.connect(DB_PATH)
        db.row_factory = sqlite3.Row
        has = db.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table'"
            " AND name = 'crm_state'").fetchone()
        if not has:
            db.close()
            return [], []
        rows = db.execute(
            "SELECT s.followup, s.status, l.name FROM crm_state s"
            " LEFT JOIN leads l ON l.phone = s.phone"
            " WHERE s.followup != '' ORDER BY s.followup").fetchall()
        db.close()
    except Exception:
        return [], []
    today = date.today().isoformat()
    overdue, due = [], []
    for r in rows:
        if r["status"] in TERMINAL:
            continue
        name = r["name"] or "A lead"
        if r["followup"] < today:
            overdue.append(name)
        elif r["followup"] == today:
            due.append(name)
    return overdue, due


def main():
    test = "--test" in sys.argv
    overdue, due = due_summary()
    if not overdue and not due and not test:
        return  # nothing due - stay silent
    lines = ["Vaani CRM - follow-ups need you today.", ""]
    if overdue:
        lines.append(f"OVERDUE ({len(overdue)}):")
        lines += [f"  - {n}" for n in overdue[:6]]
        if len(overdue) > 6:
            lines.append(f"  ... and {len(overdue) - 6} more")
        lines.append("")
    if due:
        lines.append(f"DUE TODAY ({len(due)}):")
        lines += [f"  - {n}" for n in due[:6]]
        if len(due) > 6:
            lines.append(f"  ... and {len(due) - 6} more")
        lines.append("")
    if not overdue and not due:
        lines.append("No follow-ups due today. (This was a test.)")
    lines.append("Open Vaani CRM (run-crm.bat) to work through them.")
    message = "\n".join(lines)
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0, message, "Vaani CRM - follow-ups due", 0x40)
            return
        except Exception:
            pass
    print(message)


if __name__ == "__main__":
    main()
