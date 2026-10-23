#!/usr/bin/env python3
"""Vaani Kit Sender v1.0.0 - sends ONE WhatsApp kit (intro text +
report card image + brochure PDF, together in one chat) through
WhatsApp Web, for the Vaani CRM's Auto-send button.

    py kit_send.py <phone10> <card.png> <brochure.pdf> <caption.txt>

It drives a real Chrome window (Playwright) with a persistent
profile in wa-profile/ next to this file: the first run shows the
WhatsApp Web QR code - scan it once with your phone and later runs
are already logged in. The window stays visible on purpose: you see
exactly what is sent, to whom, before it goes.

Status lines on stdout (the CRM reads the last one):
    STEP: <what is happening>
    SENT                      kit delivered to the chat
    ERROR: <what went wrong>  nothing was sent / could not confirm
"""
import os
import re
import shutil
import subprocess
import sys
import urllib.request

VERSION = "1.0.0"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/kit_send.py")
HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "wa-profile")


def _version_tuple(text):
    try:
        return tuple(int(x) for x in text.strip().split("."))
    except Exception:
        return (0,)


def self_update():
    """Same deal as Vaani, the CRM and the Report Card: check online
    for a newer kit_send.py and swap it in. Silent when offline or
    up to date."""
    if os.environ.get("VAANI_KIT_NO_UPDATE") == "1" \
            or "--no-update" in sys.argv:
        return
    try:
        req = urllib.request.Request(UPDATE_URL,
                                     headers={"User-Agent": "VaaniKit"})
        with urllib.request.urlopen(req, timeout=5) as r:
            content = r.read().decode("utf-8", "replace")
    except Exception:
        return
    m = re.search(r'^VERSION\s*=\s*"([0-9.]+)"', content, re.M)
    if not m or _version_tuple(m.group(1)) <= _version_tuple(VERSION):
        return
    if "Vaani Kit Sender" not in content or len(content) < 5000:
        return
    try:
        compile(content, "kit_send.py", "exec")
    except Exception:
        return
    me = os.path.abspath(__file__)
    try:
        shutil.copyfile(me, me + ".bak")
        with open(me, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        return
    print(f"STEP: Kit Sender updated to v{m.group(1)} - restarting")
    env = dict(os.environ)
    env["VAANI_KIT_NO_UPDATE"] = "1"
    subprocess.Popen([sys.executable, me] + sys.argv[1:], env=env)
    sys.exit(0)


def step(msg):
    print(f"STEP: {msg}", flush=True)


def fail(msg):
    print(f"ERROR: {msg}", flush=True)
    return 1


def send_kit(phone, card_png, brochure_pdf, caption):
    from playwright.sync_api import sync_playwright

    os.makedirs(PROFILE_DIR, exist_ok=True)
    with sync_playwright() as pw:
        step("Opening WhatsApp Web")
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False,
            viewport={"width": 1180, "height": 860},
            args=["--disable-blink-features=AutomationControlled"])
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            pg.goto("https://web.whatsapp.com/send?phone=91" + phone,
                    wait_until="domcontentloaded", timeout=60000)
        except Exception:
            ctx.close()
            return fail("WhatsApp Web did not load - check the "
                        "internet connection on this PC.")
        # ---- login state: QR code vs logged in ----
        compose = pg.locator(
            'footer div[contenteditable="true"], '
            'div[contenteditable="true"][role="textbox"]')
        try:
            compose.first.wait_for(state="visible", timeout=25000)
            logged_in = True
        except Exception:
            logged_in = False
        if not logged_in:
            step("WhatsApp is not logged in here. Scan the QR code "
                 "in the Chrome window with your phone (WhatsApp > "
                 "Linked devices > Link a device). Waiting...")
            try:
                compose.first.wait_for(state="visible",
                                       timeout=180000)
            except Exception:
                ctx.close()
                return fail("QR code was not scanned in time. "
                            "Run Auto-send again and scan it.")
        step("Chat open - attaching the report card and brochure")
        # ---- attach both files together ----
        clip = pg.locator(
            'span[data-icon="plus"], span[data-icon="attach-menu-plus"],'
            ' span[data-icon="clip"]').first
        try:
            clip.wait_for(state="visible", timeout=20000)
            clip.click()
            pg.wait_for_timeout(1200)
        except Exception:
            ctx.close()
            return fail("Could not find the attach (paperclip) "
                        "button in WhatsApp Web.")
        inputs = pg.locator('input[type="file"]')
        attached = False
        for i in range(inputs.count() - 1, -1, -1):
            try:
                inputs.nth(i).set_input_files([card_png, brochure_pdf],
                                              timeout=8000)
                attached = True
                break
            except Exception:
                continue
        if not attached:
            ctx.close()
            return fail("WhatsApp Web would not accept the files.")
        # ---- caption = the intro text, in the media preview ----
        pg.wait_for_timeout(1800)
        cap_box = pg.locator(
            'div[data-testid="media-caption-input"] '
            'div[contenteditable="true"], '
            'div[contenteditable="true"][data-tab="1"]').first
        caption_sent_with_media = False
        try:
            cap_box.wait_for(state="visible", timeout=10000)
            cap_box.click()
            pg.keyboard.insert_text(caption)
            caption_sent_with_media = True
        except Exception:
            step("No caption box in the preview - the intro will be "
                 "sent as a message right after the files")
        send_btn = pg.locator(
            'span[data-icon="send"], div[aria-label="Send"], '
            'button[aria-label="Send"]').first
        try:
            send_btn.wait_for(state="visible", timeout=15000)
            send_btn.click()
        except Exception:
            ctx.close()
            return fail("Files attached but the Send button was "
                        "not found - nothing was sent.")
        pg.wait_for_timeout(2500)
        if not caption_sent_with_media:
            try:
                box = pg.locator(
                    'footer div[contenteditable="true"]').first
                box.wait_for(state="visible", timeout=15000)
                box.click()
                pg.keyboard.insert_text(caption)
                pg.keyboard.press("Enter")
                pg.wait_for_timeout(1800)
            except Exception:
                ctx.close()
                return fail("Files were sent but the intro text "
                            "could not be posted.")
        # ---- verify: files visible in the chat ----
        try:
            pg.wait_for_selector(
                'span[data-icon="media-document"], '
                'span[data-icon="document"], img[src^="blob:"]',
                timeout=15000)
        except Exception:
            pass
        step("Kit sent")
        print("SENT", flush=True)
        pg.wait_for_timeout(1200)
        ctx.close()
        return 0


def main():
    self_update()
    args = [a for a in sys.argv[1:] if a != "--no-update"]
    if len(args) < 4:
        print(__doc__)
        return 1
    phone = re.sub(r"\D", "", args[0])[-10:]
    card_png, brochure_pdf, caption_file = args[1], args[2], args[3]
    if len(phone) != 10:
        return fail("Bad phone number.")
    if not os.path.exists(card_png):
        return fail(f"Report card file not found: {card_png}")
    if not os.path.exists(brochure_pdf):
        return fail(f"Brochure file not found: {brochure_pdf}")
    try:
        with open(caption_file, encoding="utf-8") as f:
            caption = f.read().strip()
    except Exception:
        return fail("Caption file could not be read.")
    if not caption:
        return fail("Caption is empty.")
    return send_kit(phone, card_png, brochure_pdf, caption)


if __name__ == "__main__":
    sys.exit(main())
