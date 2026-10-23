#!/usr/bin/env python3
"""Vaani Kit Sender v1.1.0 - sends ONE WhatsApp kit through
WhatsApp Web, for the Vaani CRM's Auto-send button, as three
messages in this order:

    1. the intro / greeting text (the file passed as caption.txt)
    2. the report card image
    3. the brochure PDF

    (Pass - instead of a brochure file to send only 1 and 2 - the
    healthcare kit works that way: intro + Content & Trust Report.)

    py kit_send.py <phone10> <card.png> <brochure.pdf> <intro.txt>

It drives a real Chrome window (Playwright) with a persistent
profile in wa-profile/ next to this file: the first run shows the
WhatsApp Web QR code - scan it once with your phone and later runs
are already logged in. The window stays visible on purpose: you see
exactly what is sent, to whom, before it goes.

Files get in by PASTING them into the chat (WhatsApp Web accepts
pasted files and opens its normal preview) - that does not depend
on the attach button, whose markup WhatsApp keeps changing. The
attach button is still tried as a fallback, in several shapes.

Status lines on stdout (the CRM reads the last one):
    STEP: <what is happening>
    SENT                      kit delivered to the chat
    ERROR: <what went wrong>  nothing was sent / could not confirm
On any failure a screenshot (wa-debug.png) and the page HTML
(wa-debug.html) are saved next to this file for diagnosis.
"""
import base64
import os
import re
import shutil
import subprocess
import sys
import urllib.request

VERSION = "1.1.0"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/kit_send.py")
HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "wa-profile")
COMPOSE = 'footer div[contenteditable="true"]'
PREVIEW_MARK = ('div[data-testid="media-caption-input"], '
                '[data-testid="media-preview"], '
                'div[data-animate-modal-popup="true"]')
SEND_BTN = ('span[data-icon="send"], div[aria-label="Send"], '
            'button[aria-label="Send"], [data-testid="media-send"], '
            '[data-testid="send"]')


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


def debug_capture(pg):
    """Save what the WhatsApp window looked like, for diagnosis."""
    try:
        pg.screenshot(path=os.path.join(HERE, "wa-debug.png"))
    except Exception:
        pass
    try:
        with open(os.path.join(HERE, "wa-debug.html"), "w",
                  encoding="utf-8") as f:
            f.write(pg.content())
    except Exception:
        pass


def fail(pg, msg):
    if pg is not None:
        debug_capture(pg)
        msg += (" (A picture of the window was saved as "
                "wa-debug.png in the Vaani folder.)")
    print(f"ERROR: {msg}", flush=True)
    return 1


# ------------------------------------------------------------------
# the three sends
# ------------------------------------------------------------------

def _preview_open(pg, timeout=9000):
    try:
        pg.locator(PREVIEW_MARK).first.wait_for(state="visible",
                                                timeout=timeout)
        return True
    except Exception:
        return False


def _preview_gone(pg, timeout=9000):
    try:
        pg.locator(PREVIEW_MARK).first.wait_for(state="hidden",
                                                timeout=timeout)
        return True
    except Exception:
        return False


def send_text(pg, text):
    """Message 1: the intro as a normal chat message. Verified the
    strict way: if our text is still sitting in the typing box
    afterwards, it did not go out."""
    box = pg.locator(COMPOSE).first
    box.click()
    pg.keyboard.insert_text(text)
    pg.wait_for_timeout(700)
    pg.keyboard.press("Enter")
    pg.wait_for_timeout(1200)
    try:
        left = box.inner_text(timeout=3000)
    except Exception:
        left = ""
    if text[:40] in left:
        try:
            pg.locator(SEND_BTN).last.click(timeout=4000)
            pg.wait_for_timeout(1200)
            left = box.inner_text(timeout=3000)
        except Exception:
            pass
    return text[:40] not in (left or "")


def _paste_one(pg, path, mime):
    """Paste one file into the chat; WhatsApp Web opens its media
    preview. True if the preview opened."""
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    js = """([b64, name, mime]) => {
      const bin = atob(b64);
      const arr = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      const dt = new DataTransfer();
      dt.items.add(new File([arr], name, {type: mime}));
      const target = document.querySelector(
          'footer [contenteditable="true"]') || document.body;
      target.focus();
      const ev = new ClipboardEvent('paste',
          {bubbles: true, cancelable: true, clipboardData: dt});
      return target.dispatchEvent(ev);
    }"""
    try:
        pg.evaluate(js, [b64, os.path.basename(path), mime])
    except Exception:
        return False
    return _preview_open(pg)


def _button_one(pg, path):
    """Fallback for one file: find the attach button in any of its
    known shapes (or by position, left of the typing box), open it
    and set the file on the input it reveals."""
    candidates = [
        'footer span[data-icon="plus"]',
        'footer span[data-icon="plus-rounded"]',
        'footer span[data-icon="attach-menu-plus"]',
        'footer span[data-icon="clip"]',
        'span[data-icon="plus"]',
        'span[data-icon="attach-menu-plus"]',
        'span[data-icon="clip"]',
        'button[aria-label="Attach"]',
        '[title="Attach"]',
        'footer [role="button"][aria-label="Attach"]',
    ]
    clicked = False
    for sel in candidates:
        loc = pg.locator(sel).first
        try:
            loc.wait_for(state="visible", timeout=1500)
            loc.click()
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        btns = pg.locator('footer button, footer [role="button"]')
        try:
            n = min(btns.count(), 4)
        except Exception:
            n = 0
        for i in range(n):
            try:
                btns.nth(i).click()
                pg.wait_for_timeout(600)
                if pg.locator('input[type="file"]').count() > 0 \
                        or _preview_open(pg, 1200):
                    clicked = True
                    break
            except Exception:
                continue
    if not clicked:
        return False
    pg.wait_for_timeout(900)
    inputs = pg.locator('input[type="file"]')
    for i in range(inputs.count() - 1, -1, -1):
        try:
            inputs.nth(i).set_input_files([path], timeout=8000)
            return _preview_open(pg)
        except Exception:
            continue
    return _preview_open(pg, 2000)


def send_file(pg, path, mime):
    """Messages 2 and 3: one file each, sent from the preview."""
    if not _paste_one(pg, path, mime):
        step("Paste did not open the preview - trying the "
             "attach button")
        if not _button_one(pg, path):
            return False
    pg.wait_for_timeout(1200)
    btn = pg.locator(SEND_BTN).first
    try:
        btn.wait_for(state="visible", timeout=10000)
        btn.click()
    except Exception:
        pg.keyboard.press("Enter")
    if not _preview_gone(pg):
        try:
            pg.locator(SEND_BTN).first.click(timeout=4000)
        except Exception:
            pass
    if not _preview_gone(pg, 4000):
        return False
    pg.wait_for_timeout(1500)
    return True


# ------------------------------------------------------------------

def send_kit(phone, card_png, brochure_pdf, intro_text):
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
            return fail(None, "WhatsApp Web did not load - check "
                              "the internet connection on this PC.")
        compose = pg.locator(COMPOSE)
        try:
            compose.first.wait_for(state="visible", timeout=30000)
        except Exception:
            step("WhatsApp is not logged in here, or the chat is "
                 "still opening. If a QR code is showing, scan it "
                 "with your phone (WhatsApp > Linked devices > "
                 "Link a device). Waiting...")
            try:
                compose.first.wait_for(state="visible",
                                       timeout=180000)
            except Exception:
                rc = fail(pg, "The chat never opened - if a QR "
                              "code was showing, it was not "
                              "scanned in time. Run Auto-send "
                              "again and scan it.")
                ctx.close()
                return rc
        pg.wait_for_timeout(1200)
        step("Sending the introduction message")
        if not send_text(pg, intro_text):
            rc = fail(pg, "The introduction message could not be "
                          "sent, so the files were not sent "
                          "either - the kit only goes out in "
                          "order: text, report card, brochure.")
            ctx.close()
            return rc
        step("Sending the report card")
        if not send_file(pg, card_png, "image/png"):
            rc = fail(pg, "The introduction went, but the report "
                          "card could not be attached - the "
                          "brochure was not sent either.")
            ctx.close()
            return rc
        if brochure_pdf:
            step("Sending the brochure")
            if not send_file(pg, brochure_pdf, "application/pdf"):
                rc = fail(pg, "The introduction and report card "
                              "went, but the brochure could not "
                              "be attached.")
                ctx.close()
                return rc
            step("Kit sent - text, report card and brochure, "
                 "in order")
        else:
            step("Kit sent - text and report card, in order")
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
    card_png, brochure_pdf, intro_file = args[1], args[2], args[3]
    if len(phone) != 10:
        return fail(None, "Bad phone number.")
    if not os.path.exists(card_png):
        return fail(None, f"Report card file not found: {card_png}")
    if brochure_pdf in ("-", ""):
        brochure_pdf = None
    elif not os.path.exists(brochure_pdf):
        return fail(None, f"Brochure file not found: {brochure_pdf}")
    try:
        with open(intro_file, encoding="utf-8") as f:
            intro_text = f.read().strip()
    except Exception:
        return fail(None, "Introduction file could not be read.")
    if not intro_text:
        return fail(None, "Introduction text is empty.")
    return send_kit(phone, card_png, brochure_pdf, intro_text)


if __name__ == "__main__":
    sys.exit(main())
