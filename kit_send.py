#!/usr/bin/env python3
"""Vaani Kit Sender v1.0.1 - sends ONE WhatsApp kit (intro text +
report card image + brochure PDF, together in one chat) through
WhatsApp Web, for the Vaani CRM's Auto-send button.

    py kit_send.py <phone10> <card.png> <brochure.pdf> <caption.txt>

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

VERSION = "1.0.1"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/kit_send.py")
HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "wa-profile")
COMPOSE = 'footer div[contenteditable="true"]'
PREVIEW_MARK = ('div[data-testid="media-caption-input"], '
                '[data-testid="media-preview"], '
                'div[data-animate-modal-popup="true"]')


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


def _preview_open(pg, timeout=9000):
    try:
        pg.locator(PREVIEW_MARK).first.wait_for(state="visible",
                                                timeout=timeout)
        return True
    except Exception:
        return False


def _paste_files(pg, card_png, brochure_pdf):
    """Paste both files into the chat - WhatsApp Web opens its
    media preview for pasted files. True if the preview opened."""
    def b64(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    js = """([cardB64, pdfB64, cardName, pdfName]) => {
      function fileFrom(b64, name, type) {
        const bin = atob(b64);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        return new File([arr], name, {type: type});
      }
      const dt = new DataTransfer();
      dt.items.add(fileFrom(cardB64, cardName, 'image/png'));
      dt.items.add(fileFrom(pdfB64, pdfName, 'application/pdf'));
      const target = document.querySelector(
          'footer [contenteditable="true"]') || document.body;
      target.focus();
      const ev = new ClipboardEvent('paste',
          {bubbles: true, cancelable: true, clipboardData: dt});
      return target.dispatchEvent(ev);
    }"""
    try:
        pg.evaluate(js, [b64(card_png), b64(brochure_pdf),
                         os.path.basename(card_png),
                         os.path.basename(brochure_pdf)])
    except Exception:
        return False
    return _preview_open(pg)


def _button_files(pg, card_png, brochure_pdf):
    """Fallback: find the attach button in any of its known
    shapes (or by position, left of the typing box), open it and
    set both files on the file input it reveals."""
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
            inputs.nth(i).set_input_files([card_png, brochure_pdf],
                                          timeout=8000)
            return _preview_open(pg)
        except Exception:
            continue
    return _preview_open(pg, 2000)


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
        step("Chat open - attaching the report card and brochure")
        if not _paste_files(pg, card_png, brochure_pdf):
            step("Paste did not open the preview - trying the "
                 "attach button")
            if not _button_files(pg, card_png, brochure_pdf):
                rc = fail(pg, "The files could not be attached "
                              "in WhatsApp Web - neither pasting "
                              "nor the attach button opened the "
                              "preview.")
                ctx.close()
                return rc
        # ---- caption = the intro text, in the media preview ----
        cap_box = pg.locator(
            'div[data-testid="media-caption-input"] '
            'div[contenteditable="true"], '
            f'{PREVIEW_MARK} div[contenteditable="true"]').first
        caption_in_preview = False
        try:
            cap_box.wait_for(state="visible", timeout=10000)
            cap_box.click()
            pg.keyboard.insert_text(caption)
            caption_in_preview = True
        except Exception:
            step("No caption box in the preview - the intro will "
                 "be sent as a message right after the files")
        send_btn = pg.locator(
            'span[data-icon="send"], div[aria-label="Send"], '
            'button[aria-label="Send"], '
            '[data-testid="media-send"]').first
        sent = False
        try:
            send_btn.wait_for(state="visible", timeout=12000)
            send_btn.click()
            sent = True
        except Exception:
            if caption_in_preview:
                try:
                    cap_box.press("Enter")
                    sent = True
                except Exception:
                    pass
        if not sent:
            rc = fail(pg, "Files attached but the Send button "
                          "was not found - nothing was sent.")
            ctx.close()
            return rc
        pg.wait_for_timeout(2500)
        if not caption_in_preview:
            try:
                box = pg.locator(COMPOSE).first
                box.wait_for(state="visible", timeout=15000)
                box.click()
                pg.keyboard.insert_text(caption)
                pg.keyboard.press("Enter")
                pg.wait_for_timeout(1800)
            except Exception:
                rc = fail(pg, "Files were sent but the intro "
                              "text could not be posted.")
                ctx.close()
                return rc
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
        return fail(None, "Bad phone number.")
    if not os.path.exists(card_png):
        return fail(None, f"Report card file not found: {card_png}")
    if not os.path.exists(brochure_pdf):
        return fail(None, f"Brochure file not found: {brochure_pdf}")
    try:
        with open(caption_file, encoding="utf-8") as f:
            caption = f.read().strip()
    except Exception:
        return fail(None, "Caption file could not be read.")
    if not caption:
        return fail(None, "Caption is empty.")
    return send_kit(phone, card_png, brochure_pdf, caption)


if __name__ == "__main__":
    sys.exit(main())
