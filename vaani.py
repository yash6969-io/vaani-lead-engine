#!/usr/bin/env python3
"""
Vaani Lead Engine v3.4
======================
Type a location (even a small area like "Gota, Ahmedabad") and get 4 lead
databases about restaurants there:

    1. restaurant_names.csv   - every restaurant name found
    2. locations.csv          - their addresses
    3. phone_numbers.csv      - MOBILE numbers only (landlines are dropped)
    4. email_ids.csv          - email addresses found on their websites

Plus a bonus: all_leads.xlsx - everything combined in one Excel file.

The scraper is BUILT IN - this one file does everything. No Docker, no
separate engine, no API keys. It opens Google Maps in a real browser
window (don't close it while it runs), reads each restaurant, then checks
the restaurant's own website for an email address.

Needs: Python + the two helpers in requirements.txt
    py -m pip install -r requirements.txt
"""
import argparse
import csv
import os
import random
import re
import subprocess
import sys
import urllib.parse
import urllib.request
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))

VERSION = "3.4.0"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/vaani.py")


# ---------------------------------------------------------------- auto-update

def _version_tuple(text):
    try:
        return tuple(int(x) for x in text.strip().split("."))
    except Exception:
        return (0,)


def self_update():
    """Check online for a newer vaani.py and, if there is one, swap it in
    and restart. Silent when offline or already up to date - the run just
    continues with the version on disk. Your leads are never uploaded;
    this only DOWNLOADS the program file."""
    if os.environ.get("VAANI_NO_UPDATE") == "1" or "--no-update" in sys.argv:
        return
    try:
        req = urllib.request.Request(UPDATE_URL,
                                     headers={"User-Agent": BROWSER_UA})
        with urllib.request.urlopen(req, timeout=5) as resp:
            content = resp.read().decode("utf-8", "replace")
    except Exception:
        return  # no internet / blocked - run as-is
    m = re.search(r'^VERSION\s*=\s*"([0-9.]+)"', content, re.M)
    if not m or _version_tuple(m.group(1)) <= _version_tuple(VERSION):
        return
    if "Vaani Lead Engine" not in content or len(content) < 10000:
        return  # not a real vaani.py - don't touch anything
    try:
        compile(content, "vaani.py", "exec")  # must be valid Python
    except Exception:
        return
    try:
        target = os.path.abspath(__file__)
        new_file = target + ".new"
        bak_file = target + ".bak"
        with open(new_file, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)
        with open(target, "r", encoding="utf-8") as f:
            old = f.read()
        with open(bak_file, "w", encoding="utf-8", newline="\n") as f:
            f.write(old)
        os.replace(new_file, target)
    except Exception:
        return
    print(f"\n  Updated Vaani to v{m.group(1)} - restarting with the "
          "new version...", flush=True)
    env = dict(os.environ, VAANI_NO_UPDATE="1")
    try:
        rc = subprocess.run([sys.executable, os.path.abspath(__file__)]
                            + sys.argv[1:], env=env).returncode
    except Exception:
        rc = 0
    sys.exit(rc)

BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/130.0.0.0 Safari/537.36")

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
JUNK_DOMAINS = ("example.", "sentry.", "wix.", "godaddy", "squarespace",
                "wordpress", "schema.org", "w3.org", "yourdomain",
                "email.com", "test.", "sample.")


# ---------------------------------------------------------------- utilities

# Indian STD (landline area) codes, without the trunk 0 - metros first,
# then the major cities. Used to tell a landline (0 + STD + number) apart
# from a mobile written with a trunk 0 (0 + 10-digit mobile).
STD_CODES = {
    "11", "22", "33", "44", "20", "79", "80", "40",
    "120", "121", "129", "135", "141", "161", "171", "172", "177", "181",
    "183", "191", "194", "231", "233", "240", "251", "253", "256", "257",
    "261", "265", "271", "281", "291", "322", "341", "343", "353", "361",
    "413", "416", "422", "427", "431", "452", "462", "471", "484", "487",
    "495", "512", "522", "532", "542", "551", "562", "571", "581", "591",
    "612", "621", "631", "641", "651", "654", "657", "671", "674", "712",
    "721", "724", "731", "734", "751", "755", "761", "771", "775", "788",
    "821", "824", "831", "832", "836", "861", "863", "866", "870", "877",
    "891",
}


def _is_std_landline(digits11):
    """True for 0 + known STD code + a landline-style subscriber number
    (subscriber numbers start 2-5; mobile numbers start 6-9)."""
    body = digits11[1:]
    for ln in (4, 3, 2):
        if body[:ln] in STD_CODES:
            sub = body[ln:]
            if sub and sub[0] in "2345":
                return True
    return False


def clean_phone(raw):
    """Return the 10-digit Indian mobile number, or None.

    Accepted: 9825012345, +91 98250 12345, and 098250 12345 (trunk 0 -
    the format most businesses use on Maps). Landlines (079-..., 011-...
    i.e. 0 + STD code + subscriber number) are dropped on purpose -
    they are useless for outreach.
    """
    if not raw:
        return None
    digits = re.sub(r"[^\d]", "", str(raw))
    if digits.startswith("91") and len(digits) == 12:   # +91 98250 12345
        digits = digits[2:]
    elif len(digits) == 11 and digits.startswith("0"):
        # 0 + 10 digits: either a trunk-0 mobile or an STD landline
        if digits[1] in "6789" and not _is_std_landline(digits):
            digits = digits[1:]
    if len(digits) == 10 and digits[0] in "6789":       # real mobile
        return digits
    return None                                          # landline / invalid


def split_phones(raw):
    """One Google Maps phone field can hold several numbers - split & clean."""
    if not raw:
        return []
    out = []
    for part in re.split(r"[,;/|]", str(raw)):
        mob = clean_phone(part)
        if mob and mob not in out:
            out.append(mob)
    return out


def split_emails(raw):
    """Clean up found emails: valid format only, junk dropped, deduped."""
    if not raw:
        return []
    out = []
    for e in EMAIL_RE.findall(str(raw).lower()):
        domain = e.rpartition("@")[2]
        if any(j in domain for j in JUNK_DOMAINS):
            continue
        if e not in out:
            out.append(e)
    return out[:3]


def slug(text):
    s = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return s or "leads"


# ---------------------------------------------------------------- scraper

class Blocked(Exception):
    """Google is rate-limiting this connection - stop politely."""


def import_playwright():
    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except ImportError:
        sys.exit(
            "\n  The browser tool (Playwright) is not installed yet.\n"
            "  Run this ONE command in this folder, then run vaani.py again:\n"
            "\n      py -m pip install -r requirements.txt\n")


def launch_browser(p, headless):
    """Use the user's own Chrome if present (no download); else the
    built-in browser; else download it once and retry."""
    args = ["--disable-blink-features=AutomationControlled"]
    kw_base = {"headless": headless, "args": args}
    proxy = (os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
             or os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy"))
    if proxy:
        kw_base["proxy"] = {"server": proxy}
    for channel, label in (("chrome", "your Google Chrome"),
                           (None, "the built-in browser")):
        try:
            kw = dict(kw_base)
            if channel:
                kw["channel"] = channel
            browser = p.chromium.launch(**kw)
            print(f"  browser ready ({label})")
            return browser
        except Exception:
            continue
    print("  downloading the browser (one time only, a few minutes)...")
    subprocess.run([sys.executable, "-m", "playwright", "install",
                    "chromium"])
    try:
        return p.chromium.launch(**kw_base)
    except Exception:
        sys.exit(
            "\n  Could not start a browser. Run this once, then try again:\n"
            "\n      py -m playwright install chromium\n")


def goto_retry(page, url, tries=3):
    """Open a page, retrying on flaky network errors instead of dying."""
    last = None
    for attempt in range(tries):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            return
        except Exception as e:
            last = e
            if attempt < tries - 1:
                page.wait_for_timeout(3000)
    raise last


def handle_consent(page):
    """Dismiss the Google cookie popup if it appears."""
    for label in ("Reject all", "Accept all"):
        try:
            btn = page.get_by_role("button", name=label)
            if btn.count():
                btn.first.click(timeout=3000)
                page.wait_for_timeout(1500)
                return
        except Exception:
            pass


def is_blocked(page):
    try:
        if "/sorry/" in page.url:
            return True
        body = page.locator("body").inner_text(timeout=3000).lower()
        return "unusual traffic" in body or "automated queries" in body
    except Exception:
        return False


def collect_links(page, query, max_results):
    """Open the Maps search, scroll the results list, gather place links."""
    url = "https://www.google.com/maps/search/" + urllib.parse.quote(query)
    goto_retry(page, url)
    page.wait_for_timeout(2500)
    handle_consent(page)
    if is_blocked(page):
        raise Blocked
    try:
        page.wait_for_selector('div[role="feed"], h1', timeout=20000)
    except Exception:
        pass
    feed = page.locator('div[role="feed"]')
    if feed.count() == 0:
        # the search went straight to a single place page
        return [page.url] if page.locator("h1").count() else []

    links, seen, stale = [], set(), 0
    while len(links) < max_results and stale < 4:
        found = page.eval_on_selector_all(
            'div[role="feed"] a[href*="/maps/place/"]',
            "els => els.map(e => e.href)")
        new = 0
        for href in found:
            key = href.split("?")[0]
            if key not in seen:
                seen.add(key)
                links.append(href)
                new += 1
        stale = 0 if new else stale + 1
        if page.locator("text=You've reached the end of the list").count():
            break
        try:
            feed.evaluate("el => el.scrollTo(0, el.scrollHeight)")
        except Exception:
            break
        page.wait_for_timeout(random.randint(1400, 2400))
    return links[:max_results]


def text_of(locator, timeout=5000):
    try:
        el = locator.first
        el.wait_for(state="attached", timeout=timeout)
        return el.inner_text(timeout=timeout).strip()
    except Exception:
        return ""


def clean_field(text):
    """Detail fields can start with an icon character - strip it."""
    return re.sub(r"^[^\w]+", "", text or "").strip()


PHONE_IN_TEXT = re.compile(
    r"\+?91[\s-]?\d{5}[\s-]?\d{5}|0\d{2,4}[\s-]?\d{6,8}|\d{5}[\s-]?\d{5}")


def best_mobile_in(text):
    """The mobile that appears most often in a blob of text/HTML - on a
    place page that is the place's own number (it repeats in the data,
    the call link and the button text). Only valid mobiles count."""
    counts, best, best_n = {}, "", 0
    for cand in PHONE_IN_TEXT.findall(text or ""):
        mob = clean_phone(cand)
        if mob:
            counts[mob] = counts.get(mob, 0) + 1
            if counts[mob] > best_n:
                best, best_n = mob, counts[mob]
    return best


def _phone_candidates(page):
    """Yield phone-number candidates from the page, most reliable first.
    Verified on a live page: the row is a button named "Phone: ..." and
    contains a "Call phone number" link whose href is tel:..."""
    # 0) give the phone row a chance to appear at all (it can load late)
    try:
        page.wait_for_selector(
            'a[href^="tel:"], [data-item-id^="phone:tel:"], '
            'button[aria-label^="Phone:"]', timeout=15000)
    except Exception:
        pass
    # 1) the "Call phone number" link carries the number in its href
    try:
        a = page.locator('a[href^="tel:"]').first
        href = a.get_attribute("href", timeout=3000) or ""
        if href.startswith("tel:") and href[4:].strip():
            yield href[4:].strip()
    except Exception:
        pass
    # 2) the phone button, found by its accessible name "Phone: ..."
    try:
        b = page.get_by_role("button", name=re.compile(r"^Phone:")).first
        m = PHONE_IN_TEXT.search(b.inner_text(timeout=3000))
        if m:
            yield m.group(0)
    except Exception:
        pass
    # 3) the phone element's data attribute carries the number
    try:
        el = page.locator('[data-item-id^="phone:tel:"]').first
        item = el.get_attribute("data-item-id", timeout=3000) or ""
        if item.startswith("phone:tel:") and item[10:].strip():
            yield item[10:].strip()
        text = el.inner_text(timeout=2000).strip()
        if text:
            yield text
    except Exception:
        pass
    # 4) the button's aria-label attribute ("Phone: ...")
    try:
        el = page.locator('[aria-label^="Phone:"]').first
        label = el.get_attribute("aria-label", timeout=3000) or ""
        if ":" in label:
            yield label.split(":", 1)[1].strip()
    except Exception:
        pass
    # 5) scan the page's full HTML - the number is in the page data even
    #    when the panel itself renders oddly
    try:
        html = page.content()
        m = re.search(r'href="tel:([^"]+)"', html)
        if m:
            yield m.group(1)
        best = best_mobile_in(html)
        if best:
            yield best
    except Exception:
        pass
    # 6) last resort - all visible text on the page
    try:
        best = best_mobile_in(page.locator("body").inner_text(timeout=4000))
        if best:
            yield best
    except Exception:
        pass


def extract_phone(page):
    """First candidate that is a valid mobile wins. If none is a mobile,
    fall back to the first candidate found (e.g. a landline - the output
    filter drops it, so the place is excluded as intended)."""
    fallback = ""
    for cand in _phone_candidates(page):
        if not cand:
            continue
        if clean_phone(cand):
            return cand
        if not fallback:
            fallback = cand
    return fallback


def extract_place(page, url):
    """Read one restaurant's Maps page: name, address, phone, website."""
    goto_retry(page, url)
    try:
        page.wait_for_selector("h1", timeout=15000)
    except Exception:
        return None
    try:
        page.wait_for_selector(
            '[data-item-id="address"], [data-item-id^="phone:tel:"], '
            'a[data-item-id="authority"]', timeout=8000)
    except Exception:
        pass
    # scroll the side panel - lower rows (incl. phone) can load lazily
    try:
        page.mouse.move(250, 400)
        page.mouse.wheel(0, 1500)
        page.wait_for_timeout(700)
    except Exception:
        pass
    if is_blocked(page):
        raise Blocked
    website = ""
    try:
        w = page.locator('a[data-item-id="authority"]').first
        w.wait_for(state="attached", timeout=3000)
        website = w.get_attribute("href") or ""
    except Exception:
        pass
    return {
        "name": text_of(page.locator("h1")),
        "address": clean_field(
            text_of(page.locator('button[data-item-id="address"]'))),
        "phone": extract_phone(page),
        "website": website,
    }


def fetch_emails(website):
    """Look for an email on the business's own website (home + contact)."""
    if not website or "google." in website:
        return []
    base = website.rstrip("/")
    for url in (website, base + "/contact", base + "/contact-us"):
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": BROWSER_UA})
            with urllib.request.urlopen(req, timeout=12) as resp:
                html = resp.read(400000).decode("utf-8", "replace")
        except Exception:
            continue
        found = split_emails(html)
        if found:
            return found
    return []


def scrape(location, category, max_results, want_email, headless):
    sync_playwright = import_playwright()
    query = f"{category} in {location}"
    results = []
    with sync_playwright() as p:
        browser = launch_browser(p, headless)
        ctx = browser.new_context(
            user_agent=BROWSER_UA, locale="en-IN",
            viewport={"width": 1366, "height": 800})
        page = ctx.new_page()
        try:
            print(f'  searching Google Maps for "{query}" ...')
            links = collect_links(page, query, max_results)
            print(f"  found {len(links)} places - reading each one...\n")
            debug_saved = 0
            for i, url in enumerate(links, 1):
                try:
                    d = extract_place(page, url)
                except Blocked:
                    print("\n  Google is rate-limiting this connection - "
                          "stopping here.")
                    print("  (Wait 30-60 minutes before the next run.)")
                    break
                if not d or not d["name"]:
                    continue
                if not split_phones(d["phone"]) and debug_saved < 3:
                    # keep copies of missed pages - the user can send one
                    # for an exact fix instead of guesswork
                    try:
                        debug_saved += 1
                        name = ("debug_place_page.html" if debug_saved == 1
                                else f"debug_place_page_{debug_saved}.html")
                        with open(os.path.join(HERE, name),
                                  "w", encoding="utf-8") as f:
                            f.write(page.content())
                    except Exception:
                        pass
                d["emails"] = (fetch_emails(d["website"])
                               if want_email and d["website"] else [])
                results.append(d)
                mob = split_phones(d["phone"])
                line = f"  [{i}/{len(links)}] {d['name'][:44]}"
                line += f"  - mobile {mob[0]}" if mob else "  - no mobile"
                if d["emails"]:
                    line += f"  - {d['emails'][0]}"
                print(line)
                page.wait_for_timeout(random.randint(700, 1600))
            missed = sum(1 for r in results
                         if not split_phones(r["phone"]))
            if results and missed:
                print(f"\n  {missed} place(s) had no readable mobile "
                      "number.")
                print("  A copy of one such page is saved as "
                      "debug_place_page.html")
                print("  in this folder - send it to me if those places "
                      "should have")
                print("  numbers, and I will match the reader to the page.")
        except Blocked:
            print("\n  Google blocked the search (rate limit). "
                  "Wait 30-60 minutes and run again.")
        except KeyboardInterrupt:
            print("\n  Stopped by you - saving what we have so far.")
        except Exception:
            print("\n  The search hit a connection problem and stopped.")
            print("  Check your internet connection and run again.")
        finally:
            browser.close()
    return results


# ---------------------------------------------------------------- outputs

def ensure_openpyxl():
    try:
        import openpyxl  # noqa
        return True
    except ImportError:
        print("  installing the Excel helper (one time)...")
        subprocess.run([sys.executable, "-m", "pip", "install",
                        "openpyxl", "-q"])
        try:
            import openpyxl  # noqa
            return True
        except ImportError:
            print("  (!) Excel export skipped - couldn't install openpyxl. "
                  "CSVs are still saved.")
            return False


def write_csv(path, header, values):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([header])
        for v in values:
            w.writerow([v])


def write_excel(path, leads):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    head_fill = PatternFill("solid", fgColor="1F4E5F")
    head_font = Font(bold=True, color="FFFFFF", size=12)
    for c, h in enumerate(["Restaurant Name", "Location",
                           "Phone Number", "Email ID"], 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(vertical="center")
    for r, lead in enumerate(leads, 2):
        ws.cell(row=r, column=1, value=lead["name"])
        ws.cell(row=r, column=2, value=lead["location"])
        ws.cell(row=r, column=3, value=lead["phone"])
        ws.cell(row=r, column=4, value=lead["email"])
    widths = [38, 55, 16, 34]
    for i, w_ in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w_
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 26
    wb.save(path)


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(
        description=f"Vaani Lead Engine v{VERSION} - location in, 4 lead databases "
                    "out. Scraper built in.")
    ap.add_argument("--location", help='e.g. "Gota, Ahmedabad"')
    ap.add_argument("--category", default="restaurants",
                    help="what to search (default: restaurants)")
    ap.add_argument("--max", type=int, default=100,
                    help="max places to read (default: 100)")
    ap.add_argument("--no-email", dest="email", action="store_false",
                    help="skip website email lookup for a faster run")
    ap.add_argument("--headless", action="store_true",
                    help="hide the browser window while it works")
    ap.add_argument("--no-update", action="store_true",
                    help="skip the automatic update check")
    a = ap.parse_args()
    self_update()

    print(f"\n  VAANI LEAD ENGINE v{VERSION}")
    print("  ----------------------")
    location = a.location or input(
        '\n  Type a location (e.g. "Gota, Ahmedabad"): ').strip()
    if not location:
        sys.exit("  No location given - nothing to search.")
    print()

    rows = scrape(location, a.category, a.max, a.email, a.headless)
    if not rows:
        sys.exit(
            "\n  No places were read. Either Google is rate-limiting right "
            "now\n  (wait 30-60 minutes and run again) or no listings "
            "matched.\n  Nothing was saved.\n")

    # ---- build the 4 databases -------------------------------------------
    # Rule: a restaurant with NO mobile number is not a lead - it is left
    # out of every file. (Landline-only counts as no number: landlines are
    # dropped by the phone filter, so nothing callable remains.)
    leads, seen_places = [], set()
    names, locations, phones, emails = [], [], [], []
    seen_phones, seen_emails = set(), set()
    skipped_no_phone = 0
    for r in rows:
        key = (r["name"].lower(), r["address"].lower())
        if key in seen_places:
            continue
        seen_places.add(key)
        mob = split_phones(r["phone"])
        if not mob:
            skipped_no_phone += 1
            continue
        em = r.get("emails") or []
        names.append(r["name"])
        if r["address"]:
            locations.append(r["address"])
        for p in mob:
            if p not in seen_phones:
                seen_phones.add(p)
                phones.append(p)
        for e in em:
            if e not in seen_emails:
                seen_emails.add(e)
                emails.append(e)
        leads.append({"name": r["name"], "location": r["address"],
                      "phone": ", ".join(mob), "email": ", ".join(em)})

    if not leads:
        sys.exit(
            f"\n  Found {len(rows)} places, but none of them had a mobile "
            "number\n  (landlines don't count). Nothing was saved - try a "
            "different area.\n")

    outdir = os.path.join(HERE,
                          f"leads_{slug(location)}_{date.today():%Y-%m-%d}")
    os.makedirs(outdir, exist_ok=True)
    write_csv(os.path.join(outdir, "restaurant_names.csv"),
              "restaurant_name", names)
    write_csv(os.path.join(outdir, "locations.csv"), "location", locations)
    write_csv(os.path.join(outdir, "phone_numbers.csv"), "phone_number",
              phones)
    write_csv(os.path.join(outdir, "email_ids.csv"), "email_id", emails)

    excel_done = False
    if ensure_openpyxl():
        try:
            write_excel(os.path.join(outdir, "all_leads.xlsx"), leads)
            excel_done = True
        except Exception as e:
            print(f"  (!) Excel file failed: {e}")

    # ---- summary -----------------------------------------------------------
    print("\n  DONE")
    print(f"  Restaurants in your database : {len(names)} "
          "(every one has a mobile number)")
    print(f"  Left out - no mobile number  : {skipped_no_phone}")
    print(f"  Email IDs                    : {len(emails)}")
    print(f"\n  Your 4 databases are in:\n  {outdir}")
    print("    - restaurant_names.csv")
    print("    - locations.csv")
    print("    - phone_numbers.csv   (mobiles only)")
    print("    - email_ids.csv")
    if excel_done:
        print("    - all_leads.xlsx      (everything combined)")
    print()


if __name__ == "__main__":
    main()
