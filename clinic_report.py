#!/usr/bin/env python3
"""Vaani Clinic Report v1.0.0 - a one-page Content & Trust
Report for one clinic or small hospital, generated from
vaani_healthcare.db (the HEALTHCARE leads Vaani scraped - a
separate database from the restaurant one; they never mix).

    py clinic_report.py 9898148202      one clinic, by phone number
    py clinic_report.py --top 10        the 10 most-reviewed clinics
    py clinic_report.py --all           every callable clinic

Like Vaani and the CRM, it updates itself from GitHub on every
start (silent when offline or up to date; --no-update skips it).

For each clinic it writes clinic-reports/<phone>-<name>.png
(WhatsApp-ready), .pdf and .html next to this script. Every number
on the card comes from the lead database - the clinic's own
scraped Google data, never invented. The report is focused on the
social content, captions and publishing rhythm Content365 provides.

The pitch is content: patients find a clinic on Google, then check
it before booking. A strong listing with silence behind it leaks
trust - that is what Content365 fills (365 ready-to-post branded
posts). Chains and hospital groups, suppressed leads, and the rare
clinic already fully set up online are skipped with reasons.
(--force overrides the skip for one phone number.)
"""
import html
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import urllib.request
from datetime import date

# ------------------------------------------------------------------ config
VERSION = "1.1.0"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/clinic_report.py")
SELLER_NAME = "Yash Tanna"
SELLER_PHONE = ""          # e.g. "+91 98XXX XXXXX" - shown on the card
C365_WEBSITE = "https://content365.si/"
BRAND = "Content365"

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "vaani_healthcare.db")
OUT_DIR = os.path.join(HERE, "clinic-reports")


def esc(x):
    return html.escape(str(x if x is not None else ""))


# ------------------------------------------------------------------ update

def _version_tuple(text):
    try:
        return tuple(int(x) for x in text.strip().split("."))
    except Exception:
        return (0,)


def self_update():
    """Same deal as Vaani and the CRM: check online for a newer
    report_card.py, swap it in and restart. Silent when offline or
    up to date. Only the program file is downloaded - your lead data
    never leaves this PC."""
    if os.environ.get("VAANI_CLINIC_NO_UPDATE") == "1" \
            or "--no-update" in sys.argv:
        return
    try:
        req = urllib.request.Request(UPDATE_URL,
                                     headers={"User-Agent": "VaaniClinic"})
        with urllib.request.urlopen(req, timeout=5) as r:
            content = r.read().decode("utf-8", "replace")
    except Exception:
        return
    m = re.search(r'^VERSION\s*=\s*"([0-9.]+)"', content, re.M)
    if not m or _version_tuple(m.group(1)) <= _version_tuple(VERSION):
        return
    if "Vaani Clinic Report" not in content or len(content) < 10000:
        return
    try:
        compile(content, "clinic_report.py", "exec")
    except Exception:
        return
    me = os.path.abspath(__file__)
    try:
        shutil.copyfile(me, me + ".bak")
        with open(me, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        return
    print(f"\n  Updated Clinic Report to v{m.group(1)} - restarting...\n")
    env = dict(os.environ)
    env["VAANI_CLINIC_NO_UPDATE"] = "1"
    subprocess.Popen([sys.executable, me] + sys.argv[1:], env=env)
    sys.exit(0)


# ------------------------------------------------------------------ data

CLINIC_CHAINS = ("apollo", "fortis", "manipal", "medanta",
                 "max hospital", "narayana", "cloudnine",
                 "columbia asia", "dr batra", "kaya", "vlcc", "oliva",
                 "clove dental", "sabka dentist", "toothsi", "dentzz",
                 "partha dental", "vasan", "dr agarwal",
                 "centre for sight", "kokilaben", "hinduja",
                 "rainbow hospitals", "motherhood", "thyrocare",
                 "dr lal", "metropolis diagnostics", "kims hospital",
                 "yashoda", "care hospitals", "aster", "wockhardt")

def verdict(lead, site=None):
    """Choose Content365 prospects without evaluating website ownership."""
    name_low = (lead["name"] or "").lower()
    if any(c in name_low for c in CLINIC_CHAINS):
        return ("skip", "chain / hospital group - marketing is "
                "centralised, the local clinic cannot buy")
    if lead["suppressed"]:
        return ("skip", "marked Do not contact / Not interested / "
                "Wrong number")
    return ("pitch", "")


def locality_tokens(address):
    segs = [s.strip() for s in str(address or "").split(",") if s.strip()]
    segs = [s for s in segs if not re.fullmatch(r"[\d\s\-]+", s)]
    if len(segs) > 1:
        segs = segs[:-1]          # last segment is city/state - too broad
    return [s.lower() for s in segs if len(s) >= 3]


def find_peers(conn, lead):
    """Area peers for comparison. National chains are excluded: an
    owner measures himself against local rivals, and a chain's review
    count would skew every average and rank on the card."""
    toks = set(locality_tokens(lead["address"]))
    rows = conn.execute(
        "SELECT * FROM leads WHERE phone != ? AND suppressed = 0",
        (lead["phone"],)).fetchall()
    peers = []
    for r in rows:
        if any(c in (r["name"] or "").lower() for c in CLINIC_CHAINS):
            continue
        if toks & set(locality_tokens(r["address"])):
            peers.append(r)
    return peers


def area_stats(lead, peers):
    group = [lead] + peers
    rated = [r for r in group if r["rating"]]
    avg_rating = (sum(r["rating"] for r in rated) / len(rated)
                  if rated else None)
    avg_reviews = (sum((r["reviews"] or 0) for r in group) / len(group)
                   if group else 0)
    by_reviews = sorted(group, key=lambda r: -(r["reviews"] or 0))
    rank = by_reviews.index(lead) + 1 if lead in by_reviews else None
    top = by_reviews[0] if by_reviews else None
    return {"n": len(group), "avg_rating": avg_rating,
            "avg_reviews": avg_reviews, "rank": rank, "top": top}


_SITE_CACHE = {}


def fetch_site(lead):
    """Compatibility placeholder for older callers; no site is scanned."""
    return {"state": "not-checked", "socials": []}


def gaps_and_wins(lead, peers, st, site):
    gaps, wins = [], []
    reviews = lead["reviews"] or 0
    rating = lead["rating"]
    # --- content opportunities: what Content365 can create and publish ---
    if reviews >= 150:
        gaps.append(("Your patient trust is not yet a content library",
                     f"You already have {reviews:,} public patient reviews. "
                     "Those real trust signals can become a steady stream "
                     "of educational, reassuring and awareness-day posts."))
    elif reviews > 0:
        gaps.append(("Your content foundation is still small",
                     f"Your {reviews:,} public reviews are a useful start. "
                     "A consistent series of helpful posts can keep your "
                     "clinic visible between patient searches."))
    else:
        gaps.append(("Your social content story has not started yet",
                     "Without a regular stream of useful health content, "
                     "patients have fewer reminders of what your clinic "
                     "does and why they should trust it."))
    top = st["top"]
    if (top is not None and top["phone"] != lead["phone"]
            and (top["reviews"] or 0) > 2 * max(1, reviews)):
        gaps.append((f"{top['name']} shows up before you",
                     f"They have {(top['reviews'] or 0):,} Google "
                     f"reviews - about "
                     f"{(top['reviews'] or 0) // max(1, reviews)}x "
                     "yours. Patients compare review counts "
                     "first, and pick the clinic that looks chosen."))
    if (rating and st["avg_rating"]
            and rating < st["avg_rating"] - 0.05):
        gaps.append(("Rating below your area's average",
                     f"Your area averages {st['avg_rating']:.1f} "
                     "stars. Patients read that gap as the area's "
                     "verdict - it costs you first visits "
                     "every week."))
    if st["rank"] == 1 and peers:
        second = sorted(peers, key=lambda r: -(r["reviews"] or 0))[0]
        diff = reviews - (second["reviews"] or 0)
        if 0 <= diff < reviews:
            gaps.append(("Your lead is thinner than it looks",
                         f"{second['name']} is only {diff:,} reviews "
                         "behind you. One strong season from them "
                         "- or one quiet one from you - and the "
                         "area's top spot changes hands."))
    if rating and rating < 4.5 and reviews >= 300:
        gaps.append(("Popularity is ahead of the rating",
                     f"{reviews:,} reviews at {rating:.1f} stars "
                     "means plenty of average-or-worse visits are "
                     "sitting in public, where every new patient "
                     "reads them first."))
    if not gaps:
        gaps.append(("No consistent content rhythm",
                     "Patients remember clinics they keep seeing. "
                     "Health tips, awareness days, seasonal guidance "
                     "and simple clinic stories give Content365 a "
                     "clear rhythm to build for you."))
    # --- wins, capped at two: praise opens the door, pain sells ---
    if rating and rating >= 4.3:
        wins.append(("Strong rating",
                     f"{rating:.1f} stars is genuinely good - "
                     "patients already trust your care. That trust "
                     "is the hardest part of marketing, and you "
                     "have it."))
    if st["rank"] and st["rank"] <= 3 and st["n"] >= 4:
        title = ("The most-reviewed clinic in your area"
                 if st["rank"] == 1 else
                 f"Top {st['rank']} most-reviewed in your area")
        wins.append((title,
                     "You are one of the names your area already "
                     "knows."))
    if reviews >= 500:
        wins.append(("A deep bank of patient trust",
                     f"{reviews:,} reviews give Content365 strong real-world "
                     "proof to turn into reassuring and educational posts."))
    elif reviews >= 50:
        wins.append(("Enough proof for a strong content series",
                     f"Your {reviews:,} reviews provide real themes and "
                     "patient trust for a useful recurring post plan."))
    if rating and rating >= 4.5:
        wins.append(("A reputation worth showing consistently",
                     "Your rating gives every educational and awareness "
                     "post a credible clinic story behind it."))
    return gaps[:3], wins[:2]


# ------------------------------------------------------------------ html

STAR = ('<svg viewBox="0 0 24 24" width="{s}" height="{s}" fill="{c}">'
        '<path d="M12 2l2.9 6.6 7.1.6-5.4 4.7 1.6 7-6.2-3.7L5.8 21l1.6-7'
        'L2 9.2l7.1-.6z"/></svg>')
CHECK = ('<svg viewBox="0 0 24 24" width="20" height="20" fill="none" '
         'stroke="#15803D" stroke-width="2.6" stroke-linecap="round" '
         'stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>')
ALERT = ('<svg viewBox="0 0 24 24" width="20" height="20" fill="none" '
         'stroke="#B45309" stroke-width="2.4" stroke-linecap="round" '
         'stroke-linejoin="round"><path d="M10.3 3.9L1.8 18a2 2 0 001.7 '
         '3h17a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/><path d="M12 9v4"/>'
         '<path d="M12 17h.01"/></svg>')

CSS = """
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI', Inter, system-ui, -apple-system, Arial,
  sans-serif; background:#E9EFED; color:#0C1F1E; }
.card { width:1080px; margin:0 auto; background:#fff; }
.head { background:linear-gradient(120deg,#052E2B 0%,#0B4F4A 58%,
  #0F766E 130%); color:#fff; padding:46px 54px 40px; position:relative;
  overflow:hidden; }
.head::after { content:''; position:absolute; width:430px; height:430px;
  border-radius:50%; top:-230px; right:-90px;
  background:radial-gradient(circle,rgba(153,246,228,.35),
  rgba(153,246,228,0) 65%); }
.kicker { font-size:15px; font-weight:700; letter-spacing:.22em;
  color:#99F6E4; text-transform:uppercase; }
.head h1 { font-size:47px; font-weight:800; letter-spacing:-.015em;
  margin:10px 0 6px; line-height:1.12; position:relative; z-index:1; }
.head .addr { font-size:19px; color:#CDEAE4; position:relative;
  z-index:1; }
.head .meta { margin-top:16px; font-size:14.5px; color:#9CCFC6;
  position:relative; z-index:1; }
.stats { display:flex; gap:16px; padding:30px 54px 4px; }
.stat { flex:1; background:#F2F8F6; border:1px solid #DDEEEA;
  border-radius:16px; padding:20px 22px 16px; }
.stat .lbl { font-size:13px; font-weight:700; letter-spacing:.09em;
  text-transform:uppercase; color:#52706A; }
.stat .big { font-size:44px; font-weight:800; letter-spacing:-.02em;
  margin-top:4px; display:flex; align-items:center; gap:9px; }
.stat .sub { font-size:14.5px; color:#52706A; margin-top:2px; }
.sec { padding:28px 54px 2px; }
.sec h2 { font-size:22px; font-weight:800; letter-spacing:-.01em;
  display:flex; align-items:center; gap:10px; }
.sec h2::after { content:''; flex:1; height:1px; background:#DDEEEA; }
.barrow { display:flex; align-items:center; gap:14px; margin-top:15px; }
.barrow .who { width:300px; font-size:16.5px; font-weight:600;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.barrow .track { flex:1; height:15px; background:#E7F1EE;
  border-radius:8px; overflow:hidden; }
.barrow .fill { height:100%; border-radius:8px; background:#BCDCD6; }
.barrow .fill.me { background:linear-gradient(90deg,#0F766E,#14B8A6); }
.barrow .fill.top { background:#0C1F1E; }
.barrow .val { width:104px; font-size:16.5px; font-weight:700;
  font-variant-numeric:tabular-nums; }
.grp { font-size:14px; font-weight:700; letter-spacing:.08em;
  text-transform:uppercase; color:#52706A; margin-top:20px; }
.item { display:flex; gap:15px; padding:15px 2px; }
.item + .item { border-top:1px solid #E7F1EE; }
.item .ic { flex:none; width:38px; height:38px; border-radius:11px;
  display:flex; align-items:center; justify-content:center; }
.item .ic.win { background:#DCFCE7; }
.item .ic.gap { background:#FEF3C7; }
.item b { font-size:18.5px; display:block; }
.item p { font-size:15.5px; color:#47605A; line-height:1.45;
  margin-top:2px; }
.math { margin:28px 54px 0; background:#F2F8F6;
  border:1px solid #DDEEEA; border-radius:18px; padding:26px 28px 22px; }
.math h2 { font-size:22px; font-weight:800; }
.math .rows { margin-top:14px; }
.mrow { display:flex; align-items:center; gap:14px; padding:11px 0; }
.mrow + .mrow { border-top:1px dashed #CDEAE4; }
.mrow .what { flex:1; font-size:17px; font-weight:600; }
.mrow .kept { font-size:21px; font-weight:800; white-space:nowrap; }
.mrow .kept.bad { color:#B42318; }
.mrow .kept.good { color:#15803D; }
.math .note { font-size:14px; color:#52706A; margin-top:10px;
  line-height:1.5; }
.foot { padding:30px 54px 40px; display:flex; gap:20px;
  align-items:center; }
.foot .cta { flex:1; }
.foot .cta b { font-size:21px; display:block; line-height:1.3; }
.foot .cta span { font-size:16px; color:#47605A; }
.foot .brand { text-align:right; }
.foot .brand .bname { font-size:19px; font-weight:800; color:#0B4F4A; }
.foot .brand .who2 { font-size:14.5px; color:#52706A; margin-top:2px; }
.fine { padding:0 54px 34px; font-size:12.5px; color:#84A39C;
  line-height:1.55; }
"""


def bar(label, value, maxv, cls, shown):
    pct = 0 if not maxv else max(3, round(value * 100 / maxv))
    return (f'<div class="barrow"><div class="who">{esc(label)}</div>'
            f'<div class="track"><div class="fill {cls}" '
            f'style="width:{pct}%"></div></div>'
            f'<div class="val">{esc(shown)}</div></div>')


def build_html(lead, peers, st, site):
    gaps, wins = gaps_and_wins(lead, peers, st, site)
    rating = lead["rating"]
    stars = ""
    if rating:
        full = int(round(rating))
        stars = "".join(STAR.format(s=26, c="#F59E0B") for _ in range(full))
        stars += "".join(STAR.format(s=26, c="#D8E8E4")
                         for _ in range(5 - full))
    area = ""
    toks = locality_tokens(lead["address"])
    if toks:
        area = toks[-1].title()
    today = date.today().strftime("%d %B %Y")
    rank_stat = ""
    if st["rank"] and st["n"] >= 4:
        rank_stat = (f'<div class="stat"><div class="lbl">Area rank '
                     f'by reviews</div><div class="big">#{st["rank"]}'
                     f'</div><div class="sub">of {st["n"]} clinics '
                     f'tracked in {esc(area)}</div></div>')
    else:
        rank_stat = (f'<div class="stat"><div class="lbl">Clinics '
                     f'tracked nearby</div><div class="big">{st["n"]}'
                     f'</div><div class="sub">in your area, from Google '
                     f'listings</div></div>')
    stats = (
        f'<div class="stats">'
        f'<div class="stat"><div class="lbl">Google rating</div>'
        f'<div class="big">'
        f'{f"{rating:.1f}" if rating else "-"} {stars}</div>'
        f'<div class="sub">what patients see first</div></div>'
        f'<div class="stat"><div class="lbl">Google reviews</div>'
        f'<div class="big">{lead["reviews"] or 0:,}</div>'
        f'<div class="sub">your social proof so far</div></div>'
        f'{rank_stat}</div>')
    maxr = max([lead["rating"] or 0] + [p["rating"] or 0 for p in peers]
               ) or 1
    maxv = max([lead["reviews"] or 0] + [p["reviews"] or 0 for p in peers]
               ) or 1
    compare = ""
    if peers:
        top = st["top"]
        compare += '<div class="grp">Rating</div>'
        compare += bar("You", rating or 0, 5, "me",
                       f'{rating:.1f}' if rating else "-")
        if st["avg_rating"]:
            compare += bar("Area average", st["avg_rating"], 5, "",
                           f'{st["avg_rating"]:.1f}')
        if top and top["phone"] != lead["phone"] and top["rating"]:
            compare += bar(top["name"], top["rating"], 5, "top",
                           f'{top["rating"]:.1f}')
        compare += '<div class="grp">Reviews</div>'
        compare += bar("You", lead["reviews"] or 0, maxv, "me",
                       f'{(lead["reviews"] or 0):,}')
        compare += bar("Area average", round(st["avg_reviews"]), maxv,
                       "", f'{round(st["avg_reviews"]):,}')
        if top and top["phone"] != lead["phone"]:
            compare += bar(top["name"], top["reviews"] or 0, maxv,
                           "top", f'{(top["reviews"] or 0):,}')
        compare = (f'<div class="sec"><h2>How you stack up in '
                   f'{esc(area)}</h2>{compare}</div>')
    def items(rows, kind, icon):
        out = ""
        for title, text in rows:
            out += (f'<div class="item"><div class="ic {kind}">'
                    f'{icon}</div><div><b>{esc(title)}</b>'
                    f'<p>{esc(text)}</p></div></div>')
        return out
    wins_html = (f'<div class="sec"><h2>What\'s working</h2>'
                 f'{items(wins, "win", CHECK)}</div>') if wins else ""
    gaps_html = (f'<div class="sec"><h2>Content opportunities</h2>'
                 f'{items(gaps, "gap", ALERT)}</div>') if gaps else ""
    panel_rows = ""
    for t in (
        "A post for every day of the year - health days, "
        "festivals and seasons land on the right dates",
        "Your clinic's name, logo and phone printed on every "
        "post - never misspelled",
        "Captions written for Instagram, Facebook, LinkedIn "
        "and X, plus Hindi",
        "You review every post before it goes public - "
        "nothing posts without your OK",
    ):
        panel_rows += (f'<div class="item"><div class="ic win">'
                       f'{CHECK}</div><div><p style="font-size:'
                       f'16.5px;color:#173B37">{esc(t)}</p></div>'
                       f'</div>')
    panel = (f'<div class="math"><h2>What Content365 makes '
             f'for you</h2>{panel_rows}<div class="note">The '
             f'full year is &#8377;3,999 including GST. See '
             f'posts made for clinics like yours: '
             f'content365.si/industries</div></div>')
    phone_line = (f' &middot; WhatsApp {esc(SELLER_PHONE)}'
                  if SELLER_PHONE else "")
    site_line = (f'<br>{esc(C365_WEBSITE)}' if C365_WEBSITE else "")
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>{CSS}</style></head><body><div class="card">
<div class="head"><div class="kicker">Content365 Social Content Report</div>
<h1>{esc(lead['name'])}</h1>
<div class="addr">{esc(lead['address'])}</div>
<div class="meta">Prepared on {today} &middot; Built from your public
Google reputation to plan better social content</div></div>
{stats}
{compare}
{wins_html}
{gaps_html}
{panel}
<div class="foot"><div class="cta"><b>Want to see your first week
of posts?</b><span>I will make 7 days of posts with your clinic's
name and logo - free. Reply on this WhatsApp and judge them
yourself.</span></div>
<div class="brand"><div class="bname">{esc(BRAND)}</div>
<div class="who2">{esc(SELLER_NAME)}{phone_line}{site_line}</div></div></div>
<div class="fine">This report uses only your public Google listing data
and the {st['n']} clinics tracked in your area - nothing here is estimated
or invented. Content365 turns the opportunity into branded social posts,
captions and a publishing calendar for your approval.</div>
</div></body></html>"""


# ------------------------------------------------------------------ render

def render(html_text, base):
    from playwright.sync_api import sync_playwright
    html_path = base + ".html"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_text)
    with sync_playwright() as pw:
        b = pw.chromium.launch()
        pg = b.new_page(viewport={"width": 1120, "height": 900},
                        device_scale_factor=2)
        pg.goto("file://" + html_path)
        pg.wait_for_timeout(350)
        h = pg.evaluate(
            "document.querySelector('.card').getBoundingClientRect()"
            ".height")
        pg.set_viewport_size({"width": 1120, "height": int(h) + 40})
        card = pg.locator(".card")
        card.screenshot(path=base + ".png")
        pg.pdf(path=base + ".pdf", width="1080px",
               height=f"{int(h)}px", print_background=True,
               page_ranges="1")
        b.close()
    return html_path


def slug(name):
    s = re.sub(r"[^a-z0-9]+", "-", str(name or "").lower()).strip("-")
    return s[:40] or "clinic"


def make_card(conn, lead):
    peers = find_peers(conn, lead)
    st = area_stats(lead, peers)
    site = {"state": "not-checked", "socials": []}
    html_text = build_html(lead, peers, st, site)
    os.makedirs(OUT_DIR, exist_ok=True)
    base = os.path.join(OUT_DIR, f"{lead['phone']}-{slug(lead['name'])}")
    render(html_text, base)
    return base


def main():
    self_update()
    args = [a for a in sys.argv[1:] if a != "--no-update"]
    if not os.path.exists(DB_PATH):
        print("vaani_healthcare.db not found next to clinic_report.py.")
        print("Put clinic_report.py in your Vaani folder and run a "
              "healthcare search first:  py vaani.py --vertical "
              "healthcare --category \"Dental clinics\"")
        return 1
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    leads = []
    if args and args[0] == "--top":
        n = int(args[1]) if len(args) > 1 else 10
        leads = conn.execute(
            "SELECT * FROM leads WHERE suppressed = 0 ORDER BY"
            " reviews DESC LIMIT ?", (n,)).fetchall()
    elif args and args[0] == "--all":
        leads = conn.execute(
            "SELECT * FROM leads WHERE suppressed = 0 ORDER BY"
            " reviews DESC").fetchall()
    elif args:
        digits = re.sub(r"\D", "", args[0])[-10:]
        leads = conn.execute(
            "SELECT * FROM leads WHERE phone = ?",
            (digits,)).fetchall()
        if not leads:
            print(f"No lead found with phone ending {digits}.")
            return 1
    else:
        print(__doc__)
        return 0
    force = "--force" in args
    skipped = []
    if not force:
        keep = []
        for lead in leads:
            decision, reason = verdict(lead)
            if decision == "skip":
                skipped.append((lead, reason))
            else:
                keep.append(lead)
        leads = keep
    if leads:
        print(f"Making {len(leads)} report card(s)...")
        for lead in leads:
            base = make_card(conn, lead)
            print(f"  {lead['name']} -> {os.path.basename(base)}.png")
    if skipped:
        print("\nSkipped - don't pitch these:")
        for lead, reason in skipped:
            print(f"  {lead['name']} - {reason}")
    conn.close()
    print(f"Done. Cards are in: {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
