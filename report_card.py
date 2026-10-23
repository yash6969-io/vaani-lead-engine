#!/usr/bin/env python3
"""Vaani Report Card v1.2 - a one-page Google-presence report for one
restaurant, generated from vaani_master.db (the leads Vaani scraped).

    py report_card.py 9898148202        one lead, by phone number
    py report_card.py --top 10          the 10 best leads (score order)
    py report_card.py --all             every callable lead

Like Vaani and the CRM, it updates itself: on every start it checks
GitHub for a newer version and swaps it in automatically (silent when
offline or up to date; --no-update skips the check).

For each lead it writes report-cards/<phone>-<name>.png (WhatsApp-ready),
.pdf and .html next to this script. Every number on the card comes from
the lead database - ratings, review counts and area comparisons are the
restaurant's own scraped Google data, never invented.

Every card leads with pain, not praise: even an area leader gets the
customer-ownership pain (his reviewers are unreachable), a thin-lead
warning when #2 is close, and a popularity-vs-rating gap when one
exists. Leads that cannot be helped - national chains, Do Not Contact,
and the genuinely fully set up - are skipped with reasons, to save
calling time. (--force overrides the skip for one phone number.)
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
VERSION = "1.2.0"
UPDATE_URL = ("https://raw.githubusercontent.com/yash6969-io/"
              "vaani-lead-engine/main/report_card.py")
SELLER_NAME = "Yash Tanna"
SELLER_PHONE = ""          # e.g. "+91 98XXX XXXXX" - shown on the card
BRAND = "Book My Food"

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "vaani_master.db")
OUT_DIR = os.path.join(HERE, "report-cards")


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
    if os.environ.get("VAANI_RC_NO_UPDATE") == "1" \
            or "--no-update" in sys.argv:
        return
    try:
        req = urllib.request.Request(UPDATE_URL,
                                     headers={"User-Agent": "VaaniRC"})
        with urllib.request.urlopen(req, timeout=5) as r:
            content = r.read().decode("utf-8", "replace")
    except Exception:
        return
    m = re.search(r'^VERSION\s*=\s*"([0-9.]+)"', content, re.M)
    if not m or _version_tuple(m.group(1)) <= _version_tuple(VERSION):
        return
    if "Vaani Report Card" not in content or len(content) < 10000:
        return
    try:
        compile(content, "report_card.py", "exec")
    except Exception:
        return
    me = os.path.abspath(__file__)
    try:
        shutil.copyfile(me, me + ".bak")
        with open(me, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        return
    print(f"\n  Updated Report Card to v{m.group(1)} - restarting...\n")
    env = dict(os.environ)
    env["VAANI_RC_NO_UPDATE"] = "1"
    subprocess.Popen([sys.executable, me] + sys.argv[1:], env=env)
    sys.exit(0)


# ------------------------------------------------------------------ data

KNOWN_CHAINS = ("mcdonald", "domino", "pizza hut", "kfc", "burger king",
                "subway", "starbucks", "haldiram", "barbeque nation",
                "wow momo", "wow! momo", "la pino", "baskin",
                "cafe coffee", "jumbo king", "behrouz", "faasos",
                "freshmenu", "oven story", "chaayos", "third wave",
                "burger singh")

def verdict(lead):
    """Should this lead be pitched at all? Returns (decision, reason).
    The skip list exists to save calling time, and it only fires on
    evidence the database actually holds: chains cannot buy locally,
    suppressed leads must not be called, and an 'Operations upgrade'
    angle is only assigned when the scrape found a website AND direct
    ordering already in place - with a strong rating and review base
    on top, that is the weakest possible fit for a pain-led pitch."""
    name_low = (lead["name"] or "").lower()
    if any(c in name_low for c in KNOWN_CHAINS):
        return ("skip", "national chain - decisions are centralised, "
                "the local owner cannot buy")
    if lead["suppressed"]:
        return ("skip", "marked Do not contact / Not interested / "
                "Wrong number")
    angle = (lead["angle"] or "").strip()
    if (angle == "Operations upgrade"
            and (lead["rating"] or 0) >= 4.2
            and (lead["reviews"] or 0) >= 200):
        return ("skip", "already has direct ordering and a solid "
                "setup - weakest fit; skip to save time")
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
        if any(c in (r["name"] or "").lower() for c in KNOWN_CHAINS):
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


def gaps_and_wins(lead, peers, st):
    gaps, wins = [], []
    angle = (lead["angle"] or "").strip()
    reviews = lead["reviews"] or 0
    rating = lead["rating"]
    # --- product pains, from the scrape intel ---
    if angle == "Aggregator-dependent":
        gaps.append(("No direct ordering channel",
                     "Orders run through Swiggy/Zomato - every single "
                     "one pays commission, even from your regulars."))
    elif angle == "No direct ordering":
        gaps.append(("No direct ordering option",
                     "Customers who find you still have to order "
                     "through an app that charges commission."))
    if not (lead["website"] or "").strip():
        gaps.append(("No website on your Google listing",
                     "A customer who wants to reach you directly "
                     "has nowhere to land."))
    top = st["top"]
    if (top is not None and top["phone"] != lead["phone"]
            and (top["reviews"] or 0) > 2 * max(1, reviews)):
        gaps.append((f"{top['name']} shows up before you",
                     f"They have {(top['reviews'] or 0):,} Google "
                     f"reviews - about "
                     f"{(top['reviews'] or 0) // max(1, reviews)}x "
                     "yours. More reviews means Google shows them "
                     "first."))
    if (rating and st["avg_rating"]
            and rating < st["avg_rating"] - 0.05):
        gaps.append(("Rating below your area's average",
                     f"Your area averages {st['avg_rating']:.1f} "
                     "stars. Even a small gap costs you clicks "
                     "every day."))
    # --- pains every restaurant has, leader or not ---
    if reviews >= 100:
        gaps.append(("Your regulars are not reachable",
                     f"{reviews:,} people have reviewed you on "
                     "Google - but their names and numbers sit with "
                     "Google and the apps. You cannot send a single "
                     "one of them tomorrow's offer."))
    elif reviews > 0:
        gaps.append(("Too little proof for new customers",
                     f"Only {reviews} Google reviews - most customers "
                     "scroll past a listing with so little proof and "
                     "pick a name they have seen reviewed more."))
    else:
        gaps.append(("Almost invisible on Google",
                     "No reviews yet - a new customer searching your "
                     "area will almost never find you."))
    if st["rank"] == 1 and peers:
        second = sorted(peers, key=lambda r: -(r["reviews"] or 0))[0]
        diff = reviews - (second["reviews"] or 0)
        if 0 <= diff < reviews:
            gaps.append(("Your lead is thinner than it looks",
                         f"{second['name']} is only {diff:,} reviews "
                         "behind you. One strong season from them - "
                         "or one quiet one from you - and the area's "
                         "top spot changes hands."))
    if rating and rating < 4.5 and reviews >= 300:
        gaps.append(("Popularity is ahead of the rating",
                     f"{reviews:,} reviews at {rating:.1f} stars "
                     "means plenty of average-or-worse experiences "
                     "are sitting in public, where every new customer "
                     "reads them first."))
    # --- wins, capped at two: praise opens the door, pain sells ---
    if rating and rating >= 4.3:
        wins.append(("Strong rating",
                     f"{rating:.1f} stars is genuinely good - "
                     "customers already trust the food."))
    if st["rank"] and st["rank"] <= 3 and st["n"] >= 4:
        title = ("The most-reviewed restaurant in your area"
                 if st["rank"] == 1 else
                 f"Top {st['rank']} most-reviewed in your area")
        wins.append((title,
                     "You are one of the names your area already "
                     "knows."))
    if reviews >= 500:
        wins.append(("A deep base of happy customers",
                     f"{reviews:,} reviews is social proof most "
                     "restaurants never build."))
    if (lead["website"] or "").strip():
        wins.append(("You have a website",
                     "That is the base a direct-ordering channel is "
                     "built on - most competitors don't have one."))
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
  sans-serif; background:#ECEEF1; color:#17122B; }
.card { width:1080px; margin:0 auto; background:#fff; }
.head { background:linear-gradient(120deg,#2A1052 0%,#4C1D95 58%,
  #7C3AED 130%); color:#fff; padding:46px 54px 40px; position:relative;
  overflow:hidden; }
.head::after { content:''; position:absolute; width:430px; height:430px;
  border-radius:50%; top:-230px; right:-90px;
  background:radial-gradient(circle,rgba(216,180,254,.35),
  rgba(216,180,254,0) 65%); }
.kicker { font-size:15px; font-weight:700; letter-spacing:.22em;
  color:#D8B4FE; text-transform:uppercase; }
.head h1 { font-size:47px; font-weight:800; letter-spacing:-.015em;
  margin:10px 0 6px; line-height:1.12; position:relative; z-index:1; }
.head .addr { font-size:19px; color:#DCD2F5; position:relative;
  z-index:1; }
.head .meta { margin-top:16px; font-size:14.5px; color:#B9A8E8;
  position:relative; z-index:1; }
.stats { display:flex; gap:16px; padding:30px 54px 4px; }
.stat { flex:1; background:#F7F5FC; border:1px solid #E9E4F5;
  border-radius:16px; padding:20px 22px 16px; }
.stat .lbl { font-size:13px; font-weight:700; letter-spacing:.09em;
  text-transform:uppercase; color:#6B6390; }
.stat .big { font-size:44px; font-weight:800; letter-spacing:-.02em;
  margin-top:4px; display:flex; align-items:center; gap:9px; }
.stat .sub { font-size:14.5px; color:#6B6390; margin-top:2px; }
.sec { padding:28px 54px 2px; }
.sec h2 { font-size:22px; font-weight:800; letter-spacing:-.01em;
  display:flex; align-items:center; gap:10px; }
.sec h2::after { content:''; flex:1; height:1px; background:#E9E4F5; }
.barrow { display:flex; align-items:center; gap:14px; margin-top:15px; }
.barrow .who { width:300px; font-size:16.5px; font-weight:600;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.barrow .track { flex:1; height:15px; background:#F0EDF7;
  border-radius:8px; overflow:hidden; }
.barrow .fill { height:100%; border-radius:8px; background:#C9BFE8; }
.barrow .fill.me { background:linear-gradient(90deg,#6D28D9,#8B5CF6); }
.barrow .fill.top { background:#17122B; }
.barrow .val { width:104px; font-size:16.5px; font-weight:700;
  font-variant-numeric:tabular-nums; }
.grp { font-size:14px; font-weight:700; letter-spacing:.08em;
  text-transform:uppercase; color:#6B6390; margin-top:20px; }
.item { display:flex; gap:15px; padding:15px 2px; }
.item + .item { border-top:1px solid #F0EDF7; }
.item .ic { flex:none; width:38px; height:38px; border-radius:11px;
  display:flex; align-items:center; justify-content:center; }
.item .ic.win { background:#DCFCE7; }
.item .ic.gap { background:#FEF3C7; }
.item b { font-size:18.5px; display:block; }
.item p { font-size:15.5px; color:#5B5470; line-height:1.45;
  margin-top:2px; }
.math { margin:28px 54px 0; background:#F7F5FC;
  border:1px solid #E9E4F5; border-radius:18px; padding:26px 28px 22px; }
.math h2 { font-size:22px; font-weight:800; }
.math .rows { margin-top:14px; }
.mrow { display:flex; align-items:center; gap:14px; padding:11px 0; }
.mrow + .mrow { border-top:1px dashed #DCD2F5; }
.mrow .what { flex:1; font-size:17px; font-weight:600; }
.mrow .kept { font-size:21px; font-weight:800; white-space:nowrap; }
.mrow .kept.bad { color:#B42318; }
.mrow .kept.good { color:#15803D; }
.math .note { font-size:14px; color:#6B6390; margin-top:10px;
  line-height:1.5; }
.foot { padding:30px 54px 40px; display:flex; gap:20px;
  align-items:center; }
.foot .cta { flex:1; }
.foot .cta b { font-size:21px; display:block; line-height:1.3; }
.foot .cta span { font-size:16px; color:#5B5470; }
.foot .brand { text-align:right; }
.foot .brand .bname { font-size:19px; font-weight:800; color:#4C1D95; }
.foot .brand .who2 { font-size:14.5px; color:#6B6390; margin-top:2px; }
.fine { padding:0 54px 34px; font-size:12.5px; color:#9A93B5;
  line-height:1.55; }
"""


def bar(label, value, maxv, cls, shown):
    pct = 0 if not maxv else max(3, round(value * 100 / maxv))
    return (f'<div class="barrow"><div class="who">{esc(label)}</div>'
            f'<div class="track"><div class="fill {cls}" '
            f'style="width:{pct}%"></div></div>'
            f'<div class="val">{esc(shown)}</div></div>')


def build_html(lead, peers, st):
    gaps, wins = gaps_and_wins(lead, peers, st)
    rating = lead["rating"]
    stars = ""
    if rating:
        full = int(round(rating))
        stars = "".join(STAR.format(s=26, c="#F59E0B") for _ in range(full))
        stars += "".join(STAR.format(s=26, c="#DDD6F0")
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
                     f'</div><div class="sub">of {st["n"]} restaurants '
                     f'tracked in {esc(area)}</div></div>')
    else:
        rank_stat = (f'<div class="stat"><div class="lbl">Restaurants '
                     f'tracked nearby</div><div class="big">{st["n"]}'
                     f'</div><div class="sub">in your area, from Google '
                     f'listings</div></div>')
    stats = (
        f'<div class="stats">'
        f'<div class="stat"><div class="lbl">Google rating</div>'
        f'<div class="big">'
        f'{f"{rating:.1f}" if rating else "-"} {stars}</div>'
        f'<div class="sub">what customers see first</div></div>'
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
    gaps_html = (f'<div class="sec"><h2>What\'s costing you</h2>'
                 f'{items(gaps, "gap", ALERT)}</div>') if gaps else ""
    phone_line = (f' &middot; WhatsApp {esc(SELLER_PHONE)}'
                  if SELLER_PHONE else "")
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>{CSS}</style></head><body><div class="card">
<div class="head"><div class="kicker">Restaurant Report Card</div>
<h1>{esc(lead['name'])}</h1>
<div class="addr">{esc(lead['address'])}</div>
<div class="meta">Prepared on {today} &middot; Based on your public
Google listing</div></div>
{stats}
{compare}
{wins_html}
{gaps_html}
<div class="math"><h2>The commission math</h2>
<div class="rows">
<div class="mrow"><div class="what">On a &#8377;500 Swiggy/Zomato order,
the platforms keep &#8377;125-175 (25-35% all-in)</div>
<div class="kept bad">You keep &#8377;325-375</div></div>
<div class="mrow"><div class="what">On a &#8377;500 order from your own
ordering page</div><div class="kept good">You keep &#8377;500</div></div>
</div>
<div class="note">Same food, same customer. The only difference is who
takes the cut - and whether the customer stays yours.</div></div>
<div class="foot"><div class="cta"><b>Want the costing-you list fixed
before the Diwali rush?</b><span>One short call - I will show you your
own numbers, on your own menu.</span></div>
<div class="brand"><div class="bname">{esc(BRAND)}</div>
<div class="who2">{esc(SELLER_NAME)}{phone_line}</div></div></div>
<div class="fine">This report uses only your public Google listing data
and the {st['n']} restaurants tracked in your area - no estimates about
your kitchen or your books. Commission ranges are reported 2026 figures
for Swiggy/Zomato (25-35% effective, incl. fees and GST effects).</div>
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
    return s[:40] or "restaurant"


def make_card(conn, lead):
    peers = find_peers(conn, lead)
    st = area_stats(lead, peers)
    html_text = build_html(lead, peers, st)
    os.makedirs(OUT_DIR, exist_ok=True)
    base = os.path.join(OUT_DIR, f"{lead['phone']}-{slug(lead['name'])}")
    render(html_text, base)
    return base


def main():
    self_update()
    args = [a for a in sys.argv[1:] if a != "--no-update"]
    if not os.path.exists(DB_PATH):
        print("vaani_master.db not found next to report_card.py.")
        print("Put report_card.py in your Vaani folder and run Vaani "
              "at least once first.")
        return 1
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    leads = []
    if args and args[0] == "--top":
        n = int(args[1]) if len(args) > 1 else 10
        leads = conn.execute(
            "SELECT * FROM leads WHERE suppressed = 0 ORDER BY"
            " score DESC LIMIT ?", (n,)).fetchall()
    elif args and args[0] == "--all":
        leads = conn.execute(
            "SELECT * FROM leads WHERE suppressed = 0 ORDER BY"
            " score DESC").fetchall()
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
