# Vaani Lead Engine v3.8 — scored leads for Book My Food

**NEW in v3.8 — works hand-in-hand with the Vaani CRM app.** If you
manage your leads in the CRM (`vaani_crm.py`, its own download), any
lead that has a status in the CRM belongs to the CRM: Vaani reads that
status and never overrides it from the tracker again. Leads the CRM
doesn't know still follow the tracker exactly as before.

**NEW in v3.7 — your Not interested leads are kept, not lost.** Anyone you
mark **Not interested** in the tracker is now saved into
`vaani-not-interested.xlsx` (kept next to Vaani): name, phone, location,
email, score, tier, the date you marked them, and your tracker notes —
your win-back list for when you want to try them again in a few months.
They stay out of your daily leads, but they are never deleted. **Do Not
Contact** stays permanent (it is never in the win-back file) and **Wrong
number** stays dropped. Changed your mind about someone? Set their
Outcome in the tracker to anything else (e.g. Interested) and the next
run un-marks them automatically.

**NEW in v3.6 — quality scoring (Phase 1).** Every lead is now scored out
of 100 from evidence on its Maps page and its own website (rating, review
count, Zomato/Swiggy presence, direct ordering, website, socials) and
lands in the tracker sorted best-first with a Tier (A = call today,
B = nurture, C = skip), a Pitch angle, and the reasons for its score.
Closed businesses are excluded, a master database (vaani_master.db)
remembers every business across runs, and anyone you mark Not
interested / Wrong number / Do not contact in the tracker never comes
back into your daily leads. The tracker's Outcome dropdown is now the
full sales funnel: Staff reached, Decision-maker reached, Demo booked,
Demo done, Won…

**NEW in v3.5 — leads go straight into your call tracker.** When a run
finishes, the leads are written into `vaani-call-tracker.xlsx` (kept next
to Vaani) in a sheet named for that day — e.g. **21 Oct 2026**. The next
day's run creates that day's sheet in the same file. No copy-paste. A
second run on the same day adds only restaurants that aren't already in
that day's sheet. v3.5 is also faster: restaurant websites are checked
for emails several at a time, and the browser skips heavy map photos.

**NEW in v3.4 — automatic updates.** Every time you start Vaani, it first
checks online (about 2 seconds) whether a newer version exists. If yes, it
updates itself and restarts with the new version. If not — or if there's no
internet — it just runs as-is. You never have to download, delete, unblock
or reinstall Vaani again. This zip is the last one you will ever install.

Your leads are never uploaded anywhere — the check only *downloads* the
program file.

Type a location (even a small area like **"Gota, Ahmedabad"**) and get 4 lead
databases about the restaurants there:

| File | What is inside |
|---|---|
| `restaurant_names.csv` | Every restaurant name found |
| `locations.csv` | Their addresses |
| `phone_numbers.csv` | **Mobile numbers only** — landlines are dropped on purpose |
| `email_ids.csv` | Email addresses found on the restaurants' own websites |
| `all_leads.xlsx` | Everything combined in one Excel file |

**What changed in v3:** the scraper is built into the script itself. No Docker,
no separate engine file to download, no ports, nothing else to start. One
folder, one command. It works the way the well-known open-source Google Maps
scrapers work — it opens Google Maps in a real browser window, reads each
restaurant, then visits the restaurant's own website to look for an email.

Also inside this folder: **`vaani-call-tracker.xlsx`** — after your scrape,
paste the leads in and track call outcomes (Interested / Callback / Deck sent…)
while you work through them.

---

## Setup — one time only

### Step 1 — Install Python (skip if you already have it)
1. Go to https://www.python.org/downloads/ and download Python for Windows.
2. Run the installer. On the **first screen**, tick **"Add python.exe to PATH"** — this step matters.
3. Finish the install.

### Step 2 — Unblock and extract this folder
1. Right-click `vaani-lead-engine-v3.4.zip` → **Properties** → tick **Unblock** → **OK**.
   (Windows blocks files downloaded from the internet — this one tick removes problems later.)
2. Extract the zip. Keep the extracted folder somewhere easy, like your Desktop.

### Step 3 — Install the helper tools
Open a terminal **inside** the extracted folder:
- Open the folder in File Explorer → **Shift + right-click** on empty space → **Open in Terminal**.

Run these two commands (see `COMMANDS.txt` — it's also in the folder):

```
py -m pip install -r requirements.txt
py -m playwright install chromium
```

The second command downloads a browser (one time, a few minutes). If your
computer already has Google Chrome, the tool will use that instead and this
download is just a backup.

---

## Every run

In the same terminal, inside the folder:

```
py vaani.py
```

1. Type your location when asked, e.g. `Gota, Ahmedabad`, and press Enter.
2. A browser window opens and searches Google Maps by itself. **Don't close it.**
3. Watch the terminal — it prints each restaurant as it reads it, with the
   mobile number and email it found.
4. When it says **DONE**, your 4 databases + Excel are in a new folder named
   `leads_<location>_<date>` next to `vaani.py`.

**Shortcut:** double-clicking `run-vaani.bat` does the setup check and the
same run for you.

Useful options:

```
py vaani.py --max 50          first 50 places only (quick test)
py vaani.py --no-email        skip the website email lookup (faster)
py vaani.py --category "tiffin services"    search a different business type
py vaani.py --vertical healthcare --category "Dental clinics"
                              clinics/hospitals go into their OWN
                              healthcare database - see below
```

## Two separate worlds: restaurants and Healthcare

Restaurant leads and healthcare leads **never share a database**.
With `--vertical healthcare`, everything a run touches is
healthcare-labelled and separate: the database
(`vaani_healthcare.db`), the call tracker
(`vaani-healthcare-tracker.xlsx`), the win-back list, and an output
folder named `leads_healthcare_<location>_<date>` with a
`clinic_names.csv` inside. Your restaurant database
(`vaani_master.db`) is not opened or changed at all - and the
restaurant CRM, which reads only `vaani_master.db`, never sees a
clinic. If you forget the flag and search a clinic-like category,
Vaani warns you before anything is saved and tells you the right
command.

## TPO cells (`--vertical tpo`) — Content365 college channel

TPO runs work differently from every other vertical, because the
number you want is NOT on Google Maps. Maps gives a college's main
number; the TPO cell's number lives on the college website's
placement page. So a TPO run does this:

1. Takes the college list from Maps for the location you give.
2. Opens each college website, goes straight to its placement
   page — following homepage links AND trying the usual
   placement addresses directly — then follows that page's
   contact/team links one level deeper (max 14 pages per
   college), and reads the number shown there as the cell's own
   contact. Pages are read in full (up to 2 MB), so footer
   contact blocks on large modern sites are no longer cut off.
3. Saves the college ONLY if such a number exists.

The rule is strict, by instruction: **only TPO cell numbers are
saved**. The college's main number is never counted as a TPO
number, and any other staff member's number is deleted — it is not
stored in the database, the Excel file, the tracker, or the CSVs.
When a number's context is ambiguous (for example a cell block that
also names a person), the number is dropped rather than guessed.
TPO office landlines ARE kept and tagged as landline, because a
cell's published number is often a landline; mobiles are preferred
when a page lists both. Generic TPO emails (`tpo@…`,
`placement@…`) are kept as a secondary contact; personal emails are
dropped.

TPO data gets its own database (`vaani_tpo.db`), tracker
(`vaani-tpo-tracker.xlsx`) and output folder
(`leads_tpo_<location>_<date>`), so it never mixes with restaurant
or clinic data. The CRM picks it up as a "TPO Cells" workspace
automatically once the database exists.

Pilot command (Ahmedabad only):

```
py vaani.py --vertical tpo --category "Colleges" --location "Ahmedabad"
```

Expect partial coverage: many colleges publish no separate TPO
number at all (only an email, a form, or an officer's name), so the
run summary's "Left out - no TPO cell number" count is the honest
yield figure. The direct-placement crawl is v3.15.0 in the
workspace and reaches your PC only when it is published.

## Delivery-first mid-tier segment (`--segment delivery`)

Restaurants only. Adds one extra file to the run folder,
`delivery_leads.xlsx` — the places most likely to feel aggregator
commission and to say yes to their own ordering channel. Nothing is
removed from the normal outputs, the tracker or the database; this is
a focused call list on top of them.

A lead qualifies only if it meets **all** of these:

- **Delivery listed, plus dine-in OR cloud kitchen** on Maps — it
  lives on delivery and either has a real dining room or is a cloud
  kitchen (delivery-only, detected from its Maps category or name).
  Cloud kitchens are in on purpose: they pay aggregator commission
  on every order and survive on repeat orders from a loyal audience.
- **Not a chain** and **not a cart / stall / kiosk** — no scaled
  brands, no tiny vendors with no dining or delivery setup.
- **Not too popular:** 30–1500 reviews, and never a near-perfect
  rating (4.7+) with 500+ reviews — those are the famous,
  dining-led brands.
- **Loyal-base evidence:** at least 30 reviews and a rating of 3.8
  or more, so regulars are demonstrably coming back.
- **Commission pain visible:** on Swiggy/Zomato, or no direct
  ordering found.
- **Everyday prices:** Maps shows ₹ or ₹₹ (budget-to-moderate).
  ₹₹₹/₹₹₹₹ places are left out. When Maps shows no price band, the
  lead still qualifies but is marked "Price unknown - check menu",
  and the summary tells you how many need a menu check.

Run it like this:

```
py vaani.py --segment delivery --location "Prahlad Nagar, Ahmedabad"
```

The end-of-run summary prints the segment count, how many are
price-known vs price-unknown, and why the rest were left out
(no delivery, no dine-in, too popular, premium price, ...) so you can
see whether the thresholds fit your area. For the repeat-orders-from-
offices effect, run it in office-dense areas — business districts,
IT parks, commercial streets — where the same people order lunch
every day.

Every row in `delivery_leads.xlsx` (and in `all_leads.xlsx`) now
carries a **Date Assigned** column with the run's date, so you can
tell at a glance whether you are looking at freshly generated leads
or ones added in an earlier run — even after files are copied or
merged. The call tracker already groups leads by run-date sheet;
this makes the date travel with each row in the Excel files too.

Honest limits: Maps' ₹ level is a rough band, not dish prices, so
₹100–400 is approximated, not proven per dish. "Loyal base" and
"commission pain" are proxies from public signals, not certainties —
the segment file shows the reason for every lead so you can judge
each one before calling.

## Honest notes

- **Only restaurants WITH a mobile number are saved.** If a restaurant
  lists no number on Maps — or only a landline — it is left out of every
  file completely. Every row in your database is a restaurant you can
  actually call. The end-of-run summary tells you how many were left out.
- **Landline numbers are dropped by design.** Only Indian mobiles
  (starting 6–9) are kept. Numbers written with a leading 0 — `098981 48202`,
  the format most businesses use on Maps — are read as mobiles correctly;
  true STD landlines (`079 …`, `011 …`) are still dropped.
- **Emails come only from the business's own website.** If a restaurant has
  no website, or no email on it, there is no email for it. Nothing is guessed
  or invented.
- **Google rate-limits scraping.** One search at a time, stay logged out of
  Google in that browser window, and if a run finds nothing or stops early,
  wait 30–60 minutes and run again. If a CAPTCHA appears in the browser
  window, solve it there and the run continues.
- A full run of 100 places takes roughly 5–10 minutes. Press **Ctrl+C** any
  time to stop — whatever was found so far still gets saved.

## Troubleshooting

| Problem | Fix |
|---|---|
| File/folder blocked, or window flashes and dies | Right-click the **zip** → Properties → **Unblock** → OK, then extract again fresh. |
| `'py' is not recognized` | Reinstall Python and tick **"Add python.exe to PATH"** on the first screen. |
| "Playwright is not installed" | Run `py -m pip install -r requirements.txt` in the folder. |
| "Could not start a browser" | Run `py -m playwright install chromium` in the folder. |
| Run finds 0 places | Google is rate-limiting — wait 30–60 minutes, run again, one search at a time. |
| Run lists restaurants but no numbers | The tool saves `debug_place_page.html` in the Vaani folder in this case — send that file to Muse for an exact fix. |
| Excel file missing | The CSVs are the real output; run `py -m pip install openpyxl` and re-run to get the Excel too. |
