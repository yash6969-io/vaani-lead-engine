# Clinic Reports (Content & Trust Report)

One-page reports for clinics and small hospitals, built from
`vaani_healthcare.db` — the separate healthcare database. They
are the report half of the Content365 WhatsApp kit the CRM's
Healthcare section sends.

## The easy way

You normally never run this by hand: open the CRM's Healthcare
section (run-crm-healthcare.bat), go to **Send Kit**, and the
card is generated automatically when you click **Auto-send kit**
(or the Generate button). If `clinic_report.py` is missing from
the folder, the CRM downloads it by itself.

## By hand

```
py clinic_report.py 9898148202     one clinic, by phone number
py clinic_report.py --top 10       the 10 most-reviewed clinics
py clinic_report.py --all          every callable clinic
```

Or double-click **run-clinic-reports.bat** (top 10 + opens the
folder). Cards land in `clinic-reports\` as PNG (WhatsApp), PDF
and HTML.

Needs a healthcare scrape first:
`py vaani.py --vertical healthcare --category "Dental clinics"`.

## What the card says

- The clinic's Google rating, reviews, and rank in its area —
  every number from the scrape, nothing invented.
- **What's costing you**: content pains only — no website, no
  social page linked from the website (checked live when the
  card is made, and only claimed when the check succeeded),
  proof sitting unused, a rival showing up first, a rating
  below the area average.
- What Content365 makes (365 posts, ₹3,999 incl. GST for the
  year) and the free first-week offer.

Chains and hospital groups (Apollo, Clove Dental, Dr Batra's…)
and already-fully-set-up clinics are skipped with reasons —
`--force` overrides for one phone number.

Like Vaani and the CRM, this tool updates itself from GitHub on
every start. Your clinic data never leaves your PC.
