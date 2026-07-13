# K1 Kart Log

A mechanic log app for the K1 Speed Arlington kart shop. Runs entirely in the
browser on the shop iPad (works on old iPads — no App Store needed), tracks
work history, service dates, kart notes, photos, and parts used for karts
1–53, and live-syncs to the shop's Google Sheet.

## ▶ Open the app

**https://robbietay1008-dev.github.io/K1-Kart-log/**

That's the live app (served from this repo's `index.html` by GitHub Pages —
it updates automatically about a minute after every push). On the shop iPad:
open the link in Safari → Share → **Add to Home Screen** → tap the icon like
any app.

## What it does

- Tap a kart → see its service dates (chain / diff / brake flush / battery
  dates — red when 90+ days old), kart notes ("personality"), weld count,
  and full work history
- **+ LOG WORK**: date, quick-pick or typed action, parts search over the
  full K1 parts catalog (~495 part numbers) with quantities, mechanic,
  notes, photos from the camera
- Logging a brake flush / chain / diff / weld auto-updates the tracked
  date / weld count
- **Export** → an `.xlsx` Excel file: ALL DATES summary, PARTS USED
  inventory tally, and one sheet per kart in the shop's standard layout
- **Google Sheet Live Sync** → every save pushes to the shop Google Sheet's
  `APP ALL DATES`, `APP LOG`, and `APP PARTS USED` tabs (setup in
  `google-sheet-sync/SETUP_AND_CODE.txt`)

## Where the data lives

Log data is stored **on the iPad itself** (browser storage for this site).
Photos too. GitHub only holds the app's code — no shop data is published
here. That means:

- Don't clear Safari website data on the iPad — it wipes the log and photos
- Use the in-app **Backup** button regularly (it copies all log text —
  email it to yourself)
- The Google Sheet sync is the off-device copy of everything except photos

## Updating the app

Replace `index.html` with the new version (GitHub → Add file → Upload files
→ drag the new `index.html` → Commit). GitHub Pages redeploys on its own in
about a minute. The iPad keeps all its data as long as the site address
stays the same.

## Repo layout

| Path | What it is |
|---|---|
| `index.html` | The app (built, self-contained — this is what the iPad runs) |
| `src/app_template.html` | App source before parts/seed/SheetJS are injected |
| `src/seed.py` | Build script — injects parts catalog, seed data, SheetJS into the template |
| `src/parts_raw.txt` | K1 parts catalog (`number\|description` per line) |
| `src/sheetjs.min.js` | SheetJS 0.18.5 (xlsx export library, embedded at build) |
| `src/tests/` | Playwright browser tests |
| `google-sheet-sync/SETUP_AND_CODE.txt` | Google Apps Script receiver + one-time setup steps |

To rebuild `index.html` after editing the source:

```
cd src && python3 seed.py && mv kart_log.html ../index.html
```
