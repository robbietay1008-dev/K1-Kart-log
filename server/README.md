# Kart Log server

The shop server that replaced the Google Sheet as the app's home base (2026-09-21). Runs on archer in
`/main/homelab/kartlog/` as the `kartlog` docker container. This folder is that directory minus secrets and data.

| File | What it is |
|---|---|
| `app.py` | Two FastAPI apps in one process: **public** (port 8090, published to the internet by Tailscale Funnel on :8443) = the app, device enrolment, sync, photos, version feed; **admin** (port 8091, tailnet only on :13443) = the owner's edition of the app at `/` and the data dashboard at `/dashboard` |
| `merge.py` | The app's own merge rules in Python: union of entries minus tombstones, newest stamp wins. The server's canonical snapshot is what any device would have computed |
| `admin.html` | The dashboard: summary, karts & log, inventory, **orders**, batteries, devices & invite codes, pushes, quarantine, scripts & export, raw data |
| `web/` | The built app served to devices (`index.html` is the server-mode build, `APP_BUILD` inside it is what devices compare against) |
| `Dockerfile`, `compose.yaml`, `entrypoint.sh`, `requirements.txt` | The container |

Not in the repo: `.env` (holds `KARTLOG_ADMIN_TOKEN`, the dashboard's key) and `data/` (`kartlog.db` SQLite + `photos/`).

## Ordering (ported from the sheet's `logic.gs` on 2026-09-25)

Dashboard → **orders** tab. The sheet's APP NEEDED / Place order / APP ORDERS / Book received flow, server side:

1. **To order**: every tracked part at or under red (NEEDED) or under green (WANTED). Tick ORDER, type ORDER QTY
   (checks and quantities are saved as you go, in `order_draft`). PLACE ORDER makes an `ORD-yyyymmdd-HHMM` order.
2. **Orders**: "copy as text" for the supplier. When boxes arrive, type what you have received so far on each line
   and BOOK RECEIVED. Only the new amount is added to stock (`inv` + `invTouched` in the canonical snapshot), so a
   split shipment is just a higher number later. Every device adopts the new counts on its next sync.
3. Cancel is allowed until something from the order is booked. Every booking is a row in `receipts` and an
   `admin-receive` entry in the pushes / activity log.

Endpoints (all under the admin token): `GET /api/admin/orders/needed`, `POST /api/admin/orders/draft`,
`POST /api/admin/orders/place`, `GET /api/admin/orders?all=1`, `GET /api/admin/orders/{id}/text`,
`POST /api/admin/orders/receive`, `POST /api/admin/orders/{id}/cancel`, `GET /api/admin/receipts`,
`POST /api/admin/parts/{num}/delete` (tombstones a part the way the app's DELETE PART does).

## Security (2026-09-28, audit batch 3)

- **Device text is data, never markup.** Everything a device sends is shown with `esc()` (quotes too) or
  `textContent`; the dashboard has no inline `onclick="..."` strings (buttons carry `data-act` / `data-*`).
- **Content-Security-Policy** on the app page and the dashboard: only the page's own inline `<script>` runs
  (matched by SHA-256, computed per response in `page_csp()`), `connect-src 'self'`, `frame-ancestors 'none'`,
  images only from the server or `data:`/`blob:`. Every response also has `nosniff`, `no-referrer`, `DENY`.
  `'unsafe-inline'` is listed too, but only browsers too old to know hashes (CSP level 1) use it; every current
  browser ignores it when a hash is present, so injected `on...=` handlers stay blocked there.
  A new inline script, `on...=` attribute or outside host in the pages will be blocked - test in a browser.
- **The merge refuses what the app never writes**: kart keys that are not kart numbers (push quarantined;
  "merge anyway" on the dashboard merges the rest and refuses just those keys), single kart status fields
  (that field keeps the server's value, the rest of the kart merges), entry dates / photo ids with markup
  characters, part numbers with angle brackets, backticks or control characters (quotes are fine), counts
  that are not numbers,
  photo links that are not `/api/photo/...` or `https://`. Refusals show as `rejected` in the pushes list.
- **Request caps**: enrol 64 KB, sync 5 MB, photo 9 MB, admin import 64 MB, anything else 64 KB (1 MB on the
  admin port) -> 413. A wrong enrolment code waits 0.8 s outside the lock; after 20 wrong codes in 10 minutes
  (all clients together) enrolment answers 429 until the window passes (the admin port counts separately).
- **Invite codes expire** (48 h by default, "valid for" on the dashboard; 0 = never).
- **Container**: `TZ=America/Chicago`, `mem_limit: 1g`, `pids_limit: 256`, a healthcheck on both ports, and a
  bash entrypoint that restarts the container (via `restart: unless-stopped`) when either uvicorn dies.
- FastAPI / Starlette / uvicorn are pinned exactly in `requirements.txt` (Starlette >= 1.3.1 for its 2025-26 CVEs).

## Deploy

```
cd /main/homelab/kartlog
cp <new files> .          # keep a .bak of what was there (outside web/)
docker compose up -d --build
curl -s http://127.0.0.1:8090/healthz; curl -s http://127.0.0.1:8091/healthz
docker inspect -f '{{.State.Health.Status}}' kartlog     # "healthy" about 30 s after the start
```

After a new app build is live, wait until every device in the dashboard's devices list shows it, then set the
minimum build (summary tab). Builds are `MMDD-HHMM` with no year: never leave a December minimum in place into
January (every device would be told to update to a build that is "newer" than anything served).

To try a change against real data without touching the live database: back the SQLite file up with
`sqlite3.Connection.backup` into `/tmp/kltest/data`, then run
`KARTLOG_DATA=/tmp/kltest/data KARTLOG_APP=<web dir> KARTLOG_ADMIN_TOKEN=x uvicorn app:admin --port 8099`.
