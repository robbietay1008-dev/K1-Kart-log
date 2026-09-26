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

## Deploy

```
cd /main/homelab/kartlog
cp <new files> .          # keep a .bak of what was there
docker compose up -d --build
curl -s http://127.0.0.1:8090/healthz; curl -s http://127.0.0.1:8091/healthz
```

To try a change against real data without touching the live database: back the SQLite file up with
`sqlite3.Connection.backup` into `/tmp/kltest/data`, then run
`KARTLOG_DATA=/tmp/kltest/data KARTLOG_APP=<web dir> KARTLOG_ADMIN_TOKEN=x uvicorn app:admin --port 8099`.
