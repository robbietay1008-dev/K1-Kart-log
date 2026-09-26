"""K1 Kart Log server: the app's home base, replacing the Google Sheet receiver.

Two FastAPI apps in one process file:
  public  (port 8090, published by Tailscale Funnel)  - the app itself, device enrolment, sync, photos, version feed
  admin   (port 8091, tailnet only)                    - the full data view, devices/invites, audit, quarantine, scripts

Data: one SQLite file (DATA_DIR/kartlog.db). The canonical snapshot is one JSON document merged with every push
(merge.py = the app's own merge rules); every push is kept verbatim in the audit table; refused pushes go to
quarantine with the reason. Photos are files under DATA_DIR/photos.
"""
import base64
import hashlib
import json
import os
import re
import secrets
import sqlite3
import threading
import time

from fastapi import APIRouter, Depends, FastAPI, Header, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, PlainTextResponse, Response

import merge as M

DATA_DIR = os.environ.get("KARTLOG_DATA", "/data")
APP_DIR = os.environ.get("KARTLOG_APP", os.path.join(os.path.dirname(os.path.abspath(__file__)), "web"))
DB_PATH = os.path.join(DATA_DIR, "kartlog.db")
PHOTO_DIR = os.path.join(DATA_DIR, "photos")
ADMIN_TOKEN = os.environ.get("KARTLOG_ADMIN_TOKEN", "")
_LOCK = threading.Lock()
_SAFE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,80}$")


# ------------------------------------------------------------------------------------------ storage
def db():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(PHOTO_DIR, exist_ok=True)
    c = sqlite3.connect(DB_PATH, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA busy_timeout=30000")
    return c


def init():
    # two processes (public + admin) start together: WAL is set once, with a retry instead of a crash
    for attempt in range(20):
        try:
            with db() as c:
                c.execute("PRAGMA journal_mode=WAL")
            break
        except sqlite3.OperationalError:
            time.sleep(0.25)
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS snapshot (id INTEGER PRIMARY KEY CHECK (id = 1), json TEXT NOT NULL, version INTEGER NOT NULL,
                                             updated_at REAL NOT NULL, updated_by TEXT);
        CREATE TABLE IF NOT EXISTS devices (token TEXT PRIMARY KEY, label TEXT, created_at REAL, last_seen REAL, build TEXT,
                                            pushes INTEGER DEFAULT 0, revoked INTEGER DEFAULT 0, invite TEXT);
        CREATE TABLE IF NOT EXISTS invites (code TEXT PRIMARY KEY, label TEXT, created_at REAL, uses INTEGER DEFAULT 0,
                                            max_uses INTEGER DEFAULT 1, revoked INTEGER DEFAULT 0);
        CREATE TABLE IF NOT EXISTS audit (id INTEGER PRIMARY KEY AUTOINCREMENT, at REAL, device TEXT, build TEXT, bytes INTEGER,
                                          saved_at TEXT, report TEXT, version_after INTEGER, raw TEXT);
        CREATE TABLE IF NOT EXISTS quarantine (id INTEGER PRIMARY KEY AUTOINCREMENT, at REAL, device TEXT, build TEXT,
                                               reason TEXT, raw TEXT, resolved INTEGER DEFAULT 0);
        CREATE TABLE IF NOT EXISTS photos (id TEXT PRIMARY KEY, kart TEXT, date TEXT, device TEXT, at REAL, path TEXT, bytes INTEGER);
        CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE IF NOT EXISTS order_draft (part TEXT PRIMARY KEY, chk INTEGER DEFAULT 0, qty INTEGER, at REAL);
        CREATE TABLE IF NOT EXISTS orders (id TEXT PRIMARY KEY, at REAL, note TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS order_lines (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id TEXT, part TEXT, name TEXT,
                                                ordered INTEGER, booked INTEGER DEFAULT 0, last_booked REAL);
        CREATE TABLE IF NOT EXISTS receipts (id INTEGER PRIMARY KEY AUTOINCREMENT, at REAL, line_id INTEGER, order_id TEXT,
                                             part TEXT, qty INTEGER, before REAL, after REAL, version INTEGER);
        """)
        for table in ("invites", "devices"):
            cols = {r[1] for r in c.execute(f"PRAGMA table_info({table})")}
            if "mechanic" not in cols:
                c.execute(f"ALTER TABLE {table} ADD COLUMN mechanic TEXT DEFAULT ''")


def setting(key, default=None):
    with db() as c:
        r = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default


def set_setting(key, value):
    with db() as c:
        c.execute("INSERT INTO settings(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))


def load_snapshot():
    with db() as c:
        r = c.execute("SELECT json, version, updated_at FROM snapshot WHERE id=1").fetchone()
    if not r:
        return None, 0, None
    return json.loads(r["json"]), r["version"], r["updated_at"]


def store_snapshot(snap, version, who):
    with db() as c:
        c.execute("INSERT INTO snapshot(id, json, version, updated_at, updated_by) VALUES(1, ?, ?, ?, ?) "
                  "ON CONFLICT(id) DO UPDATE SET json=excluded.json, version=excluded.version, updated_at=excluded.updated_at, updated_by=excluded.updated_by",
                  (json.dumps(snap, separators=(",", ":")), version, time.time(), who))


def build_of(s):
    """app builds are 'MMDD-HHMM' strings; compare as (year-less) tuples so a min build can be enforced"""
    m = re.match(r"^(\d{2})(\d{2})-(\d{2})(\d{2})$", str(s or ""))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))) if m else (0, 0, 0, 0)


def current_build():
    try:
        html = open(os.path.join(APP_DIR, "index.html"), encoding="utf-8").read()
        m = re.search(r'var APP_BUILD = (?:/\*__BUILD__\*/)?"([^"]*)"', html)
        return m.group(1) if m else "dev"
    except Exception:
        return "dev"


def apply_push(payload, device, build):
    """merge a device's snapshot into the canonical copy; returns (ok, response dict)"""
    problems = M.sanity(payload)
    now = int(time.time() * 1000)
    with _LOCK:
        if problems:
            with db() as c:
                c.execute("INSERT INTO quarantine(at, device, build, reason, raw) VALUES(?, ?, ?, ?, ?)",
                          (time.time(), device, build, "; ".join(problems), json.dumps(payload)[:4_000_000]))
            return False, {"ok": False, "quarantined": True, "reason": "; ".join(problems)}
        canon, version, _ = load_snapshot()
        merged, report = M.merge(canon, payload, now)
        version += 1
        store_snapshot(merged, version, device)
        with db() as c:
            c.execute("INSERT INTO audit(at, device, build, bytes, saved_at, report, version_after, raw) VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                      (time.time(), device, build, len(json.dumps(payload)), payload.get("savedAt") or "", json.dumps(report), version,
                       json.dumps(payload, separators=(",", ":"))[:4_000_000]))
            c.execute("UPDATE devices SET last_seen=?, build=?, pushes=pushes+1 WHERE token=?", (time.time(), build, device))
            c.execute("DELETE FROM audit WHERE id NOT IN (SELECT id FROM audit ORDER BY id DESC LIMIT 400)")
    changed = {k: v for k, v in report.items() if v}
    return True, {"ok": True, "version": version, "changes": changed, "counts": M.counts(merged)}


# ------------------------------------------------------------------------------------------ public app
public = FastAPI(title="K1 Kart Log", docs_url=None, redoc_url=None)
shared = APIRouter()


def device_auth(x_device_token: str = Header(None), x_app_build: str = Header(None)):
    if not x_device_token:
        raise HTTPException(401, "no device token")
    with db() as c:
        d = c.execute("SELECT * FROM devices WHERE token=?", (x_device_token,)).fetchone()
    if not d or d["revoked"]:
        raise HTTPException(403, "this device is not enrolled (ask for a new code)")
    return {"token": d["token"], "label": d["label"], "build": x_app_build or "", "mechanic": d["mechanic"] or ""}


def must_update(build):
    mb = setting("min_build", "")
    return bool(mb) and build_of(build) < build_of(mb)


@shared.get("/api/version")
def api_version(x_app_build: str = Header(None)):
    return {"build": current_build(), "minBuild": setting("min_build", ""), "mustUpdate": must_update(x_app_build or "")}


@shared.post("/api/enroll")
async def api_enroll(request: Request):
    body = await request.json()
    code = str(body.get("code") or "").strip().upper()
    label = str(body.get("label") or "")[:60]
    if not code:
        raise HTTPException(400, "code required")
    with _LOCK, db() as c:
        inv = c.execute("SELECT * FROM invites WHERE code=?", (code,)).fetchone()
        if not inv or inv["revoked"] or (inv["max_uses"] and inv["uses"] >= inv["max_uses"]):
            time.sleep(0.8)   # slow down guessing
            raise HTTPException(403, "that code is not valid")
        token = secrets.token_urlsafe(32)
        mech = (inv["mechanic"] or "").upper()
        c.execute("INSERT INTO devices(token, label, created_at, last_seen, invite, mechanic) VALUES(?, ?, ?, ?, ?, ?)",
                  (token, label or inv["label"] or code, time.time(), time.time(), code, mech))
        c.execute("UPDATE invites SET uses=uses+1 WHERE code=?", (code,))
    return {"ok": True, "token": token, "label": label or inv["label"] or code, "mechanic": mech}


@shared.get("/api/snapshot")
def api_snapshot(dev=Depends(device_auth)):
    if must_update(dev["build"]):
        return JSONResponse({"ok": False, "mustUpdate": True, "build": current_build()}, status_code=426)
    snap, version, updated_at = load_snapshot()
    with db() as c:
        c.execute("UPDATE devices SET last_seen=?, build=? WHERE token=?", (time.time(), dev["build"], dev["token"]))
    if not snap:
        return {"ok": False, "empty": True, "version": 0}
    out = dict(snap)
    out.update({"ok": True, "version": version, "savedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(updated_at or 0)),
                "serverTime": int(time.time() * 1000), "build": current_build(),
                "device": {"label": dev["label"], "mechanic": dev.get("mechanic") or ""}})
    return out


@shared.post("/api/sync")
async def api_sync(request: Request, dev=Depends(device_auth)):
    if must_update(dev["build"]):
        return JSONResponse({"ok": False, "mustUpdate": True, "build": current_build()}, status_code=426)
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(400, "bad json")
    ok, resp = apply_push(payload, dev["token"], dev["build"])
    resp["build"] = current_build()
    return resp


@shared.post("/api/photo")
async def api_photo(request: Request, dev=Depends(device_auth)):
    body = await request.json()
    pid = str(body.get("id") or "")
    data_url = str(body.get("dataURL") or "")
    if not _SAFE.match(pid) or not data_url.startswith("data:image/"):
        raise HTTPException(400, "bad photo")
    m = re.match(r"^data:image/(jpeg|jpg|png|webp);base64,(.+)$", data_url, re.S)
    if not m:
        raise HTTPException(400, "unsupported image")
    raw = base64.b64decode(m.group(2))
    if len(raw) > 6_000_000:
        raise HTTPException(413, "photo too large")
    ext = "jpg" if m.group(1) in ("jpeg", "jpg") else m.group(1)
    path = os.path.join(PHOTO_DIR, f"{pid}.{ext}")
    with open(path, "wb") as fh:
        fh.write(raw)
    with _LOCK:
        with db() as c:      # one connection at a time: a nested connection would wait on this one's write lock
            c.execute("INSERT INTO photos(id, kart, date, device, at, path, bytes) VALUES(?, ?, ?, ?, ?, ?, ?) "
                      "ON CONFLICT(id) DO UPDATE SET path=excluded.path, bytes=excluded.bytes",
                      (pid, str(body.get("kart") or ""), str(body.get("date") or ""), dev["token"], time.time(), path, len(raw)))
        snap, version, _ = load_snapshot()
        if snap is not None:
            snap.setdefault("photos", {})[pid] = f"/api/photo/{pid}"
            store_snapshot(snap, version + 1, dev["token"])
    return {"ok": True, "url": f"/api/photo/{pid}"}


@shared.get("/api/photo/{pid}")
def api_photo_get(pid: str, t: str = None, x_device_token: str = Header(None)):
    """an <img>/link cannot send headers: the device token may come as ?t= instead"""
    device_auth(x_device_token or t, None)
    if not _SAFE.match(pid):
        raise HTTPException(400, "bad id")
    with db() as c:
        r = c.execute("SELECT path FROM photos WHERE id=?", (pid,)).fetchone()
    if not r or not os.path.isfile(r["path"]):
        raise HTTPException(404, "no such photo")
    return FileResponse(r["path"])


@shared.get("/healthz")
def healthz():
    snap, version, _ = load_snapshot()
    return {"ok": True, "version": version, "build": current_build()}


def _static(name, media):
    p = os.path.join(APP_DIR, name)
    if not os.path.isfile(p):
        raise HTTPException(404, "not built yet")
    return FileResponse(p, media_type=media, headers={"Cache-Control": "no-cache"})


def _page(request):
    """the app page; on the admin port the ADMIN flag is switched on (same build, more screens)"""
    p = os.path.join(APP_DIR, "index.html")
    if not os.path.isfile(p):
        raise HTTPException(404, "not built yet")
    html = open(p, encoding="utf-8").read()
    if request.app is admin:
        html = html.replace("/*__ADMIN__*/false", "true", 1)
    return HTMLResponse(html, headers={"Cache-Control": "no-cache"})


@shared.get("/", response_class=HTMLResponse)
def app_index(request: Request):
    return _page(request)


@shared.get("/index.html")
def app_index2(request: Request):
    return _page(request)


@shared.get("/manifest.webmanifest")
def app_manifest():
    return _static("manifest.webmanifest", "application/manifest+json")


@shared.get("/sw.js")
def app_sw():
    return _static("sw.js", "application/javascript")


@shared.get("/icon-{size}.png")
def app_icon(size: str):
    if size not in ("192", "512", "180"):
        raise HTTPException(404)
    return _static(f"icon-{size}.png", "image/png")


# ------------------------------------------------------------------------------------------ admin app
admin = FastAPI(title="K1 Kart Log admin", docs_url=None, redoc_url=None)


def admin_auth(authorization: str = Header(None)):
    """the admin port is tailnet-only; the token is a second lock (set KARTLOG_ADMIN_TOKEN)"""
    if ADMIN_TOKEN:
        tok = (authorization or "").replace("Bearer ", "").strip()
        if not secrets.compare_digest(tok, ADMIN_TOKEN):
            raise HTTPException(403, "admin token")
    return True


@admin.get("/dashboard", response_class=HTMLResponse)
def admin_dashboard():
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "admin.html")
    return HTMLResponse(open(p, encoding="utf-8").read())


@admin.post("/api/admin/self-enroll")
async def adm_self_enroll(request: Request, _=Depends(admin_auth)):
    """the owner's own device: enrolled with the admin token instead of a code"""
    body = await request.json()
    label = "admin · " + str(body.get("label") or "device")[:50]
    token = secrets.token_urlsafe(32)
    mech = (setting("admin_mechanic", "ROBERT") or "").upper()
    with _LOCK, db() as c:
        c.execute("INSERT INTO devices(token, label, created_at, last_seen, invite, mechanic) VALUES(?, ?, ?, ?, ?, ?)",
                  (token, label, time.time(), time.time(), "admin", mech))
    return {"ok": True, "token": token, "label": label, "mechanic": mech}


@admin.get("/api/admin/activity")
def adm_activity(limit: int = 40, _=Depends(admin_auth)):
    """what everyone is doing: the last pushes, with the entries each one added, plus today's entries by mechanic"""
    with db() as c:
        labels = {r["token"]: r["label"] for r in c.execute("SELECT token, label FROM devices")}
        rows = [dict(r) for r in c.execute("SELECT id, at, device, build, report FROM audit ORDER BY id DESC LIMIT ?", (min(limit, 200),))]
        devs = [dict(r) for r in c.execute("SELECT label, last_seen, build, pushes, revoked FROM devices ORDER BY last_seen DESC")]
        q = c.execute("SELECT COUNT(*) n FROM quarantine WHERE resolved=0").fetchone()["n"]
    events = []
    for r in rows:
        try:
            rep = json.loads(r["report"])
        except Exception:
            rep = {}
        changed = {k: v for k, v in rep.items() if v and k not in ("added", "future_stamps")}
        events.append({"at": r["at"], "device": labels.get(r["device"], r["device"] if str(r["device"] or "").startswith("admin-") else "unknown device"),
                       "build": r["build"], "added": rep.get("added") or [], "changed": changed})
    snap, _v, _u = load_snapshot()
    today = time.strftime("%-m/%-d/%Y") if os.name != "nt" else time.strftime("%#m/%#d/%Y")
    by_mech = {}
    if snap:
        for k, v in (snap.get("karts") or {}).items():
            for e in v.get("entries") or []:
                if e.get("date") == today:
                    by_mech.setdefault(e.get("mechanic") or "?", []).append({"kart": k, "action": e.get("action"), "parts": e.get("parts")})
    return {"events": events, "devices": devs, "quarantine": q, "today": today, "by_mechanic": by_mech}


@admin.get("/api/admin/summary")
def adm_summary(_=Depends(admin_auth)):
    snap, version, updated_at = load_snapshot()
    with db() as c:
        devices = c.execute("SELECT COUNT(*) n FROM devices WHERE revoked=0").fetchone()["n"]
        q = c.execute("SELECT COUNT(*) n FROM quarantine WHERE resolved=0").fetchone()["n"]
        last = c.execute("SELECT at, device, build, report FROM audit ORDER BY id DESC LIMIT 1").fetchone()
        photos = c.execute("SELECT COUNT(*) n, COALESCE(SUM(bytes),0) b FROM photos").fetchone()
    return {"version": version, "updated_at": updated_at, "counts": M.counts(snap) if snap else None, "devices": devices,
            "quarantine": q, "build": current_build(), "min_build": setting("min_build", ""),
            "last_push": dict(last) if last else None, "photos": {"n": photos["n"], "bytes": photos["b"]},
            "db_bytes": os.path.getsize(DB_PATH) if os.path.isfile(DB_PATH) else 0}


@admin.get("/api/admin/snapshot")
def adm_snapshot(_=Depends(admin_auth)):
    snap, version, updated_at = load_snapshot()
    return {"version": version, "updated_at": updated_at, "snapshot": snap}


@admin.post("/api/admin/import")
async def adm_import(request: Request, _=Depends(admin_auth)):
    """import a snapshot (raw JSON, or the JSONP the sheet endpoint returns) as a push from 'admin-import'"""
    raw = (await request.body()).decode("utf-8", errors="replace")
    m = re.match(r"^\s*\w+\((.*)\)\s*;?\s*$", raw, re.S)
    try:
        payload = json.loads(m.group(1) if m else raw)
    except Exception:
        raise HTTPException(400, "not json")
    payload.setdefault("app", "k1kartlog")
    payload.setdefault("type", "snapshot")
    ok, resp = apply_push(payload, "admin-import", payload.get("appBuild") or "import")
    return resp


@admin.get("/api/admin/devices")
def adm_devices(_=Depends(admin_auth)):
    with db() as c:
        rows = [dict(r) for r in c.execute("SELECT token, label, created_at, last_seen, build, pushes, revoked, invite, mechanic FROM devices ORDER BY last_seen DESC")]
    for r in rows:
        r["token"] = r["token"][:6] + "…"
        r["id"] = hashlib.sha1(r["token"].encode()).hexdigest()[:8]
    with db() as c:
        full = [r["token"] for r in c.execute("SELECT token FROM devices ORDER BY last_seen DESC")]
    for r, t in zip(rows, full):
        r["ref"] = t
    return {"devices": rows}


@admin.post("/api/admin/devices/revoke")
async def adm_revoke(request: Request, _=Depends(admin_auth)):
    body = await request.json()
    with db() as c:
        c.execute("UPDATE devices SET revoked=? WHERE token=?", (0 if body.get("restore") else 1, str(body.get("ref") or "")))
    return {"ok": True}


@admin.post("/api/admin/devices/mechanic")
async def adm_device_mechanic(request: Request, _=Depends(admin_auth)):
    body = await request.json()
    with db() as c:
        c.execute("UPDATE devices SET mechanic=? WHERE token=?", (str(body.get("mechanic") or "").strip().upper()[:30], str(body.get("ref") or "")))
    return {"ok": True}


@admin.get("/api/admin/invites")
def adm_invites(_=Depends(admin_auth)):
    with db() as c:
        return {"invites": [dict(r) for r in c.execute("SELECT * FROM invites ORDER BY created_at DESC")]}


@admin.post("/api/admin/invites")
async def adm_invite_new(request: Request, _=Depends(admin_auth)):
    body = await request.json()
    code = "".join(secrets.choice("ABCDEFGHJKLMNPQRSTUVWXYZ23456789") for _ in range(6))
    mech = str(body.get("mechanic") or "").strip().upper()[:30]
    with db() as c:
        c.execute("INSERT INTO invites(code, label, created_at, max_uses, mechanic) VALUES(?, ?, ?, ?, ?)",
                  (code, str(body.get("label") or "")[:60], time.time(), int(body.get("max_uses") or 1), mech))
    return {"ok": True, "code": code, "mechanic": mech}


@admin.post("/api/admin/invites/revoke")
async def adm_invite_revoke(request: Request, _=Depends(admin_auth)):
    body = await request.json()
    with db() as c:
        c.execute("UPDATE invites SET revoked=1 WHERE code=?", (str(body.get("code") or ""),))
    return {"ok": True}


@admin.get("/api/admin/audit")
def adm_audit(limit: int = 60, _=Depends(admin_auth)):
    with db() as c:
        rows = [dict(r) for r in c.execute("SELECT id, at, device, build, bytes, saved_at, report, version_after FROM audit ORDER BY id DESC LIMIT ?", (min(limit, 400),))]
    for r in rows:
        r["device"] = (r["device"] or "")[:6] + "…" if r["device"] != "admin-import" else r["device"]
        try:
            r["report"] = {k: v for k, v in json.loads(r["report"]).items() if v}
        except Exception:
            pass
    return {"audit": rows}


@admin.get("/api/admin/audit/{aid}")
def adm_audit_one(aid: int, _=Depends(admin_auth)):
    with db() as c:
        r = c.execute("SELECT raw FROM audit WHERE id=?", (aid,)).fetchone()
    if not r:
        raise HTTPException(404)
    return Response(r["raw"], media_type="application/json")


@admin.get("/api/admin/quarantine")
def adm_quarantine(_=Depends(admin_auth)):
    with db() as c:
        rows = [dict(r) for r in c.execute("SELECT id, at, device, build, reason, resolved, length(raw) bytes FROM quarantine ORDER BY id DESC LIMIT 200")]
    for r in rows:
        r["device"] = (r["device"] or "")[:6] + "…"
    return {"quarantine": rows}


@admin.post("/api/admin/quarantine/{qid}/{action}")
def adm_quarantine_act(qid: int, action: str, _=Depends(admin_auth)):
    with db() as c:
        r = c.execute("SELECT * FROM quarantine WHERE id=?", (qid,)).fetchone()
    if not r:
        raise HTTPException(404)
    if action == "apply":
        payload = json.loads(r["raw"])
        payload["app"] = "k1kartlog"; payload["type"] = "snapshot"
        ok, resp = apply_push(payload, r["device"], r["build"])
        if not ok:
            return resp
    elif action != "dismiss":
        raise HTTPException(400, "apply or dismiss")
    with db() as c:
        c.execute("UPDATE quarantine SET resolved=1 WHERE id=?", (qid,))
    return {"ok": True}


@admin.get("/api/admin/settings")
def adm_settings(_=Depends(admin_auth)):
    return {"min_build": setting("min_build", ""), "build": current_build(), "admin_mechanic": setting("admin_mechanic", "ROBERT")}


@admin.post("/api/admin/settings")
async def adm_settings_set(request: Request, _=Depends(admin_auth)):
    body = await request.json()
    if "min_build" in body:
        mb = str(body["min_build"] or "").strip()
        if mb and not re.match(r"^\d{4}-\d{4}$", mb):
            raise HTTPException(400, "build looks like 0921-1430")
        if mb and build_of(mb) > build_of(current_build()):
            raise HTTPException(400, f"the minimum cannot be above the build being served ({current_build()}) - every device would be told to update to something that does not exist")
        set_setting("min_build", mb)
    if "admin_mechanic" in body:
        set_setting("admin_mechanic", str(body["admin_mechanic"] or "").strip().upper()[:30])
    return {"ok": True, "min_build": setting("min_build", ""), "admin_mechanic": setting("admin_mechanic", "ROBERT")}


@admin.get("/api/admin/needed")
def adm_needed(_=Depends(admin_auth)):
    """the 'APP NEEDED' script: parts at or under their reorder point, with the suggested order quantity"""
    snap, _v, _u = load_snapshot()
    if not snap:
        return {"needed": []}
    out = []
    for num, cfg in (snap.get("invCfg") or {}).items():
        if not isinstance(cfg, dict) or num in (snap.get("partTomb") or {}):
            continue
        qty = M._num((snap.get("inv") or {}).get(num, 0))
        r = M._num(cfg.get("r", 0))
        if r > 0 and qty <= r:
            out.append({"part": num, "name": cfg.get("n"), "qty": qty, "reorder": r, "order": max(M._num(cfg.get("g", 0)) - qty, 0)})
    out.sort(key=lambda x: (x["qty"] - x["reorder"], x["part"]))
    return {"needed": out}


@admin.get("/api/admin/export/entries.csv")
def adm_export_entries(_=Depends(admin_auth)):
    snap, _v, _u = load_snapshot()
    lines = ["kart,date,action,parts,mechanic,notes,photos"]
    for k, v in sorted((snap or {}).get("karts", {}).items(), key=lambda kv: M._num(kv[0])):
        for e in v.get("entries") or []:
            row = [k, e.get("date", ""), e.get("action", ""), e.get("parts", ""), e.get("mechanic", ""), e.get("notes", ""), str(len(e.get("photos") or []))]
            lines.append(",".join('"' + str(x).replace('"', '""') + '"' for x in row))
    return PlainTextResponse("\n".join(lines), media_type="text/csv")


# ------------------------------------------------------------------------------------------ ordering
# The sheet's ordering flow, on the server: APP NEEDED (ORDER ✓ / ORDER QTY) -> "Place order" -> APP ORDERS ->
# "Book received quantities into stock". Tables: order_draft = the checks and quantities typed on the needed
# list (kept between visits, shared by every browser); orders + order_lines = APP ORDERS (what was ordered and
# how much of it has been booked so far); receipts = every booking with the stock number before and after.
# Booking adds the delta to the canonical snapshot's inv and stamps invTouched - exactly what the app used to
# do with a sheet receipt on one device - so every device adopts the new count on its next pull.
ORDER_TZ = "America/Chicago"


def _order_now():
    try:
        from zoneinfo import ZoneInfo
        import datetime as _dt
        return _dt.datetime.now(ZoneInfo(ORDER_TZ))
    except Exception:
        import datetime as _dt
        return _dt.datetime.now()


def needed_rows(snap):
    """the sheet's APP NEEDED: parts with levels set that are at/under red (NEEDED) or under green (WANTED)"""
    out = []
    if not snap:
        return out
    inv, cfgs, tomb = M._d(snap.get("inv")), M._d(snap.get("invCfg")), M._d(snap.get("partTomb"))
    for num, cfg in cfgs.items():
        if not isinstance(cfg, dict) or num in tomb:
            continue
        r, g = M._num(cfg.get("r", 0)), M._num(cfg.get("g", 0))
        if r <= 0 and g <= 0:
            continue                                   # no levels set = not something we track
        qty = M._num(inv.get(num, 0))
        status = "NEEDED" if qty <= r else ("WANTED" if qty < g else "")
        if not status:
            continue
        out.append({"part": num, "name": cfg.get("n") or num, "qty": qty, "red": r, "green": g, "status": status,
                    "to_green": max(g - qty, 0), "fits": cfg.get("k") or ""})
    out.sort(key=lambda x: (0 if x["status"] == "NEEDED" else 1, x["qty"] - x["red"], x["part"]))
    return out


def _clean_int(v):
    try:
        n = int(float(v))
    except (TypeError, ValueError):
        return None
    return n if n >= 0 else None


@admin.get("/api/admin/orders/needed")
def adm_orders_needed(_=Depends(admin_auth)):
    snap, version, _u = load_snapshot()
    rows = needed_rows(snap)
    keep = {r["part"] for r in rows}
    with db() as c:
        draft = {r["part"]: dict(r) for r in c.execute("SELECT part, chk, qty FROM order_draft")}
        # like the sheet, a check only lives as long as the part is on the list
        stale = [p for p in draft if p not in keep]
        if stale:
            c.executemany("DELETE FROM order_draft WHERE part=?", [(p,) for p in stale])
    for r in rows:
        d = draft.get(r["part"]) or {}
        r["chk"] = bool(d.get("chk"))
        r["order_qty"] = d.get("qty")
    return {"needed": rows, "version": version}


@admin.post("/api/admin/orders/draft")
async def adm_orders_draft(request: Request, _=Depends(admin_auth)):
    """ORDER ✓ and ORDER QTY for one part, or a list of them ({"rows": [...]})"""
    body = await request.json()
    rows = body.get("rows") if isinstance(body.get("rows"), list) else [body]
    with db() as c:
        for r in rows:
            part = str(r.get("part") or "").strip().upper()
            if not part:
                continue
            qty = _clean_int(r.get("qty")) if r.get("qty") not in (None, "") else None
            c.execute("INSERT INTO order_draft(part, chk, qty, at) VALUES(?, ?, ?, ?) "
                      "ON CONFLICT(part) DO UPDATE SET chk=excluded.chk, qty=excluded.qty, at=excluded.at",
                      (part, 1 if r.get("chk") else 0, qty, time.time()))
    return {"ok": True}


@admin.post("/api/admin/orders/place")
async def adm_orders_place(request: Request, _=Depends(admin_auth)):
    """checked rows with an ORDER QTY become one order; their checks are cleared (the sheet's placeOrder)"""
    body = await request.json()
    note = str(body.get("note") or "").strip()[:120]
    snap, _v, _u = load_snapshot()
    cfgs = M._d((snap or {}).get("invCfg"))
    with _LOCK, db() as c:
        draft = [dict(r) for r in c.execute("SELECT part, chk, qty FROM order_draft WHERE chk=1")]
        lines = [(d["part"], (cfgs.get(d["part"]) or {}).get("n") or d["part"], int(d["qty"])) for d in draft if (d["qty"] or 0) > 0]
        skipped = len(draft) - len(lines)
        if not lines:
            raise HTTPException(400, f"{skipped} checked item(s) have no ORDER QTY - type how many you're ordering first"
                                if skipped else "nothing checked")
        now = _order_now()
        base = "ORD-" + now.strftime("%Y%m%d-%H%M")
        oid, n = base, 2
        while c.execute("SELECT 1 FROM orders WHERE id=?", (oid,)).fetchone():
            oid = f"{base}-{n}"
            n += 1
        c.execute("INSERT INTO orders(id, at, note) VALUES(?, ?, ?)", (oid, time.time(), note))
        c.executemany("INSERT INTO order_lines(order_id, part, name, ordered, booked) VALUES(?, ?, ?, ?, 0)",
                      [(oid, p, nm, q) for p, nm, q in lines])
        c.executemany("DELETE FROM order_draft WHERE part=?", [(p,) for p, _n, _q in lines])
    return {"ok": True, "order": oid, "lines": len(lines), "units": sum(q for _p, _n, q in lines), "skipped": skipped}


def _orders(c, show_all):
    orders = [dict(r) for r in c.execute("SELECT id, at, note FROM orders ORDER BY at DESC")]
    lines = [dict(r) for r in c.execute("SELECT id, order_id, part, name, ordered, booked, last_booked FROM order_lines ORDER BY id")]
    by = {}
    for ln in lines:
        by.setdefault(ln["order_id"], []).append(ln)
    out, closed = [], 0
    for o in orders:
        o["lines"] = by.get(o["id"], [])
        o["ordered"] = sum(ln["ordered"] or 0 for ln in o["lines"])
        o["booked"] = sum(ln["booked"] or 0 for ln in o["lines"])
        o["open"] = any((ln["booked"] or 0) < (ln["ordered"] or 0) for ln in o["lines"])
        if not o["open"]:
            closed += 1
            if not show_all:
                continue
        out.append(o)
    return out, closed


@admin.get("/api/admin/orders")
def adm_orders(all: int = 0, _=Depends(admin_auth)):
    with db() as c:
        out, closed = _orders(c, bool(all))
    return {"orders": out, "closed": closed}


@admin.get("/api/admin/orders/{oid}/text")
def adm_order_text(oid: str, _=Depends(admin_auth)):
    """the order as plain text, to paste into an email or the supplier's form"""
    with db() as c:
        o = c.execute("SELECT id, at, note FROM orders WHERE id=?", (oid,)).fetchone()
        if not o:
            raise HTTPException(404, "no such order")
        lines = c.execute("SELECT part, name, ordered, booked FROM order_lines WHERE order_id=? ORDER BY id", (oid,)).fetchall()
    when = time.strftime("%m/%d/%Y", time.localtime(o["at"]))
    txt = [f"K1 Arlington parts order {o['id']}  ({when})" + (f"  {o['note']}" if o["note"] else ""), ""]
    for ln in lines:
        txt.append(f"{ln['part']:<12} x{ln['ordered']:<5} {ln['name']}")
    txt.append("")
    txt.append(f"{len(lines)} items, {sum(l['ordered'] for l in lines)} units")
    return PlainTextResponse("\n".join(txt))


@admin.post("/api/admin/orders/receive")
async def adm_orders_receive(request: Request, _=Depends(admin_auth)):
    """the sheet's bookReceived: for each line, received (running total) minus booked = what goes into stock now.
    Split shipments: raise the received number when the rest arrives and book again - only the new amount is added."""
    body = await request.json()
    want = {}
    for r in body.get("lines") or []:
        lid, rec = _clean_int(r.get("id")), _clean_int(r.get("received"))
        if lid is not None and rec is not None:
            want[lid] = rec
    if not want:
        raise HTTPException(400, "nothing to book")
    now_ms = int(time.time() * 1000)
    booked, skipped = [], []
    with _LOCK:
        with db() as c:
            rows = {r["id"]: dict(r) for r in c.execute(
                f"SELECT id, order_id, part, name, ordered, booked FROM order_lines WHERE id IN ({','.join('?' * len(want))})", list(want))}
        snap, version, _u = load_snapshot()
        if snap is None:
            raise HTTPException(409, "the server has no data yet")
        tomb = M._d(snap.get("partTomb"))
        for lid, rec in want.items():
            ln = rows.get(lid)
            if not ln:
                skipped.append({"id": lid, "why": "no such line"})
                continue
            delta = rec - (ln["booked"] or 0)
            if delta <= 0:
                skipped.append({"id": lid, "part": ln["part"], "why": "nothing new" if delta == 0 else f"already booked {ln['booked']}, cannot lower it"})
                continue
            before = M._num(snap.setdefault("inv", {}).get(ln["part"], 0))
            after = before + delta
            snap["inv"][ln["part"]] = int(after) if float(after).is_integer() else after
            snap.setdefault("invTouched", {})[ln["part"]] = now_ms
            booked.append({"id": lid, "order": ln["order_id"], "part": ln["part"], "name": ln["name"], "qty": delta, "before": before,
                           "after": after, "received": rec, "deleted_part": ln["part"] in tomb})
        if booked:
            version += 1
            store_snapshot(snap, version, "admin-receive")
            with db() as c:
                for b in booked:
                    c.execute("UPDATE order_lines SET booked=?, last_booked=? WHERE id=?", (b["received"], time.time(), b["id"]))
                    c.execute("INSERT INTO receipts(at, line_id, order_id, part, qty, before, after, version) VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                              (time.time(), b["id"], b["order"], b["part"], b["qty"], b["before"], b["after"], version))
                report = {"inv_updated": len(booked), "received": [{"part": b["part"], "qty": b["qty"], "order": b["order"]} for b in booked]}
                c.execute("INSERT INTO audit(at, device, build, bytes, saved_at, report, version_after, raw) VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                          (time.time(), "admin-receive", "admin", 0, "", json.dumps(report), version, json.dumps({"booked": booked})))
                c.execute("DELETE FROM audit WHERE id NOT IN (SELECT id FROM audit ORDER BY id DESC LIMIT 400)")
    return {"ok": True, "booked": booked, "units": sum(b["qty"] for b in booked), "skipped": skipped,
            "version": version if booked else None}


@admin.post("/api/admin/orders/{oid}/cancel")
def adm_order_cancel(oid: str, _=Depends(admin_auth)):
    """an order placed by mistake can go, but only while nothing from it has been booked into stock"""
    with _LOCK, db() as c:
        if not c.execute("SELECT 1 FROM orders WHERE id=?", (oid,)).fetchone():
            raise HTTPException(404, "no such order")
        got = c.execute("SELECT COALESCE(SUM(booked),0) b FROM order_lines WHERE order_id=?", (oid,)).fetchone()["b"]
        if got:
            raise HTTPException(409, f"{got} unit(s) from this order are already booked into stock - it stays on the record")
        c.execute("DELETE FROM order_lines WHERE order_id=?", (oid,))
        c.execute("DELETE FROM orders WHERE id=?", (oid,))
    return {"ok": True}


@admin.post("/api/admin/parts/{num}/delete")
def adm_part_delete(num: str, _=Depends(admin_auth)):
    """delete a part from the catalog the way the app's DELETE PART does: tombstone it and drop it from every
    section, so every device removes it on its next pull and a stale device cannot bring it back"""
    num = num.strip().upper()
    if not _SAFE.match(num):
        raise HTTPException(400, "bad part number")
    now_ms = int(time.time() * 1000)
    with _LOCK:
        snap, version, _u = load_snapshot()
        if snap is None:
            raise HTTPException(409, "the server has no data yet")
        cfg = M._d(snap.get("invCfg")).get(num)
        if not cfg:
            raise HTTPException(404, f"{num} is not in the catalog")
        gone = {"part": num, "name": cfg.get("n"), "qty": M._num(M._d(snap.get("inv")).get(num, 0))}
        for sec in ("invCfg", "inv", "invTouched", "invCounted", "cfgTouched"):
            M._d(snap.get(sec)).pop(num, None)
        snap.setdefault("partTomb", {})[num] = now_ms
        version += 1
        store_snapshot(snap, version, "admin-edit")
        with db() as c:
            c.execute("DELETE FROM order_draft WHERE part=?", (num,))
            c.execute("INSERT INTO audit(at, device, build, bytes, saved_at, report, version_after, raw) VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                      (time.time(), "admin-edit", "admin", 0, "", json.dumps({"parts_deleted": 1, "deleted": [gone]}), version,
                       json.dumps({"deleted": gone})))
    return {"ok": True, "deleted": gone, "version": version}


@admin.get("/api/admin/receipts")
def adm_receipts(limit: int = 100, _=Depends(admin_auth)):
    with db() as c:
        rows = [dict(r) for r in c.execute("SELECT * FROM receipts ORDER BY id DESC LIMIT ?", (min(limit, 1000),))]
    return {"receipts": rows}


public.include_router(shared)
admin.include_router(shared)
init()
