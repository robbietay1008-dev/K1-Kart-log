"""The kart log's merge, server side.

The app keeps its whole world as one JSON snapshot (karts, shop log, quick actions, inventory, part config,
batteries, tombstones, stamps...). Every device merges the server's copy into its own before it pushes, using
the rules in the app's mergeRemote(). This module is the same rules in Python, so the server's canonical copy
is what any device would have computed: a union of entries minus tombstones, and "newest stamp wins" for
everything that has a stamp. Nothing is ever replaced wholesale, so a stale device can only ADD what it has
that is genuinely new - it cannot blank a section it does not know about (the failure the sheet receiver had).

merge(canonical, incoming, now_ms) -> (canonical, report)   report lists what changed, for the audit log.
"""
import copy
import re
import time

SECTIONS = ("karts", "shop", "quicks", "inv", "invCfg", "tomb", "stamps", "invTouched", "invCounted",
            "cfgTouched", "partTomb", "rekeys", "bat", "rc", "photos")


def empty_snapshot():
    return {"karts": {}, "shop": [], "quicks": [], "inv": {}, "invCfg": {}, "tomb": {},
            "stamps": {"karts": {}, "quicks": 0}, "invTouched": {}, "invCounted": {}, "cfgTouched": {},
            "partTomb": {}, "rekeys": [], "bat": {}, "rc": {}, "photos": {}, "parts": None}


def _d(x):
    return x if isinstance(x, dict) else {}


def _l(x):
    return x if isinstance(x, list) else []


def _num(x, default=0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return default


# ---- parts strings: "59921 x1, 59926 x2, 60290" ----------------------------------------------------------
_PART_RE = re.compile(r"^\s*([A-Za-z0-9._/-]+)\s*(?:[xX]\s*(\d+))?\s*$")


def parse_parts(s):
    out = []
    for tok in str(s or "").split(","):
        m = _PART_RE.match(tok)
        if not m:
            tok = tok.strip()
            if tok:
                out.append({"num": tok, "qty": 1})
            continue
        out.append({"num": m.group(1), "qty": int(m.group(2) or 1)})
    return out


def _swap_parts(s, frm, to):
    items = parse_parts(s)
    touched = False
    out = []
    for it in items:
        n = it["num"]
        if n.upper() == frm:
            n = to
            touched = True
        out.append(f"{n} x{it['qty']}" if it["qty"] > 1 else n)
    return ", ".join(out) if touched else None


def rekey_part(db, frm, to, stamp):
    """rekeyPart(): a part number renamed everywhere - log entries, shop entries, inventory, config; the old
    number is tombstoned so it cannot come back from another device"""
    frm = str(frm).upper()
    to = str(to).upper()
    if not frm or not to or frm == to:
        return False
    for k in db["karts"].values():
        for e in _l(k.get("entries")):
            if e.get("parts"):
                ns = _swap_parts(e["parts"], frm, to)
                if ns is not None:
                    e["parts"] = ns
    for se in db["shop"]:
        if se.get("parts"):
            ns = _swap_parts(se["parts"], frm, to)
            if ns is not None:
                se["parts"] = ns
    old = db["invCfg"].get(frm)
    if old:
        cfg = {"n": old.get("n"), "r": old.get("r"), "g": old.get("g")}
        if old.get("k"):
            cfg["k"] = old["k"]
        if old.get("t") is not None:
            cfg["t"] = old["t"]
        if to not in db["invCfg"]:
            db["invCfg"][to] = cfg
        del db["invCfg"][frm]
    if frm in db["inv"]:
        if to not in db["inv"]:
            db["inv"][to] = db["inv"][frm]
        else:
            db["inv"][to] = _num(db["inv"][to]) + _num(db["inv"][frm])
        del db["inv"][frm]
    if db["invCounted"].get(frm):
        db["invCounted"][to] = 1
        del db["invCounted"][frm]
    db["invTouched"].pop(frm, None)
    db["cfgTouched"].pop(frm, None)
    db["invTouched"][to] = stamp
    db["cfgTouched"][to] = stamp
    db["partTomb"][frm] = stamp
    db["partTomb"].pop(to, None)
    return True


def _clean_entry(en):
    photos = en.get("photos")
    if not isinstance(photos, list):
        photos = []
    return {"id": en.get("id"), "date": en.get("date") or "", "action": en.get("action") or "",
            "parts": en.get("parts") or "", "mechanic": en.get("mechanic") or "", "notes": en.get("notes") or "",
            "photos": list(photos)}


def _clamp_stamp(v, now_ms, report, what):
    """a device with a wrong clock must not win forever: stamps more than a day in the future become 'now'"""
    v = _num(v)
    if v > now_ms + 86_400_000:
        report.setdefault("future_stamps", []).append(what)
        return float(now_ms)
    return v


def merge(canonical, incoming, now_ms=None):
    """the app's mergeRemote() with the roles swapped: canonical = the server (local), incoming = a device (remote)"""
    now_ms = now_ms or int(time.time() * 1000)
    first = not canonical
    db = copy.deepcopy(canonical) if canonical else empty_snapshot()
    if first and isinstance(incoming, dict):
        # the very first import: part config and counts that were never "touched" on a device (the seeded
        # catalog) have no stamps and would never merge - the first snapshot seeds them as the base
        for num, cfg in _d(incoming.get("invCfg")).items():
            if isinstance(cfg, dict) and num not in _d(incoming.get("partTomb")):
                db["invCfg"][num] = copy.deepcopy(cfg)
        for num, qty in _d(incoming.get("inv")).items():
            if num not in _d(incoming.get("partTomb")):
                db["inv"][num] = qty
        # every kart's status is stamped at import: an unstamped kart could otherwise be overwritten by any
        # stale device carrying the smallest stamp (a stamp beats no stamp in the app's rule)
        in_stamps = _d(_d(incoming.get("stamps")).get("karts"))
        for k, v in _d(incoming.get("karts")).items():
            if isinstance(v, dict) and k not in in_stamps:
                db["stamps"]["karts"][k] = float(now_ms)
    for key, blank in empty_snapshot().items():
        if key not in db or db[key] is None:
            db[key] = copy.deepcopy(blank)
    db["stamps"].setdefault("karts", {})
    r = incoming or {}
    report = {"entries_added": 0, "entries_removed": 0, "karts_status": 0, "shop_added": 0, "quicks": False,
              "rekeys": 0, "parts_deleted": 0, "cfg_updated": 0, "inv_updated": 0, "bat_added": 0, "bat_updated": 0,
              "bat_history": 0, "rc_added": 0, "photos_added": 0, "tombs_added": 0}

    # tombstones: a union
    for tid, at in _d(r.get("tomb")).items():
        if tid not in db["tomb"]:
            db["tomb"][tid] = at
            report["tombs_added"] += 1

    # karts: entries union minus tombstones; status + notes by the kart's stamp
    rstamps = _d(r.get("stamps"))
    rk_stamps = _d(rstamps.get("karts"))
    rk = _d(r.get("karts"))
    for k, rem in rk.items():
        if not isinstance(rem, dict):
            continue
        loc = db["karts"].setdefault(k, {"status": {}, "entries": [], "knotes": ""})
        loc.setdefault("entries", [])
        have = {e.get("id") for e in loc["entries"] if e.get("id")}
        for en in _l(rem.get("entries")):
            if not isinstance(en, dict) or not en.get("id") or en["id"] in have or en["id"] in db["tomb"]:
                continue
            loc["entries"].append(_clean_entry(en))
            have.add(en["id"])
            report["entries_added"] += 1
            if len(report.setdefault("added", [])) < 40:
                report["added"].append({"kart": k, "date": en.get("date") or "", "action": (en.get("action") or "")[:80],
                                        "mechanic": en.get("mechanic") or "", "parts": (en.get("parts") or "")[:80]})
        rs = _clamp_stamp(rk_stamps.get(k, 0), now_ms, report, f"kart {k}")
        ls = _num(db["stamps"]["karts"].get(k, 0))
        if rs > ls:
            if rem.get("status"):
                loc["status"] = rem["status"]
            loc["knotes"] = rem.get("knotes") or ""
            db["stamps"]["karts"][k] = rs
            report["karts_status"] += 1
    for k, loc in db["karts"].items():
        before = len(loc.get("entries") or [])
        loc["entries"] = [e for e in loc.get("entries") or [] if not (e.get("id") and e["id"] in db["tomb"])]
        report["entries_removed"] += before - len(loc["entries"])

    # shop (non-kart parts use): union minus tombstones
    have_s = {s.get("id") for s in db["shop"] if s.get("id")}
    for se in _l(r.get("shop")):
        if not isinstance(se, dict) or not se.get("id") or se["id"] in have_s or se["id"] in db["tomb"]:
            continue
        db["shop"].append({"id": se["id"], "date": se.get("date") or "", "usedFor": se.get("usedFor") or "",
                           "parts": se.get("parts") or "", "mechanic": se.get("mechanic") or ""})
        have_s.add(se["id"])
        report["shop_added"] += 1
        if len(report.setdefault("added", [])) < 40:
            report["added"].append({"kart": "shop", "date": se.get("date") or "", "action": (se.get("usedFor") or "")[:80],
                                    "mechanic": se.get("mechanic") or "", "parts": (se.get("parts") or "")[:80]})
        if len(report.setdefault("added", [])) < 40:
            report["added"].append({"kart": "shop", "date": se.get("date") or "", "action": (se.get("usedFor") or "")[:80],
                                    "mechanic": se.get("mechanic") or "", "parts": (se.get("parts") or "")[:80]})
    db["shop"] = [s for s in db["shop"] if not (s.get("id") and s["id"] in db["tomb"])]

    # quick actions: the newer list wins
    rq = _clamp_stamp(rstamps.get("quicks", 0), now_ms, report, "quicks")
    if _l(r.get("quicks")) and rq > _num(db["stamps"].get("quicks", 0)):
        db["quicks"] = r["quicks"]
        db["stamps"]["quicks"] = rq
        report["quicks"] = True

    # part identity: renames first (in order), then deletions, then names/thresholds
    mine = {f"{x.get('from')}>{x.get('to')}@{x.get('at')}" for x in db["rekeys"]}
    todo = [x for x in _l(r.get("rekeys")) if isinstance(x, dict) and x.get("from") and x.get("to")
            and f"{x.get('from')}>{x.get('to')}@{x.get('at')}" not in mine]
    todo.sort(key=lambda x: _num(x.get("at")))
    for x in todo:
        rekey_part(db, x["from"], x["to"], _num(x.get("at")) or now_ms)
        db["rekeys"].append({"from": str(x["from"]).upper(), "to": str(x["to"]).upper(), "at": x.get("at") or 0})
        report["rekeys"] += 1
    if len(db["rekeys"]) > 100:
        db["rekeys"] = db["rekeys"][-100:]
    for dnum, at in _d(r.get("partTomb")).items():
        if dnum in db["partTomb"]:
            continue
        if _num(db["cfgTouched"].get(dnum, 0)) > _num(at):
            continue          # edited here after the delete there: the part is wanted back
        db["partTomb"][dnum] = at
        for sec in ("invCfg", "inv", "invTouched", "invCounted", "cfgTouched"):
            db[sec].pop(dnum, None)
        report["parts_deleted"] += 1
    rct = _d(r.get("cfgTouched"))
    rcfg = _d(r.get("invCfg"))
    for cnum, at in rct.items():
        if cnum in db["partTomb"] or not isinstance(rcfg.get(cnum), dict):
            continue
        at = _clamp_stamp(at, now_ms, report, f"cfg {cnum}")
        if at < _num(db["cfgTouched"].get(cnum, 0)):
            continue
        rc2 = rcfg[cnum]
        lc2 = db["invCfg"].get(cnum)
        if not lc2:
            new = {"n": rc2.get("n"), "r": rc2.get("r"), "g": rc2.get("g")}
            if rc2.get("k"):
                new["k"] = rc2["k"]
            if rc2.get("t") is not None:
                new["t"] = rc2["t"]
            db["invCfg"][cnum] = new
            report["cfg_updated"] += 1
        elif (lc2.get("n") != rc2.get("n") or lc2.get("r") != rc2.get("r") or lc2.get("g") != rc2.get("g")
              or (lc2.get("k") or "") != (rc2.get("k") or "")):
            lc2["n"], lc2["r"], lc2["g"] = rc2.get("n"), rc2.get("r"), rc2.get("g")
            if rc2.get("k"):
                lc2["k"] = rc2["k"]
            else:
                lc2.pop("k", None)
            if rc2.get("t") is not None:
                lc2["t"] = rc2["t"]
            report["cfg_updated"] += 1
        db["cfgTouched"][cnum] = at

    # inventory per part: newest touch wins; the recount checkmark travels with the count
    rit = _d(r.get("invTouched"))
    rin = _d(r.get("inv"))
    ric = _d(r.get("invCounted"))
    for num, at in rit.items():
        if num in db["partTomb"]:
            continue
        at = _clamp_stamp(at, now_ms, report, f"inv {num}")
        if num in rin and at > _num(db["invTouched"].get(num, 0)):
            db["inv"][num] = rin[num]
            db["invTouched"][num] = at
            if ric.get(num):
                db["invCounted"][num] = 1
            else:
                db["invCounted"].pop(num, None)
            report["inv_updated"] += 1

    # batteries: per entry, newest edit wins; a serial scanned on two devices collapses onto the first scan
    rb = _d(r.get("bat"))
    by_sn = {b.get("sn"): bid for bid, b in db["bat"].items() if isinstance(b, dict) and b.get("sn")}
    for bid, rbe in rb.items():
        if not isinstance(rbe, dict) or not rbe.get("sn") or bid in db["tomb"]:
            continue
        lbe = db["bat"].get(bid)
        if not lbe:
            dup = by_sn.get(rbe["sn"])
            if dup and dup != bid:
                keep = db["bat"][dup]
                if _num(rbe.get("c")) < _num(keep.get("c")):
                    # theirs is the earlier scan: it owns the row, ours becomes the duplicate
                    if not rbe.get("kart") and keep.get("kart"):
                        rbe["kart"] = keep["kart"]
                    if not rbe.get("ini") and keep.get("ini"):
                        rbe["ini"] = keep["ini"]
                    db["tomb"][dup] = now_ms
                    del db["bat"][dup]
                    db["bat"][bid] = _bat_row(rbe, keep)
                    by_sn[rbe["sn"]] = bid
                else:
                    changed = False
                    if not keep.get("kart") and rbe.get("kart"):
                        keep["kart"] = rbe["kart"]; changed = True
                    if not keep.get("ini") and rbe.get("ini"):
                        keep["ini"] = rbe["ini"]; changed = True
                    if changed:
                        keep["at"] = now_ms
                    db["tomb"][bid] = now_ms
                report["bat_updated"] += 1
                continue
            db["bat"][bid] = _bat_row(rbe, None)
            by_sn[rbe["sn"]] = bid
            report["bat_added"] += 1
        elif _num(rbe.get("at")) > _num(lbe.get("at")):
            for f in ("sn", "rcv", "bd", "st", "kart", "pos", "date", "ini"):
                lbe[f] = rbe.get(f) or ""
            if rbe.get("nf"):
                lbe["nf"] = 1
            else:
                lbe.pop("nf", None)
            lbe["at"] = rbe.get("at")
            if rbe.get("c") is not None:
                lbe["c"] = rbe.get("c")
            report["bat_updated"] += 1
        # history: a union keyed by timestamp
        cur = db["bat"].get(bid)
        rh = _l(rbe.get("h"))
        if cur is not None and rh:
            lh = cur.setdefault("h", [])
            seen = {h.get("t") for h in lh}
            for h in rh:
                if isinstance(h, dict) and h.get("t") not in seen:
                    lh.append(h)
                    seen.add(h.get("t"))
                    report["bat_history"] += 1
            lh.sort(key=lambda h: _num(h.get("t")))
    for bd in list(db["bat"].keys()):
        if bd in db["tomb"]:
            del db["bat"][bd]

    # receipts applied + photo links: unions
    for rid in _d(r.get("rc")):
        if rid not in db["rc"]:
            db["rc"][rid] = 1
            report["rc_added"] += 1
    for pid, url in _d(r.get("photos")).items():
        if pid not in db["photos"]:
            db["photos"][pid] = url
            report["photos_added"] += 1
    if r.get("parts"):
        db["parts"] = r["parts"]
    return db, report


def _bat_row(rbe, keep):
    keep = keep or {}
    row = {"sn": rbe.get("sn"), "rcv": rbe.get("rcv") or keep.get("rcv") or "", "bd": rbe.get("bd") or keep.get("bd") or "",
           "st": rbe.get("st") or keep.get("st") or "", "kart": rbe.get("kart") or "", "pos": rbe.get("pos") or "",
           "date": rbe.get("date") or "", "ini": rbe.get("ini") or "", "c": rbe.get("c") or 0, "at": rbe.get("at") or 0}
    if rbe.get("nf"):
        row["nf"] = 1
    if isinstance(rbe.get("h"), list):
        row["h"] = list(rbe["h"])
    return row


def counts(snap):
    """the numbers an admin wants at a glance"""
    karts = _d(snap.get("karts"))
    return {"karts": len(karts), "entries": sum(len(_l(k.get("entries"))) for k in karts.values() if isinstance(k, dict)),
            "shop": len(_l(snap.get("shop"))), "parts": len(_d(snap.get("invCfg"))), "batteries": len(_d(snap.get("bat"))),
            "tombstones": len(_d(snap.get("tomb"))), "photos": len(_d(snap.get("photos")))}


def sanity(incoming):
    """reasons a snapshot must not be merged at all (it is quarantined instead)"""
    problems = []
    if not isinstance(incoming, dict):
        return ["not an object"]
    if incoming.get("app") != "k1kartlog":
        problems.append("not a kart log snapshot")
    if incoming.get("type") not in (None, "snapshot"):
        problems.append(f"type {incoming.get('type')!r} is not a snapshot")
    karts = incoming.get("karts")
    if karts is not None and not isinstance(karts, dict):
        problems.append("karts is not a map")
    if isinstance(karts, dict):
        for k, v in list(karts.items())[:200]:
            if not isinstance(v, dict):
                problems.append(f"kart {k} is not an object")
                break
            if v.get("entries") is not None and not isinstance(v.get("entries"), list):
                problems.append(f"kart {k} entries is not a list")
                break
    return problems
