#!/usr/bin/env python3
"""Extract full kart history from the shop Excel workbook (karts log .xlsx)
into seed_full.json for seed.py. Usage: python3 extract_from_xlsx.py workbook.xlsx"""
import openpyxl, json, re, datetime, sys

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)

def norm(h): return re.sub(r'\s+', ' ', str(h)).strip().upper() if h is not None else ""
def fmt_date(v):
    if isinstance(v, datetime.datetime): return f"{v.month}/{v.day}/{v.year}"
    return str(v).strip() if v is not None else ""
def fmt_bat(v):
    if isinstance(v, datetime.datetime): return f"{v.month}-{str(v.year)[2:]}"
    s = str(v).strip() if v is not None else ""
    return "" if s in ("?", "N/A", "x", "SD") else s
def fmt_svc(v):
    if isinstance(v, datetime.datetime): return f"{v.month}/{v.day}/{v.year}"
    s = str(v).strip() if v is not None else ""
    return "" if s in ("?", "N/A") else s

HDRMAP = {"DATE":"date","ACTION COMPLETED":"action","PARTS USED":"parts","MECHANIC":"mech",
 "NOTES":"notes","CHAIN DATE":"chain","DIFF DATE":"diff","BRAKE FLUSH":"brake",
 "BAT 1 DATE":"bat1","BATE 1 DATE":"bat1","BAT 2 DATE":"bat2","BATE 2 DATE":"bat2",
 "BAT 3 DATE":"bat3","BAT 4 DATE":"bat4","WELD COUNT":"weldcount"}

seed = {}
for k in [str(i) for i in range(1, 54)]:
    ws = wb[k]
    rows = list(ws.iter_rows(values_only=True)) or [()]
    cols = {}
    for idx, h in enumerate(rows[0]):
        key = HDRMAP.get(norm(h))
        if key and key not in cols: cols[key] = idx
    def get(row, key):
        i = cols.get(key)
        return row[i] if i is not None and i < len(row) else None
    sr = rows[1] if len(rows) > 1 else ()
    wc = get(sr, "weldcount")
    try: wc = int(wc) if wc is not None and str(wc).strip() != "" else 0
    except Exception: wc = 0
    status = {"chain":fmt_svc(get(sr,"chain")),"diff":fmt_svc(get(sr,"diff")),
      "brake":fmt_svc(get(sr,"brake")),"bat1":fmt_bat(get(sr,"bat1")),
      "bat2":fmt_bat(get(sr,"bat2")),"bat3":fmt_bat(get(sr,"bat3")),
      "bat4":fmt_bat(get(sr,"bat4")),"weld":wc}
    kn = get(sr, "notes")
    entries = []
    for row in rows[2:]:
        d, a = get(row,"date"), get(row,"action")
        if d is None and a is None: continue
        m = get(row,"mech"); p = get(row,"parts"); n = get(row,"notes")
        entries.append({"date":fmt_date(d),
            "action":str(a).strip() if a is not None else "",
            "parts":str(p).strip() if p is not None else "",
            "mechanic":str(m).strip().upper() if m is not None else "",
            "notes":str(n).strip() if n is not None else ""})
    seed[k] = {"status":status,"entries":entries,
               "knotes":str(kn).strip() if kn is not None else ""}

json.dump(seed, open("seed_full.json","w"), separators=(",",":"))
print("karts:", len(seed), "entries:", sum(len(v["entries"]) for v in seed.values()))
