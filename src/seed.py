#!/usr/bin/env python3
"""Assemble the final single-file app.
Seed data comes from seed_full.json (extracted from the shop's Excel workbook
by extract step — full history for all 53 karts)."""
import json

seed = json.load(open("seed_full.json", encoding="utf-8"))
print("seeded karts:", len(seed), " entries:", sum(len(v["entries"]) for v in seed.values()))

# canonical mechanic chips (entry history keeps whatever spelling was recorded)
MECHANICS = ["ROBERT", "WESLEY", "EMMETT", "JOHN"]

parts = []
for line in open("parts_raw.txt", encoding="utf-8"):
    line = line.strip()
    if not line or "|" not in line: continue
    num, desc = line.split("|", 1)
    parts.append([num.strip(), desc.strip()])
print("parts:", len(parts))

tpl = open("app_template.html", encoding="utf-8").read()
sheetjs = open("sheetjs.min.js", encoding="utf-8").read()
out = tpl.replace("/*__PARTS__*/[]", json.dumps(parts, separators=(",",":")))
out = out.replace("/*__SEED__*/{}", json.dumps(seed, separators=(",",":")))
assert 'mechanics:["ROBERT","WESLEY","EMMETT"]' in out
out = out.replace('mechanics:["ROBERT","WESLEY","EMMETT"]',
                  'mechanics:' + json.dumps(MECHANICS, separators=(",",":")))
i = out.find("/*__SHEETJS__*/")
out = out[:i] + sheetjs + out[i+len("/*__SHEETJS__*/"):]
open("kart_log.html", "w", encoding="utf-8").write(out)
print("wrote kart_log.html", len(out), "chars")
