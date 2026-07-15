#!/usr/bin/env python3
"""Assemble the final single-file app.
Seed data comes from seed_full.json (extracted from the shop's Excel workbook
by extract step — full history for all 53 karts)."""
import json

seed = json.load(open("seed_full.json", encoding="utf-8"))
print("seeded karts:", len(seed), " entries:", sum(len(v["entries"]) for v in seed.values()))

# canonical mechanic chips (entry history keeps whatever spelling was recorded)
MECHANICS = ["ROBERT", "WESLEY", "JOHN"]

inv_seed = json.load(open("inv_seed.json", encoding="utf-8"))
pd = json.load(open("parts_data.json", encoding="utf-8"))
parts, thumbs = pd["parts"], pd["thumbs"]
print("parts:", len(parts), "thumbs:", len(thumbs))

SYNC_URL = "https://script.google.com/macros/s/AKfycbyw5Z3C_4yPn-MbUPLJkxrYPzNYwDkdbf7jQhmAToo7rZPuo3tgUcmWvELmcRpHUel_/exec"

import time as _t
BUILD = _t.strftime("%m%d-%H%M")
tpl = open("app_template.html", encoding="utf-8").read()
out = tpl.replace('/*__BUILD__*/"dev"', json.dumps(BUILD))
out = out.replace("/*__PARTS__*/[]", json.dumps(parts, separators=(",",":")))
out = out.replace("/*__THUMBS__*/[]", json.dumps(thumbs, separators=(",",":")))
out = out.replace("/*__INVSEED__*/{}", json.dumps(inv_seed, separators=(",",":")))
assert '/*__SYNCURL__*/""' in out
out = out.replace('/*__SYNCURL__*/""', json.dumps(SYNC_URL))
out = out.replace("/*__SEED__*/{}", json.dumps(seed, separators=(",",":")))
assert 'mechanics:["ROBERT","WESLEY","JOHN"]' in out
out = out.replace('mechanics:["ROBERT","WESLEY","JOHN"]',
                  'mechanics:' + json.dumps(MECHANICS, separators=(",",":")))
open("kart_log.html", "w", encoding="utf-8").write(out)
print("wrote kart_log.html", len(out), "chars")
