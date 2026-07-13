#!/usr/bin/env python3
"""Build parts_data.json for the app:
- names from Robbie's 'inventory' tab (his shop names) with factory
  descriptions (parts_raw.txt / 'full' tab) kept as hidden search aliases
- part photos from the inventory tab, shrunk to small thumbnails
Format: {"parts": [[num, name, alias, thumbIdx], ...], "thumbs": [dataURL, ...]}
"""
import openpyxl, json, io, base64
from PIL import Image

wb = openpyxl.load_workbook('current.xlsx', data_only=True)
ws = wb['inventory']

# factory descriptions (aliases)
catalog = {}
order = []
for line in open('parts_raw.txt', encoding='utf-8'):
    line = line.strip()
    if not line or '|' not in line: continue
    n, d = line.split('|', 1)
    n = n.strip()
    catalog[n.upper()] = d.strip()
    order.append(n)

# inventory names by row (openpyxl rows are 1-based; anchors 0-based)
inv = {}      # rownum(0-based) -> (num, name)
for r, row in enumerate(ws.iter_rows(min_row=1, values_only=True)):
    num, name = row[1], row[2]
    if r == 0 or num is None: continue
    num_s = str(num).strip()
    name_s = str(name).strip() if name else ''
    if num_s and name_s:
        inv[r] = (num_s, name_s)

# thumbnails by anchor row
thumbs = []
thumb_for_num = {}
for im in getattr(ws, '_images', []):
    try:
        r = im.anchor._from.row
        if r not in inv: continue
        num = inv[r][0].upper()
        if num in thumb_for_num: continue
        data = im._data()
        img = Image.open(io.BytesIO(data)).convert('RGB')
        img.thumbnail((96, 96))
        buf = io.BytesIO()
        img.save(buf, 'JPEG', quality=55, optimize=True)
        thumb_for_num[num] = len(thumbs)
        thumbs.append('data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode())
    except Exception:
        continue

parts = []
seen = set()
# inventory-named parts first (most-used, Robbie's names)
for r in sorted(inv):
    num, name = inv[r]
    key = num.upper()
    if key in seen: continue
    seen.add(key)
    alias = catalog.get(key, '')
    if alias.upper() == name.upper(): alias = ''
    parts.append([num, name, alias, thumb_for_num.get(key, -1)])
# remaining catalog-only parts
for n in order:
    key = n.upper()
    if key in seen: continue
    seen.add(key)
    parts.append([n, catalog[key], '', -1])

total_thumb_bytes = sum(len(t) for t in thumbs)
json.dump({'parts': parts, 'thumbs': thumbs}, open('parts_data.json', 'w'))
print('parts:', len(parts), '| with photos:', len(thumb_for_num),
      '| thumb payload:', round(total_thumb_bytes/1024), 'KB')
print('sample:', parts[:3])
