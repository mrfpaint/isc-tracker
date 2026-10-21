"""Builds dpl.json (the DPL price list the ISC wizard prices Order Details from) out of the price-list Excel.

    python tools/build_dpl.py "Price List 28 May 26.xlsx" 2026-05-28

Reads the first sheet: Material Description (col D), Rounded Pack Size (col C) and New DPL (col I). Rows with no
numeric DPL ("-") are skipped. Descriptions are matched case-insensitively, so "MC PLUS BASE 1" and
"MC Plus Base 1" become one material; where the same material + pack appears twice the first row wins.
"""
import json, re, sys
from pathlib import Path
import openpyxl

src, as_of = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(src, data_only=True, read_only=True).worksheets[0]

items, names = {}, {}
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r or r[3] is None or not isinstance(r[8], (int, float)) or not isinstance(r[2], (int, float)) or r[2] <= 0:
        continue
    name = re.sub(r'\s+', ' ', str(r[3]).replace('\xa0', ' ')).strip()
    key = name.lower()
    names.setdefault(key, name)
    items.setdefault(names[key], {}).setdefault(float(r[2]), r[8])

out = {'asOf': as_of, 'items': {n: sorted([[p if p % 1 else int(p), d] for p, d in packs.items()])
                                for n, packs in sorted(items.items())}}
dst = Path(__file__).resolve().parent.parent / 'dpl.json'
dst.write_text(json.dumps(out, separators=(',', ':'), ensure_ascii=False), encoding='utf-8')
print(f'{len(out["items"])} materials -> {dst}')
