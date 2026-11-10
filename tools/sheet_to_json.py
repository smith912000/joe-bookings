#!/usr/bin/env python3
"""Turn the bookings CSV (exported from the 'Joe's Bookings' Sheet) into data/busy.json.
Only date+hour keys are published (no client names). Usage: sheet_to_json.py bookings.csv data/busy.json"""
import csv, json, sys, datetime
src, dst = sys.argv[1], sys.argv[2]
busy = set()
for r in csv.DictReader(open(src, newline='', encoding='utf-8')):
    if not r.get('Date') or r.get('Status') == 'Cancelled':
        continue
    a = int(r['Start time'].split(':')[0]); b = int((r.get('End time') or '0').split(':')[0] or 0)
    for h in range(a, max(b, a + 1)):
        busy.add(f"{r['Date']} {h:02d}")
json.dump({"updated": datetime.datetime.now(datetime.timezone.utc).strftime('%d/%m/%Y %H:%M UTC'), "busy": sorted(busy)}, open(dst, 'w'), indent=1)
