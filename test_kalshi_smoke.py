"""Quick smoke test: fetch 2 open events and build a mini workbook."""
import sys
sys.path.insert(0, r"C:\Users\jarre\polymarket-film")

import requests, json
from fetch_kalshi_all_categories import (
    fetch_series_map, _get, _extract_markets_from_event,
    _historical_market_to_row, build_workbook, COLUMNS
)

# Grab a small series map sample
print("Loading series map (limited)...")
series_map = {}
data = _get("/series", {"limit": 50})
for s in data.get("series") or []:
    t = s.get("ticker") or ""
    series_map[t] = {"category": s.get("category",""), "title": s.get("title","")}
print(f"  {len(series_map)} series")

# Fetch 3 open events
print("Fetching 3 open events...")
ev_data = _get("/events", {"status": "open", "with_nested_markets": "true", "limit": 3})
events = ev_data.get("events") or []
print(f"  got {len(events)} events")

all_rows = []
for ev in events:
    rows = _extract_markets_from_event(ev, series_map)
    all_rows.extend(rows)
    print(f"  ev={ev.get('event_ticker')} cat={ev.get('category')} -> {len(rows)} markets")
    if rows:
        r = rows[0]
        print(f"    category={r['Category']} subcat={r['Subcategory']}")
        print(f"    vol={r['Market Volume (cts)']} oi={r['Open Interest (cts)']} liq={r['Liquidity ($)']}")
        print(f"    outcome1%={r['Outcome 1 %']} last={r['Last Trade Price']} spread={r['Spread ($)']}")
        print(f"    start={r['Event Start']} end={r['Event End']}")

# Fetch 2 historical
print("\nFetching 2 historical markets...")
h_data = _get("/historical/markets", {"limit": 2})
for m in h_data.get("markets") or []:
    row = _historical_market_to_row(m, series_map)
    print(f"  ticker={row['Market Ticker']} cat={row['Category']} vol={row['Market Volume (cts)']}")
    all_rows.append(row)

# Build mini workbook
print(f"\nBuilding workbook with {len(all_rows)} rows...")
wb = build_workbook(all_rows)
out = r"C:\Users\jarre\polymarket-film\kalshi_smoke_test.xlsx"
wb.save(out)
print(f"Saved: {out}")
print("SMOKE TEST PASSED")
