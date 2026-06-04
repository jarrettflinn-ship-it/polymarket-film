import requests, json
base = "https://external-api.kalshi.com/trade-api/v2"
h = {"Accept": "application/json"}

# Test series
r = requests.get(f"{base}/series", params={"limit": 3}, headers=h, timeout=15)
print("series status:", r.status_code)
data = r.json()
keys = list(data.keys())
print("series keys:", keys)
if "series" in data:
    s = data["series"]
    print(f"  got {len(s)} series items")
    if s:
        print("  first series:", json.dumps(s[0], indent=2)[:500])

# Test events open
r2 = requests.get(f"{base}/events", params={"status": "open", "with_nested_markets": "true", "limit": 2}, headers=h, timeout=15)
print("\nevents status:", r2.status_code)
d2 = r2.json()
print("events keys:", list(d2.keys()))
evs = d2.get("events") or []
print(f"  got {len(evs)} events")
if evs:
    ev = evs[0]
    print("  ev keys:", list(ev.keys()))
    mkts = ev.get("markets") or []
    if mkts:
        print("  market keys:", list(mkts[0].keys()))
        print("  volume_fp:", mkts[0].get("volume_fp"))
        print("  yes_bid:", mkts[0].get("yes_bid"))
        print("  last_price:", mkts[0].get("last_price"))

# Test historical
r3 = requests.get(f"{base}/historical/markets", params={"limit": 2}, headers=h, timeout=15)
print("\nhistorical status:", r3.status_code)
d3 = r3.json()
print("historical keys:", list(d3.keys()))
mhist = d3.get("markets") or []
print(f"  got {len(mhist)} historical markets")
if mhist:
    print("  hist mkt keys:", list(mhist[0].keys()))
    print("  event_ticker:", mhist[0].get("event_ticker"))
    print("  series_ticker:", mhist[0].get("series_ticker"))
    print("  volume_fp:", mhist[0].get("volume_fp"))
