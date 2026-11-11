import sys, json, datetime
sys.path.insert(0, '/Users/marc/coach-kira-runner')
import kira_garmin as g
out, prot = [], []
heute = datetime.date.today()
for y in range(2013, heute.year + 1):
    von, bis = f"{y}-01-01", min(f"{y}-12-31", heute.isoformat())
    try:
        acts = g.activities(von, bis) or []
    except Exception as e:
        prot.append(f"{y}: {e}"); continue
    n = 0
    for a in acts:
        t = ((a.get("activityType") or {}).get("typeKey") or "")
        if t in ("hiking", "walking", "mountaineering", "trail_hiking", "speed_walking", "casual_walking", "snow_shoe_ws") or "hik" in t or "mountain" in t:
            n += 1
            out.append({k: a.get(k) for k in ["startTimeLocal", "activityName", "duration", "movingDuration", "distance", "elevationGain", "averageHR", "maxHR", "activityTrainingLoad", "aerobicTrainingEffect", "anaerobicTrainingEffect", "averageSpeed", "steps"]} | {"typ": t})
    prot.append(f"{y}: {len(acts)} Aktivitäten, {n} Wandern/Berg")
json.dump({"touren": out, "protokoll": prot}, open("/tmp/ck/wandern.json", "w"), ensure_ascii=False, indent=1)
print("ok", len(out))
