import sys, json, datetime
sys.path.insert(0, '/Users/marc/coach-kira-runner')
import kira_garmin as g
out, fehler = [], []
heute = datetime.date.today()
y0 = 2012
for y in range(y0, heute.year + 1):
    von, bis = f"{y}-01-01", min(f"{y}-12-31", heute.isoformat())
    try:
        acts = g.activities(von, bis) or []
    except Exception as e:
        fehler.append(f"{y}: {e}"); continue
    for a in acts:
        t = ((a.get("activityType") or {}).get("typeKey") or "")
        if "ski" in t or "snow" in t or "skitour" in str(a.get("activityName", "")).lower():
            out.append({k: a.get(k) for k in ["startTimeLocal", "activityName", "duration", "movingDuration", "distance", "elevationGain", "averageHR", "maxHR", "activityTrainingLoad", "aerobicTrainingEffect", "anaerobicTrainingEffect", "deviceId", "activityId"]} | {"typ": t})
    out_y = len(acts)
    fehler.append(f"{y}: {out_y} Aktivitäten")
json.dump({"skitouren": out, "protokoll": fehler}, open("/tmp/ck/skitouren2.json", "w"), ensure_ascii=False, indent=1)
print("ok", len(out))
