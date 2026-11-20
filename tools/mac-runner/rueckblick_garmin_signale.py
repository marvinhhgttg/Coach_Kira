import sys, os, json, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kira_extras as kx
gt = kx.garmin_tage(1000)
tage = sorted(gt)
z = collections.Counter(); orange = []; jede = {}
for D in tage[36:]:
    sub = {d: v for d, v in gt.items() if d <= D}
    s, _ = kx.garmin_signale(sub, D)
    for x in s: z[(x["name"], x["schwere"])] += 1
    if s: jede[D] = [f'{x["name"]}:{x["schwere"]}' for x in s]
    if any(x["schwere"] == "orange" for x in s): orange.append((D, [x["name"] for x in s if x["schwere"] == "orange"]))
n = len(tage[36:])
print(f"Tage geprüft: {n} ({tage[36]} bis {tage[-1]})")
for (k, sw), c in sorted(z.items()): print(f"{k:14s} {sw:6s} {c:4d} Tage ({c/n*100:.0f} %)")
print("Orange-Tage:"); [print(" ", d, ", ".join(x)) for d, x in orange]
json.dump({"n": n, "zaehl": {f"{k}|{s}": c for (k, s), c in z.items()}, "orange": orange, "jede": jede}, open("/tmp/ck/g5/rueckblick.json", "w"), ensure_ascii=False, indent=1)
