import datetime, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kira_extras as kx
H = "2026-10-25"
def D(k): return (datetime.date.fromisoformat(H) - datetime.timedelta(days=k)).isoformat()
def gt(akt=None, z=None, zstd=5, n=60):
    out = {}
    for k in range(n):
        x = {"bb_aufwachen": 75, "tag": {"stress_avg": 25, "bb_max": 80},
             "schlaf": {"stress": 18, "atem": 12.0, "spo2": 96, "tief_s": 1800, "rem_s": 6000, "wach_s": 900, "dauer_s": 28800, "bedarf_min": 450}}
        if akt: akt(k, x)
        if z and 1 <= k <= 7:
            tot = zstd * 3600 / 7
            x["summe"] = {"zonen_s": [tot * p for p in z], "bewegt_s": tot, "km": 20, "hm": 100}
        out[D(k)] = x
    return out
ok = fail = 0
def t(name, c):
    global ok, fail
    ok += bool(c); fail += (not c); print(("OK  " if c else "FAIL") + " " + name)
def namen(s): return {(x["name"], x["schwere"]) for x in s}
s, l = kx.garmin_signale(gt(), H); t("normal: keine Signale", s == [])
def bb(v): return lambda k, x: x.__setitem__("bb_aufwachen", v) if k <= 2 else None
s, _ = kx.garmin_signale(gt(bb(30)), H); t("BB 30 -> orange", ("Body Battery", "orange") in namen(s))
s, _ = kx.garmin_signale(gt(bb(56)), H); t("BB 56 -> gelb", ("Body Battery", "gelb") in namen(s))
s, _ = kx.garmin_signale(gt(bb(62)), H); t("BB 62 -> nichts", "Body Battery" not in {x["name"] for x in s})
s, _ = kx.garmin_signale(gt(bb(40)), H); t("BB 40 (nicht < 35) -> nur gelb", ("Body Battery", "gelb") in namen(s))
def sl(key, v, tage=2): return lambda k, x: x["schlaf"].__setitem__(key, v) if k <= tage else None
s, _ = kx.garmin_signale(gt(sl("atem", 13.6)), H); t("Atem 13,6 -> orange", ("Atmung", "orange") in namen(s))
s, _ = kx.garmin_signale(gt(sl("atem", 13.1)), H); t("Atem 13,1 -> gelb", ("Atmung", "gelb") in namen(s))
s, _ = kx.garmin_signale(gt(sl("spo2", 93)), H); t("SpO2 93 -> orange", ("SpO2", "orange") in namen(s))
s, _ = kx.garmin_signale(gt(sl("stress", 23, 6)), H); t("Schlafstress 23 -> gelb", ("Schlafstress", "gelb") in namen(s))
def st(v): return lambda k, x: x["tag"].__setitem__("stress_avg", v) if 1 <= k <= 7 else None
s, _ = kx.garmin_signale(gt(st(34)), H); t("Stress tags 34 -> orange", ("Stress", "orange") in namen(s))
s, l = kx.garmin_signale(gt(z=[0.05, 0.45, 0.45, 0.04, 0.01]), H); t("Grauzone 45 % Z3 -> gelb", ("Grauzone", "gelb") in namen(s))
s, _ = kx.garmin_signale(gt(z=[0.05, 0.5, 0.4, 0.04, 0.01], zstd=2), H); t("Grauzone nur 2 h -> nichts", "Grauzone" not in {x["name"] for x in s})
s, _ = kx.garmin_signale(gt(bb(30), n=12), H); t("< 10 Basiswerte -> nichts", s == [])
t("Lage-Schlüssel", all(k in l for k in ("bb_aufwachen", "bb_schnitt", "schlaf_tief_min", "schlafbedarf_diff_min", "atem", "atem_schnitt", "spo2", "stress_gestern", "zonen7_pct", "std7", "km7", "hm7")))
txt = kx._garmin_lage_txt(l); print("   ", txt)
t("Lage-Text", "Body Battery beim Aufwachen 75 (Schnitt 75)" in txt and "Bedarf +30 min" in txt and "Z3 45 %" in txt)
d = {"erstellt": "2026-10-25T06:30", "vorausblick": {"ctl_heute": 900, "ctl_7": 905, "ctl_14": 910, "ctl_28": 920, "trend": "steigt", "basis_28": "Ø Plan", "acwr_max": None, "tage": [], "qualitaetstag": None, "mono_hinweise": [], "modell": None},
     "tourfenster": [], "fruehwarnung": {"stufe": "gelb", "signale": s, "prognose": "x"}, "garmin_lage": l}
b = kx._vorausblick_block(d)
t("Block: Garmin-Lage vor Schluss", "- Garmin-Lage:" in b and b.index("Garmin-Lage") < b.index("Nutze diese Zukunftssicht"))
t("leere Daten -> nichts", kx.garmin_signale({}, H) == ([], {}))
s2, _ = kx.garmin_signale(gt(z=[0.05, 0.6, 0.31, 0.03, 0.01]), H); t("Grauzone 31 % -> nichts", "Grauzone" not in {x["name"] for x in s2})
print(f"{ok}/{ok + fail}")
