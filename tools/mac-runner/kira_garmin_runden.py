#!/usr/bin/env python3
"""Coach Kira – Runden und Wetter einzelner Aktivitäten über den Garmin-Login (garminconnect, eigene venv).

Die offizielle Personal API liefert keine Runden (404). Dieses Skript nutzt denselben Login wie
„An Garmin senden“ (~/.garminconnect) und schreibt Ergebnisse in den Cache von kira_garmin.py.

  kira_garmin_runden.py AUFTRAG.json CACHE.json
  AUFTRAG = [{"id": 123, "typ": "running", "dauer_s": 4200}, ...]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

TOKENS = Path.home() / ".garminconnect"
ENT_TYPEN = {"running", "road_biking", "cycling", "gravel_cycling", "mountain_biking", "indoor_cycling", "virtual_ride"}
INDOOR = {"indoor_cycling", "virtual_ride", "treadmill_running", "indoor_rowing"}


def entkopplung_aus_runden(laps: list, rad: bool) -> dict | None:
    """Aerobe Entkopplung: Effizienz (Tempo bzw. Leistung je Herzschlag) 1. vs. 2. Hälfte.
    Erste 10 Minuten zählen nicht. Positiv = Puls driftet bei gleichem Tempo nach oben."""
    rs = []
    for l in laps or []:
        d = float(l.get("movingDuration") or l.get("duration") or 0)
        hr = l.get("averageHR")
        if d < 30 or not hr:
            continue
        rs.append({"d": d, "hr": float(hr), "km": float(l.get("distance") or 0) / 1000.0, "w": l.get("averagePower")})
    if len(rs) < 4:
        return None
    t, rest = 0.0, []
    for r in rs:
        if t + r["d"] / 2 >= 600:
            rest.append(r)
        t += r["d"]
    if len(rest) < 4:
        return None
    gesamt = sum(r["d"] for r in rest)
    h1, h2, t = [], [], 0.0
    for r in rest:
        (h1 if t + r["d"] / 2 < gesamt / 2 else h2).append(r)
        t += r["d"]
    if len(h1) < 2 or len(h2) < 2:
        return None
    mit_watt = rad and all(r["w"] for r in rest)

    def eff(h):
        dd = sum(r["d"] for r in h)
        hr = sum(r["hr"] * r["d"] for r in h) / dd
        x = (sum(float(r["w"]) * r["d"] for r in h) / dd) if mit_watt else (sum(r["km"] for r in h) / (dd / 3600.0))
        return hr, x

    hr1, x1 = eff(h1)
    hr2, x2 = eff(h2)
    if hr1 <= 0 or hr2 <= 0 or x1 <= 0:
        return None
    e1, e2 = x1 / hr1, x2 / hr2
    return {"pct": round((e1 - e2) / e1 * 100, 1), "basis": "watt" if mit_watt else "tempo",
            "hf1": round(hr1), "hf2": round(hr2), "x1": round(x1, 1), "x2": round(x2, 1), "runden": len(rest)}


def entkopplung_aus_verlauf(det: dict, rad: bool) -> dict | None:
    """Wie entkopplung_aus_runden, aber aus dem Sekundenverlauf (für Einheiten ohne Auto-Runden).
    Stillstand (< 1 m/s bzw. 0 W) zählt nicht, die ersten 10 Minuten auch nicht."""
    desc = {d.get("key"): d.get("metricsIndex") for d in (det or {}).get("metricDescriptors") or []}
    rows = [r.get("metrics") or [] for r in (det or {}).get("activityDetailMetrics") or []]
    iT, iH, iV, iW = desc.get("sumDuration"), desc.get("directHeartRate"), desc.get("directSpeed"), desc.get("directPower")
    if iT is None or iH is None or iV is None or len(rows) < 100:
        return None
    pts = []
    for m in rows:
        try:
            t, h, v = m[iT], m[iH], m[iV]
        except IndexError:
            continue
        if t is None or not h or v is None or v < 1.0:
            continue
        w = m[iW] if iW is not None and iW < len(m) else None
        pts.append((float(t), float(h), float(v), float(w) if w is not None else None))
    pts = [p for p in pts if p[0] >= 600]
    if len(pts) < 60:
        return None
    mit_watt = rad and sum(1 for p in pts if p[3]) >= 0.8 * len(pts)
    if mit_watt:
        pts = [p for p in pts if p[3] is not None]
    mitte = pts[len(pts) // 2][0]
    h1 = [p for p in pts if p[0] < mitte]
    h2 = [p for p in pts if p[0] >= mitte]

    def eff(h):
        hr = sum(p[1] for p in h) / len(h)
        x = (sum(p[3] for p in h) / len(h)) if mit_watt else (sum(p[2] for p in h) / len(h) * 3.6)
        return hr, x

    hr1, x1 = eff(h1)
    hr2, x2 = eff(h2)
    if hr1 <= 0 or hr2 <= 0 or x1 <= 0:
        return None
    e1, e2 = x1 / hr1, x2 / hr2
    return {"pct": round((e1 - e2) / e1 * 100, 1), "basis": "watt" if mit_watt else "tempo",
            "hf1": round(hr1), "hf2": round(hr2), "x1": round(x1, 1), "x2": round(x2, 1), "runden": 0, "quelle": "verlauf"}


def wetter_aus(w: dict | None) -> dict | None:
    """Garmin-Aktivitätswetter (Fahrenheit, mph) → °C, %, km/h."""
    if not w or w.get("temp") is None:
        return None
    c = round((float(w["temp"]) - 32) * 5 / 9, 1)
    if not -40 <= c <= 50:
        return None
    out = {"temp_c": c}
    if w.get("apparentTemp") is not None:
        out["gefuehlt_c"] = round((float(w["apparentTemp"]) - 32) * 5 / 9, 1)
    if w.get("relativeHumidity") is not None:
        out["feuchte"] = int(w["relativeHumidity"])
    if w.get("windSpeed") is not None:
        out["wind_kmh"] = round(float(w["windSpeed"]) * 1.609, 1)
    return out


def main() -> None:
    auftrag = json.loads(Path(sys.argv[1]).read_text())
    cache_p = Path(sys.argv[2])
    try:
        cache = json.loads(cache_p.read_text())
    except Exception:
        cache = {}
    from garminconnect import Garmin

    g = Garmin()
    g.login(str(TOKENS))
    heute = datetime.date.today()
    ent_n = wet_n = fehl = 0
    for i, a in enumerate(auftrag):
        aid, typ = int(a["id"]), a.get("typ") or ""
        e = {"v": 3, "ok": True}
        try:
            if typ in ENT_TYPEN and (a.get("dauer_s") or 0) >= 3600:
                sp = g.get_activity_splits(aid) or {}
                ent = entkopplung_aus_runden(sp.get("lapDTOs") or [], typ != "running")
                if not ent:  # keine/zu wenige Runden → Sekundenverlauf
                    ent = entkopplung_aus_verlauf(g.get_activity_details(aid, maxchart=4000, maxpoly=0), typ != "running")
                if ent:
                    e["ent"] = ent
                    ent_n += 1
            if typ not in INDOOR:
                try:
                    w = wetter_aus(g.get_activity_weather(aid))
                    if w:
                        e["wetter"] = w
                        wet_n += 1
                except Exception:
                    pass
            if "ent" not in e and "wetter" not in e:
                e = {"v": 3, "ok": False, "bis": (heute + datetime.timedelta(days=60)).isoformat(), "grund": "keine Runden/kein Wetter"}
        except Exception as ex:
            fehl += 1
            e = {"v": 3, "ok": False, "bis": (heute + datetime.timedelta(days=3)).isoformat(), "grund": f"{type(ex).__name__}: {str(ex)[:120]}"}
            if "429" in str(ex) or "Too Many" in str(ex):
                cache[str(aid)] = e
                break
        cache[str(aid)] = e
        if i % 10 == 9:
            cache_p.write_text(json.dumps(cache))
        time.sleep(0.6)
    cache_p.write_text(json.dumps(cache))
    print(f"{len(auftrag)} Einheiten: {ent_n} mit Entkopplung, {wet_n} mit Wetter, {fehl} Fehler")


if __name__ == "__main__":
    main()
