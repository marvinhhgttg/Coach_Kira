#!/usr/bin/env python3
"""Coach Kira – Garmin-Livewerte (23.11.2026).

Liest alle 10 Minuten über den Garmin-Login (garminconnect, eigene venv, ~/.garminconnect) die aktuellen
Garmin-Werte und meldet sie an das Cockpit (POST /api/garmin/live, Mac-Schlüssel):
  tr      Trainingsbereitschaft (neuester Wert des Tages)
  atl     akute Last, ctl chronische Last, acwr Verhältnis akut/chronisch
  status  Trainingszustand (Garmin-Kürzel)
  es      Ausdauerwert
  erholH  Erholungszeit in Stunden, vo2 VO2max

Nur Anzeige: Uhr, Zifferblätter, iPhone-Widget und Heute-Seite. Die Timeline (Wahrheit) und die
Garmin-Skills bleiben unverändert; das Votum rechnet weiter mit den Morgenwerten.
Ausnahme von der Projektregel „kein Ersatz-Workflow liest Garmin Connect“ – von Marc am 23.11.2026 freigegeben.

  kira_garmin_live.py          Abruf + Meldung (nachts 00:00–04:29 still)
  kira_garmin_live.py --test   nur anzeigen, nichts melden
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path
from urllib.request import Request, urlopen

BASE = Path.home() / "coach-kira-runner"
TOKENS = Path.home() / ".garminconnect"
STATE = BASE / "garmin_live.json"
LOG = BASE / "logs" / "garmin_live.log"
ZIEL = "https://coachkira.pplx.app/port/5000/api/garmin/live"


def log(t: str) -> None:
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {t}\n")


def zugang() -> dict:
    try:
        z = json.loads((BASE / "zugang.json").read_text(encoding="utf-8"))
        return z if isinstance(z, dict) else {}
    except Exception:
        return {}


def zustand() -> dict:
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def suche(o, schluessel: str):
    """Ersten Wert zu einem Schlüssel in verschachtelten Antworten finden."""
    if isinstance(o, dict):
        if schluessel in o and o[schluessel] is not None:
            return o[schluessel]
        for v in o.values():
            r = suche(v, schluessel)
            if r is not None:
                return r
    elif isinstance(o, list):
        for v in o:
            r = suche(v, schluessel)
            if r is not None:
                return r
    return None


def abrufen(g, tag: str) -> dict:
    w: dict = {"datum": tag}
    # Trainingsbereitschaft: Liste von Momentaufnahmen → neueste
    try:
        tr = g.get_training_readiness(tag) or []
        # Diagnose 23.11.: Rohantwort sichern (Zeitpunkte, Kontext, Score je Momentaufnahme)
        try:
            (BASE / "garmin_live_roh.json").write_text(json.dumps(tr, ensure_ascii=False)[:20000])
        except Exception:
            pass
        if isinstance(tr, dict):
            tr = [tr]
        tr = [x for x in tr if isinstance(x, dict) and x.get("score") is not None]
        if tr:
            x = max(tr, key=lambda x: str(x.get("timestampLocal") or x.get("timestamp") or ""))
            w["tr"] = x.get("score")
            w["trZeit"] = str(x.get("timestampLocal") or x.get("timestamp") or "")[11:16]
            w["trStufe"] = x.get("level")
            if x.get("recoveryTime") is not None:
                w["erholH"] = round(float(x["recoveryTime"]) / 60.0, 1)
    except Exception as e:
        log(f"Bereitschaft: {type(e).__name__}: {str(e)[:120]}")
        if "429" in str(e):
            raise
    # Trainingszustand mit akuter/chronischer Last
    try:
        ts = g.get_training_status(tag) or {}
        dto = suche(ts, "acuteTrainingLoadDTO") or {}
        if dto:
            w["atl"] = dto.get("dailyTrainingLoadAcute")
            w["ctl"] = dto.get("dailyTrainingLoadChronic")
            w["acwr"] = dto.get("dailyAcuteChronicWorkloadRatio")
        st = suche(ts, "trainingStatusFeedbackPhrase")
        if st:
            w["status"] = st
        vo2 = suche(ts, "vo2MaxPreciseValue") or suche(ts, "vo2MaxValue")
        if vo2:
            w["vo2"] = vo2
    except Exception as e:
        log(f"Trainingszustand: {type(e).__name__}: {str(e)[:120]}")
        if "429" in str(e):
            raise
    # Ausdauerwert
    try:
        es = g.get_endurance_score(tag) or {}
        v = suche(es, "overallScore")
        if v:
            w["es"] = v
    except Exception as e:
        log(f"Ausdauerwert: {type(e).__name__}: {str(e)[:120]}")
    return {k: v for k, v in w.items() if v is not None}


def melden(w: dict) -> dict:
    h = {"Content-Type": "application/json", "Origin": "https://coachkira.pplx.app", "User-Agent": "coach-kira-runner"}
    z = zugang()
    if z.get("mac_schluessel"):
        h["X-CK-Mac"] = str(z["mac_schluessel"])
    req = Request(ZIEL, data=json.dumps(w).encode("utf-8"), headers=h, method="POST")
    with urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))


def main() -> None:
    test = "--test" in sys.argv
    jetzt = datetime.datetime.now()
    s = zustand()
    if not test:
        if (jetzt.hour, jetzt.minute) < (4, 30):
            return
        if s.get("pause_bis") and time.time() < float(s["pause_bis"]):
            return
    from garminconnect import Garmin

    g = Garmin()
    g.login(str(TOKENS))
    tag = jetzt.date().isoformat()
    try:
        w = abrufen(g, tag)
    except Exception as e:
        s["pause_bis"] = time.time() + 3600
        STATE.write_text(json.dumps(s))
        log(f"Garmin bremst (429) – Pause 60 min: {str(e)[:120]}")
        return
    if test:
        print(json.dumps(w, ensure_ascii=False, indent=1))
        return
    try:
        r = melden(w)
        s.update({"zuletzt": jetzt.isoformat(timespec="seconds"), "werte": w, "antwort": r})
        s.pop("pause_bis", None)
        STATE.write_text(json.dumps(s, ensure_ascii=False))
        log("gemeldet " + " ".join(f"{k}={w[k]}" for k in ("tr", "trZeit", "atl", "ctl", "acwr", "es", "status") if k in w))
    except Exception as e:
        log(f"Meldung fehlgeschlagen: {type(e).__name__}: {str(e)[:150]}")


if __name__ == "__main__":
    main()
