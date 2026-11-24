#!/usr/bin/env python3
"""Coach Kira – Garmin-Livewerte (23.11.2026).

Liest alle 10 Minuten über den Garmin-Login (garminconnect, eigene venv, ~/.garminconnect) die aktuellen
Garmin-Werte und meldet sie an das Cockpit (POST /api/garmin/live, Mac-Schlüssel):
  tr      Trainingsbereitschaft (neuester Wert des Tages)
  atl     akute Last, ctl chronische Last, acwr Verhältnis akut/chronisch
  status  Trainingszustand (Garmin-Kürzel)
  es      Ausdauerwert
  erholH  Erholungszeit in Stunden, vo2 VO2max
  hrv     Herzratenvariabilität (Nachtschnitt), hrvLo/Hi Unter-/Obergrenze, Status
  schlafScore   Schlafqualität-Score, schlafH Schlafdauer (Stunden)
  rhr     Ruhepuls, bbAufwachen    Body Battery beim Aufwachen

Nur Anzeige: Uhr, Zifferblätter, iPhone-Widget und Heute-Seite. Die Timeline (Wahrheit) und die
Garmin-Skills bleiben unverändert; das Votum rechnet weiter mit den Morgenwerten.
Ausnahme von der Projektregel „kein Ersatz-Workflow liest Garmin Connect“ – von Marc am 23.11.2026 freigegeben.

  kira_garmin_live.py          Abruf + Meldung (nachts 00:00–04:29 still), danach kira_auto.lauf (Automatik)
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
            # 24.11.2026: Morgenwert (erste Momentaufnahme des Tages) für den automatischen Morgeneintrag
            m0 = min(tr, key=lambda x: str(x.get("timestampLocal") or x.get("timestamp") or ""))
            w["trMorgen"] = m0.get("score")
            w["trMorgenZeit"] = str(m0.get("timestampLocal") or m0.get("timestamp") or "")[11:16]
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
    # HRV (Herzfrequenzvariabilität)
    try:
        h = g.get_hrv_data(tag) or {}
        w["hrv"] = suche(h, "lastNightAvg")
        if isinstance(h.get("hrvSummary"), dict):
            bas = h["hrvSummary"].get("baseline") or {}
            bl = bas.get("balancedLow")
            bu = bas.get("balancedUpper")
        else:
            bl = bu = None
        w["hrvLo"] = suche(h, "balancedLow") if bl is None else bl
        w["hrvHi"] = suche(h, "balancedUpper") if bu is None else bu
        w["hrvStatus"] = h.get("hrvSummary", {}).get("status") if isinstance(h.get("hrvSummary"), dict) else None
        # 24.11.2026: 7-Tage-Schnitt (Formularspalte hrv_status laut Skill)
        w["hrv7"] = h.get("hrvSummary", {}).get("weeklyAvg") if isinstance(h.get("hrvSummary"), dict) else None
        w["hrv"] = w.get("hrv") if w.get("hrv") is not None else None
    except Exception as e:
        log(f"HRV: {type(e).__name__}: {str(e)[:120]}")
        if "429" in str(e):
            raise
    # Schlaf
    try:
        s = g.get_sleep_data(tag) or {}
        d = s.get("dailySleepDTO") or {}
        ss = d.get("sleepScores") or {}
        ov = ss.get("overall") or {}
        v = ov.get("value")
        if v is not None:
            w["schlafScore"] = v
        st = d.get("sleepTimeSeconds")
        if st is not None:
            w["schlafH"] = round(st / 3600, 2)
    except Exception as e:
        log(f"Schlaf: {type(e).__name__}: {str(e)[:120]}")
        if "429" in str(e):
            raise
    # Ruhepuls
    try:
        u = g.get_user_summary(tag) or {}
        rhr = u.get("restingHeartRate")
        if rhr is not None:
            w["rhr"] = rhr
        bb = u.get("bodyBatteryAtWakeTime")
        if bb is not None:
            w["bbAufwachen"] = bb
    except Exception as e:
        log(f"Ruhepuls: {type(e).__name__}: {str(e)[:120]}")
        if "429" in str(e):
            raise
    # Einheiten von heute (für Kiras Fazit nach dem Speichern und den Tagesstand im Datenfeld)
    try:
        akt = g.get_activities_by_date(tag, tag) or []
        w["akt"] = [{
            "id": x.get("activityId"), "start": str(x.get("startTimeLocal") or "")[:19],
            "name": x.get("activityName") or "", "typ": (x.get("activityType") or {}).get("typeKey") or "",
            "min": round(float(x.get("duration") or 0) / 60.0), "last": round(float(x.get("activityTrainingLoad") or 0)),
            "teAe": x.get("aerobicTrainingEffect"), "teAn": x.get("anaerobicTrainingEffect"),
        } for x in akt[:8] if isinstance(x, dict)]
    except Exception as e:
        log(f"Einheiten: {type(e).__name__}: {str(e)[:120]}")
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
        # 24.11.2026: Takt 5 min (LaunchAgent 300 s) nur morgens 5–10 Uhr, bis 2 h nach einer neuen Einheit und 21–22 Uhr,
        # sonst höchstens alle 10 min – schont Garmins Abrufgrenzen
        schnell = 5 <= jetzt.hour < 10 or 21 <= jetzt.hour < 22 or time.time() - float(s.get("akt_neu_at") or 0) < 2 * 3600
        if not schnell and time.time() - float(s.get("abruf_at") or 0) < 9.5 * 60:
            return
    from garminconnect import Garmin

    g = Garmin()
    g.login(str(TOKENS))
    tag = jetzt.date().isoformat()
    s["abruf_at"] = time.time()
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
        log("gemeldet " + " ".join(f"{k}={w[k]}" for k in ("tr", "trZeit", "atl", "ctl", "acwr", "es", "status") if k in w) + f" akt={len(w.get('akt') or [])}")
    except Exception as e:
        log(f"Meldung fehlgeschlagen: {type(e).__name__}: {str(e)[:150]}")
    # neue Einheit → 2 h schneller Takt
    ids = ",".join(sorted(str(a.get("id")) for a in (w.get("akt") or [])))
    if ids and ids != s.get("akt_ids"):
        s["akt_neu_at"] = time.time()
    s["akt_ids"] = ids
    try:
        STATE.write_text(json.dumps(s, ensure_ascii=False))
    except Exception:
        pass
    # 24.11.2026: Automatik (Formular, Sicherung, Hinweise, Workout, Wächter) – darf den Live-Abruf nie stören
    try:
        sys.path.insert(0, str(BASE))
        import kira_auto

        kira_auto.lauf(g, tag, w)
    except Exception as e:
        log(f"Automatik: {type(e).__name__}: {str(e)[:150]}")


if __name__ == "__main__":
    main()
