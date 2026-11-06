#!/usr/bin/env python3
"""Coach Kira – Workout an Garmin Connect senden (inoffizieller Weg über garminconnect ≥ 0.3.20).

Die offizielle Personal API ist nur lesend; Workouts in den Kalender legen geht nur über den Login,
den auch die Garmin-Connect-App nutzt. Tokens liegen in ~/.garminconnect (0700/0600), nie im Sheet
oder auf dem Cockpit-Server.

  kira_garmin_workout.py login        einmalig im Terminal: E-Mail, Passwort, ggf. MFA-Code
  kira_garmin_workout.py test         prüft den gespeicherten Login
  kira_garmin_workout.py job DATEI    Auftrag (JSON vom Cockpit) hochladen + einplanen, Ergebnis als JSON auf stdout
  kira_garmin_workout.py dry DATEI    nur umwandeln und Garmin-JSON ausgeben (kein Upload)
"""
from __future__ import annotations

import json
import sys
from getpass import getpass
from pathlib import Path

TOKENS = Path.home() / ".garminconnect"
STATE = Path.home() / "coach-kira-runner" / "garmin_workouts_state.json"

SPORT = {
    "running": {"sportTypeId": 1, "sportTypeKey": "running"},
    "cycling": {"sportTypeId": 2, "sportTypeKey": "cycling"},
}
OTHER = {"sportTypeId": 3, "sportTypeKey": "other"}
STEP = {
    "warmup": (1, "warmup"),
    "cooldown": (2, "cooldown"),
    "active": (3, "interval"),
    "interval": (3, "interval"),
    "recovery": (4, "recovery"),
    "rest": (5, "rest"),
}


def _client(login: bool = False):
    from garminconnect import Garmin  # erst hier – der Runner läuft auch ohne Bibliothek weiter

    if login:
        email = input("Garmin-E-Mail: ").strip()
        pw = getpass("Garmin-Passwort: ")
        g = Garmin(email, pw, prompt_mfa=lambda: input("MFA-Code (falls per Mail/App gesendet): ").strip())
    else:
        g = Garmin()
    TOKENS.mkdir(mode=0o700, exist_ok=True)
    g.login(str(TOKENS))
    return g


def umwandeln(auftrag: dict) -> dict:
    """Cockpit-Workout (wie für die FIT-Datei) → Garmin-Connect-Workout-JSON."""
    w = auftrag["workout"]
    sport = SPORT.get(w.get("sport"), OTHER)
    order = 0
    child = 0
    steps = []
    total = 0.0

    def exe(s: dict, child_id=None) -> dict:
        nonlocal order, total
        order += 1
        sec = max(1.0, round(float(s.get("minutes") or 0) * 60))
        total += sec
        tid, tkey = STEP.get(s.get("intensity") or "active", STEP["active"])
        d = {
            "type": "ExecutableStepDTO",
            "stepOrder": order,
            "stepType": {"stepTypeId": tid, "stepTypeKey": tkey, "displayOrder": tid},
            "endCondition": {"conditionTypeId": 2, "conditionTypeKey": "time", "displayOrder": 2, "displayable": True},
            "endConditionValue": float(sec),
            "description": " · ".join(x for x in [str(s.get("name") or "").strip(), str(s.get("notes") or "").strip()] if x)[:500],
        }
        lo, hi = s.get("hrLow"), s.get("hrHigh")
        if lo and hi:
            d["targetType"] = {"workoutTargetTypeId": 4, "workoutTargetTypeKey": "heart.rate.zone", "displayOrder": 1}
            d["targetValueOne"] = float(min(lo, hi))
            d["targetValueTwo"] = float(max(lo, hi))
        else:
            d["targetType"] = {"workoutTargetTypeId": 1, "workoutTargetTypeKey": "no.target", "displayOrder": 1}
        if child_id is not None:
            d["childStepId"] = child_id
        return d

    for b in w.get("blocks") or []:
        n = max(1, int(b.get("repeat") or 1))
        inner = b.get("steps") or []
        if n == 1:
            steps.extend(exe(s) for s in inner)
            continue
        child += 1
        order += 1
        rep = {
            "type": "RepeatGroupDTO",
            "stepOrder": order,
            "stepType": {"stepTypeId": 6, "stepTypeKey": "repeat", "displayOrder": 6},
            "childStepId": child,
            "numberOfIterations": n,
            "smartRepeat": False,
            "endCondition": {"conditionTypeId": 7, "conditionTypeKey": "iterations", "displayOrder": 7, "displayable": False},
            "endConditionValue": float(n),
            "workoutSteps": [],
        }
        t0 = total
        rep["workoutSteps"] = [exe(s, child) for s in inner]
        total += (total - t0) * (n - 1)
        steps.append(rep)

    datum = auftrag["datum"]
    name = f"Kira {datum[8:10]}.{datum[5:7]}. {w.get('name') or 'Workout'}"[:60]
    return {
        "workoutName": name,
        "description": (str(w.get("beschreibung") or "") + " – von Coach Kira").strip(" –")[:1000],
        "sportType": sport,
        "estimatedDurationInSecs": int(total),
        "workoutSegments": [{"segmentOrder": 1, "sportType": sport, "workoutSteps": steps}],
    }


def _state() -> dict:
    try:
        return json.loads(STATE.read_text())
    except Exception:
        return {}


def job(path: str) -> dict:
    auftrag = json.loads(Path(path).read_text())
    g = _client()
    payload = umwandeln(auftrag)
    st = _state()
    alt = (st.get("by_datum") or {}).get(auftrag["datum"])
    if alt:  # neu gesendet → altes Kira-Workout des Tages ersetzen (entfernt auch den Kalendereintrag)
        try:
            g.delete_workout(alt)
        except Exception:
            pass
    r = g.upload_workout(payload)
    wid = r.get("workoutId") if isinstance(r, dict) else None
    if not wid:
        raise RuntimeError(f"Upload ohne workoutId: {str(r)[:200]}")
    g.schedule_workout(wid, auftrag["datum"])
    st.setdefault("by_datum", {})[auftrag["datum"]] = wid
    STATE.write_text(json.dumps(st))
    return {"ok": True, "workoutId": str(wid), "result": f"„{payload['workoutName']}“ ({round(payload['estimatedDurationInSecs'] / 60)} min)"}


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "test"
    if cmd == "login":
        g = _client(login=True)
        print("Login gespeichert für", g.get_full_name() if hasattr(g, "get_full_name") else "Garmin")
    elif cmd == "test":
        g = _client()
        print("OK – angemeldet als", g.get_full_name() if hasattr(g, "get_full_name") else "?")
    elif cmd == "dry":
        print(json.dumps(umwandeln(json.loads(Path(sys.argv[2]).read_text())), ensure_ascii=False, indent=1))
    elif cmd == "job":
        try:
            print(json.dumps(job(sys.argv[2]), ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"ok": False, "result": f"{type(e).__name__}: {str(e)[:250]}"}, ensure_ascii=False))
            sys.exit(1)
    else:
        print(__doc__)
        sys.exit(2)


if __name__ == "__main__":
    main()
