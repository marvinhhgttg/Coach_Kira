#!/usr/bin/env python3
"""Coach Kira – Automatik auf dem Mac (24.11.2026, von Marc freigegeben: „Formular-Regel wird gelockert“).

Wird von kira_garmin_live.py nach jeder Live-Meldung aufgerufen (gleicher Garmin-Login, eigene venv):
  formular     Morgens (≥ 15 min nach Garmins Aufwachwert), nach jeder gespeicherten Einheit, abends ab 23:30 (bis 24.11. 21:30)
               → POST /api/auto/eintrag (Server prüft gegen die Timeline, trägt über onFormSubmit ein)
  sicherung    Server-Zustand (Kopplungen, Token, Einstellungen) lokal sichern; nach einem Neustart wiederherstellen
  hinweise     fällige Hinweise ans iPhone (Erinnerungen-App, iCloud) und als Mac-Mitteilung
  workout      ab 20 Uhr das morgige Plan-Workout in die Garmin-Warteschlange (kira_runner lädt es hoch)
  waechter     Selbstüberwachung – meldet sich nur bei Fehlern (je Fehler höchstens einmal am Tag)

Die Garmin-Skills bleiben unverändert und von Hand nutzbar. Schalter im Cockpit: Technik → Automatik.
  kira_auto.py --test     nur anzeigen, was eingetragen würde (kein Senden)
"""
from __future__ import annotations

import datetime
import json
import subprocess
import sys
import time
from pathlib import Path
from urllib.request import Request, urlopen

BASE = Path.home() / "coach-kira-runner"
STATE = BASE / "auto_state.json"
SICHERUNG = BASE / "server_sicherung.json"
LOG = BASE / "logs" / "auto.log"
SERVER = "https://coachkira.pplx.app/port/5000"
VERSION = "auto-2026-11-24"

ZUSTAND = {
    "PRODUCTIVE": "Formaufbau", "MAINTAINING": "Formerhalt", "PEAKING": "Höchstform", "RECOVERY": "Erholung",
    "UNPRODUCTIVE": "Unproduktiv", "DETRAINING": "Formverlust", "OVERREACHING": "Überlastung", "STRAINED": "Ermüdet",
}
SPORT = [
    ("Run", ("running", "trail_running", "treadmill_running", "track_running", "indoor_running", "ultra_run", "virtual_run")),
    ("Bike", ("cycling", "road_biking", "gravel_cycling", "mountain_biking", "indoor_cycling", "virtual_ride", "e_bike_fitness", "e_bike_mountain", "cyclocross", "bmx", "recumbent_cycling")),
    ("Hike", ("hiking", "mountaineering")),
    ("Row", ("rowing", "indoor_rowing")),
    ("Swim", ("lap_swimming", "open_water_swimming", "swimming")),
    ("Skike", ("inline_skating", "skike")),
    ("Ski", ("backcountry_skiing", "backcountry_snowboarding", "skate_skiing", "cross_country_skiing", "cross_country_skiing_ws", "resort_skiing", "resort_skiing_snowboarding_ws", "skiing")),
]


def log(t: str) -> None:
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {t}\n")


def lesen(p: Path, d):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return d


def schreiben(p: Path, v) -> None:
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(v, ensure_ascii=False), encoding="utf-8")
    tmp.chmod(0o600)
    tmp.replace(p)


def schluessel() -> str:
    z = lesen(BASE / "zugang.json", {})
    return str(z.get("mac_schluessel") or "") if isinstance(z, dict) else ""


def api(methode: str, pfad: str, body=None, timeout: int = 60) -> dict:
    h = {"Content-Type": "application/json", "Origin": "https://coachkira.pplx.app", "User-Agent": "coach-kira-runner", "X-CK-Mac": schluessel()}
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = Request(SERVER + pfad, data=data, headers=h, method=methode)
    try:
        with urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:  # HTTPError trägt oft JSON
        try:
            return json.loads(e.read().decode("utf-8"))  # type: ignore[attr-defined]
        except Exception:
            return {"ok": False, "error": f"{type(e).__name__}: {str(e)[:120]}"}


def sport_von(typ: str) -> str | None:
    t = (typ or "").lower()
    for name, keys in SPORT:
        if t in keys:
            return name
    for name, keys in SPORT:
        if any(k in t for k in keys):
            return name
    return None


def zone_von(sek: dict[int, float]) -> str | None:
    s = {z: v for z, v in sek.items() if 1 <= z <= 5 and v > 0}
    if not s:
        return None
    reihe = sorted(s.items(), key=lambda x: -x[1])
    z1, t1 = reihe[0]
    if len(reihe) > 1:
        z2, t2 = reihe[1]
        if abs(z2 - z1) == 1 and t2 >= 0.5 * t1:
            a, b = sorted((z1, z2))
            return f"Z{a}, Z{b}"
    return f"Z{z1}"


def nachricht(titel: str, text: str, iphone: bool = True) -> None:
    """Mac-Mitteilung und (für iPhone/Uhr) eine Erinnerung in der Liste „Coach Kira“ (iCloud)."""
    esc = lambda s: s.replace("\\", "\\\\").replace('"', '\\"')
    try:
        subprocess.run(["osascript", "-e", f'display notification "{esc(text)}" with title "{esc(titel)}"'], timeout=10)
    except Exception:
        pass
    if not iphone:
        return
    pause = BASE / "erinnerung_pause"
    try:
        if pause.exists() and time.time() < float(pause.read_text() or 0):
            return
    except Exception:
        pass
    # 25.11.2026: Aufgaben nur mit Datum (ganztägig), ohne Uhrzeit
    skript = (
        'tell application "Reminders"\n'
        '  if (name of every account) contains "iCloud" then\n'
        '    if not (exists list "Coach Kira" of account "iCloud") then make new list at account "iCloud" with properties {name:"Coach Kira"}\n'
        '    set l to list "Coach Kira" of account "iCloud"\n'
        '  else\n'
        '    if not (exists list "Coach Kira") then make new list with properties {name:"Coach Kira"}\n'
        '    set l to list "Coach Kira"\n'
        '  end if\n'
        '  set alt to (current date) - (6 * hours)\n'
        '  try\n'
        '    delete (every reminder of l whose creation date < alt)\n'
        '  end try\n'
        f'  make new reminder at end of l with properties {{name:"{esc(titel)}", body:"{esc(text)}", allday due date:(current date)}}\n'
        'end tell'
    )
    try:
        r = subprocess.run(["osascript", "-e", skript], capture_output=True, text=True, timeout=25)
        if r.returncode != 0:
            log(f"Erinnerung fehlgeschlagen: {r.stderr.strip()[:160]}")
            pause.write_text(str(time.time() + 3600))
    except Exception as e:
        log(f"Erinnerung fehlgeschlagen (Freigabe für Erinnerungen fehlt?): {str(e)[:80]}")
        pause.write_text(str(time.time() + 3600))


# ---------------------------------------------------------------- Formular
def morgens(g, tag: str, w: dict, st: dict, jetzt: datetime.datetime, test: bool) -> None:
    if st.get("morgens") == tag or jetzt.hour < 5 or jetzt.hour >= 14:
        return
    if not all(w.get(k) is not None for k in ("tr", "schlafH", "rhr")):
        return
    tz = str(w.get("trMorgenZeit") or w.get("trZeit") or "")
    if len(tz) == 5:
        t0 = jetzt.replace(hour=int(tz[:2]), minute=int(tz[3:]), second=0, microsecond=0)
        if jetzt < t0 + datetime.timedelta(minutes=15):
            return
    v: dict = {
        "sleep_hours": w.get("schlafH"), "sleep_score_0_100": w.get("schlafScore"), "rhr_bpm": w.get("rhr"),
        "Garmin_Training_Readiness": w.get("trMorgen", w.get("tr")),
    }
    if w.get("hrv7") is not None:
        v["hrv_status"] = w["hrv7"]
    if w.get("hrvLo") is not None and w.get("hrvHi") is not None:
        v["hrv_threshholds"] = f"{int(round(float(w['hrvLo'])))};{int(round(float(w['hrvHi'])))}"
    zs = ZUSTAND.get(str(w.get("status") or "").split("_")[0])
    if zs:
        v["Trainingszustand"] = zs
    # ATL/CTL/ACWR nur vor der ersten Einheit (sonst sind es schon die Werte nach der Aktivität)
    if not [a for a in (w.get("akt") or []) if (a.get("last") or 0) > 0]:
        v.update({"garminATL": w.get("atl"), "garminCTL": w.get("ctl"), "garminACWR": w.get("acwr")})
        st["atl_morgen"] = w.get("atl")
    v = {k: x for k, x in v.items() if x is not None}
    if test:
        print("Morgens", json.dumps(v, ensure_ascii=False))
        return
    r = api("POST", "/api/auto/eintrag", {"typ": "Morgens", "datum": tag, "werte": v})
    log(f"Morgens {tag}: {json.dumps(r, ensure_ascii=False)[:200]}")
    if r.get("ok"):
        st["morgens"] = tag
    else:
        st.setdefault("fehler", []).append(f"Morgens: {r.get('error')}")


def aktivitaet(g, tag: str, w: dict, st: dict, jetzt: datetime.datetime, test: bool) -> None:
    akt = [a for a in (w.get("akt") or []) if (a.get("last") or 0) > 0 and (a.get("min") or 0) >= 5]
    if not akt:
        return
    ids = ",".join(sorted(str(a.get("id")) for a in akt))
    a_st = st.get("akt") or {}
    if a_st.get("tag") != tag:
        a_st = {"tag": tag}
    if a_st.get("ids") != ids:
        a_st = {"tag": tag, "ids": ids, "gesehen": time.time(), "n": 0}
    st["akt"] = a_st
    if a_st.get("n", 0) >= 2:
        return
    alter = time.time() - float(a_st.get("gesehen") or time.time())
    if a_st.get("n", 0) == 0:
        # Garmin rechnet Last/ATL nach dem Speichern nach – 15 min warten, oder bis sich ATL bewegt hat
        if alter < 15 * 60 and not (st.get("atl_morgen") and w.get("atl") and w["atl"] != st["atl_morgen"] and alter >= 5 * 60):
            return
    else:
        # zweiter Eintrag 60 min später, falls Garmin ATL/CTL noch nachgezogen hat
        if alter < 60 * 60 or w.get("atl") == a_st.get("atl"):
            return
    haupt = max(akt, key=lambda a: a.get("last") or 0)
    sek: dict[int, float] = {}
    for a in akt:
        try:
            for z in g.get_activity_hr_in_timezones(a["id"]) or []:
                n = int(z.get("zoneNumber") or 0)
                sek[n] = sek.get(n, 0.0) + float(z.get("secsInZone") or 0)
        except Exception as e:
            log(f"HF-Zonen {a.get('id')}: {type(e).__name__}: {str(e)[:100]}")
    v: dict = {
        "load_fb_day": round(sum(a.get("last") or 0 for a in akt)),
        "fbATL_obs": w.get("atl"), "fbCTL_obs": w.get("ctl"), "fbACWR_obs": w.get("acwr"),
        "garminEnduranceScore": w.get("es"), "fb_TR_obs": w.get("tr"),
        "Aerobic_TE": haupt.get("teAe"), "Anaerobic_TE": haupt.get("teAn"),
    }
    sp = sport_von(haupt.get("typ") or "")
    if sp:
        v["Sport_x"] = sp
    zo = zone_von(sek)
    if zo:
        v["Zone"] = zo
    v = {k: x for k, x in v.items() if x is not None}
    if test:
        print("Nach Aktivität", json.dumps(v, ensure_ascii=False))
        return
    r = api("POST", "/api/auto/eintrag", {"typ": "Nach Aktivität", "datum": tag, "werte": v})
    log(f"Nach Aktivität {tag} ({ids}): {json.dumps(r, ensure_ascii=False)[:200]}")
    if r.get("ok"):
        a_st["n"] = a_st.get("n", 0) + 1
        a_st["gesehen"] = time.time()
        a_st["atl"] = w.get("atl")
    else:
        st.setdefault("fehler", []).append(f"Nach Aktivität: {r.get('error')}")


KAL = ("calories", "totalcalories", "caloriesconsumed", "energy")


def _kal(o: dict) -> float:
    for k, v in o.items():
        if k.lower() in KAL and isinstance(v, (int, float)):
            return float(v)
    return 0.0


def _suche_summen(o, tiefe: int = 0, treffer=None):
    """Tagessummen in der Garmin-Ernährungsantwort: Objekt mit Kalorien + Protein, größte Kalorienzahl gewinnt."""
    if treffer is None:
        treffer = []
    if tiefe <= 7:
        if isinstance(o, dict):
            ks = [k.lower() for k in o.keys()]
            if any(k in KAL for k in ks) and any(k.startswith("protein") for k in ks) and _kal(o) > 0:
                treffer.append(o)
            for v in o.values():
                _suche_summen(v, tiefe + 1, treffer)
        elif isinstance(o, list):
            for v in o:
                _suche_summen(v, tiefe + 1, treffer)
    return max(treffer, key=_kal) if (tiefe == 0 and treffer) else (None if tiefe == 0 else treffer)


def ernaehrung(g, tag: str) -> dict | None:
    roh = {}
    for name in ("get_nutrition_daily_food_log", "get_nutrition_daily_meals"):
        try:
            roh[name] = getattr(g, name)(tag)
        except Exception as e:
            roh[name] = {"fehler": f"{type(e).__name__}: {str(e)[:120]}"}
    try:
        (BASE / "garmin_ernaehrung_roh.json").write_text(json.dumps(roh, ensure_ascii=False)[:60000])
    except Exception:
        pass
    fl = roh.get("get_nutrition_daily_food_log")
    s = fl.get("dailyNutritionContent") if isinstance(fl, dict) and isinstance(fl.get("dailyNutritionContent"), dict) else None
    s = s or _suche_summen(fl) or _suche_summen(roh.get("get_nutrition_daily_meals"))
    if not s:
        return None
    lk = {k.lower(): v for k, v in s.items()}
    kal = next((lk[k] for k in ("calories", "totalcalories", "caloriesconsumed", "energy") if k in lk), None)
    f = lambda *ns: next((lk[k] for k in lk for n in ns if k.startswith(n) and isinstance(lk[k], (int, float))), None)
    return {"kcal_in": kal, "carb_g": f("carbohydrate", "carbs", "carb"), "protein_g": f("protein"), "fat_g": f("fat", "totalfat"), "fiber_g": f("fiber", "dietaryfiber")}


def abends(g, tag: str, st: dict, jetzt: datetime.datetime, test: bool) -> None:
    ziel = None
    if (jetzt.hour, jetzt.minute) >= (23, 30) and st.get("abends") != tag:   # 24.11.2026: Marc – abends 23:30
        ziel = tag
    gestern = (jetzt.date() - datetime.timedelta(days=1)).isoformat()
    if 5 <= jetzt.hour < 12 and st.get("abends") != gestern and st.get("abends_nachtrag") != gestern:
        ziel = gestern  # Nachtrag für gestern, falls 23:30 verpasst (Mac schlief)
    if not ziel:
        return
    e = ernaehrung(g, ziel)
    if not e or not e.get("kcal_in"):
        log(f"Abends {ziel}: keine Ernährungssummen bei Garmin gefunden (Rohdaten in garmin_ernaehrung_roh.json)")
        if ziel == gestern:
            st["abends_nachtrag"] = gestern
        return
    out = None
    try:
        u = g.get_user_summary(ziel) or {}
        out = u.get("totalKilocalories")
    except Exception as ex:
        log(f"Abends kcal_out: {ex}")
    v = {k: round(float(x)) for k, x in e.items() if x is not None}
    if out:
        v["kcal_out"] = round(float(out))
        v["deficit"] = v["kcal_out"] - v["kcal_in"]
    if test:
        print("Abends", ziel, json.dumps(v, ensure_ascii=False))
        return
    r = api("POST", "/api/auto/eintrag", {"typ": "Abends", "datum": ziel, "werte": v})
    log(f"Abends {ziel}: {json.dumps(r, ensure_ascii=False)[:200]}")
    if r.get("ok"):
        st["abends"] = ziel
    else:
        if ziel == gestern:
            st["abends_nachtrag"] = gestern
        st.setdefault("fehler", []).append(f"Abends: {r.get('error')}")


# ---------------------------------------------------------------- Sicherung
def sicherung(st: dict) -> dict:
    r = api("GET", f"/api/auto/sicherung?v={VERSION}", timeout=30)
    if not r.get("ok"):
        return r
    werte = r.get("werte") or {}
    alt = lesen(SICHERUNG, {})
    alt_werte = alt.get("werte") or {}
    neustart = alt.get("start") and r.get("start") and alt.get("start") != r.get("start")
    if neustart:
        fehlt = {k: v for k, v in alt_werte.items() if k not in werte or (k == "widget_geraete" and r.get("geraete", 0) == 0)}
        if fehlt:
            w = api("POST", "/api/auto/wiederherstellen", {"werte": fehlt})
            log(f"Server neu gestartet – wiederhergestellt: {w.get('wiederhergestellt')}")
            r2 = api("GET", f"/api/auto/sicherung?v={VERSION}", timeout=30)
            if r2.get("ok"):
                werte = r2.get("werte") or werte
                r = r2
    # Apps-Script-Token fehlt (Server frisch) → aus zugang.json nachreichen (Server setzt nur fehlende Schlüssel)
    if "dashboard_token" not in werte:
        z = lesen(BASE / "zugang.json", {})
        tok = str(z.get("apps_script_token") or "") if isinstance(z, dict) else ""
        if tok:
            w2 = api("POST", "/api/auto/wiederherstellen", {"werte": {"dashboard_token": tok}})
            log(f"Apps-Script-Token nachgereicht: {w2.get('wiederhergestellt')}")
    # Sicherung nie mit einem leeren Stand überschreiben
    if werte.get("widget_geraete") or not alt_werte.get("widget_geraete"):
        schreiben(SICHERUNG, {"at": time.time(), "start": r.get("start"), "werte": {**alt_werte, **werte} if neustart else werte})
    else:
        schreiben(SICHERUNG, {**alt, "start": r.get("start")})
    return r


# ---------------------------------------------------------------- Hinweise, Workout
def hinweise(st: dict, test: bool) -> None:
    r = api("GET", "/api/auto/hinweise", timeout=120)
    if not r.get("ok"):
        return
    gesendet = list(st.get("hinweise") or [])[-80:]  # ids: Datum-…, plan-…, ok-…, wo-…
    for h in r.get("hinweise") or []:
        if h["id"] in gesendet:
            continue
        if test:
            print("Hinweis", h)
            continue
        nachricht(h.get("titel") or "Kira", h.get("text") or "", iphone=True)
        gesendet.append(h["id"])
        log(f"Hinweis gesendet: {h['id']}")
    st["hinweise"] = gesendet


def workout(st: dict, jetzt: datetime.datetime, test: bool) -> None:
    tag = jetzt.date().isoformat()
    if jetzt.hour < 20 or st.get("workout") == tag or test:
        return
    r = api("POST", "/api/auto/workout-morgen", {}, timeout=120)
    log(f"Workout morgen: {json.dumps(r, ensure_ascii=False)[:200]}")
    if r.get("ok"):
        st["workout"] = tag
    else:
        st.setdefault("fehler", []).append(f"Workout: {r.get('error')}")


# ---------------------------------------------------------------- Wächter
def waechter(st: dict, w: dict | None, server_ok: bool, jetzt: datetime.datetime, test: bool) -> None:
    tag = jetzt.date().isoformat()
    z = st.get("waechter") or {}
    z["server_fehl"] = 0 if server_ok else int(z.get("server_fehl", 0)) + 1
    z["garmin_fehl"] = 0 if w else int(z.get("garmin_fehl", 0)) + 1
    probleme = []
    if z["server_fehl"] >= 3:
        probleme.append(("server", "Cockpit-Server seit 30 min nicht erreichbar."))
    if z["garmin_fehl"] >= 3:
        probleme.append(("garmin", "Garmin-Livewerte seit 30 min nicht lesbar (Login/Token prüfen: kira_garmin_workout.py test)."))
    if (jetzt.hour, jetzt.minute) >= (10, 30) and st.get("morgens") != tag and w and w.get("tr") is not None:
        probleme.append(("morgens", "Morgenwerte sind um 10:30 noch nicht im Sheet."))
    try:
        with urlopen("http://127.0.0.1:11434/api/tags", timeout=5) as r:
            r.read()
    except Exception:
        probleme.append(("ollama", "Lokales Modell (Ollama) antwortet nicht."))
    for f in st.pop("fehler", []) or []:
        probleme.append(("eintrag", str(f)[:140]))
    gemeldet = z.get("gemeldet") or {}
    gemeldet = {k: v for k, v in gemeldet.items() if v == tag}
    for k, text in probleme:
        if gemeldet.get(k) == tag:
            continue
        if test:
            print("Wächter", k, text)
            continue
        if 7 <= jetzt.hour < 22:
            nachricht("Kira-Wächter", text, iphone=True)
            gemeldet[k] = tag
            log(f"Wächter: {text}")
    z["gemeldet"] = gemeldet
    st["waechter"] = z


def lauf(g, tag: str, w: dict | None, test: bool = False) -> None:
    jetzt = datetime.datetime.now()
    st = lesen(STATE, {})
    server_ok = False
    try:
        server_ok = bool(sicherung(st).get("ok")) if not test else True
    except Exception as e:
        log(f"Sicherung: {e}")
    if w and g is not None:
        for f in (morgens, aktivitaet):
            try:
                f(g, tag, w, st, jetzt, test)
            except Exception as e:
                log(f"{f.__name__}: {type(e).__name__}: {str(e)[:150]}")
        try:
            abends(g, tag, st, jetzt, test)
        except Exception as e:
            log(f"abends: {type(e).__name__}: {str(e)[:150]}")
    if server_ok:
        for f in (lambda: hinweise(st, test), lambda: workout(st, jetzt, test)):
            try:
                f()
            except Exception as e:
                log(f"auto: {type(e).__name__}: {str(e)[:150]}")
    try:
        waechter(st, w, server_ok, jetzt, test)
    except Exception as e:
        log(f"waechter: {e}")
    if not test:
        schreiben(STATE, st)


if __name__ == "__main__":
    from garminconnect import Garmin

    g = Garmin()
    g.login(str(Path.home() / ".garminconnect"))
    tag = datetime.date.today().isoformat()
    sys.path.insert(0, str(BASE))
    import kira_garmin_live as kl  # type: ignore

    lauf(g, tag, kl.abrufen(g, tag), test="--test" in sys.argv)
