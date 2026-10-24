"""Coach Kira – Garmin Personal API Integration.

Einbindung in ~/coach-kira-runner neben kira_runner.py:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from kira_runner import log, http_get_json, http_post_json, notify, SCRIPT_URL, TZ, BASE
"""
from __future__ import annotations

import argparse
import base64
import datetime
import json
import os
import sys
import time
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from urllib.error import HTTPError

# ---------------------------------------------------------------------------
# Konstanten
# ---------------------------------------------------------------------------

API = "https://personal.connectapi.garmin.com"
TOKEN_PATH = "/personalapitoken-service/personal-api/personal-api-key/token"
KEY_FILE = Path.home() / ".config" / "garmin" / "connect-personal-api-key"

# Wird nach dem Import von kira_runner überschrieben: BASE ist dort definiert
from kira_runner import log, http_get_json, http_post_json, notify, SCRIPT_URL, TZ, BASE

TOKEN_CACHE = BASE / ".garmin_token.json"
STATE_PATH = BASE / "garmin_state.json"

START_TIMELINE = "2025-08-26"


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _atomar_schreiben(path: Path, data: dict) -> None:
    """JSON atomar schreiben: .tmp + rename."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(str(tmp), str(path))


def _laden_state() -> dict:
    """STATE lesen, bei Fehler leeres Dict."""
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _laden_key() -> str:
    """Garmin-API-Schlüssel aus KEY_FILE lesen, strip()."""
    if not KEY_FILE.exists():
        raise RuntimeError(f"Schlüsseldatei fehlt: {KEY_FILE}")
    return KEY_FILE.read_text(encoding="utf-8").strip()


def _b64url_decode(s: str) -> bytes:
    """Base64url-Decoding mit Padding-Ergänzung."""
    padding = 4 - len(s) % 4
    if padding != 4:
        s += "=" * padding
    return base64.urlsafe_b64decode(s)


def _user_agent() -> str:
    return "coach-kira-garmin"


# ---------------------------------------------------------------------------
# Token-Management
# ---------------------------------------------------------------------------

def get_token() -> str:
    """API-Token holen, cache-nutzen falls noch valide. Gibt JWT-Text zurück."""
    key = _laden_key()
    # Cache prüfen
    if TOKEN_CACHE.exists():
        try:
            with open(TOKEN_CACHE, "r", encoding="utf-8") as f:
                cache = json.load(f)
            exp = cache.get("exp", 0)
            if exp - 300 > datetime.datetime.now().timestamp():
                return cache["token"]
        except Exception:
            pass

    # Neues Token holen
    url = API + TOKEN_PATH
    req = Request(
        url,
        data=b"",
        headers={
            "Authorization": f"ApiKey {key}",
            "Accept": "*/*",
            "User-Agent": _user_agent(),
        },
        method="POST",
    )
    try:
        with urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
    except HTTPError as e:
        if e.code in (401, 403):
            notify("Coach Kira", "Garmin-Schlüssel ungültig oder abgelaufen")
            raise RuntimeError(f"Schlüssel ungültig oder abgelaufen (HTTP {e.code})")
        raise

    # JWT als roher Text oder JSON?
    stripped = body.strip()
    if stripped and stripped[0] == "{":
        j = json.loads(stripped)
        token = j.get("access_token") or j.get("accessToken") or j.get("token", "")
    else:
        token = stripped

    # exp aus JWT-Payload lesen
    parts = token.split(".")
    if len(parts) >= 2:
        payload_b64 = parts[1]
        try:
            payload_bytes = _b64url_decode(payload_b64)
            payload = json.loads(payload_bytes.decode("utf-8"))
            exp = payload.get("exp")
        except Exception:
            exp = None
    else:
        exp = None

    if exp is None or not isinstance(exp, (int, float)):
        exp = datetime.datetime.now().timestamp() + 23 * 3600

    # Cache schreiben (chmod 0o600)
    cache_data = {"token": token, "exp": exp}
    with open(TOKEN_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache_data, f)
    os.chmod(TOKEN_CACHE, 0o600)

    return token


def _api_get_with_token(path: str, params=None):
    """Interne GET-Funktion mit Token-Handling und Retry."""
    token = get_token()

    def _do_get():
        url = API + path
        if params:
            url += "?" + urlencode(params)
        req = Request(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "User-Agent": _user_agent(),
            },
        )
        with urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
        return resp.status, body

    for attempt in range(3):
        try:
            status, body = _do_get()
            if status == 204 or not body.strip():
                return None
            if status == 401:
                # Token-Cache löschen und wiederholen
                try:
                    TOKEN_CACHE.unlink(missing_ok=True)
                except Exception:
                    pass
                token = get_token()  # neues Token holen
                continue
            if status in (429, 500, 502, 503, 504):
                pause = 5 if attempt == 0 else 20
                time.sleep(pause)
                continue
            return json.loads(body)
        except HTTPError as e:
            if e.code == 401:
                try:
                    TOKEN_CACHE.unlink(missing_ok=True)
                except Exception:
                    pass
                token = get_token()
                continue
            if e.code in (429, 500, 502, 503, 504) and attempt < 2:
                pause = 5 if attempt == 0 else 20
                time.sleep(pause)
                continue
            raise
        except Exception:
            if attempt < 2:
                time.sleep(5)
                continue
            raise

    return None


def api_get(path: str, params=None):
    """GET API+path mit Token-Handling. Gibt json.loads oder None zurück."""
    return _api_get_with_token(path, params)


# ---------------------------------------------------------------------------
# Daten holen
# ---------------------------------------------------------------------------

def tag_aus(a: dict) -> str:
    """Datum aus einer Aktivität extrahieren: startTimeLocal[:10]."""
    st = a.get("startTimeLocal", "")
    return st[:10] if len(st) >= 10 else ""


def activities(von: str, bis: str):
    """Alle Aktivitäten zwischen von und bis holen (geblättert)."""
    all_acts = []
    start = 0
    limit = 99
    while True:
        params = {
            "startDate": von,
            "endDate": bis,
            "start": start,
            "limit": limit,
            "sortOrder": "asc",
        }
        result = api_get("/activitylist-service/activities/search/activities", params)
        if not result:
            break
        if isinstance(result, list):
            chunk = result
        elif isinstance(result, dict) and "results" in result:
            chunk = result["results"]
        else:
            break
        all_acts.extend(chunk)
        if len(chunk) < 99:
            break
        start += limit
        if start > 50 * limit:
            break

    return all_acts


def km_hm_je_tag(acts):
    """{datum: {"km": float, "hm": int, "n": int}} aus Aktivitätsliste."""
    days: dict[str, dict] = {}
    for a in acts:
        d = tag_aus(a)
        if not d:
            continue
        dist = a.get("distance") or 0
        elev = a.get("elevationGain") or 0
        if d not in days:
            days[d] = {"km": 0.0, "hm": 0, "n": 0}
        days[d]["km"] += dist / 1000.0
        days[d]["hm"] += elev
        days[d]["n"] += 1

    # Aufrunden auf 2 Dezimalstellen, hm auf int
    for d in days:
        days[d]["km"] = round(days[d]["km"], 2)
        days[d]["hm"] = round(days[d]["hm"])
    return days


def sport_map(type_key: str) -> str | None:
    """Map activityType.typeKey zu Sport-Kategorie."""
    tk = (type_key or "").lower()
    if "run" in tk:
        return "Run"
    if any(x in tk for x in ("cycl", "bik", "ride")):
        return "Bike"
    if any(x in tk for x in ("hik", "walk", "mountaineer")):
        return "Hike"
    if any(x in tk for x in ("row", "kayak", "canoe", "paddl")):
        return "Row"
    if "swim" in tk:
        return "Swim"
    # Wichtig: "skat" vor "ski"! (inline_skating -> Skike)
    if "skat" in tk:
        return "Skike"
    if any(x in tk for x in ("ski", "snowshoe")):
        return "Ski"
    return None


def send_day_metrics(rows: list, modus: str, probe: bool) -> list:
    """Zeilenpaketiert zu 120 pro POST senden. Gibt Liste der Antworten zurück."""
    if not rows:
        return []
    results = []
    for i in range(0, max(len(rows), 1), 120):
        chunk = rows[i:i + 120]
        payload = {
            "kind": "saveDayMetrics",
            "rows": chunk,
            "modus": modus,
            "probe": probe,
        }
        resp = http_post_json(SCRIPT_URL, payload)
        if not resp.get("ok"):
            err = resp.get("error", "unbekannt")
            log(f"[garmin] saveDayMetrics Fehler: {err}")
            raise RuntimeError(err)
        results.append(resp)
    return results


# ---------------------------------------------------------------------------
# Schlaf, Aktivität, Abend
# ---------------------------------------------------------------------------

def morgens(datum: str):
    """Morgens-Daten für ein Datum sammeln (Sleep, RHR, HRV)."""
    result = {}

    # Sleep
    try:
        sleep_data = api_get("/sleep-service/sleep/dailySleep", {"date": datum})
        if sleep_data:
            sleep_seconds = sleep_data.get("sleepTimeSeconds") or 0
            if sleep_seconds > 0:
                result["sleep_hours"] = round(sleep_seconds / 3600.0, 2)
            scores = sleep_data.get("sleepScores") or {}
            overall = scores.get("overall")
            if isinstance(overall, dict) and "value" in overall:
                result["sleep_score_0_100"] = overall.get("value")
    except Exception as e:
        log(f"[garmin] Schlaf({datum}): {e}")

    # RHR from usersummary
    try:
        summary = api_get("/usersummary-service/usersummary/daily", {"calendarDate": datum})
        if summary:
            rhr = summary.get("restingHeartRate")
            if rhr is not None:
                result["rhr_bpm"] = rhr
    except Exception as e:
        log(f"[garmin] RHR({datum}): {e}")

    # HRV
    try:
        hrv_data = api_get(f"/hrv-service/hrv/{datum}")
        if hrv_data:
            hrv_summary = hrv_data.get("hrvSummary") or {}
            weekly_avg = hrv_summary.get("weeklyAvg")
            if weekly_avg is not None:
                result["hrv_status"] = weekly_avg
            baseline = hrv_summary.get("baseline") or {}
            bal_low = baseline.get("balancedLow")
            bal_high = baseline.get("balancedUpper")
            if bal_low is not None and bal_high is not None:
                result["hrv_threshholds"] = f"{bal_low};{bal_high}"
    except Exception as e:
        log(f"[garmin] HRV({datum}): {e}")

    return result


def aktivitaet(datum: str, acts_des_tags: list):
    """Aktivitäts-Daten für einen Tag analysieren."""
    if not acts_des_tags:
        return None

    # load_fb_day: Summe activityTrainingLoad über alle Aktivitäten des Tages
    has_load = any(a.get("activityTrainingLoad") is not None for a in acts_des_tags)
    total_load = 0
    for a in acts_des_tags:
        load = a.get("activityTrainingLoad")
        if load is not None:
            total_load += load
    result = {}
    if has_load:
        result["load_fb_day"] = round(total_load)

    # Letzte Aktivität für TE und Zonen
    last_act = acts_des_tags[-1]

    # Aerobic_TE, Anaerobic_TE – nur setzen wenn nicht None
    aer_te = last_act.get("aerobicTrainingEffect")
    ana_te = last_act.get("anaerobicTrainingEffect")
    if aer_te is not None:
        result["Aerobic_TE"] = round(aer_te, 1)
    if ana_te is not None:
        result["Anaerobic_TE"] = round(ana_te, 1)

    # Sport_x
    sport_type = last_act.get("activityType") or {}
    type_key = sport_type.get("typeKey", "") if isinstance(sport_type, dict) else ""
    sport_name = sport_map(type_key)
    if sport_name:
        result["Sport_x"] = sport_name

    # Zone aus hrTimeInZone_1..5 der letzten Aktivität
    zone_times = []
    for i in range(1, 6):
        zt_key = f"hrTimeInZone_{i}"
        zt = last_act.get(zt_key)
        if zt is None:
            zt = 0
        zone_times.append((i, float(zt)))

    max_zones = sorted(zone_times, key=lambda x: x[1], reverse=True)
    if max_zones and max_zones[0][1] > 0:
        top_z = max_zones[0]
        second_z = max_zones[1] if len(max_zones) > 1 else None
        top_zone_num = top_z[0]

        if second_z and (second_z is not None) and second_z[1] > 0:
            diff = abs(top_zone_num - second_z[0])
            if diff == 1 and second_z[1] >= top_z[1] * 0.75:
                lower, higher = min(top_zone_num, second_z[0]), max(
                    top_zone_num, second_z[0]
                )
                result["Zone"] = f"Z{lower}, Z{higher}"
            else:
                result["Zone"] = f"Z{top_zone_num}"
        else:
            result["Zone"] = f"Z{top_zone_num}"

    # Info-Felder: Anzahl, km, hm des Tages
    result["anzahl"] = len(acts_des_tags)
    km_val = 0.0
    hm_val = 0
    for a in acts_des_tags:
        km_val += (a.get("distance") or 0) / 1000.0
        hm_val += a.get("elevationGain") or 0
    result["km"] = round(km_val, 2)
    result["hm"] = round(hm_val)

    return result


def abends(datum: str):
    """Abend-Daten (Ernährung, Kalorienverbrauch)."""
    result = {}

    try:
        nutrition = api_get("/nutrition-service/food/logs/" + datum, {"includeNutrition": "true"})
        if nutrition:
            content = nutrition.get("dailyNutritionContent") or {}
            kcal_in = content.get("calories") if content else None
            carb_g = content.get("carbs") if content else None
            protein_g = content.get("protein") if content else None
            fat_g = content.get("fat") if content else None

            if kcal_in is not None:
                result["kcal_in"] = round(kcal_in)
            if carb_g is not None:
                result["carb_g"] = round(carb_g)
            if protein_g is not None:
                result["protein_g"] = round(protein_g)
            if fat_g is not None:
                result["fat_g"] = round(fat_g)

            # Leeres Ernährungstagebuch
            if result.get("kcal_in", 0) == 0:
                result.pop("kcal_in", None)
                result.pop("carb_g", None)
                result.pop("protein_g", None)
                result.pop("fat_g", None)
    except Exception as e:
        log(f"[garmin] Abend Ernährung({datum}): {e}")

    try:
        summary = api_get("/usersummary-service/usersummary/daily", {"calendarDate": datum})
        if summary:
            kcal_out = summary.get("totalKilocalories")
            if kcal_out is not None:
                result["kcal_out"] = round(kcal_out)
    except Exception as e:
        log(f"[garmin] Abend kcal_out({datum}): {e}")

    # Deficit
    if "kcal_in" in result and "kcal_out" in result:
        result["deficit"] = result["kcal_out"] - result["kcal_in"]

    return result


def _send_save_extra(datum: str, text: str, data: dict):
    """saveExtra ans Apps Script senden. Raises RuntimeError wenn ok false."""
    payload = {
        "kind": "saveExtra",
        "typ": "garmin_api",
        "datum": datum,
        "text": text,
        "data": data,
        "quelle": "garmin-api",
    }
    resp = http_post_json(SCRIPT_URL, payload)
    if not resp.get("ok"):
        err = resp.get("error", "unbekannt")
        raise RuntimeError(err)
    return resp


# ---------------------------------------------------------------------------
# Befehle (argparse)
# ---------------------------------------------------------------------------

def cmd_test():
    """Token holen, 1 Aktivität der letzten 7 Tage prüfen."""
    now = datetime.datetime.now(TZ)
    token = get_token()

    # JWT exp lesen
    parts = token.split(".")
    if len(parts) >= 2:
        payload_b64 = parts[1]
        try:
            payload_bytes = _b64url_decode(payload_b64)
            payload = json.loads(payload_bytes.decode("utf-8"))
            exp_ts = payload.get("exp")
            if exp_ts:
                exp_dt = datetime.datetime.fromtimestamp(exp_ts, tz=TZ)
                print(f"Token ok, gültig bis {exp_dt.strftime('%H:%M')}")
        except Exception:
            log("Token ok, exp konnte nicht gelesen werden")

    # Letzte 7 Tage Aktivitäten
    today = now.date()
    von = (today - datetime.timedelta(days=7)).isoformat()
    bis = today.isoformat()
    acts = activities(von, bis)

    print(now.strftime("%Y-%m-%dT%H:%M:%S%z"))
    print(f"Aktivitäten (7 Tage): {len(acts)}")
    if acts:
        first = acts[0]
        d = tag_aus(first)
        tk = (first.get("activityType") or {}).get("typeKey", "")
        km = round((first.get("distance") or 0) / 1000, 2)
        hm = round(first.get("elevationGain") or 0)
        print(f"Erste: {d}, typeKey={tk}, km={km}, hm={hm}")


def cmd_backfill(von=None, bis=None, probe=False, alle=False):
    """Zeitrückholung: Aktivitäten sammeln und alle auf einmal senden."""
    now = datetime.datetime.now(TZ)
    today = now.date()
    if von is None:
        von = START_TIMELINE
    if bis is None:
        bis = today.isoformat()

    von_dt = datetime.date.fromisoformat(von)
    bis_dt = datetime.date.fromisoformat(bis)

    all_km_hm: dict[str, dict] = {}
    total_km = 0.0
    total_hm = 0

    while von_dt <= bis_dt:
        fenster_ende = min(von_dt + datetime.timedelta(days=31), bis_dt + datetime.timedelta(days=1))
        fenster_bis = (fenster_ende - datetime.timedelta(days=1)).isoformat()

        acts = activities(von_dt.isoformat(), fenster_bis)
        days_data = km_hm_je_tag(acts)

        for d in sorted(days_data.keys()):
            dd = days_data[d]
            all_km_hm[d] = dd
            total_km += dd["km"]
            total_hm += dd["hm"]

        von_dt = fenster_ende

    modus = "alle" if alle else "leer"
    rows = [
        {"datum": d, "km": all_km_hm[d]["km"], "hm": all_km_hm[d]["hm"]}
        for d in sorted(all_km_hm.keys())
    ]

    results = send_day_metrics(rows, modus=modus, probe=probe)

    print(f"Tage mit Aktivität: {len(all_km_hm)}, Summe km: {round(total_km, 2)}, Summe hm: {total_hm}")
    for resp in results:
        print(json.dumps(resp, ensure_ascii=False))

    return all_km_hm


def cmd_tick():
    """Aktuellen Stand abrufen und Vorschlag für heute+gestern senden."""
    now = datetime.datetime.now(TZ)
    if not (now.hour >= 5 and now.hour < 24):
        log("[garmin] tick: Außerhalb der Zeit [05:00, 23:59]")
        return

    state = _laden_state()
    state.setdefault("vorschlag", {})  # Fix Cloud-Review: Bereich vor erstem Zugriff anlegen

    # Fehlerzähler persistent in STATE["fehler"]
    fehler = state.get("fehler", 0)

    def _inc_fehler():
        nonlocal fehler
        fehler += 1
        state["fehler"] = fehler

    def _on_success():
        nonlocal fehler
        fehler = 0
        state["fehler"] = 0

    # Ablauf-Erinnerung: einmal pro Tag
    heute_str = now.strftime("%Y-%m-%d")
    ablauf_gemeldet = state.get("ablauf_gemeldet", "")

    today_str = now.strftime("%Y-%m-%d")
    gestern_dt = now - datetime.timedelta(days=1)
    gestern_str = gestern_dt.strftime("%Y-%m-%d")

    #.activities EIN Abruf, nach Tag gruppieren
    try:
        acts = activities(today_str, today_str) + activities(gestern_str, gestern_str)
    except Exception as e:
        log(f"[garmin] tick activities: {e}")
        _inc_fehler()
        acts = []

    # Nach Tag gruppieren
    acts_by_day: dict[str, list] = {}
    for a in acts:
        da = tag_aus(a)
        if da:
            acts_by_day.setdefault(da, []).append(a)

    # km/hm für heute und gestern senden (nur Tage mit Aktivität)
    changed = False
    for d in [today_str, gestern_str]:
        days_data = km_hm_je_tag(acts_by_day.get(d, []))
        if not days_data:
            continue
        for d2, dd in days_data.items():
            prev = state.get("kmhm", {}).get(d2)
            if prev != [round(dd["km"], 2), dd["hm"]]:
                try:
                    send_day_metrics([{"datum": d2, "km": dd["km"], "hm": dd["hm"]}], modus="alle", probe=False)
                    state.setdefault("kmhm", {})[d2] = [round(dd["km"], 2), dd["hm"]]
                    log(f"[garmin] tick km/hm {d2}: {dd['km']} km / {dd['hm']} hm gesendet")
                    changed = True
                except Exception as e:
                    log(f"[garmin] tick km/hm {d2}: {e}")

    if changed:
        _on_success()

    # Vorschlag HEUTE IMMER bauen (auch ohne Aktivität)
    try:
        m = morgens(today_str) if True else {}
        a_dict = aktivitaet(today_str, acts_by_day.get(today_str, [])) if acts_by_day.get(today_str) else None
        ab = abends(today_str)

        stand = now.strftime("%H:%M")
        extra_data: dict = {}
        if m:
            extra_data["morgens"] = m
        if a_dict is not None:
            extra_data["aktivitaet"] = a_dict
        if ab:
            extra_data["abends"] = ab

        # Keine "stand" Daten zum Vergleich verwenden
        standless = {k: v for k, v in extra_data.items()}

        prev_vorschlag = state.get("vorschlag", {}).get(today_str, {})
        if standless and prev_vorschlag != standless:
            text_felder = sum(len(v) for v in [m, a_dict or {}, ab if ab else {}])
            kurztext = f"Garmin-API {stand}: Morgens {len(m)}, Aktivität {len(a_dict or {})}, Abends {len(ab if ab else {})} Felder"
            try:
                resp = _send_save_extra(today_str, kurztext, extra_data)
                if resp.get("ok"):
                    state["vorschlag"][today_str] = extra_data.copy()
                else:
                    log(f"[garmin] tick heute Vorschlag gesendet (fehlerhaft)")
                    _inc_fehler()
            except Exception as e:
                log(f"[garmin] tick heute Vorschlag senden: {e}")
                _inc_fehler()

        # STATE updaten – vorschlag und kmhm auf 10 neuesten
        state.setdefault("vorschlag", {})
        state["kmhm"] = dict(list(state.get("kmhm", {}).items())[-10:])
        state["vorschlag"] = dict(list(state.get("vorschlag", {}).items())[-10:])

    except Exception as e:
        log(f"[garmin] tick Vorschlag heute: {e}")
        _inc_fehler()

    # Vorschlag GESTERN: aktivitaet + abends (kein morgens)
    try:
        a_dict_g = aktivitaet(gestern_str, acts_by_day.get(gestern_str, [])) if acts_by_day.get(gestern_str) else None
        ab_g = abends(gestern_str)

        extra_data_g: dict = {}
        if a_dict_g is not None:
            extra_data_g["aktivitaet"] = a_dict_g
        if ab_g:
            extra_data_g["abends"] = ab_g

        standless_g = {k: v for k, v in extra_data_g.items()}
        prev_vorschlag_g = state.get("vorschlag", {}).get(gestern_str, {})
        if standless_g and prev_vorschlag_g != standless_g:
            text_felder_g = sum(len(v) for v in [a_dict_g or {}, ab_g if ab_g else {}])
            kurztext_g = f"Garmin-API {stand}: Aktivität {len(a_dict_g or {})}, Abends {len(ab_g if ab_g else {})} Felder"
            try:
                resp = _send_save_extra(gestern_str, kurztext_g, extra_data_g)
                if resp.get("ok"):
                    state["vorschlag"][gestern_str] = extra_data_g.copy()
                else:
                    log(f"[garmin] tick gestern Vorschlag gesendet (fehlerhaft)")
                    _inc_fehler()
            except Exception as e:
                log(f"[garmin] tick gestern Vorschlag senden: {e}")
                _inc_fehler()

        state.setdefault("vorschlag", {})
        state["kmhm"] = dict(list(state.get("kmhm", {}).items())[-10:])
        state["vorschlag"] = dict(list(state.get("vorschlag", {}).items())[-10:])

    except Exception as e:
        log(f"[garmin] tick Vorschlag gestern: {e}")
        _inc_fehler()

    # Fehlerzähler prüfen: genau bei 3 notify
    if fehler == 3:
        notify("Coach Kira", "Garmin-Abruf klappt nicht: wiederholte Fehler")
        fehler = 0
        state["fehler"] = 0

    # Ablauf-Erinnerung: höchstens einmal pro Tag
    if ablauf_gemeldet != heute_str:
        try:
            key_expiry = state.get("key_expiry")
            if key_expiry:
                exp_date = datetime.date.fromisoformat(key_expiry)
                now_date = now.date()
                diff = (exp_date - now_date).days
                if 0 <= diff <= 7:
                    formatted = exp_date.strftime("%d.%m.%Y")
                    notify("Coach Kira", f"Garmin-Schlüssel läuft am {formatted} ab")
                    state["ablauf_gemeldet"] = heute_str
        except Exception:
            pass

    # STATE nur einmal am Ende speichern
    _atomar_schreiben(STATE_PATH, state)


def cmd_vorschau(datum=None):
    """Vorschlag für Datum als JSON ausgeben, nichts senden."""
    if datum is None:
        datum = datetime.datetime.now(TZ).date().isoformat()

    acts = activities(datum, datum)
    days_data = km_hm_je_tag(acts)

    m = {}
    a_dict = None
    ab_data = {}
    last_acts_by_day: dict[str, list] = {}
    for a in acts:
        da = tag_aus(a)
        last_acts_by_day.setdefault(da, []).append(a)

    todays = last_acts_by_day.get(datum, [])
    if todays:
        a_dict = aktivitaet(datum, todays)

    try:
        m = morgens(datum)
    except Exception as e:
        log(f"[garmin] vorschau morgens({datum}): {e}")

    try:
        ab_data = abends(datum)
    except Exception as e:
        log(f"[garmin] vorschau abends({datum}): {e}")

    stand = datetime.datetime.now(TZ).strftime("%H:%M")
    extra_data = {"stand": stand}
    if m:
        extra_data["morgens"] = m
    if a_dict:
        extra_data["aktivitaet"] = a_dict
    if ab_data:
        extra_data["abends"] = ab_data

    print(json.dumps(extra_data, ensure_ascii=False, indent=2, default=str))


def cmd_ablauf(bis):
    """Ablaufdatum des Schlüssels im STATE speichern."""
    state = _laden_state()
    state["key_expiry"] = bis
    try:
        exp_dt = datetime.date.fromisoformat(bis)
        state["key_expiry_formatted"] = exp_dt.strftime("%d.%m.%Y")
    except Exception:
        pass
    _atomar_schreiben(STATE_PATH, state)
    log(f"[garmin] Schlüssel läuft am {bis} ab")


def main():
    """Haupteingang: argparse mit Unterbefehlen."""
    parser = argparse.ArgumentParser(description="Coach Kira – Garmin Personal API")
    subparsers = parser.add_subparsers(dest="command")

    # test
    subparsers.add_parser("test", help="Token-Prüfung, 7 Tage Aktivität")

    # backfill
    p_backfill = subparsers.add_parser("backfill", help="Zeitrückholung")
    p_backfill.add_argument("--von", default=None)
    p_backfill.add_argument("--bis", default=None)
    p_backfill.add_argument("--probe", action="store_true")
    p_backfill.add_argument("--alle", action="store_true")

    # tick
    subparsers.add_parser("tick", help="Heute+Gestern Abruf")

    # vorschau
    p_vorschau = subparsers.add_parser("vorschau", help="Vorschlag als JSON")
    p_vorschau.add_argument("--datum", default=None)

    # ablauf
    p_ablauf = subparsers.add_parser("ablauf", help="Ablaufdatum speichern")
    p_ablauf.add_argument("--bis", required=True)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    try:
        if args.command == "test":
            cmd_test()
        elif args.command == "backfill":
            cmd_backfill(von=args.von, bis=args.bis, probe=bool(args.probe), alle=bool(args.alle))
        elif args.command == "tick":
            cmd_tick()
        elif args.command == "vorschau":
            cmd_vorschau(datum=args.datum)
        elif args.command == "ablauf":
            cmd_ablauf(bis=args.bis)
    except Exception as e:
        log(f"[garmin] FATAL: {e}")
        raise


if __name__ == "__main__":
    main()
