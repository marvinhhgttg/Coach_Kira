"""Coach Kira – Garmin Personal API Integration.

Einbindung in ~/coach-kira-runner neben kira_runner.py:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from kira_runner import log, http_get_json, http_post_json, notify, SCRIPT_URL, TZ, BASE
"""
from __future__ import annotations

import argparse
import base64
import datetime
import hashlib
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
# Garmin Tag-Details: tag_details, text_fuer, send_tage
# ---------------------------------------------------------------------------

def _windows_ab(von: str, bis: str, window: int):
    """Generator von Fenster-Datenpaaren (start_iso, end_iso) in window-Tages-Schritten."""
    von_dt = datetime.date.fromisoformat(von)
    bis_dt = datetime.date.fromisoformat(bis)
    while von_dt <= bis_dt:
        end = min(von_dt + datetime.timedelta(days=window), bis_dt + datetime.timedelta(days=1))
        yield von_dt.isoformat(), (end - datetime.timedelta(days=1)).isoformat()
        von_dt = end


def _dailysummaries(von: str, bis: str) -> dict:
    """{calendarDate: UserDailySummaryV2DTO} aus v2/dailysummaries."""
    result = {}
    for start, end in _windows_ab(von, bis, 365):
        try:
            data = api_get("/usersummary-service/v2/dailysummaries", {"startDate": start, "endDate": end})
            if isinstance(data, dict):
                items = data.get("data", [])
            elif isinstance(data, list):
                items = data
            else:
                items = []
            for entry in items:
                cd = entry.get("calendarDate", "")
                if cd:
                    result[cd] = entry
        except Exception as e:
            log(f"[garmin] dailysummaries({start}/{end}): {e}")
    return result


def _dailysummary_day(dailies: dict, datum: str) -> dict | None:
    """tag-Dict aus UserDailySummaryV2DTO für ein Datum, nur existierende Schlüssel."""
    entry = dailies.get(datum)
    if not entry:
        return None

    wc = entry.get("wellnessChronology") or {}
    mov = entry.get("movement") or {}
    cal = entry.get("calories") or {}
    hr = entry.get("heartRate") or {}
    im = entry.get("intensityMinutes") or {}
    stress = entry.get("stress") or {}
    bb = entry.get("bodyBattery") or {}
    hydra = entry.get("hydration") or {}
    resp = entry.get("respiration") or {}
    spo2 = entry.get("pulseOx") or {}
    jetlag = entry.get("jetLag")

    day = {}
    m_val = mov.get("steps")
    if m_val is not None:
        st = m_val or {}
        schritte = st.get("value")
        if schritte is not None: day["schritte"] = int(schritte)
    z = entry.get("dailyStepGoal") or entry.get("totalSteps")  # dailyStepGoal in V2
    s_goal = (entry or {}).get("dailyStepGoal")
    if s_goal is not None: day["schritte_ziel"] = int(s_goal)
    fl = mov.get("floorsAscended") or {}
    if fl is not None:
        et = fl.get("value")
        if et is not None: day["etagen"] = int(et)
    kc = cal.get("burnedTotal") if cal.get("burnedTotal") else cal.get("burnedResting", 0) + cal.get("burnedActive", 0)
    bt = cal.get("burnedTotal")
    if isinstance(bt, dict): bt = bt.get("value")
    if bt is not None: day["kcal_gesamt"] = int(bt)
    ba = cal.get("burnedActive")
    if isinstance(ba, dict): ba = ba.get("value")
    if ba is not None: day["kcal_aktiv"] = int(ba)
    br = cal.get("burnedResting")
    if isinstance(br, dict): br = br.get("value")
    if br is not None: day["kcal_ruhe"] = int(br)
    cg = cal.get("consumedValue")
    if isinstance(cg, dict): cg = cg.get("value")
    if cg is not None: day["kcal_gegessen"] = int(cg)
    hm_min = hr.get("minValue")
    if isinstance(hm_min, dict): hm_min = hm_min.get("value")
    if hm_min is not None: day["hf_min"] = int(hm_min)
    hm_max = hr.get("maxValue")
    if isinstance(hm_max, dict): hm_max = hm_max.get("value")
    if hm_max is not None: day["hf_max"] = int(hm_max)
    rhr = hr.get("restingValue")
    if isinstance(rhr, dict): rhr = rhr.get("value")
    if rhr is not None: day["rhr"] = int(rhr)
    intm = im.get("moderate")
    if isinstance(intm, dict): intm = intm.get("value")
    if intm is not None: day["int_mod"] = int(intm)
    ivg = im.get("vigorous")
    if isinstance(ivg, dict): ivg = ivg.get("value")
    if ivg is not None: day["int_vig"] = int(ivg)

    sa = stress.get("avgLevel")
    if isinstance(sa, dict): sa = sa.get("value")
    if sa is not None: day["stress_avg"] = int(sa)
    sm = stress.get("maxLevel")
    if isinstance(sm, dict): sm = sm.get("value")
    if sm is not None: day["stress_max"] = int(sm)
    sq = stress.get("qualifier")
    if sq: day["stress_q"] = sq

    # stress_min: Duration-Werte / 60000 gerundet, verschachtelt
    stress_min = {}
    rdur = stress.get("restDurationInMillis")
    if rdur is not None: stress_min["ruhe"] = round(rdur / 60000)
    lsd = stress.get("lowStressDurationInMillis")
    if lsd is not None: stress_min["niedrig"] = round(lsd / 60000)
    msd = stress.get("mediumStressDurationInMillis")
    if msd is not None: stress_min["mittel"] = round(msd / 60000)
    hsd = stress.get("highStressDurationInMillis")
    if hsd is not None: stress_min["hoch"] = round(hsd / 60000)
    adur = stress.get("activityDurationInMillis")
    if adur is not None: stress_min["aktiv"] = round(adur / 60000)
    if stress_min: day["stress_min"] = stress_min

    bb_min_val = bb.get("minValue")
    if bb_min_val is not None: day["bb_min"] = int(bb_min_val)
    bb_max_val = bb.get("maxValue")
    if bb_max_val is not None: day["bb_max"] = int(bb_max_val)
    charged = bb.get("chargedValue")
    if charged is not None: day["bb_geladen"] = int(charged)
    drained = bb.get("drainedValue")
    if drained is not None: day["bb_verbraucht"] = int(drained)

    end_of_day_val = bb.get("endOfDayValue")
    if end_of_day_val is not None: day["bb_ende"] = int(end_of_day_val)

    ra = resp.get("avgValue")
    if ra is not None: day["atem_avg"] = int(ra)
    rmin = resp.get("minValue")
    if rmin is not None: day["atem_min"] = int(rmin)
    rmax = resp.get("maxValue")
    if rmax is not None: day["atem_max"] = int(rmax)

    s_avg = spo2.get("avgValue")
    if s_avg is not None: day["spo2_avg"] = int(s_avg)
    s_min = spo2.get("minValue")
    if s_min is not None: day["spo2_min"] = int(s_min)

    h_consumed = hydra.get("consumedInFractionalMl")
    if h_consumed is not None and int(h_consumed) > 0: day["wasser_ml"] = round(int(h_consumed))

    if jetlag and isinstance(jetlag, dict) and jetlag.get("tripName"):
        day["jetlag"] = jetlag.get("tripName")

    return day


def _sleep(von: str, bis: str) -> dict:
    """{calendarDate: DailySleepDTO} aus dailySleepSummariesByDate."""
    result = {}
    for start, end in _windows_ab(von, bis, 31):
        try:
            data = api_get("/sleep-service/sleep/dailySleepSummariesByDate", {"startDate": start, "endDate": end})
            if isinstance(data, list):
                for entry in data:
                    cd = entry.get("calendarDate", "")
                    if cd:
                        result[cd] = entry
        except Exception as e:
            log(f"[garmin] sleep({start}/{end}): {e}")
    return result


def _sleep_day(sleeps: dict, datum: str) -> dict | None:
    """schlaf-Dict aus DailySleepDTO für ein Datum."""
    entry = sleeps.get(datum)
    if not entry:
        return None

    sn = entry.get("sleepNeed", {}) or {}
    scores = entry.get("sleepScores") or {}
    overall = scores.get("overall") if isinstance(scores, dict) else None
    oval = overall.get("value") if isinstance(overall, dict) else None

    Schlaf = {}
    ds = entry.get("sleepTimeSeconds")
    if ds is not None: Schlaf["dauer_s"] = int(ds)
    ts = entry.get("deepSleepSeconds")
    if ts is not None: Schlaf["tief_s"] = int(ts)
    ls = entry.get("lightSleepSeconds")
    if ls is not None: Schlaf["leicht_s"] = int(ls)
    rms = entry.get("remSleepSeconds")
    if rms is not None: Schlaf["rem_s"] = int(rms)
    ws = entry.get("awakeSleepSeconds")
    if ws is not None: Schlaf["wach_s"] = int(ws)
    nap = entry.get("napTimeSeconds") or entry.get("nap_s")
    if nap is not None: Schlaf["nap_s"] = int(nap)
    if oval is not None: Schlaf["score"] = int(oval)

    aw = entry.get("awakeCount")
    if aw is not None: Schlaf["wachphasen"] = int(aw)

    ss = entry.get("avgSleepStress")
    if ss is not None: Schlaf["stress"] = round(ss, 1)

    hf = entry.get("avgHeartRate")
    if hf is not None: Schlaf["hf"] = round(hf, 1)

    at = entry.get("averageRespirationValue")
    if at is not None: Schlaf["atem"] = round(at, 1)

    sp = entry.get("averageSpO2Value")
    if sp is not None: Schlaf["spo2"] = round(sp, 1)

    # start / ende als HH:MM
    st_ts = entry.get("sleepStartTimestampLocal") or entry.get("sleepStartTimestampGMT")
    en_ts = entry.get("sleepEndTimestampLocal") or entry.get("sleepEndTimestampGMT")
    def _ts_to_hhmm(val):
        if not val: return None
        try:
            ts = int(val)
            dt = datetime.datetime.fromtimestamp(ts / 1000.0, tz=TZ)
            return dt.strftime("%H:%M")
        except Exception:
            if isinstance(val, str) and len(val) == 5:
                return val[:5]
            try:
                dt = datetime.datetime.fromisoformat(str(val)[:19])
                return dt.strftime("%H:%M")
            except Exception:
                return None

    start_mm = _ts_to_hhmm(st_ts)
    if start_mm: Schlaf["start"] = start_mm
    end_mm = _ts_to_hhmm(en_ts)
    if end_mm: Schlaf["ende"] = end_mm

    need_actual = sn.get("actual")
    if need_actual is not None and isinstance(need_actual, (int, float)):
        Schlaf["bedarf_min"] = int(need_actual)

    fb = entry.get("sleepScoreFeedback")
    if fb: Schlaf["feedback"] = fb

    return Schlaf


def _hrv(von: str, bis: str) -> dict:
    """{calendarDate: HrvSummaryDTO} aus hrv-service/hrv/daily."""
    result = {}
    for start, end in _windows_ab(von, bis, 31):
        try:
            data = api_get(f"/hrv-service/hrv/daily/{start}/{end}")
            if isinstance(data, dict):
                items = data.get("hrvSummaries", [])
            elif isinstance(data, list):
                items = data
            else:
                items = []
            for entry in items:
                cd = entry.get("calendarDate", "")
                if cd:
                    result[cd] = entry
        except Exception as e:
            log(f"[garmin] hrv({start}/{end}): {e}")
    return result


def _hrv_day(hrvs: dict, datum: str) -> dict | None:
    """hrv-Dict aus HrvSummaryDTO für ein Datum."""
    entry = hrvs.get(datum)
    if not entry:
        return None

    hrv_dict = {}
    ln = entry.get("lastNightAvg")
    if ln is not None: hrv_dict["nacht"] = int(ln)
    wk = entry.get("weeklyAvg")
    if wk is not None: hrv_dict["woche"] = int(wk)
    m5 = entry.get("lastNight5MinHigh")
    if m5 is not None: hrv_dict["max5"] = int(m5)
    st = entry.get("status")
    if st: hrv_dict["status"] = st

    base = entry.get("baseline") or {}
    bl = base.get("balancedLow")
    if bl is not None: hrv_dict["lo"] = int(bl)
    bh = base.get("balancedUpper")
    if bh is not None: hrv_dict["hi"] = int(bh)

    return hrv_dict


def _gewicht(von: str, bis: str) -> dict:
    """{summaryDate: kg_float} aus weight-service/weight/range."""
    result = {}
    for start, end in _windows_ab(von, bis, 31):
        try:
            data = api_get(f"/weight-service/weight/range/{start}/{end}")
            if isinstance(data, dict):
                items = data.get("dailyWeightSummaries", [])
            elif isinstance(data, list):
                items = data
            else:
                items = []
            for item in items:
                sd = item.get("summaryDate", "")
                lw = (item or {}).get("latestWeight") or {}
                wt = lw.get("weight") if isinstance(lw, dict) else None
                if sd and wt is not None:
                    val = wt / 1000.0 if wt > 1000 else wt
                    result[sd] = round(val, 1)
        except Exception as e:
            log(f"[garmin] gewicht({start}/{end}): {e}")
    return result


def _fitnessalter(heute: str) -> float | None:
    """Fitness Age für ein Datum (1 Nachkommastelle)."""
    try:
        data = api_get(f"/fitnessage-service/fitnessage/v2/{heute}")
        if isinstance(data, list) and len(data) > 0:
            return round(data[0].get("fitnessAge", 0), 1)
    except Exception as e:
        log(f"[garmin] fitnessalter({heute}): {e}")
    return None


def _bb_aufwachen(datum: str) -> int | None:
    """Body Battery at Wake Time für ein Datum."""
    try:
        data = api_get("/usersummary-service/usersummary/daily", {"calendarDate": datum})
        if data:
            val = data.get("bodyBatteryAtWakeTime")
            return int(val) if val is not None else None
    except Exception as e:
        log(f"[garmin] bb_aufwachen({datum}): {e}")
    return None


def _activities_block(von: str, bis: str) -> list:
    """Alle Aktivitäten in 31-Tage-Fenstern (wie backfill, aber Liste)."""
    all_acts = []
    for start, end in _windows_ab(von, bis, 31):
        try:
            acts = activities(start, end)
            all_acts.extend(acts)
        except Exception as e:
            log(f"[garmin] activities({start}/{end}): {e}")
    return all_acts


def _activity_day(acts: list, datum: str) -> dict | None:
    """Liste von Aktivitäts-Dicts für ein Datum (sortiert nach Startzeit)."""
    day_acts = [a for a in acts if tag_aus(a) == datum]
    day_acts.sort(key=lambda x: x.get("startTimeLocal", ""))
    if not day_acts:
        return None

    aus = []
    for a in day_acts:
        item = {}
        aid = a.get("activityId")
        if aid is not None: item["id"] = int(aid)
        aname = a.get("activityName")
        if aname: item["name"] = aname

        atype = a.get("activityType") or {}
        tkey = atype.get("typeKey", "") if isinstance(atype, dict) else ""
        if tkey: item["typ"] = tkey
        sport = sport_map(tkey)
        if sport: item["sport"] = sport

        st = a.get("startTimeLocal", "")
        if len(st) >= 16: item["start"] = st[11:16]

        dur = a.get("duration")
        if dur is not None: item["dauer_s"] = int(dur)
        md = a.get("movingDuration")
        if md is not None: item["bewegt_s"] = int(md)

        dist = a.get("distance")
        if dist is not None: item["km"] = round(dist / 1000.0, 2)

        eg = a.get("elevationGain")
        if eg is not None: item["hm"] = int(eg)
        el = a.get("elevationLoss")
        if el is not None: item["hm_ab"] = int(el)

        ahr = a.get("averageHR")
        if ahr is not None: item["hf_avg"] = int(ahr)
        mhr = a.get("maxHR")
        if mhr is not None: item["hf_max"] = int(mhr)

        # Zonen
        zones = []
        for i in range(1, 6):
            zk = f"hrTimeInZone_{i}"
            zv = a.get(zk)
            zones.append(int(zv) if zv is not None else 0)  # Fix Cloud-Review: immer 5 Werte
        if sum(zones) > 0: item["zonen_s"] = zones

        te_ae = a.get("aerobicTrainingEffect")
        if te_ae is not None: item["te_ae"] = round(te_ae, 1)
        te_an = a.get("anaerobicTrainingEffect")
        if te_an is not None: item["te_an"] = round(te_an, 1)
        tel = a.get("trainingEffectLabel")
        if tel: item["te_label"] = tel

        atl = a.get("activityTrainingLoad")
        if atl is not None: item["last"] = int(atl)

        cal = a.get("calories")
        if cal is not None: item["kcal"] = int(cal)

        bb_diff = a.get("differenceBodyBattery")
        if bb_diff is not None: item["bb"] = int(bb_diff)

        tmn = a.get("minTemperature")
        if tmn is not None: item["temp_min"] = int(tmn)
        tmx = a.get("maxTemperature")
        if tmx is not None: item["temp_max"] = int(tmx)

        speed = a.get("averageSpeed")
        if speed is not None: item["v_kmh"] = round(speed * 3.6, 1)

        p = a.get("avgPower")
        if p is not None: item["watt"] = int(p)
        np_pwr = a.get("normPower")
        if np_pwr is not None: item["watt_np"] = int(np_pwr)
        p20 = a.get("max20MinPower")
        if p20 is not None: item["watt_20"] = int(p20)

        cad = a.get("averageRunningCadenceInStepsPerMinute") or a.get("averageBikingCadenceInRevPerMinute")
        if cad is not None: item["kadenz"] = int(cad)

        vo2 = a.get("vO2MaxValue")
        if vo2 is not None: item["vo2max"] = round(vo2, 1)

        tss = a.get("trainingStressScore")
        if tss is not None: item["tss"] = round(tss, 1)
        if_val = a.get("intensityFactor")
        if if_val is not None: item["if"] = round(if_val, 2)

        loc = a.get("locationName")
        if loc: item["ort"] = loc

        steps = a.get("steps")
        if steps is not None: item["schritte"] = int(steps)

        im_m = a.get("moderateIntensityMinutes")
        if im_m is not None: item["int_mod"] = int(im_m)
        iv_v = a.get("vigorousIntensityMinutes")
        if iv_v is not None: item["int_vig"] = int(iv_v)

        aret = a.get("avgRespirationRate")
        if aret is not None: item["atem"] = round(aret, 1)

        sd_diff = a.get("differenceStress")
        if sd_diff is not None: item["stress_diff"] = int(sd_diff)

        rc = a.get("lapCount")
        if rc is not None: item["runden"] = int(rc)

        he_max = a.get("maxElevation")
        if he_max is not None: item["hoehe_max"] = int(he_max)

        is_pr = a.get("isPR")
        if is_pr is True: item["pr"] = True

        aus.append(item)
    return aus


def tag_details(von: str, bis: str, extras_heute=False) -> dict:
    """{datum: data} für alle Tage von-von bis bis, mit allen Quellen."""
    today_str = datetime.datetime.now(TZ).date().isoformat()

    dailies = _dailysummaries(von, bis)
    sleeps = _sleep(von, bis)
    acts = _activities_block(von, bis)

    hrvs = {}
    try:
        hrvs = _hrv(von, bis)
    except Exception:
        pass

    gewicht = {}
    try:
        gewicht = _gewicht(von, bis)
    except Exception:
        pass

    all_dates = set()
    for cd in dailies: all_dates.add(cd)
    for cd in sleeps: all_dates.add(cd)
    seen_days = set()
    for a in acts:
        da = tag_aus(a)
        if da: seen_days.add(da)
    for cd in seen_days: all_dates.add(cd)

    result = {}
    for datum in sorted(all_dates):
        tag_data = _dailysummary_day(dailies, datum) or {}
        schlaf_data = _sleep_day(sleeps, datum)
        akt_data = _activity_day(acts, datum)
        hrv_data = _hrv_day(hrvs, datum) or {}

        has_tag = bool(tag_data)
        has_schlaf = schlaf_data is not None
        has_akt = akt_data is not None

        if not (has_tag or has_schlaf or has_akt):
            continue

        data = {"v": 1}  # Fix Cloud-Review: Versionskennung auf oberster Ebene
        if has_tag:
            if tag_data.get("v") == 1 or all(k != "v" for k in tag_data):
                data["tag"] = {k: v for k, v in tag_data.items() if v is not None}
        if has_schlaf:
            data["schlaf"] = {k: v for k, v in schlaf_data.items() if v is not None}
        if hrv_data:
            data["hrv"] = {k: v for k, v in hrv_data.items() if k != "v" or v is not None}

        # Gewicht
        g = gewicht.get(datum)
        if g is not None: data["gewicht_kg"] = g

        # Fitness Alter und BB aufwachen nur für HEUTE (wenn extras_heute=True)
        if datum == today_str and extras_heute:
            fa = _fitnessalter(datum)
            if fa is not None: data["fitnessalter"] = fa
            bbw = _bb_aufwachen(datum)
            if bbw is not None: data["bb_aufwachen"] = bbw

        # Aktivitäten und Summe
        if has_akt:
            data["akt"] = akt_data
            sum_entry = {"n": 0, "km": 0.0, "hm": 0, "dauer_s": 0, "bewegt_s": 0,
                         "last": 0.0, "kcal": 0, "bb": 0, "zonen_s": [0] * 5}
            for act in akt_data:
                sum_entry["n"] += 1
                sum_entry["km"] += act.get("km", 0)
                sum_entry["hm"] += act.get("hm", 0)
                sum_entry["dauer_s"] += act.get("dauer_s", 0)
                sum_entry["bewegt_s"] += act.get("bewegt_s", 0)
                l = act.get("last", 0)
                if isinstance(l, (int, float)): sum_entry["last"] += l
                sum_entry["kcal"] += act.get("kcal", 0)
                sum_entry["bb"] += act.get("bb", 0)
                zones = act.get("zonen_s")
                if isinstance(zones, list):
                    for i in range(min(len(zones), 5)):
                        sum_entry["zonen_s"][i] += zones[i]

            # Runde km und hm
            sum_entry["km"] = round(sum_entry["km"], 2)
            sum_entry["hm"] = int(round(sum_entry["hm"]))

            data["summe"] = {k: v for k, v in sum_entry.items()}

        # Clean: keine None-Werte
        def _clean(d):
            cleaned = {}
            for k, v in d.items():
                if v is not None:
                    if isinstance(v, dict):
                        cleaned[k] = {kk: vv for kk, vv in v.items() if vv is not None}
                        if cleaned[k] == {}: pass  # keep empty dicts if no keys survived? no, skip
                    else:
                        cleaned[k] = v
            return cleaned

        data = _clean(data)
        result[datum] = data

    return result


def text_fuer(data: dict) -> str:
    """Kurztext für einen Tag aus data."""
    parts = []

    # Anzahl Aktivitäten, km, hm, Dauer
    akt = data.get("akt", [])
    n = len(akt) if isinstance(akt, list) else 0
    sm = data.get("summe") or {}
    if isinstance(akt, list) and len(akt) > 0:
        km_val = sm.get("km", 0) if isinstance(sm, dict) else 0
        hm_val = sm.get("hm", 0) if isinstance(sm, dict) else 0
        dur = sm.get("dauer_s", 0) if isinstance(sm, dict) else 0
    elif "tag" in data:
        km_val = 0.0
        hm_val = 0
        dur = 0
    else:
        km_val = 0.0
        hm_val = 0
        dur = 0

    if n > 0:
        parts.append(f"{n} Akt")
        parts.append(f"{km_val:.1f} km".replace(".", ","))
        if hm_val > 0: parts.append(f"{hm_val} hm")
        # Dauer als H:MM oder MM
        if dur > 0:
            h = dur // 3600
            m = (dur % 3600) // 60
            if h > 0:
                parts.append(f"{h}:{m:02d} h")
            else:
                parts.append(f"{m} min")

    # Schlaf
    schlaf = data.get("schlaf", {})
    if isinstance(schlaf, dict) and "dauer_s" in schlaf:
        ds = schlaf["dauer_s"]
        sh = ds // 3600
        sm2 = (ds % 3600) // 60
        parts.append(f"Schlaf {sh}:{sm2:02d}")

    # BB
    tag_data = data.get("tag", {})
    bb_min = tag_data.get("bb_min") if isinstance(tag_data, dict) else None
    bb_max = tag_data.get("bb_max") if isinstance(tag_data, dict) else None
    if bb_min is not None and bb_max is not None:
        parts.append(f"BB {bb_min}-{bb_max}")

    return " · ".join(parts) if parts else ""


def send_tage(tage: dict, nur_geaendert=True, state=None):
    """tage = {datum: data}. Sende via saveExtra für garmin_tag.
    
    Mit state=None (Default): lädt/sichert State, printet Ergebnis.
    Mit state=...: updates state["tag_hash"], gibt (gesendet, fehlert) zurück.
    """
    if state is None:
        state = _laden_state()
        load_save = True
    else:
        load_save = False

    tag_hash = state.get("tag_hash", {}) if isinstance(state, dict) else {}
    gesendet = 0
    fehlert = 0

    for datum in sorted(tage.keys()):
        data = tage[datum]
        h = hashlib.sha1(json.dumps(data, sort_keys=True).encode()).hexdigest()

        if nur_geaendert:
            old_h = tag_hash.get(datum, "")
            if h == old_h and old_h:
                continue

        payload = {
            "kind": "saveExtra",
            "typ": "garmin_tag",
            "datum": datum,
            "text": text_fuer(data),
            "data": data,
            "quelle": "garmin-api",
        }
        resp = http_post_json(SCRIPT_URL, payload)
        if not resp.get("ok"):
            err = resp.get("error", "unbekannt")
            log(f"[garmin] send_tage({datum}): {err}")
            fehlert += 1
        else:
            tag_hash[datum] = h
            gesendet += 1

        # Pause 0.3s zwischen Aufrufen (nicht nach letztem)
        if datum != sorted(tage.keys())[-1]:
            time.sleep(0.3)

    # tag_hash speichern, auf 30 neueste kürzen
    state["tag_hash"] = dict(list(tag_hash.items())[-30:])
    if load_save:
        _atomar_schreiben(STATE_PATH, state)

    if state is None:
        print(f"Tage gesendet: {gesendet}, Fehler: {fehlert}")
    return gesendet, fehlert


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


# Takt: LaunchAgent startet alle 10 min. Morgens (04:30–10:00) wird jedes Mal abgerufen,
# sonst nur, wenn der letzte Abruf mindestens 28 min her ist. Nachts (00:00–04:29) gar nicht.
# --force (Knopf „Garmin jetzt abrufen“ im Cockpit) ignoriert Fenster und Takt.
MORGEN_VON = (4, 30)
MORGEN_BIS = (10, 0)
TAKT_SONST_MIN = 28


def cmd_tick(force: bool = False):
    """Takt/Fenster prüfen, doppelte Läufe sperren, dann abrufen."""
    import fcntl
    now = datetime.datetime.now(TZ)
    hm = (now.hour, now.minute)
    if not force:
        if hm < MORGEN_VON:
            return  # nachts still – kein Log-Rauschen
        if not (MORGEN_VON <= hm < MORGEN_BIS):
            letzter = _laden_state().get("letzter_tick")
            try:
                alt = (now - datetime.datetime.fromisoformat(letzter)).total_seconds() / 60 if letzter else 1e9
            except Exception:
                alt = 1e9
            if alt < TAKT_SONST_MIN:
                return
    lock = open(STATE_PATH.parent / "garmin_tick.lock", "w")
    try:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        log("[garmin] tick: läuft bereits – übersprungen")
        return
    try:
        log(f"[garmin] tick{' (auf Anfrage)' if force else ''} {now.strftime('%H:%M')}")
        _cmd_tick_inner()
    finally:
        fcntl.flock(lock, fcntl.LOCK_UN)
        lock.close()


def _cmd_tick_inner():
    """Aktuellen Stand abrufen und Vorschlag für heute+gestern senden."""
    now = datetime.datetime.now(TZ)
    state = _laden_state()
    state["letzter_tick"] = now.isoformat(timespec="seconds")
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

    # Ablauf-Erinnerung: höchstens einmal pro Tag
    if ablauf_gemeldet != heute_str:
        try:
            key_expiry = state.get("key_expiry")
            if key_expiry:
                exp_date = datetime.date.fromisoformat(key_expiry)
                now_date = now.date()
                diff = (exp_date - now_date).days
                if 0 <= diff <= 7:
                    formatted = exp_date.strftime("%d.%m.%Y")
                    notify("Coach Kira", f"Garmin-Schlüssel läuft am {formatted} ab")
                    state["ablauf_gemeldet"] = heute_str
        except Exception:
            pass

    # --- Garmin Tag: Details für gestern+heute abrufen und senden ---
    try:
        yesterday_str = (datetime.datetime.now(TZ) - datetime.timedelta(days=1)).date().isoformat()
        today_str = datetime.datetime.now(TZ).date().isoformat()

        tage_data = tag_details(yesterday_str, today_str, extras_heute=True)
        if not tage_data:
            log("[garmin] tick tag_details: keine Tage gefunden")

        send_tage(tage_data, nur_geaendert=True, state=state)
    except Exception as e:
        log(f"[garmin] tick tag_details/send_tage: {e}")

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
    """Ablaufdatum des Schlüssels im STATE speichern."""
    state = _laden_state()
    state["key_expiry"] = bis
    try:
        exp_dt = datetime.date.fromisoformat(bis)
        state["key_expiry_formatted"] = exp_dt.strftime("%d.%m.%Y")
    except Exception:
        pass
    _atomar_schreiben(STATE_PATH, state)
    log(f"[garmin] Schlüssel läuft am {bis} ab")


def cmd_details(von=None, bis=None, probe=False):
    """Tag-Details für einen Zeitraum abrufen und optional senden."""
    today = datetime.datetime.now(TZ).date().isoformat()
    start_date = START_TIMELINE  # Vorlage für --von
    
    if von:
        start_date = von
    if bis:
        end_date = bis
    else:
        end_date = today

    tage = tag_details(start_date, end_date)
    if not tage:
        log("[garmin] details: keine Tage gefunden")
        return

    if probe:
        # Ausgabe Beispiel (erster Tag mit akt) und Anzahl Tage
        example_day = None
        for d, data in sorted(tage.items()):
            if "akt" in data:
                example_day = d
                break
        
        print(f"Tage gesamt: {len(tage)}")
        if example_day:
            print(f"Beispiel ({example_day}): {json.dumps(tage[example_day], ensure_ascii=False, indent=2, default=str)[:600]}")
        print(f"Texte:")
        for d, data in sorted(tage.items()):
            t = text_fuer(data)
            print(f"  {d}: {t}")
    else:
        send_tage(tage, nur_geaendert=False)


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
    p_tick = subparsers.add_parser("tick", help="Heute+Gestern Abruf")
    p_tick.add_argument("--force", action="store_true", help="Fenster/Takt ignorieren (Abruf auf Anfrage)")

    # vorschau
    p_vorschau = subparsers.add_parser("vorschau", help="Vorschlag als JSON")
    p_vorschau.add_argument("--datum", default=None)

    # ablauf
    p_ablauf = subparsers.add_parser("ablauf", help="Ablaufdatum speichern")
    p_ablauf.add_argument("--bis", required=True)

    # details
    p_details = subparsers.add_parser("details", help="Tag-Details abrufen/senden")
    p_details.add_argument("--von", default=None)
    p_details.add_argument("--bis", default=None)
    p_details.add_argument("--probe", action="store_true")

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
            cmd_tick(force=bool(getattr(args, "force", False)))
        elif args.command == "vorschau":
            cmd_vorschau(datum=args.datum)
        elif args.command == "ablauf":
            cmd_ablauf(bis=args.bis)
        elif args.command == "details":
            cmd_details(von=args.von, bis=args.bis, probe=bool(args.probe))
    except Exception as e:
        log(f"[garmin] FATAL: {e}")
        raise


if __name__ == "__main__":
    main()
