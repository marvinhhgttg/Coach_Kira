"""Coach Kira Runner – Hintergrunddienst auf dem Mac.

Erzeugt morgens die Coach-Kira-Briefings mit dem lokalen Ollama-Modell
und speichert sie im Google-Apps-Script.

NUR Standardbibliothek: urllib, json, datetime, zoneinfo, pathlib,
subprocess, sys, time, argparse.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
import socket

# IPv4 bevorzugen: über IPv6 hängen Anfragen an Google hier 30-50 s (IPv4: ~2,5 s)
_orig_getaddrinfo = socket.getaddrinfo


def _gai_v4(host, port, family=0, *args, **kwargs):
    if family in (0, socket.AF_UNSPEC):
        try:
            return _orig_getaddrinfo(host, port, socket.AF_INET, *args, **kwargs)
        except socket.gaierror:
            pass
    return _orig_getaddrinfo(host, port, family, *args, **kwargs)


socket.getaddrinfo = _gai_v4
from zoneinfo import ZoneInfo

# ---------------------------------------------------------------------------
# Konstanten
# ---------------------------------------------------------------------------
SCRIPT_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec"
)
OLLAMA = "http://127.0.0.1:11434"
MODEL = "coach-kira"
TZ = ZoneInfo("Europe/Berlin")
BASE = Path.home() / "coach-kira-runner"
STATE = BASE / "state.json"
LOG = BASE / "logs" / "runner.log"
MORNING_TYPES = ["status", "plan", "consolidated"]
WINDOW = (5, 12)  # Stunden, lokale Zeit [5, 12)
MAX_TRIES = 3  # Fehlversuche pro Tag, danach Ruhe bis morgen

# Modell-Routing: welcher Typ nutzt welches Modell
ROUTES: dict[str, str] = {
    "activity": "coach-kira",  # Test 18.10.: e4b 2,4 s aber Zahlenfehler, coach-kira 3,5 s korrekt
    "status": "coach-kira",
    "plan": "coach-kira",
    "consolidated": "coach-kira",
    "supervisor": "coach-kira",
}
FALLBACK_MODEL = "coach-kira"

# Tagesdaten-Metriken (für fetch_days)
ACT_METRICS = (
    "activity_done,load_fb_day,coachE_ESS_day,Sport_x,Zone,Aerobic_TE,"
    "Anaerobic_TE,Target_Aerobic_TE,Target_Anaerobic_TE,fbATL_obs,"
    "fbCTL_obs,fbACWR_obs,coachE_ATL_forecast,coachE_CTL_forecast,"
    "coachE_ACWR_forecast,Garmin_Training_Readiness"
)

# Activity-Prompt
ACTIVITY_SYSTEM = (
    "Du bist Coach Kira, eine erfahrene Sportwissenschaftlerin. Du sprichst Marc direkt per Du an. "
    "Antworte nur mit dem Kommentartext, ohne Anführungszeichen, ohne JSON, ohne Einleitung."
)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------


def log(msg: str) -> None:
    """Zeile 'YYYY-MM-DD HH:MM:SS msg' an LOG anhängen und auf stdout drucken."""
    LOG.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} {msg}\n"
    with LOG.open("a", encoding="utf-8") as fh:
        fh.write(line)
    print(line, end="", flush=True)


# ---------------------------------------------------------------------------
# HTTP-Hilfen
# ---------------------------------------------------------------------------

def http_get_json(url: str, timeout: int = 60) -> dict:
    """urllib GET, User-Agent 'coach-kira-runner', folgt Redirects, json.loads."""
    req = Request(url, headers={"User-Agent": "coach-kira-runner"})
    with urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode("utf-8")
    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Antwort ist kein JSON (erste 200 Zeichen: {body[:200]!r})"
        ) from exc


def http_post_json(url: str, obj: object, timeout: int = 120) -> dict:
    """POST mit Body json.dumps(obj, ensure_ascii=False).encode('utf-8').

    Apps Script antwortet auf POST mit 302 auf eine andere URL, die per GET
    geholt werden muss. urllib macht das bei 302 automatisch (POST -> GET
    ohne Body) – so lassen. Antwort json.loads.

    timeout: Sekunden bis zum Schließen der Verbindung (default 120).
    """
    data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    req = Request(
        url,
        data=data,
        headers={
            "Content-Type": "text/plain;charset=utf-8",
            "User-Agent": "coach-kira-runner",
        },
        method="POST",
    )
    with urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode("utf-8")
    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Antwort ist kein JSON (erste 200 Zeichen: {body[:200]!r})"
        ) from exc


# ---------------------------------------------------------------------------
# Ollama
# ---------------------------------------------------------------------------


def ollama_ready(model: str = MODEL) -> bool:
    """GET OLLAMA/api/tags (timeout 5); True, wenn ein Modell mit name == model
    oder name beginnt mit model + ':'. Fehler -> False."""
    try:
        req = Request(OLLAMA + "/api/tags", headers={"User-Agent": "coach-kira-runner"})
        with urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        models = data.get("models") or []
        for m in models:
            name = m.get("name") or ""
            if name == model or name.startswith(model + ":"):
                return True
        return False
    except Exception:
        return False


def ollama_chat(
    model: str,
    messages: list[dict],
    num_ctx: int,
    fmt: str | None = None,
    timeout: int = 1200,
) -> tuple[str, float | None]:
    """POST OLLAMA/api/chat mit stream=False, think=False, keep_alive='30m'.

    Gibt (message.content.strip(), eval_count/(eval_duration/1e9) oder None) zurück.
    """
    body: dict = {
        "model": model,
        "stream": False,
        "think": False,
        "keep_alive": "30m",
        "options": {"temperature": 0.3, "num_ctx": num_ctx},
        "messages": messages,
    }
    if fmt is not None:
        body["format"] = fmt
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = Request(
        OLLAMA + "/api/chat",
        data=data,
        headers={
            "Content-Type": "application/json",
            "User-Agent": "coach-kira-runner",
        },
        method="POST",
    )
    with urlopen(req, timeout=timeout) as resp:
        j = json.loads(resp.read().decode("utf-8"))
    text = (j.get("message") or {}).get("content")
    if not text:
        raise RuntimeError("Ollama antwortete ohne Text")
    text = text.strip()
    if not text:
        raise RuntimeError("Ollama antwortete mit leerem Text")
    eval_count = j.get("eval_count")
    eval_duration = j.get("eval_duration")
    tok_s: float | None = None
    if eval_count is not None and eval_duration is not None:
        try:
            tok_s = eval_count / (eval_duration / 1e9)
        except Exception:
            pass
    return text, tok_s


# ---------------------------------------------------------------------------
# Tagesdaten
# ---------------------------------------------------------------------------


def fetch_days(days: int) -> list[dict]:
    """http_get_json(SCRIPT_URL + f'?mode=chartData&range={days}d&metrics=' + ACT_METRICS)['data']."""
    url = SCRIPT_URL + f"?mode=chartData&range={days}d&metrics=" + ACT_METRICS
    return http_get_json(url).get("data") or []


def num(v: object) -> float | None:
    """Zahl aus int/float/String mit Komma; '' oder None -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    try:
        return float(s.replace(",", "."))
    except (ValueError, TypeError):
        return None


def is_activity(row: dict) -> bool:
    """str(row.get('activity_done','')).strip().lower() == 'x' und (num(load_fb_day)
    oder num(coachE_ESS_day) oder 0) > 0 und Sport_x (klein) nicht in ('off','krank','')."""
    if str(row.get("activity_done", "")).strip().lower() != "x":
        return False
    load = num(row.get("load_fb_day")) or num(row.get("coachE_ESS_day")) or 0.0
    if load <= 0:
        return False
    sport = str(row.get("Sport_x") or "").strip().lower()
    if sport in ("off", "krank", ""):
        return False
    return True


def ist_load(row: dict) -> float | None:
    """num(load_fb_day) falls > 0, sonst num(coachE_ESS_day)."""
    v = num(row.get("load_fb_day"))
    if v is not None and v > 0:
        return v
    return num(row.get("coachE_ESS_day"))


# ---------------------------------------------------------------------------
# Plan-Schnappschuss (SOLL ist nach dem Eintragen im Sheet überschrieben,
# deshalb morgens merken)
# ---------------------------------------------------------------------------


def snapshot_plan(
    st: dict, rows: list[dict], today: str
) -> bool:
    """Zeile mit date == today suchen; wenn vorhanden und activity_done != 'x':
    st.setdefault('plan_snapshot', {})[today] = {...}; danach nur die 14 neuesten Schlüssel behalten.
    Rückgabe True wenn geändert."""
    changed = False
    for row in rows:
        if row.get("date") != today:
            continue
        if str(row.get("activity_done", "")).strip().lower() == "x":
            continue
        snap: dict = st.setdefault("plan_snapshot", {})
        snap[today] = {
            "load": num(row.get("coachE_ESS_day")),
            "sport": row.get("Sport_x"),
            "zone": row.get("Zone"),
            "atl": num(row.get("coachE_ATL_forecast")),
            "ctl": num(row.get("coachE_CTL_forecast")),
            "acwr": num(row.get("coachE_ACWR_forecast")),
            "te_ae": num(row.get("Target_Aerobic_TE")),
            "te_an": num(row.get("Target_Anaerobic_TE")),
        }
        changed = True
        break
    # nur die 14 neuesten Schlüssel behalten
    snap = st.get("plan_snapshot")
    if snap:
        keys = list(snap.keys())
        if len(keys) > 14:
            st["plan_snapshot"] = dict(list(snap.items())[-14:])
    return changed


# ---------------------------------------------------------------------------
# Activity-Prompt
# ---------------------------------------------------------------------------


def build_activity_messages(row: dict, soll: dict | None) -> list[dict]:
    """user-Text (deutsch) mit Abschnitten: PLANVORGABE FÜR HEUTE, TATSÄCHLICH DURCHGEFÜHRT,
    AUFGABE (MAX. 3 SÄTZE), REGELN. Rückgabe [{'role':'system',...},{'role':'user',...}]."""
    soll_load = soll.get("load") if soll else None
    soll_sport = soll.get("sport") if soll else None
    soll_zone = soll.get("zone") if soll else None
    soll_te_ae = soll.get("te_ae") if soll else num(row.get("Target_Aerobic_TE"))
    soll_te_an = soll.get("te_an") if soll else num(row.get("Target_Anaerobic_TE"))

    # PLANVORGABE FÜR HEUTE:
    if soll_load is not None:
        plan_load = f"{soll_load:.0f}"
    else:
        plan_load = "unbekannt"
    plan_sport = str(soll_sport or "unbekannt")
    plan_zone = str(soll_zone or "unbekannt")
    if soll_te_ae is not None:
        plan_te_ae = f"{soll_te_ae:.1f}"
    else:
        plan_te_ae = "unbekannt"
    if soll_te_an is not None:
        plan_te_an = f"{soll_te_an:.1f}"
    else:
        plan_te_an = "unbekannt"

    # TATSÄCHLICH DURCHGEFÜHRT:
    date = row.get("date")
    sport_x = str(row.get("Sport_x") or "")
    zone = str(row.get("Zone") or "")
    load = ist_load(row)
    aer = num(row.get("Aerobic_TE"))
    ana = num(row.get("Anaerobic_TE"))
    atl = num(row.get("fbATL_obs"))
    ctl = num(row.get("fbCTL_obs"))
    acwr = num(row.get("fbACWR_obs"))

    text = (
        f"PLANVORGABE FÜR HEUTE:\n"
        f"  Geplanter Load (ESS): {plan_load}\n"
        f"  Geplante Sportart: {plan_sport}\n"
        f"  Geplante Zone: {plan_zone}\n"
        f"  Ziel-TE aerob: {plan_te_ae}\n"
        f"  Ziel-TE anaerob: {plan_te_an}\n\n"
        f"TATSÄCHLICH DURCHGEFÜHRT:\n"
        f"  Datum: {date}\n"
        f"  Aktivität (Sport_x): {sport_x}\n"
        f"  Zone(n): {zone}\n"
        f"  Load: {load}\n"
        f"  TE aerob: {aer}\n"
        f"  TE anaerob: {ana}\n"
        f"  ATL: {atl}\n"
        f"  CTL: {ctl}\n"
        f"  ACWR: {acwr}\n\n"
        f"AUFGABE (MAX. 3 SÄTZE):\n"
        f"1. Vergleiche SOLL mit IST.\n"
        f"2. Bewerte die Abweichung (positiv, negativ, neutral).\n"
        f"3. Gib einen kurzen, motivierenden Kommentar mit Blick auf morgen.\n\n"
        f"REGELN: Nur Werte aus diesen Daten verwenden, nichts erfinden. "
        f"Wenn die Planvorgabe unbekannt ist, nur die TE-Ziele vergleichen und das kurz sagen."
    )
    return [
        {"role": "system", "content": ACTIVITY_SYSTEM},
        {"role": "user", "content": text},
    ]


def sane(text: str) -> bool:
    """40 <= len(text) <= 900 und nicht text.lstrip().startswith('{') und any(ch.isdigit() for ch in text)."""
    if not (40 <= len(text) <= 900):
        return False
    if text.lstrip().startswith("{"):
        return False
    if not any(ch.isdigit() for ch in text):
        return False
    return True


def clean(text: str) -> str:
    """äußere Anführungszeichen ('" „ “ ') entfernen, strip."""
    s = text.strip()
    if len(s) >= 2:
        if (s[0] == '"' and s[-1] == '"') or (s[0] == "“" and s[-1] == "”") or (s[0] == "„" and s[-1] == "”"):
            s = s[1:-1].strip()
    return s


def review_activity(row: dict, soll: dict | None) -> tuple[str, float]:
    """model = ROUTES['activity']; wenn nicht ollama_ready(model): model = FALLBACK_MODEL.
    text, tok = ollama_chat(model, build_activity_messages(row, soll), 8192, timeout=600); text = clean(text).
    Wenn nicht sane(text) und model != FALLBACK_MODEL: log('  Kommentar von <model> unbrauchbar – nehme coach-kira'); model = FALLBACK_MODEL; erneut.
    Wenn immer noch nicht sane: RuntimeError('Kommentar unbrauchbar').
    Speichern: http_post_json(SCRIPT_URL + '?mode=saveActivityReview', {...}); bei nicht ok RuntimeError(error).
    Zeiten und tok/s loggen. Rückgabe (model, sekunden)."""
    model = ROUTES["activity"]
    if not ollama_ready(model):
        model = FALLBACK_MODEL
    messages = build_activity_messages(row, soll)
    start = time.time()
    text, tok_s = ollama_chat(model, messages, 8192, timeout=600)
    text = clean(text)
    sec = time.time() - start
    if tok_s is not None:
        log(f"  {model}: {tok_s:.0f} tok/s")
    if not sane(text) and model != FALLBACK_MODEL:
        log(f"  Kommentar von {model} unbrauchbar – nehme coach-kira")
        model = FALLBACK_MODEL
        messages = build_activity_messages(row, soll)
        start = time.time()
        text, tok_s = ollama_chat(model, messages, 8192, timeout=600)
        text = clean(text)
        sec = time.time() - start
        if tok_s is not None:
            log(f"  {model}: {tok_s:.0f} tok/s")
    if not sane(text):
        raise RuntimeError("Kommentar unbrauchbar")
    ist = {
        "load": ist_load(row),
        "atl": num(row.get("fbATL_obs")),
        "ctl": num(row.get("fbCTL_obs")),
        "acwr": num(row.get("fbACWR_obs")),
    }
    soll_d: dict = {}
    if soll:
        soll_d = {
            "load": soll.get("load"),
            "atl": soll.get("atl"),
            "ctl": soll.get("ctl"),
            "acwr": soll.get("acwr"),
        }
    saved = http_post_json(
        SCRIPT_URL + "?mode=saveActivityReview",
        {
            "kind": "saveActivityReview",
            "date": row["date"],
            "text": text,
            "model": model,
            "ist": ist,
            "soll": soll_d,
        },
    )
    if not saved.get("ok"):
        raise RuntimeError(saved.get("error") or "saveActivityReview fehlgeschlagen")
    return model, sec


# ---------------------------------------------------------------------------
# Activity-Befehl
# ---------------------------------------------------------------------------


def activity_patch_ready() -> bool:
    """True, wenn das Apps Script den Modus activityReviews kennt (Patch eingespielt)."""
    try:
        j = http_get_json(SCRIPT_URL + "?mode=activityReviews&limit=1")
        return bool(j.get("ok")) and isinstance(j.get("reviews"), list)
    except Exception:
        return False


def cmd_activity() -> None:
    now = datetime.now(TZ)
    if not (6 <= now.hour < 24):
        return  # still beenden (kein Log)
    st = load_state()
    today0 = now.strftime("%Y-%m-%d")
    rows0 = fetch_days(3)
    if snapshot_plan(st, rows0, today0):
        save_state(st)  # Plan-Soll auch ohne Patch schon merken
    if not activity_patch_ready():
        if st.get("patch_wait_log") != now.strftime("%Y-%m-%d"):
            log("Aktivitäten: Apps-Script-Patch (activityReviews) fehlt noch – nichts gespeichert")
            st["patch_wait_log"] = now.strftime("%Y-%m-%d")
            save_state(st)
        return
    rows = fetch_days(3)
    today = now.strftime("%Y-%m-%d")
    if snapshot_plan(st, rows, today):
        save_state(st)
    # Gestern und heute
    yesterday = (now - timedelta(days=1)).strftime("%Y-%m-%d")
    dates = (yesterday, today)
    # Beim allerersten Lauf (st hat noch kein 'reviewed'): Gestern-Keys direkt als reviewed eintragen
    if "reviewed" not in st:
        for row in rows:
            if row.get("date") == yesterday and is_activity(row):
                key = f"{row['date']}:{ist_load(row):.0f}"
                st.setdefault("reviewed", []).append(key)
        save_state(st)
    for row in rows:
        date = row.get("date")
        if date not in dates:
            continue
        if not is_activity(row):
            continue
        key = f"{date}:{ist_load(row):.0f}"
        if key in st.get("reviewed", []):
            continue
        if st.get("act_fail", {}).get(key, 0) >= 3:
            continue  # dreimal gescheitert – nicht weiter versuchen
        soll = st.get("plan_snapshot", {}).get(date)
        try:
            model, sec = review_activity(row, soll)
            log(f"Aktivität {date}: kommentiert mit {model} in {sec:.0f} s")
            st["reviewed"] = (st.get("reviewed", []) + [key])[-60:]
            save_state(st)
            notify("Coach Kira", "Kira hat deine Einheit vom " + date + " kommentiert.")
        except Exception as e:
            log(f"Aktivität {date}: FEHLER {e}")
            fails = st.get("act_fail", {})
            fails[key] = fails.get(key, 0) + 1
            st["act_fail"] = fails
            save_state(st)


# ---------------------------------------------------------------------------
# Backfill-Befehl
# ---------------------------------------------------------------------------


def cmd_backfill(days: int, dry: bool) -> None:
    if not activity_patch_ready():
        raise RuntimeError("Apps-Script-Patch (activityReviews) fehlt noch")
    reviews = http_get_json(SCRIPT_URL + "?mode=activityReviews").get("reviews", [])
    vorhandene = {r["date"] for r in reviews}
    rows = fetch_days(days)
    heute = datetime.now(TZ).strftime("%Y-%m-%d")
    gezählt = 0
    for row in rows:
        date = row.get("date")
        if not is_activity(row):
            continue
        if date in vorhandene or date >= heute:
            continue
        gezählt += 1
        if dry:
            print(date, row.get("Sport_x"), ist_load(row))
        else:
            try:
                model, sec = review_activity(row, None)
                log(f"Backfill {date}: kommentiert mit {model} in {sec:.0f} s")
            except Exception as e:
                log(f"Backfill {date}: FEHLER {e}")
    log(f"Backfill: {gezählt} Aktivität(en) bearbeitet")


# ---------------------------------------------------------------------------
# Tick-Befehl
# ---------------------------------------------------------------------------


def cmd_supervisor() -> None:
    """Supervisor-Aufträge abholen; wurde ein Auftrag während der Rechnung ersetzt,
    sofort den neueren rechnen (max. 3 Runden)."""
    import fcntl
    lock = open(Path.home() / "coach-kira-runner" / "supervisor.lock", "w")
    try:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        return  # läuft schon – nicht doppelt rechnen
    try:
        for _ in range(3):
            if _supervisor_once() != "superseded":
                return
            log("Supervisor: neuerer Auftrag vorhanden – rechne ihn sofort")
    finally:
        fcntl.flock(lock, fcntl.LOCK_UN)
        lock.close()


def _supervisor_once() -> str:
    """Einen Supervisor-Auftrag abholen und lokal mit Ollama rechnen."""
    st = load_state()
    try:
        j = http_get_json(SCRIPT_URL + "?mode=supervisorJob", timeout=60)
    except RuntimeError:
        return "none"  # kein JSON = Patch fehlt, still beenden
    job = j.get("job")
    if not job:
        return "none"  # still beenden
    if job["id"] in st.get("sup_done", []):
        return "none"  # schon erledigt
    fails = st.get("sup_fail", {})
    if fails.get(job["id"], 0) >= 2:
        return "none"  # zu oft gescheitert
    model = ROUTES["supervisor"]
    if not ollama_ready(model):
        log("Supervisor: Ollama/coach-kira nicht erreichbar – später erneut")
        return "none"
    log(f"Supervisor: Auftrag {job['id']} übernommen (Prompt {len(job.get('prompt', ''))} Zeichen)")
    messages = [
        {"role": "system", "content": job.get("system") or ""},
        {"role": "user", "content": job["prompt"]},
    ]
    try:
        t0 = time.time()
        text, tok = ollama_chat(model, messages, 32768, fmt=job.get("schema") or "json", timeout=1800)
        if tok is not None:
            log(f"  Supervisor: {model} {time.time() - t0:.0f} s, {tok:.0f} tok/s")
        else:
            log(f"  Supervisor: {model} {time.time() - t0:.0f} s")
        text = json_clean(text)  # Codeblock/Vortext entfernen; muss gültiges JSON sein, sonst Exception
        saved = http_post_json(
            SCRIPT_URL + "?mode=saveSupervisor",
            {"kind": "saveSupervisor", "id": job["id"], "content": text, "model": model},
            timeout=600,
        )
        if not saved.get("ok"):
            err = str(saved.get("error") or "saveSupervisor fehlgeschlagen")
            if "neuerer Auftrag" in err:
                log(f"Supervisor: Ergebnis zu {job['id']} verworfen (inzwischen neu gestartet)")
                return "superseded"
            raise RuntimeError(err)
        st["sup_done"] = (st.get("sup_done", []) + [job["id"]])[-30:]
        save_state(st)
        log(f"Supervisor: Auftrag {job['id']} lokal erledigt und gespeichert")
        notify("Coach Kira", "KI-Report lokal erstellt (coach-kira).")
        return "ok"
    except Exception as e:
        fails[job["id"]] = fails.get(job["id"], 0) + 1
        st["sup_fail"] = fails
        save_state(st)
        log(f"Supervisor: FEHLER {e}")
        if fails[job["id"]] >= 2:
            notify("Coach Kira", "KI-Report lokal fehlgeschlagen – siehe runner.log")
        return "error"


def cmd_tick() -> None:
    try:
        cmd_morning()
    except Exception as e:
        log(f"morning: FEHLER {e}")
    try:
        cmd_activity()
    except Exception as e:
        log(f"activity: FEHLER {e}")
    # Supervisor-Aufträge holt der eigene LaunchAgent de.coachkira.supervisor (jede Minute)


# ---------------------------------------------------------------------------
# Morning data
# ---------------------------------------------------------------------------


def morning_data_ready(today_str: str) -> tuple[bool, dict | None]:
    """http_get_json(SCRIPT_URL + '?mode=chartData&range=3d&metrics=...').

    Suche in j['data'] den Eintrag mit date == today_str; bereit, wenn beide
    Felder vorhanden und nicht None/'' sind. Rückgabe (bool, dict_or_None).
    """
    url = (
        SCRIPT_URL
        + "?mode=chartData"
        "&range=3d"
        "&metrics=Garmin_Training_Readiness,sleep_score_0_100"
    )
    j = http_get_json(url)
    for row in (j.get("data") or []):
        if row.get("date") == today_str:
            readiness = row.get("Garmin_Training_Readiness")
            sleep = row.get("sleep_score_0_100")
            if (
                readiness is not None
                and readiness != ""
                and sleep is not None
                and sleep != ""
            ):
                return True, row
            return False, row
    return False, None


# ---------------------------------------------------------------------------
# Quellen
# ---------------------------------------------------------------------------


def sources() -> dict:
    """http_get_json(SCRIPT_URL + '?mode=briefingSources')['sources'] (bei Fehler {})."""
    try:
        return http_get_json(SCRIPT_URL + "?mode=briefingSources").get("sources") or {}
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Done-Check
# ---------------------------------------------------------------------------


def done_today(src: dict, type_: str, today_str: str) -> bool:
    """src.get(type_) hat 'timestamp' (ISO, UTC mit Z); in TZ umrechnen
    (datetime.fromisoformat nach Ersetzen von 'Z' durch '+00:00'); True, wenn
    Datum == today_str."""
    entry = src.get(type_)
    if not entry:
        return False
    ts = entry.get("timestamp")
    if not ts:
        return False
    iso = ts.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(iso)
    except (ValueError, TypeError):
        return False
    return dt.astimezone(TZ).strftime("%Y-%m-%d") == today_str


# ---------------------------------------------------------------------------
# Briefing erzeugen
# ---------------------------------------------------------------------------


def run_briefing(type_: str) -> float:
    """Briefing für 'type_' erzeugen und speichern. Rückgabe Sekunden (float)."""
    p = http_get_json(SCRIPT_URL + "?mode=briefingPrompt&type=" + type_)
    if not p.get("ok"):
        raise RuntimeError(p.get("error") or "BriefingPrompt fehlgeschlagen")
    model = ROUTES.get(type_, MODEL)
    body: dict = {
        "model": model,
        "stream": False,
        "think": False,
        "keep_alive": "30m",
        "options": {"temperature": 0.3, "num_ctx": 32768},
        "messages": p["messages"],
    }
    if p.get("format") == "json":
        body["format"] = "json"
    start = time.time()
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = Request(
        OLLAMA + "/api/chat",
        data=data,
        headers={
            "Content-Type": "application/json",
            "User-Agent": "coach-kira-runner",
        },
        method="POST",
    )
    with urlopen(req, timeout=1200) as resp:
        j = json.loads(resp.read().decode("utf-8"))
    elapsed = time.time() - start
    text = (j.get("message") or {}).get("content")
    if not text:
        raise RuntimeError("Ollama antwortete ohne Text")
    text = text.strip()
    if not text:
        raise RuntimeError("Ollama antwortete mit leerem Text")
    # tok/s loggen (falls vorhanden)
    eval_count = j.get("eval_count")
    eval_duration = j.get("eval_duration")
    if eval_count is not None and eval_duration is not None:
        try:
            tok_s = eval_count / (eval_duration / 1e9)
            log(f"  Ollama: {tok_s:.0f} tok/s")
        except Exception:
            pass
    saved = http_post_json(
        SCRIPT_URL + "?mode=saveBriefing",
        {"kind": "saveBriefing", "type": type_, "text": text, "model": model},
    )
    if not saved.get("ok"):
        raise RuntimeError(saved.get("error") or "saveBriefing fehlgeschlagen")
    if type_ == "plan":
        recs = saved.get("recommendations") or []
        log(f"  Plan: {len(recs)} Empfehlung(en) gespeichert")
    return elapsed


# ---------------------------------------------------------------------------
# Benachrichtigung
# ---------------------------------------------------------------------------


def _as_str(s: str) -> str:
    """AppleScript-String: nur Anführungszeichen und Backslashes escapen (Umlaute bleiben)."""
    return '"' + str(s).replace("\\", "\\\\").replace('"', '\\"') + '"'


def json_clean(text: str) -> str:
    """Modellantwort zu gültigem JSON-Text machen: ```json-Blöcke und Vor-/Nachtext entfernen."""
    import re
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```\s*$", "", t, flags=re.S).strip()
    try:
        json.loads(t)
        return t
    except Exception:
        a, b = t.find("{"), t.rfind("}")
        if a >= 0 and b > a:
            cand = t[a:b + 1]
            json.loads(cand)
            return cand
        raise ValueError("Modellantwort ist kein JSON: " + (t[:80] or "leer"))


def notify(title: str, msg: str) -> None:
    """subprocess.run(['osascript', '-e', ...], timeout=10) in try/except."""
    try:
        subprocess.run(
            [
                "osascript",
                "-e",
                f'display notification {_as_str(msg)} with title {_as_str(title)}',
            ],
            timeout=10,
        )
    except Exception:
        pass


# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------


def load_state() -> dict:
    """JSON in STATE lesen; bei Fehler leeres dict."""
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(d: dict) -> None:
    """JSON in STATE schreiben."""
    STATE.parent.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Morning-Befehl
# ---------------------------------------------------------------------------


def cmd_morning(force: bool = False) -> None:
    now = datetime.now(TZ)
    today = now.strftime("%Y-%m-%d")
    # Zeitfenster prüfen
    if not force and not (WINDOW[0] <= now.hour < WINDOW[1]):
        return  # still beenden (kein Log)
    st = load_state()
    if not force and st.get("morning_done") == today:
        return  # still beenden
    if not ollama_ready():
        log("Ollama/coach-kira nicht erreichbar – später erneut")
        return
    ready, row = morning_data_ready(today)
    if not ready:
        hour = now.hour
        if st.get("last_wait_log") != hour:
            log("Morgenwerte für heute noch nicht da")
            st["last_wait_log"] = hour
            save_state(st)
        return
    if not force:
        tries = st.get("tries", {})
        if tries.get("date") != today:
            tries = {"date": today, "n": 0}
        if tries["n"] >= MAX_TRIES:
            return  # heute schon zu oft gescheitert – still beenden
        tries["n"] += 1
        st["tries"] = tries
        save_state(st)
    src = sources()
    ergebnisse: list[str] = []
    fehler: list[str] = []
    for t in MORNING_TYPES:
        if not force and done_today(src, t, today):
            log(f"{t}: heute schon vorhanden – übersprungen")
            continue
        try:
            sec = run_briefing(t)
            log(f"{t}: lokal erzeugt in {sec:.0f} s")
            ergebnisse.append(t)
        except Exception as e:
            log(f"{t}: FEHLER {e}")
            fehler.append(t)
    if not fehler:
        st["morning_done"] = today
        save_state(st)
    if ergebnisse:
        names: list[str] = []
        for t in ergebnisse:
            if t == "status":
                names.append("KI-Analyse")
            elif t == "plan":
                names.append("Plan-Briefing")
            elif t == "consolidated":
                names.append("Konsolidierte Analyse")
            else:
                names.append(t)
        notify("Coach Kira", "Morgen-Briefing lokal erstellt: " + ", ".join(names))
    if fehler and (force or st.get("tries", {}).get("n", 0) >= MAX_TRIES):
        notify("Coach Kira", "Morgen-Briefing: Fehler bei " + ", ".join(fehler) + " – siehe ~/coach-kira-runner/logs/runner.log")


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------


def main() -> None:
    ap = argparse.ArgumentParser(description="Coach Kira Runner")
    sub = ap.add_subparsers(dest="command")

    morning = sub.add_parser("morning", help="Morgen-Briefing erzeugen")
    morning.add_argument("--force", action="store_true", help="Zeitfenster und Status ignorieren")

    check = sub.add_parser("check", help="Nur prüfen, nichts schreiben")

    run = sub.add_parser("run", help="Einzelnes Briefing erzeugen")
    run.add_argument("types", nargs="+", choices=MORNING_TYPES)

    tick = sub.add_parser("tick", help="Morgen + Aktivität (alle 6-24 Uhr)")

    activity = sub.add_parser("activity", help="Aktivitäten von gestern und heute kommentieren")

    backfill = sub.add_parser("backfill", help="Vergangene Aktivitäten nachträglich kommentieren")
    backfill.add_argument("--days", type=int, default=14, help="Anzahl der zurückliegenden Tage (default 14)")
    backfill.add_argument("--dry", action="store_true", help="Nur anzeigen, nichts speichern")

    supervisor = sub.add_parser("supervisor", help="Supervisor-Auftrag abholen und lokal rechnen")

    args = ap.parse_args()
    if args.command == "check":
        today = datetime.now(TZ).strftime("%Y-%m-%d")
        print("Ollama-Modelle bereit:")
        for m in set(ROUTES.values()):
            print(f"  {m}: {ollama_ready(m)}")
        print("Ollama/coach-kira (default):", ollama_ready())
        ready, row = morning_data_ready(today)
        print("Morgenwerte heute:", ready, row)
        src = sources()
        for t in MORNING_TYPES:
            print(f"{t}: heute vorhanden={done_today(src, t, today)}", src.get(t))
        print("Status:", load_state())
        # letzte 3 Tage pro Zeile: date, is_activity, ist_load, Sport_x
        rows = fetch_days(3)
        print("Letzte 3 Tage (Tagesdaten):")
        for row in rows:
            date = row.get("date")
            if is_activity(row):
                print(f"  {date}: activity=True  load={ist_load(row)}  Sport_x={row.get('Sport_x')}")
            else:
                print(f"  {date}: activity=False  load={ist_load(row)}  Sport_x={row.get('Sport_x')}")
        # Supervisor-Patch prüfen
        try:
            sj = http_get_json(SCRIPT_URL + "?mode=supervisorJob")
            engine = sj.get("engine")
            job = sj.get("job")
            print(f"Supervisor-Patch: engine={engine!r}, job vorhanden={bool(job)}, job id={job.get('id') if job else None}, prompt Länge={len(job.get('prompt','')) if job else 0}")
        except Exception as e:
            print(f"Supervisor-Patch: {e}")
        return
    if args.command == "morning":
        try:
            cmd_morning(force=args.force)
        except Exception as e:
            log(f"morning: FEHLER {e}")
            sys.exit(1)
    elif args.command == "run":
        for t in args.types:
            try:
                sec = run_briefing(t)
                log(f"{t}: lokal erzeugt in {sec:.0f} s")
            except Exception as e:
                log(f"{t}: FEHLER {e}")
    elif args.command == "tick":
        try:
            cmd_tick()
        except Exception as e:
            log(f"tick: FEHLER {e}")
    elif args.command == "activity":
        try:
            cmd_activity()
        except Exception as e:
            log(f"activity: FEHLER {e}")
    elif args.command == "backfill":
        try:
            cmd_backfill(days=args.days, dry=args.dry)
        except Exception as e:
            log(f"backfill: FEHLER {e}")
    elif args.command == "supervisor":
        try:
            cmd_supervisor()
        except Exception as e:
            log(f"supervisor: FEHLER {e}")
            sys.exit(1)
    else:
        ap.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
