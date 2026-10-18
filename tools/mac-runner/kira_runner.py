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
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
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


def ollama_ready() -> bool:
    """GET OLLAMA/api/tags (timeout 5); True, wenn ein Modell mit name == MODEL
    oder name beginnt mit MODEL + ':'. Fehler -> False."""
    try:
        req = Request(OLLAMA + "/api/tags", headers={"User-Agent": "coach-kira-runner"})
        with urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        models = data.get("models") or []
        for m in models:
            name = m.get("name") or ""
            if name == MODEL or name.startswith(MODEL + ":"):
                return True
        return False
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Morning data
# ---------------------------------------------------------------------------


def morning_data_ready(today_str: str) -> tuple[bool, dict | None]:
    """http_get_json(SCRIPT_URL + '?mode=chartData&range=3d&metrics=...').

    Suche in j['data'] den Eintrag mit date == today_str; bereit, wenn beide
    Felder vorhanden und nicht None/'' sind. Rückgabe (bool, dict_or_None).
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
    """Briefing für 'type_' erzeugen und speichern. Rückgabe Sekunden (float)."""
    p = http_get_json(SCRIPT_URL + "?mode=briefingPrompt&type=" + type_)
    if not p.get("ok"):
        raise RuntimeError(p.get("error") or "BriefingPrompt fehlgeschlagen")
    body: dict = {
        "model": MODEL,
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
        {"kind": "saveBriefing", "type": type_, "text": text, "model": MODEL},
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


def notify(title: str, msg: str) -> None:
    """subprocess.run(['osascript', '-e', ...], timeout=10) in try/except."""
    try:
        subprocess.run(
            [
                "osascript",
                "-e",
                f'display notification {json.dumps(msg)} with title {json.dumps(title)}',
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

    sub.add_parser("check", help="Nur prüfen, nichts schreiben")

    run = sub.add_parser("run", help="Einzelnes Briefing erzeugen")
    run.add_argument("types", nargs="+", choices=MORNING_TYPES)

    args = ap.parse_args()
    if args.command == "check":
        today = datetime.now(TZ).strftime("%Y-%m-%d")
        print("Ollama/coach-kira bereit:", ollama_ready())
        ready, row = morning_data_ready(today)
        print("Morgenwerte heute:", ready, row)
        src = sources()
        for t in MORNING_TYPES:
            print(f"{t}: heute vorhanden={done_today(src, t, today)}", src.get(t))
        print("Status:", load_state())
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
    else:
        ap.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
