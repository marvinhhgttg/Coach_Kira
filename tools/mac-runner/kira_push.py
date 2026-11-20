"""Coach Kira – Push-Benachrichtigung (Standardbibliothek nur).

Wird aus kira_runner.py importiert.
Importierte Symbole: SCRIPT_URL, BASE, TZ, http_get_json, num, log, notify
"""

from __future__ import annotations

import datetime
import json
import subprocess
from pathlib import Path

from kira_runner import BASE, SCRIPT_URL, TZ, http_get_json, log, notify, num


def _zahl(v) -> str:
    """62.0 -> '62', 7.5 -> '7,5'"""
    x = num(v)
    if x is None:
        return "–"
    return str(int(round(x))) if abs(x - round(x)) < 0.05 else f"{x:.1f}".replace(".", ",")


def _schlaf_hmm(v) -> str:
    x = num(v)
    if x is None:
        return ""
    h = int(x)
    m = int(round((x - h) * 60))
    if m == 60:
        h, m = h + 1, 0
    return f"{h}:{m:02d}"


def kurzfassung(heute: str) -> str:
    """Kurze Morgen-Nachricht (höchstens 280 Zeichen)."""
    teile: list[str] = []
    try:
        j = http_get_json(SCRIPT_URL + "?mode=chartData&range=8d&metrics=Garmin_Training_Readiness,sleep_hours,hrv_status,hrv_threshholds,rhr_bpm")
        rows = sorted(j.get("data") or [], key=lambda r: r.get("date", ""))
        h = next((r for r in rows if r.get("date") == heute), None)
        if h:
            t = [f"Readiness {_zahl(h.get('Garmin_Training_Readiness'))}"]
            sl = _schlaf_hmm(h.get("sleep_hours"))
            if sl:
                t.append(f"Schlaf {sl}")
            hv = num(h.get("hrv_status"))
            band = str(h.get("hrv_threshholds") or "")
            if hv is not None:
                t.append(f"HRV {_zahl(hv)}" + (f" ({band.replace(';', '–')})" if ";" in band else ""))
            werte = [num(r.get("sleep_hours")) for r in rows if r.get("date", "") <= heute][-7:]
            werte = [w for w in werte if w is not None]
            if len(werte) >= 5:
                schuld = sum(7.5 - w for w in werte)
                t.append(f"Schlafschuld {schuld:.1f} h".replace(".", ",") if schuld > 0 else f"Schlafguthaben {-schuld:.1f} h".replace(".", ","))
            teile.append(" · ".join(t))
    except Exception as e:
        log(f"push: Morgenwerte nicht lesbar ({e})")
    try:
        p = next((e for e in (http_get_json(SCRIPT_URL + "?mode=planJsonV2").get("plan") or []) if e.get("date") == heute), None)
        if p and p.get("training_1"):
            teile.append(f"Heute: {p['training_1']} ({_zahl(p.get('recommended_load_ess'))} ESS {p.get('recommended_zone') or ''})".replace(" )", ")"))
    except Exception as e:
        log(f"push: Plan nicht lesbar ({e})")
    try:
        it = (http_get_json(SCRIPT_URL + "?mode=extras&typ=prognose&n=1").get("items") or [None])[0]
        if it and str(it.get("datum", ""))[:10] == heute:
            d = it.get("data") or {}
            fw = (d.get("fruehwarnung") or {}).get("stufe")
            if fw and fw != "gruen":
                teile.append(f"Frühwarnung {fw}")
            q = ((d.get("vorausblick") or {}).get("qualitaetstag") or {}).get("tag")
            if q:
                teile.append(f"Qualität {q}")
            tf = d.get("tourfenster") or []
            if tf and tf[0].get("tag"):
                teile.append(f"Tour {tf[0]['tag']}")
    except Exception as e:
        log(f"push: Vorausblick nicht lesbar ({e})")
    msg = " · ".join(teile)
    return msg if len(msg) <= 280 else msg[:279] + "…"


def _esc(s: str) -> str:
    """AppleScript-Text escapen (Backslash + Anführungszeichen)."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def imessage_senden(adresse: str, text: str) -> bool:
    """iMessage via osascript senden. True bei Erfolg."""
    a = _esc(adresse)
    t = _esc(text)
    skript = (
        'tell application "Messages"\n'
        '  set s to 1st account whose service type = iMessage\n'
        f'  set b to participant "{a}" of s\n'
        f'  send "{t}" to b\n'
        'end tell'
    )
    try:
        result = subprocess.run(
            ["osascript", "-e", skript],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            return True
    except Exception as exc:
        log(f"iMessage-Fehler: {exc}")
    log(f"iMessage fehlgeschlagen (Adresse={adresse})")
    return False


# ── 3. morgen_push ───────────────────────────────────────────

def _load_json(path: Path) -> dict | None:
    """JSON-Datei laden, bei Fehler None."""
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _save_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def morgen_push(heute: str) -> None:
    """Täglicher Push – nur einmal pro Tag."""
    state_path = Path(BASE) / "push_state.json"
    settings_path = Path(BASE) / "push.json"

    # Zustand prüfen (nur einmal pro Tag)
    state = _load_json(state_path) or {}
    if state.get("zuletzt") == heute:
        return

    # Einstellungen lesen
    settings = _load_json(settings_path)
    if settings is None:
        return
    if not settings.get("aktiv"):
        return

    # Nachricht aufbauen und senden
    text = kurzfassung(heute)
    notify("Coach Kira · Morgen", text)

    adresse = settings.get("imessage", "")
    if adresse:
        imessage_senden(adresse, "Coach Kira: " + text)

    # Zustand speichern
    _save_json(state_path, {"zuletzt": heute})

    suffix = ", iMessage" if adresse else ""
    log(f"push: gesendet ({len(text)} Zeichen{suffix})")


# ── 4. CLI-Hauptteil ─────────────────────────────────────────

if __name__ == "__main__":
    import sys

    heute = datetime.datetime.now(TZ).strftime("%Y-%m-%d")
    if sys.argv[1:] == ["test"]:
        print(kurzfassung(heute))
