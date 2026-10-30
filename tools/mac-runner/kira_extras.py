"""Erweiterungen für Coach Kira: Plausi-Check, Selbsttest lokales Modell, CLI.

Liest kira_runner.py aus demselben Verzeichnis und nutzt dort log, http_get_json
usw.  NUR Standardbibliothek.
"""
from __future__ import annotations

import argparse
import datetime
import json
import re
import os
import sys
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Import aus kira_runner
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).resolve().parent))

from kira_runner import (
    log,
    http_get_json,
    http_post_json,
    ollama_chat,
    ollama_ready,
    notify,
    num,
    SCRIPT_URL,
    MODEL,
    TZ,
    BASE,
)

# ---------------------------------------------------------------------------
# State-Datei für Extras (atomar schreiben)
# ---------------------------------------------------------------------------

STATE_X = BASE / "extras_state.json"


def _load_json(path: Path) -> dict:
    """JSON-Datei lesen, bei Fehler leeres dict."""
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_json_atomic(path: Path, d: dict) -> None:
    """JSON atomar schreiben über tmp + rename."""
    tmp = path.with_suffix(".tmp")
    d2 = json.dumps(d, ensure_ascii=False, indent=2)
    tmp.write_text(d2, encoding="utf-8")
    os.replace(str(tmp), str(path))


def load_x() -> dict:
    """Zustandsdatei extras_state.json laden."""
    return _load_json(STATE_X)


def save_x(d: dict) -> None:
    """Zustandsdatei extras_state.json atomar speichern."""
    STATE_X.parent.mkdir(parents=True, exist_ok=True)
    _save_json_atomic(STATE_X, d)


# ---------------------------------------------------------------------------
# Zeit-Helfer (lokale Zeit in Europe/Berlin)
# ---------------------------------------------------------------------------

tz_local = TZ


def today() -> str:
    """'YYYY-MM-DD' im Zeitzone TZ."""
    return datetime.datetime.now(tz_local).strftime("%Y-%m-%d")


def now_h() -> int:
    """Stunde des Tages in TZ."""
    return datetime.datetime.now(tz_local).hour


# ---------------------------------------------------------------------------
# Daten-Helfer
# ---------------------------------------------------------------------------

METRICS_QUERY = (
    "sleep_hours,sleep_score_0_100,rhr_bpm,hrv_status,hrv_threshholds,"
    "Garmin_Training_Readiness,garminATL,garminCTL,fbATL_obs,fbCTL_obs,"
    "kcal_in,kcal_out,carb_g,protein_g,fat_g,load_fb_day,"
    "Aerobic_TE,Anaerobic_TE"
)


def fetch_rows(days: int, metrics: list[str]) -> list[dict]:
    """http_get_json(SCRIPT_URL + '?mode=chartData&range=%dd&metrics=' % days + ','.join(metrics))['data']."""
    url = SCRIPT_URL + "?mode=chartData&range=%dd&metrics=" % days + ",".join(metrics)
    return http_get_json(url).get("data") or []


def save_extra(typ: str, datum: str, text: str, data: object, quelle: str) -> bool:
    """Ergebnis in AI_KIRA_EXTRAS speichern; ein zweiter Versuch bei Fehlern, nie eine Ausnahme."""
    body = {"kind": "saveExtra", "typ": typ, "datum": datum, "text": text, "data": data, "quelle": quelle}
    for versuch in (1, 2):
        try:
            result = http_post_json(SCRIPT_URL + "?mode=saveExtra", body)
            if result.get("ok"):
                return True
            log(f"save_extra({typ},{datum}): FEHLER – {result.get('error', '')}")
        except Exception as e:
            log(f"save_extra({typ},{datum}): Versuch {versuch} – {str(e)[:120]}")
        time.sleep(3)
    return False


def parse_loose(text: str):
    """JSON aus Modell-Antwort lesen (ohne erzwungenes Format): Codeblöcke entfernen, erstes {...} nehmen."""
    import re
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t, flags=re.S).strip()
    try:
        return json.loads(t)
    except Exception:
        pass
    a, b = t.find("{"), t.rfind("}")
    if a >= 0 and b > a:
        return json.loads(t[a:b + 1])
    raise ValueError("keine JSON-Antwort: " + (t[:80] or "leer"))


def f(v: object) -> float | None:
    """Alias für num()."""
    return num(v)


# ---------------------------------------------------------------------------
# A) PLAUSIBILITÄTS-PRÜFUNG
# ---------------------------------------------------------------------------


def _median(values: list[float]) -> float:
    """Median einer sortierten Liste von Zahlen."""
    s = sorted(values)
    n = len(s)
    if n == 0:
        return 0.0
    mid = n // 2
    if n % 2 == 1:
        return s[mid]
    return (s[mid - 1] + s[mid]) / 2.0


def _fmt(v: float | None) -> str:
    """Zahl mit Komma formatieren (1 Dezimalstelle), None als '–'."""
    if v is None:
        return "–"
    if not isinstance(v, (int, float)):
        return str(v) or "–"
    return f"{v:.1f}".replace(".", ",")


def plausi_check(row: dict, hist: list[dict]) -> list[dict]:
    """Prüfe row nach den Regeln aus plausi-spec.md.

    hist = Liste der Zeilen VOR dem Tag, sortiert alt -> neu.
    prev = letzte Zeile in hist (Vortags-Wert-Quelle).
    Jeder Befund: {feld, wert, erwartet, grund, schwere ('warnung'|'hinweis')}.
    """
    def _median_vals(feld: str, min_count: int = 5) -> float | None:
        vals = []
        for h in hist:
            v = num(h.get(feld))
            if v is not None:
                vals.append(v)
        if len(vals) < min_count:
            return None
        return _median(sorted(vals))

    befunde: list[dict] = []
    prev_hist = hist[-1] if hist else None

    # --- sleep_hours ----------------------------------------------------------
    sh = num(row.get("sleep_hours"))
    if sh is not None:
        med = _median_vals("sleep_hours")
        if sh < 3.0 or sh > 12.0:
            befunde.append({
                "feld": "sleep_hours", "wert": sh,
                "erwartet": f"ca. {_fmt(med)} h",
                "grund": "außerhalb 3–12 Stunden", "schwere": "warnung"
            })
        elif med is not None and abs(sh - med) > 3.0:
            befunde.append({
                "feld": "sleep_hours", "wert": sh,
                "erwartet": f"ca. {_fmt(med)} h",
                "grund": f"Ausreißer – {sh:.1f} statt ca. {_fmt(med)}",
                "schwere": "hinweis"
            })

        # sleep_score_0_100
        ss = num(row.get("sleep_score_0_100"))
        if ss is not None:
            if ss < 0 or ss > 100:
                befunde.append({
                    "feld": "sleep_score_0_100", "wert": ss,
                    "erwartet": "0–100",
                    "grund": "außerhalb gültigen Bereichs", "schwere": "warnung"
                })
            else:
                if ss >= 80 and sh is not None and sh < 5:
                    befunde.append({
                        "feld": "sleep_score_0_100", "wert": ss,
                        "erwartet": f"höher bei {sh:.1f}h Schlaf",
                        "grund": "Score passt nicht zur Schlafdauer",
                        "schwere": "hinweis"
                    })
                if ss < 40 and sh is not None and sh > 8:
                    befunde.append({
                        "feld": "sleep_score_0_100", "wert": ss,
                        "erwartet": f"niedriger bei {sh:.1f}h Schlaf",
                        "grund": "Score passt nicht zur Schlafdauer",
                        "schwere": "hinweis"
                    })

    # --- rhr_bpm --------------------------------------------------------------
    rhr = num(row.get("rhr_bpm"))
    if rhr is not None:
        med_rhr = _median_vals("rhr_bpm")
        if rhr < 30 or rhr > 100:
            befunde.append({
                "feld": "rhr_bpm", "wert": rhr,
                "erwartet": f"ca. {_fmt(med_rhr)} bpm",
                "grund": "außerhalb 30–100 bpm", "schwere": "warnung"
            })
        elif med_rhr is not None and abs(rhr - med_rhr) > 8:
            befunde.append({
                "feld": "rhr_bpm", "wert": rhr,
                "erwartet": f"ca. {_fmt(med_rhr)} bpm",
                "grund": f"Ausreißer – {rhr:.0f} statt ca. {_fmt(med_rhr)}",
                "schwere": "hinweis"
            })

    # --- hrv_status -----------------------------------------------------------
    hrv_s = num(row.get("hrv_status"))
    if hrv_s is not None:
        med_hrv = _median_vals("hrv_status")
        if hrv_s < 10 or hrv_s > 200:
            befunde.append({
                "feld": "hrv_status", "wert": hrv_s,
                "erwartet": f"ca. {_fmt(med_hrv)} ms",
                "grund": "außerhalb 10–200 ms", "schwere": "warnung"
            })
        elif med_hrv is not None:
            if hrv_s < 0.6 * med_hrv or hrv_s > 1.5 * med_hrv:
                befunde.append({
                    "feld": "hrv_status", "wert": hrv_s,
                    "erwartet": f"ca. {_fmt(med_hrv)} ms",
                    "grund": f"Ausreißer – {hrv_s:.0f} statt ca. {_fmt(med_hrv)}",
                    "schwere": "hinweis"
                })

    # --- hrv_threshholds ------------------------------------------------------
    hrv_t = row.get("hrv_threshholds")
    if hrv_t is not None and str(hrv_t).strip() != "":
        s = str(hrv_t).strip()
        parts = [p.strip() for p in s.split(";")] if ";" in s else []
        valid = False
        if len(parts) == 2:
            try:
                a, b = float(parts[0].replace(",", ".")), float(parts[1].replace(",", "."))
                if a < b:
                    valid = True
            except (ValueError, TypeError):
                pass
        if not valid:
            befunde.append({
                "feld": "hrv_threshholds", "wert": s,
                "erwartet": "Format a;b mit Zahlen a<b",
                "grund": "Normalband unten;oben", "schwere": "warnung"
            })

    # --- Garmin_Training_Readiness --------------------------------------------
    tr = num(row.get("Garmin_Training_Readiness"))
    if tr is not None:
        med_tr = _median_vals("Garmin_Training_Readiness")
        if tr < 0 or tr > 100:
            befunde.append({
                "feld": "Garmin_Training_Readiness", "wert": tr,
                "erwartet": f"0–100",
                "grund": "außerhalb 0–100", "schwere": "warnung"
            })
        elif med_tr is not None and abs(tr - med_tr) > 35:
            befunde.append({
                "feld": "Garmin_Training_Readiness", "wert": tr,
                "erwartet": f"ca. {_fmt(med_tr)}",
                "grund": f"Ausreißer – {tr:.0f} statt ca. {_fmt(med_tr)}",
                "schwere": "hinweis"
            })

    # --- ATL / CTL (garminATL, garminCTL, fbATL_obs, fbCTL_obs) -------------
    atl_ctl_fields = ("garminATL", "garminCTL", "fbATL_obs", "fbCTL_obs")
    for af in atl_ctl_fields:
        val = num(row.get(af))
        if val is not None:
            if val <= 0:
                befunde.append({
                    "feld": af, "wert": val, "erwartet": "> 0",
                    "grund": f"{af} ≤ 0", "schwere": "warnung"
                })
            elif prev_hist is not None:
                pval = num(prev_hist.get(af))
                if pval is not None and pval > 0:
                    rel = abs(val - pval) / pval
                    if rel > 0.25:
                        befunde.append({
                            "feld": af, "wert": val,
                            "erwartet": f"ca. {_fmt(pval)} (Vortag)",
                            "grund": f"Sprung > 25 % zum Vortag ({rel*100:.0f}%)",
                            "schwere": "hinweis"
                        })

    # --- kcal_in / kcal_out ---------------------------------------------------
    ki = num(row.get("kcal_in"))
    if ki is not None and (ki < 800 or ki > 6500):
        befunde.append({
            "feld": "kcal_in", "wert": ki, "erwartet": "800–6500",
            "grund": f"{ki:.0f} kcal außerhalb Bereich", "schwere": "warnung"
        })
    ko = num(row.get("kcal_out"))
    if ko is not None and (ko < 1200 or ko > 8000):
        befunde.append({
            "feld": "kcal_out", "wert": ko, "erwartet": "1200–8000",
            "grund": f"{ko:.0f} kcal außerhalb Bereich", "schwere": "warnung"
        })

    # --- Makros: carb_g, protein_g, fat_g ------------------------------------
    cg = num(row.get("carb_g"))
    pg = num(row.get("protein_g"))
    fg = num(row.get("fat_g"))
    if cg is not None and pg is not None and fg is not None and ki is not None:
        berechnet = 4.0 * cg + 4.0 * pg + 9.0 * fg
        if ki > 0 and abs(berechnet - ki) / ki > 0.30:
            befunde.append({
                "feld": "Makros", "wert": berechnet,
                "erwartet": f"{ki:.0f} kcal",
                "grund": f"Makros ergeben {berechnet:.0f} kcal, eingetragen {ki:.0f}",
                "schwere": "hinweis"
            })

    if pg is not None and pg > 400:
        befunde.append({
            "feld": "protein_g", "wert": pg, "erwartet": "< 400",
            "grund": f"{pg:.0f} g zu viel", "schwere": "warnung"
        })
    if cg is not None and cg > 1000:
        befunde.append({
            "feld": "carb_g", "wert": cg, "erwartet": "< 1000",
            "grund": f"{cg:.0f} g zu viel", "schwere": "warnung"
        })
    if fg is not None and fg > 400:
        befunde.append({
            "feld": "fat_g", "wert": fg, "erwartet": "< 400",
            "grund": f"{fg:.0f} g zu viel", "schwere": "warnung"
        })

    # --- load_fb_day ----------------------------------------------------------
    lf = num(row.get("load_fb_day"))
    if lf is not None:
        med_lf = _median_vals("load_fb_day")
        if lf < 0 or lf > 800:
            befunde.append({
                "feld": "load_fb_day", "wert": lf,
                "erwartet": f"0–800 (ca. {_fmt(med_lf) if med_lf else 'mind. 5 Werte'})",
                "grund": f"{lf:.0f} außerhalb Bereich", "schwere": "warnung"
            })
        elif med_lf is not None and med_lf > 20 and lf > 2.5 * med_lf:
            befunde.append({
                "feld": "load_fb_day", "wert": lf,
                "erwartet": f"ca. {_fmt(med_lf)}",
                "grund": f"{lf:.0f} > 2.5× Median", "schwere": "hinweis"
            })

    # --- Aerobic_TE, Anaerobic_TE ---------------------------------------------
    for te_key in ("Aerobic_TE", "Anaerobic_TE"):
        te_val = num(row.get(te_key))
        if te_val is not None and (te_val < 0 or te_val > 5):
            befunde.append({
                "feld": te_key, "wert": te_val, "erwartet": "0–5",
                "grund": f"{te_val:.1f} außerhalb Bereich", "schwere": "warnung"
            })

    return befunde


def cmd_plausi() -> None:
    """Plausibilitäts-Check für heute und gestern."""
    metrics = [
        "sleep_hours", "sleep_score_0_100", "rhr_bpm", "hrv_status",
        "hrv_threshholds", "Garmin_Training_Readiness", "garminATL",
        "garminCTL", "fbATL_obs", "fbCTL_obs", "kcal_in", "kcal_out",
        "carb_g", "protein_g", "fat_g", "load_fb_day",
        "Aerobic_TE", "Anaerobic_TE",
    ]

    rows = fetch_rows(35, metrics)
    if not rows:
        log("plausi: keine Daten vorhanden")
        return

    # Zeilen gruppieren nach datum
    row_map: dict[str, dict] = {}
    for r in rows:
        d = r.get("date")
        if d:
            row_map[d] = r

    # Alle Datums-Werte sortiert (alt -> neu)
    alle_dates = sorted(row_map.keys())

    today_str = today()
    yesterday_dt = datetime.datetime.now(tz_local) - datetime.timedelta(days=1)
    yesterday_str = yesterday_dt.strftime("%Y-%m-%d")

    target_dates = [today_str, yesterday_str]
    state = load_x()
    plausi_fp: dict[str, str] = state.get("plausi_fp", {})

    alle_geprueft: set[str] = set()

    for datum in target_dates:
        if datum not in row_map:
            continue

        row = row_map[datum]
        # Zeilen VOR diesem Tag, sortiert alt->neu (hist)
        hist = [row_map[d] for d in alle_dates if d < datum and d in row_map]

        # Fingerabdruck: json.dumps der METRICS-Werte (sort_keys)
        fp_values: dict[str, float | None] = {}
        for m in metrics:
            v = row.get(m)
            fp_values[m] = num(v) if v is not None else None

        fingerprint = json.dumps(fp_values, sort_keys=True)
        prev_fp = plausi_fp.get(datum)

        befunde: list[dict] = []
        if fingerprint != prev_fp:
            befunde = plausi_check(row, hist)

        # Felder sammeln
        for m in metrics:
            if row.get(m) is not None and num(row.get(m)) is not None:
                alle_geprueft.add(m)

        # Text aufbauen
        if not befunde:
            text = "Keine Auffälligkeiten."
        else:
            lines: list[str] = []
            for b in befunde:
                lines.append(
                    f"- {b['feld']}: {_fmt(b['wert'])} ({b['grund']}; erwartet {b['erwartet']})"
                )
            text = "\n".join(lines)

        ok = save_extra("plausi", datum, text,
                        {"befunde": befunde, "geprueft": sorted(alle_geprueft)},
                        "mac-regeln")
        if not ok:
            continue

        # Warnungen prüfen -> Notify
        warner_felder = [b["feld"] for b in befunde if b.get("schwere") == "warnung"]
        if warner_felder:
            notify("Coach Kira", "Eingabe prüfen: " + ", ".join(warner_felder))

        # Fingerabdruck speichern
        plausi_fp[datum] = fingerprint

    state["plausi_fp"] = plausi_fp
    save_x(state)


# ---------------------------------------------------------------------------
# B) SELBSTTEST lokales Modell
# ---------------------------------------------------------------------------


def cmd_selbsttest(force: bool = False) -> None:
    """Selbsttest des lokalen Ollama-Modells mit 6 Tests."""
    st = load_x()

    # Nur einmal pro Tag (außer force)
    if not force and st.get("selbsttest_done") == today():
        pass
        return

    if not ollama_ready():
        log("selbsttest: Ollama nicht erreichbar – abgebrochen")
        return

    # Plan laden für echte Testdaten
    plan = http_get_json(SCRIPT_URL + "?mode=planJsonV2").get("plan") or []
    tage: list[dict] = [p for p in plan if "date" in p and "day" in p and "original_load_ess" in p]

    if len(tage) < 5:
        # Synthetischer Plan mit 8 Tagen ab heute und Loads [150,190,170,155,0,145,125,145]
        loads = [150, 190, 170, 155, 0, 145, 125, 145]
        tage = []
        for i in range(8):
            d = (datetime.datetime.now(tz_local) + datetime.timedelta(days=i)).strftime("%Y-%m-%d")
            wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][i % 7]
            tage.append({"date": d, "day": wd, "original_load_ess": loads[i]})

    # --- Tests ---------------------------------------------------------------
    ergebnisse: list[dict] = []
    gesamt_ssekunden: float = 0.0

    # T1 'Zahlen übernehmen'
    def test_1() -> dict:
        text = ""
        for t in tage:
            original = int(t["original_load_ess"])
            text += f"{t['date']} ({t['day']}): Geplanter Load {original} (ORIGINAL-PLAN: {original})\n"
        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Übernimm diese Tage exakt als JSON. Aufgabe: {{\"tage\":[{{\"datum\":..,\"original_load_ess\":..}}]}}\n\n{text}"},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = parse_loose(text_out)
            antwort_tage = data.get("tage", [])
            ok = True
            detail = ""
            if len(antwort_tage) != len(tage):
                ok = False
                detail = f"Anzahl: {len(antwort_tage)} statt {len(tage)}"
            else:
                for i, at in enumerate(antwort_tage):
                    ref = tage[i]
                    if at.get("datum") != ref["date"]:
                        ok = False
                        detail = f"Datum {at.get('datum')} statt {ref['date']}"
                        break
                    if int(at.get("original_load_ess", -999)) != ref["original_load_ess"]:
                        ok = False
                        detail = f"Load {at.get('original_load_ess')} statt {ref['original_load_ess']}"
                        break
            return {"id": "T1", "name": "Zahlen übernehmen", "ok": ok,
                    "detail": detail or "Alle Tage und Loads exakt gleich",
                    "sekunden": round(sekunden, 1), "tps": tps}
        except Exception as e:
            return {"id": "T1", "name": "Zahlen übernehmen", "ok": False,
                    "detail": str(e), "sekunden": 0, "tps": None}

    # T2 'Gesperrte Tage'
    def test_2() -> dict:
        text = ""
        for i, t in enumerate(tage):
            marker = " [FIX]" if (i % 2 == 1) else ""
            text += f"{t['date']} ({t['day']}): Geplanter Load {int(t['original_load_ess'])}{marker}\n"
        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Empfehlung je Tag = 80 % des Plans gerundet. FIX-Tage unverändert.\nAufgabe: JSON {{\"tage\":[{{\"datum\", ... , \"empfohlen\"}}]}}\n\n{text}"},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = parse_loose(text_out)
            antwort_tage = data.get("tage", [])
            ok = True
            detail = ""
            if len(antwort_tage) != len(tage):
                ok = False
                detail = f"Anzahl: {len(antwort_tage)} statt {len(tage)}"
            else:
                for i, at in enumerate(antwort_tage):
                    ref = tage[i]
                    if at.get("datum") != ref["date"]:
                        ok = False
                        detail += f"Datum {at.get('datum')} statt {ref['date']}\n"
                        continue
                    if i % 2 == 1:
                        if int(at.get("empfohlen", -999)) != ref["original_load_ess"]:
                            ok = False
                            detail += f"FIX-Tag Load {at.get('empfohlen')} statt {ref['original_load_ess']}\n"
                    else:
                        expected = round(0.8 * ref["original_load_ess"])
                        if abs(at.get("empfohlen", 999) - expected) > 2:
                            ok = False
                            detail += f"Nicht-FIX {at.get('empfohlen')} statt {expected} (±2 toleriert)\n"
            return {"id": "T2", "name": "Gesperrte Tage", "ok": ok,
                    "detail": detail.strip() or "80 %-Empfehlung und FIX-Tage korrekt",
                    "sekunden": round(sekunden, 1), "tps": tps}
        except Exception as e:
            return {"id": "T2", "name": "Gesperrte Tage", "ok": False,
                    "detail": str(e), "sekunden": 0, "tps": None}

    # T3 'Rechnen'
    def test_3() -> dict:
        loads_list = [int(t["original_load_ess"]) for t in tage]
        waechte_sum = sum(loads_list)
        wahr_mittel = waechte_sum / len(loads_list) if loads_list else 0
        text = f"Summe und Mittelwert berechnen.\nLoads:\n"
        for lv in loads_list:
            text += f"  {lv}\n"
        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Aufgabe: JSON {{\"summe\":..,\"mittelwert\":..}}\n\n{text}"},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = parse_loose(text_out)
            s_val = data.get("summe")
            mw_val = data.get("mittelwert")
            ok = True
            detail = ""
            if s_val is None or mw_val is None:
                ok = False
                detail = "Fehlende Schlüssel summe/mittelwert"
            else:
                if int(s_val) != waechte_sum:
                    ok = False
                    detail += f"Summe {s_val} statt {waechte_sum}\n"
                if abs(float(mw_val) - wahr_mittel) > 0.6:
                    ok = False
                    detail += f"Mittelwert {mw_val} statt {wahr_mittel:.2f}"
            return {"id": "T3", "name": "Rechnen", "ok": ok,
                    "detail": detail.strip() or f"Summe {waechte_sum}, Mittelwert {wahr_mittel:.2f}",
                    "sekunden": round(sekunden, 1), "tps": tps}
        except Exception as e:
            return {"id": "T3", "name": "Rechnen", "ok": False,
                    "detail": str(e), "sekunden": 0, "tps": None}

    # T4 'Werte auslesen'
    def _stunden(v) -> float:
        """Zahl oder Text wie '7h 24min', '7 h 24 min', '7:24', '7,4' in Dezimalstunden."""
        if isinstance(v, (int, float)):
            return float(v)
        t = str(v).strip().lower().replace(",", ".")
        m = re.match(r"^(\d+):(\d{1,2})$", t)
        if m:
            return int(m.group(1)) + int(m.group(2)) / 60
        m = re.match(r"^(\d+(?:\.\d+)?)\s*(?:h|std\.?|stunden?)?\s*(?:(\d+)\s*(?:m|min\.?|minuten?))?$", t)
        if m:
            return float(m.group(1)) + (int(m.group(2)) / 60 if m.group(2) else 0)
        return float(t)

    def test_4() -> dict:
        prompt_text = "Heute Nacht 7 h 24 min geschlafen, Schlafscore 81. Ruhepuls 46, HRV 50 ms, Trainingsbereitschaft 77."
        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Aus dem Text extrahieren:\n{prompt_text}\nAufgabe: JSON {{\"sleep_hours\",\"sleep_score\",\"rhr\",\"hrv\",\"readiness\"}}. Alle Werte als Zahlen ohne Einheit; sleep_hours als Dezimalzahl in Stunden (z. B. 6 h 30 min = 6.5)."},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = parse_loose(text_out)
            ok = True
            detail = ""

            sh_raw = data.get("sleep_hours", -999)
            sh_val = _stunden(sh_raw)
            if not isinstance(sh_raw, (int, float)):
                detail += f"sleep_hours als Text geliefert ({sh_raw!r}), umgerechnet\n"
            if abs(sh_val - 7.4) > 0.01:
                ok = False
                detail += f"sleep_hours {sh_val} statt 7.4\n"

            if int(data.get("sleep_score", -999)) != 81:
                ok = False
                detail += f"sleep_score {data.get('sleep_score')} statt 81\n"

            if int(data.get("rhr", -999)) != 46:
                ok = False
                detail += f"rhr {data.get('rhr')} statt 46\n"

            if int(data.get("hrv", -999)) != 50:
                ok = False
                detail += f"hrv {data.get('hrv')} statt 50\n"

            if int(data.get("readiness", -999)) != 77:
                ok = False
                detail += f"readiness {data.get('readiness')} statt 77\n"

            return {"id": "T4", "name": "Werte auslesen", "ok": ok,
                    "detail": detail.strip() or "Alle Werte korrekt extrahiert",
                    "sekunden": round(sekunden, 1), "tps": tps}
        except Exception as e:
            return {"id": "T4", "name": "Werte auslesen", "ok": False,
                    "detail": str(e), "sekunden": 0, "tps": None}

    # T5 'Wochentage'
    def test_5() -> dict:
        wd_map = {0: "Mo", 1: "Di", 2: "Mi", 3: "Do", 4: "Fr", 5: "Sa", 6: "So"}
        text = ""
        test_dates = []
        for t in tage[:5]:
            dt = datetime.datetime.strptime(t["date"], "%Y-%m-%d")
            wd = wd_map[dt.weekday()]
            text += f"{t['date']}: {wd}\n"
            test_dates.append((t["date"], wd))

        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Wochentage zu diesen Daten ermitteln. Kürzel: Mo Di Mi Do Fr Sa So\nAufgabe: JSON {{\"tage\":[{{\"datum\",\"wochentag\"}}]}}\n\n{text}"},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = parse_loose(text_out)
            antwort_tage = data.get("tage", [])
            ok = True
            detail = ""
            if len(antwort_tage) != len(test_dates):
                ok = False
                detail = f"Anzahl: {len(antwort_tage)} statt {len(test_dates)}"
            else:
                for i, at in enumerate(antwort_tage):
                    ref_d, ref_wd = test_dates[i]
                    if at.get("datum") != ref_d:
                        ok = False
                        detail += f"Datum {at.get('datum')} statt {ref_d}\n"
                    elif at.get("wochentag") != ref_wd:
                        ok = False
                        detail += f"Wochentag {at.get('wochentag')} statt {ref_wd}\n"
            return {"id": "T5", "name": "Wochentage", "ok": ok,
                    "detail": detail.strip() or "Alle Wochentage korrekt",
                    "sekunden": round(sekunden, 1), "tps": tps}
        except Exception as e:
            return {"id": "T5", "name": "Wochentage", "ok": False,
                    "detail": str(e), "sekunden": 0, "tps": None}

    # T6 'Schema'
    def test_6() -> dict:
        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Bewerte kurz 'Schlaf 7 h, HRV im Normalband' als JSON.\nAufgabe: JSON mit Schlüsseln \"ampel\" (GRÜN|GELB|ROT), \"text\" (1 Satz), \"punkte\" (0-100 Zahl)."},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = parse_loose(text_out)
            ok = True
            detail = ""

            ampel = data.get("ampel")
            text_val = data.get("text", "")
            punkte = data.get("punkte")

            if ampel not in ("GRÜN", "GELB", "ROT"):
                ok = False
                detail += f"ampel '{ampel}' ungültig (GRÜN/GELB/ROT)\n"
            if not text_val or not text_val.strip():
                ok = False
                detail += "text ist leer\n"

            punkte_val: float | None = None
            if punkte is not None:
                try:
                    punkte_val = float(punkte)
                    if punkte_val < 0 or punkte_val > 100:
                        ok = False
                        detail += f"punkte {punkte_val} außerhalb 0-100\n"
                except (ValueError, TypeError):
                    ok = False
                    detail += f"punkte '{punkte}' ist keine Zahl\n"
            else:
                ok = False
                detail += "Fehlender Schlüssel 'punkte'\n"

            if not all(k in data for k in ("ampel", "text", "punkte")):
                ok = False
                detail += "Nicht alle drei Schlüssel vorhanden\n"

            return {"id": "T6", "name": "Schema", "ok": ok,
                    "detail": detail.strip() or "Ampel, Text und Punkte valide",
                    "sekunden": round(sekunden, 1), "tps": tps}
        except Exception as e:
            return {"id": "T6", "name": "Schema", "ok": False,
                    "detail": str(e), "sekunden": 0, "tps": None}

    # Alle Tests ausführen
    test_funktionen = [test_1, test_2, test_3, test_4, test_5, test_6]
    tps_werte: list[float | None] = []

    for tf in test_funktionen:
        r = tf()
        ergebnisse.append(r)
        gesamt_ssekunden += r["sekunden"]
        if r.get("tps") is not None:
            tps_werte.append(r["tps"])

    bestanden = sum(1 for e in ergebnisse if e["ok"])
    avg_tps: float = sum(tps_werte) / len(tps_werte) if tps_werte else 0

    # Text zusammenbauen
    lines_out = [f"{bestanden}/6 Tests bestanden ({MODEL}, Ø {avg_tps:.0f} tok/s, {gesamt_ssekunden:.0f} s)."]
    for e in ergebnisse:
        if not e["ok"]:
            lines_out.append(f"- {e['name']}: {e['detail']}")
    text = "\n".join(lines_out)

    # Speichern
    save_extra(
        "selbsttest", today(), text,
        {
            "tests": ergebnisse,
            "modell": MODEL,
            "tok_s": round(avg_tps, 1),
            "sekunden": round(gesamt_ssekunden, 1),
        },
        "mac",
    )

    st["selbsttest_done"] = today()
    save_x(st)

    # Notify bei unvollständig
    if bestanden < 6:
        notify("Coach Kira", f"Selbsttest lokales Modell: {bestanden}/6")

    log(f"selbsttest: {text}")


# ---------------------------------------------------------------------------
# C) CLI – main()
# ---------------------------------------------------------------------------

def main() -> None:
    """CLI mit Unterbefehlen: plausi, selbsttest [--force], tick."""
    ap = argparse.ArgumentParser(description="Coach Kira Extras")
    sub = ap.add_subparsers(dest="command")

    sub.add_parser("plausi", help="Plausibilitäts-Check")
    st_p = sub.add_parser("selbsttest", help="Selbsttest lokales Modell")
    st_p.add_argument("--force", action="store_true", help="Tag-Check ignorieren")
    sub.add_parser("tick", help="Plausi immer + Selbsttest 2-5 Uhr")

    args = ap.parse_args()
    if args.command is None:
        ap.print_help()
        return

    if args.command == "plausi":
        try:
            cmd_plausi()
        except Exception as e:
            log(f"extras: FEHLER plausi – {e}")

    elif args.command == "selbsttest":
        try:
            cmd_selbsttest(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER selbsttest – {e}")

    elif args.command == "tick":
        try:
            cmd_plausi()
        except Exception as e:
            log(f"extras: FEHLER tick/plausi – {e}")

        try:
            nh = now_h()
            # nachts 2-5 Uhr; schläft der Mac dann, ab 13 Uhr nachholen
            if 2 <= nh < 5 or nh >= 13:
                cmd_selbsttest(force=False)
        except Exception as e:
            log(f"extras: FEHLER tick/selbsttest – {e}")

    # --- Treffsicherheit / Ernährung folgen ---

# ---------------------------------------------------------------------------
# D) EMPFEHLUNGS-SNAPSHOT
# ---------------------------------------------------------------------------


def cmd_snapshot() -> None:
    """Empfehlung-Snapshot für heute in extras_state speichern."""
    try:
        plan = http_get_json(SCRIPT_URL + "?mode=planJsonV2").get("plan") or []
    except Exception as e:
        log(f"snapshot: FEHLER beim Plan-Laden – {e}")
        return

    heute = today()
    eintrag = None
    for p in plan:
        if p.get("date") == heute:
            eintrag = p
            break

    st = load_x()
    rec = st.get("rec", {})

    if eintrag is not None:
        empfohlen = eintrag.get("recommended_load_ess")
        plan_load = eintrag.get("original_load_ess")
        zone = eintrag.get("recommended_zone", "")
        training = eintrag.get("training_1", "")
        if empfohlen is None:
            empfohlen = 0
        if plan_load is None:
            plan_load = int(empfohlen) if empfohlen else 0

        hhmm = datetime.datetime.now(tz_local).strftime("%H:%M")
        rec[heute] = {
            "empfohlen": empfohlen,
            "plan": plan_load,
            "zone": zone,
            "training": training,
            "quelle": "supervisor",
            "zeit": hhmm,
        }

    # Nur speichern solange now_h() < 10
    if not (5 <= now_h() < 10):
        log(f"snapshot: Zeitfenster {now_h()}h nicht erfüllt (5-10) – übersprungen")
        return

    # Empfehlungen der ganzen Historie behalten (Obergrenze nur als Schutz)
    if rec:
        keys = sorted(rec.keys(), reverse=True)[:1000]
        st["rec"] = {k: rec[k] for k in keys}

    save_x(st)
    log(f"snapshot: {heute} gespeichert – empfohlen={rec.get(heute, {}).get('empfohlen', '?')}")


# ---------------------------------------------------------------------------
# E) TREFFSICHERHEIT TAG
# ---------------------------------------------------------------------------

METRICS_T = [
    "load_fb_day",
    "Garmin_Training_Readiness",
    "hrv_status",
    "rhr_bpm",
    "sleep_score_0_100",
    "befinden_1_5",
]


def treffer_fuer(datum: str, rec: dict, rows_by_date: dict) -> dict | None:
    """Empfehlung vs. Ist für Tag D und Reaktion am Morgen D+1 (rows_by_date: date -> Zeile)."""
    row_d = rows_by_date.get(datum)
    n = (datetime.datetime.strptime(datum, "%Y-%m-%d") + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    row_n = rows_by_date.get(n)
    if not row_d or not row_n:
        return None
    empfohlen = num(rec.get("empfohlen"))
    if empfohlen is None:
        return None
    ist = num(row_d.get("load_fb_day"))
    if ist is None:
        if empfohlen > 0:
            return None
        ist = 0.0
    if num(row_n.get("Garmin_Training_Readiness")) is None:
        return None

    def basis(feld: str) -> float | None:
        werte = []
        d0 = datetime.datetime.strptime(datum, "%Y-%m-%d")
        for i in range(6, -1, -1):
            tag = (d0 - datetime.timedelta(days=i)).strftime("%Y-%m-%d")
            v = num((rows_by_date.get(tag) or {}).get(feld))
            if v is not None:
                werte.append(v)
        return sum(werte) / len(werte) if werte else None

    felder = {"readiness": "Garmin_Training_Readiness", "hrv": "hrv_status", "rhr": "rhr_bpm",
              "schlaf": "sleep_score_0_100", "befinden": "befinden_1_5"}
    reaktion: dict = {}
    for k, feld in felder.items():
        w, bv = num(row_n.get(feld)), basis(feld)
        if w is not None and bv is not None:
            reaktion[k] = {"wert": w, "basis": round(bv, 1), "delta": round(w - bv, 1)}

    abw = round((ist - empfohlen) / empfohlen * 100, 1) if empfohlen > 0 else None
    gefolgt = (abs(abw) <= 20) if abw is not None else (ist <= 15)
    if empfohlen <= 10 and ist <= 15:
        gefolgt = True
    dl = lambda k: reaktion.get(k, {}).get("delta", 0)  # noqa: E731
    score = round(dl("readiness") / 10 + dl("hrv") / 5 - dl("rhr") / 3 + dl("schlaf") / 10, 2)

    def de(v: float) -> str:
        return (f"{int(round(v)):+d}").replace("-", "−")

    teile = []
    for k, name in (("readiness", "Bereitschaft"), ("hrv", "HRV"), ("rhr", "Ruhepuls"), ("schlaf", "Schlaf")):
        if k in reaktion:
            teile.append(f"{name} {reaktion[k]['wert']:.0f} ({de(reaktion[k]['delta'])})")
    abw_txt = f"{de(abw)} %" if abw is not None else "Ruhetag"
    text = (f"Empfohlen {empfohlen:.0f}, gemacht {ist:.0f} ({abw_txt}, {'gefolgt' if gefolgt else 'nicht gefolgt'}). "
            f"Morgen danach: {', '.join(teile)}.")
    return {"datum": datum, "empfohlen": empfohlen, "plan": rec.get("plan"), "ist": ist, "abweichung_pct": abw,
            "gefolgt": gefolgt, "reaktion": reaktion, "score": score, "quelle": rec.get("quelle", "supervisor"),
            "text": text}


def cmd_treffer(backfill_file: str | None = None) -> None:
    """Verarbeitet Treffer für alle noch nicht erledigten Tage."""
    rows = fetch_rows(370, METRICS_T)  # ganze Historie (Schnittstelle liefert max. ca. 365 Tage)
    rows_by_date: dict[str, dict] = {r["date"]: r for r in rows if r.get("date")}

    st = load_x()
    rec: dict[str, dict] = st.get("rec", {})

    # backfill ergänzen
    if backfill_file:
        try:
            bf = json.load(Path(backfill_file).open("r", encoding="utf-8"))
            for d, v in bf.items():
                if isinstance(v, dict) and "empfohlen" in v and d not in rec:
                    rec[d] = {
                        "empfohlen": v["empfohlen"],
                        "plan": v.get("plan"),
                        "zone": v.get("zone", ""),
                        "quelle": "plan-briefing",
                        "stand": v.get("stand"),
                    }
        except Exception as e:
            log(f"treffer: FEHLER beim Laden von backfill_file – {e}")

    treffer_done = st.get("treffer_done", [])
    alle_datum = sorted(rec.keys())
    neu = 0

    for datum in alle_datum:
        if datum in treffer_done:
            continue
        rec_eintrag = rec[datum]
        t = treffer_fuer(datum, rec_eintrag, rows_by_date)
        if t is None:
            continue

        ok = save_extra("treffer_tag", datum, t["text"], t, t.get("quelle", "unknown"))
        if ok:
            st.setdefault("treffer", {})
            treffer_dict = st.get("treffer", {})
            treffer_dict[datum] = t
            # ganze Historie behalten (Obergrenze nur als Schutz)
            if len(treffer_dict) > 1000:
                keys_sorted = sorted(treffer_dict.keys(), reverse=True)[:1000]
                st["treffer"] = {k: treffer_dict[k] for k in keys_sorted}

            if datum not in st.get("treffer_done", []):
                treffer_done.append(datum)
            st["treffer_done"] = treffer_done[-1000:]
            neu += 1
            save_x(st)

    save_x(st)
    if neu:
        log(f"treffer: {neu} neue Tagesauswertung(en) gespeichert")


# ---------------------------------------------------------------------------
# F) TREFFSICHERHEIT WOCHE
# ---------------------------------------------------------------------------


def cmd_treffer_woche(force: bool = False) -> None:
    """Erstellt eine Wochenbilanz aus den letzten Treffer-Einträgen."""
    st = load_x()
    heute = today()

    # nur wenn force oder (Montag und st['woche_done'] != heute)
    if not force:
        wday = datetime.datetime.now(tz_local).weekday()  # 0=Mo
        if wday != 0:
            pass
            return
        if st.get("woche_done") == heute:
            log("treffer_woche: bereits heute ausgeführt – übersprungen")
            return

    treffer_dict = st.get("treffer", {})
    jetzt = datetime.datetime.now(tz_local)

    # Letzte 28 Tage sammeln
    werte = []
    for d_str, t in treffer_dict.items():
        try:
            dt = datetime.datetime.strptime(d_str, "%Y-%m-%d").replace(tzinfo=jetzt.tzinfo)
        except (ValueError, TypeError):
            continue
        if (jetzt - dt).days <= 28:
            werte.append((d_str, t))

    if len(werte) < 5:
        log(f"treffer_woche: nur {len(werte)} Treffer vorhanden (mind. 5 nötig) – abgebrochen")
        return

    werte.sort(key=lambda x: x[0])

    # Statistik
    n = len(werte)
    n_gefolgt = sum(1 for _, t in werte if t.get("gefolgt"))
    scores_gefolgt = [t.get("score", 0) for _, t in werte if t.get("gefolgt")]
    scores_nicht = [t.get("score", 0) for _, t in werte if not t.get("gefolgt")]
    avg_gefolgt = round(sum(scores_gefolgt) / len(scores_gefolgt), 2) if scores_gefolgt else 0
    avg_nicht = round(sum(scores_nicht) / len(scores_nicht), 2) if scores_nicht else 0
    abweichungen = [t.get("abweichung_pct", 0) for _, t in werte if t.get("abweichung_pct") is not None]
    avg_abw = round(sum(abweichungen) / len(abweichungen), 2) if abweichungen else 0

    schlechteste = sorted(werte, key=lambda x: x[1].get("score", 0))[:3]
    beste = sorted(werte, key=lambda x: x[1].get("score", 0), reverse=True)[:3]

    statistik = {
        "n": n,
        "n_gefolgt": n_gefolgt,
        "mittel_score_gefolgt": avg_gefolgt,
        "mittel_score_nicht": avg_nicht,
        "mittel_abweichung_pct": avg_abw,
        "schlechteste_3": schlechteste,
        "beste_3": beste,
    }

    # LLM-Analyse
    tabelzeile = ""
    for d, t in werte:
        tabelzeile += (
            f"{d}, {t.get('empfohlen', '?')}, {t.get('ist', '?')}, "
            f"{t.get('abweichung_pct', '')}%, {t.get('gefolgt')}, {t.get('score', '?')}\n"
        )

    user_text = (
        f"Statistik: n={n}, n_gefolgt={n_gefolgt}, "
        f"Ø Score gefolgt={avg_gefolgt}, Ø Score nicht={avg_nicht}, "
        f"Ø Abweichung={avg_abw}%.\n\nTabelle:\n{tabelzeile}\n"
        f"Schlechteste 3: {', '.join(d for d, _ in schlechteste)}\n"
        f"Beste 3: {', '.join(d for d, _ in beste)}\n\n"
        f"Aufgabe: Wochenbilanz in 3 bis 5 Stichpunkten (max. 120 Wörter)."
    )

    msgs = [
        {"role": "system", "content": (
            "Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, sachlich, "
            "keine Emojis, keine erfundenen Zahlen."
        )},
        {"role": "user", "content": user_text},
    ]

    try:
        text_out, _ = ollama_chat(MODEL, msgs, num_ctx=8192, timeout=600)
        # sane-Prüfung
        text_out = text_out.strip()
        if not text_out or len(text_out) < 80:
            log(f"treffer_woche: LLM-Antwort zu kurz ({len(text_out)} Zeichen) – abgebrochen")
            return
    except Exception as e:
        log(f"treffer_woche: LLM-FEHLER – {e}")
        return

    save_extra(
        "treffer_woche", heute, text_out,
        {"statistik": statistik},
        "mac:" + MODEL,
    )
    st["woche_done"] = heute
    save_x(st)
    log(f"treffer_woche: Wochenbilanz für {heute} gespeichert.")


# ---------------------------------------------------------------------------
# G) ERNÄHRUNG FÜR MORGEN
# ---------------------------------------------------------------------------


def cmd_ernaehrung(force: bool = False) -> None:
    """Erstellt Ernährungsplan für morgen."""
    st = load_x()
    heute = today()

    # 17 <= now_h() < 23 und st['ern_done'] != heute
    if not force:
        nh = now_h()
        if not (17 <= nh < 23):
            pass  # außerhalb 17–23 Uhr still
            return
        if st.get("ern_done") == heute:
            pass
            return

    morgen = (datetime.datetime.now(tz_local) + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

    # Plan morgen laden
    plan = http_get_json(SCRIPT_URL + "?mode=planJsonV2").get("plan") or []
    plan_morgen = None
    for p in plan:
        if p.get("date") == morgen:
            plan_morgen = p
            break

    last = 0
    sport = ""
    zone = ""
    if plan_morgen:
        last = plan_morgen.get("recommended_load_ess") or plan_morgen.get("original_load_ess") or 0
        sport = plan_morgen.get("training_1", "")
        zone = plan_morgen.get("recommended_zone", "")
    else:
        last = 0

    # Baseline laden
    try:
        baseline_data = http_get_json(SCRIPT_URL + "?mode=physBaseline").get("baseline", {})
    except Exception:
        baseline_data = {}

    gewicht = 76
    grundumsatz = 2028
    for k, v in baseline_data.items():
        if "Weight" in k and gewicht == 76:
            gewicht = v
        if "Resting_Calories" in k and grundumsatz == 2028:
            grundumsatz = v

    # Energiemodell: Regression kcal_out = a + b * load
    rows = fetch_rows(42, ["kcal_out", "load_fb_day"])
    xy_pairs: list[tuple[float, float]] = []
    for r in rows:
        ko = num(r.get("kcal_out"))
        lf = num(r.get("load_fb_day"))
        if ko is not None and lf is not None:
            xy_pairs.append((lf, ko))

    if len(xy_pairs) >= 10:
        x_m = sum(p[0] for p in xy_pairs) / len(xy_pairs)
        y_m = sum(p[1] for p in xy_pairs) / len(xy_pairs)
        num_b = sum((p[0] - x_m) * (p[1] - y_m) for p in xy_pairs)
        den_b = sum((p[0] - x_m) ** 2 for p in xy_pairs)
        if den_b != 0:
            b = num_b / den_b
            a = y_m - b * x_m
        else:
            a = grundumsatz * 1.15
            b = 4.5
    else:
        a = grundumsatz * 1.15
        b = 4.5

    kcal_ziel = round(a + b * last, -1) if last else 0

    # Makroziele
    kg = gewicht
    protein_kg = 1.6 if last < 80 else (1.8 if last < 180 else 2.0)
    if last < 40:
        carb_kg = 3
    elif last < 120:
        carb_kg = 4.5
    elif last < 220:
        carb_kg = 6
    else:
        carb_kg = 7.5

    protein_g = round(kg * protein_kg)
    carbs_g = round(kg * carb_kg)

    fett_g: float | None
    if kcal_ziel > 0 and protein_g + carbs_g > 0:
        fett_g = round((kcal_ziel - 4 * protein_g - 4 * carbs_g) / 9)
    fett_g = max(round(0.8 * kg), fett_g if fett_g is not None else 0)

    # Heute-Bilanz (Ist heute gegen das Ziel für die heutige Last)
    bilanz: dict | None = None
    try:
        rows_h = fetch_rows(7, ["kcal_in", "load_fb_day", "protein_g", "carb_g"])
        row_h = next((r for r in rows_h if r.get("date") == today()), None)
    except Exception:
        row_h = None
    if row_h:
        kcal_in_h = num(row_h.get("kcal_in"))
        load_h = num(row_h.get("load_fb_day")) or 0.0
        if kcal_in_h is not None:
            bilanz = {
                "kcal_in": kcal_in_h,
                "kcal_ziel": round(a + b * load_h, -1),
                "protein": num(row_h.get("protein_g")),
                "protein_ziel": round(kg * (1.6 if load_h < 80 else (1.8 if load_h < 180 else 2.0))),
                "kh": num(row_h.get("carb_g")),
                "last": load_h,
            }

    # LLM aufrufen
    user_text = (
        f"Ziele morgen: {kcal_ziel} kcal · Protein {protein_g} g · KH {carbs_g} g · Fett {fett_g} g\n"
        f"Training morgen: Last={last}, Sport={sport}, Zone={zone}\n"
        f"Heute Ist: {json.dumps(bilanz)}\n"
        f"Vorlieben: Familie (2 Erwachsene, Kind 4 Jahre), mag Fleischgerichte, Kartoffelknödel, "
        f"Sushi, Gerichte im HelloFresh-Stil, unter der Woche schnell (max. 30 min).\n"
        f"Aufgabe: JSON mit 'kurz' (1-2 Sätze), 'mahlzeiten' (4-5: wann, gericht, warum), 'tipp'."
    )

    msgs = [
        {"role": "system", "content": (
            "Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, sachlich, "
            "keine Emojis, keine erfundenen Zahlen."
        )},
        {"role": "user", "content": user_text},
    ]

    try:
        llm_out, _ = ollama_chat(MODEL, msgs, num_ctx=8192, timeout=600)
        llm_data = parse_loose(llm_out)
        kurz = llm_data.get("kurz", "")
        mahlzeiten = llm_data.get("mahlzeiten", [])
        tipp = llm_data.get("tipp", "")

        # bauen
        lines_ern = [f"Ziel morgen ({last} ESS, {sport}): {kcal_ziel} kcal · Protein {protein_g} g · KH {carbs_g} g · Fett {fett_g} g"]
        if kurz:
            lines_ern.append(kurz)
        for mh in mahlzeiten:
            wann = mh.get("wann", "")
            gericht = mh.get("gericht", "")
            lines_ern.append(f"- {wann}: {gericht}")
        if tipp:
            lines_ern.append(f"Tipp: {tipp}")

        text = "\n".join(lines_ern)
    except Exception as e:
        log(f"ernaehrung: LLM-FEHLER – {e}")
        return

    save_extra(
        "ernaehrung", morgen, text,
        {
            "ziele": {"kcal": kcal_ziel, "protein": protein_g, "carbs": carbs_g, "fett": fett_g},
            "last": last,
            "sport": sport,
            "zone": zone,
            "modell": {"a": round(a, 2), "b": round(b, 2)},
            "heute": bilanz,
            "mahlzeiten": mahlzeiten,
            "kurz": kurz,
            "tipp": tipp,
        },
        "mac:" + MODEL,
    )
    st["ern_done"] = heute
    save_x(st)
    log(f"ernaehrung: Plan für {morgen} gespeichert.")


# ---------------------------------------------------------------------------
# H) CLI – main() ERWEITERT
# ---------------------------------------------------------------------------


# ============================================================
# Zukunftsplanung: Monotonie-Prüfung, Wochen-Autopilot, Störungen
# ============================================================

def ref_load(loads: list) -> float:
    xs = [float(x) for x in loads[-28:]]
    return sum(xs) / len(xs) if xs else 0.0


def is_demanding(l: float, ref: float) -> bool:
    return l >= max(60.0, 0.7 * ref)


RELIEF_ANTEIL = 0.6  # Entlastung = lockerer oder freier Tag: <= 40 ESS bzw. <= 60 % der Referenz-Tageslast


def is_relief(l: float, ref: float) -> bool:
    return l <= max(40.0, RELIEF_ANTEIL * ref)


def mono7(w: list) -> float:
    w = [float(x) for x in w[-7:]]
    if len(w) < 7:
        return 0.0
    m = sum(w) / 7
    sd = (sum((x - m) ** 2 for x in w) / 7) ** 0.5
    return (9.99 if m > 0 else 0.0) if sd == 0 else m / sd


def _k(v: float) -> str:
    return f"{v:.2f}".replace(".", ",")


def mono_verstoesse(hist: list, plan: list) -> list:
    """Monotonie-Regel (wie lib/monotony.ts, monotonyPlanViolations) für einen Plan [{'datum','last','fest'}]."""
    allv = [float(x) for x in hist]
    ref = ref_load(allv + [float(d["last"]) for d in plan])
    out = []
    streak = 0
    for x in reversed(allv):
        if is_relief(x, ref):
            break
        streak += 1
    run25 = 0
    for i, d in enumerate(plan):
        last = float(d["last"])
        prev_m = mono7(allv) if len(allv) >= 7 else 0.0  # Monotonie des Vortags
        allv.append(last)
        streak = 0 if is_relief(last, ref) else streak + 1
        m = mono7(allv)
        run25 = run25 + 1 if m > 2.5 else 0
        nxt = [plan[j] for j in (i + 1, i + 2) if j < len(plan)]
        next_relief = any(is_relief(float(n["last"]), ref) for n in nxt)
        bewertbar = i + 2 < len(plan)
        # 5 Tage ohne Entlastung: kein Verstoß mehr, nur Hinweis im Tagesvotum (29.10.26)
        if bewertbar and m > 3.0 and not next_relief:
            out.append(f"{d['datum']}: Monotonie {_k(m)} > 3,0 ohne Entlastung in 48 h")
        elif bewertbar and run25 == 2 and not next_relief:
            out.append(f"{d['datum']}: Monotonie seit 2 Tagen > 2,5 ohne Ruhetag in 48 h")
        if i + 1 < len(plan) and m > 2.0 and prev_m > 2.0:  # ein einzelner Tag über 2,0 ist bedeutungslos
            dem = 0
            for x in reversed(allv):
                if is_relief(x, ref):
                    break
                if is_demanding(x, ref):
                    dem += 1
            n = plan[i + 1]
            nl = float(n["last"])
            if dem >= 4 and not n.get("fest") and not is_relief(nl, ref):
                out.append(f"{n['datum']}: sollte nach 2 Tagen Monotonie über 2,0 ({_k(m)}) und {dem} fordernden Tagen sehr locker oder frei sein")
    return out


_WTAG = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]


def _wetter(tage: int = 9) -> dict:
    try:
        w = http_get_json(
            "https://api.open-meteo.com/v1/forecast?latitude=48.31&longitude=10.12"
            "&daily=temperature_2m_max,precipitation_sum,wind_speed_10m_max&timezone=Europe%2FBerlin"
            f"&forecast_days={tage}"
        )["daily"]
        return {d: {"tmax": w["temperature_2m_max"][i], "regen": w["precipitation_sum"][i], "wind": w["wind_speed_10m_max"][i]}
                for i, d in enumerate(w["time"])}
    except Exception as e:
        log(f"autopilot: Wetter nicht verfügbar ({e})")
        return {}


def _kalender(tage: int = 9) -> dict:
    try:
        return {d["date"]: d for d in http_get_json(SCRIPT_URL + f"?mode=calendarDays&days={tage}").get("days", [])}
    except Exception as e:
        log(f"autopilot: Kalender nicht verfügbar ({e})")
        return {}


def cmd_autopilot(force: bool = False) -> None:
    """Sonntagabend: Plan der neuen Woche prüfen und Verbesserungen vorschlagen."""
    st = load_x()
    jetzt = datetime.datetime.now(tz_local)
    if not force and not (jetzt.weekday() == 6 and 17 <= jetzt.hour < 23 and st.get("autopilot_done") != today()):
        return
    if not ollama_ready():
        log("autopilot: Ollama nicht bereit – später")
        return
    tage_bis_mo = (7 - jetzt.weekday()) % 7 or 7
    montag_d = (jetzt + datetime.timedelta(days=tage_bis_mo)).date()
    woche_d = [montag_d + datetime.timedelta(days=i) for i in range(7)]
    woche = [d.isoformat() for d in woche_d]
    plan_all = {d["date"]: d for d in http_get_json(SCRIPT_URL + "?mode=planSimulationJson").get("days", [])}
    plan = []
    for i, d in enumerate(woche):
        p = plan_all.get(d) or {}
        plan.append({"datum": d, "tag": _WTAG[i], "last": round(float(p.get("load") or 0)), "sport": p.get("sport") or "Off",
                     "zone": p.get("zone") or "Off", "fest": bool(p.get("locked"))})
    rows = {r["date"]: r for r in fetch_rows(35, ["load_fb_day"])}
    hist = []
    for k in range(28, 0, -1):
        d = (montag_d - datetime.timedelta(days=k)).isoformat()
        if d in rows and num(rows[d].get("load_fb_day")) is not None and d < today():
            hist.append(num(rows[d].get("load_fb_day")))
        elif d in plan_all:
            hist.append(float(plan_all[d].get("load") or 0))
        else:
            hist.append(0.0)
    # Tagesobergrenze aus der eigenen Historie (95. Perzentil der Trainingstage, letzte 360 Tage)
    try:
        lang = sorted(v for v in (num(r.get("load_fb_day")) for r in fetch_rows(360, ["load_fb_day"])) if v and v > 0)
        spitze = round(lang[int(len(lang) * 0.95)]) if len(lang) >= 30 else 240
    except Exception:
        spitze = 240

    # Wochenrhythmus (weiche Präferenz): Ø Trainingsbereitschaft je Wochentag, letzte 180 Tage
    rhythmus = ""
    try:
        wt = {}
        for r in fetch_rows(180, ["Garmin_Training_Readiness"]):
            v = num(r.get("Garmin_Training_Readiness"))
            if v is not None:
                wt.setdefault(datetime.date.fromisoformat(r["date"]).weekday(), []).append(v)
        mw = {k: sum(v) / len(v) for k, v in wt.items() if len(v) >= 8}
        if len(mw) == 7:
            ordn = sorted(mw, key=lambda k: -mw[k])
            gut = ", ".join(_WTAG[k] for k in ordn[:3])
            schwach = _WTAG[ordn[-1]]
            rhythmus = (f"Weiche Präferenz aus deinem Wochenrhythmus: Bereitschaft im Schnitt am höchsten {gut}, am niedrigsten {schwach} "
                        f"({round(mw[ordn[0]])} gegen {round(mw[ordn[-1]])}). Qualität eher auf die starken Tage, den schwächsten Tag eher locker – nur wenn Kalender und Regeln es zulassen. ")
    except Exception:
        rhythmus = ""

    def tag_max(alt: int) -> int:
        """Höchstens +40 % je Tag und nie über die eigene Spitze (außer der Plan lag schon darüber)."""
        if alt <= 0:
            return round(spitze * 0.7)
        return max(alt, min(spitze, round(alt * 1.4)))

    kal, wet = _kalender(), _wetter()
    ziel_ess, saison_txt, ziele_txt = None, "", []
    try:
        items = http_get_json(SCRIPT_URL + "?mode=extras&typ=saison,ziel&n=20").get("items", [])
        sa = next((x for x in items if x["typ"] == "saison"), None)
        if sa and sa.get("data"):
            w = next((x for x in sa["data"].get("wochen", []) if x.get("start") == woche[0]), None)
            if w:
                ziel_ess = w.get("ziel_ess")
                saison_txt = (f"Saisonrahmen: Phase {w.get('phase')}, Ziel-Wochenlast ca. {ziel_ess} ESS, Fokus {w.get('fokus') or w.get('sportmix') or '-'}"
                              + (f", Zeitfenster: {w.get('fenster')}" if w.get("fenster") else ""))
                if w.get("phase") == "Pause":
                    saison_txt += " – Marc hat in dieser Woche kaum Zeit (Zeitfenster „kein Training möglich“): nur kurze lockere Einheiten an freien Tagen, keine harten Reize."
                elif w.get("phase") == "Aktivurlaub":
                    saison_txt += " – Aktivurlaub: lange Tage (Bergtouren, Ski) sind gewollt, dazwischen lockere Tage; Monotonie- und ACWR-Regeln gelten weiter."
                elif w.get("phase") == "Entlastung" and w.get("fenster"):
                    saison_txt += " – Entlastungswoche passend zum Zeitfenster: locker, Schlaf und Erholung im Blick."

        ziele_txt = [f"{x['datum']}: {x['text'][:160]}" for x in items if x["typ"] == "ziel" and x["datum"] >= woche[0]][:3]
    except Exception as e:
        log(f"autopilot: Saison/Ziele nicht lesbar ({e})")
    v0 = mono_verstoesse(hist, plan)

    def zeile(p):
        k = kal.get(p["datum"]) or {}
        w = wet.get(p["datum"]) or {}
        kt = f"{k.get('termine', '?')} Termine, {k.get('belegt_h', '?')} h belegt{', abends' if k.get('abends') else ''}{', REISE' if k.get('reise') else ''}" if k else "unbekannt"
        wt = f"{w.get('tmax')}°, Regen {w.get('regen')} mm, Wind {w.get('wind')} km/h" if w else "unbekannt"
        return f"{p['datum']} {p['tag']} | Plan: {p['last']} ESS {p['sport']} {p['zone']}{' FEST' if p['fest'] else ''} | Kalender: {kt} | Wetter: {wt}"

    regeln = ("Regeln: FEST-Tage nie ändern. "
              + (f"Wochenlast höchstens 10 % über {ziel_ess} ESS, gern bis 15 % darunter. " if ziel_ess is not None else "Wochenlast etwa beibehalten (15 % weniger ist in Ordnung, höchstens 10 % mehr). ")
              + rhythmus
              + f"Kein Tag über {spitze} ESS und kein Tag mehr als 40 % über seinem bisherigen Plan. Gestrichene Last nicht auf einen Tag packen: auf zwei bis drei Tage verteilen oder die Wochenlast senken. "
              + "Monotonie-Regel: als Entlastung zählt ein lockerer oder freier Tag (<= 40 ESS bzw. 60 % der üblichen Tageslast); 5 Tage ohne Entlastung sind nur ein Hinweis; nach 2 Tagen Monotonie über 2,0 und 4 fordernden Tagen ein sehr lockerer oder freier Tag; harte Tage klar hart, lockere wirklich locker. "
              "Lange Einheiten auf Tage mit wenig Terminen und gutem Wetter. Rad draußen nicht bei Regen > 5 mm oder Wind > 40 km/h. Nur nötige Änderungen.")
    beispiel = '{"tage":[{"datum":"' + woche[0] + '","last":130,"sport":"Bike","zone":"Z2","grund":"kurz"}],"fazit":"1-2 Sätze"}'
    user = ("Plan der Woche ab " + woche[0] + ":\n" + "\n".join(zeile(p) for p in plan) + "\n\n"
            + (saison_txt + "\n" if saison_txt else "") + ("Ziele: " + "; ".join(ziele_txt) + "\n" if ziele_txt else "")
            + ("Gefundene Regelverstöße: " + "; ".join(v0) + "\n" if v0 else "Keine Regelverstöße gefunden.\n")
            + regeln + "\nAntworte nur mit JSON in diesem Format, tage enthält ALLE 7 Tage: " + beispiel)
    msgs = [{"role": "system", "content": "Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, keine Emojis. Antworte nur mit JSON."},
            {"role": "user", "content": user}]

    def anwenden(antwort: str):
        j = parse_loose(antwort)
        by = {str(x.get("datum", ""))[:10]: x for x in (j.get("tage") or [])}
        neu = []
        for p in plan:
            x = by.get(p["datum"])
            if p["fest"] or not x:
                neu.append(dict(p, grund=""))
                continue
            try:
                last = max(0, min(tag_max(p["last"]), round(float(x.get("last", p["last"])))))
            except Exception:
                last = p["last"]
            neu.append(dict(p, last=last, sport=str(x.get("sport") or p["sport"]), zone=str(x.get("zone") or p["zone"]), grund=str(x.get("grund") or "")))
        return neu, str(j.get("fazit") or "")

    t0 = time.time()
    txt, _ = ollama_chat(MODEL, msgs, num_ctx=12288, timeout=900)
    neu, fazit = anwenden(txt)
    v1 = mono_verstoesse(hist, neu)
    if v1:
        msgs2 = msgs + [{"role": "assistant", "content": txt},
                        {"role": "user", "content": "Dein Plan verletzt noch: " + "; ".join(v1) + ". Korrigiere und gib das vollständige JSON erneut aus."}]
        try:
            txt2, _ = ollama_chat(MODEL, msgs2, num_ctx=12288, timeout=900)
            neu2, fazit2 = anwenden(txt2)
            v2 = mono_verstoesse(hist, neu2)
            if len(v2) < len(v1):
                neu, fazit, v1 = neu2, fazit2, v2
        except Exception as e:
            log(f"autopilot: Nachfrage fehlgeschlagen ({e})")
    aend = [{"datum": n["datum"], "tag": n["tag"], "alt": {"last": p["last"], "sport": p["sport"], "zone": p["zone"]},
             "neu": {"last": n["last"], "sport": n["sport"], "zone": n["zone"]}, "grund": n["grund"]}
            for p, n in zip(plan, neu) if (p["last"], p["sport"], p["zone"]) != (n["last"], n["sport"], n["zone"])]
    wa, wn = sum(p["last"] for p in plan), sum(n["last"] for n in neu)
    zeilen = [f"Plan für die Woche ab {woche[0]}: {len(aend)} Änderung(en), Wochenlast {wa} → {wn} ESS.", fazit]
    zeilen += [f"- {a['tag']} {a['datum'][8:10]}.{a['datum'][5:7]}.: {a['alt']['last']} → {a['neu']['last']} ESS {a['neu']['sport']} {a['neu']['zone']} – {a['grund']}" for a in aend]
    if v1:
        zeilen.append("Noch offen: " + "; ".join(v1))
    save_extra("autopilot", woche[0], "\n".join(z for z in zeilen if z),
               {"tage": neu, "aenderungen": aend, "verstoesse_vorher": v0, "verstoesse": v1, "wochenlast_alt": wa,
                "wochenlast_neu": wn, "saison_ziel": ziel_ess, "fazit": fazit, "sekunden": round(time.time() - t0)}, "mac:" + MODEL)
    notify("Coach Kira", f"Plan für nächste Woche bereit: {len(aend)} Änderung(en)")
    st["autopilot_done"] = today()
    save_x(st)
    log(f"autopilot: Woche ab {woche[0]} – {len(aend)} Änderung(en), Verstöße {len(v0)} → {len(v1)}")


def cmd_stoerung(force: bool = False) -> None:
    """Morgens: Infekt-Verdacht, starke Müdigkeit, ausgefallenen Tag oder Reise erkennen."""
    st = load_x()
    if not force and (st.get("stoerung_done") == today() or not (6 <= now_h() < 22)):
        return
    rows = fetch_rows(21, ["hrv_status", "hrv_threshholds", "rhr_bpm", "sleep_score_0_100", "Garmin_Training_Readiness", "load_fb_day", "befinden_1_5"])
    by = {r["date"]: r for r in rows}
    h = by.get(today())
    if not h or num(h.get("hrv_status")) is None:
        return  # Morgenwerte noch nicht da
    vor = [r for r in rows if r["date"] < today()]
    befunde = []
    hrv, rhr = num(h.get("hrv_status")), num(h.get("rhr_bpm"))
    unten = None
    try:
        unten = float(str(h.get("hrv_threshholds") or "").split(";")[0].replace(",", "."))
    except Exception:
        pass
    h7 = [num(r.get("hrv_status")) for r in vor[-7:] if num(r.get("hrv_status")) is not None]
    r14 = sorted(num(r.get("rhr_bpm")) for r in vor[-14:] if num(r.get("rhr_bpm")) is not None)
    rmed = r14[len(r14) // 2] if r14 else None
    hrv_tief = hrv is not None and ((unten is not None and hrv < unten) or (h7 and hrv < 0.85 * sum(h7) / len(h7)))
    rhr_hoch = rhr is not None and rmed is not None and rhr >= rmed + 5
    bef = num(h.get("befinden_1_5"))
    if hrv_tief and rhr_hoch:
        befunde.append({"art": "infekt", "schwere": "hoch", "text": f"Infekt-Verdacht: HRV {hrv:.0f} unter Normalband/Schnitt, Ruhepuls {rhr:.0f} (Median {rmed:.0f})."})
    elif bef is not None and bef <= 2 and (hrv_tief or rhr_hoch):
        befunde.append({"art": "infekt", "schwere": "hoch", "text": f"Befinden {bef:.0f}/5 und auffällige Morgenwerte – Infekt möglich."})
    tr, sl = num(h.get("Garmin_Training_Readiness")), num(h.get("sleep_score_0_100"))
    if not befunde and tr is not None and sl is not None and tr <= 35 and sl < 60:
        befunde.append({"art": "muede", "schwere": "mittel", "text": f"Sehr müde: Trainingsbereitschaft {tr:.0f}, Schlafscore {sl:.0f}."})
    gestern = (datetime.datetime.now(tz_local).date() - datetime.timedelta(days=1)).isoformat()
    try:
        rv = http_get_json(SCRIPT_URL + "?mode=activityReviews&n=5").get("reviews", [])
        g = next((x for x in rv if x.get("date") == gestern), None)
        soll = num(g.get("loadSoll")) if g else None
        ist = num((by.get(gestern) or {}).get("load_fb_day"))
        if soll is not None and soll >= 60 and (ist is None or ist <= 10):
            befunde.append({"art": "ausfall", "schwere": "mittel", "text": f"Gestern {soll:.0f} ESS geplant, nicht trainiert."})
    except Exception as e:
        log(f"stoerung: Aktivitäten nicht lesbar ({e})")
    kal = _kalender(7)
    reise = [d for d, k in sorted(kal.items()) if d > today() and (k.get("reise") or (k.get("belegt_h") or 0) >= 10)]
    if reise:
        befunde.append({"art": "reise", "schwere": "niedrig", "text": "Reise/sehr volle Tage: " + ", ".join(d[8:10] + "." + d[5:7] + "." for d in reise)})
    st["stoerung_done"] = today()
    save_x(st)
    if not befunde:
        return
    text = "\n".join(b["text"] for b in befunde)
    save_extra("umplanung", today(), text, {"befunde": befunde, "werte": {"hrv": hrv, "rhr": rhr, "readiness": tr, "schlaf": sl, "befinden": bef}}, "mac-regeln")
    wichtig = [b for b in befunde if b["schwere"] in ("hoch", "mittel")]
    if wichtig:
        notify("Coach Kira", "Störung erkannt: " + wichtig[0]["text"][:80] + " – Rest-Plan im Cockpit anpassen")
    log(f"stoerung: {len(befunde)} Befund(e) – {', '.join(b['art'] for b in befunde)}")



# ---------------------------------------------------------------------------
# Vorausblick, Tourfenster, Frühwarnung (morgens vor den Briefings, abends erneut)
# ---------------------------------------------------------------------------

_ATL = (-8.819849, 0.83645661, 1.24059231)
_CTL = (-3.95042, 0.97283824, 0.2336734, 0.7, 7.1)


def _atl_step(atl: float, load: float) -> float:
    return _ATL[0] + _ATL[1] * atl + _ATL[2] * load


def _ctl_step(ctl: float, load: float, mean28: float) -> float:
    a, b, c, g, k = _CTL
    return g * (a + b * ctl + c * load) + (1 - g) * k * mean28


_ES = (155.91432, 0.96845, 0.1537, 0.2993)  # wie ES_MODEL in lib/ctlModel.ts


def _es_step(es: float, load: float, mean56: float) -> float:
    a, b, c, d = _ES
    return a + b * es + c * load + d * mean56


def _mean(xs: list) -> float | None:
    xs = [float(x) for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else None


def _slope_per_week(ys: list) -> float | None:
    pts = [(i, float(y)) for i, y in enumerate(ys) if y is not None]
    if len(pts) < 7:
        return None
    mx = sum(p[0] for p in pts) / len(pts)
    my = sum(p[1] for p in pts) / len(pts)
    den = sum((p[0] - mx) ** 2 for p in pts) or 1.0
    return 7.0 * sum((p[0] - mx) * (p[1] - my) for p in pts) / den


# --- Bereitschafts-Modell (wie lib/readinessModel.ts, Ridge-Regression) -------

def _feats(load, prev, mean7, acwr, mono, start):
    return [load, prev, mean7, acwr, min(mono, 4.0), start, 1.0 if load <= 20 else 0.0]


def _solve(A, b):
    n = len(b)
    M = [row[:] + [b[i]] for i, row in enumerate(A)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(M[r][c]))
        M[c], M[p] = M[p], M[c]
        d = M[c][c] or 1e-9
        M[c] = [v / d for v in M[c]]
        for r in range(n):
            if r != c:
                f_ = M[r][c]
                M[r] = [M[r][k] - f_ * M[c][k] for k in range(n + 1)]
    return [M[r][n] for r in range(n)]


def _fit(X, y, lam=3.0):
    p = len(X[0])
    mean = [sum(r[j] for r in X) / len(X) for j in range(p)]
    sd = [((sum((r[j] - mean[j]) ** 2 for r in X) / len(X)) ** 0.5) or 1.0 for j in range(p)]
    Z = [[1.0] + [(r[j] - mean[j]) / sd[j] for j in range(p)] for r in X]
    k = p + 1
    A = [[sum(z[i] * z[j] for z in Z) + (lam if i == j and i > 0 else 0.0) for j in range(k)] for i in range(k)]
    bb = [sum(z[i] * y[t] for t, z in enumerate(Z)) for i in range(k)]
    return {"w": _solve(A, bb), "mean": mean, "sd": sd}


def _pred(m, x):
    return m["w"][0] + sum(m["w"][j + 1] * ((v - m["mean"][j]) / m["sd"][j]) for j, v in enumerate(x))


def _bereitschaft_modell(hist: list) -> dict | None:
    """hist: [{'date','load','readiness','acwr','mono'}] chronologisch."""
    X, y, naive = [], [], []
    for i in range(7, len(hist) - 1):
        d, n = hist[i], hist[i + 1]
        if n["readiness"] is None or d["readiness"] is None or d["load"] is None:
            continue
        prev = hist[i - 1]["load"] or 0.0
        w7 = [r["load"] or 0.0 for r in hist[i - 6:i + 1]]
        X.append(_feats(d["load"], prev, sum(w7) / 7, d["acwr"] or 1.0, d["mono"] or 1.5, d["readiness"]))
        y.append(n["readiness"])
        naive.append(d["readiness"])
    if len(X) < 30:
        return None
    cut = int(len(X) * 0.8)
    mt = _fit(X[:cut], y[:cut])
    err = [y[cut + i] - _pred(mt, X[cut + i]) for i in range(len(X) - cut)]
    mae = sum(abs(e) for e in err) / len(err)
    base = sum(abs(y[cut + i] - naive[cut + i]) for i in range(len(err))) / len(err)
    m = _fit(X, y)
    m.update({"n": len(X), "mae": mae, "naiv": base})
    return m


def _wetter_tage(tage: int = 10) -> dict:
    try:
        w = http_get_json(
            "https://api.open-meteo.com/v1/forecast?latitude=48.31&longitude=10.12"
            "&daily=temperature_2m_max,precipitation_sum,precipitation_probability_max,wind_speed_10m_max,sunshine_duration"
            f"&timezone=Europe%2FBerlin&forecast_days={tage}"
        )["daily"]
        out = {}
        for i, d in enumerate(w["time"]):
            out[d] = {"tmax": w["temperature_2m_max"][i], "regen": w["precipitation_sum"][i],
                      "regen_p": (w.get("precipitation_probability_max") or [None] * 99)[i],
                      "wind": w["wind_speed_10m_max"][i],
                      "sonne_h": round(((w.get("sunshine_duration") or [0] * 99)[i] or 0) / 3600, 1)}
        return out
    except Exception as e:
        log(f"vorausblick: Wetter nicht verfügbar ({e})")
        return {}


def _dm(d: str) -> str:
    dt = datetime.date.fromisoformat(d)
    return f"{_WTAG[dt.weekday()]} {dt.day:02d}.{dt.month:02d}."


def _de(v: float, nd: int = 0) -> str:
    return (f"{v:.{nd}f}").replace(".", ",")


# ---------------------------------------------------------------------------
# Garmin Personal API: Signale für Frühwarnung und Lage für den Prompt
# (Daten: AI_KIRA_EXTRAS typ garmin_tag, geschrieben von kira_garmin.py)
# ---------------------------------------------------------------------------
def garmin_tage(n: int = 60) -> dict:
    """Garmin-Tagesdetails als {datum: data}. Fehler -> {}."""
    try:
        j = http_get_json(SCRIPT_URL + f"?mode=extras&typ=garmin_tag&n={n}")
        return {it["datum"]: it.get("data") or {} for it in (j.get("items") or []) if it.get("datum")}
    except Exception as e:
        log(f"vorausblick: Garmin-Tage nicht verfügbar ({e})")
        return {}


def _g_tage(heute: str, von_offset: int, bis_offset: int) -> list:
    """Datumsliste heute-von_offset … heute-bis_offset (von_offset >= bis_offset)."""
    h = datetime.date.fromisoformat(heute)
    return [(h - datetime.timedelta(days=k)).isoformat() for k in range(von_offset, bis_offset - 1, -1)]


def _g_get(data: dict, *pfad):
    x = data
    for p in pfad:
        if not isinstance(x, dict):
            return None
        x = x.get(p)
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else None


def garmin_signale(gt: dict, heute: str) -> tuple:
    """Frühwarn-Signale aus Garmin-Tageswerten gegen den eigenen Schnitt (Tag -35 … -8)."""
    sig: list = []
    lage: dict = {}
    if not gt:
        return sig, lage

    def bb_auf(d):
        g = gt.get(d) or {}
        v = _g_get(g, "bb_aufwachen")
        return v if v is not None else _g_get(g, "tag", "bb_max")

    def werte(tage, f):
        return [v for v in (f(d) for d in tage) if v is not None]

    def med(xs):
        return _median(xs) if xs else None

    basis = _g_tage(heute, 35, 8)
    l7 = _g_tage(heute, 6, 0)
    l3 = _g_tage(heute, 2, 0)
    vorwoche = _g_tage(heute, 7, 1)

    def pruefe(name, f, aktuell, grenze_o, grenze_g, richtung, text, einheit="", nd=0, extra=None):
        b = werte(basis, f)
        a = werte(aktuell, f)
        if len(b) < 10 or len(a) < 2:
            return
        bm, am = med(b), sum(a) / len(a)
        d = (am - bm) * richtung
        if extra is not None and not extra(am, bm):
            schwere = "gelb" if d >= grenze_g else None
        else:
            schwere = "orange" if d >= grenze_o else "gelb" if d >= grenze_g else None
        if schwere:
            sig.append({"name": name, "schwere": schwere, "art": "koerper",
                        "text": text.format(m=_de(am, nd), b=_de(bm, nd), d=_de(abs(am - bm), nd), e=einheit)})

    # a) Body Battery beim Aufwachen (orange nur, wenn zusätzlich absolut niedrig)
    pruefe("Body Battery", bb_auf, l3, 20, 18, -1,
           "Body Battery beim Aufwachen 3 Tage Ø {m} (dein Schnitt {b})", extra=lambda am, bm: am < 35)
    # b) Schlafstress
    pruefe("Schlafstress", lambda d: _g_get(gt.get(d) or {}, "schlaf", "stress"), l7, 8, 4, 1,
           "Schlafstress 7 Tage Ø {m} – {d} über deinem Schnitt")
    # c) Atmung im Schlaf
    pruefe("Atmung", lambda d: _g_get(gt.get(d) or {}, "schlaf", "atem"), l3, 1.5, 1.0, 1,
           "Atmung im Schlaf 3 Tage Ø {m}/min – {d} über deinem Schnitt (frühes Krankheitszeichen möglich)", nd=1)
    # d) SpO2 im Schlaf
    pruefe("SpO2", lambda d: _g_get(gt.get(d) or {}, "schlaf", "spo2"), l3, 3, 2, -1,
           "SpO2 im Schlaf 3 Tage Ø {m} % – {d} unter deinem Schnitt")
    # e) Stress tagsüber (heute unvollständig -> Vorwoche)
    pruefe("Stress", lambda d: _g_get(gt.get(d) or {}, "tag", "stress_avg"), vorwoche, 8, 5, 1,
           "Stress tagsüber 7 Tage Ø {m} – {d} über deinem Schnitt")

    # f) Grauzone nach Zonenzeit (Last-Signal)
    z = [0.0] * 5
    std = km = hm = 0.0
    for d in vorwoche:
        sm = (gt.get(d) or {}).get("summe") or {}
        for i, v in enumerate((sm.get("zonen_s") or [])[:5]):
            z[i] += v or 0
        std += (sm.get("bewegt_s") or sm.get("dauer_s") or 0) / 3600
        km += sm.get("km") or 0
        hm += sm.get("hm") or 0
    zsum = sum(z)
    if zsum >= 3 * 3600:
        z3, hart = z[2] / zsum, (z[3] + z[4]) / zsum
        if z3 >= 0.40 and hart < 0.10:  # am Verlauf geprüft: ab 30 % an 35 % aller Tage -> zu laut
            sig.append({"name": "Grauzone", "schwere": "gelb", "art": "last",
                        "text": f"Grauzone: {_de(z3 * 100)} % der Trainingszeit der letzten 7 Tage in Z3, nur {_de(hart * 100)} % hart"})
    if zsum > 0:
        lage["zonen7_pct"] = [round(v / zsum * 100) for v in z]
    if std > 0:
        lage["std7"], lage["km7"], lage["hm7"] = round(std, 1), round(km), round(hm)

    # Lage für den Prompt (nur vorhandene Werte)
    g0 = gt.get(heute) or {}
    gest = gt.get(_g_tage(heute, 1, 1)[0]) or {}
    if bb_auf(heute) is not None:
        lage["bb_aufwachen"] = bb_auf(heute)
    b = werte(basis, bb_auf)
    if len(b) >= 10:
        lage["bb_schnitt"] = med(b)
    for k, src in (("schlaf_tief_min", "tief_s"), ("schlaf_rem_min", "rem_s"), ("schlaf_wach_min", "wach_s")):
        v = _g_get(g0, "schlaf", src)
        if v is not None:
            lage[k] = round(v / 60)
    dauer, bedarf = _g_get(g0, "schlaf", "dauer_s"), _g_get(g0, "schlaf", "bedarf_min")
    if dauer is not None and bedarf is not None:
        lage["schlafbedarf_diff_min"] = round(dauer / 60 - bedarf)
    for k, pfad in (("schlafstress", ("schlaf", "stress")), ("atem", ("schlaf", "atem")), ("spo2", ("schlaf", "spo2"))):
        v = _g_get(g0, *pfad)
        if v is not None:
            lage[k] = v
    for k, pfad in (("atem_schnitt", ("schlaf", "atem")), ("spo2_schnitt", ("schlaf", "spo2"))):
        xs = werte(basis, lambda d, p=pfad: _g_get(gt.get(d) or {}, *p))
        if len(xs) >= 10:
            lage[k] = med(xs)
    v = _g_get(gest, "tag", "stress_avg")
    if v is not None:
        lage["stress_gestern"] = v
    return sig, lage


def _garmin_lage_txt(l: dict) -> str:
    t = []
    if "bb_aufwachen" in l:
        t.append(f"Body Battery beim Aufwachen {_de(l['bb_aufwachen'])}" + (f" (Schnitt {_de(l['bb_schnitt'])})" if "bb_schnitt" in l else ""))
    sl = [f"{n} {l[k]}" for n, k in (("tief", "schlaf_tief_min"), ("REM", "schlaf_rem_min"), ("wach", "schlaf_wach_min")) if k in l]
    if sl:
        s = "Schlaf " + " / ".join(sl) + " min"
        if "schlafbedarf_diff_min" in l:
            dm = l["schlafbedarf_diff_min"]
            s += f", Bedarf {'+' if dm >= 0 else '−'}{abs(dm)} min"
        t.append(s)
    if "schlafstress" in l:
        t.append(f"Schlafstress {_de(l['schlafstress'])}")
    if "atem" in l:
        t.append(f"Atmung {_de(l['atem'], 1)}/min" + (f" (Schnitt {_de(l['atem_schnitt'], 1)})" if "atem_schnitt" in l else ""))
    if "spo2" in l:
        t.append(f"SpO2 {_de(l['spo2'])} %" + (f" (Schnitt {_de(l['spo2_schnitt'])})" if "spo2_schnitt" in l else ""))
    if "stress_gestern" in l:
        t.append(f"Stress gestern {_de(l['stress_gestern'])}")
    if "std7" in l:
        s = f"letzte 7 Tage {_de(l['std7'], 1)} h, {_de(l['km7'])} km, {_de(l['hm7'])} hm"
        if "zonen7_pct" in l:
            s += ", Zonen " + " · ".join(f"Z{i + 1} {p} %" for i, p in enumerate(l["zonen7_pct"]))
        t.append(s)
    return "; ".join(t)


def vorausblick_rechnen() -> dict:
    """Alles regelbasiert/statistisch – das Sprachmodell formuliert danach nur den Text."""
    heute = today()
    metr = ["load_fb_day", "Garmin_Training_Readiness", "hrv_status", "hrv_threshholds", "fbACWR_obs", "Monotony7",
            "rhr_bpm", "sleep_hours", "sleep_score_0_100", "garminEnduranceScore"]
    rows = sorted(fetch_rows(200, metr), key=lambda r: r["date"])
    vor = [r for r in rows if r["date"] < heute]
    h = next((r for r in rows if r["date"] == heute), {})
    sim = http_get_json(SCRIPT_URL + "?mode=planSimulationJson")
    base = sim.get("base") or {}
    plan = [d for d in (sim.get("days") or []) if d.get("date", "") >= heute][:14]
    if not plan:
        raise RuntimeError("planSimulationJson ohne Plantage")

    # --- Simulation ATL/CTL/ACWR/Monotonie (wie Plan-Cockpit) + Verlängerung auf 28 Tage
    hist_loads = [num(r.get("load_fb_day")) or 0.0 for r in vor][-60:]
    loads = [float(d.get("load") or 0) for d in plan]
    saison_wo = None
    try:
        sx = http_get_json(SCRIPT_URL + "?mode=extras&typ=saison&n=1").get("items") or []
        saison_wo = (sx[0].get("data") or {}).get("wochen") if sx else None
    except Exception:
        pass
    datum0 = datetime.date.fromisoformat(plan[0]["date"])

    def _saison_woche(d: str):
        for w in saison_wo or []:
            st = w.get("start", "")
            if st and st <= d < (datetime.date.fromisoformat(st) + datetime.timedelta(days=7)).isoformat():
                return w
        return None

    # Horizont: 28 Tage – liegt Tag 28 in einer Entlastungswoche, bis zum Ende der Folgewoche weiterrechnen
    horizont = 28
    # Entlastungswoche in den Tagen 22–28? Dann bis zum Ende der Folgewoche (Wirkung der Entlastung sichtbar machen)
    for k in range(22, 29):
        wk = _saison_woche((datum0 + datetime.timedelta(days=k)).isoformat())
        if wk and any(x in str(wk.get("phase", "")).lower() for x in ("entlast", "pause")):
            ende = datetime.date.fromisoformat(wk["start"]) + datetime.timedelta(days=13)
            horizont = max(horizont, min(42, (ende - datum0).days))
            break
    ext_dates = [(datum0 + datetime.timedelta(days=i)).isoformat() for i in range(horizont + 1)]
    ext_loads = loads[:]
    fallback = (sum(loads[1:]) / max(1, len(loads) - 1)) if len(loads) > 1 else 120.0
    for i in range(len(loads), horizont + 1):
        d = ext_dates[i]
        wl = None
        for w in saison_wo or []:
            s = w.get("start", "")
            if s and s <= d < (datetime.date.fromisoformat(s) + datetime.timedelta(days=7)).isoformat():
                wl = float(w.get("ziel_ess") or 0) / 7
        # Pause-Wochen (Zeitfenster „kein Training möglich“) dürfen 0 sein – nur fehlende Wochen fallen auf den Ø zurück
        ext_loads.append(wl if wl is not None else fallback)
    atl, ctl = float(base.get("atl") or 0), float(base.get("ctl") or 0)
    serie = hist_loads[:]
    tage = []
    for i, l in enumerate(ext_loads):
        serie.append(l)
        if i > 0:
            atl = _atl_step(atl, l)
            ctl = _ctl_step(ctl, l, sum(serie[-28:]) / len(serie[-28:]))
        acwr = atl / ctl if ctl > 0 else 0.0
        mono = mono7(serie)
        tage.append({"datum": ext_dates[i], "last": round(l), "atl": round(atl), "ctl": round(ctl), "acwr": round(acwr, 2),
                     "mono": round(min(mono, 9.99), 2)})
    # Endurance Score (Näherung wie lib/ctlModel.ts esStep) – Startwert = letzter gemessener Wert vor heute
    es0 = next((num(r.get("garminEnduranceScore")) for r in reversed(vor) if num(r.get("garminEnduranceScore"))), None)
    if es0:
        es = float(es0)
        se = hist_loads[:]
        for i, l in enumerate(ext_loads):
            se.append(l)
            es = _es_step(es, l, sum(se[-56:]) / len(se[-56:]))
            tage[i]["es"] = round(es)
    for i, d in enumerate(plan):
        tage[i].update({"sport": d.get("sport") or "Off", "zone": d.get("zone") or "Off", "fest": bool(d.get("locked")), "tag": _dm(d["date"])})

    # --- Bereitschafts-Prognose (Morgenwert)
    hm = [{"date": r["date"], "load": num(r.get("load_fb_day")), "readiness": num(r.get("Garmin_Training_Readiness")),
           "acwr": num(r.get("fbACWR_obs")), "mono": num(r.get("Monotony7"))} for r in vor]
    # Krankheitstage (Ereignis-Marker) und den Folgetag nicht als Lernbeispiel verwenden
    try:
        krank = set()
        for e in http_get_json(SCRIPT_URL + "?mode=extras&typ=ereignis&n=400").get("items") or []:
            if "krank" in ((e.get("data") or {}).get("arten") or []):
                d0 = datetime.date.fromisoformat(str(e.get("datum"))[:10])
                krank.update({d0.isoformat(), (d0 + datetime.timedelta(days=1)).isoformat()})
        for x in hm:
            if x["date"] in krank:
                x["readiness"] = None
    except Exception:
        pass
    modell = _bereitschaft_modell(hm)
    bereit = {}
    r_heute = num(h.get("Garmin_Training_Readiness"))
    if r_heute is not None:
        bereit[heute] = round(r_heute)
    if modell:
        r = r_heute if r_heute is not None else next((x["readiness"] for x in reversed(hm) if x["readiness"] is not None), 60.0)
        ls = hist_loads[:]
        for i in range(min(len(plan), 13)):
            ls.append(tage[i]["last"])
            prev = ls[-2] if len(ls) > 1 else 0.0
            m7 = sum(ls[-7:]) / len(ls[-7:])
            rn = max(1.0, min(100.0, _pred(modell, _feats(tage[i]["last"], prev, m7, tage[i]["acwr"], tage[i]["mono"], r))))
            bereit[tage[i + 1]["datum"]] = round(rn)
            r = rn
    for t in tage:
        if t["datum"] in bereit:
            t["bereitschaft"] = bereit[t["datum"]]

    # --- Monotonie-Hinweise im Plan
    mono_w = mono_verstoesse(hist_loads[-28:], [{"datum": t["datum"], "last": t["last"], "fest": t.get("fest")} for t in tage[:len(plan)]])
    mono_daten = {w[:10] for w in mono_w}
    ref = ref_load(hist_loads[-28:])

    # --- Kalender + Wetter
    kal = _kalender(11)
    wet = _wetter_tage(10)

    def kal_score(k: dict | None, d: str) -> tuple[float, str]:
        if not k:
            return 0.7, "Kalender unbekannt"
        if k.get("reise"):
            return 0.1, "Reise"
        bh = float(k.get("belegt_h") or 0)
        wochenende = datetime.date.fromisoformat(d).weekday() >= 5
        if k.get("frei") or bh <= 1:
            s, t = 1.0, "frei"
        elif bh <= 3:
            s, t = 0.8, f"{_de(bh)} h belegt"
        elif bh <= 5:
            s, t = 0.55, f"{_de(bh)} h belegt"
        else:
            s, t = 0.25, f"{_de(bh)} h belegt"
        if wochenende and bh <= 3:
            s = min(1.0, s + 0.1)
        return s, t

    def wetter_score(w: dict | None) -> tuple[float, str]:
        if not w:
            return 0.6, "Wetter unbekannt"
        rg, wd, tm = float(w.get("regen") or 0), float(w.get("wind") or 0), float(w.get("tmax") or 0)
        s = (1.0 if rg <= 1 else 0.6 if rg <= 4 else 0.2) * (1.0 if wd <= 25 else 0.7 if wd <= 40 else 0.3) * (1.0 if 5 <= tm <= 26 else 0.75)
        return s, f"{_de(tm)}°, {_de(rg, 1)} mm, Wind {_de(wd)} km/h"

    # --- Bester Qualitätstag (nächste 7 Tage)
    qual = None
    for i in range(1, min(8, len(plan))):
        t = tage[i]
        b = t.get("bereitschaft")
        if b is None or b < 55 or t["acwr"] >= 1.3 or t["datum"] in mono_daten or (t.get("fest") and t["last"] <= 0):
            continue
        ks, kt = kal_score(kal.get(t["datum"]), t["datum"])
        if ks < 0.5:
            continue
        vortag = tage[i - 1]["last"]
        score = b + (5 if vortag <= max(60.0, 0.7 * ref) else -5) + 10 * ks
        if not qual or score > qual["score"]:
            qual = {"datum": t["datum"], "tag": t["tag"], "score": round(score), "bereitschaft": b,
                    "grund": f"Bereitschaft ~{b}, Vortag {'locker' if vortag <= max(60.0, 0.7 * ref) else 'fordernd'} ({vortag} ESS), {kt}, ACWR {_de(t['acwr'], 2)}"}

    # --- Tourfenster (nächste 9 Tage)
    fenster = []
    for i in range(1, min(10, len(tage))):
        t = tage[i]
        d = t["datum"]
        ws, wt = wetter_score(wet.get(d))
        ks, kt = kal_score(kal.get(d), d)
        b = t.get("bereitschaft")
        bs = 0.7 if b is None else (1.0 if b >= 65 else 0.75 if b >= 50 else 0.4)
        ps = 1.0
        if t.get("fest") and t["last"] <= 0:
            ps = 0.4
        if d in mono_daten:
            ps = min(ps, 0.6)
        if tage[i - 1]["last"] > 1.3 * max(ref, 1):
            ps = min(ps, 0.8)
        score = round(100 * (0.35 * ws + 0.3 * ks + 0.25 * bs + 0.1 * ps))
        if ks <= 0.1:
            score = min(score, 40)  # Reisetag
        elif ks < 0.5:
            score = min(score, 50)  # zu voller Tag für eine lange Tour
        fenster.append({"datum": d, "tag": _dm(d), "score": score, "wetter": wt, "kalender": kt,
                        "bereitschaft": b, "plan": f"{t['last']} ESS {t.get('sport', '')} {t.get('zone', '')}".strip()})
    fenster.sort(key=lambda x: (-x["score"], x["datum"]))
    top = [x for x in fenster if x["score"] >= 65][:3]

    # --- Frühwarnung (Trends der letzten 2–5 Wochen)
    def col(k, rs):
        return [num(r.get(k)) for r in rs]
    r_all = vor + ([h] if h else [])
    letzte7, basis = r_all[-7:], r_all[-35:-7]
    sig = []

    def add(name, schwere, text, art):
        sig.append({"name": name, "schwere": schwere, "text": text, "art": art})

    def _band(v):
        """Pars hrv_threshholds: 'lo;hi' oder 'lo,hi'. Gibt (lo, hi) oder None."""
        if not v or str(v).strip() == "":
            return None
        try:
            s = str(v).strip()
            parts = [p.strip().replace(",", ".") for p in s.split(";")] if ";" in s else [p.strip() for p in s.split(",")]
            if len(parts) != 2:
                return None
            lo, hi = float(parts[0]), float(parts[1])
            if lo >= hi:
                return None
            return (lo, hi)
        except Exception:
            return None

    # Für letzte 14 Tage: Liste von (hrv, lo, hi)
    tage_14 = r_all[-14:]  # letzte 14 Tage chronologisch
    hrv_bands = []
    for r in tage_14:
        hrv_val = num(r.get("hrv_status"))
        band = _band(r.get("hrv_threshholds"))
        if hrv_val is None or band is None:
            continue
        hrv_bands.append((hrv_val, band[0], band[1]))

    # unter_folge: aufeinanderfolgende Tage am Ende (rückwärts), an denen hrv < lo
    unter_folge = 0
    if hrv_bands:
        for entry in reversed(hrv_bands):
            h, lo, hi = entry
            if h < lo:
                unter_folge += 1
            else:
                break

    # unter_7: Anzahl Tage unter lo unter den letzten 7 Einträgen
    letzteren7 = hrv_bands[-7:] if len(hrv_bands) >= 7 else hrv_bands
    unter_7 = sum(1 for entry in letzteren7 if entry[0] < entry[1])

    # Signale
    if hrv_bands:  # nicht-leere Liste
        neu = hrv_bands[-1]
        # Schwellen am eigenen Verlauf geprüft (400 Tage): orange 14 Tage, gelb 77 Tage; Krankheit 12/2025 neun Tage vorher orange
        if unter_folge >= 7 and neu[0] < neu[1] - 3:
            add("HRV", "orange", f"HRV seit {unter_folge} Tagen unter deinem Garmin-Normalband ({_de(neu[1])}–{_de(neu[2])} ms), heute {_de(neu[0])} ms", "koerper")
        elif unter_folge >= 4 or unter_7 >= 6:
            add("HRV", "gelb", f"HRV an {max(unter_folge, unter_7)} der letzten 7 Tage unter deinem Garmin-Normalband ({_de(neu[1])}–{_de(neu[2])} ms)", "koerper")
        elif len(hrv_bands) >= 10:
            sl = _slope_per_week([e[1] for e in hrv_bands])
            if sl is not None and sl <= -1.0:
                add("HRV-Basis", "gelb", f"Deine HRV-Basis (untere Bandgrenze) sinkt um {_de(-sl, 1)} ms pro Woche", "koerper")
    rh7, rhb = _mean(col("rhr_bpm", letzte7)), _mean(col("rhr_bpm", basis))
    if rh7 and rhb:
        dlt = rh7 - rhb
        if dlt >= 4:
            add("Ruhepuls", "orange", f"Ruhepuls 7 Tage Ø {_de(rh7)} – {_de(dlt, 1)} Schläge über deinem Schnitt", "koerper")
        elif dlt >= 2:
            add("Ruhepuls", "gelb", f"Ruhepuls 7 Tage Ø {_de(rh7)} – {_de(dlt, 1)} Schläge über deinem Schnitt", "koerper")
    sh = [x for x in col("sleep_hours", letzte7) if x is not None]
    if len(sh) >= 5:
        schuld = sum(max(0.0, 7.5 - x) for x in sh)
        if schuld >= 7:
            add("Schlafschuld", "orange", f"Schlafschuld {_de(schuld, 1)} h in 7 Tagen (Ziel 7,5 h/Nacht)", "koerper")
        elif schuld >= 4:
            add("Schlafschuld", "gelb", f"Schlafschuld {_de(schuld, 1)} h in 7 Tagen (Ziel 7,5 h/Nacht)", "koerper")
    ss7 = _mean(col("sleep_score_0_100", letzte7))
    if ss7 is not None and ss7 < 70:
        add("Schlafqualität", "gelb", f"Schlafscore 7 Tage Ø {_de(ss7)}", "koerper")
    rd7, rdb = _mean(col("Garmin_Training_Readiness", letzte7)), _mean(col("Garmin_Training_Readiness", basis))
    if rd7 is not None and rdb is not None:
        dlt = rd7 - rdb
        if dlt <= -18:
            add("Bereitschaft", "orange", f"Trainingsbereitschaft 7 Tage Ø {_de(rd7)}, {_de(-dlt)} Punkte unter Schnitt", "koerper")
        elif dlt <= -10:
            add("Bereitschaft", "gelb", f"Trainingsbereitschaft 7 Tage Ø {_de(rd7)}, {_de(-dlt)} Punkte unter Schnitt", "koerper")
    a0 = tage[0]["acwr"]
    if a0 > 1.5:
        add("ACWR", "orange", f"ACWR {_de(a0, 2)} – akute Last deutlich über der Grundlage", "last")
    elif a0 > 1.3:
        add("ACWR", "gelb", f"ACWR {_de(a0, 2)} – akute Last über der Grundlage", "last")
    amax = max(tage[1:len(plan)], key=lambda t: t["acwr"]) if len(plan) > 1 else None
    if amax and amax["acwr"] > 1.35 and a0 <= 1.3:
        add("ACWR geplant", "gelb", f"Plan führt zu ACWR {_de(amax['acwr'], 2)} am {_dm(amax['datum'])}", "last")
    # Monotonie-Verlauf (tatsächliche Tage)
    mono_seq = [mono7(hist_loads[:k]) for k in range(max(7, len(hist_loads) - 6), len(hist_loads) + 1)]
    ueber2 = 0
    for m in reversed(mono_seq):
        if m > 2.0:
            ueber2 += 1
        else:
            break
    if mono_seq and mono_seq[-1] > 2.5 and ueber2 >= 2:
        add("Monotonie", "orange", f"Monotonie {_de(mono_seq[-1], 2)}, seit {ueber2} Tagen über 2,0", "last")
    elif ueber2 >= 3:
        add("Monotonie", "gelb", f"Monotonie seit {ueber2} Tagen über 2,0", "last")
    # Geplante Wochenlast (nächste 7 Tage) vs. Schnitt der letzten 4 Wochen – Monotonie zählt separat
    if len(hist_loads) >= 28 and len(plan) >= 8:
        wo_schnitt = sum(hist_loads[-28:]) / 4
        wo_plan = sum(t["last"] for t in tage[1:8])
        if wo_schnitt > 0:
            plus = (wo_plan / wo_schnitt - 1) * 100
            if plus >= 45:
                add("Wochenlast geplant", "orange", f"Nächste 7 Tage {_de(wo_plan)} ESS – {_de(plus)} % über deinem 4-Wochen-Schnitt", "last")
            elif plus >= 25:
                add("Wochenlast geplant", "gelb", f"Nächste 7 Tage {_de(wo_plan)} ESS – {_de(plus)} % über deinem 4-Wochen-Schnitt", "last")
    if mono_w:
        add("Monotonie-Regel", "gelb", f"{len(mono_w)} Hinweis(e) im Plan, zuerst {_mono_txt(mono_w[0])}", "last")
    g_sig, g_lage = garmin_signale(garmin_tage(60), heute)
    for s in g_sig:
        add(s["name"], s["schwere"], s["text"], s["art"])
        sig[-1]["quelle"] = "garmin"
    kor = [s for s in sig if s["art"] == "koerper"]
    las = [s for s in sig if s["art"] == "last"]
    orange_k = [s for s in kor if s["schwere"] == "orange"]
    orange_l = [s for s in las if s["schwere"] == "orange"]
    if len(orange_k) >= 2 or (orange_k and orange_l):
        stufe = "rot"
    elif orange_k or orange_l or len([s for s in sig if not (s.get("quelle") == "garmin" and s["schwere"] == "gelb")]) >= 3:
        stufe = "orange"
    elif sig:
        stufe = "gelb"
    else:
        stufe = "gruen"
    if kor and las:
        prog = "Körperwerte und Last zeigen in dieselbe Richtung: Bleibt der Plan so, droht in 5–7 Tagen Überlastung – Entlastungstag früh einplanen."
    elif kor:
        prog = "Der Körper zeigt erste Ermüdung, die Last ist moderat – Qualität nur bei guten Morgenwerten, sonst locker."
    elif las:
        prog = "Der Plan braucht mehr Entlastung, die Körperwerte sind noch stabil – Entlastungstag einplanen."
    else:
        prog = "Keine Frühwarnzeichen in den letzten Wochen."

    i7, i14, i28 = min(7, len(tage) - 1), min(14, len(tage) - 1), min(28, len(tage) - 1)
    c0, c28 = tage[0]["ctl"], tage[i28]["ctl"]
    c_ende = tage[-1]["ctl"]
    trend = "steigt" if c_ende > c0 * 1.03 else "sinkt" if c_ende < c0 * 0.97 else "stabil"

    # --- Szenarien ab dem Tag nach dem Plan: Saisonrahmen / Fitness halten / wie die letzten 4 Wochen
    def _lauf(fuell: float | None) -> float:
        a_, c_ = float(base.get("atl") or 0), float(base.get("ctl") or 0)
        se = hist_loads[:]
        for i, l in enumerate(ext_loads):
            l = l if (fuell is None or i < len(loads)) else fuell
            se.append(l)
            if i > 0:
                a_ = _atl_step(a_, l)
                c_ = _ctl_step(c_, l, sum(se[-28:]) / len(se[-28:]))
        return round(c_)

    def _halte_last(ctl0: float) -> float:
        """Tageslast, bei der die CTL im Gleichgewicht bleibt (numerisch, 90 Tage konstante Last)."""
        lo, hi = 0.0, 400.0
        for _ in range(30):
            mid = (lo + hi) / 2
            c_ = ctl0
            for _k in range(90):
                c_ = _ctl_step(c_, mid, mid)
            lo, hi = (mid, hi) if c_ < ctl0 else (lo, mid)
        return (lo + hi) / 2

    halte = _halte_last(float(c0))
    vier = sum(hist_loads[-28:]) / max(1, len(hist_loads[-28:]))
    # Füll-Last nach dem Plan, mit der die CTL am Horizont-Ende wieder auf dem heutigen Wert steht
    lo, hi = 0.0, 500.0
    for _ in range(30):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if _lauf(mid) < c0 else (lo, mid)
    fuell_halten = (lo + hi) / 2
    szen = {
        "plan_saison": {"ctl": c_ende, "tag": round(sum(ext_loads[len(loads):]) / max(1, len(ext_loads) - len(loads)))},
        "halten": {"ctl": _lauf(fuell_halten), "tag": round(fuell_halten)},
        "vier_wochen": {"ctl": _lauf(vier), "tag": round(vier)},
    }

    # --- Wochen (Mo–So) mit Quelle und Abstand zur Haltelast
    wochen = []
    i = 1
    while i < len(ext_dates):
        d0 = datetime.date.fromisoformat(ext_dates[i])
        n_bis_so = 6 - d0.weekday()
        j = min(len(ext_dates) - 1, i + n_bis_so)
        idx = list(range(i, j + 1))
        summe = sum(ext_loads[k] for k in idx)
        im_plan = [k < len(loads) for k in idx]
        if all(im_plan):
            quelle = "Plan"
        elif not any(im_plan):
            w = _saison_woche(ext_dates[i])
            quelle = f"Saison KW {w.get('kw')} ({w.get('phase')})" if w else "Ø Plan"
        else:
            quelle = "Plan + Saison"
        soll = halte * len(idx)
        wochen.append({"von": ext_dates[i], "bis": ext_dates[j], "tage": len(idx), "quelle": quelle,
                       "last": round(summe), "halten": round(soll), "diff": round(summe - soll),
                       "ctl_ende": tage[j]["ctl"], "unter": bool(summe < 0.9 * soll and all(im_plan))})
        i = j + 1
    return {
        "erstellt": datetime.datetime.now(tz_local).isoformat(timespec="minutes"),
        "vorausblick": {
            "ctl_heute": c0, "ctl_7": tage[i7]["ctl"], "ctl_14": tage[i14]["ctl"], "ctl_28": c28, "trend": trend,
            "es_heute": round(es0) if es0 else None, "es_7": tage[i7].get("es"), "es_14": tage[i14].get("es"), "es_28": tage[i28].get("es"),
            "horizont": len(tage) - 1, "ctl_ende": c_ende, "datum_ende": tage[-1]["datum"],
            "halte_tag": round(halte), "halte_woche": round(halte * 7), "szenarien": szen, "wochen": wochen,
            "basis_28": "Saisonrahmen" if saison_wo else "Ø Plan",
            "acwr_max": {"wert": amax["acwr"], "datum": amax["datum"]} if amax else None,
            "mono_hinweise": mono_w, "qualitaetstag": qual, "tage": tage[:len(plan)],
            "modell": {"n": modell["n"], "mae": round(modell["mae"], 1), "naiv": round(modell["naiv"], 1)} if modell else None,
        },
        "tourfenster": top,
        "tourfenster_alle": fenster,
        "fruehwarnung": {"stufe": stufe, "signale": sig, "prognose": prog},
        "garmin_lage": g_lage,
    }


def _mono_txt(x: str) -> str:
    """'2026-10-28: …' → 'Mi 28.10.: …'"""
    try:
        return _dm(x[:10]) + x[10:]
    except Exception:
        return x


def _vorausblick_block(d: dict) -> str:
    v, fw = d["vorausblick"], d["fruehwarnung"]
    z = [f"**VORAUSBLICK (vom Mac berechnet, {d['erstellt'][11:16]} Uhr):**",
         f"- Fitness (CTL): heute {v['ctl_heute']} → 7 Tage {v['ctl_7']} → 14 Tage {v['ctl_14']} → 28 Tage {v['ctl_28']}"
         + (f" → {v['horizont']} Tage {v['ctl_ende']}" if v.get('horizont', 28) > 28 else "")
         + f" ({v['trend']}; nach dem Plan {v['basis_28']})."]
    if v.get("szenarien"):
        sz = v["szenarien"]
        z.append(f"- Nach dem Plan (bis {_dm(v['datum_ende'])}): laut Saisonrahmen CTL {sz['plan_saison']['ctl']} · Fitness halten (zurück auf {v['ctl_heute']}) bräuchte {sz['halten']['tag']} ESS/Tag · "
                 f"wie letzte 4 Wochen ({sz['vier_wochen']['tag']} ESS/Tag) {sz['vier_wochen']['ctl']}.")
    if v.get("wochen"):
        wt = "; ".join(f"{_dm(w['von'])}–{_dm(w['bis'])} {w['quelle']} {w['last']} ESS ({'+' if w['diff'] >= 0 else '−'}{abs(w['diff'])} zu Fitness halten)" for w in v["wochen"])
        z.append(f"- Wochen (Fitness halten auf Dauer = {v['halte_woche']} ESS/Woche): {wt}.")
        unter = [w for w in v["wochen"] if w.get("unter")]
        if unter:
            z.append("- Achtung: " + "; ".join(f"Plan-Woche {_dm(w['von'])}–{_dm(w['bis'])} liegt {abs(w['diff'])} ESS unter Fitness halten" for w in unter) + ".")
    if v.get("acwr_max"):
        z.append(f"- Höchster ACWR im Plan: {_de(v['acwr_max']['wert'], 2)} am {_dm(v['acwr_max']['datum'])}")
    br = [f"{t['tag']} {t['bereitschaft']}" for t in v["tage"][1:8] if t.get("bereitschaft") is not None]
    if br:
        z.append("- Bereitschafts-Prognose (Morgenwert): " + ", ".join(br) + (f" (Fehler ±{_de(v['modell']['mae'])})" if v.get("modell") else "") + ".")
    if v.get("qualitaetstag"):
        q = v["qualitaetstag"]
        z.append(f"- Bester Tag für eine Qualitätseinheit: {q['tag']} – {q['grund']}.")
    if d["tourfenster"]:
        z.append("- Tourfenster: " + "; ".join(f"{x['tag']} (Eignung {x['score']}/100: {x['wetter']}, {x['kalender']})" for x in d["tourfenster"]) + ".")
    else:
        z.append("- Tourfenster: in den nächsten 9 Tagen kein gutes Fenster (Wetter/Kalender/Bereitschaft).")
    if v["mono_hinweise"]:
        z.append("- Monotonie-Regel im Plan: " + "; ".join(_mono_txt(x) for x in v["mono_hinweise"][:3]) + ".")
    z.append(f"- Frühwarnung: {fw['stufe'].upper()}" + (" – " + "; ".join(s["text"] for s in fw["signale"][:4]) if fw["signale"] else "") + f". {fw['prognose']}")
    if d.get("garmin_lage"):
        gl = _garmin_lage_txt(d["garmin_lage"])
        if gl:
            z.append(f"- Garmin-Lage: {gl}.")
    z.append("Nutze diese Zukunftssicht: Nenne kurz die nächste Chance (Qualitätstag oder Tourfenster) und reagiere auf die Frühwarnung, ohne die Monotonie- und ACWR-Regeln zu verletzen.")
    return "\n".join(z)


def cmd_vorausblick(force: bool = False) -> None:
    """Morgens (sobald die Morgenwerte da sind) und ab 17 Uhr ein zweites Mal."""
    st = load_x()
    jetzt = datetime.datetime.now(tz_local)
    slot = "abend" if jetzt.hour >= 17 else "morgen"
    key = f"{today()}:{slot}"
    if not force and (st.get("vorausblick_done") == key or not (5 <= jetzt.hour < 23)):
        return
    if slot == "morgen" and not force:
        r0 = {r["date"]: r for r in fetch_rows(3, ["Garmin_Training_Readiness"])}.get(today()) or {}
        if num(r0.get("Garmin_Training_Readiness")) is None:
            return  # Morgenwerte noch nicht da
    t0 = time.time()
    d = vorausblick_rechnen()
    block = _vorausblick_block(d)
    text = ""
    if ollama_ready():
        try:
            msgs = [
                {"role": "system", "content": "Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, keine Emojis, keine Überschriften."},
                {"role": "user", "content": block + "\n\nFasse das in 3–5 Sätzen für Marc zusammen: zuerst die Frühwarnung (nur wenn nicht grün), dann die nächste Chance (Qualitätstag, Tourfenster) und der Fitnesstrend – sag dabei, ob ein Rückgang aus dem Plan (Woche unter Fitness halten) oder aus einer geplanten Entlastungswoche kommt. Konkrete Tage nennen, keine neuen Zahlen erfinden."},
            ]
            text, _ = ollama_chat(MODEL, msgs, num_ctx=8192, timeout=600)
            text = text.strip().strip('"')
        except Exception as e:
            log(f"vorausblick: Text lokal nicht erzeugt ({e}) – nur Regeltext")
    if not text:
        fw = d["fruehwarnung"]
        teile = []
        if fw["stufe"] != "gruen":
            teile.append(f"Frühwarnung {fw['stufe']}: {fw['prognose']}")
        if d["vorausblick"].get("qualitaetstag"):
            teile.append(f"Bester Qualitätstag: {d['vorausblick']['qualitaetstag']['tag']}.")
        if d["tourfenster"]:
            teile.append(f"Tourfenster: {d['tourfenster'][0]['tag']}.")
        _v = d['vorausblick']
        teile.append(f"Fitness {_v['trend']} ({_v['ctl_heute']} → {_v.get('ctl_ende', _v['ctl_28'])} in {_v.get('horizont', 28)} Tagen; Fitness halten wären {_v.get('halte_woche', '?')} ESS/Woche).")
        text = " ".join(teile)
    d["text"] = text
    d["sekunden"] = round(time.time() - t0)
    save_extra("prognose", today(), text, d, "mac:" + MODEL)
    st = load_x()
    st["vorausblick_done"] = key
    st["vorausblick_datum"] = today()
    st["vorausblick_block"] = block
    save_x(st)
    fw = d["fruehwarnung"]
    log(f"vorausblick: {slot} gespeichert – Frühwarnung {fw['stufe']}, {len(d['tourfenster'])} Tourfenster, Qualitätstag {(d['vorausblick'].get('qualitaetstag') or {}).get('tag', '–')} ({d['sekunden']} s)")
    if fw["stufe"] in ("orange", "rot") and st.get("fw_gemeldet") != today():
        st["fw_gemeldet"] = today()
        save_x(st)
        notify("Coach Kira", f"Frühwarnung {fw['stufe']}: {fw['signale'][0]['text'] if fw['signale'] else fw['prognose']}")


def _main() -> None:
    """CLI mit Unterbefehlen für alle Extras."""
    ap = argparse.ArgumentParser(description="Coach Kira Extras")
    sub_parsers = ap.add_subparsers(dest="command")

    sub_parsers.add_parser("plausi", help="Plausibilitäts-Check")
    st_p = sub_parsers.add_parser("selbsttest", help="Selbsttest lokales Modell")
    st_p.add_argument("--force", action="store_true", help="Tag-Check ignorieren")
    sub_parsers.add_parser("tick", help="Plausi immer + Selbsttest 2-5 Uhr")
    sub_parsers.add_parser("snapshot", help="Empfehlungs-Snapshot speichern")

    tt_p = sub_parsers.add_parser("treffer", help="Treffsicherheit-Tag verarbeiten")
    tt_p.add_argument("--backfill", metavar="DATEI", help="Backfill-JSON laden")

    tw_p = sub_parsers.add_parser("treffer-woche", help="Wochenbilanz erstellen")
    tw_p.add_argument("--force", action="store_true", help="Montag-Check ignorieren")

    ap_p = sub_parsers.add_parser("autopilot", help="Wochen-Autopilot")
    ap_p.add_argument("--force", action="store_true")
    sp_p = sub_parsers.add_parser("stoerung", help="Störungen erkennen")
    sp_p.add_argument("--force", action="store_true")
    vb_p = sub_parsers.add_parser("vorausblick", help="Vorausblick, Tourfenster, Frühwarnung")
    vb_p.add_argument("--force", action="store_true")
    ern_p = sub_parsers.add_parser("ernaehrung", help="Ernährungsplan für morgen")
    ern_p.add_argument("--force", action="store_true", help="Zeitfenster-Check ignorieren")

    args = ap.parse_args()
    if args.command is None:
        ap.print_help()
        return

    if args.command == "plausi":
        try:
            cmd_plausi()
        except Exception as e:
            log(f"extras: FEHLER plausi – {e}")

    elif args.command == "selbsttest":
        try:
            cmd_selbsttest(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER selbsttest – {e}")

    elif args.command == "tick":
        try:
            cmd_plausi()
        except Exception as e:
            log(f"extras: FEHLER tick/plausi – {e}")

        try:
            nh = now_h()
            # nachts 2-5 Uhr; schläft der Mac dann, ab 13 Uhr nachholen
            if 2 <= nh < 5 or nh >= 13:
                cmd_selbsttest(force=False)
        except Exception as e:
            log(f"extras: FEHLER tick/selbsttest – {e}")

        # D-E-F-G tick-trigger
        try:
            nh2 = now_h()
            if 5 <= nh2 < 10:
                cmd_snapshot()
        except Exception as e:
            log(f"extras: FEHLER tick/snapshot – {e}")

        try:
            nh3 = now_h()
            if 6 <= nh3 < 23:
                cmd_treffer()
        except Exception as e:
            log(f"extras: FEHLER tick/treffer – {e}")

        try:
            nh4 = now_h()
            if 6 <= nh4 < 12:
                cmd_treffer_woche()
        except Exception as e:
            log(f"extras: FEHLER tick/treffer-woche – {e}")

        try:
            cmd_ernaehrung()
        except Exception as e:
            log(f"extras: FEHLER tick/ernaehrung – {e}")

        try:
            cmd_stoerung()
        except Exception as e:
            log(f"extras: FEHLER tick/stoerung – {e}")

        try:
            cmd_autopilot()
        except Exception as e:
            log(f"extras: FEHLER tick/autopilot – {e}")

        try:
            cmd_vorausblick()
        except Exception as e:
            log(f"extras: FEHLER tick/vorausblick – {e}")

    elif args.command == "snapshot":
        try:
            cmd_snapshot()
        except Exception as e:
            log(f"extras: FEHLER snapshot – {e}")

    elif args.command == "treffer":
        try:
            cmd_treffer(backfill_file=args.backfill)
        except Exception as e:
            log(f"extras: FEHLER treffer – {e}")

    elif args.command == "treffer-woche":
        try:
            cmd_treffer_woche(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER treffer-woche – {e}")

    elif args.command == "ernaehrung":
        try:
            cmd_ernaehrung(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER ernährung – {e}")

    elif args.command == "autopilot":
        try:
            cmd_autopilot(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER autopilot – {e}")

    elif args.command == "vorausblick":
        try:
            cmd_vorausblick(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER vorausblick – {e}")
    elif args.command == "stoerung":
        try:
            cmd_stoerung(force=args.force)
        except Exception as e:
            log(f"extras: FEHLER stoerung – {e}")


# Alias für backward-compat
main = _main



if __name__ == "__main__":
    main()
