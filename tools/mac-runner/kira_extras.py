"""Erweiterungen für Coach Kira: Plausi-Check, Selbsttest lokales Modell, CLI.

Liest kira_runner.py aus demselben Verzeichnis und nutzt dort log, http_get_json
usw.  NUR Standardbibliothek.
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
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
    """http_post_json(SCRIPT_URL, {'kind':'saveExtra',...}); ok-Feld prüfen."""
    result = http_post_json(
        SCRIPT_URL,
        {
            "kind": "saveExtra",
            "typ": typ,
            "datum": datum,
            "text": text,
            "data": data,
            "quelle": quelle,
        },
    )
    if not result.get("ok"):
        log(f"save_extra({typ},{datum}): FEHLER – {result.get('error', '')}")
        return False
    return True


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
    if hrv_t is not None:
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
        log("selbsttest: bereits heute ausgeführt – übersprungen")
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
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, fmt="json", timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = json.loads(text_out)
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
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, fmt="json", timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = json.loads(text_out)
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
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, fmt="json", timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = json.loads(text_out)
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
    def test_4() -> dict:
        prompt_text = "Heute Nacht 7 h 24 min geschlafen, Schlafscore 81. Ruhepuls 46, HRV 50 ms, Trainingsbereitschaft 77."
        messages = [
            {"role": "system", "content": "Du bist ein exakter Datenassistent. Antworte nur mit JSON."},
            {"role": "user", "content": f"Aus dem Text extrahieren:\n{prompt_text}\nAufgabe: JSON {{\"sleep_hours\",\"sleep_score\",\"rhr\",\"hrv\",\"readiness\"}}"},
        ]
        try:
            t0 = datetime.datetime.now(tz_local).timestamp()
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, fmt="json", timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = json.loads(text_out)
            ok = True
            detail = ""

            sh_val = float(data.get("sleep_hours", -999))
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
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, fmt="json", timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = json.loads(text_out)
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
            text_out, tps = ollama_chat(MODEL, messages, num_ctx=8192, fmt="json", timeout=300)
            sekunden = datetime.datetime.now(tz_local).timestamp() - t0
            data = json.loads(text_out)
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

    # Nur die 60 neuesten Tage behalten
    if rec:
        keys = sorted(rec.keys(), reverse=True)[:60]
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
    rows = fetch_rows(90, METRICS_T)
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
                        "quelle": "plan-briefing",
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
            # max. 90 behalten
            if len(treffer_dict) > 90:
                keys_sorted = sorted(treffer_dict.keys(), reverse=True)[:90]
                st["treffer"] = {k: treffer_dict[k] for k in keys_sorted}

            if datum not in st.get("treffer_done", []):
                treffer_done.append(datum)
            st["treffer_done"] = treffer_done[-200:]
            neu += 1

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
            log("treffer_woche: kein Montag – übersprungen")
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
            dt = datetime.datetime.strptime(d_str, "%Y-%m-%d")
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
            log(f"ernaehrung: Zeitfenster {nh}h nicht erfüllt (17-23) – übersprungen")
            return
        if st.get("ern_done") == heute:
            log(f"ernaehrung: bereits heute ausgeführt – übersprungen")
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
        llm_out, _ = ollama_chat(MODEL, msgs, fmt="json", num_ctx=8192, timeout=600)
        llm_data = json.loads(llm_out)
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


# Alias für backward-compat
main = _main



if __name__ == "__main__":
    main()
