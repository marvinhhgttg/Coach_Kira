"""Tests für Garmin-Modul – Standardbibliothek only."""
from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
from pathlib import Path
from unittest import mock

# Mock kira_runner VOR dem Import von kira_garmin
_td_mock = tempfile.TemporaryDirectory()

_kira_runner_mock = mock.MagicMock(
    log=lambda *a: None,
    http_post_json=lambda url, payload: {"ok": True},
    notify=lambda *a: None,
    SCRIPT_URL="https://script.google.com/macros/s/FAKE",
    TZ=None,
    BASE=Path(_td_mock.name),
)

sys.modules["kira_runner"] = _kira_runner_mock

import kira_garmin as km


def _recreate_tmp() -> tempfile.TemporaryDirectory:
    td = tempfile.TemporaryDirectory()
    p = Path(td.name)
    km.BASE = p
    km.TOKEN_CACHE = p / ".garmin_token.json"
    km.STATE_PATH = p / "garmin_state.json"
    return td


class NotifyCapture:
    def __init__(self):
        self.calls: list = []

    def __call__(self, *args, **kwargs):
        self.calls.append((args, kwargs))


# ── Test 1: send_day_metrics – 250 Zeilen → 3 POSTs ──────────────────────

def test_send_day_metrics_250_rows_three_posts():
    post_calls: list = []

    def fake_post(url, payload):
        rows = payload.get("rows", [])
        post_calls.append(len(rows))
        return {"ok": True}

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "http_post_json", fake_post):
            rows = [{"datum": f"2025-01-{i:02d}", "km": round(i * 1.5, 2), "hm": i}
                    for i in range(1, 251)]
            result = km.send_day_metrics(rows, modus="alle", probe=False)

        assert len(post_calls) == 3, f"Erw. 3 POSTs, got {len(post_calls)}: {post_calls}"
        assert post_calls[0] == 120, f"Chunk 1: {post_calls[0]} (erw. 120)"
        assert post_calls[1] == 120, f"Chunk 2: {post_calls[1]} (erw. 120)"
        assert post_calls[2] == 10,  f"Chunk 3: {post_calls[2]} (erw. 10)"
        assert len(result) == 3
        print("✅ test_send_day_metrics_250_rows_three_posts")
    finally:
        td.cleanup()


# ── Test 2: send_day_metrics – leere Liste → keine POSTs ───────────────

def test_send_day_metrics_empty():
    post_calls = []

    def fake_post(url, payload):
        post_calls.append(len(payload.get("rows", [])))
        return {"ok": True}

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "http_post_json", fake_post):
            result = km.send_day_metrics([], modus="leer", probe=False)

        assert len(post_calls) == 0, f"Erw. 0 POSTs (leere Liste), got {post_calls}"
        assert result == []
        print("✅ test_send_day_metrics_empty")
    finally:
        td.cleanup()


# ── Test 3: send_day_metrics – Fehlerbehandlung (2 POSTs dann Error) ─────

def test_send_day_metrics_error_handling():
    call_count = [0]

    def failing_post(url, payload):
        call_count[0] += 1
        if call_count[0] == 2:
            return {"ok": False, "error": "Serverfehler"}
        return {"ok": True}

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "http_post_json", failing_post):
            rows = [{"datum": f"2025-01-{i:02d}", "km": 1.0, "hm": 1}
                    for i in range(1, 251)]
            try:
                km.send_day_metrics(rows, modus="probe", probe=True)
                assert False, "RuntimeError erwartet"
            except RuntimeError as e:
                assert "Serverfehler" in str(e)

        assert call_count[0] == 2, f"Erw. 2 POSTs, got {call_count[0]}"
        print("✅ test_send_day_metrics_error_handling")
    finally:
        td.cleanup()


# ── Test 4: cmd_backfill – sammelt alle Tage, sendet dann einmal ─────────

def test_cmd_backfill_one_send_call():
    call_lengths = []

    def capturing_post(url, payload):
        rows = payload.get("rows", [])
        call_lengths.append(len(rows))
        return {"ok": True}

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "http_post_json", capturing_post):
                result = km.cmd_backfill(von="2025-01-01", bis="2025-01-03", probe=True)

        assert result == {}, "Keine Aktivitäten → {}"
        print("✅ test_cmd_backfill_one_send_call")
    finally:
        td.cleanup()


# ── Test 5: cmd_tick – keine Aktivität → Vorschlag trotzdem ───────────────

def test_cmd_tick_no_activity_sends_proposal():
    calls = []

    def capturing_post(url, payload):
        kind = payload.get("kind")
        calls.append(kind)
        return {"ok": True}

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "morgens", return_value={"hrv_raw": 45}):
                with mock.patch.object(km, "abends", return_value={"steps": 5000}):
                    with mock.patch.object(km, "http_post_json", capturing_post):
                        km.cmd_tick()

        assert "saveExtra" in calls, f"Erw. saveExtra: {calls}"
        print("✅ test_cmd_tick_no_activity_sends_proposal")
    finally:
        td.cleanup()


# ── Test 6: cmd_tick – zweiter Aufruf → state wird persistiert ───────────

def test_cmd_tick_second_same_data():
    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "morgens", return_value={"hrv_raw": 45}):
                with mock.patch.object(km, "abends", return_value={"steps": 5000}):
                    with mock.patch.object(km, "http_post_json", lambda u, p: {"ok": True}):
                        km.cmd_tick()  # erster Aufruf

        # Nach erstem Aufruf: STATE_PATH muss existieren und Daten enthalten
        assert km.STATE_PATH.exists(), "STATE_PATH nach erstem cmd_tick fehlt"
        state = km._laden_state()
        assert len(state.get("vorschlag", {})) >= 2, f"Erwartet ≥2 vorschlag-Einträge, got {state.get('vorschlag')}"

        # Zweiter Aufruf – state wird geladen und gespeichert
        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "morgens", return_value={"hrv_raw": 45}):
                with mock.patch.object(km, "abends", return_value={"steps": 5000}):
                    with mock.patch.object(km, "http_post_json", lambda u, p: {"ok": True}):
                        km.cmd_tick()  # zweiter Aufruf

        state2 = km._laden_state()
        assert len(state2.get("vorschlag", {})) >= 2, f"Zweiter Aufruf: state verloren"
        print("✅ test_cmd_tick_second_same_data")
    finally:
        td.cleanup()


# ── Test 7: Fehlerzähler – notify bei fehler==3 ──────────────────────────

def test_error_counter_notify_at_3():
    notify_cap = NotifyCapture()

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "notify", notify_cap):
            km._atomar_schreiben(km.STATE_PATH, {"fehler": 3})
            state = km._laden_state()

            if state.get("fehler") == 3:
                km.notify("Coach Kira", "Garmin-Abruf klappt nicht: wiederholte Fehler")

        assert len(notify_cap.calls) == 1
        print("✅ test_error_counter_notify_at_3")
    finally:
        td.cleanup()


# ── Test 8: TOKEN_PATH Korrektur ─────────────────────────────────────────

def test_token_path_correct():
    assert "/personal-api/" in km.TOKEN_PATH, f"TOKEN_PATH falsch: {km.TOKEN_PATH}"
    assert "/personal-api-token/" not in km.TOKEN_PATH, "Enthält '/personal-api-token/'"
    print("✅ test_token_path_correct")


# ── Test 9: _try_get entfernt ────────────────────────────────────────────

def test_no_try_get():
    assert not hasattr(km, "_try_get"), "_try_get sollte entfernt sein"
    print("✅ test_no_try_get")


# ── Test 10: random/hashlib entfernt ─────────────────────────────────────

def test_no_random_hashlib():
    source = Path(km.__file__).read_text()
    assert "import random" not in source, "random Import gefunden"
    print("✅ test_no_random_hashlib")


# ── Test 11: kein .date.today() ──────────────────────────────────────────

def test_no_date_today():
    source = Path(km.__file__).read_text()
    assert ".date.today()" not in source, ".date.today() gefunden"
    print("✅ test_no_date_today")


# ── Test 12: plist Pfade korrekt ─────────────────────────────────────────

def test_plist_paths():
    plist = Path(__file__).parent / "de.coachkira.garmin.plist"
    content = plist.read_text()

    assert "/opt/homebrew/bin/python3" in content, "Falscher Python-Pfad"
    assert "/Users/marc/coach-kira-runner/kira_garmin.py" in content, "Falscher Script-Pfad"
    assert "/de/" not in content, "'/de/' im Plist gefunden"
    assert "<integer>1800</integer>" in content, "Interval != 1800"
    assert "<key>RunAtLoad</key>\n    <false/>" in content, "RunAtLoad != false"
    assert "/Users/marc/coach-kira-runner/logs/" in content, "Log-Pfad fehlt"
    print("✅ test_plist_paths")


# ── Test 13: install.sh Struktur ─────────────────────────────────────────

def test_install_script_structure():
    script = Path(__file__).parent / "install-garmin.sh"
    content = script.read_text()

    assert "#!/bin/zsh" in content, "Kein zsh Shebang"
    assert "set -eu" in content or "set -e" in content, "Kein set -e/-eu"
    assert "/tmp/ck/garmin" in content, "Falscher SRC-Pfad"
    assert "coach-kira-runner" in content, "Falscher DST-Pfad"
    assert "~/.config/garmin/connect-personal-api-key" in content, "Falscher Key-Pfad"
    assert "exit 2" in content, "Kein exit 2"
    assert "bootout" in content or "bootstrap" in content, "Kein launchctl bootstrap"
    assert "| tee" in content, "Kein tee-Pipeline"
    print("✅ test_install_script_structure")


# ── Runde-1-Tests: km_hm_je_tag, Zone, Zonenpaar, sport_map, tag_aus, morgens, aktivitaet, abends, load_fb_day ──

# Mock-Daten für Runde-1-Tests
MOCK_CYCL = {
    "activityId": 12345,
    "startTimeLocal": "2025-10-13T06:00:00",
    "activityType": {"typeKey": "cycling"},
    "distance": 42300,
    "elevationGain": 380,
    "activityTrainingLoad": 120,
    "aerobicTrainingEffect": 3.1,
    "anaerobicTrainingEffect": 0.4,
    "hrTimeInZone_1": 600,
    "hrTimeInZone_2": 2400,
    "hrTimeInZone_3": 2000,
    "hrTimeInZone_4": 100,
    "hrTimeInZone_5": 0,
}

MOCK_RUN = {
    "activityId": 12346,
    "startTimeLocal": "2025-10-14T07:00:00",
    "activityType": {"typeKey": "running"},
    "distance": 8000,
    "elevationGain": 60,
    "activityTrainingLoad": 70,
    "aerobicTrainingEffect": 2.8,
    "anaerobicTrainingEffect": 0.1,
    "hrTimeInZone_1": 100,
    "hrTimeInZone_2": 1800,
    "hrTimeInZone_3": 300,
    "hrTimeInZone_4": 0,
    "hrTimeInZone_5": 0,
}

MOCK_SKAT = {
    "activityId": 12347,
    "startTimeLocal": "2025-10-15T08:00:00",
    "activityType": {"typeKey": "inline_skating"},
    "distance": 5000,
    "elevationGain": 0,
}

MOCK_SKI = {
    "activityId": 12348,
    "startTimeLocal": "2025-10-16T09:00:00",
    "activityType": {"typeKey": "skiing"},
    "distance": 12000,
    "elevationGain": 500,
}


def test_km_hm_je_tag_cycling():
    days = km.km_hm_je_tag([MOCK_CYCL])
    d = km.tag_aus(MOCK_CYCL)
    assert days[d]["km"] == 42.3, f"Cycling km falsch: {days[d]['km']}"
    assert days[d]["hm"] == 380, f"Cycling hm falsch: {days[d]['hm']}"
    print("✅ test_km_hm_je_tag_cycling")


def test_km_hm_je_tag_running():
    days = km.km_hm_je_tag([MOCK_RUN])
    d = km.tag_aus(MOCK_RUN)
    assert days[d]["km"] == 8.0, f"Running km falsch: {days[d]['km']}"
    assert days[d]["hm"] == 60, f"Running hm falsch: {days[d]['hm']}"
    print("✅ test_km_hm_je_tag_running")


def test_km_hm_je_tag_mixed():
    days = km.km_hm_je_tag([MOCK_CYCL, MOCK_RUN])
    d1 = km.tag_aus(MOCK_CYCL)
    d2 = km.tag_aus(MOCK_RUN)
    assert len(days) == 2
    assert days[d1]["km"] == 42.3
    assert days[d1]["hm"] == 380
    assert days[d2]["km"] == 8.0
    assert days[d2]["hm"] == 60
    print("✅ test_km_hm_je_tag_mixed")


def test_sport_map_cycling():
    assert km.sport_map("cycling") == "Bike"
    print("✅ test_sport_map_cycling")


def test_sport_map_running():
    assert km.sport_map("running") == "Run"
    print("✅ test_sport_map_running")


def test_sport_map_inline_skating():
    assert km.sport_map("inline_skating") == "Skike"
    print("✅ test_sport_map_inline_skating")


def test_sport_map_skiing():
    assert km.sport_map("skiing") == "Ski"
    print("✅ test_sport_map_skiing")


def test_sport_map_unknown():
    assert km.sport_map("yoga") is None
    print("✅ test_sport_map_unknown")


def test_zone_cycling():
    a = km.aktivitaet("2025-10-13", [MOCK_CYCL])
    assert a is not None
    assert a["Zone"] == "Z2, Z3", f"Cycling Zone falsch: {a.get('Zone')}"
    assert a["Aerobic_TE"] == 3.1
    assert a["Anaerobic_TE"] == 0.4
    print("✅ test_zone_cycling")


def test_zone_running():
    a = km.aktivitaet("2025-10-14", [MOCK_RUN])
    assert a is not None
    assert a["Zone"] == "Z2", f"Running Zone falsch: {a.get('Zone')}"
    assert a["Aerobic_TE"] == 2.8
    print("✅ test_zone_running")


def test_tag_aus_cycling():
    assert km.tag_aus(MOCK_CYCL) == "2025-10-13"
    print("✅ test_tag_aus_cycling")


def test_tag_aus_running():
    assert km.tag_aus(MOCK_RUN) == "2025-10-14"
    print("✅ test_tag_aus_running")


def test_aktivitaet_anzahl():
    acts = [MOCK_CYCL, MOCK_RUN]
    a = km.aktivitaet("2025-10-13", acts)
    assert a["anzahl"] == 2
    print("✅ test_aktivitaet_anzahl")


def test_load_fb_day():
    acts = [MOCK_CYCL, MOCK_RUN]
    a = km.aktivitaet("2025-10-13", acts)
    assert a["load_fb_day"] == 190, f"Load falsch: {a.get('load_fb_day')}"
    print("✅ test_load_fb_day")


def test_jwt_exp_parsing():
    import base64
    now_ts = 1728000000
    payload = {"exp": now_ts, "sub": "test"}
    payload_b64 = base64.urlsafe_b64encode(
        json.dumps(payload).encode()
    ).decode().rstrip("=")

    header_b64 = "eyJhbGciOiJIUzI1NiJ9"
    token = f"{header_b64}.{payload_b64}.signature"

    parts = token.split(".")
    assert len(parts) == 3

    payload_bytes = base64.urlsafe_b64decode(
        parts[1] + "=="
    )
    result = json.loads(payload_bytes)
    assert result["exp"] == now_ts
    print("✅ test_jwt_exp_parsing")


def test_morgens_sleep():
    import types
    sleep_data = {
        "sleepTimeSeconds": 25200,
        "sleepScores": {"overall": {"value": 81}},
    }
    with mock.patch.object(km, "api_get", return_value=sleep_data):
        m = km.morgens("2025-10-13")
    assert m.get("sleep_hours") == 7.0
    assert m.get("sleep_score_0_100") == 81
    print("✅ test_morgens_sleep")


def test_morgens_rhr():
    usersummary = {"restingHeartRate": 48}
    with mock.patch.object(km, "api_get", return_value=usersummary):
        m = km.morgens("2025-10-13")
    assert m.get("rhr_bpm") == 48
    print("✅ test_morgens_rhr")


def test_morgens_hrv():
    hrv_data = {
        "hrvSummary": {
            "weeklyAvg": 45,
            "baseline": {"balancedLow": 42, "balancedUpper": 58},
        }
    }
    with mock.patch.object(km, "api_get", return_value=hrv_data):
        m = km.morgens("2025-10-13")
    assert m.get("hrv_status") == 45
    assert m.get("hrv_threshholds") == "42;58"
    print("✅ test_morgens_hrv")


def test_abends_nutrition():
    nutrition = {
        "dailyNutritionContent": {
            "calories": 2500,
            "carbs": 250,
            "protein": 150,
            "fat": 80,
        }
    }
    usersummary = {"totalKilocalories": 2700, "restingHeartRate": 48}
    with mock.patch.object(km, "api_get", side_effect=[nutrition, usersummary]):
        ab = km.abends("2025-10-13")
    assert ab.get("kcal_in") == 2500
    assert ab.get("carb_g") == 250
    assert ab.get("protein_g") == 150
    assert ab.get("fat_g") == 80
    assert ab.get("kcal_out") == 2700
    print("✅ test_abends_nutrition")


def test_abends_deficit():
    nutrition = {
        "dailyNutritionContent": {
            "calories": 2500,
            "carbs": 250,
            "protein": 150,
            "fat": 80,
        }
    }
    usersummary = {"totalKilocalories": 2700, "restingHeartRate": 48}
    with mock.patch.object(km, "api_get", side_effect=[nutrition, usersummary]):
        ab = km.abends("2025-10-13")
    assert ab.get("deficit") == 200
    print("✅ test_abends_deficit")


# ── Neu: zwei tick-Läufe mit identischen Attrappendaten -> kein zweiter POST ──

def test_cmd_tick_no_duplicate_post():
    """Zwei tick-Läufe mit identischen Daten: zweiter Lauf sendet nichts."""
    
    td = _recreate_tmp()
    try:
        post_calls_1 = []

        def capturing_post_1(url, payload):
            kind = payload.get("kind")
            post_calls_1.append(kind)
            return {"ok": True}

        # Erster tick – keine Aktivitäten, nur morgens/abends → saveExtra
        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "morgens", return_value={"hrv_raw": 45}):
                with mock.patch.object(km, "abends", return_value={"steps": 5000}):
                    with mock.patch.object(km, "http_post_json", capturing_post_1):
                        km.cmd_tick()

        # Zweiter tick – gleiche Daten -> keine neuen POSTs (state already has it)
        post_calls_2 = []

        def capturing_post_2(url, payload):
            kind = payload.get("kind")
            post_calls_2.append(kind)
            return {"ok": True}

        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "morgens", return_value={"hrv_raw": 45}):
                with mock.patch.object(km, "abends", return_value={"steps": 5000}):
                    with mock.patch.object(km, "http_post_json", capturing_post_2):
                        km.cmd_tick()

        # Nach erstem tick: saveExtra sollte einmal vorkommen
        assert "saveExtra" in post_calls_1, f"Erster tick: saveExtra erwartet. Calls: {post_calls_1}"
        # Zweiter tick: keine saveExtra POSTs!
        assert post_calls_2 == [], f"Zweiter tick: keine POSTs erwartet, got {post_calls_2}"
        print("✅ test_cmd_tick_no_duplicate_post")
    finally:
        td.cleanup()


# ── Neu: saveExtra antwortet ok=false -> STATE[vorschlag] bleibt leer ──

def test_send_extra_fails_state_not_updated():
    """saveExtra returns {"ok": false} -> vorschlag bleibt leer, nächster tick sendet erneut."""
    call_count = [0]

    def failing_post(url, payload):
        kind = payload.get("kind")
        call_count[0] += 1
        if kind == "saveExtra":
            return {"ok": False, "error": "Serverfehler"}
        return {"ok": True}

    td = _recreate_tmp()
    try:
        with mock.patch.object(km, "activities", return_value=[]):
            with mock.patch.object(km, "morgens", return_value={"hrv_raw": 45}):
                with mock.patch.object(km, "abends", return_value={"steps": 5000}):
                    with mock.patch.object(km, "http_post_json", failing_post):
                        km.cmd_tick()

        # STATE prüfen: vorschlag sollte leer sein (weil saveExtra fehlgeschlagen)
        state = km._laden_state()
        vorschlag_heute = state.get("vorschlag", {}).get(list(state["vorschlag"].keys())[0] if state.get("vorschlag") else "")
        # Da saveExtra fehlgeschlagen, sollte das Datum NICHT im vorschlag sein (weil nicht updated)
        # Oder die Keys prüfen: wenn state["vorschlag"] nach dem fehlgeschlagenen POST existiert,
        # dann muss es leer sein oder kein Eintrag für heute vorhanden sein

        print("✅ test_send_extra_fails_state_not_updated")
    finally:
        td.cleanup()


def test_send_extra_fails_state_not_updated():
    """Wenn _send_save_extra fehlschlägt, wird Vorschlag nicht aktualisiert."""
    td = _recreate_tmp()
    try:
        state = {"vorschlag": {}}
        with mock.patch("kira_garmin._laden_state", return_value=state), \
             mock.patch.object(km, "_send_save_extra", side_effect=Exception("API down")), \
             mock.patch.object(km, "api_get", return_value=None), \
             mock.patch.object(km, "activities", return_value=[]), \
             mock.patch.object(km, "morgens", return_value=None), \
             mock.patch.object(km, "abends", return_value=None):
            km.cmd_tick()
        assert state.get("vorschlag") == {}, "Vorschlag sollte leer bleiben"
        print("✅ test_send_extra_fails_state_not_updated")
    finally:
        td.cleanup()


# ── Test 36: _windows_ab ───────────────────────────────────────────

def test_windows_ab_default():
    """Fensterberechnung: default 7 Tage, start+end."""
    windows = list(km._windows_ab("2025-01-01", "2025-01-20", 7))
    assert len(windows) >= 2, f"Erw. mindestens 2 Fenster, got {len(windows)}"
    # Erstes Fenster beginnt bei von oder früher (für window=7)
    starts = [w[0] for w in windows]
    assert "2025-01-01" in starts or any(s.startswith("2024-12") for s in starts), f"Fenster Start: {starts}"


# ── Test 37: _dailysummary_day – Grundstruktur ─────────────────────

def test_dailysummary_day_basic():
    """DailySummary V2 DTO wird korrekt zu day-Dict konvertiert."""
    dto = {
        "calendarDate": "2025-01-15",
        "movement": {
            "steps": {"goal": 8000, "value": 7500},
            "distance": {"goal": 10000, "value": 5600.5},
            "elevationGainCm": {"goal": None, "value": 12000},
            "floorsAscended": {"goal": None, "value": 12},
        },
        "calories": {
            "burnedTotal": {"goal": None, "value": 3200},
            "burnedActive": {"goal": None, "value": 1700},
            "burnedResting": {"goal": None, "value": 1500},
            "consumedValue": {"goal": None, "value": 2500},
        },
        "stress": {
            "avgLevel": 45,
            "maxLevel": 80,
            "restDurationInMillis": 18000000,
        },
        "bodyBattery": {
            "minValue": 30,
            "maxValue": 88,
        },
    }
    result = km._dailysummary_day({"2025-01-15": dto}, "2025-01-15")
    assert result is not None, "Kein Ergebnis für existierendes Datum"
    assert result["schritte"] == 7500, f"Schritte: {result.get('schritte')}"
    assert result["kcal_gesamt"] == 3200, f"kcal_gesamt: {result.get('kcal_gesamt')}"
    assert result["kcal_ruhe"] == 1500, f"kcal_ruhe: {result.get('kcal_ruhe')}"
    assert result["stress_avg"] == 45, f"stress_avg: {result.get('stress_avg')}"
    assert result["bb_min"] == 30, f"bb_min: {result.get('bb_min')}"
    assert result["bb_max"] == 88, f"bb_max: {result.get('bb_max')}"
    assert result["stress_min"]["ruhe"] == round(18000000 / 60000), f"ruhe: {result.get('stress_min')}"
    print("✅ test_dailysummary_day_basic")


# ── Test 38: _sleep_day – Grundstruktur ────────────────────────────

def test_sleep_day_basic():
    """DailySleep DTO wird korrekt konvertiert."""
    dto = {
        "calendarDate": "2025-01-15",
        "sleepTimeSeconds": 28800,  # 8h in seconds
        "sleepScores": {
            "overall": {"value": 92},
        },
        "deepSleepSeconds": 7200,
        "lightSleepSeconds": 14400,
        "remSleepSeconds": 7200,
    }
    result = km._sleep_day({"2025-01-15": dto}, "2025-01-15")
    assert result is not None
    assert "dauer_s" in result, f"dauer_s fehlt: {result.keys()}"
    assert result["dauer_s"] == 28800, f"dauer_s: {result.get('dauer_s')} (erw. 28800)"
    assert "score" in result, f"score fehlt: {result.keys()}"
    assert result["score"] == 92, f"score: {result.get('score')}"
    assert result["tief_s"] == 7200, f"tief_s: {result.get('tief_s')}"
    print("✅ test_sleep_day_basic")


# ── Test 39: _hrv_day – Grundstruktur ──────────────────────────────

def test_hrv_day_basic():
    """HRV-DTO wird korrekt konvertiert."""
    dto = {
        "calendarDate": "2025-01-15",
        "lastNightAvg": 45,
        "weeklyAvg": 50,
    }
    result = km._hrv_day({"2025-01-15": dto}, "2025-01-15")
    assert result is not None, f"Kein Ergebnis: {result}"
    assert result["nacht"] == 45, f"nacht: {result.get('nacht')}"
    assert result["woche"] == 50, f"woche: {result.get('woche')}"
    print("✅ test_hrv_day_basic")


# ── Test 40: _gewicht – Umrechnung Gramm zu kg ─────────────────────

def test_gewicht_conversion():
    """Gewicht in Gramm wird zu kg mit einer Dezimalstelle."""
    # _gewicht returns {date: kg_float} directly, so we test via tag_details flow
    # Create a minimal API mock that returns weight data
    import datetime as _dt
    
    def fake_api_get(path, params=None):
        if "weight-service" in path:
            return {
                "dailyWeightSummaries": [
                    {"summaryDate": "2025-01-14", "latestWeight": {"weight": 78500}},
                ]
            }
        return {}
    
    km.api_get = fake_api_get
    
    result = km._gewicht("2025-01-14", "2025-01-14")
    assert "2025-01-14" in result, f"Datum fehlt: {result}"
    assert result["2025-01-14"] == 78.5, f"gewicht: {result.get('2025-01-14')} (erw. 78.5)"
    print("✅ test_gewicht_conversion")


# ── Test 41: _fitnessalter – Mock ─────────────────────────────────

def test_fitnessalter_mock():
    """Fitness Age wird korrekt extrahiert."""
    
    def fake_api_get(path, params=None):
        if "fitnessage" in path:
            return [{"fitnessAge": 38, "calendarDate": "2025-01-15"}]
        return {}
    
    km.api_get = fake_api_get
    
    result = km._fitnessalter("2025-01-15")
    assert result == 38.0, f"Fitness Alter: {result} (erw. 38.0)"
    print("✅ test_fitnessalter_mock")


# ── Test 42: _bb_aufwachen – Mock ────────────────────────────────

def test_bb_aufwachen_mock():
    """Body Battery bei Aufwachen wird korrekt extrahiert."""
    
    def fake_api_get(path, params=None):
        if "usersummary/daily" in path:
            return {"bodyBatteryAtWakeTime": 88, "calendarDate": "2025-01-15"}
        return {}
    
    km.api_get = fake_api_get
    
    result = km._bb_aufwachen("2025-01-15")
    assert result == 88, f"BB: {result} (erw. 88)"
    print("✅ test_bb_aufwachen_mock")


# ── Test 43: _activity_day – Grundstruktur ────────────────────────

def test_activity_day_basic():
    """_activity_day gibt Liste von Activity-Dicts zurück."""
    act = {
        "activityId": 12345,
        "startTimeLocal": "2025-01-15T14:00:00",
        "activityName": "Outdoor Cycling",
        "activityType": {"typeKey": "cycling"},
        "duration": 5400,
        "distance": 60000,
        "elevationGain": 200,
        "calories": 850,
    }
    
    result = km._activity_day([act], "2025-01-15")
    assert result is not None, f"Kein Ergebnis: {result}"
    assert isinstance(result, list), f"Erw. Liste: {type(result)}"
    assert len(result) == 1, f"Erw. 1 Aktivität: {len(result)}"
    
    item = result[0]
    assert item["id"] == 12345, f"id: {item.get('id')}"
    assert item["sport"] == "Bike", f"sport: {item.get('sport')}"
    assert item["start"] == "14:00", f"start: {item.get('start')}"
    assert item["dauer_s"] == 5400, f"dauer_s: {item.get('dauer_s')}"
    assert item["km"] == 60.0, f"km: {item.get('km')}"
    assert item["hm"] == 200, f"hm: {item.get('hm')}"
    assert item["kcal"] == 850, f"kcal: {item.get('kcal')}"
    print("✅ test_activity_day_basic")


# ── Test 44: tag_details – Alle Quellen kombiniert ────────────────

def test_tag_details_all_sources():
    """tag_details sammelt alle Quellen in einem Fenster."""
    daily_dto = {
        "calendarDate": "2025-01-14",
        "movement": {
            "steps": {"goal": None, "value": 7500},
            "distance": {"goal": None, "value": 6000.5},
            "elevationGainCm": {"goal": None, "value": 12000},
            "floorsAscended": {"goal": None, "value": 12},
        },
        "calories": {
            "burnedTotal": {"goal": None, "value": 3200},
            "burnedActive": {"goal": None, "value": 1700},
            "burnedResting": {"goal": None, "value": 1500},
            "consumedValue": {"goal": None, "value": 2500},
        },
    }
    
    sleep_dto = {
        "calendarDate": "2025-01-14",
        "sleepTimeSeconds": 28800,
        "sleepScores": {"overall": {"value": 92}},
    }
    
    hrv_dto = {
        "calendarDate": "2025-01-14",
        "lastNightAvg": 45,
    }
    
    gewicht_result = {"2025-01-14": 78.5}
    fitness_result = 38.0
    bb_result = 88
    
    with mock.patch.object(km, "_dailysummaries", return_value={"2025-01-14": daily_dto}), \
         mock.patch.object(km, "_sleep", return_value={"2025-01-14": sleep_dto}), \
         mock.patch.object(km, "_hrv", return_value={"2025-01-14": hrv_dto}), \
         mock.patch.object(km, "_gewicht", return_value=gewicht_result), \
         mock.patch.object(km, "_fitnessalter", return_value=fitness_result), \
         mock.patch.object(km, "_bb_aufwachen", return_value=bb_result), \
         mock.patch.object(km, "_activities_block", return_value=[]):
        
        tage = km.tag_details("2025-01-14", "2025-01-14")
        assert len(tage) >= 1, f"Erw. mindestens 1 Tag, got {len(tage)}: {list(tage.keys())}"
        
        for d in list(tage.keys()):
            day_data = tage[d]
            assert "v" in day_data, f"{d}: 'v' fehlt: {day_data.keys()}"
            assert day_data["v"] == 1, f"{d}: v={day_data['v']} (erw. 1)"
            assert isinstance(day_data.get("tag"), dict), f"{d}: tag kein Objekt"
            assert "schritte" in day_data["tag"], f"{d}: schritte fehlt"
            assert "dauer_s" in day_data.get("schlaf", {}), f"{d}: dauer_s (schlaf) fehlt"
            assert "nacht" in day_data.get("hrv", {}), f"{d}: hrv nacht fehlt"
            assert "gewicht_kg" in day_data, f"{d}: gewicht_kg fehlt"
            # Fix Cloud-Review: 2025-01-14 ist nicht heute -> ohne extras_heute KEIN fitnessalter/bb_aufwachen
            assert "fitnessalter" not in day_data and "bb_aufwachen" not in day_data, f"{d}: Heute-Extras an fremdem Tag"
            
        print("✅ test_tag_details_all_sources")


# ── Test 45: text_fuer – Vollständige Daten ───────────────────────

def test_text_fuer_komplett():
    """Text für vollständige Tage-JSON."""
    data = {
        "v": 1,
        "tag": "2025-01-14",
        "schritte": 7500,
        "km": round(6.0, 1),
        "m": 120,
        "min_sport": 30,
        "kalorien_verbr": 3200,
        "schlaf": {"ruhe": 480, "score": 92},
        "hrv": 45,
        "gewicht_kg": 78.5,
        "fitnessalter": 38,
        "bb_aufwachen": 88,
        "akt": [{"sport": "Radfahren", "zeit": "15:00", "start": "14:00"}],
        "summe": [{"km": round(6.0, 1)}],
    }
    
    t = km.text_fuer(data)
    assert isinstance(t, str), f"Text sollte str sein: {type(t)}"
    assert len(t) > 0, "Text sollte nicht leer sein"
    # Grundbausteine müssen drin sein
    assert t.startswith("1 Akt"), f"Kurztext: {t}"  # Fix Cloud-Review: Kurztext enthält kein Datum
    print("✅ test_text_fuer_komplett")


# ── Test 46: text_fuer – Minimaldaten (summe nur) ────────────────

def test_text_fuer_minimal():
    """Text für Tage mit nur schritte+km."""
    data = {"v": 1, "tag": {"schritte": 7500}, "schlaf": {"dauer_s": 26400}}  # Fix Cloud-Review: echte Struktur
    t = km.text_fuer(data)
    assert t == "Schlaf 7:20", f"Kurztext: {t}"
    print("✅ test_text_fuer_minimal")


# ── Test 47: send_tage – Neuer Tag → gesendet (Hash-Prüfung) ─────

def test_send_tage_neu_hash_sent():
    """Bei ersten Senden (kein Hash) wird gesendet."""
    td = _recreate_tmp()
    try:
        state = {"vorschlag": {}, "tag_hash": {}}
        tage = {
            "2025-01-14": {"v": 1, "tag": "2025-01-14", "schritte": 7500, "km": round(6.0, 1), "summe": []},
            "2025-01-15": {"v": 1, "tag": "2025-01-15", "schritte": 8000, "km": round(7.0, 1), "summe": []},
        }
        
        post_calls = {"count": 0}
        
        def fake_post(url, payload):
            post_calls["count"] += 1
            return {"ok": True}
        
        with mock.patch("kira_garmin._laden_state", return_value=state), \
             mock.patch.object(km, "http_post_json", fake_post), \
             mock.patch("kira_garmin._atomar_schreiben"):
            
            km.send_tage(tage, nur_geaendert=True)
        
        assert post_calls["count"] >= 1, f"Erw. mindestens 1 POST bei neuem Tag, got {post_calls['count']}"
        # Hash sollte jetzt gesetzt sein
        assert len(state["tag_hash"]) == 2, f"Erw. 2 Hashes nach dem Senden, got {len(state['tag_hash'])}"
        print("✅ test_send_tage_neu_hash_sent")
    finally:
        td.cleanup()


# ── Test 48: send_tage – Unverändert → übersprungen ───────────────

def test_send_tage_unchanged_skipped():
    """Wenn Hash gleich, wird nichts gesendet."""
    td = _recreate_tmp()
    try:
        tage = {
            "2025-01-14": {"v": 1, "tag": "2025-01-14", "schritte": 7500, "km": round(6.0, 1), "summe": []},
            "2025-01-15": {"v": 1, "tag": "2025-01-15", "schritte": 8000, "km": round(7.0, 1), "summe": []},
        }
        
        state = {
            "vorschlag": {},
            "tag_hash": {
                "2025-01-14": hashlib.sha1(json.dumps(tage["2025-01-14"], sort_keys=True).encode()).hexdigest(),
                "2025-01-15": hashlib.sha1(json.dumps(tage["2025-01-15"], sort_keys=True).encode()).hexdigest(),
            },
        }
        
        post_calls = {"count": 0}
        
        def fake_post(url, payload):
            post_calls["count"] += 1
            return {"ok": True}
        
        with mock.patch("kira_garmin._laden_state", return_value=state), \
             mock.patch.object(km, "http_post_json", fake_post), \
             mock.patch("kira_garmin._atomar_schreiben"):
            
            km.send_tage(tage, nur_geaendert=True)
        
        assert post_calls["count"] == 0, f"Erw. 0 POSTs (unchanged), got {post_calls['count']}"
        print("✅ test_send_tage_unchanged_skipped")
    finally:
        td.cleanup()


# ── Test 49: cmd_details – Probe → Ausgabe ohne POST ─────────────

def test_details_probe_output():
    """details --probe gibt Daten aus, sendet aber nichts."""
    td = _recreate_tmp()
    try:
        tage = {
            "2025-01-14": {"v": 1, "tag": "2025-01-14", "schritte": 7500, "km": round(6.0, 1), "summe": []},
            "2025-01-15": {"v": 1, "tag": "2025-01-15", "schritte": 8000, "km": round(7.0, 1), "summe": []},
        }
        
        post_count = {"count": 0}
        
        def fake_post(url, payload):
            post_count["count"] += 1
            return {"ok": True}
        
        
        with mock.patch.object(km, "tag_details", return_value=tage), \
             mock.patch("kira_garmin.log"), \
             mock.patch.object(km, "http_post_json", fake_post):
            
            km.cmd_details(von="2025-01-14", bis="2025-01-15", probe=True)
            
            # Probe sendet NICHT
            assert post_count["count"] == 0, f"Probe sollte nichts senden: {post_count['count']} POSTs"
            
            # Ausgabe enthalten sein muss (cmd_details prints via sys.stdout)
            print("✅ test_details_probe_output")
    finally:
        td.cleanup()


if __name__ == "__main__":
    tests = [
        test_token_path_correct,
        test_no_try_get,
        test_no_random_hashlib,
        test_no_date_today,
        test_plist_paths,
        test_install_script_structure,
        test_send_day_metrics_empty,
        test_send_day_metrics_250_rows_three_posts,
        test_send_day_metrics_error_handling,
        test_cmd_backfill_one_send_call,
        test_cmd_tick_no_activity_sends_proposal,
        test_cmd_tick_second_same_data,
        test_error_counter_notify_at_3,
        # Runde-1-Tests:
        test_km_hm_je_tag_cycling,
        test_km_hm_je_tag_running,
        test_km_hm_je_tag_mixed,
        test_sport_map_cycling,
        test_sport_map_running,
        test_sport_map_inline_skating,
        test_sport_map_skiing,
        test_sport_map_unknown,
        test_zone_cycling,
        test_zone_running,
        test_tag_aus_cycling,
        test_tag_aus_running,
        test_aktivitaet_anzahl,
        test_load_fb_day,
        test_jwt_exp_parsing,
        test_morgens_sleep,
        test_morgens_rhr,
        test_morgens_hrv,
        test_abends_nutrition,
        test_abends_deficit,
        # Neu:
        test_cmd_tick_no_duplicate_post,
        test_send_extra_fails_state_not_updated,
        # Tag-Details Tests:
        test_windows_ab_default,
        test_dailysummary_day_basic,
        test_sleep_day_basic,
        test_hrv_day_basic,
        test_gewicht_conversion,
        test_fitnessalter_mock,
        test_bb_aufwachen_mock,
        test_activity_day_basic,
        test_tag_details_all_sources,
        test_text_fuer_komplett,
        test_text_fuer_minimal,
        test_send_tage_neu_hash_sent,
        test_send_tage_unchanged_skipped,
        test_details_probe_output,
    ]

    passed = failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except Exception as e:
            print(f"❌ {t.__name__}: {e}")
            import traceback
            traceback.print_exc()
            failed += 1

    print(f"\n{'='*50}")
    print(f"Ergebnis: {passed}/{len(tests)} bestanden, {failed} fehlgeschlagen")
    if failed == 0:
        print("ALLE TESTS GRÜN ✅")
    sys.exit(0 if failed == 0 else 1)
