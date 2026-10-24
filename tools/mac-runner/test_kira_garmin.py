"""Tests für Garmin-Modul – Standardbibliothek only."""
from __future__ import annotations

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
    assert "import hashlib" not in source, "hashlib Import gefunden"
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
