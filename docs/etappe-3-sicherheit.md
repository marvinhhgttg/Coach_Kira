# Etappe 3 – Live absichern

Stand 20.11.2026, Abend. Freigabe: Marc („Mach die Abnahme. Und dann Etappe 3“).

## Ausgangslage (gemessen auf coachkira.pplx.app vor dem Umbau)

| Punkt | Befund |
|---|---|
| PIN/Ersteinrichtung | Nie eingerichtet (`configured: false`). `POST /api/auth/setup` war für jeden offen. Wer zuerst kommt, setzt PIN und Token. |
| Herkunftsprüfung | Präfixvergleich plus Joker für **jede** `*.pplx.app`-Seite. Fremde Seiten unter pplx.app hätten Kopplungscodes erzeugen, Geräte widerrufen und die Cloud-KI nutzen können. |
| Gerätekopplung, Geräteliste, Widerruf, Cloud-KI | Ohne Origin-Header (curl) offen |
| Mac-Endpunkte | ohne gemeinsames Geheimnis |
| npm audit | 1 kritisch (proxy-addr), 1 hoch, 1 mittel (qs) |
| Apps Script | Die Web-App ist für „Jeder“ freigegeben, der Token-Schutz für `action=…` ist seit 28.09.2026 auf Wunsch aus. Die URL steht im öffentlichen Cockpit-Code. Damit kann jeder, der sie kennt, lesen **und** schreiben. |

## Etappe 3a – umgesetzt und live (ohne Änderung an Bedienung, Mac, iPhone oder Apps Script)

1. **Ersteinrichtung gesperrt:** `POST /api/auth/setup` nimmt jetzt nur noch Anfragen an, wenn der Server mit `CK_SETUP_OFFEN=1` gestartet ist. Eine bestehende Einrichtung lässt sich nur mit gültiger Sitzung ändern. Das Cockpit selbst nutzt diesen Weg nicht.
2. **Exakte Herkunft statt Joker:** Erlaubt sind nur noch `https://coachkira.pplx.app` und für Tests `http://localhost` bzw. `127.0.0.1`. Andere `*.pplx.app`-Seiten und Täuschungen wie `coachkira.pplx.app.evil.com` werden abgewiesen.
3. **Nur aus dem Cockpit erreichbar:** Für folgende Endpunkte muss die Anfrage nachweislich aus dem Cockpit kommen (Origin bzw. `Sec-Fetch-Site: same-origin` mit passendem Referer): Cloud-KI (`/api/llm`), „Garmin jetzt abrufen“, Workout senden, Kopplungscode, Geräteliste, Widerruf und Abgleich. Eine Anfrage ohne jede Herkunft, etwa per curl, bekommt 403.
4. **Abhängigkeiten:** `npm audit fix`. Danach ist express 5.2.1 drin, proxy-addr 2.0.8 und qs 6.17.1. Ergebnis: **0 Lücken**.

**Unverändert erreichbar** sind die Wege, die der Mac-Runner und die Apps nutzen:
- Mac-Runner: `garmin/*` GET und Status, `texte/*`, `lokal/naechster` und `lokal/ergebnis`.
- App mit Geräteschlüssel: `/api/widget`, `/api/widget/aktion`, `/api/widget/verbinden`.

**Prüfung nach Veröffentlichung (live):**
- Ersteinrichtung, Kopplung mit fremder Herkunft, Geräteliste ohne Herkunft und Cloud-KI ohne bzw. mit fremder Herkunft liefern jeweils 403.
- Die Mac-Wege antworten mit 200, `/api/widget` ohne Schlüssel mit 401.
- Der Mac meldet sich weiter (letzter Puls vor 13 s).
- Im Browser lädt das Cockpit ohne Fehler, und die Geräteliste ist sichtbar.
- **Alle 7 gekoppelten Geräte sind erhalten.**

**Grenze von 3a:** Herkunftsangaben lassen sich außerhalb eines Browsers fälschen. 3a schützt deshalb vor Missbrauch über fremde Webseiten, auch über andere pplx.app-Seiten, und vor einfachen Direktaufrufen. Gegen gezielte Angriffe mit gefälschten Headern hilft nur eine echte Anmeldung. Das ist Etappe 3b.

## Etappe 3b – umgesetzt und live (20.11.2026)

Reihenfolge wie entschieden: 1 → 2 → 3.

### 1. PIN-Anmeldung im Cockpit
- Ersteinrichtung nur mit einmaligem Einrichtungscode (Server kennt nur den Hash, `CK_SETUP_HASH`). PIN: 6–12 Ziffern, zweimal eingeben.
- Nach 5 Fehlversuchen 15 Minuten Sperre. Sitzung gleitet 45 Tage, Cookie hält 400 Tage.
- Frontend: `AnmeldeTor.tsx` (Anmeldemaske und 401-Wächter). Widgets laufen ohne Anmeldung weiter.
- Zentrale Routenprüfung in `server/zugang.mjs` mit drei Klassen: `cockpit` (Browser + Sitzung), `mac` (Mac-Schlüssel), `gemischt` (Sitzung oder Mac).

### 2. Apps Script mit Token und Schreiben über den Server
- Patch D (additiv) in `KiraGeminiSupervisor.gs`: `ckSec_pruefen_` prüft `token` als URL-Parameter bei `action=…` (GET) und bei jedem POST. Ausnahme: Telegram-Webhook (`update_id`).
- Schalter in der Skripteigenschaft `CK_TOKEN_PFLICHT`: `aus` | `protokoll` | `an`. Seit 20.11.2026, 19:58 Uhr: **an**. Abgewiesene Aufrufe zählt `CK_SEC_OHNE`, abrufbar mit `?mode=sicherheit`.
- Der Token liegt nur auf dem Server (`/api/auth/token`, einmalig) und auf dem Mac (`zugang.json`). Der Browser schreibt über `/api/gs` (`gsWeg.ts`), der Server hängt den Token an.
- Live-Bereitstellung v978 → v982. Backup: `apps-script-backup/2026-11-20-nach-patch-d/` (ohne Token).
- Nachweis: POST ohne Token → `token_required`. POST mit Token → `ok`. Lesen (`mode=status`, `mode=chartData`) unverändert offen.

### 3. Mac-Schlüssel
- `kira_runner.py` sendet `x-ck-mac` und den Apps-Script-Token über den Helfer `mit_zugang` (Schlüssel aus `~/coach-kira-runner/zugang.json`).
- Server erzwingt den Schlüssel (`CK_MAC_PFLICHT=an`). Ohne Schlüssel oder nur mit Origin → 403. Mac-Puls 4 s nach dem Scharfschalten bestätigt.

### Rückweg 3b
- Apps Script: Funktion `ckSec_modusProtokoll` oder `ckSec_modusAus` im Editor ausführen. Das wirkt sofort, ohne neue Bereitstellung.
- Mac: `kira_runner.vor-3b.py` zurückkopieren. Server: `CK_MAC_PFLICHT` weglassen und neu veröffentlichen.

### Bekannte Restrisiken
- Die Apps-Script-Web-App ist für „Jeder“ lesbar.
- HTML-Seiten mit `google.script.run` können öffentliche Funktionen aufrufen.
- Die Telegram-Ausnahme lässt sich mit einem gefälschten `update_id` vortäuschen.

## Rückweg 3a

- Die vorige `server.mjs` liegt im Git-Verlauf (Commit vor Etappe 3), `package-lock.json` ebenso.
- Danach erneut veröffentlichen.


## Restrisiken abgearbeitet (25.11.2026)

- Telegram: Es gibt keinen Telegram-Eingang (kein Webhook, kein getUpdates). Die `update_id`-Ausnahme von der Token-Pflicht gilt nur noch mit gültigem `tg_key` (TELEGRAM_WEBHOOK_KEY, nicht gesetzt → Ausnahme zu). Gefälschte Telegram-POSTs werden mit `token_required` abgewiesen. Ausgang (sendMessage) unverändert, Test `ckSec_telegramTest()` erfolgreich.
- Alte HTML-Seiten (page=chat/plan/…, Standardseite WebApp_V2): abgeschaltet, Antwort ist ein kurzer Text. Rückweg `ckSec_seitenAn()`. Zähler: `?mode=sicherheitLesen`.
- Lesen: Schalter CK_LESE_PFLICHT (aus | protokoll | an), Stand `protokoll` (nur zählen). Server und Rechner-Kindprozesse hängen den Token an (server/gsFetch.mjs). Browser liest über /api/gs, ohne Sitzung fällt er auf den direkten Weg zurück (wird gezählt). Scharf schalten mit `ckSec_leseAn()` erst, wenn der Zähler nur noch erwartete Quellen zeigt; Rückweg `ckSec_leseAus()`.
- Beide Bereitstellungen (live und v2) auf Version 983. Sicherungen: apps-script-backup/2026-11-25-vor-restrisiken und -nach-restrisiken.

### Lese-Pflicht scharf (26.11.2026, 05:38)

- Zähler vor dem Scharfschalten: fast nur Testinstanz v2 (status, planSimulationJson, kernVotum), dazu wenige extras/rohdaten.
- Entscheidung Marc: v2 ohne Daten. `ckSec_leseAn()` ausgeführt, CK_LESE_PFLICHT = an.
- Geprüft: Aufruf ohne Token → `token_required`; Live-Votum, Zeitreihe und Cockpit (Comet, angemeldet) laden normal.
- Rückweg: `ckSec_leseAus()` oder `ckSec_leseProtokoll()`.
