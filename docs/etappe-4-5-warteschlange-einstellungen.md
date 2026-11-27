# Etappe 4 und 5 – Dauerhafte Warteschlange und zentrale Einstellungen

Stand 20.11.2026, live auf coachkira.pplx.app (v2 mit demselben Code, dort ohne Wirkung).

## Etappe 4 – Aufträge an den Mac dauerhaft

**Vorher:** Lokale KI-Anfragen von iPhone, iPad oder einem fremden Rechner lagen nur im Arbeitsspeicher des Cockpit-Servers (`server/lokal_relay.mjs`). Ein Neustart oder eine Veröffentlichung ließ offene Anfragen verschwinden.

**Jetzt:**
- **Speicherort:** Die Aufträge liegen in der Server-Datenbank (`settings` › `lokal_auftraege`) und überstehen Neustarts und Veröffentlichungen.
- **Datensparsam:**
  - Die Nachrichten mit den Gesundheitswerten werden gelöscht, sobald der Mac den Auftrag abholt.
  - Die Antwort wird 2 Minuten nach dem ersten Abruf gelöscht.
  - Spätestens nach 30 Minuten verschwindet der ganze Auftrag.
- **Unverändert:**
  - Höchstens 5 offene Aufträge.
  - Nur das Modell `coach-kira`, kein Cloud-Modell.
  - Holt der Mac einen Auftrag nicht innerhalb von 3 Minuten ab, gilt er als Fehler.
- **Nachweis (lokal):**
  - Auftrag anlegen, dann Neustart: Er wartet weiter.
  - Der Mac holt ihn ab: Die Nachricht ist danach aus der Datenbank gelöscht.
  - Zweiter Neustart, dann liefert der Mac das Ergebnis: Das Cockpit bekommt die Antwort.

## Etappe 5 – Einstellungen für alle Geräte

**Zentral auf dem Server** (`server/einstellungen.mjs`, `GET/POST /api/einstellungen`, nur mit PIN-Sitzung):

| Schlüssel | Bedeutung | Erlaubte Werte |
|---|---|---|
| `ck_votum_erfahrung_v2` | Erfahrungsmodell ins Votum einrechnen | `an` · `aus` |
| `ck_kira_engine_v2` | KI für Chat und Briefings | `local` · `pplx` |
| `ck_es_datum` | Ausdauer-Ziel: Datum | `JJJJ-MM-TT` |
| `ck_es_ziel` | Ausdauer-Ziel: Score | 1000–20000 |
| `ck_aufstehen` | Schlafkonto: Aufstehzeit | `hh:mm` |

**Bleibt je Gerät:** Hell/Dunkel, der zuletzt offene Reiter und Zwischenspeicher wie Chatverlauf, Fazit, Wetter und Kalender.

**Ablauf** (`src/lib/geraeteSync.ts`):
- Nach der Anmeldung lädt das Cockpit die Werte vom Server, bevor es etwas anzeigt. Es wartet dabei höchstens 3 Sekunden.
- Der Server gewinnt. Was er noch nicht kennt, aber auf dem Gerät gesetzt ist, lädt das Gerät hoch. Das erste Gerät, das du nach dem Update öffnest, legt also die Startwerte fest.
- Jede Änderung meldet das Cockpit automatisch an den Server.
- Beim Zurückkehren in die App lädt es erneut: Die KI-Wahl und der Erfahrungsschalter wechseln sofort, die übrigen Werte beim nächsten Öffnen der Kachel.

**Fehler behoben:** Das Ausdauer-Ziel (Score) wurde gespeichert, beim Öffnen aber nie gelesen, und stand deshalb immer auf 7000 (`EsHebelTile.tsx`).

**Nachweis (lokal, zwei getrennte Browser):**
- Gerät A ändert KI-Wahl, Ziel-Score und Erfahrungsschalter: Gerät B hat die Werte nach der Anmeldung.
- Gerät B ändert die KI-Wahl: Gerät A übernimmt sie beim Zurückkehren.
- Hell/Dunkel bleibt getrennt.
- Ohne Sitzung kommt 401, ungültige Werte lehnt der Server ab.

## Rückweg

Die vorigen Dateien liegen im Git-Verlauf. Danach erneut veröffentlichen. Die gespeicherten Werte bleiben in der Datenbank und schaden nicht.
