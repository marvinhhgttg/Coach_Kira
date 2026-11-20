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

## Etappe 3b – braucht deine Entscheidung

Die folgenden Punkte ändern Bedienung oder Geräte:

1. **Anmeldung im Cockpit:**
   - Einmal je Gerät den PIN eingeben. Die Sitzung gilt danach 45 Tage, auf Wunsch länger.
   - Cloud-KI, Kopplung, Garmin-Knöpfe und Schreibaktionen laufen dann nur mit gültiger Sitzung. Die iPhone-, iPad- und Mac-App fragen den PIN einmal in ihrer Ansicht ab.
2. **Mac-Geheimnis:**
   - Der Server und `kira_runner.py` auf dem Mac teilen sich einen Schlüssel.
   - Eine Übergangszeit, in der der Server beides annimmt, verhindert Ausfälle.
3. **Apps Script wieder schützen (größte Lücke):**
   - Den Token-Schutz für `action=…` wieder einschalten. Das hattest du am 28.09. bewusst abgeschaltet.
   - Schreibaktionen gehen dann nur noch über den Cockpit-Server (`/api/action`, mit Sitzung). Der Token liegt dann nur dort und nicht mehr im Browser.
   - Lesen über die Web-App-URL bleibt offen, solange sie für „Jeder“ freigegeben ist.

Empfohlene Reihenfolge: 1 → 3 → 2. Jeder Schritt wird zuerst in v2 getestet und dann live geschaltet.

## Rückweg 3a

- Die vorige `server.mjs` liegt im Git-Verlauf (Commit vor Etappe 3), `package-lock.json` ebenso.
- Danach erneut veröffentlichen.
