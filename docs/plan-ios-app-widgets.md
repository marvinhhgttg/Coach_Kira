# Coach Kira als iPhone-App mit Homescreen-Widgets – Plan

Stand: 19.11.2026 · Stufe 2: native Hülle plus Widgets. Das Cockpit bleibt die eine Quelle.

## 1. Ziel

- **App:** Coach Kira ist eine echte App auf iPhone (und optional iPad/Mac). Sie zeigt das bestehende Cockpit, ohne Safari-Leiste und mit eigenem Symbol.
- **Widgets:** Auf dem Homescreen und dem Sperrbildschirm steht der Tagesstand, ohne die App zu öffnen.
- **Gleiche Zahlen:** Alle Werte im Widget sind dieselben wie im Cockpit. Es gibt keine zweite Rechenlogik („Basis sind immer die Zahlen der Vergangenheit“, Konsistenz vor allem).
- **KI-Regeln bleiben:** Lokal hat Vorrang. Cloud läuft nur nach deiner Wahl oder über den freigegebenen Cloud-Ersatz nach 3 × 5 min.

## 2. Aufbau

```
iPhone
 ├─ Coach-Kira-App (SwiftUI)
 │   └─ Web-Ansicht: https://coachkira.pplx.app  (Heute, Eingabe, Plan, Analyse …)
 ├─ Widget-Erweiterung (WidgetKit)
 │   └─ holt /api/widget  →  zeigt Tagesstand
 └─ Gemeinsamer Speicher (App Group)
     └─ Zugangsschlüssel + letzter Widget-Stand (für offline)

Cockpit-Server (coachkira.pplx.app, server.mjs)
 └─ NEU /api/widget  →  rechnet mit derselben Logik wie das Cockpit
     └─ Daten aus Apps Script (chartData, planJsonV2, briefingSources …)
```

- **Die App rechnet nichts selbst.** Die Widget-Werte kommen vom Cockpit-Server.
- **Dieselbe Logik wie im Cockpit:** Votum, Readiness-Einordnung, ESS-Soll und ACWR berechnet der Server mit demselben TypeScript-Code wie das Cockpit. Er wird dafür gebündelt, genauso wie `kira_prompts.mjs` für den Mac. Ändert sich eine Regel im Cockpit, ändert sie sich automatisch auch im Widget.
- **Ohne Netz:** Das Widget zeigt den letzten Stand mit Uhrzeit, zum Beispiel „Stand 07:42“.

## 3. Widgets

### Klein: Tagesentscheid

```
┌────────────────┐
│ HEUTE      ●   │  ● = Votum-Farbe (grün/gelb/orange/rot/blau)
│ GRÜN           │
│ Training ok    │
│ Readiness 72   │
│ Stand 07:42    │
└────────────────┘
```

### Mittel: Morgen-Lage

```
┌──────────────────────────────────┐
│ ● GRÜN · Training ok    Do 19.11.│
│ Readiness 72   Schlaf 84         │
│ HRV ausgeglichen   RHR 48        │
│ ESS heute: Soll 120 · Ist 0      │
│ Kira-Texte: 3/3 fertig (lokal)   │
└──────────────────────────────────┘
```

### Groß: Woche und Ziel

```
┌──────────────────────────────────┐
│ ● GRÜN · Training ok    Do 19.11.│
│ Readiness 72 · Schlaf 84 · HRV ok│
│ ESS heute Soll 120 · Ist 0       │
│ ACWR 1,08 (im Korridor)          │
│ ▁▃▅▂▇▁▄  Last 7 Tage (Ist/Soll)  │
│ Nächstes Ziel: Skitour Allgäu    │
│   in 23 Tagen · Tourform 86 %    │
│ Kira-Texte: KI-Analyse fertig …  │
│ [Garmin abrufen] [Briefing lokal]│  ← Knöpfe direkt im Widget
└──────────────────────────────────┘
```

### Sperrbildschirm (und Apple Watch Smart Stack)

- **Rund:** Readiness als Ring, mittig die Zahl, Ringfarbe ist das Votum.
- **Rechteckig:** „GRÜN · Soll 120 ESS · ACWR 1,08“.
- **Inline:** über der Uhrzeit „Kira: GRÜN · 120 ESS“.

### Tippen und Knöpfe

- **Tippen:** Das Widget öffnet die App an der passenden Stelle. Das kleine und das mittlere Widget führen zu Heute, die Zeile „Kira-Texte“ zur Eingabe, die Zielzeile zu Plan → Ziele.
- **Knöpfe im großen Widget:** Sie lösen eine Aktion aus, ohne die App zu öffnen.
  - „Garmin abrufen“ nutzt den bestehenden Mac-Auftrag.
  - „Briefing lokal“ nutzt den bestehenden Lauf „Manuell starten“ auf dem Mac.
  - In der Cloud läuft hier nichts. Ein Cloud-Knopf kommt nur dazu, wenn du das ausdrücklich willst.

## 4. Neue Server-Schnittstelle `/api/widget`

Antwort (Beispiel):

```json
{
  "ok": true,
  "stand": "2026-11-19T07:42:00+01:00",
  "datum": "2026-11-19",
  "votum": { "stufe": "gruen", "text": "Training ok", "grund": "Readiness 72, Schlaf 84" },
  "morgen": { "readiness": 72, "schlaf": 84, "hrvStatus": "ausgeglichen", "rhr": 48, "eingetragen": true },
  "last": { "sollHeute": 120, "istHeute": 0, "acwr": 1.08, "acwrBand": "korridor",
            "woche": [{ "d": "2026-11-13", "ist": 95, "soll": 100 }] },
  "ziel": { "name": "Skitour Allgäu", "datum": "2026-12-12", "tage": 23, "tourform": 86 },
  "texte": { "fertig": 3, "gesamt": 3, "offen": [], "herkunft": "lokal" },
  "mac": { "erreichbar": true }
}
```

- **Berechnung:** Werte, die das Cockpit schon zeigt, kommen aus derselben Funktion. Das gilt für Votum, Readiness, ACWR-Korridor, Tourform und den Texte-Status aus dem Texte-Wächter.
- **Zwischenspeicher:** Der Server hält das Ergebnis 5 Minuten vor. So treffen viele Widget-Abfragen nicht jedes Mal das Apps Script.
- **Abgleich:** Ein Test vergleicht `/api/widget` mit den Werten auf der Heute-Seite. Weichen sie ab, gilt das als Fehler.

## 5. Sicherheit

- **Gesundheitsdaten:** `/api/widget` enthält solche Daten und ist deshalb nicht offen. Die App bekommt einen eigenen Zugangsschlüssel.
- **Einrichtung:** Einmalig in der App „Mit Cockpit verbinden“. Das Cockpit zeigt dafür einen Kopplungscode, unter Technik → „App koppeln“.
- **Ablage:** Der Schlüssel liegt im iOS-Schlüsselbund und im gemeinsamen App-Speicher. So können App und Widget ihn nutzen, ohne dass er im Code steht.
- **Kontrolle:** Der Schlüssel lässt sich im Cockpit jederzeit widerrufen.

## 6. Aktualisierung

- **Grenzen:** iOS gibt Widgets nur ein begrenztes Budget an Aktualisierungen. Geplant ist etwa alle 30 Minuten, morgens zwischen 6 und 10 Uhr öfter.
- **Sofort nachladen:**
  - beim Öffnen der App
  - nach „Absenden“ auf der Eingabe-Seite, die App gibt dem Widget dann Bescheid
  - nach einem Knopfdruck im Widget
- **Später:** Eine Push-Mitteilung „Briefing fertig“ kann das Widget aktuell halten. Push braucht das bezahlte Apple-Entwicklerkonto (siehe 8).

## 7. Projekt im Repo

```
ios/
  project.yml              XcodeGen-Beschreibung (kein Klick-Projekt, gut versionierbar)
  App/                     SwiftUI-App: Web-Ansicht, Kopplung, Deep Links
  Widgets/                 WidgetKit: Klein, Mittel, Groß, Sperrbildschirm, Knöpfe (App Intents)
  Shared/                  Datenmodell zu /api/widget, Abruf, Zwischenspeicher
  README.md                Bauen und Installieren Schritt für Schritt
```

- **Wer was macht:** Ich schreibe Code und Projektdatei. Gebaut und aufs iPhone gebracht wird auf deinem Mac mit Xcode, weil das nur dort geht.
- **Ein Befehl:** `xcodegen && open CoachKira.xcodeproj` erzeugt und öffnet das Projekt. Dann wählst du in Xcode dein iPhone und klickst auf Start.

## 8. Voraussetzungen und Entscheidungen für dich

| Punkt | Optionen | Empfehlung |
|---|---|---|
| Xcode auf dem Mac | aus dem Mac App Store, ca. 15 GB | nötig |
| Apple-Konto zum Signieren | **kostenlos:** App läuft 7 Tage, dann neu installieren; keine Push-Mitteilungen · **Entwicklerkonto 99 €/Jahr:** TestFlight, läuft dauerhaft, Push möglich | für Alltagsnutzung das Entwicklerkonto |
| Geräte | iPhone · optional iPad und Mac (gleiche Widgets) | iPhone zuerst |
| Apple Watch | Sperrbildschirm-Widgets erscheinen im Smart Stack, eine eigene Watch-App wäre ein späterer Schritt | später |

## 9. Etappen

| Etappe | Inhalt | Fertig, wenn … |
|---|---|---|
| E1 Server | `/api/widget` mit derselben Logik wie das Cockpit, Kopplungscode, Abgleichtest | Werte von `/api/widget` = Heute-Seite an 3 Tagen hintereinander |
| E2 App-Hülle | SwiftUI-App mit Cockpit, Symbol, Kopplung, Deep Links | App startet vom Homescreen, Heute und Eingabe funktionieren, Absenden klappt |
| E3 Widgets | Klein, Mittel, Groß, Sperrbildschirm | Widgets zeigen den Morgenstand innerhalb von 30 min nach dem Absenden |
| E4 Knöpfe | „Garmin abrufen“, „Briefing lokal“ im Widget | Knopf → Auftrag erscheint im Cockpit → Mac erledigt ihn |
| E5 optional | Push „Briefing fertig“, iPad/Mac, Watch | nach Bedarf |

E1 kann ich sofort umsetzen und wie gewohnt erst lokal prüfen, dann veröffentlichen. Für E2 bis E4 schreibe ich den Code fertig, und du baust ihn mit Xcode auf deinem Mac. Bei Problemen in Xcode hilft ein Screenshot der Fehlermeldung.

## 10. Offene Fragen

1. **Konto:** Kostenloses Apple-Konto (7 Tage) oder Entwicklerkonto (99 €/Jahr)?
2. **Xcode:** Ist Xcode schon installiert? Ich konnte es von hier aus nicht prüfen.
3. **Inhalte:** Passen die Inhalte der drei Widget-Größen, oder fehlt dir ein Wert, zum Beispiel Body Battery, Ruhepuls-Trend oder die Ernährung heute?
4. **Symbol:** Hast du einen Wunsch, oder soll ich eins im Garmin-Stil entwerfen?
