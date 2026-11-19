# Coach Kira – iPhone-App mit Widgets

Die App zeigt das Cockpit (coachkira.pplx.app). Dazu kommen Widgets für Homescreen und Sperrbildschirm. Die Werte liefert der Cockpit-Server über `/api/widget`, mit derselben Logik wie die Heute-Seite.

## Einmalig einrichten (ca. 10 Minuten)

1. **XcodeGen installieren** (erzeugt das Xcode-Projekt aus `project.yml`). Im Terminal:
   ```
   brew install xcodegen
   ```
2. **Projekt erzeugen und öffnen.** Im Terminal:
   ```
   cd ~/coach-kira-ios
   xcodegen
   open CoachKira.xcodeproj
   ```
3. **Apple-ID in Xcode anmelden:** Xcode → Einstellungen (⌘,) → Accounts → „+“ → Apple-ID. Die kostenlose Apple-ID reicht.
4. **Signieren:** Links das blaue Projekt „CoachKira“ anklicken. Dann für **beide** Targets, also „CoachKira“ und „CoachKiraWidgets“:
   - Reiter „Signing & Capabilities“ öffnen.
   - „Automatically manage signing“ muss ein Häkchen haben.
   - Unter Team „Marc Reichpietsch (Personal Team)“ wählen.

   Meldet Xcode, dass die Bundle-ID schon vergeben ist: in beiden Targets `de.reichpietsch.coachkira` durch etwas Eigenes ersetzen, etwa `de.reichpietsch.coachkira2`, und beim Widget `….widgets` behalten. Die App Group heißt dann ebenfalls neu. Xcode passt sie an, wenn du unter „App Groups“ auf „+“ gehst.
5. **iPhone vorbereiten:**
   - iPhone per Kabel an den Mac anschließen und „Diesem Computer vertrauen“ bestätigen.
   - Auf dem iPhone den Entwicklermodus einschalten: Einstellungen → Datenschutz & Sicherheit → Entwicklermodus → an. Das iPhone startet dabei neu.
6. **Starten:** In Xcode oben in der Leiste dein iPhone als Ziel wählen und auf ▶ (Run) klicken. Der erste Bau dauert 1 bis 2 Minuten.
7. **Beim ersten Start:** iOS fragt, ob du dem Entwickler vertraust. Einstellungen → Allgemein → VPN und Geräteverwaltung → deine Apple-ID → „Vertrauen“. Danach die App erneut öffnen.

## App koppeln

1. Im Cockpit, am Mac oder in der App: Heute → ganz unten „Technik“ → Kachel „iPhone-App & Widgets“ → „Kopplungscode erzeugen“.
2. Die App zeigt beim ersten Start „Mit Cockpit verbinden“. Den 6-stelligen Code eingeben und „Verbinden“ tippen. Später erreichst du die Kopplung über den runden Knopf unten rechts.

## Widgets hinzufügen

- **Homescreen:** Den Homescreen lange drücken → „Bearbeiten“ → „Widget hinzufügen“ → „Coach Kira“ suchen. Es gibt drei Größen: Klein, Mittel, Groß. Das große Widget hat die Knöpfe „Garmin abrufen“ und „Analyse lokal“.
- **Sperrbildschirm:** Den Sperrbildschirm lange drücken → „Anpassen“ → Sperrbildschirm → Widget-Leiste → „Coach Kira“. Es gibt einen Ring für die Readiness, eine Zeile und den Text über der Uhrzeit.

## Kostenloses Konto: alle 7 Tage neu signieren

Mit der kostenlosen Apple-ID läuft die App 7 Tage. Danach öffnet sie sich nicht mehr. Dann iPhone anschließen, Xcode öffnen und erneut ▶ klicken. Kopplung und Widgets bleiben erhalten. Mit dem Entwicklerkonto (99 € im Jahr) entfällt das.

## Wenn etwas nicht klappt

- **„Personal development teams do not support the App Groups capability“:** Laut Apple sollten App Groups auch mit der kostenlosen Apple-ID gehen. Kommt diese Meldung trotzdem, bitte einen Screenshot schicken. Dann stelle ich die Weitergabe des Schlüssels zwischen App und Widget anders um.
- **Rote Fehlermeldungen beim Bau:** Fehlerliste (⌘5) öffnen und einen Screenshot schicken.
- **Widget zeigt „App ist noch nicht gekoppelt“:** App öffnen → runder Knopf unten rechts → Kopplungscode eingeben.
- **Widget aktualisiert sich langsam:** iOS teilt Widgets ein Budget zu. Geplant ist morgens alle 15 Minuten und sonst alle 30 Minuten. Sofort frisch werden die Widgets beim Öffnen der App, nach „Absenden“ auf der Eingabe-Seite in der App und über „Widgets jetzt aktualisieren“ in der Kopplung.

## Mac-App mit Widgets

Gleiches Xcode-Projekt, zwei neue Targets: `CoachKiraMac` (App mit Fenster) und `CoachKiraMacWidgets`.

1. Neue Dateien liegen in `~/coach-kira-ios` (Ordner `MacApp/`, geänderte `project.yml`, `Widgets/KiraWidgets.swift`, `Shared/KiraDaten.swift`, `App/CoachKiraApp.swift`).
2. Xcode schließen, im Terminal neu erzeugen:
   ```
   cd ~/coach-kira-ios && xcodegen generate && open CoachKira.xcodeproj
   ```
3. Für alle vier Targets unter „Signing & Capabilities“ dein Team wählen (Personal Team).
4. Oben das Schema `CoachKiraMac` und als Ziel „My Mac“ wählen, dann ▶.
5. In der App: Toolbar „Koppeln“ (oder ⌘,) → Code aus dem Cockpit (Technik → App-Kopplung) eingeben. Der Mac wird als eigenes Gerät gekoppelt.

Bedienung im Fenster:
- ⌘1 bis ⌘7 springen zu Heute, Eingabe, Plan, Ziele, Analyse, Berichte, Log.
- ⌘R lädt neu, ⇧⌘R aktualisiert die Widgets, „Im Browser öffnen“ im Menü „Cockpit“.
- Das lokale Modell läuft direkt über Caddy (https://localhost:11435), ohne Umweg über das Relay.

Widgets auf dem Mac: Rechtsklick auf den Schreibtisch → „Widgets bearbeiten …“ (oder Mitteilungszentrale → „Widgets bearbeiten“), nach „Coach Kira“ suchen. Auf dem Mac gibt es nur die Größen klein, mittel und groß; die Sperrbildschirm-Varianten sind iPhone-only.

Hinweise:
- Auch die Mac-App muss mit dem kostenlosen Konto alle 7 Tage neu aus Xcode gestartet werden.
- Wenn Xcode beim App Group meckert: beim Mac-Target ist die Gruppe `<TeamID>.de.reichpietsch.coachkira` (macOS-Schreibweise) – Xcode legt sie beim Signieren selbst an.
- Bei Buildfehlern: Fehlerliste (⌘5) als Screenshot schicken.
