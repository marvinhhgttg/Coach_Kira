# Coach_Kira — Apps Script ↔ GitHub (Backup-Workflow)

Dieses Repository dient **vorerst als Sicherheitskopie** (Backup) für das **Google Sheets–gebundene** Google Apps Script Projekt.

- Entwicklung/Editing passiert **im Google Apps Script Editor** (Browser).
- Versionierung/Backup passiert über **GitHub**.
- Der Sync zwischen Editor und Repo läuft über **clasp**.

> Hinweis: Triggers/Deployments/Berechtigungen sind nicht vollständig „Code-only“ reproduzierbar.

---

## Voraussetzungen (einmalig)

1) **Node.js + npm** installiert
2) `clasp` installieren:

```bash
npm i -g @google/clasp
```

3) **Apps Script API aktivieren**

- Öffne: https://script.google.com/home/usersettings
- Aktiviere: **Apps Script API**

4) `clasp` mit Google verbinden:

```bash
clasp login
```

---

## Projekt lokal klonen (einmalig)

Im Apps Script Editor:
- **Project Settings** → **Script ID** kopieren

Dann lokal:

```bash
cd /Users/marcreichpietsch/Coach_Kira
clasp clone <SCRIPT_ID>
```

Das erzeugt u. a.:
- `.clasp.json` (enthält die Script ID)
- `appsscript.json` (Manifest)
- `*.js` / `*.gs` und `*.html` Dateien

---

## Backup-Workflow (Browser-Editor → GitHub)

Du änderst Code im Browser. Wenn du sichern willst:

```bash
cd /Users/marcreichpietsch/Coach_Kira/Coach_Kira
clasp pull

git status

git add -A
git commit -m "Backup: <kurze Beschreibung>"

git push
```

**Empfehlung:** Lieber häufig kleine Backups als selten große.

---

## (Optional) Repo → Browser-Editor (nur bewusst!)

Wenn du **absichtlich** Änderungen aus dem Repo ins Apps Script Projekt übertragen willst:

```bash
cd /Users/marcreichpietsch/Coach_Kira/Coach_Kira

git pull
clasp push
```

⚠️ Achtung:
- Das überschreibt den Code im Apps Script Projekt.
- Wenn du gerade im Browser ungespeicherte Änderungen hast: erst speichern.

---

## Häufige Stolpersteine

### 1) „clasp pull“ überschreibt lokale Änderungen
Wenn du lokal noch uncommitted Änderungen hast:

```bash
git status
```

Entweder committen oder stashen, bevor du `clasp pull` machst.

### 2) Trigger / Properties
- Trigger (Zeit, onEdit, etc.) müssen oft **manuell** im Apps Script UI gesetzt werden.
- Secrets gehören **nicht** ins Repo. Nutze `PropertiesService` und dokumentiere im README, welche Keys erwartet werden.

### 3) `.clasp.json`
Enthält i. d. R. nur Script ID / RootDir.
Wenn du das Repo öffentlich machst, prüfen ob darin nichts Sensibles steht (normalerweise ok).

---

## Live beim Coden helfen

Sag mir einfach:
- Was soll das Script im Sheet tun?
- Welche Datei/Funktion ist betroffen?
- Und kopiere den relevanten Code-Abschnitt (oder den Fehler-Stacktrace) hier rein.

Dann kann ich:
- Bugs finden/fixen
- Refactorings vorschlagen
- Performance/Quota-Probleme entschärfen
- UI (Sidebar/Dialog) mit `.html` verbessern
