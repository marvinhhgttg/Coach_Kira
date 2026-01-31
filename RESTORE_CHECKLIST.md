# Restore Checklist — GitHub → Google Apps Script (Coach_Kira)

Diese Checkliste beschreibt, wie du **bewusst** Code aus **GitHub** zurück in das **Google Apps Script** (Sheets-bound) Projekt spielst.

**Ziel:** Repo dient als Backup. Restore soll sicher sein (Rollback möglich, keine unabsichtlichen Überschreibungen).

> Hinweis: `git push/pull` betrifft nur Repo-Dateien. `clasp push/pull` synchronisiert Code-Dateien mit Apps Script – **nicht** Sheet-Daten, Trigger oder Deployments.

---

## 0) Vorher klären
- Welcher Stand soll zurückgespielt werden?
  - ein bestimmter Commit / Branch / Tag?
- Im richtigen Ordner?
  - `/Users/marcreichpietsch/Coach_Kira/Coach_Kira`

---

## 1) Apps Script Projekt absichern (Rollback)
Im Apps Script Editor:
- **Project Version** anlegen ("Save new version")

Optional:
- Sheet einmal neu laden, damit UI "clean" ist.

---

## 2) Lokalen Stand sauber machen

```bash
cd /Users/marcreichpietsch/Coach_Kira/Coach_Kira
git status
```

- Wenn lokale Änderungen vorhanden sind: committen oder staschen (sonst entstehen Mischzustände).

---

## 3) Aktuellen Stand aus Apps Script sichern (Backup vor Restore)

```bash
clasp pull

git add -A
git commit -m "Backup before restore from GitHub"
git push
```

Damit ist der aktuelle Ist-Stand des Projekts auf GitHub konserviert.

---

## 4) Ziel-Stand aus GitHub holen

```bash
git pull

# optional gezielt:
# git checkout <commit|tag|branch>
```

---

## 5) Killer-Check: Keys / Properties / Services

### A) Script Properties
- Prüfen, ob **`GEMINI_API_KEY`** in den Script/Project Properties existiert.
  - Im Code gibt es eine harte Abbruchstelle, wenn der Key fehlt.

### B) APIs / Services
- Falls Advanced Google Services genutzt werden (z. B. Calendar/Drive): prüfen, ob sie im Projekt noch aktiviert sind.
- `appsscript.json` wird mit übertragen, aber Aktivierung/Scopes können trotzdem Interaktion erfordern.

---

## 6) Restore: Code ins Apps Script pushen (bewusster Schritt)

```bash
clasp push
```

⚠️ Das überschreibt den Code im Apps Script Projekt mit deinem lokalen Stand.

---

## 7) Smoke Test (2 Minuten)
Im Apps Script Editor:
- Sheet neu laden → Menü **Coach Kira** erscheint? (`onOpen`)
- Eine harmlose Funktion ausführen (z. B. `debugDump()`)
- "Executions" prüfen: Fehler? Scope-Prompts?

---

## 8) Trigger / Deployments prüfen (falls genutzt)

### Trigger
- Installierte Trigger (time-based, onEdit-installiert) werden nicht via Git/clasp versioniert.
- Nach Restore prüfen und ggf. neu anlegen.

### Deployments / WebApp
- Wenn du eine feste URL verwendest (`FIXED_WEBAPP_URL` im Code), prüfen:
  - Ist diese URL noch der gewünschte Deployment-Stand?
  - Falls Deployments neu gemacht wurden: URL/ID kann veraltet sein.

---

## Häufigste Fehlerquellen in diesem Projekt
1) `GEMINI_API_KEY` fehlt/anders benannt → harter Fehler.
2) Unabsichtliches Überschreiben neuerer Browser-Änderungen → Code weg.
3) WebApp-Deployment/URL driftet → falscher Endpunkt.
