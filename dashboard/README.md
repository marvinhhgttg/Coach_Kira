# Coach Kira Cockpit (React/Vite)

Produktives Dashboard, live unter https://coachkira.pplx.app.
Datenquelle: Google Apps Script Web-App (`src/lib/api.ts`, `server.mjs`) auf Basis von `KK_TIMELINE`.

## Aufbau
- `src/screens/CommandCenter.tsx` – Status, Recovery · Tagesentscheid (`RecoveryPanel.tsx`)
- `src/screens/PlanCockpit.tsx` – 14-Tage-Simulation (ATL/CTL rekursiv ab Tag 2, SG_Flags)
- `src/screens/ChartDeck.tsx` – Charts mit Vollbild, RHR-Trend, KEI-Achse ab −50
- `server.mjs` – privater Proxy (PIN-Login, Token serverseitig) für Schreib-Aktionen

## Lokal
```bash
npm install
npm run build        # erzeugt dist/
node server.mjs      # Proxy + statische Dateien auf Port 5000
```

## Apps-Script-Abhängigkeiten
Das Dashboard erwartet diese Routen in `doGet`: `status`, `chartData`, `planSimulationJson`,
`planBriefing`, `planJsonV2`, `saveSimulatedPlan`, `saveWellbeing` u. a.
Die zuletzt eingespielten Patches liegen in `docs/`.

Nicht im Repo: `screenshots/` (persönliche Gesundheitsdaten), `deploy/` (Hetzner), `data.db` (PIN/Sessions).
