# Garmin-Historie

Einmalige Auswertung aus Garmin Connect (21.11.2026, im angemeldeten Browser gelesen):

- Aktivitätsliste (`/gc-api/activitylist-service/activities/search/activities`): 4334 Aktivitäten seit 2013, Last (ESS) seit 2020.
- Endurance Score täglich (`/gc-api/metrics-service/metrics/endurancescore/stats?aggregation=daily`): erst ab 21.08.2025.

`ana.py` berechnet das Residuum von ES_MODEL (Tage 0–3) je Dauerklasse, `gen.py` schreibt `dashboard/src/lib/garminHistorie.ts`.
Rohdaten (`/tmp/garmin_hist.json`) liegen bewusst nicht im Repo.
