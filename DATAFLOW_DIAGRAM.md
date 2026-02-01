# Coach_Kira — Dataflow Diagram

This doc summarizes the main runtime flows between Sheets, Apps Script, external APIs, and UI.

---

## 1) Core daily run (Supervisor)

```mermaid
flowchart TD
  A[Trigger / Menu / Manual Run] --> B[runKiraGeminiSupervisor]

  B --> C[Read timeline + KK_TIMELINE]
  C --> C1[Find today row via is_today]
  C --> C2[Compute scores / RTP / readiness]
  C --> C3[Load factors + week_config]

  C3 --> D[Build prompt/context]
  D --> E[LLM calls]
  E -->|OpenAI| E1[callOpenAI]
  E -->|Gemini| E2[callGeminiAPI]

  E --> F[Write outputs]
  F --> F1[AI_REPORT_STATUS]
  F --> F2[AI_REPORT_PLAN]
  F --> F3[AI_DATA_HISTORY / AI_DATA_FORECAST]
  F --> F4[LOOKER_DATA]
  F --> F5[AI_LOG / AI_HEARTBEAT]

  F --> G[schedulePostWork_]
  G --> H[time trigger ~15s]
  H --> I[runPostWork_]
  I --> J[Heavy exports]
  J --> F4
  I --> K[Cleanup trigger]
```

---

## 2) Two timeline sheets (source vs values)

```mermaid
flowchart LR
  T[timeline\n(formulas/source)] -->|copyTimelineData / delta sync| K[KK_TIMELINE\n(values/target)]
  K -->|fast reads| S[Supervisor / API endpoints]
  T -->|plan writes| P[Planning flows]\n(updatePlannedLoad etc.)
```

Key idea:
- Writes that should preserve formula propagation usually target `timeline`.
- Reads that must be fast and stable may prefer `KK_TIMELINE`.

---

## 3) WebApp / API endpoints

```mermaid
flowchart TD
  U[Client browser / widgets] -->|GET /exec?page=...| H[HTML templates]
  U -->|GET /exec?mode=json| J[Dashboard JSON]\n(cache + lock
  U -->|GET /exec?mode=timeline| L[Timeline JSON]\n(cache + lock

  H -->|calls back| J
  H -->|calls back| L
```

---

*Last updated: 2026-02-01*
