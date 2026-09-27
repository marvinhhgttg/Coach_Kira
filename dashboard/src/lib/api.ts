// Coach Kira Apps Script API client.
// All requests are GET with a cache-buster `cb` parameter.
// Token is never persisted — it lives in React state only.

export const BASE_URL =
  'https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec';

function cb() {
  return Date.now().toString();
}

function buildUrl(params: Record<string, string | number | undefined>): string {
  const usp = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === '') continue;
    usp.set(k, String(v));
  }
  usp.set('cb', cb());
  return `${BASE_URL}?${usp.toString()}`;
}

export class ApiError extends Error {
  constructor(message: string, public readonly status?: number) {
    super(message);
    this.name = 'ApiError';
  }
}

export class TimeoutError extends ApiError {
  constructor(message = 'Zeitüberschreitung') {
    super(message);
    this.name = 'TimeoutError';
  }
}

/**
 * Wrap a promise with a timeout. If the wrapped call accepts an AbortSignal, pass
 * `signal` from `withTimeout(ms, (signal) => fetcher(signal))` so the underlying
 * fetch is actually aborted instead of just detached.
 */
export function withTimeout<T>(
  ms: number,
  factory: (signal: AbortSignal) => Promise<T>,
): Promise<T> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ms);
  return factory(controller.signal)
    .catch((err: any) => {
      if (controller.signal.aborted) throw new TimeoutError(`Timeout nach ${ms} ms`);
      throw err;
    })
    .finally(() => clearTimeout(timer));
}

/**
 * Lesende Anfragen (mode=…) werden bei HTTP 404/5xx oder Netzwerkfehler einmal wiederholt.
 * Hintergrund: Apps Script liefert bei ausgelastetem Sheet sporadisch eine 404-Seite
 * ("unable to open the file at this time"). Schreibende Aktionen (action=…) werden nie wiederholt.
 */
async function getJson<T>(url: string, signal?: AbortSignal): Promise<T> {
  const isRead = /[?&]mode=/.test(url) && !/[?&]action=/.test(url);
  try {
    return await getJsonOnce<T>(url, signal);
  } catch (err) {
    const retryable =
      isRead &&
      !signal?.aborted &&
      err instanceof ApiError &&
      (err.status == null || err.status === 404 || err.status >= 500);
    if (!retryable) throw err;
    await new Promise((r) => setTimeout(r, 1500));
    return getJsonOnce<T>(url.replace(/([?&]cb=)\d+/, `$1${Date.now()}`), signal);
  }
}

async function getJsonOnce<T>(url: string, signal?: AbortSignal): Promise<T> {
  let res: Response;
  try {
    res = await fetch(url, {
      method: 'GET',
      redirect: 'follow',
      mode: 'cors',
      cache: 'no-store',
      signal,
    });
  } catch (err: any) {
    throw new ApiError(`Netzwerkfehler: ${err?.message || 'unbekannt'}`);
  }
  if (!res.ok) {
    throw new ApiError(`HTTP ${res.status}`, res.status);
  }
  const text = await res.text();
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new ApiError('Antwort ist kein gültiges JSON');
  }
}

// ---------- Types ----------

export type Ampel = 'GRÜN' | 'GELB' | 'ORANGE' | 'ROT' | 'GRAU' | 'BLAU' | 'LILA' | string;

export interface StatusItem {
  kategorie?: string;
  metrik?: string;
  ampel?: Ampel;
  wert?: number | string | null;
  score?: number | null;
  text?: string;
}

export interface StatusResponse {
  ok: boolean;
  timestamp: string;
  heartbeat?: {
    timestamp?: string;
    stage?: string;
    runId?: string;
    extra?: string;
    header?: string[];
  };
  summary?: {
    gesamtAmpel?: Ampel;
    gesamtScore?: number;
    recoveryScore?: number;
    trainingScore?: number;
    planStatus?: string;
    items?: StatusItem[];
  };
}

export interface PlanDay {
  index?: number;
  date: string;
  day?: string;
  original_load_ess?: number | null;
  recommended_load_ess?: number | null;
  recommended_zone?: string;
  training_1?: string;
  training_2?: string;
  projected_effect?: string;
  weather_recommendation?: string;
  timeline_atl?: number | null;
  timeline_ctl?: number | null;
  timeline_acwr?: number | null;
  timeline_kei?: number | null;
  timeline_monotony?: number | null;
  timeline_hrv_status?: number | null;
  timeline_hrv_thresholds?: string | null;
  timeline_sg_flags?: string | null;
}

export interface PlanResponse {
  ok: boolean;
  timestamp: string;
  rows?: number;
  plan: PlanDay[];
}

export interface ChartDataResponse {
  ok: boolean;
  timestamp: string;
  range: string;
  requestedMetrics: string[];
  returnedMetrics: string[];
  rows: number;
  data: Array<Record<string, any>>;
}

export interface LogEntry {
  row?: number;
  timestamp: string;
  level: string;
  message: string;
  raw?: any[];
}

export interface LogsResponse {
  ok: boolean;
  rows: number;
  logs: LogEntry[];
}

export interface RunStatusResponse {
  ok: boolean;
  timestamp: string;
  requestedRunId: string | null;
  lastQueued?: {
    runId?: string;
    functionName?: string;
    queuedAt?: string;
  } | null;
  heartbeat?: {
    timestamp?: string;
    stage?: string;
    runId?: string;
    extra?: string;
  };
}

export interface SupervisorResponse {
  ok: boolean;
  timestamp?: string;
  queued?: boolean;
  functionName?: string;
  runId?: string;
  error?: string;
}

export interface PlanSimulationResponse {
  ok: boolean;
  timestamp?: string;
  source?: string;
  base?: {
    atl: number;
    ctl: number;
    ctlHistory?: number[];
    config?: Record<string, any>;
    todayIsClosed?: boolean;
    startDate?: string;
  };
  days?: Array<{
    index: number;
    date: string;
    day: string;
    phase: string;
    load: number;
    sport: string;
    zone: string;
    te_ae: number;
    te_an: number;
    locked: boolean;
    atl?: number;
    ctl?: number;
    acwr?: number;
    kei?: number;
    monotony?: number;
  }>;
  error?: string;
}

export interface SaveSimulationResponse {
  ok: boolean;
  timestamp?: string;
  message?: string;
  raw?: string;
  error?: string;
}

export interface StatusAiAnalysisResponse {
  ok: boolean;
  timestamp?: string;
  analysisTimestamp?: string | null;
  analysis?: string | null;
  error?: string;
}

export interface ConsolidatedAnalysisResponse {
  ok: boolean;
  timestamp?: string;
  analysisTimestamp?: string | null;
  analysis?: string | null;
  raw?: string | null;
  error?: string;
}

export interface PlanBriefingResponse {
  ok: boolean;
  timestamp?: string;
  briefingTimestamp?: string | null;
  briefing?: string | null;
  recommendations?: Array<{
    date: string;
    day?: string;
    recommended_load_ess: number;
    recommended_zone?: string;
    rationale?: string;
  }> | null;
  raw?: string | null;
  error?: string;
}

// ---------- Endpoints ----------

export function fetchStatus(signal?: AbortSignal) {
  return getJson<StatusResponse>(buildUrl({ mode: 'status' }), signal);
}

export function fetchRunStatus(signal?: AbortSignal) {
  return getJson<RunStatusResponse>(buildUrl({ mode: 'runStatus' }), signal);
}

export function fetchPlan(signal?: AbortSignal) {
  return getJson<PlanResponse>(buildUrl({ mode: 'planJsonV2' }), signal);
}

export function fetchLogs(limit = 15, signal?: AbortSignal) {
  return getJson<LogsResponse>(buildUrl({ mode: 'logs', limit }), signal);
}

export const LOAD_METRICS = [
  'load_fb_day',
  'fbATL_obs',
  'fbCTL_obs',
  'fbACWR_obs',
  'coachE_Smart_Gains',
  'garminEnduranceScore',
] as const;
export const RECOVERY_METRICS = [
  'sleep_hours',
  'sleep_score_0_100',
  'rhr_bpm',
  'Garmin_Training_Readiness',
] as const;

export const ACTIVITY_METRICS = [
  'load_fb_day',
  'Sport_x',
  'Zone',
  'Aerobic_TE',
  'Anaerobic_TE',
  'activity_done',
] as const;

export type Range = '7d' | '14d' | '28d' | '60d' | '90d' | '180d' | '360d';

export function fetchChartData(
  range: Range,
  metrics: readonly string[],
  signal?: AbortSignal
) {
  return getJson<ChartDataResponse>(
    buildUrl({ mode: 'chartData', range, metrics: metrics.join(',') }),
    signal
  );
}

// Action — token stays in caller scope
export function runSupervisor(token: string, signal?: AbortSignal) {
  if (!token) {
    return Promise.reject(new ApiError('Token fehlt'));
  }
  return getJson<SupervisorResponse>(
    buildUrl({ action: 'runSupervisor', token }),
    signal
  );
}

export function fetchPlanSimulation(signal?: AbortSignal) {
  return getJson<PlanSimulationResponse>(buildUrl({ mode: 'planSimulationJson' }), signal);
}

/** planSimulationJson with a hard timeout (Apps Script may 404 slowly). */
export function fetchPlanSimulationWithTimeout(ms = 8000) {
  return withTimeout(ms, (signal) => fetchPlanSimulation(signal));
}

export function savePlanSimulation(
  token: string,
  payload: {
    loads: number[];
    teAe: number[];
    teAn: number[];
    sports: string[];
    zones: string[];
    locks: boolean[];
  },
  signal?: AbortSignal
) {
  if (!token) {
    return Promise.reject(new ApiError('Token fehlt'));
  }
  return getJson<SaveSimulationResponse>(
    buildUrl({
      action: 'saveSimulatedPlan',
      token,
      loads: JSON.stringify(payload.loads),
      teAe: JSON.stringify(payload.teAe),
      teAn: JSON.stringify(payload.teAn),
      sports: JSON.stringify(payload.sports),
      zones: JSON.stringify(payload.zones),
      locks: JSON.stringify(payload.locks),
    }),
    signal
  );
}

export interface SaveWellbeingResponse {
  ok: boolean;
  timestamp?: string;
  date?: string;
  value?: number;
  written?: string[];
  error?: string;
}

export function saveWellbeing(token: string, date: string, value: number, signal?: AbortSignal) {
  if (!token) return Promise.reject(new ApiError('Token fehlt'));
  return getJson<SaveWellbeingResponse>(buildUrl({ action: 'saveWellbeing', token, date, value }), signal);
}

export type EntryType = 'Morgens' | 'Nach Aktivität' | 'Abends';

export interface SubmitEntryResponse {
  ok: boolean;
  timestamp?: string;
  type?: EntryType;
  date?: string;
  row?: number;
  written?: string[];
  readback?: Record<string, unknown>;
  triggersStarted?: boolean;
  durationMs?: number;
  error?: string;
}

export function submitCockpitEntry(token: string, type: EntryType, date: string, values: Record<string, string | number>) {
  if (!token) return Promise.reject(new ApiError('Token fehlt'));
  return getJson<SubmitEntryResponse>(
    buildUrl({ action: 'submitCockpitEntry', token, type, date, values: JSON.stringify(values) }),
  );
}

export function fetchStatusAiAnalysis(signal?: AbortSignal) {
  return getJson<StatusAiAnalysisResponse>(buildUrl({ mode: 'statusAiAnalysis' }), signal);
}

export function runStatusAiAnalysis(token: string, signal?: AbortSignal) {
  if (!token) {
    return Promise.reject(new ApiError('Token fehlt'));
  }
  return getJson<StatusAiAnalysisResponse>(
    buildUrl({ action: 'runStatusAiAnalysis', token }),
    signal
  );
}

export function fetchConsolidatedAnalysis(signal?: AbortSignal) {
  return getJson<ConsolidatedAnalysisResponse>(buildUrl({ mode: 'consolidatedAnalysis' }), signal);
}

export function runConsolidatedAnalysis(token: string, signal?: AbortSignal) {
  if (!token) {
    return Promise.reject(new ApiError('Token fehlt'));
  }
  return getJson<ConsolidatedAnalysisResponse>(
    buildUrl({ action: 'runConsolidatedAnalysis', token }),
    signal
  );
}

export function fetchPlanBriefing(signal?: AbortSignal) {
  return getJson<PlanBriefingResponse>(buildUrl({ mode: 'planBriefing' }), signal);
}

export function generatePlanBriefing(token: string, signal?: AbortSignal) {
  if (!token) {
    return Promise.reject(new ApiError('Token fehlt'));
  }
  return getJson<PlanBriefingResponse>(
    buildUrl({ action: 'generatePlanBriefing', token }),
    signal
  );
}

// ---------- Helpers ----------

export function dedupeByDateKeepLast<T extends { date: string }>(rows: T[]): T[] {
  const map = new Map<string, T>();
  rows.forEach((row) => map.set(row.date, row));
  return Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
}

export function toNum(v: unknown): number | null {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : Number(v);
  return Number.isFinite(n) ? n : null;
}
