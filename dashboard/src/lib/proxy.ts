import { ApiError, type EntryType, type SubmitEntryResponse, type SaveWellbeingResponse, type ConsolidatedAnalysisResponse, type PlanBriefingResponse, type SaveSimulationResponse, type StatusAiAnalysisResponse, type SupervisorResponse } from './api';

export interface ProxyAuthStatus {
  ok: boolean;
  configured: boolean;
  authenticated: boolean;
}

function apiUrl(path: string) {
  const normalizedPath = path.startsWith('/') ? path.slice(1) : path;
  const href = window.location.href.endsWith('/') ? window.location.href : `${window.location.href}/`;
  const base = new URL('api/', href).pathname.replace(/\/$/, '');
  return `${base}/${normalizedPath.replace(/^api\//, '')}`;
}

async function proxyJson<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, {
    ...init,
    headers: { 'content-type': 'application/json', ...(init?.headers || {}) },
    cache: 'no-store',
  });
  const text = await res.text();
  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    throw new ApiError('Proxy-Antwort ist kein JSON');
  }
  if (!res.ok) throw new ApiError(json?.error || `Proxy HTTP ${res.status}`, res.status);
  return json as T;
}

export function getProxyStatus() {
  return proxyJson<ProxyAuthStatus>(apiUrl('auth/status'));
}

export function setupProxy(token: string, pin: string) {
  return proxyJson<ProxyAuthStatus>(apiUrl('auth/setup'), {
    method: 'POST',
    body: JSON.stringify({ token, pin }),
  });
}

export function loginProxy(pin: string) {
  return proxyJson<ProxyAuthStatus>(apiUrl('auth/login'), {
    method: 'POST',
    body: JSON.stringify({ pin }),
  });
}

export function logoutProxy() {
  return proxyJson<{ ok: boolean }>(apiUrl('auth/logout'), { method: 'POST', body: '{}' });
}

export function proxyRunSupervisor() {
  return proxyJson<SupervisorResponse>(apiUrl('action/runSupervisor'), { method: 'POST', body: '{}' });
}

export function proxyRunStatusAiAnalysis() {
  return proxyJson<StatusAiAnalysisResponse>(apiUrl('action/runStatusAiAnalysis'), { method: 'POST', body: '{}' });
}

export function proxyRunConsolidatedAnalysis() {
  return proxyJson<ConsolidatedAnalysisResponse>(apiUrl('action/runConsolidatedAnalysis'), { method: 'POST', body: '{}' });
}

export function proxyGeneratePlanBriefing() {
  return proxyJson<PlanBriefingResponse>(apiUrl('action/generatePlanBriefing'), { method: 'POST', body: '{}' });
}

export function proxySavePlanSimulation(payload: {
  loads: number[];
  teAe: number[];
  teAn: number[];
  sports: string[];
  zones: string[];
  locks: boolean[];
}) {
  return proxyJson<SaveSimulationResponse>(apiUrl('action/saveSimulatedPlan'), {
    method: 'POST',
    body: JSON.stringify(payload),
  });
}

export function proxySaveWellbeing(date: string, value: number) {
  return proxyJson<SaveWellbeingResponse>(apiUrl('action/saveWellbeing'), {
    method: 'POST',
    body: JSON.stringify({ date, value }),
  });
}

export function proxySubmitEntry(type: EntryType, date: string, values: Record<string, string | number>) {
  return proxyJson<SubmitEntryResponse>(apiUrl('action/submitEntry'), {
    method: 'POST',
    body: JSON.stringify({ type, date, values }),
  });
}
