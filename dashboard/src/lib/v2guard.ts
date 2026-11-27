/**
 * Zweitinstanz coachkira-v2: läuft vor allem anderen.
 * 1. Nur lesend: Alles, was direkt vom Browser ins Apps Script schreiben würde (POST oder action=…),
 *    wird hier abgefangen. Schreiben geht nur im Live-Cockpit coachkira.pplx.app.
 * 2. Etappe 1b: Verlaufsdaten (mode=chartData) kommen aus dem Formelblatt `timeline` – der Wahrheit –
 *    über den v2-Server statt aus der Wertekopie KK_TIMELINE.
 * Lokal testen: localStorage.setItem('ck-v2-test', '1').
 */
import { SERVER_BASE } from './cloudLLM';

export const IST_V2 =
  typeof window !== 'undefined' &&
  (/^coachkira-v2\./.test(window.location.hostname) || (() => { try { return localStorage.getItem('ck-v2-test') === '1'; } catch { return false; } })());

if (IST_V2) {
  const original = window.fetch.bind(window);
  const antwort = (obj: unknown) =>
    new Response(JSON.stringify(obj), { status: 200, headers: { 'content-type': 'application/json' } });

  window.fetch = (input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    const methode = String(init?.method || (typeof input === 'object' && 'method' in input ? input.method : 'GET')).toUpperCase();
    if (/script\.google\.com\/macros\//.test(url)) {
      const u = new URL(url);
      if (methode !== 'GET' || u.searchParams.has('action')) {
        return Promise.resolve(antwort({ ok: false, error: 'v2-Testinstanz ist nur lesend. Speichern bitte im Cockpit auf coachkira.pplx.app.' }));
      }
      if (u.searchParams.get('mode') === 'chartData') {
        const q = new URLSearchParams();
        for (const k of ['range', 'metrics']) { const v = u.searchParams.get(k); if (v) q.set(k, v); }
        return original(`${SERVER_BASE}/api/rechenkern/chartData?${q}`, { ...init, mode: 'same-origin' });
      }
    }
    return original(input as RequestInfo, init);
  };
}
