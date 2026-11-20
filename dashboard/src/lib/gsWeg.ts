/**
 * Etappe 3b – Schreibwege ins Apps Script laufen über den Cockpit-Server (20.11.2026).
 * Der Browser schickt Schreibaufrufe (POST oder action=…) nicht mehr direkt an Apps Script,
 * sondern an /api/gs. Der Server prüft die PIN-Sitzung und hängt den Apps-Script-Token an –
 * der Token kommt nie in den Browser. Lesen (GET mode=…) bleibt direkt.
 * Nicht in der Testinstanz v2 (dort sperrt v2guard alle Schreibwege).
 */
import { SERVER_BASE } from './cloudLLM';

const IST_V2 =
  typeof window !== 'undefined' &&
  (/^coachkira-v2\./.test(window.location.hostname) || (() => { try { return localStorage.getItem('ck-v2-test') === '1'; } catch { return false; } })());

if (typeof window !== 'undefined' && !IST_V2 && !(window as any).__ckGsWeg) {
  (window as any).__ckGsWeg = true;
  const orig = window.fetch.bind(window);
  window.fetch = (input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(url)) {
      const methode = String(init?.method || (typeof input === 'object' && 'method' in input ? input.method : 'GET')).toUpperCase();
      const u = new URL(url);
      if (methode === 'POST' || u.searchParams.has('action')) {
        const query: Record<string, string> = {};
        u.searchParams.forEach((v, k) => { if (k !== 'token') query[k] = v; });
        const h = new Headers(init?.headers || {});
        return orig(`${SERVER_BASE}/api/gs`, {
          method: 'POST',
          credentials: 'same-origin',
          signal: init?.signal,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ methode, query, body: typeof init?.body === 'string' ? init.body : undefined, contentType: h.get('content-type') || undefined }),
        });
      }
    }
    return orig(input as any, init);
  };
}
