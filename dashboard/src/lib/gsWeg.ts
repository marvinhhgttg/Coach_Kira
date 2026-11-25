/**
 * Etappe 3b – Schreibwege ins Apps Script laufen über den Cockpit-Server (20.11.2026).
 * Der Browser schickt Schreibaufrufe (POST oder action=…) nicht mehr direkt an Apps Script,
 * sondern an /api/gs. Der Server prüft die PIN-Sitzung und hängt den Apps-Script-Token an –
 * der Token kommt nie in den Browser.
 * Restrisiko 3 (25.11.2026): auch Lesen (GET mode=…) läuft jetzt über /api/gs, damit Apps Script
 * Lesezugriffe ohne Token abweisen kann (Schalter CK_LESE_PFLICHT).
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
      if (methode === 'POST' || u.searchParams.has('action') || u.searchParams.has('mode')) {
        const query: Record<string, string> = {};
        u.searchParams.forEach((v, k) => { if (k !== 'token') query[k] = v; });
        const h = new Headers(init?.headers || {});
        const proxy = orig(`${SERVER_BASE}/api/gs`, {
          method: 'POST',
          credentials: 'same-origin',
          signal: init?.signal,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ methode, query, body: typeof init?.body === 'string' ? init.body : undefined, contentType: h.get('content-type') || undefined }),
        });
        const nurLesen = methode !== 'POST' && !u.searchParams.has('action');
        if (!nurLesen) return proxy;
        // Lesen: ohne Sitzung (401/403) oder ohne Server-Token (503) direkt wie bisher – wird in Apps Script gezählt
        return proxy.then((r) => (r.status === 401 || r.status === 403 || r.status === 503 ? orig(input as any, init) : r))
          .catch((e) => { if ((e as any)?.name === 'AbortError') throw e; return orig(input as any, init); });
      }
    }
    return orig(input as any, init);
  };
}
