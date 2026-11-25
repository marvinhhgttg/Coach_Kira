/**
 * Restrisiko 3 (25.11.2026): Lesezugriffe aufs Apps Script tragen den Token.
 * Hängt bei allen fetch-Aufrufen an …/macros/s/…/exec den Apps-Script-Token an (falls noch keiner dabei ist).
 * Token-Quelle: globalThis.__ckGsToken() (Server, aus den Einstellungen) oder process.env.CK_GS_TOKEN (Rechner-Kindprozesse).
 * Ohne Token bleibt alles wie bisher.
 */
const RE = /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/;
if (!globalThis.__ckGsFetch && typeof globalThis.fetch === 'function') {
  globalThis.__ckGsFetch = true;
  const orig = globalThis.fetch;
  globalThis.fetch = (input, init) => {
    try {
      const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input && input.url;
      if (url && RE.test(url)) {
        const t = typeof globalThis.__ckGsToken === 'function' ? globalThis.__ckGsToken() : (process.env.CK_GS_TOKEN || '');
        if (t) {
          const u = new URL(url);
          if (!u.searchParams.has('token')) { u.searchParams.set('token', t); return orig(u.toString(), init); }
        }
      }
    } catch { /* unverändert weiter */ }
    return orig(input, init);
  };
}
export {};
