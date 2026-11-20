import { abgleich } from './rechenkern.mjs';

// Etappe 1: Abgleich über die eigene Apps-Script-Bereitstellung „v2 Schattenbetrieb“ (nur lesend)
const V2_URL = process.env.CK_V2_SCRIPT_URL ||
  'https://script.google.com/macros/s/AKfycbwfLqLJ56WtQj6CdfufJW3VMclCGyF1XL6EUOBR--Jm3UOlqALIB0Y32wJ6xIZhFw93/exec';

async function holen(params) {
  const usp = new URLSearchParams({ mode: 'rohdaten', ab: '2025-08-26', ...params, cb: String(Date.now()) });
  const r = await fetch(`${V2_URL}?${usp}`, { redirect: 'follow' });
  const t = await r.text();
  try { return JSON.parse(t); } catch { return { ok: false, error: `Apps Script antwortet nicht mit JSON (HTTP ${r.status})` }; }
}

export function setupRechenkern(app) {
  let cache = null;
  app.get('/api/rechenkern/abgleich', async (req, res) => {
    try {
      if (cache && Date.now() - cache.at < 10 * 60 * 1000 && req.query.neu !== '1') return res.json(cache.daten);
      const [tl, kk] = await Promise.all([holen({ blatt: 'timeline', formeln: 'alle' }), holen({})]);
      if (!tl.ok) return res.json({ ok: false, error: tl.error || 'Rohdaten fehlen' });
      if (tl.blatt !== 'timeline' || !tl.formelnAlle) return res.json({ ok: false, error: 'Patch B fehlt noch: die Route liefert kein Formelblatt timeline mit allen Formeln.' });
      const heute = new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
      const daten = { ...abgleich(tl, kk, heute), stand: Date.now() };
      cache = { at: Date.now(), daten };
      res.json(daten);
    } catch (e) {
      res.json({ ok: false, error: String(e?.message || e).slice(0, 200) });
    }
  });
}
