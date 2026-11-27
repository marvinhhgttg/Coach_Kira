/**
 * 21.11.2026: Kiras Wochenplan zentral auf dem Server – gleich auf iPhone, iPad, Mac-App und Browser.
 *   GET  /api/wochenplan   → { ok, plan, at }      (nur mit Sitzung, Regel in zugang.mjs)
 *   POST /api/wochenplan   { plan } → { ok, at }   (neuerer Plan gewinnt; ältere werden abgelehnt)
 * Gespeichert wird nur der Plan (Last, Zone, Sport, Training, Begründung, lange Einheit), keine Gesundheitswerte.
 */
const MAX_BYTES = 200 * 1024;

function pruefen(p) {
  if (!p || typeof p !== 'object' || !Array.isArray(p.tage) || p.tage.length === 0 || p.tage.length > 31) return false;
  if (typeof p.erstellt !== 'string' || Number.isNaN(Date.parse(p.erstellt))) return false;
  return p.tage.every((t) => t && typeof t.date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(t.date) && Number.isFinite(Number(t.load)));
}

export function setupWochenplan(app, { getSetting, setSetting }) {
  const lesen = () => {
    try { return JSON.parse(getSetting('wochenplan') || 'null'); } catch { return null; }
  };
  app.get('/api/wochenplan', (_req, res) => {
    const d = lesen();
    res.json({ ok: true, plan: d ? d.plan : null, at: d ? d.at : null });
  });
  app.post('/api/wochenplan', (req, res) => {
    const p = req.body && req.body.plan;
    if (!pruefen(p)) return res.status(400).json({ ok: false, error: 'Plan ungültig.' });
    const roh = JSON.stringify(p);
    if (roh.length > MAX_BYTES) return res.status(413).json({ ok: false, error: 'Plan zu groß.' });
    const alt = lesen();
    if (alt && alt.plan && Date.parse(alt.plan.erstellt) > Date.parse(p.erstellt)) return res.json({ ok: true, at: alt.at, neuer: true });
    const at = Date.now();
    setSetting('wochenplan', JSON.stringify({ plan: p, at }));
    res.json({ ok: true, at });
  });
}
