/**
 * Etappe 5 (20.11.2026): geräteübergreifende Einstellungen zentral auf dem Server.
 * Gilt für iPhone, iPad, Mac-App und Browser gleich. Nur eine feste Liste von Schlüsseln mit Prüfung;
 * Anzeige-Vorlieben (Hell/Dunkel, offener Reiter) und Zwischenspeicher bleiben je Gerät.
 *   GET  /api/einstellungen   → { ok, werte, at }        (nur mit Sitzung, Regel in zugang.mjs)
 *   POST /api/einstellungen   { werte: { schluessel: wert } } → { ok, werte, at }
 */
export const PRUEFUNG = {
  ck_votum_erfahrung_v2: (v) => v === 'an' || v === 'aus', // Erfahrungsmodell ins Votum einrechnen
  ck_kira_engine_v2: (v) => v === 'local' || v === 'pplx', // KI: lokal (Mac) oder Cloud
  ck_es_datum: (v) => /^\d{4}-\d{2}-\d{2}$/.test(v), // Ausdauer-Ziel: Datum
  ck_es_ziel: (v) => /^\d{3,5}$/.test(v) && Number(v) >= 1000 && Number(v) <= 20000, // Ausdauer-Ziel: Score
  ck_aufstehen: (v) => /^([01]\d|2[0-3]):[0-5]\d$/.test(v), // Schlafkonto: Aufstehzeit
};

export function setupEinstellungen(app, { getSetting, setSetting }) {
  const lesen = () => {
    try { const j = JSON.parse(getSetting('geraete_einstellungen') || '{}'); return { werte: j.werte || {}, at: j.at || {} }; } catch { return { werte: {}, at: {} }; }
  };
  app.get('/api/einstellungen', (_req, res) => res.json({ ok: true, ...lesen() }));
  app.post('/api/einstellungen', (req, res) => {
    const neu = (req.body && req.body.werte) || {};
    const d = lesen();
    const abgelehnt = [];
    for (const [k, roh] of Object.entries(neu)) {
      const v = String(roh ?? '');
      if (!PRUEFUNG[k] || !PRUEFUNG[k](v)) { abgelehnt.push(k); continue; }
      if (d.werte[k] !== v) { d.werte[k] = v; d.at[k] = Date.now(); }
    }
    setSetting('geraete_einstellungen', JSON.stringify(d));
    res.json({ ok: true, ...d, abgelehnt });
  });
}
