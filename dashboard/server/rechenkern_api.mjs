import { abgleich, simuliere } from './rechenkern.mjs';

// Etappe 1: eigene Apps-Script-Bereitstellung „v2 Schattenbetrieb“ (nur lesend).
// Entscheidung 20.11.2026: Das Formelblatt `timeline` ist die Wahrheit.
const V2_URL = process.env.CK_V2_SCRIPT_URL ||
  'https://script.google.com/macros/s/AKfycbwfLqLJ56WtQj6CdfufJW3VMclCGyF1XL6EUOBR--Jm3UOlqALIB0Y32wJ6xIZhFw93/exec';
const TZ = 'Europe/Berlin';
const heute = () => new Date().toLocaleDateString('sv-SE', { timeZone: TZ });

async function holen(params) {
  const usp = new URLSearchParams({ mode: 'rohdaten', ab: '2025-08-26', ...params, cb: String(Date.now()) });
  const r = await fetch(`${V2_URL}?${usp}`, { redirect: 'follow' });
  const t = await r.text();
  try { return JSON.parse(t); } catch { return { ok: false, error: `Apps Script antwortet nicht mit JSON (HTTP ${r.status})` }; }
}

// Zeitpunkt „datum hh:mm:ss“ in Berlin als ms (berücksichtigt Sommer-/Winterzeit)
function berlin(datum, h = 0, m = 0, s = 0) {
  const [y, mo, d] = datum.split('-').map(Number);
  const utc = Date.UTC(y, mo - 1, d, h, m, s);
  const off = (ms) => {
    const t = new Intl.DateTimeFormat('en-US', { timeZone: TZ, timeZoneName: 'shortOffset' }).formatToParts(new Date(ms)).find((p) => p.type === 'timeZoneName')?.value || 'GMT+0';
    const mm = t.match(/GMT([+-]\d+)(?::(\d+))?/);
    return mm ? (Number(mm[1]) * 60 + Math.sign(Number(mm[1])) * Number(mm[2] || 0)) * 60000 : 0;
  };
  return utc - off(utc - off(utc));
}

export function setupRechenkern(app, { getSetting, setSetting, aktiv = true } = {}) {
  // ---------- Rohdaten aus `timeline` (5 Minuten zwischengespeichert) ----------
  let tlCache = null;
  async function timeline(neu = false) {
    if (!neu && tlCache && Date.now() - tlCache.at < 5 * 60 * 1000) return tlCache.d;
    const d = await holen({ blatt: 'timeline', formeln: 'alle' });
    if (d.ok && d.blatt === 'timeline') tlCache = { at: Date.now(), d };
    return d;
  }

  // ---------- Abgleich ----------
  let cache = null;
  async function abgleichen(neu = false) {
    if (!neu && cache && Date.now() - cache.at < 10 * 60 * 1000) return cache.daten;
    const [tl, kk] = await Promise.all([timeline(neu), holen({})]);
    if (!tl.ok) return { ok: false, error: tl.error || 'Rohdaten fehlen' };
    if (tl.blatt !== 'timeline' || !tl.formelnAlle) return { ok: false, error: 'Patch B fehlt noch: die Route liefert kein Formelblatt timeline mit allen Formeln.' };
    const daten = { ...abgleich(tl, kk, heute()), stand: Date.now() };
    cache = { at: Date.now(), daten };
    return daten;
  }

  // ---------- Täglicher Abgleich + 14-Tage-Serie ----------
  const historie = () => { try { return JSON.parse(getSetting?.('rk_historie') || '[]'); } catch { return []; } };
  function serie(h) {
    let n = 0;
    for (let i = h.length - 1; i >= 0; i--) { if (h[i].alleOk) n++; else break; }
    return n;
  }
  async function taeglich() {
    const tag = heute();
    const h = historie();
    if (h.length && h[h.length - 1].datum === tag) return;
    const stunde = Number(new Date().toLocaleString('de-DE', { timeZone: TZ, hour: '2-digit', hour12: false }));
    if (stunde < 10) return; // nach dem Morgenformular
    const d = await abgleichen(true);
    if (!d.ok) return;
    const abweichend = d.spalten.filter((s) => s.treffer !== s.geprueft).map((s) => ({ spalte: s.spalte, n: s.geprueft - s.treffer, bsp: s.beispiele[0] || '' }));
    h.push({ datum: tag, alleOk: abweichend.length === 0, abweichend, zeit: Date.now() });
    setSetting?.('rk_historie', JSON.stringify(h.slice(-60)));
  }
  if (aktiv && getSetting && setSetting) {
    setInterval(() => taeglich().catch(() => {}), 30 * 60 * 1000);
    setTimeout(() => taeglich().catch(() => {}), 20000);
  }

  app.get('/api/rechenkern/abgleich', async (req, res) => {
    try {
      const d = await abgleichen(req.query.neu === '1');
      const h = historie();
      res.json({ ...d, historie: h.slice(-21), serie: serie(h), ziel: 14 });
    } catch (e) {
      res.json({ ok: false, error: String(e?.message || e).slice(0, 200) });
    }
  });

  // ---------- Etappe 1d: Plan durchrechnen (Was-wäre-wenn), nur lesend ----------
  // GET /api/rechenkern/simulation?ess={"2026-11-22":250}
  app.get('/api/rechenkern/simulation', async (req, res) => {
    try {
      let ess = {};
      if (req.query.ess) {
        const roh = JSON.parse(String(req.query.ess));
        for (const [k, v] of Object.entries(roh)) if (/^\d{4}-\d{2}-\d{2}$/.test(k) && Number.isFinite(Number(v))) ess[k] = Math.max(0, Math.min(2000, Number(v)));
      }
      const tl = await timeline();
      if (!tl.ok) return res.json({ ok: false, error: tl.error || 'timeline nicht verfügbar' });
      const S = tl.spalten;
      const rows = tl.zeilen.map((z) => Object.fromEntries(S.map((s, j) => [s, z[j]])));
      const tag = heute();
      const start = rows.findIndex((r) => String(r.date) > tag);
      if (start < 0) return res.json({ ok: false, error: 'Keine Planzeilen ab morgen.' });
      const basis = simuliere(rows, start);
      const plan = Object.keys(ess).length ? simuliere(rows, start, ess) : basis;
      res.json({ ok: true, ab: rows[start].date, geaendert: ess, basis, plan, quelle: 'timeline · Rechenkern' });
    } catch (e) {
      res.json({ ok: false, error: String(e?.message || e).slice(0, 200) });
    }
  });

  // ---------- Etappe 1b: chartData aus `timeline` (gleiches Format wie Apps Script mode=chartData) ----------
  app.get('/api/rechenkern/chartData', async (req, res) => {
    try {
      const rangeParam = String(req.query.range || '28d');
      const m = rangeParam.toLowerCase().match(/^(\d+)d$/);
      const days = m ? Math.max(1, Number(m[1])) : 28;
      const requested = String(req.query.metrics || 'load_fb_day,fbATL_obs,fbCTL_obs,fbACWR_obs,sleep_hours,sleep_score_0_100,rhr_bpm,Garmin_Training_Readiness')
        .split(',').map((s) => s.trim()).filter(Boolean);
      const tl = await timeline();
      if (!tl.ok) return res.json({ ok: false, error: tl.error || 'timeline nicht verfügbar' });
      const S = tl.spalten;
      const iDate = S.indexOf('date');
      const metr = requested.map((n) => ({ name: n, i: S.indexOf(n) })).filter((x) => x.i >= 0);
      const tag = heute();
      const bis = berlin(tag, 23, 59, 59);
      const von = bis - (days - 1) * 86400000;
      const data = [];
      for (const z of tl.zeilen) {
        const d = z[iDate];
        if (!d) continue;
        const t = berlin(d);
        if (t < von || t > bis) continue;
        const o = { row: z[0], date: d };
        for (const x of metr) o[x.name] = z[x.i] === null ? '' : z[x.i];
        data.push(o);
      }
      res.json({ ok: true, timestamp: new Date().toISOString(), range: rangeParam, requestedMetrics: requested, returnedMetrics: metr.map((x) => x.name), rows: data.length, data, quelle: 'timeline' });
    } catch (e) {
      res.json({ ok: false, error: String(e?.message || e).slice(0, 200) });
    }
  });
}
