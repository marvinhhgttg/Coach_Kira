/**
 * Etappe 3b – Zugang (20.11.2026)
 *
 * Eine zentrale Prüfung vor allen API-Routen:
 *  - COCKPIT: nur aus dem Cockpit (exakte Herkunft) UND mit PIN-Sitzung
 *  - MAC:     Mac-Runner mit Mac-Schlüssel (Header x-ck-mac). Übergang: ohne Schlüssel erlaubt,
 *             solange CK_MAC_PFLICHT nicht "an" ist – jeder Aufruf ohne Schlüssel wird gezählt.
 *  - GEMISCHT: Sitzung ODER Mac
 *  - alles andere bleibt wie bisher (Widget mit Geräteschlüssel, Status-Abfragen, Anmeldung).
 *
 * Dazu:
 *  - POST /api/auth/token   Apps-Script-Token einmalig hinterlegen (Einrichtungscode)
 *  - POST|GET /api/gs       Schreibwege ins Apps Script über den Server (Token bleibt auf dem Server)
 *  - GET  /api/zugang/status Übersicht für das Cockpit (mit Sitzung)
 */
import crypto from 'node:crypto';

const sha = (s) => crypto.createHash('sha256').update(String(s)).digest('hex');
const gleich = (a, b) => {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

const REGELN = [
  // [Methode, Pfad, Klasse]
  ['POST', /^\/api\/garmin\/(anfrage|workout)\/status$/, 'mac'],
  // 23.11.2026: Garmin-Livewerte vom Mac (Bereitschaft, akute/chronische Last, Ausdauerwert) – nur Anzeige
  ['POST', /^\/api\/garmin\/live$/, 'mac'],
  ['GET', /^\/api\/garmin\/live$/, 'gemischt'],
  ['POST', /^\/api\/garmin\/(anfrage|workout)$/, 'cockpit'],
  ['GET', /^\/api\/garmin\/(anfrage|workout)$/, 'gemischt'],
  ['POST', /^\/api\/llm$/, 'gemischt'],
  ['POST', /^\/api\/widget\/(koppeln|widerrufen)$/, 'cockpit'],
  ['POST', /^\/api\/widget\/uhr\/bestaetigen$/, 'cockpit'],
  ['GET', /^\/api\/widget\/(geraete|abgleich)$/, 'cockpit'],
  ['POST', /^\/api\/texte\/(starten|einstellung|pruefen)$/, 'cockpit'],
  ['POST', /^\/api\/texte\/(mac|claim|auftrag\/status)$/, 'mac'],
  ['GET', /^\/api\/texte\/auftrag$/, 'mac'],
  ['POST', /^\/api\/lokal\/auftrag$/, 'cockpit'],
  ['GET', /^\/api\/lokal\/auftrag\/[^/]+$/, 'cockpit'],
  ['GET', /^\/api\/lokal\/naechster$/, 'mac'],
  ['POST', /^\/api\/lokal\/ergebnis$/, 'mac'],
  ['POST', /^\/api\/action\//, 'cockpit'],
  ['*', /^\/api\/gs$/, 'cockpit'],
  ['GET', /^\/api\/zugang\/status$/, 'cockpit'],
  ['*', /^\/api\/einstellungen$/, 'cockpit'],
  ['*', /^\/api\/wochenplan$/, 'cockpit'],
  // Etappe 6: der Formel-Abgleich kostet Rechenzeit → nur Sitzung oder Mac (chartData, votum, soll bleiben lesbar wie das Apps Script)
  ['GET', /^\/api\/rechenkern\/abgleich$/, 'gemischt'],
];

function klasse(method, pfad) {
  for (const [m, re, k] of REGELN) if ((m === '*' || m === method) && re.test(pfad)) return k;
  return null;
}

export function setupZugang(app, { getSetting, setSetting, isAuthed, browserOk, scriptUrl, aktiv = true }) {
  const lesen = (k, d) => { try { const v = getSetting(k); return v ? JSON.parse(v) : d; } catch { return d; } };
  const macHash = () => process.env.CK_MAC_HASH || getSetting('mac_hash') || '';
  const macPflicht = () => (process.env.CK_MAC_PFLICHT || 'aus') === 'an';

  function macOk(req) {
    const s = String(req.headers['x-ck-mac'] || '');
    const h = macHash();
    if (s && h && gleich(sha(s), h)) return 'schluessel';
    // Übergang: der Mac-Runner schickt bisher „Origin: https://coachkira.pplx.app“ mit – fremde Herkunft nie
    const o = String(req.headers.origin || '');
    if (o && o !== 'https://coachkira.pplx.app') return false;
    if (!macPflicht()) return 'uebergang';
    return false;
  }
  function ohneSchluesselMerken(req) {
    const z = lesen('mac_ohne', { n: 0, pfade: {} });
    z.n += 1;
    z.zuletzt = Date.now();
    z.pfade[req.method + ' ' + req.path] = (z.pfade[req.method + ' ' + req.path] || 0) + 1;
    setSetting('mac_ohne', JSON.stringify(z));
  }
  function mitSchluesselMerken() {
    const z = lesen('mac_mit', { n: 0 });
    z.n += 1;
    z.zuletzt = Date.now();
    setSetting('mac_mit', JSON.stringify(z));
  }

  // ---------- zentrale Prüfung ----------
  app.use((req, res, next) => {
    if (!aktiv || !req.path.startsWith('/api/')) return next();
    const k = klasse(req.method, req.path);
    if (!k) return next();
    const sitzung = isAuthed(req);
    if (k === 'cockpit') {
      if (!browserOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
      if (!sitzung) return res.status(401).json({ ok: false, error: 'not_authenticated' });
      return next();
    }
    if (k === 'gemischt' && sitzung && browserOk(req)) return next();
    const m = macOk(req);
    if (m === 'schluessel') { mitSchluesselMerken(); return next(); }
    if (m === 'uebergang') { ohneSchluesselMerken(req); return next(); }
    return res.status(k === 'gemischt' ? 401 : 403).json({ ok: false, error: k === 'gemischt' ? 'not_authenticated' : 'mac_key_required' });
  });

  // ---------- Fehlversuche (Einrichtungscode) ----------
  function codeOk(code) {
    const f = lesen('auth_fehl', { n: 0, bis: 0 });
    if (f.bis > Date.now()) return { ok: false, status: 429, error: 'Zu viele Fehlversuche – bitte 15 Minuten warten.' };
    const h = process.env.CK_SETUP_HASH || '';
    if (h && code && gleich(sha(String(code).trim()), h)) { setSetting('auth_fehl', JSON.stringify({ n: 0, bis: 0 })); return { ok: true }; }
    f.n += 1;
    if (f.n >= 5) { f.n = 0; f.bis = Date.now() + 15 * 60 * 1000; }
    setSetting('auth_fehl', JSON.stringify(f));
    return { ok: false, status: 403, error: h ? 'Einrichtungscode falsch.' : 'setup_disabled' };
  }

  // ---------- Apps-Script-Token einmalig hinterlegen ----------
  app.post('/api/auth/token', (req, res) => {
    const b = req.body || {};
    const vorhanden = !!getSetting('dashboard_token');
    if (vorhanden && !isAuthed(req)) return res.status(403).json({ ok: false, error: 'already_configured' });
    if (!vorhanden) {
      const c = codeOk(b.code);
      if (!c.ok) return res.status(c.status).json({ ok: false, error: c.error });
    }
    const t = String(b.token || '');
    if (t.length < 32) return res.status(400).json({ ok: false, error: 'token_too_short' });
    setSetting('dashboard_token', t);
    res.json({ ok: true });
  });

  // ---------- Schreibwege ins Apps Script über den Server ----------
  // Body: { methode: 'GET'|'POST', query: {…}, body?: string }  → Apps Script mit Token
  app.post('/api/gs', async (req, res) => {
    const token = getSetting('dashboard_token');
    if (!token) return res.status(503).json({ ok: false, error: 'token_not_configured' });
    const b = req.body || {};
    const q = new URLSearchParams();
    for (const [k, v] of Object.entries(b.query || {})) if (k !== 'token') q.set(k, String(v));
    q.set('token', token);
    const url = `${scriptUrl}?${q.toString()}`;
    try {
      const r = await fetch(url, b.methode === 'POST'
        ? { method: 'POST', redirect: 'follow', headers: { 'content-type': String(b.contentType || 'text/plain;charset=utf-8') }, body: typeof b.body === 'string' ? b.body : JSON.stringify(b.body ?? {}) }
        : { method: 'GET', redirect: 'follow' });
      const text = await r.text();
      res.status(200).type(r.headers.get('content-type') || 'application/json').send(text);
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e?.message || e).slice(0, 200) });
    }
  });

  app.get('/api/zugang/status', (req, res) => {
    res.json({
      ok: true,
      pin: !!getSetting('pin_hash'),
      appsScriptToken: !!getSetting('dashboard_token'),
      mac: { schluessel: !!macHash(), pflicht: macPflicht(), ohne: lesen('mac_ohne', { n: 0, pfade: {} }), mit: lesen('mac_mit', { n: 0 }) },
    });
  });

  return { macOk, codeOk };
}
