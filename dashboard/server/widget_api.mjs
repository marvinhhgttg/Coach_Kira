/**
 * Schnittstelle für die Coach-Kira-App und ihre Widgets.
 * - /api/widget            Tagesstand (Votum, Morgenwerte, Last, Ziel, Kira-Texte) – nur mit App-Schlüssel
 * - /api/widget/koppeln    Cockpit erzeugt einen 6-stelligen Kopplungscode (10 min gültig)
 * - /api/widget/verbinden  App tauscht den Code gegen einen eigenen Schlüssel
 * - /api/widget/geraete    gekoppelte Geräte anzeigen · /api/widget/widerrufen  Schlüssel löschen
 * - /api/widget/aktion     Knöpfe im Widget: Garmin abrufen, Briefing lokal (nie Cloud)
 * Die Werte rechnet server/widget.mjs mit denselben Funktionen wie das Cockpit.
 */
import { spawn } from 'node:child_process';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const KERN = path.join(__dirname, 'widget.mjs');
const CACHE_MS = 5 * 60 * 1000;
const CODE_MS = 10 * 60 * 1000;

const sha = (s) => crypto.createHash('sha256').update(String(s)).digest('hex');
const heute = () => new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });

export function setupWidget(app, { getSetting, setSetting, originOk, texte, appsScript }) {
  const lesen = (k, def) => {
    try {
      const v = getSetting(k);
      return v ? JSON.parse(v) : def;
    } catch {
      return def;
    }
  };
  const schreiben = (k, v) => setSetting(k, JSON.stringify(v));
  const geraete = () => lesen('widget_geraete', []);

  function geraetZu(req) {
    const h = String(req.headers.authorization || '');
    const tok = h.startsWith('Bearer ') ? h.slice(7).trim() : '';
    if (!tok) return null;
    const xs = geraete();
    const g = xs.find((x) => x.hash === sha(tok));
    if (!g) return null;
    g.zuletzt = Date.now();
    schreiben('widget_geraete', xs);
    return g;
  }

  // ---------- Rechnen mit Zwischenspeicher ----------
  let cache = lesen('widget_cache', null);
  let laufend = null;
  function rechnen() {
    if (laufend) return laufend;
    laufend = new Promise((resolve, reject) => {
      const p = spawn(process.execPath, [KERN], { stdio: ['ignore', 'pipe', 'pipe'] });
      let out = '';
      let err = '';
      const to = setTimeout(() => p.kill('SIGKILL'), 120000);
      p.stdout.on('data', (d) => (out += d));
      p.stderr.on('data', (d) => (err += d));
      p.on('close', (code) => {
        clearTimeout(to);
        laufend = null;
        if (code !== 0) return reject(new Error('Widget-Rechnung: ' + (err || 'Abbruch').slice(0, 200)));
        try {
          const j = JSON.parse(out);
          cache = { at: Date.now(), daten: j };
          schreiben('widget_cache', cache);
          resolve(cache);
        } catch {
          reject(new Error('Widget-Rechnung: keine gültige Antwort'));
        }
      });
    });
    return laufend;
  }
  function texteKurz() {
    const s = texte.statusAntwort();
    const its = s.items.filter((i) => i.datum >= heute() || i.art === 'ernaehrung');
    const fertig = its.filter((i) => i.phase === 'fertig');
    return {
      fertig: fertig.length,
      gesamt: its.length,
      cloud: fertig.filter((i) => i.engine === 'cloud').length,
      offen: its.filter((i) => i.phase !== 'fertig').map((i) => ({ label: i.label, phase: i.phase, versuche: i.versuche })),
      mac: { erreichbar: s.mac.erreichbar, ollama: s.mac.ollama },
    };
  }
  async function antwort(frisch = false) {
    const alt = !cache || cache.daten?.datum !== heute() || Date.now() - cache.at > CACHE_MS;
    if (frisch || alt) {
      try {
        await rechnen();
      } catch (e) {
        if (!cache) throw e;
        return { ...cache.daten, texte: texteKurz(), cacheAlter: Math.round((Date.now() - cache.at) / 1000), fehler: String(e.message || e) };
      }
    }
    return { ...cache.daten, texte: texteKurz(), cacheAlter: Math.round((Date.now() - cache.at) / 1000) };
  }

  app.get('/api/widget', async (req, res) => {
    if (!geraetZu(req)) return res.status(401).json({ ok: false, error: 'App nicht gekoppelt.' });
    try {
      const a = await antwort(req.query.frisch === '1');
      if (req.query.uhr === '1') {
        // Garmin-Uhr: wenig Speicher im Hintergrund – nur was die Uhr anzeigt
        const { texte: _t, uhr: u, ...rest } = a;
        // lange Einheiten aus Kiras Wochenplan (Server) an die nächsten Tage hängen
        let wp = null;
        try { wp = JSON.parse(getSetting('wochenplan') || 'null'); } catch { wp = null; }
        const lang = new Map(((wp && wp.plan && wp.plan.tage) || []).filter((t) => t && t.langH != null).map((t) => [t.date, t.langH]));
        const naechste = (a.naechste || []).map((x) => (lang.has(x.d) ? { ...x, langH: lang.get(x.d) } : x));
        // 22.11.2026: Diagramme nur im Vordergrund (?mehr=1) – der Hintergrunddienst bleibt unter 8 KB
        const mehr = req.query.mehr === '1' && u ? { uhr: { l14: u.l14, m7: u.m7, form: u.form, letzte: u.letzte } } : {};
        return res.json({ ...rest, naechste: req.query.mehr === '1' ? naechste : naechste.slice(0, 4), votum: a.votum ? { ...a.votum, gruende: (a.votum.gruende || []).slice(0, 2) } : null, ...mehr });
      }
      const { uhr: _u, ...ohneUhr } = a;
      res.json(ohneUhr);
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });
  // 22.11.2026: Datenfeld „Kira ESS“ – Rahmen des Tages und ESS je Stunde in Zone 1–5 je Sportgruppe
  app.get('/api/widget/feld', async (req, res) => {
    if (!geraetZu(req)) return res.status(401).json({ ok: false, error: 'nicht gekoppelt' });
    try {
      const a = await antwort(false);
      const f = a.uhr?.feld;
      if (!f) return res.status(503).json({ ok: false, error: 'noch keine Daten' });
      res.json({ ok: true, ...f });
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });
  // 22.11.2026: Befinden 1–5 von der Uhr (gleicher Weg wie im Cockpit: Apps Script saveWellbeing)
  app.post('/api/widget/befinden', async (req, res) => {
    if (!geraetZu(req)) return res.status(401).json({ ok: false, error: 'nicht gekoppelt' });
    const v = Number(req.body?.value);
    const datum = /^\d{4}-\d{2}-\d{2}$/.test(String(req.body?.date || '')) ? String(req.body.date) : heute();
    if (!Number.isInteger(v) || v < 1 || v > 5) return res.status(400).json({ ok: false, error: 'Wert 1–5' });
    if (datum > heute() || datum < new Date(Date.now() - 2 * 86400000).toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' })) return res.status(400).json({ ok: false, error: 'nur heute oder gestern' });
    if (!appsScript) return res.status(503).json({ ok: false, error: 'nicht verfügbar' });
    try {
      const r = await appsScript({ action: 'saveWellbeing', date: datum, value: String(v) });
      if (!r || !r.ok) return res.status(502).json({ ok: false, error: r?.error || 'Speichern fehlgeschlagen' });
      cache = null; // Votum neu rechnen
      res.json({ ok: true, date: datum, value: v });
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });
  // Für den Abgleichtest im Cockpit (gleiche Herkunft, ohne Schlüssel): nur Kennzahlen, keine Texte
  app.get('/api/widget/abgleich', async (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    try {
      const a = await antwort(false);
      res.json({ ok: true, datum: a.datum, vote: a.votum?.vote ?? null, readiness: a.morgen?.readiness ?? null, sollHeute: a.last?.sollHeute ?? null, sollText: a.last?.sollText ?? null, planHeute: a.last?.planHeute ?? null, acwr: a.last?.acwr ?? null, tourform: a.ziel?.tourform ?? null });
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });

  // ---------- Kopplung ----------
  app.post('/api/widget/koppeln', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const code = String(crypto.randomInt(0, 1000000)).padStart(6, '0');
    schreiben('widget_code', { hash: sha(code), bis: Date.now() + CODE_MS, fehl: 0 });
    res.json({ ok: true, code, gueltigBis: Date.now() + CODE_MS });
  });
  app.post('/api/widget/verbinden', (req, res) => {
    const b = req.body || {};
    const c = lesen('widget_code', null);
    if (!c || Date.now() > c.bis || c.fehl >= 5) return res.status(400).json({ ok: false, error: 'Kein gültiger Kopplungscode – im Cockpit neu erzeugen.' });
    if (sha(String(b.code || '').trim()) !== c.hash) {
      c.fehl += 1;
      schreiben('widget_code', c);
      return res.status(400).json({ ok: false, error: 'Code falsch.' });
    }
    schreiben('widget_code', null);
    const token = crypto.randomBytes(32).toString('hex');
    const xs = geraete();
    xs.push({ id: crypto.randomUUID(), hash: sha(token), name: String(b.geraet || 'iPhone').slice(0, 60), at: Date.now(), zuletzt: null });
    schreiben('widget_geraete', xs.slice(-10));
    res.json({ ok: true, token });
  });
  // ---------- Kopplung der Garmin-Uhr (21.11.2026) ----------
  // Die Uhr hat keine Tastatur: Sie fordert einen Code an und zeigt ihn; Marc bestätigt ihn im Cockpit (Sitzung).
  //   POST /api/widget/uhr/anfrage {geraet} → {id, code}      (offen, höchstens 3 gleichzeitig, 10 min)
  //   GET  /api/widget/uhr/status?id=…      → {wartet} | {token} (Token nur einmal)
  //   POST /api/widget/uhr/bestaetigen {code}                   (nur Cockpit)
  const anfragen = () => lesen('widget_uhr_anfragen', []).filter((a) => a.bis > Date.now());
  app.post('/api/widget/uhr/anfrage', (req, res) => {
    const xs = anfragen().filter((a) => !a.token);
    if (xs.length >= 3) return res.status(429).json({ ok: false, error: 'Zu viele offene Anfragen – in 10 min erneut.' });
    const code = String(crypto.randomInt(0, 1000000)).padStart(6, '0');
    const id = crypto.randomBytes(16).toString('hex');
    const alle = anfragen();
    alle.push({ id: sha(id), code, geraet: String((req.body || {}).geraet || 'Garmin-Uhr').slice(0, 60), bis: Date.now() + CODE_MS, token: null });
    schreiben('widget_uhr_anfragen', alle);
    res.json({ ok: true, id, code, gueltigBis: Date.now() + CODE_MS });
  });
  app.get('/api/widget/uhr/status', (req, res) => {
    const id = sha(String(req.query.id || ''));
    const alle = anfragen();
    const a = alle.find((x) => x.id === id);
    if (!a) return res.status(404).json({ ok: false, error: 'Anfrage abgelaufen.' });
    if (!a.token) return res.json({ ok: true, wartet: true });
    schreiben('widget_uhr_anfragen', alle.filter((x) => x !== a));
    res.json({ ok: true, token: a.token });
  });
  app.post('/api/widget/uhr/bestaetigen', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const code = String((req.body || {}).code || '').replace(/\D/g, '');
    const alle = anfragen();
    const a = alle.find((x) => !x.token && x.code === code);
    if (!a) return res.status(400).json({ ok: false, error: 'Kein offener Code von der Uhr – auf der Uhr neu koppeln.' });
    const token = crypto.randomBytes(32).toString('hex');
    const xs = geraete();
    xs.push({ id: crypto.randomUUID(), hash: sha(token), name: a.geraet, at: Date.now(), zuletzt: null });
    schreiben('widget_geraete', xs.slice(-10));
    a.token = token;
    a.code = null;
    schreiben('widget_uhr_anfragen', alle);
    res.json({ ok: true, name: a.geraet, geraete: xs.slice(-10).map(({ hash, ...g }) => g) });
  });

  app.get('/api/widget/geraete', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    res.json({ ok: true, geraete: geraete().map(({ hash, ...g }) => g) });
  });
  app.post('/api/widget/widerrufen', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const id = String((req.body || {}).id || '');
    schreiben('widget_geraete', geraete().filter((g) => g.id !== id));
    res.json({ ok: true, geraete: geraete().map(({ hash, ...g }) => g) });
  });

  // ---------- Knöpfe im Widget ----------
  app.post('/api/widget/aktion', (req, res) => {
    if (!geraetZu(req)) return res.status(401).json({ ok: false, error: 'App nicht gekoppelt.' });
    const aktion = String((req.body || {}).aktion || '');
    if (aktion === 'garmin') {
      // gleiche Warteschlange wie der Knopf „Garmin jetzt abrufen“ im Cockpit
      const a = lesen('garmin_anfrage', {});
      const pending = !!a.requestedAt && !(a.doneAt && a.doneAt >= a.requestedAt);
      if (!pending) {
        a.requestedAt = Date.now();
        a.pickedAt = null;
        schreiben('garmin_anfrage', a);
      }
      return res.json({ ok: true, text: 'Garmin-Abruf an den Mac gegeben.' });
    }
    if (aktion === 'briefing_lokal') {
      const [code, body] = texte.auftragStarten({ art: 'status', engine: 'local' });
      return res.status(code).json(code === 200 ? { ok: true, text: 'KI-Analyse lokal an den Mac gegeben.' } : body);
    }
    res.status(400).json({ ok: false, error: 'Unbekannte Aktion.' });
  });
}
