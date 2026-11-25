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
      const p = spawn(process.execPath, [KERN], { stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, CK_LIVE: getSetting('garmin_live') || '' } });
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
  // 23.11.2026: Garmin-Livewerte vom Mac drüberlegen (nur Anzeige, höchstens 3 h alt, nur heute)
  function live(a) {
    let w = null;
    try { w = JSON.parse(getSetting('garmin_live') || 'null'); } catch { w = null; }
    if (!w || w.datum !== heute() || !a || Date.now() - Number(w.at) > 3 * 3600 * 1000) return a;
    const o = { ...a, live: { at: w.at, trZeit: w.trZeit ?? null, status: w.status ?? null, erholH: w.erholH ?? null } };
    // Einheiten von heute live vom Mac (nach dem Speichern auf der Uhr, ca. 10 min Verzug)
    const akt = Array.isArray(w.akt) ? w.akt : [];
    const aktSumme = Math.round(akt.reduce((s, x) => s + (Number(x.last) > 0 ? Number(x.last) : 0), 0));
    if (akt.length) o.liveAkt = akt;
    if (o.uhr && o.uhr.feld && aktSumme > (o.uhr.feld.vorher || 0)) o.uhr = { ...o.uhr, feld: { ...o.uhr.feld, vorher: aktSumme } };
    if (o.last && aktSumme > (o.last.istHeute || 0)) o.last = { ...o.last, istHeute: aktSumme, erledigt: true };
    if (w.tr != null && o.morgen) o.morgen = { ...o.morgen, readiness: w.tr, readinessQuelle: 'garmin live', readinessZeit: w.trZeit ?? null };
    if (o.last && w.acwr != null) o.last = { ...o.last, acwr: w.acwr, acwrBand: w.acwr > 1.5 ? 'hoch' : w.acwr > 1.3 ? 'erhoeht' : w.acwr < 0.8 ? 'niedrig' : 'korridor' };
    if (o.form) o.form = { ...o.form, ...(w.ctl != null ? { ctl: w.ctl } : {}), ...(w.atl != null ? { atl: w.atl } : {}), ...(w.es != null ? { es: w.es } : {}) };
    const u = o.uhr;
    if (u && u.form) {
      const letzt = (xs, v) => (Array.isArray(xs) && xs.length && v != null ? [...xs.slice(0, -1), Math.round(v)] : xs);
      o.uhr = { ...u, form: { ...u.form, ctl: letzt(u.form.ctl, w.ctl), atl: letzt(u.form.atl, w.atl), es: letzt(u.form.es, w.es) } };
    }
    return o;
  }
  // 23.11.2026: Wochenbudget je Kalenderwoche einmal festhalten (erster Abruf der Woche = Kiras Plan Mo..So)
  function woche(o) {
    const w = o && o.woche;
    if (!w || !w.kw) return o;
    let snap = lesen('wochen_budget', null);
    if (!snap || snap.kw !== w.kw || !(snap.budget > 0)) {
      snap = { kw: w.kw, budget: w.budgetBasis, at: Date.now(), rekonstruiert: w.istVorher > 0 };
      if (snap.budget > 0) schreiben('wochen_budget', snap);
    }
    const istHeute = o.last && o.last.istHeute != null ? o.last.istHeute : 0;
    const ist = Math.max(w.ist || 0, (w.istVorher || 0) + istHeute);
    return { ...o, woche: { kw: w.kw, ist, budget: snap.budget, rekonstruiert: !!snap.rekonstruiert } };
  }
  // 23.11.2026: Kiras Fazit nach dem Speichern – die Uhr meldet es einmal je neuer Einheit
  function fazit(o) {
    try {
      if (!o || o.datum !== heute()) return o;
      let ein = [];
      if (Array.isArray(o.liveAkt) && o.liveAkt.length) ein = o.liveAkt.filter((x) => Number(x.last) > 0 && Number(x.min) >= 10);
      else if (o.uhr && o.uhr.letzte && o.uhr.letzte.d === o.datum) ein = (o.uhr.letzte.einheiten || []).map((x) => ({ id: null, start: '', name: x.name, min: x.min, last: x.ess }));
      if (!ein.length) return o;
      const letzte = ein[ein.length - 1];
      const tag = Math.round(ein.reduce((s, x) => s + Number(x.last || 0), 0));
      const lo = o.last?.sollMin, hi = o.last?.sollMax;
      const lage = hi == null ? '' : tag > hi ? `über Rahmen ${lo ?? 0}–${hi}` : lo != null && tag < lo ? `unter Rahmen ${lo}–${hi}` : `im Rahmen ${lo ?? 0}–${hi}`;
      const n = (o.naechste || [])[0];
      const morgen = !n ? '' : !n.ess ? 'Morgen Ruhetag.' : `Morgen ${n.sport || 'Training'}${n.zone ? ' ' + n.zone : ''} ${Math.round(n.ess)} ESS.`;
      const id = String(letzte.id || `${o.datum}-${ein.length}-${tag}`);
      return { ...o, fazit: { id, titel: `Kira: ${tag} ESS heute`, unter: lage, text: `${ein.length > 1 ? ein.length + ' Einheiten · ' : ''}${String(letzte.name || '').slice(0, 24)} ${letzte.min}' ${Math.round(letzte.last)} ESS. ${morgen}`.trim() } };
    } catch {
      return o;
    }
  }
  // 24.11.2026: Stufe 3 – neue Livewerte vom Mac → sofort neu rechnen (Uhr und Widgets bekommen den Stand beim nächsten Abruf)
  globalThis.ckWidgetNeu = () => { rechnen().catch(() => {}); };
  // 24.11.2026: Automatik (server/auto.mjs) liest denselben Stand wie Uhr und Widgets
  globalThis.ckWidgetAntwort = (frisch = false) => antwort(frisch);
  async function antwort(frisch = false) {
    return fazit(woche(live(await antwortRoh(frisch))));
  }
  async function antwortRoh(frisch = false) {
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
        const { texte: _t, uhr: u, liveAkt: _la, ...rest } = a;
        // lange Einheiten aus Kiras Wochenplan (Server) an die nächsten Tage hängen
        let wp = null;
        try { wp = JSON.parse(getSetting('wochenplan') || 'null'); } catch { wp = null; }
        const lang = new Map(((wp && wp.plan && wp.plan.tage) || []).filter((t) => t && t.langH != null).map((t) => [t.date, t.langH]));
        const naechste = (a.naechste || []).map((x) => (lang.has(x.d) ? { ...x, langH: lang.get(x.d) } : x));
        // 22.11.2026: Diagramme nur im Vordergrund (?mehr=1) – der Hintergrunddienst bleibt unter 8 KB
        const mehr = req.query.mehr === '1' && u ? { uhr: { l14: u.l14, m7: u.m7, form: u.form, letzte: u.letzte, plan7: u.plan7 || null, x: u.x || null } } : {};
        // 24.11.2026: Hinweise (einmal je id auf der Uhr) und Plan-Vorschlag (Uhr-Menü: übernehmen/verwerfen)
        let hinw = [];
        try { hinw = (globalThis.ckHinweise?.(a, 'uhr') || []).slice(0, 6).map(({ id, titel, text }) => ({ id, titel, text })); } catch {}
        const st = globalThis.ckVorschlag?.status?.() || {};
        const pv = a.planVorschlag && st.id !== a.planVorschlag.id ? { id: a.planVorschlag.id, text: a.planVorschlag.aenderungen.map((x) => `${x.d.slice(8, 10)}.${x.d.slice(5, 7)}. ${x.alt}>${x.neu}`).join(' · '), max: a.planVorschlag.maxNeu } : null;
        const { planVorschlag: _pv, schlaf: _sl, ...rest2 } = rest;
        return res.json({ ...rest2, naechste: req.query.mehr === '1' ? naechste : naechste.slice(0, 4), votum: a.votum ? { ...a.votum, gruende: (a.votum.gruende || []).slice(0, 2) } : null, hinweise: hinw, vorschlag: pv, ...mehr });
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
  // 25.11.2026: Push-Gerätekennung der iPhone-/iPad-App; der Mac verschickt darüber Kiras Hinweise (APNs)
  app.post('/api/widget/push', (req, res) => {
    const g = geraetZu(req);
    if (!g) return res.status(401).json({ ok: false, error: 'App nicht gekoppelt.' });
    const token = String(req.body?.token || '').toLowerCase();
    if (!/^[0-9a-f]{64,200}$/.test(token)) return res.status(400).json({ ok: false, error: 'Ungültige Gerätekennung.' });
    const umgebung = req.body?.umgebung === 'development' ? 'development' : 'production';
    const xs = geraete();
    for (const x of xs) if (x.push?.token === token && x.id !== g.id) delete x.push;
    const ich = xs.find((x) => x.id === g.id);
    if (ich) ich.push = { token, umgebung, at: Date.now() };
    schreiben('widget_geraete', xs);
    res.json({ ok: true });
  });

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

  // ---------- 25.11.2026: Tour-Rechner für die Uhr (eigene Touren, sonst ESS-Rechner) ----------
  app.post('/api/widget/uhr/tour', (req, res) => {
    if (!geraetZu(req)) return res.status(401).json({ ok: false, error: 'App nicht gekoppelt.' });
    const teile = Array.isArray(req.body?.teile) ? req.body.teile.slice(0, 6).map((x) => ({ s: String(x.s || '').slice(0, 10), km: Number(x.km) || 0, hm: Number(x.hm) || 0 })) : [];
    if (!teile.length) return res.status(400).json({ ok: false, error: 'Keine Abschnitte.' });
    const p = spawn(process.execPath, [KERN, '--tour', JSON.stringify({ teile })], { stdio: ['ignore', 'pipe', 'pipe'] });
    let out = '';
    const to = setTimeout(() => p.kill('SIGKILL'), 90000);
    p.stdout.on('data', (d) => (out += d));
    p.on('close', () => {
      clearTimeout(to);
      try { res.json(JSON.parse(out)); } catch { res.status(502).json({ ok: false, error: 'Tour-Rechnung fehlgeschlagen.' }); }
    });
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
    if (aktion === 'plan_annehmen' || aktion === 'plan_verwerfen') {
      const id = String((req.body || {}).id || '');
      const v = globalThis.ckVorschlag;
      if (!v) return res.status(503).json({ ok: false, error: 'Automatik nicht aktiv.' });
      return Promise.resolve(aktion === 'plan_annehmen' ? v.annehmen(id) : v.verwerfen(id))
        .then((r) => res.status(r.ok ? 200 : 409).json(r))
        .catch((e) => res.status(502).json({ ok: false, error: String(e.message || e) }));
    }
    if (aktion === 'plan_aendern') {
      const f = globalThis.ckPlanAendern;
      if (!f) return res.status(503).json({ ok: false, error: 'Automatik nicht aktiv.' });
      return Promise.resolve(f(req.body || {}))
        .then((r) => res.status(r.ok ? 200 : 409).json(r))
        .catch((e) => res.status(502).json({ ok: false, error: String(e.message || e) }));
    }
    if (aktion === 'briefing_lokal') {
      const [code, body] = texte.auftragStarten({ art: 'status', engine: 'local' });
      return res.status(code).json(code === 200 ? { ok: true, text: 'KI-Analyse lokal an den Mac gegeben.' } : body);
    }
    res.status(400).json({ ok: false, error: 'Unbekannte Aktion.' });
  });
}
