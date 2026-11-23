import express from 'express';
import cookieParser from 'cookie-parser';
import Database from 'better-sqlite3';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { setupTexte } from './server/texte.mjs';
import { setupWidget } from './server/widget_api.mjs';
import { setupLokalRelay } from './server/lokal_relay.mjs';
import { setupEinstellungen } from './server/einstellungen.mjs';
import { setupWochenplan } from './server/wochenplan.mjs';
import { setupZugang } from './server/zugang.mjs';
import { setupRechenkern } from './server/rechenkern_api.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 5000);
// Etappe 6: Server-seitige Rechner (Widget, Votum, Prompts) lesen den Verlauf aus timeline über diesen Server
if (!process.env.CK_ZEITREIHE_URL) process.env.CK_ZEITREIHE_URL = `http://127.0.0.1:${PORT}/api/rechenkern/chartData`;
const APP_BASE = (process.env.APP_BASE || '').replace(/\/$/, '');
const APPS_SCRIPT_URL =
  'https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec';
const PROD = process.env.NODE_ENV === 'production';
const COOKIE_NAME = PROD ? '__Host-coachkira_sid' : 'coachkira_sid';

const db = new Database(path.join(__dirname, 'data.db'));
db.pragma('journal_mode = WAL');
db.exec(`
CREATE TABLE IF NOT EXISTS settings (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions (
  id TEXT PRIMARY KEY,
  created_at INTEGER NOT NULL,
  expires_at INTEGER NOT NULL
);
`);

function now() {
  return Date.now();
}

function getSetting(key) {
  return db.prepare('SELECT value FROM settings WHERE key=?').get(key)?.value || null;
}

function setSetting(key, value) {
  db.prepare('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value')
    .run(key, value);
}

function hashPin(pin, salt = crypto.randomBytes(16).toString('hex')) {
  const hash = crypto.pbkdf2Sync(String(pin), salt, 210000, 32, 'sha256').toString('hex');
  return { salt, hash };
}

function verifyPin(pin) {
  const salt = getSetting('pin_salt');
  const expected = getSetting('pin_hash');
  if (!salt || !expected) return false;
  const { hash } = hashPin(pin, salt);
  return crypto.timingSafeEqual(Buffer.from(hash, 'hex'), Buffer.from(expected, 'hex'));
}

function createSession(res) {
  const id = crypto.randomBytes(32).toString('hex');
  const expires = now() + 1000 * 60 * 60 * 24 * 45;
  db.prepare('INSERT INTO sessions(id,created_at,expires_at) VALUES(?,?,?)').run(id, now(), expires);
  res.cookie(COOKIE_NAME, id, {
    httpOnly: true,
    sameSite: 'lax',
    secure: PROD,
    path: '/',
    maxAge: 1000 * 60 * 60 * 24 * 400, // Ablauf steuert der Server (gleitend 45 Tage ab letzter Nutzung)
  });
  return id;
}

function clearSession(req, res) {
  const id = req.cookies?.[COOKIE_NAME];
  if (id) db.prepare('DELETE FROM sessions WHERE id=?').run(id);
  res.clearCookie(COOKIE_NAME, { path: '/', secure: PROD, sameSite: 'lax' });
}

function isAuthed(req) {
  const id = req.cookies?.[COOKIE_NAME];
  if (!id) return false;
  const row = db.prepare('SELECT expires_at FROM sessions WHERE id=?').get(id);
  if (!row || row.expires_at < now()) {
    if (row) db.prepare('DELETE FROM sessions WHERE id=?').run(id);
    return false;
  }
  // gleitend: bei Nutzung wieder 45 Tage (höchstens einmal pro Stunde schreiben)
  const neu = now() + 1000 * 60 * 60 * 24 * 45;
  if (neu - row.expires_at > 1000 * 60 * 60) db.prepare('UPDATE sessions SET expires_at=? WHERE id=?').run(neu, id);
  return true;
}

function requireAuth(req, res, next) {
  if (!isAuthed(req)) return res.status(401).json({ ok: false, error: 'not_authenticated' });
  next();
}

function actionToAppsScriptParams(name, body) {
  if (name === 'runSupervisor') return { action: 'runSupervisor' };
  if (name === 'runStatusAiAnalysis') return { action: 'runStatusAiAnalysis' };
  if (name === 'runConsolidatedAnalysis') return { action: 'runConsolidatedAnalysis' };
  if (name === 'generatePlanBriefing') return { action: 'generatePlanBriefing' };
  if (name === 'saveSimulatedPlan') {
    return {
      action: 'saveSimulatedPlan',
      loads: JSON.stringify(body.loads || []),
      teAe: JSON.stringify(body.teAe || []),
      teAn: JSON.stringify(body.teAn || []),
      sports: JSON.stringify(body.sports || []),
      zones: JSON.stringify(body.zones || []),
      locks: JSON.stringify(body.locks || []),
    };
  }
  if (name === 'submitEntry') {
    return {
      action: 'submitCockpitEntry',
      type: String(body.type || ''),
      date: String(body.date || ''),
      values: JSON.stringify(body.values || {}),
    };
  }
  if (name === 'saveWellbeing') {
    return { action: 'saveWellbeing', date: String(body.date || ''), value: String(body.value ?? '') };
  }
  return null;
}

async function callAppsScript(params) {
  const token = getSetting('dashboard_token');
  if (!token) return { ok: false, error: 'token_not_configured' };
  const usp = new URLSearchParams({ ...params, token, cb: String(Date.now()) });
  const res = await fetch(`${APPS_SCRIPT_URL}?${usp.toString()}`, {
    method: 'GET',
    redirect: 'follow',
    headers: { 'user-agent': 'coach-kira-private-proxy/1.0' },
  });
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    return { ok: false, error: 'apps_script_non_json', raw: text.slice(0, 1000) };
  }
}

const app = express();
app.set('trust proxy', 1);
app.use(express.json({ limit: '1mb' }));
app.use(cookieParser());
// Zweitinstanz (v2): nur lesend. Alle schreibenden API-Aufrufe sind gesperrt –
// keine Ersteinrichtung/PIN, keine Cloud-KI, keine Gerätekopplung, keine Aufträge an den Mac.
const INSTANZ_V2 = (process.env.CK_INSTANZ || 'live') !== 'live';
app.use((req, res, next) => {
  if (INSTANZ_V2 && req.method !== 'GET' && req.method !== 'HEAD' && req.path.startsWith('/api/')) {
    return res.status(409).json({ ok: false, error: 'v2-Testinstanz ist nur lesend. Schreiben bitte im Cockpit auf coachkira.pplx.app.' });
  }
  next();
});
app.use((req, _res, next) => {
  if (APP_BASE && req.url === APP_BASE) req.url = '/';
  else if (APP_BASE && req.url.startsWith(`${APP_BASE}/`)) req.url = req.url.slice(APP_BASE.length) || '/';
  next();
});

const zugang = setupZugang(app, { getSetting, setSetting, isAuthed, browserOk, scriptUrl: APPS_SCRIPT_URL, aktiv: !INSTANZ_V2 });

app.get('/api/auth/status', (req, res) => {
  res.json({
    ok: true,
    configured: !!getSetting('pin_hash'),
    authenticated: isAuthed(req),
  });
});

app.post('/api/auth/setup', (req, res) => {
  // Etappe 3b: PIN festlegen – nur mit Einrichtungscode (CK_SETUP_HASH), nur einmal, nur aus dem Cockpit.
  if (!browserOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
  if (getSetting('pin_hash') && !isAuthed(req)) return res.status(403).json({ ok: false, error: 'already_configured' });
  const { code, pin } = req.body || {};
  if (!/^\d{6,12}$/.test(String(pin || ''))) return res.status(400).json({ ok: false, error: 'PIN: 6 bis 12 Ziffern.' });
  if (!getSetting('pin_hash')) {
    const c = zugang.codeOk(code);
    if (!c.ok) return res.status(c.status).json({ ok: false, error: c.error });
  }
  const { salt, hash } = hashPin(pin);
  setSetting('pin_salt', salt);
  setSetting('pin_hash', hash);
  createSession(res);
  res.json({ ok: true, configured: true, authenticated: true });
});

app.post('/api/auth/login', (req, res) => {
  const { pin } = req.body || {};
  // Etappe 3b: 5 Fehlversuche → 15 Minuten Pause
  let f = {};
  try { f = JSON.parse(getSetting('login_fehl') || '{}'); } catch {}
  if ((f.bis || 0) > now()) return res.status(429).json({ ok: false, error: 'Zu viele Fehlversuche – bitte 15 Minuten warten.' });
  if (!verifyPin(pin)) {
    f.n = (f.n || 0) + 1;
    if (f.n >= 5) { f.n = 0; f.bis = now() + 15 * 60 * 1000; }
    setSetting('login_fehl', JSON.stringify(f));
    return res.status(401).json({ ok: false, error: 'invalid_pin' });
  }
  setSetting('login_fehl', JSON.stringify({ n: 0, bis: 0 }));
  createSession(res);
  res.json({ ok: true, authenticated: true });
});

app.post('/api/auth/logout', (req, res) => {
  clearSession(req, res);
  res.json({ ok: true });
});

// ---------------------------------------------------------------------------
// Cloud-KI (Perplexity) – ersetzt alle KI-Aufrufe im Apps Script.
// Schlüssel liegt im Perplexity-Credential-Tresor und wird über den Proxy
// (CUSTOM_CRED_API_PERPLEXITY_AI_URL + x-api-key) eingespielt, nie im Code.
// ---------------------------------------------------------------------------
const PPLX_URL = process.env.CUSTOM_CRED_API_PERPLEXITY_AI_URL || '';
const PPLX_KEY = process.env.CUSTOM_CRED_API_PERPLEXITY_AI_PROXY_AUTH_KEY || '';
const LLM_DAILY_LIMIT = Number(process.env.CK_LLM_DAILY_LIMIT || 80);
const LLM_MODELS = new Set(['sonar', 'sonar-pro']);

function llmDayKey() {
  const d = new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
  return 'llm_count_' + d;
}
function llmUsed() {
  return Number(getSetting(llmDayKey()) || 0);
}
// Etappe 3a (20.11.2026): exakte Herkunft statt Präfix bzw. *.pplx.app-Joker.
function herkunftErlaubt(o) {
  try {
    const u = new URL(o);
    if (u.origin === 'https://coachkira.pplx.app') return true;
    if (u.protocol === 'http:' && (u.hostname === 'localhost' || u.hostname === '127.0.0.1')) return true;
    return false;
  } catch {
    return false;
  }
}
/** Für Endpunkte, die auch der Mac-Runner ohne Browser aufruft: fremde Browser-Herkunft wird abgewiesen. */
function llmOriginOk(req) {
  const o = String(req.headers.origin || req.headers.referer || '');
  if (!o) return true; // Mac-Runner (kein Browser) – Absicherung mit Geheimnis folgt in Etappe 3b
  return herkunftErlaubt(o);
}
/** Nur für Aufrufe aus dem Cockpit selbst (Browser bzw. App-WebView): Herkunft muss vorhanden und exakt sein. */
function browserOk(req) {
  const o = String(req.headers.origin || '');
  if (o) return herkunftErlaubt(o);
  const site = String(req.headers['sec-fetch-site'] || '');
  if (site && site !== 'same-origin') return false;
  const r = String(req.headers.referer || '');
  return !!r && herkunftErlaubt(r);
}

app.get('/api/llm/status', (req, res) => {
  res.json({ ok: true, configured: !!(PPLX_URL && PPLX_KEY), used: llmUsed(), limit: LLM_DAILY_LIMIT, models: [...LLM_MODELS] });
});

/** Cloud-KI für serverseitige Aufgaben (Texte-Wächter) – zählt ins gleiche Tageslimit. */
async function cloudChatServer(messages, opts = {}) {
  if (!PPLX_URL || !PPLX_KEY) throw new Error('Cloud-KI ist auf dem Server nicht eingerichtet.');
  const used = llmUsed();
  if (used >= LLM_DAILY_LIMIT) throw new Error(`Tageslimit von ${LLM_DAILY_LIMIT} Cloud-Anfragen erreicht.`);
  setSetting(llmDayKey(), String(used + 1));
  const msgs = opts.json
    ? [{ role: 'system', content: 'Antworte ausschließlich mit gültigem JSON – ohne Codeblock, ohne Text davor oder danach.' }, ...messages]
    : messages;
  const model = LLM_MODELS.has(opts.model) ? opts.model : 'sonar-pro';
  const r = await fetch(PPLX_URL.replace(/\/$/, '') + '/chat/completions', {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-api-key': PPLX_KEY },
    body: JSON.stringify({ model, messages: msgs, temperature: 0.3, max_tokens: Math.max(100, Math.min(8000, Number(opts.maxTokens) || 2500)), disable_search: true }),
  });
  const text = await r.text();
  let j;
  try { j = JSON.parse(text); } catch { throw new Error(`Perplexity HTTP ${r.status}`); }
  if (!r.ok) throw new Error(j?.error?.message || `Perplexity HTTP ${r.status}`);
  const answer = String(j?.choices?.[0]?.message?.content || '').trim();
  if (!answer) throw new Error('Cloud-KI lieferte keinen Text.');
  return { answer, model };
}

app.post('/api/llm', async (req, res) => {
  try {
    // Cockpit (Browser) oder Mac-Runner (ohne Herkunft, Cloud-Ersatz) – fremde Browser-Herkunft nie
    if (!llmOriginOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    if (!PPLX_URL || !PPLX_KEY) return res.status(503).json({ ok: false, error: 'Cloud-KI ist auf dem Server nicht eingerichtet (Perplexity-Schlüssel fehlt).' });
    const b = req.body || {};
    const msgs = Array.isArray(b.messages) ? b.messages : [];
    const clean = msgs
      .filter((m) => m && ['system', 'user', 'assistant'].includes(m.role) && typeof m.content === 'string')
      .map((m) => ({ role: m.role, content: m.content.slice(0, 60000) }));
    if (!clean.length || clean[clean.length - 1].role !== 'user') return res.status(400).json({ ok: false, error: 'messages fehlen oder letzte Nachricht ist nicht vom Nutzer.' });
    const total = clean.reduce((s, m) => s + m.content.length, 0);
    if (total > 120000) return res.status(413).json({ ok: false, error: 'Anfrage zu groß.' });
    const used = llmUsed();
    if (used >= LLM_DAILY_LIMIT) return res.status(429).json({ ok: false, error: `Tageslimit von ${LLM_DAILY_LIMIT} Cloud-Anfragen erreicht.`, used, limit: LLM_DAILY_LIMIT });
    setSetting(llmDayKey(), String(used + 1));
    const model = LLM_MODELS.has(b.model) ? b.model : 'sonar-pro';
    const payload = {
      model,
      messages: clean,
      temperature: typeof b.temperature === 'number' ? Math.max(0, Math.min(1, b.temperature)) : 0.3,
      max_tokens: Math.max(100, Math.min(8000, Number(b.max_tokens) || 2500)),
    };
    if (b.sources) payload.search_mode = 'academic';
    else payload.disable_search = true;
    const t0 = Date.now();
    const r = await fetch(PPLX_URL.replace(/\/$/, '') + '/chat/completions', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': PPLX_KEY },
      body: JSON.stringify(payload),
    });
    const text = await r.text();
    let j;
    try {
      j = JSON.parse(text);
    } catch {
      return res.status(502).json({ ok: false, error: `Perplexity HTTP ${r.status}: ${text.slice(0, 200)}` });
    }
    if (!r.ok) return res.status(502).json({ ok: false, error: j?.error?.message || `Perplexity HTTP ${r.status}` });
    const answer = j?.choices?.[0]?.message?.content || '';
    res.json({ ok: true, answer, model, citations: j.citations || [], usage: j.usage || null, used: used + 1, limit: LLM_DAILY_LIMIT, ms: Date.now() - t0 });
  } catch (err) {
    res.status(502).json({ ok: false, error: err?.message || String(err) });
  }
});

// Garmin-Abruf auf Knopfdruck: Cockpit legt einen Auftrag an, der Mac (Supervisor-Takt, 60 s) holt ihn ab,
// ruft Garmin sofort ab und meldet zurück. Nur ein Flag – keine Garmin-Daten auf diesem Server.
function garminAnfrage() {
  try { return JSON.parse(getSetting('garmin_anfrage') || '{}'); } catch { return {}; }
}
app.get('/api/garmin/anfrage', (req, res) => {
  const a = garminAnfrage();
  res.json({ ok: true, ...a, pending: !!a.requestedAt && !(a.doneAt && a.doneAt >= a.requestedAt), now: now() });
});
app.post('/api/garmin/anfrage', (req, res) => {
  if (!browserOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
  const a = garminAnfrage();
  const pending = !!a.requestedAt && !(a.doneAt && a.doneAt >= a.requestedAt);
  if (!pending) {
    a.requestedAt = now();
    a.pickedAt = null;
    setSetting('garmin_anfrage', JSON.stringify(a));
  }
  res.json({ ok: true, ...a, pending: true, now: now() });
});
app.post('/api/garmin/anfrage/status', (req, res) => {
  if (!llmOriginOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
  const b = req.body || {};
  const a = garminAnfrage();
  if (b.phase === 'abgeholt') a.pickedAt = now();
  if (b.phase === 'fertig') {
    a.doneAt = now();
    a.result = String(b.result || '').slice(0, 300);
    a.ok = !!b.ok;
  }
  setSetting('garmin_anfrage', JSON.stringify(a));
  res.json({ ok: true });
});

// 23.11.2026: Garmin-Livewerte. Der Mac (kira_garmin_live.py, alle 10 min) liest über den Garmin-Login
// Trainingsbereitschaft, akute/chronische Last, ACWR, Trainingszustand und Ausdauerwert und meldet sie hier.
// Nur für die Anzeige (Uhr, Zifferblätter, Widget, Heute-Seite). Die Timeline bleibt die Wahrheit, das Votum
// rechnet weiter mit den Morgenwerten.
const LIVE_FELDER = { tr: [0, 100], atl: [0, 5000], ctl: [0, 5000], acwr: [0, 10], es: [0, 15000], erholH: [0, 200], vo2: [10, 100] };
app.post('/api/garmin/live', (req, res) => {
  const b = req.body || {};
  const datum = String(b.datum || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(datum)) return res.status(400).json({ ok: false, error: 'datum fehlt' });
  const w = { datum, at: now(), quelle: 'garmin' };
  for (const [k, [lo, hi]] of Object.entries(LIVE_FELDER)) {
    const v = Number(b[k]);
    if (b[k] != null && b[k] !== '' && Number.isFinite(v) && v >= lo && v <= hi) w[k] = Math.round(v * 100) / 100;
  }
  for (const k of ['trZeit', 'status', 'trStufe']) if (b[k] != null) w[k] = String(b[k]).slice(0, 40);
  setSetting('garmin_live', JSON.stringify(w));
  res.json({ ok: true, gespeichert: Object.keys(w).length - 3 });
});
app.get('/api/garmin/live', (req, res) => {
  let w = null;
  try { w = JSON.parse(getSetting('garmin_live') || 'null'); } catch { w = null; }
  res.json({ ok: true, live: w });
});

// Workout an Garmin senden: Cockpit legt das fertige Workout als Auftrag ab, der Mac (60-s-Takt) lädt es
// mit dem eigenen Garmin-Login hoch und legt es in den Garmin-Kalender. Der Server hält nur die Warteschlange.
function garminWorkouts() {
  try { const x = JSON.parse(getSetting('garmin_workouts') || '[]'); return Array.isArray(x) ? x : []; } catch { return []; }
}
function garminWorkoutsSpeichern(xs) {
  setSetting('garmin_workouts', JSON.stringify(xs.slice(-15)));
}
app.get('/api/garmin/workout', (req, res) => {
  const xs = garminWorkouts();
  const id = String(req.query.id || '');
  if (id) {
    const x = xs.find((q) => q.id === id);
    return x ? res.json({ ok: true, auftrag: x }) : res.status(404).json({ ok: false, error: 'Auftrag unbekannt (Server neu gestartet?)' });
  }
  // Mac: offene Aufträge (ohne Abholung oder hängengeblieben > 10 min)
  const t = now();
  res.json({ ok: true, offen: xs.filter((q) => !q.doneAt && (!q.pickedAt || t - q.pickedAt > 600000)) });
});
app.post('/api/garmin/workout', (req, res) => {
  if (!browserOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
  const b = req.body || {};
  const w = b.workout;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.datum || ''))) return res.status(400).json({ ok: false, error: 'Datum fehlt.' });
  if (!w || !Array.isArray(w.blocks) || !w.blocks.length) return res.status(400).json({ ok: false, error: 'Workout fehlt.' });
  const roh = JSON.stringify(w);
  if (roh.length > 20000) return res.status(400).json({ ok: false, error: 'Workout zu groß.' });
  const xs = garminWorkouts().filter((q) => !(q.datum === b.datum && !q.doneAt)); // älteren offenen Auftrag für den Tag ersetzen
  const a = { id: crypto.randomUUID(), datum: String(b.datum), workout: w, requestedAt: now(), pickedAt: null, doneAt: null, ok: null, result: '' };
  xs.push(a);
  garminWorkoutsSpeichern(xs);
  res.json({ ok: true, auftrag: { ...a, workout: undefined } });
});
app.post('/api/garmin/workout/status', (req, res) => {
  if (!llmOriginOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
  const b = req.body || {};
  const xs = garminWorkouts();
  const a = xs.find((q) => q.id === String(b.id || ''));
  if (!a) return res.status(404).json({ ok: false, error: 'Auftrag unbekannt.' });
  if (b.phase === 'abgeholt') a.pickedAt = now();
  if (b.phase === 'fertig') {
    a.doneAt = now();
    a.ok = !!b.ok;
    a.result = String(b.result || '').slice(0, 300);
    if (b.workoutId) a.workoutId = String(b.workoutId);
  }
  garminWorkoutsSpeichern(xs);
  res.json({ ok: true });
});

app.post('/api/action/:name', requireAuth, async (req, res) => {
  const params = actionToAppsScriptParams(req.params.name, req.body || {});
  if (!params) return res.status(404).json({ ok: false, error: 'unknown_action' });
  try {
    const result = await callAppsScript(params);
    res.json(result);
  } catch (err) {
    res.status(502).json({ ok: false, error: err?.message || String(err) });
  }
});

const INSTANZ = process.env.CK_INSTANZ || 'live';
const texte = setupTexte(app, { getSetting, setSetting, cloudChat: cloudChatServer, originOk: llmOriginOk, scriptUrl: APPS_SCRIPT_URL, passiv: INSTANZ !== 'live' });
app.get('/api/instanz', (_req, res) => res.json({ instanz: INSTANZ }));
setupRechenkern(app, { getSetting, setSetting, aktiv: INSTANZ !== 'live' });
setupWidget(app, { getSetting, setSetting, originOk: browserOk, texte, appsScript: callAppsScript });
setupEinstellungen(app, { getSetting, setSetting });
setupWochenplan(app, { getSetting, setSetting });
setupLokalRelay(app, {
  getSetting,
  setSetting,
  originOk: llmOriginOk,
  macStatus: () => {
    const m = texte.statusAntwort().mac;
    return { erreichbar: !!m.erreichbar && m.ollama !== false, zuletzt: m.at, ollama: m.ollama };
  },
  macPuls: () => {
    let m = {};
    try { m = JSON.parse(getSetting('texte_mac') || '{}'); } catch {}
    m.at = now();
    setSetting('texte_mac', JSON.stringify(m));
  },
});

app.use(express.static(path.join(__dirname, 'dist')));
app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Coach Kira proxy listening on ${PORT}${APP_BASE || ''}`);
});
