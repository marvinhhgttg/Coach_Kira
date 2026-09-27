import express from 'express';
import cookieParser from 'cookie-parser';
import Database from 'better-sqlite3';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 5000);
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
    maxAge: 1000 * 60 * 60 * 24 * 45,
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
app.use((req, _res, next) => {
  if (APP_BASE && req.url === APP_BASE) req.url = '/';
  else if (APP_BASE && req.url.startsWith(`${APP_BASE}/`)) req.url = req.url.slice(APP_BASE.length) || '/';
  next();
});

app.get('/api/auth/status', (req, res) => {
  res.json({
    ok: true,
    configured: !!getSetting('dashboard_token') && !!getSetting('pin_hash'),
    authenticated: isAuthed(req),
  });
});

app.post('/api/auth/setup', (req, res) => {
  const { token, pin } = req.body || {};
  if (!token || String(token).length < 20) return res.status(400).json({ ok: false, error: 'token_too_short' });
  if (!pin || String(pin).length < 6) return res.status(400).json({ ok: false, error: 'pin_too_short' });
  const { salt, hash } = hashPin(pin);
  setSetting('dashboard_token', String(token));
  setSetting('pin_salt', salt);
  setSetting('pin_hash', hash);
  createSession(res);
  res.json({ ok: true, configured: true, authenticated: true });
});

app.post('/api/auth/login', (req, res) => {
  const { pin } = req.body || {};
  if (!verifyPin(pin)) return res.status(401).json({ ok: false, error: 'invalid_pin' });
  createSession(res);
  res.json({ ok: true, authenticated: true });
});

app.post('/api/auth/logout', (req, res) => {
  clearSession(req, res);
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

app.use(express.static(path.join(__dirname, 'dist')));
app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Coach Kira proxy listening on ${PORT}${APP_BASE || ''}`);
});
