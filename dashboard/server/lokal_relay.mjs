/**
 * Lokal-Weiterleitung: Geräte ohne eigenes Ollama (iPhone-App, iPad, fremder Rechner) schicken lokale
 * KI-Anfragen hierher. Der Mac holt sie ab (Supervisor-Takt, danach im 2-s-Takt), rechnet mit coach-kira
 * und liefert die Antwort zurück. Kein Cloud-Modell beteiligt – der Server ist nur Briefkasten.
 * Aufträge liegen nur im Arbeitsspeicher (keine Gesundheitsdaten auf der Platte) und verfallen nach 30 min.
 */
import crypto from 'node:crypto';

const VERFALL_MS = 30 * 60 * 1000;
const ABHOLEN_MS = 3 * 60 * 1000; // so lange darf ein Auftrag auf den Mac warten
const MODELLE = /^coach-kira(?::[\w.\-]+)?$/;

export function setupLokalRelay(app, { originOk, macStatus, macPuls }) {
  const auftraege = new Map(); // id → { body, phase, at, abgeholtAt, text, fehler, tokPerS }

  const aufraeumen = () => {
    const jetzt = Date.now();
    for (const [id, a] of auftraege) {
      if (jetzt - a.at > VERFALL_MS) auftraege.delete(id);
      else if (a.phase === 'wartet' && jetzt - a.at > ABHOLEN_MS) Object.assign(a, { phase: 'fehler', fehler: 'Der Mac hat die Anfrage nicht abgeholt (nicht erreichbar oder im Ruhezustand).' });
    }
  };

  // Browser/App: Auftrag anlegen
  app.post('/api/lokal/auftrag', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    aufraeumen();
    const b = req.body || {};
    const body = b.body || {};
    if (!Array.isArray(body.messages) || !body.messages.length) return res.status(400).json({ ok: false, error: 'Keine Nachrichten.' });
    if (!MODELLE.test(String(body.model || ''))) return res.status(400).json({ ok: false, error: 'Nur das lokale Modell coach-kira ist erlaubt.' });
    const m = macStatus();
    if (!m.erreichbar) return res.status(503).json({ ok: false, error: 'Der Mac ist gerade nicht erreichbar.' });
    if ([...auftraege.values()].filter((x) => x.phase === 'wartet' || x.phase === 'rechnet').length >= 5) {
      return res.status(429).json({ ok: false, error: 'Zu viele offene lokale Anfragen – kurz warten.' });
    }
    const id = crypto.randomUUID();
    auftraege.set(id, {
      body: { model: String(body.model), messages: body.messages, format: body.format, options: body.options || {}, think: false, keep_alive: '30m' },
      phase: 'wartet',
      at: Date.now(),
      zweck: String(b.zweck || 'chat').slice(0, 40),
    });
    res.json({ ok: true, id });
  });

  // Browser/App: Stand abfragen
  app.get('/api/lokal/auftrag/:id', (req, res) => {
    aufraeumen();
    const a = auftraege.get(req.params.id);
    if (!a) return res.status(404).json({ ok: false, error: 'Anfrage unbekannt oder abgelaufen.' });
    const r = { ok: true, phase: a.phase, wartetSeit: a.at, abgeholtAt: a.abgeholtAt || null };
    if (a.phase === 'fertig') Object.assign(r, { text: a.text, tokPerS: a.tokPerS ?? null });
    if (a.phase === 'fehler') r.fehler = a.fehler;
    res.json(r);
  });

  // Mac: nächsten Auftrag holen
  app.get('/api/lokal/naechster', (req, res) => {
    aufraeumen();
    macPuls(); // Abholen zählt als Lebenszeichen des Macs
    const e = [...auftraege.entries()].find(([, a]) => a.phase === 'wartet');
    if (!e) return res.json({ ok: true, pending: false });
    const [id, a] = e;
    Object.assign(a, { phase: 'rechnet', abgeholtAt: Date.now() });
    res.json({ ok: true, pending: true, id, body: a.body, zweck: a.zweck });
  });

  // Mac: Ergebnis abliefern
  app.post('/api/lokal/ergebnis', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const b = req.body || {};
    const a = auftraege.get(String(b.id || ''));
    if (!a) return res.json({ ok: false, error: 'unbekannt' });
    if (b.ok) Object.assign(a, { phase: 'fertig', text: String(b.text || ''), tokPerS: b.tokPerS ?? null, body: null });
    else Object.assign(a, { phase: 'fehler', fehler: String(b.fehler || 'Fehler auf dem Mac').slice(0, 300), body: null });
    res.json({ ok: true });
  });

  // Status für die Anzeige „Lokal über den Mac“
  app.get('/api/lokal/status', (req, res) => {
    const m = macStatus();
    res.json({ ok: true, mac: m, offen: [...auftraege.values()].filter((x) => x.phase === 'wartet' || x.phase === 'rechnet').length });
  });
}
