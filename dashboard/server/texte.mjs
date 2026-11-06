/**
 * Texte-Wächter: Nach dem Absenden (morgens · Aktivität · abends) entstehen die Kira-Texte normalerweise
 * lokal auf dem Mac. Der Wächter prüft alle 5 Minuten, ob die Texte fehlen. Arbeitet der Mac nicht daran
 * und fehlt der Text nach 3 Prüfungen (≈ 15 min), darf – nur wenn Marc den Cloud-Ersatz eingeschaltet hat –
 * die Cloud-KI den Text erzeugen. Gespeichert wird wie gewohnt im Apps Script, mit Herkunft „cloud“.
 */
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROMPTS = path.join(__dirname, 'kira_prompts.mjs');
const TAKT_MS = Number(process.env.TEXTE_TAKT_MS) || 5 * 60 * 1000;
const VERSUCHE = 3; // 3 × 5 min ohne lokalen Text → Cloud-Ersatz
const CLOUD_VERSUCHE = 3;
const MAC_FRISCH_MS = 7 * 60 * 1000; // Herzschlag älter → Mac gilt als nicht erreichbar
const MAC_ARBEIT_MAX_MS = 45 * 60 * 1000; // länger „in Arbeit“ zählt nicht mehr als Arbeit
const ACT_METRICS =
  'activity_done,load_fb_day,coachE_ESS_day,Sport_x,Zone,Aerobic_TE,Anaerobic_TE,Target_Aerobic_TE,Target_Anaerobic_TE,' +
  'fbATL_obs,fbCTL_obs,fbACWR_obs,coachE_ATL_forecast,coachE_CTL_forecast,coachE_ACWR_forecast,Garmin_Training_Readiness,sleep_score_0_100,kcal_in';
const ACTIVITY_SYSTEM =
  'Du bist Coach Kira, eine erfahrene Sportwissenschaftlerin. Du sprichst Marc direkt per Du an. ' +
  'Antworte nur mit dem Kommentartext, ohne Anführungszeichen, ohne JSON, ohne Einleitung.';
const MORGEN = [
  ['status', 'KI-Analyse'],
  ['plan', 'Plan-Briefing'],
  ['consolidated', 'Konsolidierte Analyse'],
];

const heute = (d = new Date()) => d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
const stunde = () => Number(new Date().toLocaleString('en-GB', { timeZone: 'Europe/Berlin', hour: '2-digit', hour12: false }));
const plusTage = (iso, n) => {
  const d = new Date(iso + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};
const num = (v) => {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = String(v).trim().replace(',', '.');
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
};
const istLoad = (r) => {
  const v = num(r.load_fb_day);
  return v !== null && v > 0 ? v : num(r.coachE_ESS_day);
};
const istAktivitaet = (r) => {
  if (String(r.activity_done || '').trim().toLowerCase() !== 'x') return false;
  if (!((num(r.load_fb_day) || num(r.coachE_ESS_day) || 0) > 0)) return false;
  const s = String(r.Sport_x || '').trim().toLowerCase();
  return !['off', 'krank', ''].includes(s);
};
const f = (v, d = 0) => (v === null || v === undefined ? 'unbekannt' : Number(v).toFixed(d));

export function setupTexte(app, { getSetting, setSetting, cloudChat, originOk, scriptUrl }) {
  const lesen = (k, def) => {
    try {
      const v = getSetting(k);
      return v ? JSON.parse(v) : def;
    } catch {
      return def;
    }
  };
  const schreiben = (k, v) => setSetting(k, JSON.stringify(v));
  const zustand = () => lesen('texte_zustand', { items: {}, letzterTakt: 0, letzterFehler: null });
  const einstellung = () => ({ cloudErsatz: true, ...lesen('texte_einstellung', {}) });
  const mac = () => lesen('texte_mac', {});

  async function holen(params) {
    const usp = new URLSearchParams({ ...params, cb: String(Date.now()) });
    const r = await fetch(`${scriptUrl}?${usp}`, { redirect: 'follow', headers: { 'user-agent': 'coach-kira-texte/1.0' } });
    const t = await r.text();
    try {
      return JSON.parse(t);
    } catch {
      throw new Error('Apps Script antwortet nicht mit JSON');
    }
  }
  async function senden(mode, body) {
    const r = await fetch(`${scriptUrl}?mode=${mode}&cb=${Date.now()}`, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'content-type': 'text/plain;charset=utf-8', 'user-agent': 'coach-kira-texte/1.0' },
      body: JSON.stringify(body),
    });
    const t = await r.text();
    let j;
    try {
      j = JSON.parse(t);
    } catch {
      throw new Error(`${mode}: Apps Script antwortet nicht mit JSON`);
    }
    if (!j.ok) throw new Error(j.error || `${mode} fehlgeschlagen`);
    return j;
  }
  function prompts(art, payload) {
    return new Promise((resolve, reject) => {
      const p = spawn(process.execPath, [PROMPTS, art], { stdio: ['pipe', 'pipe', 'pipe'] });
      let out = '';
      let err = '';
      const to = setTimeout(() => p.kill('SIGKILL'), 90000);
      p.stdout.on('data', (d) => (out += d));
      p.stderr.on('data', (d) => (err += d));
      p.on('close', (code) => {
        clearTimeout(to);
        if (code !== 0) return reject(new Error('Prompt: ' + (err || 'Abbruch').slice(0, 200)));
        try {
          resolve(JSON.parse(out));
        } catch {
          reject(new Error('Prompt: keine gültige Antwort'));
        }
      });
      p.stdin.end(JSON.stringify(payload));
    });
  }

  // ---------- Cloud-Erzeugung je Textart ----------
  async function cloudBriefing(type) {
    const d = await holen({ mode: 'briefingData', type });
    if (!d.ok || !d.daten) throw new Error(d.error || 'Keine Briefing-Daten');
    const b = await prompts('briefing', { type, daten: d.daten });
    const r = await cloudChat(b.messages, { json: b.format === 'json', model: 'sonar-pro', maxTokens: 4000 });
    await senden('saveBriefing', { kind: 'saveBriefing', type, text: r.answer, model: r.model, engine: 'cloud' });
    return r.model;
  }
  async function cloudAktivitaet(row, soll) {
    const sl = soll || {};
    const text =
      `PLANVORGABE FÜR HEUTE:\n  Geplanter Load (ESS): ${f(sl.load)}\n  Geplante Sportart: ${sl.sport || 'unbekannt'}\n` +
      `  Geplante Zone: ${sl.zone || 'unbekannt'}\n  Ziel-TE aerob: ${f(sl.te_ae ?? num(row.Target_Aerobic_TE), 1)}\n` +
      `  Ziel-TE anaerob: ${f(sl.te_an ?? num(row.Target_Anaerobic_TE), 1)}\n\n` +
      `TATSÄCHLICH DURCHGEFÜHRT:\n  Datum: ${row.date}\n  Aktivität (Sport_x): ${row.Sport_x || ''}\n  Zone(n): ${row.Zone || ''}\n` +
      `  Load: ${istLoad(row)}\n  TE aerob: ${num(row.Aerobic_TE)}\n  TE anaerob: ${num(row.Anaerobic_TE)}\n` +
      `  ATL: ${num(row.fbATL_obs)}\n  CTL: ${num(row.fbCTL_obs)}\n  ACWR: ${num(row.fbACWR_obs)}\n\n` +
      `AUFGABE (MAX. 3 SÄTZE):\n1. Vergleiche SOLL mit IST.\n2. Bewerte die Abweichung (positiv, negativ, neutral).\n` +
      `3. Gib einen kurzen, motivierenden Kommentar mit Blick auf morgen.\n\n` +
      `REGELN: Nur Werte aus diesen Daten verwenden, nichts erfinden. Wenn die Planvorgabe unbekannt ist, nur die TE-Ziele vergleichen und das kurz sagen.`;
    const r = await cloudChat(
      [{ role: 'system', content: ACTIVITY_SYSTEM }, { role: 'user', content: text }],
      { model: 'sonar-pro', maxTokens: 600 },
    );
    const t = String(r.answer || '').trim().replace(/^["„“]|["”]$/g, '').trim();
    if (t.length < 40 || t.length > 900 || t.startsWith('{') || !/\d/.test(t)) throw new Error('Kommentar unbrauchbar');
    await senden('saveActivityReview', {
      kind: 'saveActivityReview',
      date: row.date,
      text: t,
      model: r.model,
      engine: 'cloud',
      ist: { load: istLoad(row), atl: num(row.fbATL_obs), ctl: num(row.fbCTL_obs), acwr: num(row.fbACWR_obs) },
      soll: soll ? { load: sl.load, atl: sl.atl, ctl: sl.ctl, acwr: sl.acwr } : {},
    });
    return r.model;
  }
  async function cloudErnaehrung(morgen) {
    const plan = (await holen({ mode: 'planJsonV2' })).plan || [];
    const pm = plan.find((p) => p.date === morgen);
    const last = pm ? pm.recommended_load_ess || pm.original_load_ess || 0 : 0;
    const sport = pm ? pm.training_1 || '' : '';
    const zone = pm ? pm.recommended_zone || '' : '';
    let base = {};
    try {
      base = (await holen({ mode: 'physBaseline' })).baseline || {};
    } catch {}
    let gewicht = 76;
    let grund = 2028;
    for (const [k, v] of Object.entries(base)) {
      if (k.includes('Weight') && gewicht === 76) gewicht = v;
      if (k.includes('Resting_Calories') && grund === 2028) grund = v;
    }
    const rows = (await holen({ mode: 'chartData', range: '42d', metrics: 'kcal_out,load_fb_day' })).data || [];
    const xy = rows.map((r) => [num(r.load_fb_day), num(r.kcal_out)]).filter(([x, y]) => x !== null && y !== null);
    let a = grund * 1.15;
    let b = 4.5;
    if (xy.length >= 10) {
      const xm = xy.reduce((s, p) => s + p[0], 0) / xy.length;
      const ym = xy.reduce((s, p) => s + p[1], 0) / xy.length;
      const nb = xy.reduce((s, p) => s + (p[0] - xm) * (p[1] - ym), 0);
      const db = xy.reduce((s, p) => s + (p[0] - xm) ** 2, 0);
      if (db !== 0) {
        b = nb / db;
        a = ym - b * xm;
      }
    }
    const kcal = Math.round((a + b * (last || 0)) / 10) * 10; // Ruhetag: eigener Grundbedarf aus der Regression
    const kg = gewicht;
    const pk = last < 80 ? 1.6 : last < 180 ? 1.8 : 2.0;
    const ck = last < 40 ? 3 : last < 120 ? 4.5 : last < 220 ? 6 : 7.5;
    const protein = Math.round(kg * pk);
    const carbs = Math.round(kg * ck);
    let fett = kcal > 0 && protein + carbs > 0 ? Math.round((kcal - 4 * protein - 4 * carbs) / 9) : 0;
    fett = Math.max(Math.round(0.8 * kg), fett);
    let bilanz = null;
    try {
      const rh = ((await holen({ mode: 'chartData', range: '7d', metrics: 'kcal_in,load_fb_day,protein_g,carb_g' })).data || []).find(
        (r) => r.date === heute(),
      );
      if (rh && num(rh.kcal_in) !== null) {
        const lh = num(rh.load_fb_day) || 0;
        bilanz = {
          kcal_in: num(rh.kcal_in),
          kcal_ziel: Math.round((a + b * lh) / 10) * 10,
          protein: num(rh.protein_g),
          protein_ziel: Math.round(kg * (lh < 80 ? 1.6 : lh < 180 ? 1.8 : 2.0)),
          kh: num(rh.carb_g),
          last: lh,
        };
      }
    } catch {}
    const user =
      `Ziele morgen: ${kcal} kcal · Protein ${protein} g · KH ${carbs} g · Fett ${fett} g\n` +
      `Training morgen: Last=${last}, Sport=${sport}, Zone=${zone}\nHeute Ist: ${JSON.stringify(bilanz)}\n` +
      `Vorlieben: Familie (2 Erwachsene, Kind 4 Jahre), mag Fleischgerichte, Kartoffelknödel, Sushi, Gerichte im HelloFresh-Stil, unter der Woche schnell (max. 30 min).\n` +
      `Aufgabe: JSON mit 'kurz' (1-2 Sätze), 'mahlzeiten' (4-5: wann, gericht, warum), 'tipp'.`;
    const r = await cloudChat(
      [
        { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, sachlich, keine Emojis, keine erfundenen Zahlen.' },
        { role: 'user', content: user },
      ],
      { json: true, model: 'sonar-pro', maxTokens: 1500 },
    );
    let j = {};
    const t = String(r.answer || '').replace(/^```(?:json)?\s*|\s*```$/g, '').trim();
    try {
      j = JSON.parse(t);
    } catch {
      const i = t.indexOf('{');
      const k = t.lastIndexOf('}');
      if (i >= 0 && k > i) j = JSON.parse(t.slice(i, k + 1));
      else throw new Error('Ernährungsplan: keine JSON-Antwort');
    }
    const zeilen = [`Ziel morgen (${last} ESS, ${sport}): ${kcal} kcal · Protein ${protein} g · KH ${carbs} g · Fett ${fett} g`];
    if (j.kurz) zeilen.push(j.kurz);
    for (const m of j.mahlzeiten || []) zeilen.push(`- ${m.wann || ''}: ${m.gericht || ''}`);
    if (j.tipp) zeilen.push(`Tipp: ${j.tipp}`);
    await senden('saveExtra', {
      kind: 'saveExtra',
      typ: 'ernaehrung',
      datum: morgen,
      text: zeilen.join('\n'),
      data: {
        ziele: { kcal, protein, carbs, fett },
        last,
        sport,
        zone,
        modell: { a: Math.round(a * 100) / 100, b: Math.round(b * 100) / 100 },
        heute: bilanz,
        mahlzeiten: j.mahlzeiten || [],
        kurz: j.kurz || '',
        tipp: j.tipp || '',
      },
      quelle: 'cloud:' + r.model,
    });
    return r.model;
  }

  // ---------- Erkennen, was fehlt ----------
  async function lage() {
    const tag = heute();
    const [chart, quellen, reviews, ern] = await Promise.all([
      holen({ mode: 'chartData', range: '3d', metrics: ACT_METRICS }),
      holen({ mode: 'briefingSources' }).catch(() => ({ sources: {} })),
      holen({ mode: 'activityReviews', limit: '20' }).catch(() => ({ reviews: [] })),
      holen({ mode: 'extras', typ: 'ernaehrung', n: '5' }).catch(() => ({ items: [] })),
    ]);
    const rows = chart.data || [];
    const src = quellen.sources || {};
    const revs = new Map((reviews.reviews || []).map((r) => [r.date, r]));
    const ernItems = new Map((ern.items || []).map((x) => [x.datum, x]));
    const out = [];
    const rh = rows.find((r) => r.date === tag);
    // Morgens: Readiness + Schlafscore von heute da
    if (rh && num(rh.Garmin_Training_Readiness) !== null && num(rh.sleep_score_0_100) !== null) {
      for (const [t, label] of MORGEN) {
        const s = src[t];
        const fertig = s && s.timestamp && heute(new Date(s.timestamp)) === tag;
        out.push({ key: `morgen:${tag}:${t}`, gruppe: 'morgens', art: t, label, datum: tag, fertig: !!fertig, engine: fertig ? s.engine : null, model: fertig ? s.model : null });
      }
    }
    // Nach Aktivität: heute und gestern
    for (const r of rows) {
      if (r.date !== tag && r.date !== plusTage(tag, -1)) continue;
      if (!istAktivitaet(r)) continue;
      const rv = revs.get(r.date);
      out.push({
        key: `aktivitaet:${r.date}`,
        gruppe: 'aktivitaet',
        art: 'aktivitaet',
        label: `Aktivitätskommentar ${r.date.slice(8, 10)}.${r.date.slice(5, 7)}. (${r.Sport_x || 'Einheit'})`,
        datum: r.date,
        fertig: !!rv,
        engine: rv ? rv.engine : null,
        model: rv ? rv.model : null,
        row: r,
      });
    }
    // Abends: Ernährungsplan für morgen, sobald abends eingetragen (kcal_in) und ab 17 Uhr
    const morgen = plusTage(tag, 1);
    if (rh && num(rh.kcal_in) !== null && stunde() >= 17) {
      const x = ernItems.get(morgen);
      const q = x ? String(x.quelle || x.source || '') : '';
      out.push({
        key: `ernaehrung:${morgen}`,
        gruppe: 'abends',
        art: 'ernaehrung',
        label: 'Ernährungsplan für morgen',
        datum: morgen,
        fertig: !!x,
        engine: x ? (q.startsWith('cloud') ? 'cloud' : 'local') : null,
        model: x ? q.replace(/^(mac|cloud):/, '') || null : null,
      });
    }
    // Plan-Soll für Aktivitätskommentare merken (wird nach dem Eintragen mit Ist überschrieben)
    if (rh && String(rh.activity_done || '').trim().toLowerCase() !== 'x') {
      const snap = lesen('texte_soll', {});
      snap[tag] = {
        load: num(rh.coachE_ESS_day), sport: rh.Sport_x, zone: rh.Zone, atl: num(rh.coachE_ATL_forecast), ctl: num(rh.coachE_CTL_forecast),
        acwr: num(rh.coachE_ACWR_forecast), te_ae: num(rh.Target_Aerobic_TE), te_an: num(rh.Target_Anaerobic_TE),
      };
      for (const k of Object.keys(snap).sort().slice(0, -7)) delete snap[k];
      schreiben('texte_soll', snap);
    }
    return out;
  }

  let laeuft = false;
  async function takt(zaehlen = true) {
    if (laeuft) return;
    laeuft = true;
    const z = zustand();
    try {
      const jetzt = Date.now();
      const items = await lage();
      const m = mac();
      const macFrisch = m.at && jetzt - m.at < MAC_FRISCH_MS;
      const ein = einstellung();
      for (const it of items) {
        const alt = z.items[it.key];
        const e = alt || { key: it.key, gruppe: it.gruppe, art: it.art, label: it.label, datum: it.datum, erkannt: jetzt, versuche: 0, phase: 'wartet', cloudVersuche: 0 };
        e.label = it.label;
        if (it.fertig) {
          if (e.phase !== 'fertig') e.fertigAt = jetzt;
          e.phase = 'fertig';
          e.engine = it.engine;
          e.model = it.model;
          e.fehler = null;
          z.items[it.key] = e;
          continue;
        }
        if (e.phase === 'fertig' && e.engine === 'cloud') {
          // Gerade per Cloud gespeichert, Sheet zeigt es evtl. noch nicht – nie doppelt erzeugen
          z.items[it.key] = e;
          continue;
        }
        if (e.phase === 'fertig') {
          // Text verschwunden (z. B. gelöscht) → neu warten
          Object.assign(e, { phase: 'wartet', versuche: 0, erkannt: jetzt, engine: null, model: null });
        }
        const lokalArbeit = macFrisch && m.arbeit === it.key && m.arbeitSeit && jetzt - m.arbeitSeit < MAC_ARBEIT_MAX_MS;
        const macBeschaeftigt = macFrisch && m.arbeit && m.arbeitSeit && jetzt - m.arbeitSeit < MAC_ARBEIT_MAX_MS;
        if (lokalArbeit) {
          e.phase = 'lokal';
        } else if (zaehlen && alt && e.phase !== 'cloud') {
          // Ein Versuch zählt nur, wenn der Mac nicht gerade an einem anderen Text rechnet
          if (!macBeschaeftigt) e.versuche = Math.min(VERSUCHE, (e.versuche || 0) + 1);
          e.phase = 'wartet';
          e.grund = !macFrisch ? 'Mac nicht erreichbar' : m.ollama === false ? 'Ollama auf dem Mac nicht bereit' : macBeschaeftigt ? 'Mac rechnet gerade einen anderen Text' : 'lokaler Text noch nicht da';
        }
        z.items[it.key] = e;
      }
      // Cloud-Ersatz: der Reihe nach (Plan vor Konsolidierung)
      if (zaehlen && ein.cloudErsatz) {
        const reihenfolge = ['status', 'plan', 'consolidated', 'aktivitaet', 'ernaehrung'];
        const offen = items
          .filter((it) => !it.fertig)
          .map((it) => [it, z.items[it.key]])
          .filter(([, e]) => e.phase === 'wartet' && e.versuche >= VERSUCHE && (e.cloudVersuche || 0) < CLOUD_VERSUCHE)
          .sort((a, b) => reihenfolge.indexOf(a[0].art) - reihenfolge.indexOf(b[0].art));
        for (const [it, e] of offen) {
          if (it.art === 'consolidated' && items.some((x) => x.art === 'plan' && !x.fertig && z.items[x.key]?.phase !== 'fertig')) continue;
          e.phase = 'cloud';
          e.cloudSeit = Date.now();
          e.cloudVersuche = (e.cloudVersuche || 0) + 1;
          schreiben('texte_zustand', z);
          try {
            let model;
            if (it.gruppe === 'morgens') model = await cloudBriefing(it.art);
            else if (it.art === 'aktivitaet') model = await cloudAktivitaet(it.row, lesen('texte_soll', {})[it.datum] || null);
            else model = await cloudErnaehrung(it.datum);
            Object.assign(e, { phase: 'fertig', engine: 'cloud', model, fertigAt: Date.now(), fehler: null });
          } catch (err) {
            Object.assign(e, { phase: 'wartet', fehler: String(err?.message || err).slice(0, 200) });
          }
          schreiben('texte_zustand', z);
        }
      }
      // Alte Einträge (älter als 3 Tage) aufräumen
      const grenze = plusTage(heute(), -3);
      for (const [k, e] of Object.entries(z.items)) if (e.datum < grenze) delete z.items[k];
      if (zaehlen) z.letzterTakt = jetzt;
      z.letzterCheck = jetzt;
      z.letzterFehler = null;
    } catch (err) {
      z.letzterFehler = String(err?.message || err).slice(0, 200);
      z.letzterCheck = Date.now();
    } finally {
      schreiben('texte_zustand', z);
      laeuft = false;
    }
  }

  setInterval(() => {
    if (Date.now() - (zustand().letzterTakt || 0) >= TAKT_MS * 0.95) takt(true);
  }, Math.min(60000, TAKT_MS / 4));
  setTimeout(() => takt(false), 5000);

  function statusAntwort() {
    const z = zustand();
    const m = mac();
    const jetzt = Date.now();
    const items = Object.values(z.items)
      .filter((e) => e.datum >= plusTage(heute(), -1))
      .sort((a, b) => (a.erkannt || 0) - (b.erkannt || 0));
    return {
      ok: true,
      now: jetzt,
      takt: { letzter: z.letzterTakt || null, check: z.letzterCheck || null, naechster: (z.letzterTakt || jetzt) + TAKT_MS, fehler: z.letzterFehler || null, versuche: VERSUCHE, minuten: TAKT_MS / 60000 },
      einstellung: einstellung(),
      mac: { at: m.at || null, erreichbar: !!(m.at && jetzt - m.at < MAC_FRISCH_MS), ollama: m.ollama ?? null, arbeit: m.arbeit || null, arbeitSeit: m.arbeitSeit || null },
      items,
    };
  }

  app.get('/api/texte/status', (req, res) => res.json(statusAntwort()));
  // Nach dem Absenden: sofort nachsehen (zählt nicht als Versuch)
  app.post('/api/texte/pruefen', async (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    await takt(false);
    res.json(statusAntwort());
  });
  app.post('/api/texte/einstellung', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const b = req.body || {};
    schreiben('texte_einstellung', { ...einstellung(), cloudErsatz: !!b.cloudErsatz });
    res.json(statusAntwort());
  });
  // Mac-Herzschlag (alle 5 min aus kira_runner tick) und Arbeitsmeldung je Text
  app.post('/api/texte/mac', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const b = req.body || {};
    const m = mac();
    m.at = Date.now();
    if (typeof b.ollama === 'boolean') m.ollama = b.ollama;
    if ('arbeit' in b) {
      m.arbeit = b.arbeit ? String(b.arbeit).slice(0, 80) : null;
      m.arbeitSeit = b.arbeit ? Date.now() : null;
    }
    schreiben('texte_mac', m);
    res.json({ ok: true });
  });
  // Mac fragt vor dem Rechnen: Hat die Cloud diesen Text schon übernommen?
  app.post('/api/texte/claim', (req, res) => {
    if (!originOk(req)) return res.status(403).json({ ok: false, error: 'Herkunft nicht erlaubt.' });
    const key = String((req.body || {}).key || '');
    const z = zustand();
    const e = z.items[key];
    const cloud = !!e && (e.phase === 'cloud' || (e.phase === 'fertig' && e.engine === 'cloud'));
    const m = mac();
    m.at = Date.now();
    if (!cloud) {
      m.arbeit = key;
      m.arbeitSeit = Date.now();
      if (e) {
        e.phase = 'lokal';
        schreiben('texte_zustand', z);
      }
    }
    schreiben('texte_mac', m);
    res.json({ ok: true, erlaubt: !cloud });
  });
}
