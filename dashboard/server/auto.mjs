/**
 * 24.11.2026: Coach Kira – Automatik (von Marc freigegeben: „Formular-Regel wird gelockert“).
 *
 * Mac (Mac-Schlüssel, Klasse 'mac' in zugang.mjs):
 *   POST /api/auto/eintrag          { typ: Morgens|Nach Aktivität|Abends, datum, werte }  → Apps Script submitCockpitEntry
 *                                   (dieselbe Verarbeitung wie das Formular; Doppel-Check gegen die Timeline)
 *   GET  /api/auto/sicherung        Server-Zustand (Kopplungen, Token, Einstellungen) für die Mac-Sicherung
 *   POST /api/auto/wiederherstellen { werte } → nur fehlende Schlüssel setzen (nach einem Neustart/Publish)
 *   GET  /api/auto/hinweise         fällige Hinweise (Morgenvotum, Befinden, Licht aus, Warnungen, Plan-Vorschlag)
 *   POST /api/auto/workout-morgen   morgiges Plan-Workout bauen und in die Garmin-Warteschlange legen
 * Cockpit (Sitzung):
 *   GET  /api/auto/status · POST /api/auto/einstellung { formular, hinweise, plancheck, workout }
 *   POST /api/auto/vorschlag        { aktion: annehmen|verwerfen, id }
 *
 * Die Timeline bleibt die Wahrheit: Einträge laufen über onFormSubmit (unverändert), vorhandene Werte werden nie
 * überschrieben (außer „Nach Aktivität“, wenn die Tageslast gestiegen ist – wie beim zweiten Formulareintrag).
 */

const ERLAUBT = {
  Morgens: ['garminATL', 'garminCTL', 'garminACWR', 'sleep_hours', 'sleep_score_0_100', 'rhr_bpm', 'hrv_status', 'hrv_threshholds', 'Garmin_Training_Readiness', 'Trainingszustand'],
  'Nach Aktivität': ['load_fb_day', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs', 'garminEnduranceScore', 'Sport_x', 'Zone', 'fb_TR_obs', 'Aerobic_TE', 'Anaerobic_TE'],
  Abends: ['kcal_in', 'kcal_out', 'deficit', 'carb_g', 'protein_g', 'fat_g', 'fiber_g'],
};
const TEXT = new Set(['hrv_threshholds', 'Trainingszustand', 'Sport_x', 'Zone']);
const SPORT_OK = new Set(['Run', 'Bike', 'Hike', 'Row', 'Swim', 'Skike', 'Ski']);
const ZUSTAND_OK = new Set(['Formaufbau', 'Formerhalt', 'Höchstform', 'Höchstleistung', 'Erholung', 'Unproduktiv', 'Formverlust', 'Überlastung', 'Ermüdet']);
const SICHERN = ['widget_geraete', 'dashboard_token', 'pin_hash', 'pin_salt', 'wochenplan', 'wochen_budget', 'geraete_einstellungen', 'texte_einstellung', 'auto_einstellung', 'auto_vorschlag', 'garmin_workouts', 'sitzungen', 'garmin_vorschlag'];
const STANDARD = { formular: true, hinweise: true, plancheck: true, workout: true, erfolg: true };

const heute = () => new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
const gestern = () => new Date(Date.now() - 86400000).toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
const uhrzeit = () => new Date().toLocaleTimeString('de-DE', { timeZone: 'Europe/Berlin', hour: '2-digit', minute: '2-digit', hour12: false });
const minuten = (hhmm) => { const [h, m] = String(hhmm).split(':').map(Number); return h * 60 + m; };
const num = (v) => { const n = typeof v === 'number' ? v : Number(String(v ?? '').replace(',', '.')); return Number.isFinite(n) ? n : null; };
const komma = (v, d = 0) => (d ? Number(v).toFixed(d) : String(Math.round(Number(v)))).replace('.', ',');
const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const tagKurz = (d) => { const x = new Date(d + 'T12:00:00Z'); return `${WT[x.getUTCDay()]} ${d.slice(8, 10)}.${d.slice(5, 7)}.`; };

export function setupAuto(app, { getSetting, setSetting, macOk, callAppsScript, scriptUrl, aktiv, workoutEinreihen, workouts }) {
  const workoutsLesen = () => { try { return workouts ? workouts() : []; } catch { return []; } };
  const lesen = (k, def) => { try { const v = getSetting(k); return v ? JSON.parse(v) : def; } catch { return def; } };
  const schreiben = (k, v) => setSetting(k, JSON.stringify(v));
  const einstellung = () => ({ ...STANDARD, ...lesen('auto_einstellung', {}) });
  const protokoll = (e) => {
    const xs = lesen('auto_log', []);
    xs.push({ at: Date.now(), ...e });
    schreiben('auto_log', xs.slice(-80));
  };
  const streng = (req) => macOk(req) === 'schluessel';

  async function zeile(datum, spalten) {
    const usp = new URLSearchParams({ mode: 'chartData', range: '5d', metrics: spalten.join(','), cb: String(Date.now()) });
    const r = await fetch(`${scriptUrl}?${usp}`, { redirect: 'follow' });
    const j = await r.json();
    return (j?.data || []).find((x) => String(x.date).slice(0, 10) === datum) || null;
  }

  // ---------- Werte prüfen und formatieren wie im Formular ----------
  function saeubern(typ, werte) {
    const out = {};
    for (const k of ERLAUBT[typ]) {
      const v = werte?.[k];
      if (v == null || v === '') continue;
      if (TEXT.has(k)) {
        const s = String(v).trim().slice(0, 40);
        if (k === 'Sport_x' && !SPORT_OK.has(s)) continue;
        if (k === 'Zone' && !/^Z[1-5](, Z[1-5])?$/.test(s)) continue;
        if (k === 'Trainingszustand' && !ZUSTAND_OK.has(s)) continue;
        if (k === 'hrv_threshholds' && !/^\d{1,3};\d{1,3}$/.test(s)) continue;
        out[k] = s;
        continue;
      }
      const n = num(v);
      if (n == null || n < -5000 || n > 20000) continue;
      const d = /ACWR/.test(k) ? 2 : /TE$|sleep_hours/.test(k) ? (k === 'sleep_hours' ? 2 : 1) : 0;
      out[k] = komma(n, d);
    }
    return out;
  }

  async function eintragen({ typ, datum, werte, quelle }) {
    if (!ERLAUBT[typ]) return { ok: false, error: 'Typ unbekannt.' };
    if (!einstellung().formular) return { ok: true, uebersprungen: true, grund: 'Automatik Formular ist aus.' };
    if (![heute(), gestern()].includes(datum) || (typ === 'Morgens' && datum !== heute())) return { ok: false, error: 'Datum nicht erlaubt.' };
    const v = saeubern(typ, werte);
    if (!Object.keys(v).length) return { ok: false, error: 'Keine gültigen Werte.' };
    // Doppel-Check gegen die Timeline
    const z = await zeile(datum, ERLAUBT[typ]);
    if (!z) return { ok: false, error: 'Timeline-Zeile nicht lesbar.' };
    if (typ === 'Morgens' && [z.Garmin_Training_Readiness, z.rhr_bpm, z.sleep_hours].some((x) => num(x) > 0)) return { ok: true, uebersprungen: true, grund: 'Morgenwerte stehen schon in der Timeline.' };
    if (typ === 'Abends' && num(z.kcal_in) > 0) return { ok: true, uebersprungen: true, grund: 'Abendwerte stehen schon in der Timeline.' };
    if (typ === 'Nach Aktivität') {
      const alt = num(z.load_fb_day) || 0, neu = num(v.load_fb_day) || 0;
      if (!(neu > 0)) return { ok: false, error: 'Tageslast fehlt.' };
      if (alt >= neu) return { ok: true, uebersprungen: true, grund: `Tageslast ${alt} steht schon in der Timeline.` };
    }
    const r = await callAppsScript({ action: 'submitCockpitEntry', type: typ, date: datum, values: JSON.stringify(v) });
    protokoll({ art: 'eintrag', typ, datum, ok: !!r?.ok, werte: v, quelle: quelle || 'mac', fehler: r?.ok ? undefined : String(r?.error || 'unbekannt').slice(0, 160) });
    if (r?.ok) try { globalThis.ckWidgetNeu?.(); } catch {}
    return r?.ok ? { ok: true, eingetragen: Object.keys(v), zeile: r.row ?? null } : { ok: false, error: r?.error || 'Apps Script lehnt ab.' };
  }

  // ---------- Hinweise (Uhr + iPhone) ----------
  function hinweise(a, kanal) {
    const e = einstellung();
    if (!e.hinweise || !a || a.datum !== heute()) return [];
    const jetzt = minuten(uhrzeit());
    const h = Math.floor(jetzt / 60);
    const xs = [];
    const v = a.votum, m = a.morgen, l = a.last;
    if (v?.vote && m?.eingetragen && !v.vorlaeufig && h >= 5 && h < 13) xs.push({ id: `${a.datum}-votum`, kanal: 'iphone', titel: `Kira: ${v.label}`, text: `${l?.sollText ? `Soll ${l.sollText} ESS. ` : ''}${(v.gruende || [])[0] || ''}`.trim() });
    if (v?.vote && v.vorlaeufig && h >= 5 && h < 11) xs.push({ id: `${a.datum}-vorl`, kanal: 'beide', titel: `Kira vorläufig: ${v.label}`, text: `Aus Garmin live, Sheet folgt. ${(v.gruende || [])[0] || ''}`.trim() });
    if (m && m.befinden == null && h >= 9 && h < 20) xs.push({ id: `${a.datum}-befinden`, kanal: 'beide', titel: 'Kira: Befinden?', text: 'Wie fühlst du dich heute (1–5)? Uhr: Menü → Befinden eintragen.' });
    if (a.schlaf?.lichtAus) {
      const la = minuten(a.schlaf.lichtAus);
      if (jetzt >= la - 20 && jetzt <= la + 60) xs.push({ id: `${a.datum}-licht`, kanal: 'beide', titel: `Licht aus um ${a.schlaf.lichtAus}`, text: `Schlafbedarf ${komma(a.schlaf.bedarfH, 1)} h bis ${a.schlaf.aufstehen}.` });
    }
    if (l?.acwrBand === 'hoch' && l.acwr != null && h >= 6 && h < 22) xs.push({ id: `${a.datum}-acwr`, kanal: 'beide', titel: `Kira: ACWR ${komma(l.acwr, 2)}`, text: 'Akute Last deutlich über der chronischen. Heute nichts draufpacken.' });
    if (m?.rhrDelta != null && m.rhrDelta >= 5 && h >= 6 && h < 22) xs.push({ id: `${a.datum}-rhr`, kanal: 'beide', titel: `Kira: Ruhepuls +${Math.round(m.rhrDelta)}`, text: `Ruhepuls ${m.rhr ?? '–'} liegt deutlich über deinem Schnitt. Locker bleiben, auf Infekt achten.` });
    const pv = a.planVorschlag;
    const st = lesen('auto_vorschlag', {});
    if (e.plancheck && pv && st.id !== pv.id && h >= 7 && h < 22) {
      const t = pv.aenderungen.map((x) => `${tagKurz(x.d)} ${x.alt}→${x.neu}`).join(', ');
      xs.push({ id: `plan-${pv.id}`.slice(0, 60), kanal: 'beide', titel: `Plan: ACWR ${komma(pv.maxAlt, 2)} ${tagKurz(pv.tagAlt)}`, text: `Vorschlag ${t} ESS → max ${komma(pv.maxNeu, 2)}. Bestätigen: Uhr-Menü oder Cockpit.` });
    }
    // 24.11.2026: Erfolgsmeldungen der Automatik (je Schritt einmal, höchstens 12 h alt)
    if (e.erfolg && h >= 6 && h < 23) {
      const seit = Date.now() - 12 * 3600 * 1000;
      const NAME = { Morgens: 'Morgenwerte', 'Nach Aktivität': 'Aktivität', Abends: 'Abendwerte' };
      for (const x of lesen('auto_log', []).filter((x) => x.at >= seit && x.ok && !x.uebersprungen)) {
        if (x.art === 'eintrag') {
          const w = x.werte || {};
          const kurz = x.typ === 'Morgens' ? [w.Garmin_Training_Readiness && `Bereitschaft ${w.Garmin_Training_Readiness}`, w.sleep_hours && `Schlaf ${w.sleep_hours} h`, w.hrv_status && `HRV ${w.hrv_status}`].filter(Boolean).join(' · ')
            : x.typ === 'Abends' ? [w.kcal_in && `${w.kcal_in} kcal`, w.deficit && `Defizit ${w.deficit}`, w.protein_g && `Eiweiß ${w.protein_g} g`].filter(Boolean).join(' · ')
            : [w.load_fb_day && `${w.load_fb_day} ESS`, w.Sport_x, w.Zone, w.fbACWR_obs && `ACWR ${w.fbACWR_obs}`].filter(Boolean).join(' · ');
          xs.push({ id: `ok-${x.at}`, kanal: 'beide', titel: `Kira: ${NAME[x.typ] || x.typ} eingetragen`, text: `${x.datum === a.datum ? '' : x.datum + ': '}${kurz || 'ins Sheet übernommen'}.` });
        } else if (x.art === 'plan' && !x.verworfen) xs.push({ id: `ok-${x.at}`, kanal: 'beide', titel: 'Kira: Plan angepasst', text: 'Der Plan-Vorschlag ist in der Timeline gespeichert.' });
        else if (x.art === 'wiederherstellung') xs.push({ id: `ok-${x.at}`, kanal: 'beide', titel: 'Kira: Kopplung wiederhergestellt', text: `Nach dem Server-Neustart zurückgeholt: ${(x.schluessel || []).join(', ')}.` });
      }
      for (const q of (workoutsLesen() || []).filter((q) => q.auto && q.doneAt && q.doneAt >= seit)) {
        xs.push(q.ok
          ? { id: `wo-${q.id}`.slice(0, 60), kanal: 'beide', titel: 'Kira: Workout im Garmin-Kalender', text: `${tagKurz(q.datum)} ${q.workout?.name || ''} – auf der Uhr unter Training → Workouts.` }
          : { id: `wo-${q.id}`.slice(0, 60), kanal: 'beide', titel: 'Kira: Workout fehlgeschlagen', text: `${tagKurz(q.datum)}: ${String(q.result || 'unbekannter Fehler').slice(0, 100)}` });
      }
    }
    return xs.filter((x) => kanal === 'mac' || x.kanal !== 'iphone').map((x) => ({ ...x, text: x.text.slice(0, 160) }));
  }
  globalThis.ckHinweise = hinweise;

  // ---------- Plan-Vorschlag übernehmen ----------
  async function vorschlagAnnehmen(id) {
    const a = await globalThis.ckWidgetAntwort?.(true);
    const pv = a?.planVorschlag;
    if (!pv || pv.id !== id) return { ok: false, error: 'Vorschlag ist nicht mehr aktuell.' };
    const r = await fetch(`${scriptUrl}?mode=planSimulationJson&cb=${Date.now()}`, { redirect: 'follow' });
    const sim = await r.json();
    if (!sim?.ok || !Array.isArray(sim.days)) return { ok: false, error: 'Plan-Simulation nicht lesbar.' };
    const neu = new Map(pv.aenderungen.map((x) => [x.d, x.neu]));
    const days = sim.days;
    const loads = days.map((d) => (neu.has(String(d.date).slice(0, 10)) && !d.locked ? neu.get(String(d.date).slice(0, 10)) : Number(d.load) || 0));
    const res = await callAppsScript({
      action: 'saveSimulatedPlan',
      loads: JSON.stringify(loads),
      teAe: JSON.stringify(days.map((d) => Number(d.te_ae) || 0)),
      teAn: JSON.stringify(days.map((d) => Number(d.te_an) || 0)),
      sports: JSON.stringify(days.map((d) => d.sport || '')),
      zones: JSON.stringify(days.map((d) => d.zone || '')),
      locks: JSON.stringify(days.map((d) => !!d.locked)),
    });
    protokoll({ art: 'plan', ok: !!res?.ok, id, fehler: res?.ok ? undefined : String(res?.error || '').slice(0, 160) });
    if (!res?.ok) return { ok: false, error: res?.error || 'Speichern fehlgeschlagen.' };
    schreiben('auto_vorschlag', { id, status: 'angenommen', at: Date.now() });
    try { globalThis.ckWidgetNeu?.(); } catch {}
    return { ok: true, text: `Plan angepasst: ${pv.aenderungen.map((x) => `${tagKurz(x.d)} ${x.neu}`).join(', ')} ESS.` };
  }
  function vorschlagVerwerfen(id) {
    schreiben('auto_vorschlag', { id, status: 'verworfen', at: Date.now() });
    protokoll({ art: 'plan', ok: true, id, verworfen: true });
    return { ok: true, text: 'Vorschlag verworfen.' };
  }
  globalThis.ckVorschlag = { annehmen: vorschlagAnnehmen, verwerfen: vorschlagVerwerfen, status: () => lesen('auto_vorschlag', {}) };

  // 25.11.2026: Plan direkt von der Uhr ändern – ESS eines Tages setzen, Ruhetag, zwei Tage tauschen (gesperrte Tage bleiben)
  async function planAendern(b) {
    const art = String(b.art || ''), d = String(b.d || '').slice(0, 10), d2 = String(b.d2 || '').slice(0, 10);
    const heute = new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d <= heute) return { ok: false, error: 'Nur künftige Tage lassen sich ändern.' };
    const r = await fetch(`${scriptUrl}?mode=planSimulationJson&cb=${Date.now()}`, { redirect: 'follow' });
    const sim = await r.json();
    if (!sim?.ok || !Array.isArray(sim.days)) return { ok: false, error: 'Plan-Simulation nicht lesbar.' };
    const days = sim.days.map((x) => ({ ...x }));
    const i = days.findIndex((x) => String(x.date).slice(0, 10) === d);
    if (i < 0) return { ok: false, error: 'Tag nicht im Plan.' };
    if (days[i].locked) return { ok: false, error: `${tagKurz(d)} ist gesperrt.` };
    let text = '';
    if (art === 'ess' || art === 'ruhe') {
      const ess = art === 'ruhe' ? 0 : Math.max(0, Math.min(600, Math.round(Number(b.ess) || 0)));
      text = `${tagKurz(d)}: ${Math.round(Number(days[i].load) || 0)} → ${ess} ESS`;
      days[i].load = ess;
      if (ess === 0) { days[i].sport = 'Off'; days[i].zone = 'Off'; days[i].te_ae = 0; days[i].te_an = 0; }
    } else if (art === 'tausch') {
      const j = days.findIndex((x) => String(x.date).slice(0, 10) === d2);
      if (j < 0 || d2 <= heute) return { ok: false, error: 'Zieltag nicht im Plan.' };
      if (days[j].locked) return { ok: false, error: `${tagKurz(d2)} ist gesperrt.` };
      for (const k of ['load', 'sport', 'zone', 'te_ae', 'te_an']) { const tmp = days[i][k]; days[i][k] = days[j][k]; days[j][k] = tmp; }
      text = `${tagKurz(d)} ↔ ${tagKurz(d2)} getauscht`;
    } else return { ok: false, error: 'Aktion unbekannt.' };
    const res = await callAppsScript({
      action: 'saveSimulatedPlan',
      loads: JSON.stringify(days.map((x) => Number(x.load) || 0)),
      teAe: JSON.stringify(days.map((x) => Number(x.te_ae) || 0)),
      teAn: JSON.stringify(days.map((x) => Number(x.te_an) || 0)),
      sports: JSON.stringify(days.map((x) => x.sport || '')),
      zones: JSON.stringify(days.map((x) => x.zone || '')),
      locks: JSON.stringify(days.map((x) => !!x.locked)),
    });
    protokoll({ art: 'plan_uhr', ok: !!res?.ok, text, fehler: res?.ok ? undefined : String(res?.error || '').slice(0, 160) });
    if (!res?.ok) return { ok: false, error: res?.error || 'Speichern fehlgeschlagen.' };
    try { globalThis.ckWidgetNeu?.(); } catch {}
    return { ok: true, text: `Plan gespeichert: ${text}.` };
  }
  globalThis.ckPlanAendern = planAendern;

  // ---------- Workout für morgen bauen ----------
  let faktoren = null;
  async function faktorenLaden() {
    if (faktoren && Date.now() - faktoren.at < 86400000) return faktoren;
    const r = await fetch(`${scriptUrl}?mode=loadFactors&cb=${Date.now()}`, { redirect: 'follow' });
    const j = await r.json();
    const f = {};
    for (const x of j?.load || []) { f[x.ActivityType] = f[x.ActivityType] || {}; f[x.ActivityType][x.Zone] = Number(x.Factor); }
    let lthr = 170;
    try {
      const p = await (await fetch(`${scriptUrl}?mode=physBaseline&cb=${Date.now()}`, { redirect: 'follow' })).json();
      const b = p?.baseline || {};
      const k = Object.keys(b).find((x) => x.toLowerCase().startsWith('laktatschwelle'));
      const v = k ? num(b[k]) : null;
      if (v && v > 100) lthr = v;
    } catch {}
    faktoren = { at: Date.now(), f, lthr };
    return faktoren;
  }
  function workoutBauen(tag, fk) {
    const s = String(tag.sport || '');
    const sport = s === 'Run' ? 'running' : s === 'Bike' ? 'cycling' : s === 'Hike' ? 'hiking' : 'training';
    const key = s === 'Run' || s === 'Row' ? 'Run' : s === 'Bike' ? 'Bike' : 'Other';
    const zs = String(tag.zone || '').match(/Z[1-5]/g) || ['Z2'];
    const haupt = zs[0], zweit = zs[1] && zs[1] > zs[0] ? zs[1] : null;
    const F = (z) => fk.f[key]?.[z] || fk.f[key]?.Default || fk.f.Other?.Default || 80;
    const f = [[0, 0.85], [0.85, 0.89], [0.9, 0.94], [0.95, 0.99], [1.0, 1.06]];
    const fb = [[0, 0.81], [0.81, 0.89], [0.9, 0.93], [0.94, 0.99], [1.0, 1.06]];
    const hf = (z) => { const i = Number(z[1]) - 1; const t = (sport === 'cycling' ? fb : f)[i]; return { hrLow: Math.round(fk.lthr * t[0]) || 100, hrHigh: Math.round(fk.lthr * t[1]) }; };
    const ess = Number(tag.ess) || 0;
    const rest = Math.max(0, ess - (15 / 60) * F('Z1'));
    let hauptMin = Math.max(10, Math.round((rest / F(haupt)) * 60));
    let zweitMin = 0;
    if (zweit) { zweitMin = Math.round(hauptMin * 0.2); hauptMin -= zweitMin; zweitMin = Math.max(5, Math.round(zweitMin * F(haupt) / F(zweit))); }
    const steps = [
      { name: 'Einrollen', minutes: 10, intensity: 'warmup', ...hf('Z1') },
      { name: `Hauptteil ${haupt}`, minutes: hauptMin, intensity: 'active', ...hf(haupt) },
      ...(zweit ? [{ name: `Schluss ${zweit}`, minutes: zweitMin, intensity: 'active', ...hf(zweit) }] : []),
      { name: 'Ausrollen', minutes: 5, intensity: 'cooldown', ...hf('Z1') },
    ];
    const name = `Kira ${s || 'Training'} ${zs.slice(0, 2).join('/')} ${Math.round(ess)} ESS`;
    return { name, sport, blocks: [{ repeat: 1, steps }], beschreibung: `Automatisch aus Kiras Plan (${tagKurz(tag.d)}): ${Math.round(ess)} ESS ${tag.zone || ''}. HF nach Laktatschwelle ${fk.lthr}.`, auto: true };
  }

  if (!aktiv) return { hinweise };

  // ---------- Routen: Mac ----------
  app.post('/api/auto/eintrag', async (req, res) => {
    if (!streng(req)) return res.status(403).json({ ok: false, error: 'mac_key_required' });
    const b = req.body || {};
    try {
      const r = await eintragen({ typ: String(b.typ || ''), datum: String(b.datum || ''), werte: b.werte || {}, quelle: String(b.quelle || 'mac').slice(0, 20) });
      if (r.uebersprungen) protokoll({ art: 'eintrag', typ: b.typ, datum: b.datum, ok: true, uebersprungen: r.grund });
      res.json(r);
    } catch (e) {
      protokoll({ art: 'eintrag', typ: b.typ, datum: b.datum, ok: false, fehler: String(e.message || e).slice(0, 160) });
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });
  app.get('/api/auto/sicherung', (req, res) => {
    if (!streng(req)) return res.status(403).json({ ok: false, error: 'mac_key_required' });
    const werte = {};
    for (const k of SICHERN) { const v = getSetting(k); if (v) werte[k] = v; }
    schreiben('auto_puls', { at: Date.now(), version: String(req.query.v || '').slice(0, 20) });
    res.json({ ok: true, werte, geraete: lesen('widget_geraete', []).length, start: globalThis.ckStart || null });
  });
  app.post('/api/auto/wiederherstellen', (req, res) => {
    if (!streng(req)) return res.status(403).json({ ok: false, error: 'mac_key_required' });
    const w = (req.body || {}).werte || {};
    const neu = [];
    for (const k of SICHERN) {
      const v = w[k];
      if (typeof v !== 'string' || !v || v.length > 300000) continue;
      if (k === 'widget_geraete') {
        let alt = [], dazu = [];
        try { alt = JSON.parse(getSetting(k) || '[]'); dazu = JSON.parse(v); } catch { continue; }
        if (!Array.isArray(alt) || !Array.isArray(dazu)) continue;
        const ids = new Set(alt.map((g) => g.id));
        const plus = dazu.filter((g) => g && g.id && g.hash && !ids.has(g.id));
        if (plus.length) { setSetting(k, JSON.stringify([...alt, ...plus].slice(-10))); neu.push(`${k}+${plus.length}`); }
        continue;
      }
      if (k === 'sitzungen') {
        // Anmeldungen zusammenführen (nur Hash + Ablauf), abgelaufene fallen weg
        let alt = [], dazu = [];
        try { alt = JSON.parse(getSetting(k) || '[]'); dazu = JSON.parse(v); } catch { continue; }
        if (!Array.isArray(alt) || !Array.isArray(dazu)) continue;
        const jetzt = Date.now();
        const hs = new Set(alt.map((x) => x && x.h));
        const plus = dazu.filter((x) => x && typeof x.h === 'string' && x.h.length === 64 && Number(x.e) > jetzt && !hs.has(x.h));
        if (plus.length) { setSetting(k, JSON.stringify([...alt, ...plus].slice(-50))); neu.push(`${k}+${plus.length}`); }
        continue;
      }
      if (!getSetting(k)) { setSetting(k, v); neu.push(k); }
    }
    if (neu.length) protokoll({ art: 'wiederherstellung', ok: true, schluessel: neu });
    res.json({ ok: true, wiederhergestellt: neu });
  });
  app.get('/api/auto/hinweise', async (req, res) => {
    if (!streng(req)) return res.status(403).json({ ok: false, error: 'mac_key_required' });
    try {
      const a = await globalThis.ckWidgetAntwort?.(false);
      const push = lesen('widget_geraete', []).filter((g) => g?.push?.token).map((g) => ({ token: g.push.token, umgebung: g.push.umgebung, name: g.name || g.geraet || '' }));
      res.json({ ok: true, push, hinweise: hinweise(a, 'mac'), votum: a?.votum ? { vote: a.votum.vote, vorlaeufig: !!a.votum.vorlaeufig } : null });
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });
  app.post('/api/auto/workout-morgen', async (req, res) => {
    if (!streng(req)) return res.status(403).json({ ok: false, error: 'mac_key_required' });
    try {
      if (!einstellung().workout) return res.json({ ok: true, uebersprungen: 'Automatik Workout ist aus.' });
      const a = await globalThis.ckWidgetAntwort?.(false);
      const tag = (a?.naechste || [])[0];
      if (!tag) return res.json({ ok: true, uebersprungen: 'Kein Plan für morgen.' });
      if (!(Number(tag.ess) > 0) || /off|ruhe/i.test(String(tag.sport))) return res.json({ ok: true, uebersprungen: `Morgen Ruhetag (${tag.d}).` });
      if (workouts().some((q) => q.datum === tag.d && (q.ok || !q.doneAt))) return res.json({ ok: true, uebersprungen: `Für ${tag.d} liegt schon ein Workout vor.` });
      const w = workoutBauen(tag, await faktorenLaden());
      const auftrag = workoutEinreihen(tag.d, w);
      protokoll({ art: 'workout', ok: true, datum: tag.d, name: w.name });
      res.json({ ok: true, auftrag: { id: auftrag.id, datum: tag.d, name: w.name } });
    } catch (e) {
      protokoll({ art: 'workout', ok: false, fehler: String(e.message || e).slice(0, 160) });
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });

  // ---------- Routen: Cockpit ----------
  app.get('/api/auto/status', async (_req, res) => {
    let a = null;
    try { a = await globalThis.ckWidgetAntwort?.(false); } catch {}
    res.json({
      ok: true, einstellung: einstellung(), log: lesen('auto_log', []).slice(-40).reverse(), puls: lesen('auto_puls', null),
      vorschlag: a?.planVorschlag || null, vorschlagStatus: lesen('auto_vorschlag', {}), hinweise: hinweise(a, 'mac'),
      votum: a?.votum || null, morgen: a?.morgen || null, schlaf: a?.schlaf || null,
    });
  });
  app.post('/api/auto/einstellung', (req, res) => {
    const b = req.body || {};
    const e = einstellung();
    for (const k of Object.keys(STANDARD)) if (typeof b[k] === 'boolean') e[k] = b[k];
    schreiben('auto_einstellung', e);
    res.json({ ok: true, einstellung: e });
  });
  app.post('/api/auto/vorschlag', async (req, res) => {
    const b = req.body || {};
    try {
      const r = b.aktion === 'annehmen' ? await vorschlagAnnehmen(String(b.id || '')) : b.aktion === 'verwerfen' ? vorschlagVerwerfen(String(b.id || '')) : { ok: false, error: 'Aktion unbekannt.' };
      res.status(r.ok ? 200 : 409).json(r);
    } catch (e) {
      res.status(502).json({ ok: false, error: String(e.message || e) });
    }
  });
  return { hinweise };
}
