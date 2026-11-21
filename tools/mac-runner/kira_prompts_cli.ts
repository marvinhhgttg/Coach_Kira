/**
 * Coach Kira · Prompt-CLI für den Mac. Gebaut zu kira_prompts.mjs (esbuild, ohne Abhängigkeiten).
 *   node kira_prompts.mjs supervisor   < {datenpaket aus supervisorJob.data}      → {system, prompt, schema}
 *   node kira_prompts.mjs briefing     < {type, daten aus briefingData}            → {format, messages}
 *   node kira_prompts.mjs telegram     < {frage, status, vorausblick, verlauf}     → {messages}
 *   node kira_prompts.mjs erfahrung                                                → {text, gruppen}
 *   node kira_prompts.mjs votum                                                    → maßgebliches Tagesvotum für die KI-Texte
 * Alle drei Prompt-Arten bekommen zusätzlich das Erfahrungsmodell (lib/votumErfahrung.ts, ohne KI),
 * berechnet aus der ganzen Timeline (chartData). Ist die Timeline nicht erreichbar, bleibt der Prompt wie bisher.
 * Quelle der Texte: dashboard/src/lib/kiraPrompts.ts (gleich wie im Cockpit).
 */
import { supervisorPrompt, briefingMessages, telegramMessages } from '../../dashboard/src/lib/kiraPrompts';
import { erfahrungAusRohdaten, erfahrungText } from '../../dashboard/src/lib/votumErfahrung';
import { RECOVERY_FETCH_METRICS } from '../../dashboard/src/screens/RecoveryPanel';
import { kernVotumAusRohdaten, heuteBerlin, type KernVotumKi } from '../../dashboard/src/lib/kernVotumKi';
import { langKontext, parseLangEinstellung, type LangKontext } from '../../dashboard/src/lib/esLang';

const COCKPIT_STATUS = 'https://coachkira.pplx.app/port/5000/api/texte/status';
const APPS_SCRIPT = 'https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec';

/** Schalter Technik › Kira-Texte (votumQuelle). Server übergibt ihn im Paket, der Mac fragt den Cockpit-Server. Standard: kern. */
async function schalter(d: any): Promise<string> {
  if (d && (d.votumQuelle === 'kern' || d.votumQuelle === 'ampel')) return d.votumQuelle;
  try {
    const ctl = new AbortController();
    const to = setTimeout(() => ctl.abort(), 8000);
    const j: any = await (await fetch(COCKPIT_STATUS + '?cb=' + Date.now(), { signal: ctl.signal })).json();
    clearTimeout(to);
    const q = j && j.einstellung && j.einstellung.votumQuelle;
    return q === 'ampel' ? 'ampel' : 'kern';
  } catch {
    return 'kern';
  }
}

async function erfahrung(d: any): Promise<{ text: string; gruppen: unknown[]; kernVotum: KernVotumKi }> {
  const q = await schalter(d);
  try {
    const ctl = new AbortController();
    const to = setTimeout(() => ctl.abort(), 90000);
    // Etappe 6: zuerst timeline (Wahrheit) über den Cockpit-Server, Rückfall Apps Script (KK_TIMELINE)
    const q = '?range=9999d&metrics=' + RECOVERY_FETCH_METRICS.join(',');
    const z = process.env.CK_ZEITREIHE_URL === 'aus' ? null : process.env.CK_ZEITREIHE_URL || 'https://coachkira.pplx.app/port/5000/api/rechenkern/chartData';
    let j: any = null;
    if (z) {
      try { j = await (await fetch(z + q, { signal: ctl.signal })).json(); } catch { j = null; }
      if (!j || !j.ok || !Array.isArray(j.data) || !j.data.length) j = null;
    }
    if (!j) j = await (await fetch(APPS_SCRIPT + '?mode=chartData' + q.replace('?', '&'), { redirect: 'follow', signal: ctl.signal })).json();
    clearTimeout(to);
    const e = erfahrungAusRohdaten(j.data || []);
    return { text: erfahrungText(e), gruppen: e.gruppen.filter((g) => g.n >= 8), kernVotum: kernVotumAusRohdaten(j.data || [], q) };
  } catch {
    return { text: '', gruppen: [], kernVotum: kernVotumAusRohdaten([], q) };
  }
}

/** Lange Einheiten fürs Plan-Briefing: Einstellung (Cockpit-Server), Score-Verlauf (timeline), Garmin-Dauern (Apps Script). */
async function langeEinheiten(): Promise<LangKontext | null> {
  const mitZeit = async (url: string, ms: number) => {
    const ctl = new AbortController();
    const to = setTimeout(() => ctl.abort(), ms);
    try { return await (await fetch(url, { redirect: 'follow', signal: ctl.signal })).json(); } finally { clearTimeout(to); }
  };
  try {
    let ein: any = {};
    try { ein = ((await mitZeit(COCKPIT_STATUS + '?cb=' + Date.now(), 8000)) || {}).einstellung || {}; } catch { ein = {}; }
    const le = ein.langeEinheiten || {};
    const e = parseLangEinstellung(le.tage ?? null, le.sehr ?? null);
    const z = process.env.CK_ZEITREIHE_URL === 'aus' ? null : process.env.CK_ZEITREIHE_URL || 'https://coachkira.pplx.app/port/5000/api/rechenkern/chartData';
    const q = '?range=9999d&metrics=load_fb_day,garminEnduranceScore';
    let j: any = null;
    if (z) { try { j = await mitZeit(z + q, 60000); } catch { j = null; } }
    if (!j || !j.ok || !Array.isArray(j.data)) j = await mitZeit(APPS_SCRIPT + '?mode=chartData' + q.replace('?', '&'), 90000);
    const num = (x: any) => { const n = parseFloat(String(x ?? '').replace(',', '.')); return Number.isFinite(n) ? n : null; };
    const heute = heuteBerlin();
    const t = (j.data || []).filter((r: any) => String(r.date || '').slice(0, 10) <= heute).map((r: any) => ({ date: String(r.date).slice(0, 10), load: num(r.load_fb_day), es: num(r.garminEnduranceScore) }));
    const g: any = await mitZeit(APPS_SCRIPT + '?mode=extras&typ=garmin_tag&n=1000', 90000);
    const m = new Map<string, any>();
    for (const it of (g && g.items) || []) if (it && it.data && typeof it.data === 'object') m.set(it.datum, it.data);
    if (!t.length || !m.size) return null;
    return langKontext(t as any, m as any, e, heute);
  } catch {
    return null;
  }
}

async function lesen(): Promise<string> {
  const teile: Buffer[] = [];
  for await (const t of process.stdin) teile.push(t as Buffer);
  return Buffer.concat(teile).toString('utf8');
}

(async () => {
  const art = process.argv[2];
  const d = JSON.parse((await lesen()) || '{}');
  let out: unknown;
  const mitDaten = ['supervisor', 'briefing', 'telegram', 'erfahrung', 'votum'].includes(art);
  const erf = mitDaten ? await erfahrung(d) : { text: '', gruppen: [], kernVotum: null as any };
  const kv = d.kernVotum ?? erf.kernVotum;
  if (art === 'supervisor') out = supervisorPrompt({ ...d, erfahrung: d.erfahrung ?? erf.text, kernVotum: kv });
  else if (art === 'briefing') {
    const lk = d.type === 'plan' ? ((d.daten || {}).langeEinheiten ?? (await langeEinheiten())) : undefined;
    out = briefingMessages(d.type, { ...(d.daten || {}), erfahrung: (d.daten || {}).erfahrung ?? erf.text, kernVotum: kv, langeEinheiten: lk });
  }
  else if (art === 'lange') out = await langeEinheiten();
  else if (art === 'telegram') out = { messages: telegramMessages(String(d.frage || ''), { ...d, erfahrung: d.erfahrung ?? erf.text, kernVotum: kv }) };
  else if (art === 'erfahrung') out = { text: erf.text, gruppen: erf.gruppen };
  else if (art === 'votum') out = kv;
  else throw new Error('Unbekannte Art: ' + art + ' (supervisor | briefing | telegram | erfahrung | votum | lange)');
  process.stdout.write(JSON.stringify(out));
})().catch((e) => {
  process.stderr.write(String((e && e.message) || e));
  process.exit(1);
});
