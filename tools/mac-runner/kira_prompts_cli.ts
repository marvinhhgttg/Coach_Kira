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
import { kernVotumAusRohdaten, type KernVotumKi } from '../../dashboard/src/lib/kernVotumKi';

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
    const to = setTimeout(() => ctl.abort(), 30000);
    const url = APPS_SCRIPT + '?mode=chartData&range=9999d&metrics=' + RECOVERY_FETCH_METRICS.join(',');
    const j: any = await (await fetch(url, { redirect: 'follow', signal: ctl.signal })).json();
    clearTimeout(to);
    const e = erfahrungAusRohdaten(j.data || []);
    return { text: erfahrungText(e), gruppen: e.gruppen.filter((g) => g.n >= 8), kernVotum: kernVotumAusRohdaten(j.data || [], q) };
  } catch {
    return { text: '', gruppen: [], kernVotum: kernVotumAusRohdaten([], q) };
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
  else if (art === 'briefing') out = briefingMessages(d.type, { ...(d.daten || {}), erfahrung: (d.daten || {}).erfahrung ?? erf.text, kernVotum: kv });
  else if (art === 'telegram') out = { messages: telegramMessages(String(d.frage || ''), { ...d, erfahrung: d.erfahrung ?? erf.text, kernVotum: kv }) };
  else if (art === 'erfahrung') out = { text: erf.text, gruppen: erf.gruppen };
  else if (art === 'votum') out = kv;
  else throw new Error('Unbekannte Art: ' + art + ' (supervisor | briefing | telegram | erfahrung | votum)');
  process.stdout.write(JSON.stringify(out));
})().catch((e) => {
  process.stderr.write(String((e && e.message) || e));
  process.exit(1);
});
