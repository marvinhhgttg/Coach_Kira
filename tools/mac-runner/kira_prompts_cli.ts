/**
 * Coach Kira · Prompt-CLI für den Mac. Gebaut zu kira_prompts.mjs (esbuild, ohne Abhängigkeiten).
 *   node kira_prompts.mjs supervisor   < {datenpaket aus supervisorJob.data}      → {system, prompt, schema}
 *   node kira_prompts.mjs briefing     < {type, daten aus briefingData}            → {format, messages}
 *   node kira_prompts.mjs telegram     < {frage, status, vorausblick, verlauf}     → {messages}
 *   node kira_prompts.mjs erfahrung                                                → {text, gruppen}
 * Alle drei Prompt-Arten bekommen zusätzlich das Erfahrungsmodell (lib/votumErfahrung.ts, ohne KI),
 * berechnet aus der ganzen Timeline (chartData). Ist die Timeline nicht erreichbar, bleibt der Prompt wie bisher.
 * Quelle der Texte: dashboard/src/lib/kiraPrompts.ts (gleich wie im Cockpit).
 */
import { supervisorPrompt, briefingMessages, telegramMessages } from '../../dashboard/src/lib/kiraPrompts';
import { erfahrungAusRohdaten, erfahrungText } from '../../dashboard/src/lib/votumErfahrung';
import { RECOVERY_FETCH_METRICS } from '../../dashboard/src/screens/RecoveryPanel';

const APPS_SCRIPT = 'https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec';

async function erfahrung(): Promise<{ text: string; gruppen: unknown[] }> {
  try {
    const ctl = new AbortController();
    const to = setTimeout(() => ctl.abort(), 30000);
    const url = APPS_SCRIPT + '?mode=chartData&range=9999d&metrics=' + RECOVERY_FETCH_METRICS.join(',');
    const j: any = await (await fetch(url, { redirect: 'follow', signal: ctl.signal })).json();
    clearTimeout(to);
    const e = erfahrungAusRohdaten(j.data || []);
    return { text: erfahrungText(e), gruppen: e.gruppen.filter((g) => g.n >= 8) };
  } catch {
    return { text: '', gruppen: [] };
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
  const erf = art === 'supervisor' || art === 'briefing' || art === 'telegram' || art === 'erfahrung' ? await erfahrung() : { text: '', gruppen: [] };
  if (art === 'supervisor') out = supervisorPrompt({ ...d, erfahrung: d.erfahrung ?? erf.text });
  else if (art === 'briefing') out = briefingMessages(d.type, { ...(d.daten || {}), erfahrung: (d.daten || {}).erfahrung ?? erf.text });
  else if (art === 'telegram') out = { messages: telegramMessages(String(d.frage || ''), { ...d, erfahrung: d.erfahrung ?? erf.text }) };
  else if (art === 'erfahrung') out = erf;
  else throw new Error('Unbekannte Art: ' + art + ' (supervisor | briefing | telegram | erfahrung)');
  process.stdout.write(JSON.stringify(out));
})().catch((e) => {
  process.stderr.write(String((e && e.message) || e));
  process.exit(1);
});
