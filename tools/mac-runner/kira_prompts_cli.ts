/**
 * Coach Kira · Prompt-CLI für den Mac. Gebaut zu kira_prompts.mjs (esbuild, ohne Abhängigkeiten).
 *   node kira_prompts.mjs supervisor   < {datenpaket aus supervisorJob.data}      → {system, prompt, schema}
 *   node kira_prompts.mjs briefing     < {type, daten aus briefingData}            → {format, messages}
 *   node kira_prompts.mjs telegram     < {frage, status, vorausblick, verlauf}     → {messages}
 * Quelle der Texte: dashboard/src/lib/kiraPrompts.ts (gleich wie im Cockpit).
 */
import { supervisorPrompt, briefingMessages, telegramMessages } from '../../dashboard/src/lib/kiraPrompts';

async function lesen(): Promise<string> {
  const teile: Buffer[] = [];
  for await (const t of process.stdin) teile.push(t as Buffer);
  return Buffer.concat(teile).toString('utf8');
}

(async () => {
  const art = process.argv[2];
  const d = JSON.parse((await lesen()) || '{}');
  let out: unknown;
  if (art === 'supervisor') out = supervisorPrompt(d);
  else if (art === 'briefing') out = briefingMessages(d.type, d.daten || {});
  else if (art === 'telegram') out = { messages: telegramMessages(String(d.frage || ''), d) };
  else throw new Error('Unbekannte Art: ' + art + ' (supervisor | briefing | telegram)');
  process.stdout.write(JSON.stringify(out));
})().catch((e) => {
  process.stderr.write(String((e && e.message) || e));
  process.exit(1);
});
