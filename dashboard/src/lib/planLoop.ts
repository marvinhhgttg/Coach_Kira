/**
 * Plan-Vorschlag mit Prüfschleife: coach-kira schlägt 14 Tage vor, die bestehende
 * Simulation (ATL/CTL/ACWR/Monotonie) und planChecks prüfen, Verstöße gehen als
 * Rückmeldung zurück an Kira – höchstens MAX_ROUNDS Runden. Gerechnet wird lokal.
 */
import { localChat, parseJsonLoose, type LlmMsg } from './localLLM';
import { MONOTONIE_REGEL_TEXT, monotonyPlanViolations } from './monotony';

export type LoopDay = { date: string; day?: string; load: number; zone: string; sport: string; locked: boolean };
export type LoopSimRow = { sim_load: number; acwr: number; monotony: number; kei?: number; ctl: number; atl: number };
export type LoopCheck = { issues: string[]; level: 'ok' | 'warn' | 'bad'; cap?: string };
export type LoopRound = { loads: number[]; zones: string[]; why: string[]; summary: string; violations: string[]; maxAcwr: number; restDays: number; finalCtl: number; sports?: string[]; training?: string[]; alt?: string[]; langH?: (number | null)[]; esEffekt?: (number | null)[] };
export type Llm = (msgs: LlmMsg[]) => Promise<string>;
export const localLlm: Llm = (msgs) => localChat(msgs, { format: 'json', numCtx: 16384 });

export const MAX_ROUNDS = 3;
const ACWR_MAX = 1.3;
const ACWR_MIN = 0.8;
const CTL_MAX_LOSS = 0.05; // höchstens 5 % CTL-Verlust bis Tag 14 gegenüber heute

/** Langer Tag aus Marcs Vorgabe (Endurance-Score-Hebel); Kira darf ihn um einen Tag verschieben. */
export type LangTermin = { date: string; art: string };
const HART = /^Z[3-5]$/i;
/**
 * Regeln für lange Einheiten (21.11.2026):
 * – jeder lange Termin braucht eine lange Einheit am Termin oder einen Tag davor/danach,
 * – Vortag und Folgetag einer langen Einheit locker (kein Z3–Z5), damit sie ruhig und frisch gelingt.
 */
export function langRegeln(days: LoopDay[], loads: number[], zones: string[], langH: (number | null)[], termine: LangTermin[]): string[] {
  const v: string[] = [];
  const tag = (i: number) => `${days[i].day ? days[i].day + ' ' : ''}${days[i].date}`;
  const idx = new Map(days.map((d, i) => [d.date, i]));
  for (const t of termine) {
    if (t.art === 'ziel' || t.art === 'taper') continue;
    const i = idx.get(t.date);
    if (i == null || i === 0) continue;
    const kandidaten = [i, i - 1, i + 1].filter((k) => k >= 1 && k < days.length);
    if (!kandidaten.some((k) => langH[k] != null)) {
      const frei = kandidaten.filter((k) => !days[k].locked).map(tag);
      v.push(`${tag(i)}: lange Einheit laut Vorgabe fehlt – am Termin oder einen Tag davor/danach einplanen${frei.length ? ` (frei: ${frei.join(', ')})` : ''}`);
    }
  }
  langH.forEach((h, i) => {
    if (h == null || i === 0) return;
    for (const k of [i - 1, i + 1]) {
      if (k < 0 || k >= days.length || langH[k] != null) continue;
      if (HART.test(zones[k] || '') && loads[k] > 0)
        v.push(`${tag(k)}: ${zones[k]} direkt ${k < i ? 'vor' : 'nach'} der langen Einheit am ${tag(i)} – ${k < i ? 'Vortag' : 'Folgetag'} locker (Z1–Z2) oder frei, sonst die lange Einheit um einen Tag verschieben${days[k].locked ? '' : ' oder die harte Einheit verlegen'}`);
    }
  });
  return v;
}
const MIN_REST = 2;

const f = (v: number, d = 2) => v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });

export function evaluate(
  days: LoopDay[],
  loads: number[],
  zones: string[],
  simulate: (l: number[], z: string[]) => LoopSimRow[],
  check: (rows: LoopSimRow[]) => LoopCheck[],
  lang?: { langH: (number | null)[]; termine: LangTermin[] },
) {
  const rows = simulate(loads, zones);
  const checks = check(rows);
  const v: string[] = [];
  rows.forEach((r, i) => {
    if (i === 0) return; // heute = Anker aus der timeline
    const tag = `${days[i].day ? days[i].day + ' ' : ''}${days[i].date}`;
    if (r.acwr > ACWR_MAX) v.push(`${tag}: ACWR ${f(r.acwr)} > ${f(ACWR_MAX, 1)}`);
    if (i >= 2 && r.acwr < ACWR_MIN) v.push(`${tag}: ACWR ${f(r.acwr)} < ${f(ACWR_MIN, 1)} (Formverlust)`);
    for (const iss of checks[i]?.issues || []) if (!iss.startsWith('ACWR')) v.push(`${tag}: ${iss}`);
  });
  // Monotonie-Regel: Verlauf statt Einzelwert (lib/monotony.ts)
  for (const m of monotonyPlanViolations(
    rows.map((r, i) => ({ label: `${days[i].day ? days[i].day + ' ' : ''}${days[i].date}`, load: loads[i], mono: i === 0 ? null : r.monotony, locked: days[i].locked || i === 0 })),
  ))
    v.push(m);
  if (lang) v.push(...langRegeln(days, loads, zones, lang.langH, lang.termine));
  const restDays = loads.slice(1).filter((x) => x <= 0).length;
  if (restDays < MIN_REST) v.push(`Nur ${restDays} Ruhetag(e) mit 0 ESS in 13 Tagen (mindestens ${MIN_REST})`);
  const maxAcwr = Math.max(...rows.slice(1).map((r) => r.acwr));
  const ctl0 = rows[0]?.ctl ?? 0;
  const ctlEnd = rows[rows.length - 1]?.ctl ?? 0;
  if (ctl0 > 0 && ctlEnd < ctl0 * (1 - CTL_MAX_LOSS))
    v.push(`CTL fällt bis Tag 14 von ${f(ctl0, 0)} auf ${f(ctlEnd, 0)} (mehr als ${CTL_MAX_LOSS * 100} % Verlust) – mehr Last an den Trainingstagen`);
  return { rows, checks, violations: v, maxAcwr, restDays, finalCtl: rows[rows.length - 1]?.ctl ?? 0 };
}

function planTable(days: LoopDay[], loads: number[], zones: string[], rows: LoopSimRow[], checks: LoopCheck[], sports?: string[]) {
  return days
    .map((d, i) => `${d.date} ${d.day || ''} | ${d.locked || i === 0 ? 'FEST' : 'frei'} | ${loads[i]} ESS | ${zones[i] || '-'}${sports ? ` | ${sports[i] || '-'}` : ''} | ACWR ${f(rows[i]?.acwr ?? 0)} | Mono ${f(rows[i]?.monotony ?? 0)}${checks[i]?.cap ? ` | erlaubt max. ${checks[i].cap}` : ''}`)
    .join('\n');
}

export async function runPlanLoop(args: {
  days: LoopDay[];
  simulate: (l: number[], z: string[]) => LoopSimRow[];
  check: (rows: LoopSimRow[]) => LoopCheck[];
  context: string;
  onProgress?: (msg: string) => void;
  /** Modellaufruf – Standard: coach-kira lokal. */
  llm?: Llm;
  /** Wochenplan-Modus: zusätzlich Sport, Training und Alternative je Tag. */
  sportsAllowed?: string[];
  /** Lange Termine aus Marcs Vorgabe – werden mitgeprüft (Wochenplan). */
  langTermine?: LangTermin[];
}): Promise<{ rounds: LoopRound[]; best: LoopRound }> {
  const { days, simulate, check, context, onProgress } = args;
  const llm = args.llm || localLlm;
  const wp = !!args.sportsAllowed?.length;
  const base = evaluate(days, days.map((d) => d.load), days.map((d) => d.zone), simulate, check);
  const system =
    'Du bist Coach Kira, Sportwissenschaftlerin. Du planst Marcs nächste 14 Tage konservativ und regelkonform. Die summary sprichst du Marc per Du an und beschreibst die Idee des Plans. ' +
    (wp
      ? 'Antworte ausschließlich als JSON: {"summary": "2-3 Sätze auf Deutsch", "days": [{"date": "YYYY-MM-DD", "load": Zahl 0-300, "zone": "Off|Z1|Z2|Z3|Z4|Z5", "sport": "' +
        args.sportsAllowed!.join('|') +
        '", "training": "Einheit in max. 6 Wörtern", "alt": "Alternative in max. 5 Wörtern", "why": "max. 10 Wörter"}]} mit genau den 14 Tagen. Ruhetage: load 0, zone Off, sport Off. An Tagen mit langer Einheit zusätzlich "lang_h": Stunden und "es_effekt": Punkte Endurance Score.'
      : 'Antworte ausschließlich als JSON: {"summary": "2-3 Sätze auf Deutsch", "days": [{"date": "YYYY-MM-DD", "load": Zahl 0-300, "zone": "Off|Z1|Z2|Z3|Z4|Z5", "why": "kurz"}]} mit genau den 14 Tagen.');
  const rules =
    `REGELN (werden nach deinem Vorschlag automatisch nachgerechnet):\n` +
    `- Tage mit FEST nicht verändern (gleiche ESS und Zone).\n` +
    `- ACWR an jedem Tag zwischen ${f(ACWR_MIN, 1)} und ${f(ACWR_MAX, 1)}.\n` +
    `- ${MONOTONIE_REGEL_TEXT.replace(/\n/g, '\n- ')}\n` +
    `- Geprüft wird: höchstens 4 Tage in Folge ohne Entlastung (Entlastung = lockerer oder freier Tag mit höchstens 40 ESS bzw. 60 % der üblichen Tageslast); liegt die Monotonie 2 Tage in Folge über 2,0 und gab es 4 fordernde Tage, folgt ein sehr lockerer oder freier Tag; bleibt die Monotonie 2 Tage über 2,5 oder steigt über 3,0, kommt innerhalb von 48 h ein Ruhetag.\n` +
    `- CTL am Tag 14 höchstens ${CTL_MAX_LOSS * 100} % unter der heutigen CTL.\n` +
    `- Mindestens ${MIN_REST} Ruhetage mit 0 ESS (Zone Off).\n` +
    `- Anspruch einer Einheit: REST = 0 ESS/Off; EASY = unter 120 ESS und nur Z1/Z2; TRAIN = 120–199 ESS oder Z3; QUALITY = ab 200 ESS oder Z4/Z5.\n` +
    `- Steht bei einem Tag "erlaubt max. X", darf die Einheit dort höchstens Anspruch X haben (die Grenze ergibt sich aus Monotonie/ACWR des Vortags).\n` +
    `- Harte Einheiten (QUALITY) nicht direkt hintereinander.\n` +
    `- Ziel: CTL halten oder leicht steigern, ohne Regelverstoß.`;
  const msgs: LlmMsg[] = [
    { role: 'system', content: system },
    {
      role: 'user',
      content:
        `${context}\n\nAKTUELLER PLAN (Simulation):\n${planTable(days, days.map((d) => d.load), days.map((d) => d.zone), base.rows, base.checks, wp ? days.map((d) => d.sport) : undefined)}\n\n` +
        `PRÜFUNG AKTUELLER PLAN: ${base.violations.length ? base.violations.join('; ') : 'keine Verstöße'}\n\n${rules}`,
    },
  ];
  const rounds: LoopRound[] = [];
  for (let r = 1; r <= MAX_ROUNDS; r++) {
    onProgress?.(`Runde ${r}/${MAX_ROUNDS}: Kira plant…`);
    const text = await llm(msgs);
    const j = parseJsonLoose<{ summary?: string; days?: { date: string; load: number; zone: string; why?: string; sport?: string; training?: string; alt?: string; lang_h?: number; es_effekt?: number }[] }>(text);
    const byDate = new Map((j.days || []).map((x) => [x.date, x]));
    const loads = days.map((d, i) => {
      const x = byDate.get(d.date);
      if (d.locked || i === 0 || !x || !Number.isFinite(Number(x.load))) return d.load;
      return Math.max(0, Math.min(300, Math.round(Number(x.load))));
    });
    const zones = days.map((d, i) => {
      const x = byDate.get(d.date);
      if (d.locked || i === 0 || !x?.zone) return d.zone;
      return loads[i] <= 0 ? 'Off' : String(x.zone);
    });
    const why = days.map((d) => String(byDate.get(d.date)?.why || ''));
    const sports = days.map((d, i) => {
      const x = byDate.get(d.date);
      if (d.locked || i === 0 || !x?.sport) return d.sport;
      return loads[i] <= 0 ? 'Off' : String(x.sport);
    });
    const training = days.map((d) => String(byDate.get(d.date)?.training || ''));
    const alt = days.map((d) => String(byDate.get(d.date)?.alt || ''));
    const zahl = (v: unknown) => { const n = Number(v); return v != null && v !== '' && Number.isFinite(n) ? n : null; };
    const langH = days.map((d, i) => { const n = zahl(byDate.get(d.date)?.lang_h); return n != null && n > 0 && n <= 12 && loads[i] > 0 ? Math.round(n * 4) / 4 : null; });
    const esEffekt = days.map((d, i) => (langH[i] != null ? zahl(byDate.get(d.date)?.es_effekt) : null));
    onProgress?.(`Runde ${r}/${MAX_ROUNDS}: Simulation prüft…`);
    const ev = evaluate(days, loads, zones, simulate, check, wp && args.langTermine?.length ? { langH, termine: args.langTermine } : undefined);
    rounds.push({ loads, zones, why, summary: String(j.summary || ''), violations: ev.violations, maxAcwr: ev.maxAcwr, restDays: ev.restDays, finalCtl: ev.finalCtl, ...(wp ? { sports, training, alt, langH, esEffekt } : {}) });
    if (!ev.violations.length) break;
    msgs.push({ role: 'assistant', content: text });
    msgs.push({
      role: 'user',
      content:
        `Die Simulation deines Vorschlags ergibt diese Regelverstöße:\n- ${ev.violations.join('\n- ')}\n\n` +
        `Simulation deines Vorschlags:\n${planTable(days, loads, zones, ev.rows, ev.checks, wp ? sports : undefined)}\n\nKorrigiere den Plan so, dass alle Regeln erfüllt sind. Gleiches JSON-Format. Die summary beschreibt den fertigen Plan für Marc (Wochenidee, Schwerpunkte, Entlastung) – nicht die Korrektur oder Regelverstöße.`,
    });
  }
  const best = [...rounds].sort((a, b) => a.violations.length - b.violations.length || b.finalCtl - a.finalCtl)[0];
  return { rounds, best };
}
