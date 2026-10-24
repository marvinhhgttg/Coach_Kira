import { useState } from 'react';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { fmtDateShort, fmtNum } from '../lib/format';
import { monotonyPlanViolations } from '../lib/monotony';
import { planChecks, type PlanRow, type PlanState } from '../screens/StrategyPanel';

interface Aenderung { datum: string; last?: number; sport?: string; zone?: string }
interface Kennzahlen { ctlEnde: number; formEnde: number; maxAcwr: number; maxMono: number; konflikte: number; verstoesse: string[] }

const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const wd = (d: string) => WT[new Date(d + 'T12:00:00').getDay()];

function kennzahlen(rows: PlanRow[], recGap: number): Kennzahlen {
  const last = rows[rows.length - 1];
  const free = rows.slice(1);
  return {
    ctlEnde: last?.ctl ?? 0,
    formEnde: last && last.ctl > 0 ? ((last.ctl - last.atl) / last.ctl) * 100 : 0,
    maxAcwr: Math.max(0, ...free.map((r) => r.acwr)),
    maxMono: Math.max(0, ...free.map((r) => r.monotony)),
    konflikte: planChecks(rows, recGap).filter((c) => c.issues.length).length,
    verstoesse: monotonyPlanViolations(rows.map((r, i) => ({ label: `${wd(r.date)} ${fmtDateShort(r.date)}`, load: r.sim_load, mono: i === 0 ? null : r.monotony, locked: r.locked || i === 0 }))),
  };
}

export function WhatIfTile({
  sim,
  state,
  locks,
  simulateWith,
  recGap,
  onApply,
}: {
  sim: PlanRow[];
  state: PlanState;
  locks: boolean[];
  simulateWith: (loads: number[]) => PlanRow[];
  recGap: number;
  onApply: (loads: number[], sports: string[], zones: string[]) => void;
}) {
  const [frage, setFrage] = useState('');
  const [busy, setBusy] = useState<string | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [res, setRes] = useState<{ aend: Aenderung[]; abgelehnt: string[]; vorher: Kennzahlen; nachher: Kennzahlen; antwort: string; loads: number[]; sports: string[]; zones: string[] } | null>(null);

  async function run() {
    if (!frage.trim()) return;
    setErr(null);
    setRes(null);
    try {
      setBusy('Kira versteht die Frage…');
      const tabelle = sim
        .map((r, i) => `${r.date} ${wd(r.date)} | ${Math.round(state.loads[i] ?? r.sim_load)} ESS ${state.sports[i] || r.sport} ${state.zones[i] || r.zone}${locks[i] ? ' FEST' : ''}`)
        .join('\n');
      const t1 = await localChat(
        [
          { role: 'system', content: 'Du übersetzt Trainingsfragen in Planänderungen. Antworte nur mit JSON, keine Erklärung.' },
          {
            role: 'user',
            content:
              `Aktueller Plan (Datum Wochentag | Last Sport Zone):\n${tabelle}\n\nFrage von Marc: „${frage.trim()}“\n\n` +
              'Welche Tage würden sich wie ändern? Nur Tage aus der Tabelle. Fehlt eine Angabe, lass das Feld weg. ' +
              'Dauer in ESS umrechnen: Rad Z2 ≈ 85 ESS/h, Lauf Z2 ≈ 120 ESS/h, Wandern ≈ 50 ESS/h, Ruhetag = 0. ' +
              'Format: {"aenderungen":[{"datum":"2026-10-24","last":300,"sport":"Bike","zone":"Z2"}]}',
          },
        ],
        { temperature: 0.1, numCtx: 8192 },
      );
      const raw: any = parseJsonLoose<any>(t1) || {};
      const list: any[] = Array.isArray(raw.aenderungen) ? raw.aenderungen : Array.isArray(raw) ? raw : Object.values(raw).find((x) => Array.isArray(x)) || [];
      const idx = new Map(sim.map((r, i) => [r.date, i]));
      const loads = state.loads.slice();
      const sports = state.sports.slice();
      const zones = state.zones.slice();
      const aend: Aenderung[] = [];
      const abgelehnt: string[] = [];
      for (const a of list) {
        const d = String(a?.datum ?? a?.date ?? '').slice(0, 10);
        const i = idx.get(d);
        if (i == null) continue;
        if (locks[i] || i === 0) {
          abgelehnt.push(`${wd(d)} ${fmtDateShort(d)} (${i === 0 ? 'heute' : 'gesperrt'})`);
          continue;
        }
        const l = Number(a?.last ?? a?.load);
        if (Number.isFinite(l)) loads[i] = Math.max(0, Math.min(500, Math.round(l)));
        if (a?.sport) sports[i] = String(a.sport);
        if (a?.zone) zones[i] = String(a.zone);
        aend.push({ datum: d, last: loads[i], sport: sports[i], zone: zones[i] });
      }
      if (!aend.length) throw new Error(abgelehnt.length ? `Nur gesperrte Tage betroffen: ${abgelehnt.join(', ')}.` : 'Kira konnte aus der Frage keine Planänderung ableiten. Nenne Tag und Dauer oder Last, z. B. „Samstag 5 h Rad“.');

      setBusy('Simuliere…');
      const before = sim;
      const after = simulateWith(loads).map((r, i) => ({ ...r, sport: sports[i] ?? r.sport, zone: zones[i] ?? r.zone }));
      const vorher = kennzahlen(before, recGap);
      const nachher = kennzahlen(after, recGap);

      setBusy('Kira bewertet…');
      const t2 = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, sachlich, keine Emojis. Nutze nur die gelieferten Zahlen.' },
          {
            role: 'user',
            content:
              `Frage: „${frage.trim()}“\nÄnderungen: ${aend.map((a) => `${a.datum} → ${a.last} ESS ${a.sport} ${a.zone}`).join('; ')}\n` +
              `Vorher: CTL am Ende ${vorher.ctlEnde.toFixed(0)}, Form ${vorher.formEnde.toFixed(1)} %, max ACWR ${vorher.maxAcwr.toFixed(2)}, Konflikte ${vorher.konflikte}, Monotonie-Verstöße ${vorher.verstoesse.length}\n` +
              `Nachher: CTL am Ende ${nachher.ctlEnde.toFixed(0)}, Form ${nachher.formEnde.toFixed(1)} %, max ACWR ${nachher.maxAcwr.toFixed(2)}, Konflikte ${nachher.konflikte}, Monotonie-Verstöße ${nachher.verstoesse.length}` +
              (nachher.verstoesse.length ? ` (${nachher.verstoesse.join('; ')})` : '') +
              '\n\nBeantworte die Frage in 2–4 Sätzen: lohnt es sich, was kostet es, was müsste ich danach anpassen? Regeln: ACWR über 1,3 ist kritisch; Monotonie-Regel: Einzelwerte sind unkritisch, Verstöße nicht.',
          },
        ],
        { temperature: 0.3, numCtx: 8192 },
      );
      setRes({ aend, abgelehnt, vorher, nachher, antwort: t2.trim(), loads, sports, zones });
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  }

  const delta = (a: number, b: number, d = 0, invert = false) => {
    const x = b - a;
    const cls = Math.abs(x) < 10 ** -d / 2 ? 'text-ink-dim' : (x > 0) !== invert ? 'text-green-300' : 'text-orange-300';
    return <span className={`tnum ${cls}`}>{x > 0 ? '+' : ''}{fmtNum(x, d)}</span>;
  };

  return (
    <div className="panel-raised p-4">
      <div className="label">Coach Kira · Was wäre, wenn …?</div>
      <p className="text-2xs text-ink-dim mt-1">Frag in Worten, Kira rechnet es mit der echten Simulation durch. Beispiele: „Samstag 5 h Rad statt Lauf“, „Mittwoch frei und dafür Donnerstag 200 ESS“.</p>
      <div className="mt-3 flex flex-wrap gap-2">
        <input
          className="input flex-1 min-w-[16rem]"
          placeholder="Was wäre, wenn …"
          value={frage}
          onChange={(e) => setFrage(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && !busy && run()}
        />
        <button className="btn btn-primary text-xs px-3 py-1" onClick={run} disabled={!!busy || !frage.trim()}>
          {busy || 'Durchrechnen'}
        </button>
      </div>
      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}
      {res && (
        <div className="mt-3 flex flex-col gap-2">
          <div className="text-xs text-ink-muted">
            Geändert: {res.aend.map((a) => `${wd(a.datum)} ${fmtDateShort(a.datum)} → ${a.last} ESS ${a.sport} ${a.zone}`).join(' · ')}
            {res.abgelehnt.length > 0 && <span className="text-orange-300"> · nicht geändert: {res.abgelehnt.join(', ')}</span>}
          </div>
          <table className="text-xs w-full max-w-xl">
            <thead>
              <tr className="text-ink-dim border-b border-border">
                <th className="text-left font-normal py-1">Kennzahl</th>
                <th className="text-right font-normal py-1">vorher</th>
                <th className="text-right font-normal py-1">nachher</th>
                <th className="text-right font-normal py-1">Δ</th>
              </tr>
            </thead>
            <tbody className="tnum">
              <tr><td className="py-0.5">CTL am Planende</td><td className="text-right">{fmtNum(res.vorher.ctlEnde)}</td><td className="text-right">{fmtNum(res.nachher.ctlEnde)}</td><td className="text-right">{delta(res.vorher.ctlEnde, res.nachher.ctlEnde)}</td></tr>
              <tr><td className="py-0.5">Form am Planende</td><td className="text-right">{fmtNum(res.vorher.formEnde, 1)} %</td><td className="text-right">{fmtNum(res.nachher.formEnde, 1)} %</td><td className="text-right">{delta(res.vorher.formEnde, res.nachher.formEnde, 1)}</td></tr>
              <tr><td className="py-0.5">max. ACWR</td><td className="text-right">{fmtNum(res.vorher.maxAcwr, 2)}</td><td className="text-right">{fmtNum(res.nachher.maxAcwr, 2)}</td><td className="text-right">{delta(res.vorher.maxAcwr, res.nachher.maxAcwr, 2, true)}</td></tr>
              <tr><td className="py-0.5">Konflikte im Plan</td><td className="text-right">{res.vorher.konflikte}</td><td className="text-right">{res.nachher.konflikte}</td><td className="text-right">{delta(res.vorher.konflikte, res.nachher.konflikte, 0, true)}</td></tr>
              <tr><td className="py-0.5">Verstöße Monotonie-Regel</td><td className="text-right">{res.vorher.verstoesse.length}</td><td className="text-right">{res.nachher.verstoesse.length}</td><td className="text-right">{delta(res.vorher.verstoesse.length, res.nachher.verstoesse.length, 0, true)}</td></tr>
            </tbody>
          </table>
          {res.nachher.verstoesse.length > 0 && <ul className="text-2xs text-orange-300 list-disc pl-5">{res.nachher.verstoesse.map((v, i) => <li key={i}>{v}</li>)}</ul>}
          <p className="text-sm text-ink whitespace-pre-line">{res.antwort}</p>
          <div className="flex gap-2">
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => onApply(res.loads, res.sports, res.zones)}>In Slider übernehmen</button>
            <span className="text-2xs text-ink-dim self-center">Gespeichert wird erst mit „In timeline speichern“.</span>
          </div>
        </div>
      )}
    </div>
  );
}
