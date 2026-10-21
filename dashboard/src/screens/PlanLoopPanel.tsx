/**
 * Kira-Plan mit Prüfschleife: 14-Tage-Vorschlag, Simulation rechnet ATL/CTL/ACWR/Monotonie,
 * planChecks prüft die Regeln, Verstöße gehen als Rückmeldung zurück an Kira – höchstens 3 Runden.
 * Berechnet lokal auf dem Mac.
 */
import { useCallback, useMemo, useState } from 'react';
import { Panel } from '../components/UI';
import { useEngine } from '../lib/engine';
import { runPlanLoop, evaluate, type LoopDay, type LoopRound } from '../lib/planLoop';
import type { LoopSimRow } from '../lib/planLoop';
import { planChecks, type PlanRow } from './StrategyPanel';
import { fmtDateShort, fmtNum } from '../lib/format';

export function PlanLoopPanel({
  sim,
  loads,
  zones,
  sports,
  locks,
  simulateWith,
  recGap,
  onApply,
}: {
  sim: PlanRow[];
  loads: number[];
  zones: string[];
  sports: string[];
  locks: boolean[];
  simulateWith: (l: number[], z: string[]) => PlanRow[];
  recGap: number;
  onApply: (loads: number[], zones: string[]) => void;
}) {
  const [engine] = useEngine();
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState('');
  const [result, setResult] = useState<{ rounds: LoopRound[]; best: LoopRound } | null>(null);
  const [err, setErr] = useState<string | null>(null);

  const days: LoopDay[] = useMemo(
    () =>
      sim.map((r, i) => ({
        date: r.date,
        day: r.day,
        load: loads[i] ?? r.sim_load,
        zone: zones[i] ?? r.zone,
        sport: sports[i] ?? r.sport,
        locked: !!locks[i],
      })),
    [sim, loads, zones, sports, locks],
  );

  const simulate = useCallback(
    (l: number[], z: string[]) => simulateWith(l, z),
    [simulateWith],
  );
  const check = useCallback(
    (rows: LoopSimRow[]) => planChecks(rows as unknown as PlanRow[], recGap).map((c) => ({ ...c, cap: c.cap === 'QUALITY' ? undefined : c.cap })),
    [recGap],
  );

  const context = useMemo(() => {
    const s = sim[0];
    return (
      'STAND HEUTE (' +
      (s?.date || '') +
      '): ATL ' +
      fmtNum(s?.atl) +
      ', CTL ' +
      fmtNum(s?.ctl) +
      ', ACWR ' +
      fmtNum(s?.acwr, 2) +
      ', Monotonie ' +
      fmtNum(s?.monotony, 2) +
      '.'
    );
  }, [sim]);

  const run = useCallback(async () => {
    setBusy(true);
    setErr(null);
    setResult(null);
    try {
      const r = await runPlanLoop({
        days,
        simulate,
        check,
        context,
        onProgress: setProgress,
      });
      setResult(r);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
      setProgress('');
    }
  }, [days, simulate, check, context]);

  const current = useMemo(
    () => evaluate(days, days.map((d) => d.load), days.map((d) => d.zone), simulate, check),
    [days, simulate, check],
  );

  const bestEval = useMemo(
    () => (result ? evaluate(days, result.best.loads, result.best.zones, simulate, check) : null),
    [result, days, simulate, check],
  );

  const violationsText =
    current.violations.length === 0
      ? 'Aktueller Plan: 0 Regelverstöße'
      : `Aktueller Plan: ${current.violations.length} Regelverstöße`;

  const roundSummary =
    result && result.rounds.length
      ? result.rounds
          .map((r, i) => 'Runde ' + (i + 1) + ': ' + r.violations.length + ' Verstöße')
          .join(' · ')
      : '';

  return (
    <Panel title="Kira-Plan mit Prüfschleife">
      <p className="text-2xs text-ink-dim leading-relaxed">
        Kira schlägt einen 14-Tage-Plan vor, deine Simulation rechnet ATL, CTL, ACWR und
        Monotonie nach und prüft die Regeln (ACWR 0,8–1,3, Monotonie als Verlauf (max. 4 Tage ohne Entlastung, Ruhetag nach anhaltend über 2,5), CTL-Verlust höchstens 5 %, mindestens 2
        Ruhetage, Abstand harter Einheiten, feste Tage bleiben). Bei Verstößen korrigiert
        Kira – höchstens 3 Runden. Läuft lokal auf deinem Mac.
      </p>

      <div className="mt-3 flex flex-wrap items-center justify-between gap-2">
        <div className="text-xs tnum">{violationsText}</div>
        {current.violations.length > 0 && (
          <details className="text-2xs text-orange-300">
            <summary className="cursor-pointer">Verstöße anzeigen</summary>
            <ul className="mt-1 space-y-0.5">
              {current.violations.map((v) => (
                <li key={v}>· {v}</li>
              ))}
            </ul>
          </details>
        )}
      </div>

      <div className="mt-3">
        <button
          className="btn btn-primary text-xs"
          onClick={run}
          disabled={busy || engine !== 'local' || !sim.length}
        >
          {busy
            ? progress || 'Kira plant…'
            : engine === 'local'
              ? 'Kira-Plan prüfen lassen (lokal)'
              : 'nur lokal – Schalter auf „Lokal (Mac)“'}
        </button>
      </div>

      {err && <p className="mt-2 text-xs text-red-300">{err}</p>}

      {result && (
        <div className="mt-4 space-y-3">
          {roundSummary && (
            <div className="text-2xs text-ink-dim">{roundSummary}</div>
          )}
          {result.best.summary && (
            <p className="text-sm text-ink-muted leading-relaxed">
              {result.best.summary}
            </p>
          )}

          <div className="overflow-x-auto">
            <table className="min-w-[900px] text-xs">
              <thead className="sticky top-0 bg-bg-panel z-10">
                <tr className="text-left border-b border-border">
                  {['Tag', 'Jetzt', 'Kira', 'ACWR Kira', 'Fest', 'Begründung'].map((h) => (
                    <th key={h} className="label px-2 py-2">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {days.map((d, i) => {
                  const nowLoad = d.load;
                  const nowZone = d.zone;
                  const kLoad = result.best.loads[i];
                  const kZone = result.best.zones[i];
                  const kRow = bestEval?.rows[i];
                  const kAcwr = kRow?.acwr ?? 0;
                  const isLocked = d.locked || i === 0;
                  const diff = kLoad !== nowLoad || kZone !== nowZone;
                  return (
                    <tr key={d.date} className="align-top">
                      <td className="px-2 py-2.5">
                        <div className="tnum font-medium">
                          {d.day} {fmtDateShort(d.date)}
                        </div>
                        <div className="text-2xs text-ink-dim">Tag {i + 1}</div>
                      </td>
                      <td className="px-2 py-2.5">
                        <div className="tnum">{fmtNum(nowLoad)} ESS</div>
                        <div className="text-2xs text-ink-dim">{nowZone}</div>
                      </td>
                      <td className={`px-2 py-2.5 ${diff ? 'font-semibold text-accent' : ''}`}>
                        <div className="tnum">{fmtNum(kLoad)} ESS</div>
                        <div className="text-2xs text-ink-dim">{kZone}</div>
                      </td>
                      <td className={`px-2 py-2.5 tnum ${kAcwr > 1.3 ? 'text-orange-300' : ''}`}>
                        {fmtNum(kAcwr, 2)}
                      </td>
                      <td className="px-2 py-2.5 text-2xs text-ink-dim">
                        {isLocked ? 'fest' : '—'}
                      </td>
                      <td className="px-2 py-2.5 text-2xs text-ink-dim">
                        {result.best.why[i] || '—'}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="grid gap-2 grid-cols-2 md:grid-cols-4 text-xs tnum">
            <div className="panel-raised p-2.5">
              <div className="text-2xs text-ink-dim">Max. ACWR</div>
              <div className={`mt-1 font-semibold ${(bestEval?.maxAcwr ?? 0) > 1.3 ? 'text-orange-300' : 'text-green-300'}`}>
                {fmtNum(bestEval?.maxAcwr, 2)}
              </div>
            </div>
            <div className="panel-raised p-2.5">
              <div className="text-2xs text-ink-dim">Ruhetage</div>
              <div className="mt-1 font-semibold">{bestEval?.restDays}</div>
            </div>
            <div className="panel-raised p-2.5">
              <div className="text-2xs text-ink-dim">CTL am Ende (Kira)</div>
              <div className="mt-1 font-semibold text-accent">{fmtNum(bestEval?.finalCtl)}</div>
            </div>
            <div className="panel-raised p-2.5">
              <div className="text-2xs text-ink-dim">CTL am Ende (aktuell)</div>
              <div className="mt-1 font-semibold">{fmtNum(current.finalCtl)}</div>
            </div>
          </div>

          {bestEval && bestEval.violations.length > 0 ? (
            <div className="text-orange-300 text-xs">
              <div className="font-semibold">Verbleibende Verstöße:</div>
              <ul className="mt-1 space-y-0.5">
                {bestEval.violations.map((v) => (
                  <li key={v}>· {v}</li>
                ))}
              </ul>
            </div>
          ) : (
            <div className="text-green-300 text-xs">Alle Regeln erfüllt.</div>
          )}

          <div className="space-y-1.5">
            <button
              className="btn btn-ghost text-xs"
              onClick={() => onApply(result.best.loads, result.best.zones)}
              disabled={busy}
            >
              In Slider übernehmen
            </button>
            <p className="text-2xs text-ink-dim">
              Übernimmt nur in die Simulation; gespeichert wird erst mit „In timeline
              speichern“.
            </p>
          </div>
        </div>
      )}
    </Panel>
  );
}
