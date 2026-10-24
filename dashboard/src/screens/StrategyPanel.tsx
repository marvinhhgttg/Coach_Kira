import { useEffect, useMemo, useState, type ReactNode } from 'react';
import {
  Bar,
  CartesianGrid,
  ComposedChart,
  Legend,
  Line,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { fetchChartData } from '../lib/api';
import { fmtDateShort, fmtNum } from '../lib/format';
import { Panel } from '../components/UI';
import { buildDays, demandOf, RECOVERY_FETCH_METRICS, VOTE_LABEL, VOTE_ORDER, localIsoDate } from './RecoveryPanel';
import { atlStep, ctlStep, windowMean } from '../lib/ctlModel';
import { assessMonotony, monoTone } from '../lib/monotony';
import { monotonyPlanViolations } from '../lib/monotony';

// ---------------------------------------------------------------------------
// Strategische Planung für das Plan Cockpit (alles lokal, deterministisch):
// 1 Tagesprüfung + Hebel · 2 Ziel rückwärts · 3 Szenarien A/B/C · 4 Vorlagen
// 5 Horizont 4/8 Wochen · 6 Events & Formspitze · 7 Erholungsregel & 80/20 · 8 Spielraum ±10 %
// ---------------------------------------------------------------------------

type Vote = (typeof VOTE_ORDER)[number];
export type PlanRow = {
  date: string;
  day?: string;
  sim_load: number;
  sport: string;
  zone: string;
  te_ae: number;
  te_an: number;
  atl: number;
  ctl: number;
  acwr: number;
  monotony: number;
  locked?: boolean;
};
export type PlanState = { loads: number[]; sports: string[]; zones: string[]; teAe: number[]; teAn: number[] };
export type PlanEvent = { id: string; date: string; label: string };

// Fortschreibung wie im Plan Cockpit (src/lib/ctlModel.ts)
/** series: Lastreihe inkl. heutigem Load (für das 28-Tage-Mittel der CTL). */
const step = (atl: number, ctl: number, load: number, series: number[]) => ({
  atl: atlStep(atl, load),
  ctl: ctlStep(ctl, load, windowMean(series)),
});

// Belastungsregeln wie im Tagesentscheid (Stand Vortag): ACWR fest, Monotonie als Verlauf (lib/monotony.ts)
const ACWR_WARN = 1.3;
const ACWR_HIGH = 1.5;
export function capFromLoad(
  acwr: number | null,
  mono: number | null,
  loads: number[] = [],
  monoHist: (number | null)[] = [],
): { cap: Vote; why: string | null } {
  const order = (v: Vote) => VOTE_ORDER.indexOf(v);
  let cap: Vote = 'QUALITY';
  let why: string | null = null;
  if ((acwr ?? 0) > ACWR_HIGH) { cap = 'EASY'; why = `ACWR ${fmtNum(acwr, 2)} > 1,5`; }
  else if ((acwr ?? 0) > ACWR_WARN) { cap = 'TRAIN'; why = `ACWR ${fmtNum(acwr, 2)} > 1,3`; }
  const ma = assessMonotony({ mono, loads, monoHistory: monoHist });
  if (ma.cap && order(ma.cap) < order(cap)) { cap = ma.cap; why = `${ma.text.replace(/ → .*$/, '')} (Stand Vortag)`; }
  return { cap, why };
}
const formPct = (atl: number, ctl: number) => (ctl > 0 ? ((ctl - atl) / ctl) * 100 : 0);
const isHard = (d: PlanRow) => d.sim_load >= 200 || demandOf(d.sim_load, d.sport, d.zone, d.te_ae, d.te_an) === 'QUALITY';

export type DayCheck = { demand: Vote; cap: Vote; why: string | null; form: number; issues: string[]; level: 'ok' | 'warn' | 'bad' };
export function planChecks(sim: PlanRow[], recGap: number): DayCheck[] {
  let lastHard = -99;
  return sim.map((d, i) => {
    const demand = demandOf(d.sim_load, d.sport, d.zone, d.te_ae, d.te_an);
    const prev = i > 0 ? sim[i - 1] : null;
    const { cap, why } = prev
      ? capFromLoad(prev.acwr, prev.monotony, sim.slice(0, i).map((r) => r.sim_load), sim.slice(Math.max(0, i - 8), i - 1).map((r) => r.monotony))
      : { cap: 'QUALITY' as Vote, why: null };
    const issues: string[] = [];
    if (i > 0 && !d.locked && VOTE_ORDER.indexOf(demand) > VOTE_ORDER.indexOf(cap)) issues.push(`${VOTE_LABEL[demand]} geplant, ${why} erlaubt max. ${VOTE_LABEL[cap]}`);
    if (isHard(d)) {
      if (i - lastHard < recGap && !d.locked) issues.push(`nur ${i - lastHard} T nach letzter harter Einheit (Regel ≥ ${recGap} T)`);
      lastHard = i;
    }
    if (d.acwr > 1.5) issues.push(`ACWR ${fmtNum(d.acwr, 2)} nach der Einheit`);
    const level = issues.some((x) => x.includes('erlaubt max. Easy') || x.includes('ACWR')) ? 'bad' : issues.length ? 'warn' : 'ok';
    return { demand, cap, why, form: formPct(d.atl, d.ctl), issues, level };
  });
}

// ---------------------------------------------------------------------------
// LocalStorage-Helfer (geräteweise; Events auch im Plan-Tabellenkopf genutzt)
// ---------------------------------------------------------------------------
const EV_KEY = 'ck_plan_events_v1';
const SC_KEY = 'ck_plan_scenarios_v1';
function readLS<T>(k: string, def: T): T {
  try {
    const v = localStorage.getItem(k);
    return v ? (JSON.parse(v) as T) : def;
  } catch {
    return def;
  }
}
function writeLS(k: string, v: unknown) {
  try {
    localStorage.setItem(k, JSON.stringify(v));
    window.dispatchEvent(new CustomEvent('ck-ls', { detail: k }));
  } catch {
    /* ignore */
  }
}
// Modulweiter Speicher: funktioniert auch, wenn localStorage blockiert ist (dann nur bis zum Neuladen)
let evStore: PlanEvent[] = readLS<PlanEvent[]>(EV_KEY, []);
const evSubs = new Set<(e: PlanEvent[]) => void>();
export function useEvents(): [PlanEvent[], (e: PlanEvent[]) => void] {
  const [ev, setEv] = useState<PlanEvent[]>(evStore);
  useEffect(() => {
    evSubs.add(setEv);
    return () => {
      evSubs.delete(setEv);
    };
  }, []);
  const set = (v: PlanEvent[]) => {
    evStore = [...v].sort((x, y) => x.date.localeCompare(y.date));
    writeLS(EV_KEY, evStore);
    evSubs.forEach((f) => f(evStore));
  };
  return [ev, set];
}

// ---------------------------------------------------------------------------
// Muster & Projektion
// ---------------------------------------------------------------------------
const WDN = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const wdOf = (iso: string) => WDN[new Date(iso + 'T12:00:00').getDay()];
const addDays = (iso: string, n: number) => {
  const d = new Date(iso + 'T12:00:00');
  d.setDate(d.getDate() + n);
  return localIsoDate(d);
};
const diffDays = (a: string, b: string) => Math.round((new Date(b + 'T12:00:00').getTime() - new Date(a + 'T12:00:00').getTime()) / 86400000);

type Pattern = { id: string; label: string; w: Record<string, number>; hint: string };
const PATTERNS: Pattern[] = [
  { id: 'polar', label: 'Hart–leicht (polarisiert)', w: { Mo: 0.7, Di: 1.35, Mi: 0.55, Do: 1.1, Fr: 0, Sa: 1.5, So: 0.8 }, hint: 'Di/Do/Sa betont, Fr Ruhetag – senkt Monotonie' },
  { id: 'hlmr', label: 'Hart–leicht–mittel–Ruhe', w: { Mo: 1.3, Di: 0.6, Mi: 1.0, Do: 0, Fr: 1.3, Sa: 0.6, So: 1.0 }, hint: '4-Tage-Rhythmus, Do Ruhetag' },
  { id: 'even', label: 'Gleichmäßig + 1 Ruhetag', w: { Mo: 1, Di: 1, Mi: 1, Do: 1, Fr: 0, Sa: 1, So: 1 }, hint: 'Einfach, aber hohe Monotonie' },
];
function patternLoads(dates: string[], weekly: number, p: Pattern) {
  const sumW = Object.values(p.w).reduce((a, b) => a + b, 0);
  return dates.map((d) => Math.round(((p.w[wdOf(d)] ?? 1) * weekly) / sumW / 5) * 5);
}
function monoOf(xs: number[]) {
  const w = xs.slice(-7);
  if (w.length < 7) return null;
  const m = w.reduce((a, b) => a + b, 0) / 7;
  const sd = Math.sqrt(w.reduce((a, b) => a + (b - m) ** 2, 0) / 7);
  return sd > 0 ? m / sd : 2.5;
}
/** Tagesweise Projektion ab Anker (Tag 0 = sim[0], bereits nach Belastung). */
function project(anchor: PlanRow, seed: number[], loads: number[]) {
  let atl = anchor.atl;
  let ctl = anchor.ctl;
  const series = [...seed];
  return loads.map((l, i) => {
    if (i > 0) series.push(l);
    if (i > 0) ({ atl, ctl } = step(atl, ctl, l, series));
    return { i, load: l, atl, ctl, acwr: ctl > 0 ? atl / ctl : 0, mono: monoOf(series), form: formPct(atl, ctl) };
  });
}

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------
function Box({ title, right, children, hint }: { title: string; right?: ReactNode; children: ReactNode; hint?: string }) {
  return (
    <div className="panel-raised p-3 min-w-0">
      <div className="flex flex-wrap items-baseline justify-between gap-2 mb-2">
        <div>
          <div className="label">{title}</div>
          {hint && <p className="text-2xs text-ink-dim mt-0.5">{hint}</p>}
        </div>
        {right}
      </div>
      {children}
    </div>
  );
}
function Kpi({ label, value, sub, tone }: { label: string; value: ReactNode; sub?: ReactNode; tone?: string }) {
  return (
    <div className="rounded border border-border px-2.5 py-2 min-w-0">
      <div className="text-2xs text-ink-dim">{label}</div>
      <div className={`tnum text-sm font-semibold whitespace-nowrap ${tone || ''}`}>{value}</div>
      {sub != null && <div className="text-2xs text-ink-dim tnum">{sub}</div>}
    </div>
  );
}
const GRID = '#1f2731';
const amp = (lvl: 'ok' | 'warn' | 'bad') => (lvl === 'ok' ? 'text-green-300' : lvl === 'warn' ? 'text-yellow-300' : 'text-orange-300');

export function StrategyPanel({
  sim,
  state,
  loadSeed,
  simulateWith,
  applyLoads,
  applyState,
  recGap,
  setRecGap,
  kiPlan,
}: {
  sim: PlanRow[];
  state: PlanState;
  loadSeed: number[];
  simulateWith: (loads: number[]) => PlanRow[];
  applyLoads: (loads: number[], msg: string) => void;
  applyState: (s: PlanState, msg: string) => void;
  recGap: number;
  setRecGap: (n: number) => void;
  /** Wochenplan aus dem KI-Report, auf die 14 Plantage gelegt (Rest = aktueller Plan) */
  kiPlan?: { state: PlanState; days: number } | null;
}) {
  const [events, setEvents] = useEvents();
  const checks = useMemo(() => planChecks(sim, recGap), [sim, recGap]);
  const dates = sim.map((d) => d.date);
  const curWeekly = Math.round((sim.slice(1).reduce((a, d) => a + d.sim_load, 0) / Math.max(1, sim.length - 1)) * 7);

  // ---------- 7) persönliche Erholungszeit aus Historie ----------
  const [recInfo, setRecInfo] = useState<{ avg: number; n: number } | null>(null);
  useEffect(() => {
    fetchChartData('360d', RECOVERY_FETCH_METRICS)
      .then((r) => {
        const days = buildDays(r.data as any[]);
        const res: number[] = [];
        for (let i = 7; i < days.length - 3; i++) {
          const d = days[i];
          const e = d.essDay ?? 0;
          const dem = demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
          if (e < 200 && !(dem === 'QUALITY' && e >= 120)) continue;
          const prior = days.slice(Math.max(0, i - 28), i).map((x) => x.readiness).filter((x): x is number => x != null);
          if (prior.length < 7) continue;
          const base = prior.reduce((a, b) => a + b, 0) / prior.length;
          let t = 4;
          for (let k = 1; k <= 3; k++) if (days[i + k]?.readiness != null && days[i + k].readiness! >= base - 2) { t = k; break; }
          res.push(t);
        }
        if (res.length >= 5) {
          const avg = res.reduce((a, b) => a + b, 0) / res.length;
          setRecInfo({ avg, n: res.length });
          setRecGap(Math.max(1, Math.min(3, Math.round(avg))));
        }
      })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ---------- 1) Hebel ----------
  const lever = useMemo(() => {
    // Monotonie-Regel statt Ø-Monotonie: Verstöße zählen, nicht den Mittelwert drücken
    const maxMono = (rows: PlanRow[]) =>
      monotonyPlanViolations(rows.map((r) => ({ label: r.date, load: r.sim_load, mono: r.monotony, locked: r.locked }))).length;
    const conflicts = (rows: PlanRow[]) => planChecks(rows, recGap).filter((c) => c.issues.length).length;
    const base = { mono: maxMono(sim), conf: conflicts(sim) };
    let best: { i: number; mode: 'rest' | 'half'; mono: number; conf: number } | null = null;
    for (let i = 1; i < sim.length; i++) {
      if (sim[i].locked || sim[i].sim_load <= 0) continue;
      for (const mode of ['rest', 'half'] as const) {
        const l = state.loads.slice();
        l[i] = mode === 'rest' ? 0 : Math.round(l[i] / 2 / 5) * 5;
        const rows = simulateWith(l);
        const r = { i, mode, mono: maxMono(rows), conf: conflicts(rows) };
        const score = (x: { mono: number; conf: number }) => x.conf * 10 + x.mono * 5 + (mode === 'half' ? 0 : 0.5);
        if (!best || score(r) < score(best) - 0.01) best = r;
      }
    }
    return { base, best };
  }, [sim, state.loads, simulateWith, recGap]);
  const applyLever = () => {
    if (!lever.best) return;
    const l = state.loads.slice();
    const { i, mode } = lever.best;
    l[i] = mode === 'rest' ? 0 : Math.round(l[i] / 2 / 5) * 5;
    applyLoads(l, `${mode === 'rest' ? 'Ruhetag' : 'Halbierte Last'} am ${wdOf(sim[i].date)} ${fmtDateShort(sim[i].date).slice(0, 6)} übernommen (noch nicht gespeichert).`);
  };

  // ---------- 2) Ziel rückwärts ----------
  const anchor = sim[0];
  const [goalCtl, setGoalCtl] = useState<number>(() => Math.round(((anchor?.ctl ?? 900) * 1.08) / 10) * 10);
  const [goalDate, setGoalDate] = useState<string>(() => addDays(localIsoDate(), 42));
  const goal = useMemo(() => {
    if (!anchor) return null;
    const n = Math.max(1, diffDays(anchor.date, goalDate));
    // Gleichmäßige Tageslast per Bisektion so wählen, dass die CTL am Zieldatum passt
    const run = (daily: number) => {
      let atl = anchor.atl;
      let ctl = anchor.ctl;
      let maxAcwr = 0;
      let ctl7 = anchor.ctl;
      const series = [...loadSeed];
      for (let k = 1; k <= n; k++) {
        series.push(daily);
        ({ atl, ctl } = step(atl, ctl, daily, series));
        maxAcwr = Math.max(maxAcwr, ctl > 0 ? atl / ctl : 0);
        if (k === 7) ctl7 = ctl;
      }
      return { atl, ctl, maxAcwr, ctl7 };
    };
    let lo = 0;
    let hi = 600;
    for (let it = 0; it < 40; it++) {
      const mid = (lo + hi) / 2;
      if (run(mid).ctl < goalCtl) lo = mid;
      else hi = mid;
    }
    const daily = (lo + hi) / 2;
    const { atl, ctl, maxAcwr, ctl7 } = run(daily);
    const ramp = ((Math.pow(goalCtl / anchor.ctl, 7 / n) - 1) * 100);
    const lvl: 'ok' | 'warn' | 'bad' = daily < 0 ? 'ok' : ramp > 8 || maxAcwr > 1.5 ? 'bad' : ramp > 5 || maxAcwr > 1.3 ? 'warn' : 'ok';
    return { n, daily, weekly: daily * 7, ramp, maxAcwr, ctl7, formEnd: formPct(atl, ctl), lvl };
  }, [anchor, goalCtl, goalDate, loadSeed]);

  // ---------- 3) Szenarien ----------
  type Scen = PlanState & { name: 'A' | 'B' | 'C'; savedAt: string; startDate: string };
  const [scens, setScens] = useState<Scen[]>(() => readLS<Scen[]>(SC_KEY, []));
  const saveScen = (name: 'A' | 'B' | 'C') => {
    const s: Scen = { ...state, name, savedAt: new Date().toISOString(), startDate: dates[0] };
    const next = [...scens.filter((x) => x.name !== name), s].sort((a, b) => a.name.localeCompare(b.name));
    setScens(next);
    writeLS(SC_KEY, next);
  };
  const delScen = (name: string) => {
    const next = scens.filter((x) => x.name !== name);
    setScens(next);
    writeLS(SC_KEY, next);
  };
  const scenRows = useMemo(() => {
    const all = [
      { name: 'Aktuell', st: state },
      ...(kiPlan ? [{ name: 'Kira', st: kiPlan.state }] : []),
      ...scens.map((s) => ({ name: s.name, st: s as PlanState })),
    ];
    return all.map((x) => {
      const loads = dates.map((_, i) => (sim[i]?.locked ? state.loads[i] : x.st.loads[i] ?? 0));
      const rows = simulateWith(loads);
      const ch = planChecks(rows.map((r, i) => ({ ...r, sport: x.st.sports[i] ?? r.sport, zone: x.st.zones[i] ?? r.zone })), recGap);
      return {
        name: x.name,
        st: x.st,
        rows,
        sum: loads.slice(1).reduce((a, b) => a + b, 0),
        ctlEnd: rows[rows.length - 1]?.ctl,
        maxAcwr: Math.max(...rows.slice(1).map((r) => r.acwr)),
        maxMono: Math.max(...rows.slice(1).map((r) => r.monotony)),
        formEnd: rows.length ? formPct(rows[rows.length - 1].atl, rows[rows.length - 1].ctl) : 0,
        conf: ch.filter((c) => c.issues.length).length,
      };
    });
  }, [scens, state, sim, dates, simulateWith, recGap, kiPlan]);
  const SC_COL: Record<string, string> = { Aktuell: '#94a3b8', Kira: '#facc15', A: '#38bdf8', B: '#f97316', C: '#a855f7' };
  const scenChart = dates.map((d, i) => {
    const o: any = { d: `${wdOf(d)} ${fmtDateShort(d).slice(0, 6)}` };
    for (const s of scenRows) o[s.name] = Math.round(s.rows[i]?.ctl ?? 0);
    return o;
  });

  // ---------- 4) Vorlagen ----------
  const [tplWeekly, setTplWeekly] = useState<number>(curWeekly || 900);
  useEffect(() => setTplWeekly((v) => v || curWeekly), [curWeekly]);
  const applyPattern = (p: Pattern) => {
    const gen = patternLoads(dates, tplWeekly, p);
    const l = state.loads.map((v, i) => (i === 0 || sim[i]?.locked ? v : gen[i]));
    applyLoads(l, `Vorlage „${p.label}“ mit ${fmtNum(tplWeekly)} ESS/Woche übernommen (noch nicht gespeichert).`);
  };
  const applyDeload = () => {
    const l = state.loads.map((v, i) => (i === 0 || i > 7 || sim[i]?.locked ? v : Math.round((v * 0.65) / 5) * 5));
    applyLoads(l, 'Entlastungswoche (−35 % für die nächsten 7 Tage) übernommen (noch nicht gespeichert).');
  };
  const applyTaper = (ev: PlanEvent) => {
    const idx = dates.indexOf(ev.date);
    if (idx < 1) return;
    // Grundform: Vortag sehr locker, davor linear ansteigend bis 7 Tage vorher.
    const baseF: Record<number, number> = { 1: 0.35, 2: 0.55, 3: 0.65, 4: 0.72, 5: 0.8, 6: 0.86, 7: 0.92 };
    const make = (k: number) =>
      state.loads.map((v, i) => {
        if (i === 0 || sim[i]?.locked) return v;
        const f = baseF[idx - i];
        return f ? Math.round((v * (1 - k * (1 - f))) / 5) * 5 : v;
      });
    // Stärke so wählen, dass die Form am Event möglichst nahe +10 % liegt
    let best = { k: 1, l: make(1), err: 1e9, form: 0 };
    for (let k = 0.2; k <= 1.31; k += 0.1) {
      const l = make(k);
      const r = simulateWith(l)[idx];
      const f = formPct(r.atl, r.ctl);
      if (Math.abs(f - 10) < best.err) best = { k, l, err: Math.abs(f - 10), form: f };
    }
    applyLoads(best.l, `Tapering für „${ev.label}“ übernommen – Form am Event ≈ ${best.form > 0 ? '+' : ''}${fmtNum(best.form, 1)} % (noch nicht gespeichert).`);
  };

  // ---------- 5) Horizont ----------
  const [weeks, setWeeks] = useState<4 | 8>(8);
  const [rampPct, setRampPct] = useState(5);
  const [cycle31, setCycle31] = useState(true);
  const [hzPattern, setHzPattern] = useState('polar');
  const horizon = useMemo(() => {
    if (!anchor) return null;
    const nDays = weeks * 7;
    const allDates = Array.from({ length: nDays }, (_, i) => addDays(anchor.date, i));
    const pat = PATTERNS.find((p) => p.id === hzPattern) || PATTERNS[0];
    const base = curWeekly || 800;
    const loads: number[] = [];
    const phases: string[] = [];
    let weekIdx = 0;
    for (let i = 0; i < nDays; i++) {
      if (i < sim.length) {
        loads.push(state.loads[i] ?? 0);
        phases.push('Plan');
        continue;
      }
      weekIdx = Math.floor((i - sim.length) / 7); // 0 = erste Woche nach Plan
      const wk = Math.floor(i / 7); // Kalenderindex ab heute
      const deload = cycle31 && (wk + 1) % 4 === 0;
      const growth = Math.pow(1 + rampPct / 100, weekIdx + 1 - (cycle31 ? Math.floor((wk + 1) / 4) : 0));
      const weekly = base * growth * (deload ? 0.65 : 1);
      loads.push(patternLoads([allDates[i]], weekly, pat)[0]);
      phases.push(deload ? 'Entlastung' : 'Aufbau');
    }
    const daily = project(anchor, loadSeed, loads);
    const wk = Array.from({ length: weeks }, (_, w) => {
      const s = daily.slice(w * 7, w * 7 + 7);
      const prevEnd = w > 0 ? daily[w * 7 - 1].ctl : anchor.ctl;
      const end = s[s.length - 1];
      const ph = phases.slice(w * 7, w * 7 + 7);
      return {
        label: `W${w + 1} · ${fmtDateShort(allDates[w * 7]).slice(0, 6)}`,
        Wochenlast: Math.round(s.reduce((a, b) => a + (b.i === 0 ? 0 : b.load), 0) + (w === 0 ? state.loads[0] ?? 0 : 0)),
        CTL: Math.round(end.ctl),
        'Form %': Math.round(end.form * 10) / 10,
        Rampe: Math.round(((end.ctl - prevEnd) / prevEnd) * 1000) / 10,
        maxAcwr: Math.max(...s.map((x) => x.acwr)),
        maxMono: Math.max(...s.map((x) => x.mono ?? 0)),
        phase: ph.includes('Plan') ? (ph.every((p) => p === 'Plan') ? 'Plan' : 'Plan/Aufbau') : ph[0],
      };
    });
    return { allDates, daily, wk };
  }, [anchor, weeks, rampPct, cycle31, hzPattern, curWeekly, sim.length, state.loads, loadSeed]);

  // ---------- 6) Events ----------
  const [evDate, setEvDate] = useState(addDays(localIsoDate(), 21));
  const [evLabel, setEvLabel] = useState('');
  const evInfo = (ev: PlanEvent) => {
    const i14 = dates.indexOf(ev.date);
    if (i14 >= 0) return { src: 'Plan', ctl: sim[i14].ctl, form: formPct(sim[i14].atl, sim[i14].ctl), inPlan: true };
    const ih = horizon?.allDates.indexOf(ev.date) ?? -1;
    if (ih >= 0 && horizon) return { src: `Horizont (${PATTERNS.find((p) => p.id === hzPattern)?.label})`, ctl: horizon.daily[ih].ctl, form: horizon.daily[ih].form, inPlan: false };
    return null;
  };

  // ---------- 7) 80/20 des Plans ----------
  const dist = useMemo(() => {
    let low = 0, mid = 0, high = 0;
    sim.slice(1).forEach((d) => {
      if (d.sim_load <= 0) return;
      const m = (d.zone || '').toUpperCase().match(/[0-5]/g)?.map(Number) || [2];
      const share = d.sim_load / m.length;
      m.forEach((z) => (z <= 2 ? (low += share) : z === 3 ? (mid += share) : (high += share)));
    });
    const t = low + mid + high || 1;
    return { low: Math.round((low / t) * 100), mid: Math.round((mid / t) * 100), high: Math.round((high / t) * 100) };
  }, [sim]);
  const hardDays = sim.map((d, i) => ({ d, i })).filter((x) => x.i > 0 && isHard(x.d));

  // ---------- 8) Spielraum ----------
  const band = useMemo(() => {
    const lo = simulateWith(state.loads.map((v, i) => (i === 0 || sim[i]?.locked ? v : v * 0.9)));
    const hi = simulateWith(state.loads.map((v, i) => (i === 0 || sim[i]?.locked ? v : v * 1.1)));
    return {
      chart: sim.map((d, i) => ({
        d: `${wdOf(d.date)} ${fmtDateShort(d.date).slice(0, 6)}`,
        'CTL −10 %': Math.round(lo[i].ctl),
        CTL: Math.round(d.ctl),
        'CTL +10 %': Math.round(hi[i].ctl),
        'ACWR −10 %': Math.round(lo[i].acwr * 100) / 100,
        ACWR: Math.round(d.acwr * 100) / 100,
        'ACWR +10 %': Math.round(hi[i].acwr * 100) / 100,
      })),
      ctlLo: lo[lo.length - 1]?.ctl,
      ctlHi: hi[hi.length - 1]?.ctl,
      acwrLo: Math.max(...lo.slice(1).map((r) => r.acwr)),
      acwrHi: Math.max(...hi.slice(1).map((r) => r.acwr)),
      confHi: planChecks(hi, recGap).filter((c) => c.issues.length).length,
    };
  }, [state.loads, sim, simulateWith, recGap]);

  if (!sim.length) return null;
  const confCount = checks.filter((c) => c.issues.length).length;

  return (
    <Panel title="Strategische Planung">
      <div className="space-y-4">
        {/* 1) Tagesprüfung */}
        <Box
          title="1 · Prüfung der Plantage"
          hint="Je Tag: Anspruch der Einheit gegen das maximal mögliche Votum aus ACWR/Monotonie des Vortags (Regeln wie im Tagesentscheid), dazu Form und Erholungsabstand."
          right={<span className={`text-xs tnum font-semibold ${confCount ? 'text-orange-300' : 'text-green-300'}`}>{confCount ? `${confCount} Konflikt${confCount > 1 ? 'e' : ''}` : 'keine Konflikte'}</span>}
        >
          <div className="overflow-x-auto">
            <table className="w-full min-w-[720px] text-xs tnum">
              <thead>
                <tr className="text-ink-dim text-left">
                  {['Tag', 'Last', 'Anspruch', 'max. Votum', 'Monotonie', 'ACWR', 'Form', 'Prüfung'].map((h) => (
                    <th key={h} className="py-1 pr-2 font-normal">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sim.map((d, i) => {
                  const c = checks[i];
                  const ev = events.filter((e) => e.date === d.date);
                  return (
                    <tr key={d.date} className="border-t border-border">
                      <td className="py-1 pr-2 whitespace-nowrap">
                        {wdOf(d.date)} {fmtDateShort(d.date).slice(0, 6)}
                        {ev.map((e) => (
                          <span key={e.id} className="ml-1 chip border border-accent/50 text-accent text-2xs">★ {e.label}</span>
                        ))}
                      </td>
                      <td className="py-1 pr-2">{fmtNum(d.sim_load)}{d.locked ? ' · fix' : ''}</td>
                      <td className="py-1 pr-2">{d.sim_load > 0 ? VOTE_LABEL[c.demand] : 'Ruhe'}</td>
                      <td className="py-1 pr-2">{i === 0 ? 'Tagesentscheid' : VOTE_LABEL[c.cap]}</td>
                      <td className={`py-1 pr-2 ${monoTone(d.monotony)}`}>{fmtNum(d.monotony, 2)}</td>
                      <td className={`py-1 pr-2 ${d.acwr > 1.3 ? 'text-orange-300' : ''}`}>{fmtNum(d.acwr, 2)}</td>
                      <td className="py-1 pr-2">{c.form > 0 ? '+' : ''}{fmtNum(c.form, 1)} %</td>
                      <td className={`py-1 ${amp(c.level)}`}>{c.issues.length ? c.issues.join(' · ') : '✓'}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="mt-2 rounded border border-border px-3 py-2 text-xs flex flex-wrap items-center justify-between gap-2">
            <span>
              <span className="label mr-2">Hebel</span>
              {lever.best && (lever.best.conf < lever.base.conf || lever.best.mono < lever.base.mono) ? (
                <>
                  {lever.best.mode === 'rest' ? 'Ruhetag' : 'Last halbieren'} am <b>{wdOf(sim[lever.best.i].date)} {fmtDateShort(sim[lever.best.i].date).slice(0, 6)}</b> (statt {fmtNum(sim[lever.best.i].sim_load)} ESS):
                  Konflikte {lever.base.conf} → <b>{lever.best.conf}</b>, Verstöße gegen die Monotonie-Regel {lever.base.mono} → <b>{lever.best.mono}</b>.
                </>
              ) : (
                lever.base.conf === 0 && lever.base.mono === 0
                  ? 'Keine Konflikte und keine Verstöße gegen die Monotonie-Regel – ein einzelner hoher Monotonie-Wert ist kein Grund zum Eingreifen.'
                  : 'Kein einzelner Eingriff verbessert den Plan spürbar.'
              )}
            </span>
            {lever.best && (lever.best.conf < lever.base.conf || lever.best.mono < lever.base.mono) && (
              <button className="btn btn-primary text-xs px-2 py-1" onClick={applyLever}>Übernehmen</button>
            )}
          </div>
        </Box>

        <div className="grid gap-4 grid-cols-[minmax(0,1fr)] xl:grid-cols-2">
          {/* 2) Ziel rückwärts */}
          <Box title="2 · Ziel rückwärts rechnen" hint="Welche gleichmäßige Last braucht es bis zum Zieldatum? Rampe = CTL-Zuwachs pro Woche.">
            <div className="flex flex-wrap items-end gap-3 text-xs">
              <label className="flex flex-col gap-1">
                <span className="text-ink-dim">Ziel-CTL</span>
                <input type="number" className="input w-24" value={goalCtl} onChange={(e) => setGoalCtl(Number(e.target.value) || 0)} />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-ink-dim">Zieldatum</span>
                <input type="date" className="input" value={goalDate} onChange={(e) => setGoalDate(e.target.value)} />
              </label>
              <span className="text-ink-dim pb-1">Start: CTL {fmtNum(anchor.ctl)} am {fmtDateShort(anchor.date).slice(0, 6)}</span>
            </div>
            {goal && (
              <>
                <div className="mt-3 grid gap-2 grid-cols-2 md:grid-cols-4">
                  <Kpi label="Nötige Wochenlast" value={goal.daily > 0 ? `${fmtNum(goal.weekly)} ESS` : 'Ziel ohne Last'} sub={goal.daily > 0 ? `Ø ${fmtNum(goal.daily)} / Tag` : 'CTL fällt sonst darunter'} tone={amp(goal.lvl)} />
                  <Kpi label="Rampe / Woche" value={`${goal.ramp > 0 ? '+' : ''}${fmtNum(goal.ramp, 1)} %`} sub="Ziel 3–5 %, > 8 % zu steil" tone={goal.ramp > 8 ? 'text-orange-300' : goal.ramp > 5 ? 'text-yellow-300' : 'text-green-300'} />
                  <Kpi label="Max. ACWR" value={fmtNum(goal.maxAcwr, 2)} tone={goal.maxAcwr > 1.5 ? 'text-orange-300' : goal.maxAcwr > 1.3 ? 'text-yellow-300' : 'text-green-300'} />
                  <Kpi label="Form am Zieltag" value={`${goal.formEnd > 0 ? '+' : ''}${fmtNum(goal.formEnd, 1)} %`} sub="bei gleichmäßiger Last" />
                </div>
                <p className={`mt-2 text-xs ${amp(goal.lvl)}`}>
                  {goal.lvl === 'ok' ? 'Machbar' : goal.lvl === 'warn' ? 'Ambitioniert' : 'Zu steil'}: {goal.n} Tage,{' '}
                  {goal.daily > 0 ? `${fmtNum(goal.weekly)} ESS/Woche gegenüber aktuell ${fmtNum(curWeekly)} ESS/Woche im Plan.` : 'das Ziel wird schon mit Ruhe erreicht.'}
                  {goal.lvl === 'bad' ? ' Zieldatum später setzen oder Ziel-CTL senken.' : ''}
                </p>
                {goal.daily > 0 && (
                  <button className="mt-2 btn btn-ghost text-xs px-2 py-1" onClick={() => setTplWeekly(Math.round(goal.weekly / 10) * 10)}>
                    Wochenlast als Vorlagen-Ziel übernehmen
                  </button>
                )}
              </>
            )}
          </Box>

          {/* 4) Vorlagen */}
          <Box title="4 · Wochenmuster & Vorlagen" hint="Überträgt eine Lastverteilung auf alle offenen Tage (gesperrte und heute bleiben). Danach weiter per Regler anpassbar.">
            <div className="flex flex-wrap items-end gap-3 text-xs mb-3">
              <label className="flex flex-col gap-1">
                <span className="text-ink-dim">Wochenlast-Ziel (ESS)</span>
                <input type="number" step={10} className="input w-28" value={tplWeekly} onChange={(e) => setTplWeekly(Number(e.target.value) || 0)} />
              </label>
              <span className="text-ink-dim pb-1">Plan aktuell Ø {fmtNum(curWeekly)} ESS/Woche</span>
            </div>
            <div className="grid gap-2 sm:grid-cols-2">
              {PATTERNS.map((p) => (
                <button key={p.id} className="text-left rounded border border-border px-3 py-2 hover:border-accent" onClick={() => applyPattern(p)}>
                  <div className="text-xs font-semibold">{p.label}</div>
                  <div className="text-2xs text-ink-dim">{p.hint}</div>
                  <div className="mt-1 text-2xs tnum text-ink-muted">{Object.entries(p.w).map(([k, v]) => `${k} ${Math.round((v * tplWeekly) / Object.values(p.w).reduce((a, b) => a + b, 0))}`).join(' · ')}</div>
                </button>
              ))}
              <button className="text-left rounded border border-border px-3 py-2 hover:border-accent" onClick={applyDeload}>
                <div className="text-xs font-semibold">Entlastungswoche</div>
                <div className="text-2xs text-ink-dim">Nächste 7 Tage −35 %, Woche 2 bleibt</div>
              </button>
            </div>
            <p className="mt-2 text-2xs text-ink-dim">Tapering vor einem Termin: bei den Events (Punkt 6).</p>
          </Box>
        </div>

        {/* 3) Szenarien */}
        <Box
          title="3 · Szenarien vergleichen"
          hint="Aktuellen Planstand als A, B oder C merken (nur in diesem Browser). In die Timeline kommt erst, was du über „In timeline speichern“ sicherst."
          right={
            <div className="flex gap-1">
              {(['A', 'B', 'C'] as const).map((n) => (
                <button key={n} className="btn btn-ghost text-xs px-2 py-1" onClick={() => saveScen(n)}>Als {n} merken</button>
              ))}
            </div>
          }
        >
          <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(0,1fr)_minmax(0,1fr)]">
            <div className="overflow-x-auto">
              <table className="w-full text-xs tnum">
                <thead>
                  <tr className="text-ink-dim text-left">
                    {['', 'Σ Last', 'CTL Ende', 'max ACWR', 'max Mono', 'Form Ende', 'Konflikte', ''].map((h, i) => (
                      <th key={i} className="py-1 pr-2 font-normal">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {scenRows.map((s) => (
                    <tr key={s.name} className="border-t border-border">
                      <td className="py-1 pr-2 font-semibold" style={{ color: SC_COL[s.name] }}>{s.name}</td>
                      <td className="py-1 pr-2">{fmtNum(s.sum)}</td>
                      <td className="py-1 pr-2">{fmtNum(s.ctlEnd)}</td>
                      <td className={`py-1 pr-2 ${s.maxAcwr > 1.3 ? 'text-orange-300' : ''}`}>{fmtNum(s.maxAcwr, 2)}</td>
                      <td className={`py-1 pr-2 ${monoTone(s.maxMono)}`}>{fmtNum(s.maxMono, 2)}</td>
                      <td className="py-1 pr-2">{s.formEnd > 0 ? '+' : ''}{fmtNum(s.formEnd, 1)} %</td>
                      <td className={`py-1 pr-2 ${s.conf ? 'text-orange-300' : 'text-green-300'}`}>{s.conf}</td>
                      <td className="py-1 whitespace-nowrap">
                        {s.name === 'Kira' && (
                          <button className="text-accent hover:underline" onClick={() => applyState(s.st, 'Kira-Wochenplan geladen (noch nicht gespeichert).')}>laden</button>
                        )}
                        {s.name !== 'Aktuell' && s.name !== 'Kira' && (
                          <>
                            <button className="text-accent hover:underline mr-2" onClick={() => applyState(s.st, `Szenario ${s.name} geladen (noch nicht gespeichert).`)}>laden</button>
                            <button className="text-ink-dim hover:text-red-300" onClick={() => delScen(s.name)}>×</button>
                          </>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {kiPlan && (
                <p className="mt-2 text-2xs text-ink-dim">
                  Kira = Wochenplan aus dem KI-Report ({kiPlan.days} Tage); gesperrte Tage und die übrigen Tage wie aktueller Plan.
                </p>
              )}
              {!scens.length && <p className="mt-2 text-2xs text-ink-dim">Noch keine Szenarien gemerkt.</p>}
            </div>
            <div style={{ height: 200 }}>
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={scenChart} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                  <CartesianGrid stroke={GRID} vertical={false} />
                  <XAxis dataKey="d" tickLine={false} tick={{ fontSize: 9 }} minTickGap={16} />
                  <YAxis width={40} tickLine={false} tick={{ fontSize: 10 }} domain={['auto', 'auto']} />
                  <Tooltip formatter={(v: any, n: string) => [fmtNum(v), `CTL ${n}`]} />
                  <Legend wrapperStyle={{ fontSize: 10 }} />
                  {scenRows.map((s) => (
                    <Line key={s.name} dataKey={s.name} stroke={SC_COL[s.name]} strokeWidth={s.name === 'Aktuell' ? 2.2 : 1.5} dot={false} isAnimationActive={false} />
                  ))}
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>
        </Box>

        {/* 5) Horizont */}
        <Box
          title={`5 · Horizont ${weeks} Wochen`}
          hint="Woche 1–2 = aktueller Plan, danach Fortschreibung mit Wochenmuster, Steigerung und optional 3:1-Zyklus (jede 4. Woche −35 %)."
          right={
            <div className="flex flex-wrap items-center gap-2 text-xs">
              <div className="inline-flex bg-bg rounded border border-border p-0.5">
                {[4, 8].map((w) => (
                  <button key={w} onClick={() => setWeeks(w as 4 | 8)} className={`px-2 py-1 rounded ${weeks === w ? 'bg-bg-subtle text-ink' : 'text-ink-muted'}`}>{w} Wo.</button>
                ))}
              </div>
              <label className="flex items-center gap-1 text-ink-muted">
                +<input type="number" className="input w-14" value={rampPct} step={1} onChange={(e) => setRampPct(Number(e.target.value) || 0)} />%/Wo.
              </label>
              <label className="flex items-center gap-1 text-ink-muted">
                <input type="checkbox" checked={cycle31} onChange={(e) => setCycle31(e.target.checked)} className="accent-accent" /> 3:1
              </label>
              <select className="input text-xs" value={hzPattern} onChange={(e) => setHzPattern(e.target.value)}>
                {PATTERNS.map((p) => (
                  <option key={p.id} value={p.id}>{p.label}</option>
                ))}
              </select>
            </div>
          }
        >
          {horizon && (
            <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(0,1.4fr)_minmax(0,1fr)]">
              <div style={{ height: 240 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={horizon.wk} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                    <CartesianGrid stroke={GRID} vertical={false} />
                    <XAxis dataKey="label" tickLine={false} tick={{ fontSize: 9 }} />
                    <YAxis yAxisId="l" width={44} tickLine={false} tick={{ fontSize: 10 }} />
                    <YAxis yAxisId="c" orientation="right" width={40} tickLine={false} tick={{ fontSize: 10 }} domain={['auto', 'auto']} />
                    <Tooltip formatter={(v: any, n: string) => [n === 'Form %' ? `${fmtNum(v, 1)} %` : fmtNum(v), n]} />
                    <Legend wrapperStyle={{ fontSize: 10 }} />
                    <Bar yAxisId="l" dataKey="Wochenlast" fill="#7dd3fc" fillOpacity={0.45} isAnimationActive={false} />
                    <Line yAxisId="c" dataKey="CTL" stroke="#22c55e" strokeWidth={2} dot={{ r: 2 }} isAnimationActive={false} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs tnum">
                  <thead>
                    <tr className="text-ink-dim text-left">
                      {['Woche', 'Phase', 'Last', 'Rampe', 'Form', 'max ACWR'].map((h) => (
                        <th key={h} className="py-1 pr-2 font-normal">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {horizon.wk.map((w) => (
                      <tr key={w.label} className="border-t border-border">
                        <td className="py-1 pr-2 whitespace-nowrap">{w.label}</td>
                        <td className={`py-1 pr-2 ${w.phase === 'Entlastung' ? 'text-accent' : 'text-ink-muted'}`}>{w.phase}</td>
                        <td className="py-1 pr-2">{fmtNum(w.Wochenlast)}</td>
                        <td className={`py-1 pr-2 ${w.Rampe > 8 ? 'text-orange-300' : w.Rampe > 5 ? 'text-yellow-300' : w.Rampe < 0 ? 'text-accent' : 'text-green-300'}`}>{w.Rampe > 0 ? '+' : ''}{fmtNum(w.Rampe, 1)} %</td>
                        <td className="py-1 pr-2">{w['Form %'] > 0 ? '+' : ''}{fmtNum(w['Form %'], 1)} %</td>
                        <td className={`py-1 ${w.maxAcwr > 1.3 ? 'text-orange-300' : ''}`}>{fmtNum(w.maxAcwr, 2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </Box>

        <div className="grid gap-4 grid-cols-[minmax(0,1fr)] xl:grid-cols-2">
          {/* 6) Events */}
          <Box title="6 · Events & Formspitze" hint="Termine merken (nur in diesem Browser). Prognose aus Plan (≤ 14 T) bzw. Horizont. Ziel am Event: Form +5 bis +15 %.">
            <div className="flex flex-wrap items-end gap-2 text-xs">
              <input type="date" className="input" value={evDate} onChange={(e) => setEvDate(e.target.value)} />
              <input className="input flex-1 min-w-[140px]" placeholder="z. B. Gravel-Rennen" value={evLabel} onChange={(e) => setEvLabel(e.target.value)} />
              <button
                className="btn btn-primary text-xs px-2 py-1"
                disabled={!evLabel.trim()}
                onClick={() => {
                  setEvents([...events, { id: String(Date.now()), date: evDate, label: evLabel.trim() }]);
                  setEvLabel('');
                }}
              >
                Hinzufügen
              </button>
            </div>
            <div className="mt-3 space-y-2">
              {!events.length && <p className="text-2xs text-ink-dim">Noch keine Events.</p>}
              {events.map((ev) => {
                const info = evInfo(ev);
                const n = diffDays(localIsoDate(), ev.date);
                const taperStart = addDays(ev.date, -7);
                return (
                  <div key={ev.id} className="rounded border border-border px-3 py-2 text-xs">
                    <div className="flex flex-wrap items-baseline justify-between gap-2">
                      <span className="font-semibold">★ {ev.label}</span>
                      <span className="tnum text-ink-dim">
                        {wdOf(ev.date)} {fmtDateShort(ev.date)} · {n >= 0 ? `in ${n} T` : 'vorbei'}
                        <button className="ml-2 hover:text-red-300" onClick={() => setEvents(events.filter((e) => e.id !== ev.id))}>×</button>
                      </span>
                    </div>
                    {info ? (
                      <div className="mt-1 tnum">
                        CTL {fmtNum(info.ctl)} · Form{' '}
                        <b className={info.form >= 5 && info.form <= 15 ? 'text-green-300' : info.form < 0 ? 'text-orange-300' : 'text-yellow-300'}>
                          {info.form > 0 ? '+' : ''}{fmtNum(info.form, 1)} %
                        </b>{' '}
                        <span className="text-ink-dim">({info.src})</span>
                        <div className="mt-0.5 text-ink-muted">
                          {info.form >= 5 && info.form <= 15
                            ? 'Formfenster passt.'
                            : info.form < 5
                              ? `Zu müde für eine Formspitze – ab ${wdOf(taperStart)} ${fmtDateShort(taperStart).slice(0, 6)} reduzieren (Tapering).`
                              : 'Sehr frisch – evtl. Fitness verschenkt, Tapering später beginnen.'}
                        </div>
                        {info.inPlan && info.form < 5 && (
                          <button className="mt-1 btn btn-ghost text-xs px-2 py-0.5" onClick={() => applyTaper(ev)}>Tapering anwenden</button>
                        )}
                      </div>
                    ) : (
                      <div className="mt-1 text-ink-dim">Außerhalb des Horizonts ({weeks} Wochen).</div>
                    )}
                  </div>
                );
              })}
            </div>
          </Box>

          {/* 7) Erholungsregel & 80/20 */}
          <Box title="7 · Erholungsregel & Intensität des Plans">
            <div className="text-xs space-y-2">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-ink-muted">Abstand zwischen harten Einheiten (≥ 200 ESS oder Quality):</span>
                <div className="inline-flex bg-bg rounded border border-border p-0.5">
                  {[1, 2, 3].map((g) => (
                    <button key={g} onClick={() => setRecGap(g)} className={`px-2 py-0.5 rounded tnum ${recGap === g ? 'bg-bg-subtle text-ink' : 'text-ink-muted'}`}>≥ {g} T</button>
                  ))}
                </div>
              </div>
              <p className="text-2xs text-ink-dim">
                {recInfo
                  ? `Aus deinen Daten: nach ${recInfo.n} harten Einheiten war die Readiness im Schnitt nach ${fmtNum(recInfo.avg, 1)} Tagen wieder normal → Regel ≥ ${Math.max(1, Math.min(3, Math.round(recInfo.avg)))} T voreingestellt.`
                  : 'Persönliche Erholungszeit wird aus der Historie berechnet …'}
              </p>
              <div>
                <span className="text-ink-muted">Harte Tage im Plan: </span>
                {hardDays.length ? (
                  hardDays.map((x, k) => {
                    const gap = k > 0 ? x.i - hardDays[k - 1].i : null;
                    return (
                      <span key={x.d.date} className={`tnum mr-2 ${gap != null && gap < recGap ? 'text-orange-300' : ''}`}>
                        {wdOf(x.d.date)} {fmtDateShort(x.d.date).slice(0, 6)} ({fmtNum(x.d.sim_load)}){gap != null ? ` · +${gap} T` : ''}
                      </span>
                    );
                  })
                ) : (
                  <span className="text-ink-dim">keine</span>
                )}
              </div>
              <div className="pt-1">
                <div className="flex justify-between text-2xs text-ink-dim mb-1">
                  <span>Intensitätsverteilung des Plans (Last nach Zone)</span>
                  <span>Ziel ≥ 80 % Z1–Z2</span>
                </div>
                <div className="flex h-3 rounded overflow-hidden border border-border">
                  <div style={{ width: `${dist.low}%`, background: '#22c55e' }} title={`Z1–Z2 ${dist.low} %`} />
                  <div style={{ width: `${dist.mid}%`, background: '#eab308' }} title={`Z3 ${dist.mid} %`} />
                  <div style={{ width: `${dist.high}%`, background: '#ef4444' }} title={`Z4+ ${dist.high} %`} />
                </div>
                <div className="mt-1 tnum">
                  Z1–Z2 <b className={dist.low >= 80 ? 'text-green-300' : 'text-yellow-300'}>{dist.low} %</b> · Z3 <b className={dist.mid >= 20 ? 'text-orange-300' : ''}>{dist.mid} %</b> · Z4+ <b>{dist.high} %</b>
                  <span className="text-ink-dim ml-2">
                    {dist.mid >= 20 ? 'Viel Z3 – Graubereich.' : dist.low >= 80 && dist.high < 5 ? 'Locker verteilt, kaum hohe Intensität.' : dist.low >= 80 ? 'Passt zu 80/20.' : 'Lockerer Anteil unter 80 %.'}
                  </span>
                </div>
              </div>
            </div>
          </Box>
        </div>

        {/* 8) Spielraum */}
        <Box title="8 · Spielraum ±10 % Last" hint="Wie empfindlich reagiert der Plan, wenn du an allen offenen Tagen 10 % weniger oder mehr machst?">
          <div className="grid gap-2 grid-cols-2 md:grid-cols-4 mb-2">
            <Kpi label="CTL Ende" value={`${fmtNum(band.ctlLo)} – ${fmtNum(band.ctlHi)}`} sub={`Plan ${fmtNum(sim[sim.length - 1].ctl)}`} />
            <Kpi label="max ACWR" value={`${fmtNum(band.acwrLo, 2)} – ${fmtNum(band.acwrHi, 2)}`} tone={band.acwrHi > 1.3 ? 'text-orange-300' : ''} />
            <Kpi label="Konflikte bei +10 %" value={band.confHi} sub={`Plan ${confCount}`} tone={band.confHi > confCount ? 'text-orange-300' : ''} />
            <Kpi label="Puffer" value={band.acwrHi <= 1.3 && band.confHi <= confCount ? 'robust' : 'knapp'} tone={band.acwrHi <= 1.3 && band.confHi <= confCount ? 'text-green-300' : 'text-yellow-300'} sub="bei +10 % Mehrlast" />
          </div>
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={band.chart} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="d" tickLine={false} tick={{ fontSize: 9 }} minTickGap={16} />
                <YAxis yAxisId="c" width={40} tickLine={false} tick={{ fontSize: 10 }} domain={['auto', 'auto']} />
                <YAxis yAxisId="a" orientation="right" width={34} tickLine={false} tick={{ fontSize: 10 }} domain={[0.6, 1.8]} />
                <ReferenceLine yAxisId="a" y={1.3} stroke="#f97316" strokeDasharray="4 4" />
                <Tooltip formatter={(v: any, n: string) => [n.startsWith('ACWR') ? fmtNum(v, 2) : fmtNum(v), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Line yAxisId="c" dataKey="CTL −10 %" stroke="#22c55e" strokeOpacity={0.45} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
                <Line yAxisId="c" dataKey="CTL" stroke="#22c55e" strokeWidth={2} dot={false} isAnimationActive={false} />
                <Line yAxisId="c" dataKey="CTL +10 %" stroke="#22c55e" strokeOpacity={0.45} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
                <Line yAxisId="a" dataKey="ACWR −10 %" stroke="#a855f7" strokeOpacity={0.45} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
                <Line yAxisId="a" dataKey="ACWR" stroke="#a855f7" strokeWidth={1.8} dot={false} isAnimationActive={false} />
                <Line yAxisId="a" dataKey="ACWR +10 %" stroke="#a855f7" strokeOpacity={0.45} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </Box>
      </div>
    </Panel>
  );
}
