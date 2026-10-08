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
import { ApiError, fetchChartData, toNum } from '../lib/api';
import { fmtDateShort, fmtNum } from '../lib/format';
import { ErrorBox, Panel, Skeleton } from '../components/UI';
import {
  buildDays,
  demandOf,
  fmtHm,
  LEVEL_CELL,
  LEVEL_TEXT,
  localIsoDate,
  RECOVERY_FETCH_METRICS,
  SLEEP_GOAL_H,
  VOTE_LABEL,
  VOTE_LEVEL,
  VOTE_ORDER,
  type RecoveryDay,
} from './RecoveryPanel';
import { MonthReport } from './MonthReport';

// ---------------------------------------------------------------------------
// Berichte: Monat · Woche · Erholung · Jahr · Schlaf · Ernährung · Ausfälle
// Alles regelbasiert aus der Timeline (keine KI).
// ---------------------------------------------------------------------------

const EXTRA_METRICS = ['kcal_in', 'kcal_out', 'carb_g', 'protein_g', 'fat_g', 'distance_km_day', 'elev_m_day', 'fbCTL_obs'] as const;
type Extra = { kcalIn: number | null; kcalOut: number | null; carb: number | null; protein: number | null; fat: number | null; dist: number | null; elev: number | null; ctl: number | null };
type D = RecoveryDay & { x: Extra };

type Sub = 'month' | 'week' | 'recovery' | 'year' | 'sleep' | 'nutrition' | 'outage';
const SUBS: { id: Sub; label: string }[] = [
  { id: 'month', label: 'Monat' },
  { id: 'week', label: 'Woche' },
  { id: 'recovery', label: 'Erholung' },
  { id: 'year', label: 'Jahr' },
  { id: 'sleep', label: 'Schlaf' },
  { id: 'nutrition', label: 'Ernährung' },
  { id: 'outage', label: 'Ausfälle' },
];

const TITLES: Record<Sub, string> = {
  month: 'Monatsbericht',
  week: 'Wochenbericht',
  recovery: 'Erholung nach harten Einheiten',
  year: 'Jahresrückblick · 12 Monate',
  sleep: 'Schlafbericht',
  nutrition: 'Ernährungsbericht',
  outage: 'Krankheit & Datenlücken',
};
const GRID = '#1f2731';
const MONTHS_DE = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];
const WD = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];

// ---------------------------------------------------------------------------
// Helfer
// ---------------------------------------------------------------------------
const avg = (xs: (number | null | undefined)[]) => {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};
const sum = (xs: (number | null | undefined)[]) => xs.reduce<number>((a, b) => a + (b ?? 0), 0);
const sgn = (v: number | null, d = 0, unit = '') => (v == null ? '—' : `${v > 0 ? '+' : v < 0 ? '−' : '±'}${fmtNum(Math.abs(v), d)}${unit}`);
const wdIdx = (iso: string) => (new Date(iso + 'T12:00:00').getDay() + 6) % 7; // Mo=0
const isSick = (d: RecoveryDay) => (d.actual.sport || '').trim().toLowerCase() === 'krank';
const hasMorning = (d: RecoveryDay) => d.readiness != null || d.sleepH != null || d.hrv != null;
const normSport = (s: string) => {
  const t = (s || '').trim().toLowerCase();
  if (!t || t === 'off' || t === 'krank') return '';
  if (t.startsWith('run') || t === 'laufen') return 'Run';
  if (t.includes('bike') || t.includes('rad')) return 'Bike';
  if (t.includes('hike') || t === 'walk') return 'Hike/Walk';
  if (t.startsWith('ski')) return 'Ski';
  return 'Sonstige';
};
const SPORT_COLORS: Record<string, string> = { Bike: '#7dd3fc', Run: '#f97316', 'Hike/Walk': '#22c55e', Ski: '#a855f7', Sonstige: '#94a3b8' };

function isoWeek(iso: string) {
  const d = new Date(iso + 'T12:00:00');
  const day = (d.getDay() + 6) % 7;
  const mon = new Date(d);
  mon.setDate(d.getDate() - day);
  const th = new Date(mon);
  th.setDate(mon.getDate() + 3);
  const y = th.getFullYear();
  const jan4 = new Date(y, 0, 4, 12);
  const kw = 1 + Math.round(((th.getTime() - jan4.getTime()) / 86400000 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
  return { key: `${y}-W${String(kw).padStart(2, '0')}`, kw, year: y, monday: localIsoDate(mon) };
}

function pearson(xs: number[], ys: number[]): number | null {
  const n = xs.length;
  if (n < 8) return null;
  const mx = xs.reduce((a, b) => a + b, 0) / n;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let c = 0, vx = 0, vy = 0;
  for (let i = 0; i < n; i++) {
    c += (xs[i] - mx) * (ys[i] - my);
    vx += (xs[i] - mx) ** 2;
    vy += (ys[i] - my) ** 2;
  }
  return vx && vy ? c / Math.sqrt(vx * vy) : null;
}

function Kpi({ label, value, sub, tone }: { label: string; value: ReactNode; sub?: ReactNode; tone?: string }) {
  return (
    <div className="panel-raised p-2.5 min-w-0">
      <div className="text-2xs text-ink-dim">{label}</div>
      <div className={`tnum text-sm font-semibold whitespace-nowrap ${tone || ''}`}>{value}</div>
      {sub != null && <div className="text-2xs text-ink-dim tnum">{sub}</div>}
    </div>
  );
}
function Box({ title, right, children }: { title: string; right?: ReactNode; children: ReactNode }) {
  return (
    <div className="panel-raised p-3 min-w-0">
      <div className="flex flex-wrap items-baseline justify-between gap-2 mb-2">
        <div className="label">{title}</div>
        {right}
      </div>
      {children}
    </div>
  );
}
function Chips<T extends string | number>({ items, value, onChange, fmt }: { items: T[]; value: T; onChange: (v: T) => void; fmt?: (v: T) => string }) {
  return (
    <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5">
      {items.map((v) => (
        <button key={String(v)} onClick={() => onChange(v)} className={`px-2 py-1 text-xs rounded tnum ${value === v ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}>
          {fmt ? fmt(v) : String(v)}
        </button>
      ))}
    </div>
  );
}
function Facts({ items }: { items: string[] }) {
  if (!items.length) return null;
  return (
    <div className="rounded border border-border px-3 py-2">
      <div className="label mb-1">Fazit</div>
      <ul className="text-sm space-y-0.5">
        {items.map((t) => (
          <li key={t}>· {t}</li>
        ))}
      </ul>
      <p className="mt-1 text-2xs text-ink-dim">Regelbasiert aus der Timeline, keine KI-Bewertung.</p>
    </div>
  );
}
const ttFmt = (unit = '', d = 0) => (v: any, n: string) => [v == null ? '—' : `${fmtNum(Number(v), d)}${unit}`, n];

// ---------------------------------------------------------------------------
// 1) Wochenbericht
// ---------------------------------------------------------------------------
type WeekStats = ReturnType<typeof weekStats>;
function weekStats(days: D[], key: string) {
  const w = days.filter((d) => isoWeek(d.date).key === key);
  const { kw, monday, year } = isoWeek(w[0]?.date || localIsoDate());
  const sleep = w.filter((d) => d.sleepH != null && d.sleepH > 0);
  const checked = w.filter((d) => d.vote && d.essDay != null && (d.activityDone || d.essDay === 0));
  const misses = checked.filter((d) => VOTE_ORDER.indexOf(demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn)) > VOTE_ORDER.indexOf(d.vote!));
  const best = (f: (d: D) => number | null, min = false) =>
    w.reduce<D | null>((b, d) => {
      const v = f(d);
      if (v == null) return b;
      const bv = b ? f(b)! : null;
      return bv == null || (min ? v < bv : v > bv) ? d : b;
    }, null);
  const ctlEnd = [...w].reverse().find((d) => d.ctlPost != null)?.ctlPost ?? null;
  return {
    key,
    kw,
    year,
    monday,
    n: w.length,
    days: w,
    ess: sum(w.map((d) => d.essDay)),
    sessions: w.filter((d) => (d.essDay ?? 0) > 0).length,
    readiness: avg(w.map((d) => d.readiness)),
    sleepH: avg(sleep.map((d) => d.sleepH)),
    nightsGoal: sleep.filter((d) => d.sleepH! >= SLEEP_GOAL_H - 1e-9).length,
    nights: sleep.length,
    hrv: avg(w.map((d) => d.hrv)),
    rhr: avg(w.map((d) => d.rhr)),
    mono: [...w].reverse().find((d) => d.monotonyPost != null)?.monotonyPost ?? null,
    ctlEnd,
    votes: VOTE_ORDER.slice()
      .reverse()
      .map((v) => ({ v, n: w.filter((d) => d.vote === v).length }))
      .filter((x) => x.n > 0),
    checked: checked.length,
    ok: checked.length - misses.length,
    misses,
    peak: best((d) => d.essDay),
    bestR: best((d) => d.readiness),
    longS: best((d) => d.sleepH),
    shortS: best((d) => (d.sleepH && d.sleepH > 0 ? d.sleepH : null), true),
  };
}

function WeekReport({ days }: { days: D[] }) {
  const keys = useMemo(() => [...new Set(days.map((d) => isoWeek(d.date).key))].sort(), [days]);
  const [idx, setIdx] = useState(keys.length - 1);
  useEffect(() => setIdx(keys.length - 1), [keys.length]);
  const key = keys[Math.max(0, Math.min(idx, keys.length - 1))];
  const cur = useMemo(() => weekStats(days, key), [days, key]);
  const prev = useMemo(() => (idx > 0 ? weekStats(days, keys[idx - 1]) : null), [days, keys, idx]);
  const table = useMemo(() => keys.slice(-12).map((k) => weekStats(days, k)).reverse(), [days, keys]);
  const end = new Date(cur.monday + 'T12:00:00');
  end.setDate(end.getDate() + 6);
  const facts: string[] = [];
  if (prev && prev.ess > 0) facts.push(`Last ${sgn(((cur.ess - prev.ess) / prev.ess) * 100, 0, ' %')} ggü. Vorwoche (${fmtNum(prev.ess)} → ${fmtNum(cur.ess)} ESS)${cur.n < 7 ? ' – Woche läuft noch' : ''}.`);
  if (cur.checked) facts.push(`${cur.ok} von ${cur.checked} Tagen passten zum Votum.`);
  if (cur.mono != null && cur.mono > 2) facts.push(`Monotonie zum Wochenende ${fmtNum(cur.mono, 2)} – zu gleichförmig.`);
  if (cur.sleepH != null && cur.sleepH < SLEEP_GOAL_H - 0.25) facts.push(`Schlaf Ø ${fmtHm(cur.sleepH)} h, nur ${cur.nightsGoal}/${cur.nights} Nächte ≥ 7:30 h.`);
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="text-base font-semibold tnum">
          KW {cur.kw} / {cur.year}
          <span className="ml-2 text-2xs text-ink-dim font-normal">
            {fmtDateShort(cur.monday).slice(0, 6)}–{fmtDateShort(localIsoDate(end)).slice(0, 6)}
            {cur.n < 7 ? ` · laufend (${cur.n} Tage)` : ''}
          </span>
        </div>
        <div className="flex gap-1">
          <button className="btn btn-ghost text-xs px-2 py-1" disabled={idx <= 0} onClick={() => setIdx((i) => i - 1)}>◂ Vorwoche</button>
          <button className="btn btn-ghost text-xs px-2 py-1" disabled={idx >= keys.length - 1} onClick={() => setIdx((i) => i + 1)}>Nächste ▸</button>
        </div>
      </div>
      <div className="grid gap-3 grid-cols-2 md:grid-cols-4 xl:grid-cols-8">
        <Kpi label="Last" value={`${fmtNum(cur.ess)} ESS`} sub={prev ? `Vorwoche ${fmtNum(prev.ess)}` : undefined} />
        <Kpi label="Einheiten · Ruhe" value={`${cur.sessions} · ${cur.n - cur.sessions}`} />
        <Kpi label="Ø Readiness" value={fmtNum(cur.readiness)} sub={prev ? sgn(cur.readiness != null && prev.readiness != null ? cur.readiness - prev.readiness : null) : undefined} />
        <Kpi label="Ø Schlaf" value={`${fmtHm(cur.sleepH)} h`} sub={`${cur.nightsGoal}/${cur.nights} ≥ 7:30`} />
        <Kpi label="Ø HRV" value={`${fmtNum(cur.hrv)} ms`} />
        <Kpi label="Ø RHR" value={`${fmtNum(cur.rhr, 1)} bpm`} />
        <Kpi label="Monotonie (Ende)" value={fmtNum(cur.mono, 2)} tone={cur.mono != null && cur.mono > 2 ? 'text-orange-300' : cur.mono != null && cur.mono > 1.6 ? 'text-yellow-300' : ''} />
        <Kpi label="CTL (Ende)" value={fmtNum(cur.ctlEnd)} sub={prev?.ctlEnd ? sgn(cur.ctlEnd != null ? cur.ctlEnd - prev.ctlEnd : null) : undefined} />
      </div>
      <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-3">
        <Box title="Tage">
          <div className="space-y-1 text-xs tnum">
            {cur.days.map((d) => {
              const lvl = d.vote ? VOTE_LEVEL[d.vote] : 'leer';
              return (
                <div key={d.date} className="flex items-center justify-between gap-2">
                  <span className="w-14 text-ink-muted">{WD[wdIdx(d.date)]} {fmtDateShort(d.date).slice(0, 6)}</span>
                  <span className="flex-1 truncate">{(d.essDay ?? 0) > 0 ? `${fmtNum(d.essDay)} ESS ${d.actual.sport} ${d.actual.zone}` : isSick(d) ? 'krank' : 'Ruhe'}</span>
                  {d.vote && <span className={`chip border text-2xs ${LEVEL_CELL[lvl]} ${LEVEL_TEXT[lvl]} border-border`}>{VOTE_LABEL[d.vote]}</span>}
                </div>
              );
            })}
          </div>
        </Box>
        <Box title="Voten & Einhaltung">
          <div className="flex flex-wrap gap-1">
            {cur.votes.map(({ v, n }) => (
              <span key={v} className={`chip border text-2xs ${LEVEL_CELL[VOTE_LEVEL[v]]} ${LEVEL_TEXT[VOTE_LEVEL[v]]} border-border`}>{n}× {VOTE_LABEL[v]}</span>
            ))}
          </div>
          <div className="mt-2 text-xs flex justify-between">
            <span className="text-ink-muted">Passend zum Votum</span>
            <span className={`tnum font-semibold ${cur.checked && cur.ok / cur.checked >= 0.7 ? 'text-green-300' : 'text-orange-300'}`}>{cur.checked ? `${cur.ok} / ${cur.checked}` : '—'}</span>
          </div>
          {cur.misses.length > 0 && (
            <ul className="mt-1 text-2xs text-ink-dim space-y-0.5">
              {cur.misses.map((d) => (
                <li key={d.date} className="tnum">
                  {fmtDateShort(d.date).slice(0, 6)}: {fmtNum(d.essDay)} ESS {d.actual.sport} bei {VOTE_LABEL[d.vote!]}
                </li>
              ))}
            </ul>
          )}
        </Box>
        <Box title="Höhen & Tiefen">
          <dl className="text-xs space-y-1 tnum">
            {[
              ['Höchste Last', cur.peak, cur.peak ? `${fmtNum(cur.peak.essDay)} ESS ${cur.peak.actual.sport}` : ''],
              ['Beste Readiness', cur.bestR, cur.bestR ? fmtNum(cur.bestR.readiness) : ''],
              ['Längster Schlaf', cur.longS, cur.longS ? `${fmtHm(cur.longS.sleepH)} h` : ''],
              ['Kürzester Schlaf', cur.shortS, cur.shortS ? `${fmtHm(cur.shortS.sleepH)} h` : ''],
            ].map(([l, d, v]: any) => (
              <div key={l} className="flex justify-between gap-2">
                <dt className="text-ink-muted">{l}</dt>
                <dd>{d ? `${v} · ${WD[wdIdx(d.date)]}` : '—'}</dd>
              </div>
            ))}
          </dl>
        </Box>
      </div>
      <Facts items={facts} />
      <Box title="Archiv · letzte 12 Wochen" right={<span className="text-2xs text-ink-dim">Klick = Woche öffnen</span>}>
        <div className="overflow-x-auto">
          <table className="w-full text-xs tnum">
            <thead>
              <tr className="text-ink-dim text-left">
                <th className="py-1 pr-2 font-normal">KW</th>
                <th className="py-1 pr-2 font-normal text-right">Last</th>
                <th className="py-1 pr-2 font-normal text-right">Einh.</th>
                <th className="py-1 pr-2 font-normal text-right">Readiness</th>
                <th className="py-1 pr-2 font-normal text-right">Schlaf</th>
                <th className="py-1 pr-2 font-normal text-right">Monotonie</th>
                <th className="py-1 font-normal text-right">Votum ok</th>
              </tr>
            </thead>
            <tbody>
              {table.map((w) => (
                <tr key={w.key} className={`border-t border-border cursor-pointer hover:bg-bg-subtle ${w.key === key ? 'bg-bg-subtle' : ''}`} onClick={() => setIdx(keys.indexOf(w.key))}>
                  <td className="py-1 pr-2">KW {w.kw}{w.n < 7 ? '*' : ''}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(w.ess)}</td>
                  <td className="py-1 pr-2 text-right">{w.sessions}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(w.readiness)}</td>
                  <td className="py-1 pr-2 text-right">{fmtHm(w.sleepH)}</td>
                  <td className={`py-1 pr-2 text-right ${w.mono != null && w.mono > 2 ? 'text-orange-300' : ''}`}>{fmtNum(w.mono, 2)}</td>
                  <td className="py-1 text-right">{w.checked ? `${w.ok}/${w.checked}` : '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Box>
    </div>
  );
}

// ---------------------------------------------------------------------------
// 3) Erholung nach harten Einheiten
// ---------------------------------------------------------------------------
function RecoveryAfterHard({ days }: { days: D[] }) {
  const [minLoad, setMinLoad] = useState<150 | 200 | 250>(200);
  const res = useMemo(() => {
    const items: { d: D; group: string; base: number; r: (number | null)[]; hrv: (number | null)[]; rhr: (number | null)[]; daysTo: number | null }[] = [];
    for (let i = 7; i < days.length - 1; i++) {
      const d = days[i];
      const e = d.essDay ?? 0;
      const dem = demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
      if (e < minLoad && !(dem === 'QUALITY' && e >= 120)) continue;
      const prior = days.slice(Math.max(0, i - 28), i).map((x) => x.readiness).filter((x): x is number => x != null);
      if (prior.length < 7) continue;
      const base = prior.reduce((a, b) => a + b, 0) / prior.length;
      const priorH = avg(days.slice(Math.max(0, i - 28), i).map((x) => x.hrv));
      const priorR = avg(days.slice(Math.max(0, i - 28), i).map((x) => x.rhr));
      const next = [1, 2, 3].map((k) => days[i + k]);
      const r = next.map((n) => (n?.readiness != null ? n.readiness - base : null));
      const hrv = next.map((n) => (n?.hrv != null && priorH != null ? n.hrv - priorH : null));
      const rhr = next.map((n) => (n?.rhr != null && priorR != null ? n.rhr - priorR : null));
      let daysTo: number | null = null;
      for (let k = 0; k < 3; k++) if (r[k] != null && r[k]! >= -2) { daysTo = k + 1; break; }
      if (daysTo == null && r.some((x) => x != null)) daysTo = 4; // > 3 Tage
      items.push({ d, group: normSport(d.actual.sport) || 'Sonstige', base, r, hrv, rhr, daysTo });
    }
    const groups = ['Alle', ...[...new Set(items.map((x) => x.group))]];
    const rows = groups.map((g) => {
      const its = g === 'Alle' ? items : items.filter((x) => x.group === g);
      return {
        g,
        n: its.length,
        load: avg(its.map((x) => x.d.essDay)),
        r1: avg(its.map((x) => x.r[0])),
        r2: avg(its.map((x) => x.r[1])),
        r3: avg(its.map((x) => x.r[2])),
        h1: avg(its.map((x) => x.hrv[0])),
        p1: avg(its.map((x) => x.rhr[0])),
        to: avg(its.map((x) => x.daysTo)),
      };
    });
    return { items, rows };
  }, [days, minLoad]);
  const curve = ['T+1', 'T+2', 'T+3'].map((t, k) => {
    const o: any = { t };
    for (const r of res.rows) if (r.n >= 3) o[r.g] = [r.r1, r.r2, r.r3][k];
    return o;
  });
  const all = res.rows[0];
  const facts: string[] = [];
  if (all?.n) {
    facts.push(`${all.n} harte Einheiten (≥ ${minLoad} ESS oder Quality): Readiness am Folgetag im Schnitt ${sgn(all.r1, 0)} Punkte ggü. deinem 28-Tage-Schnitt.`);
    if (all.to != null) facts.push(`Bis die Readiness wieder auf Normalniveau ist (≥ Schnitt − 2), vergehen im Schnitt ${fmtNum(all.to, 1)} Tage.`);
    const g = res.rows.slice(1).filter((r) => r.n >= 3).sort((a, b) => (a.r1 ?? 0) - (b.r1 ?? 0));
    if (g.length >= 2) facts.push(`Am stärksten wirkt ${g[0].g} (${sgn(g[0].r1)} am Folgetag), am mildesten ${g[g.length - 1].g} (${sgn(g[g.length - 1].r1)}).`);
    if (all.to != null) facts.push(all.to >= 2.5 ? 'Persönliche Regel: nach einer harten Einheit 2 lockere Tage einplanen.' : all.to >= 1.5 ? 'Persönliche Regel: nach einer harten Einheit mindestens 1 lockerer Tag.' : 'Du erholst dich schnell – ein lockerer Folgetag reicht meist.');
  }
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <p className="text-xs text-ink-muted">Wie reagieren Readiness, HRV und RHR an den 3 Tagen nach einer harten Einheit? Bezug = dein Schnitt der 28 Tage davor.</p>
        <Chips items={[150, 200, 250] as const as any} value={minLoad as any} onChange={(v: any) => setMinLoad(v)} fmt={(v) => `≥ ${v} ESS`} />
      </div>
      <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(0,1.4fr)_minmax(0,1fr)]">
        <Box title="Ø Abweichung je Sport">
          <div className="overflow-x-auto">
            <table className="w-full text-xs tnum">
              <thead>
                <tr className="text-ink-dim text-left">
                  <th className="py-1 pr-2 font-normal">Gruppe</th>
                  <th className="py-1 pr-2 font-normal text-right">n</th>
                  <th className="py-1 pr-2 font-normal text-right">Ø Last</th>
                  <th className="py-1 pr-2 font-normal text-right">Ready T+1</th>
                  <th className="py-1 pr-2 font-normal text-right">T+2</th>
                  <th className="py-1 pr-2 font-normal text-right">T+3</th>
                  <th className="py-1 pr-2 font-normal text-right">HRV T+1</th>
                  <th className="py-1 pr-2 font-normal text-right">RHR T+1</th>
                  <th className="py-1 font-normal text-right">Tage bis normal</th>
                </tr>
              </thead>
              <tbody>
                {res.rows.map((r) => (
                  <tr key={r.g} className={`border-t border-border ${r.g === 'Alle' ? 'font-semibold' : ''}`}>
                    <td className="py-1 pr-2">{r.g}</td>
                    <td className="py-1 pr-2 text-right">{r.n}</td>
                    <td className="py-1 pr-2 text-right">{fmtNum(r.load)}</td>
                    {[r.r1, r.r2, r.r3].map((v, k) => (
                      <td key={k} className={`py-1 pr-2 text-right ${v == null ? '' : v < -8 ? 'text-orange-300' : v < -2 ? 'text-yellow-300' : 'text-green-300'}`}>{sgn(v)}</td>
                    ))}
                    <td className="py-1 pr-2 text-right">{sgn(r.h1, 1)}</td>
                    <td className="py-1 pr-2 text-right">{sgn(r.p1, 1)}</td>
                    <td className="py-1 text-right">{r.to == null ? '—' : r.to >= 3.5 ? '> 3' : fmtNum(r.to, 1)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="mt-1 text-2xs text-ink-dim">„Tage bis normal“ = erster Folgetag mit Readiness ≥ Schnitt − 2 (max. 3 geprüft).</p>
        </Box>
        <Box title="Readiness-Verlauf nach harter Einheit">
          <div style={{ height: 200 }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={curve} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="t" tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis width={34} tickLine={false} tick={{ fontSize: 10 }} />
                <ReferenceLine y={0} stroke="#475569" />
                <Tooltip formatter={ttFmt(' Pkt', 1)} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                {res.rows.filter((r) => r.n >= 3).map((r) => (
                  <Line key={r.g} dataKey={r.g} stroke={r.g === 'Alle' ? '#e2e8f0' : SPORT_COLORS[r.g] || '#94a3b8'} strokeWidth={r.g === 'Alle' ? 2.5 : 1.5} dot={{ r: 3 }} isAnimationActive={false} />
                ))}
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </Box>
      </div>
      <Facts items={facts} />
      <Box title="Letzte harte Einheiten">
        <div className="overflow-x-auto">
          <table className="w-full text-xs tnum">
            <thead>
              <tr className="text-ink-dim text-left">
                <th className="py-1 pr-2 font-normal">Datum</th>
                <th className="py-1 pr-2 font-normal">Einheit</th>
                <th className="py-1 pr-2 font-normal text-right">Readiness T+1 / T+2 / T+3</th>
                <th className="py-1 font-normal text-right">Tage bis normal</th>
              </tr>
            </thead>
            <tbody>
              {res.items.slice(-10).reverse().map((x) => (
                <tr key={x.d.date} className="border-t border-border">
                  <td className="py-1 pr-2">{WD[wdIdx(x.d.date)]} {fmtDateShort(x.d.date).slice(0, 6)}</td>
                  <td className="py-1 pr-2">{fmtNum(x.d.essDay)} ESS {x.d.actual.sport} {x.d.actual.zone}</td>
                  <td className="py-1 pr-2 text-right">{x.r.map((v) => sgn(v)).join(' / ')}</td>
                  <td className="py-1 text-right">{x.daysTo == null ? '—' : x.daysTo > 3 ? '> 3' : x.daysTo}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Box>
    </div>
  );
}

// ---------------------------------------------------------------------------
// 4) Jahresrückblick (12 Monate)
// ---------------------------------------------------------------------------
function YearReport({ days }: { days: D[] }) {
  const months = useMemo(() => {
    const keys = [...new Set(days.map((d) => d.date.slice(0, 7)))].sort().slice(-12);
    return keys.map((k) => {
      const ds = days.filter((d) => d.date.startsWith(k));
      const o: any = { key: k, label: `${MONTHS_DE[Number(k.slice(5, 7)) - 1]} ${k.slice(2, 4)}` };
      for (const s of Object.keys(SPORT_COLORS)) o[s] = 0;
      for (const d of ds) {
        const s = normSport(d.actual.sport);
        if (s && (d.essDay ?? 0) > 0) o[s] += d.essDay!;
      }
      o.ess = sum(ds.map((d) => d.essDay));
      o.sessions = ds.filter((d) => (d.essDay ?? 0) > 0).length;
      o.readiness = avg(ds.map((d) => d.readiness));
      o.sleep = avg(ds.map((d) => (d.sleepH && d.sleepH > 0 ? d.sleepH : null)));
      o.sick = ds.filter(isSick).length;
      o.CTL = [...ds].reverse().find((d) => d.ctlPost != null)?.ctlPost ?? null;
      o.dist = sum(ds.map((d) => d.x.dist));
      o.elev = sum(ds.map((d) => d.x.elev));
      return o;
    });
  }, [days]);
  const yearDays = days.filter((d) => d.date >= `${months[0]?.key}-01`);
  const roll7 = yearDays.map((d, i) => ({ d, v: sum(yearDays.slice(Math.max(0, i - 6), i + 1).map((x) => x.essDay)) }));
  const best = <T,>(xs: T[], f: (x: T) => number | null, min = false) =>
    xs.reduce<T | null>((b, x) => {
      const v = f(x);
      if (v == null) return b;
      if (!b) return x;
      return min ? (v < f(b)! ? x : b) : v > f(b)! ? x : b;
    }, null);
  const pDay = best(yearDays, (d) => d.essDay);
  const pWeek = best(roll7, (x) => x.v);
  const pCtl = best(yearDays, (d) => d.ctlPost);
  const pR = best(yearDays, (d) => d.readiness);
  const pRhr = best(yearDays, (d) => d.rhr, true);
  const pHrv = best(yearDays, (d) => d.hrv);
  const totalEss = sum(months.map((m) => m.ess));
  const totalDist = sum(months.map((m) => m.dist));
  const totalElev = sum(months.map((m) => m.elev));
  const sportTotals = Object.keys(SPORT_COLORS).map((s) => ({ s, v: sum(months.map((m) => m[s])) })).filter((x) => x.v > 0).sort((a, b) => b.v - a.v);
  const ctl0 = months.find((m) => m.CTL != null)?.CTL ?? null;
  const ctl1 = [...months].reverse().find((m) => m.CTL != null)?.CTL ?? null;
  return (
    <div className="space-y-4">
      <div className="grid gap-3 grid-cols-2 md:grid-cols-4 xl:grid-cols-6">
        <Kpi label="Last 12 Monate" value={`${fmtNum(totalEss)} ESS`} sub={`Ø ${fmtNum(totalEss / Math.max(1, yearDays.length / 7))} / Woche`} />
        <Kpi label="Einheiten" value={sum(months.map((m) => m.sessions))} sub={`${sum(months.map((m) => m.sick))} Kranktage`} />
        <Kpi label="Ø Readiness" value={fmtNum(avg(yearDays.map((d) => d.readiness)))} />
        <Kpi label="Ø Schlaf" value={`${fmtHm(avg(yearDays.map((d) => (d.sleepH && d.sleepH > 0 ? d.sleepH : null))))} h`} />
        <Kpi label="CTL" value={`${fmtNum(ctl0)} → ${fmtNum(ctl1)}`} sub={ctl0 && ctl1 ? sgn(((ctl1 - ctl0) / ctl0) * 100, 0, ' %') : undefined} />
        <Kpi label="Distanz · Höhe" value={totalDist > 0 ? `${fmtNum(totalDist)} km` : '—'} sub={totalDist > 0 ? `${fmtNum(totalElev)} Hm · lückenhaft erfasst` : 'nicht erfasst'} />
      </div>
      <Box title="Monatslast nach Sport · CTL am Monatsende">
        <div style={{ height: 260 }}>
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={months} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
              <CartesianGrid stroke={GRID} vertical={false} />
              <XAxis dataKey="label" tickLine={false} tick={{ fontSize: 10 }} />
              <YAxis yAxisId="l" width={44} tickLine={false} tick={{ fontSize: 10 }} />
              <YAxis yAxisId="c" orientation="right" width={40} tickLine={false} tick={{ fontSize: 10 }} />
              <Tooltip formatter={ttFmt()} />
              <Legend wrapperStyle={{ fontSize: 10 }} />
              {Object.keys(SPORT_COLORS).map((s) => (
                <Bar key={s} yAxisId="l" stackId="s" dataKey={s} fill={SPORT_COLORS[s]} fillOpacity={0.75} isAnimationActive={false} />
              ))}
              <Line yAxisId="c" dataKey="CTL" stroke="#22c55e" strokeWidth={2} dot={{ r: 2 }} connectNulls isAnimationActive={false} />
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </Box>
      <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-2">
        <Box title="Bestwerte">
          <dl className="text-xs space-y-1 tnum">
            {[
              ['Höchste Tageslast', pDay ? `${fmtNum(pDay.essDay)} ESS ${pDay.actual.sport}` : '—', pDay?.date],
              ['Höchste 7-Tage-Last', pWeek ? `${fmtNum(pWeek.v)} ESS` : '—', pWeek?.d.date],
              ['Höchste CTL', pCtl ? fmtNum(pCtl.ctlPost) : '—', pCtl?.date],
              ['Beste Readiness', pR ? fmtNum(pR.readiness) : '—', pR?.date],
              ['Niedrigster RHR', pRhr ? `${fmtNum(pRhr.rhr)} bpm` : '—', pRhr?.date],
              ['Höchste HRV', pHrv ? `${fmtNum(pHrv.hrv)} ms` : '—', pHrv?.date],
            ].map(([l, v, d]) => (
              <div key={l as string} className="flex justify-between gap-2">
                <dt className="text-ink-muted">{l}</dt>
                <dd>{v}{d ? ` · ${fmtDateShort(d as string)}` : ''}</dd>
              </div>
            ))}
          </dl>
        </Box>
        <Box title="Sportarten (12 Monate)">
          <div className="space-y-1.5">
            {sportTotals.map((x) => (
              <div key={x.s} className="text-xs">
                <div className="flex justify-between tnum">
                  <span>{x.s}</span>
                  <span>{fmtNum(x.v)} ESS <span className="text-ink-dim">({fmtNum((x.v / Math.max(1, totalEss)) * 100)} %)</span></span>
                </div>
                <div className="h-1.5 rounded bg-bg-subtle mt-0.5 overflow-hidden">
                  <div className="h-full" style={{ width: `${(x.v / Math.max(1, totalEss)) * 100}%`, background: SPORT_COLORS[x.s] }} />
                </div>
              </div>
            ))}
          </div>
        </Box>
      </div>
      <Box title="Monate im Überblick">
        <div className="overflow-x-auto">
          <table className="w-full text-xs tnum">
            <thead>
              <tr className="text-ink-dim text-left">
                <th className="py-1 pr-2 font-normal">Monat</th>
                <th className="py-1 pr-2 font-normal text-right">Last</th>
                <th className="py-1 pr-2 font-normal text-right">Einh.</th>
                <th className="py-1 pr-2 font-normal text-right">Readiness</th>
                <th className="py-1 pr-2 font-normal text-right">Schlaf</th>
                <th className="py-1 pr-2 font-normal text-right">CTL</th>
                <th className="py-1 font-normal text-right">krank</th>
              </tr>
            </thead>
            <tbody>
              {[...months].reverse().map((m) => (
                <tr key={m.key} className="border-t border-border">
                  <td className="py-1 pr-2">{m.label}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(m.ess)}</td>
                  <td className="py-1 pr-2 text-right">{m.sessions}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(m.readiness)}</td>
                  <td className="py-1 pr-2 text-right">{fmtHm(m.sleep)}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(m.CTL)}</td>
                  <td className={`py-1 text-right ${m.sick ? 'text-orange-300' : 'text-ink-dim'}`}>{m.sick || '–'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Box>
    </div>
  );
}

// ---------------------------------------------------------------------------
// 5) Schlafbericht
// ---------------------------------------------------------------------------
function SleepReport({ days }: { days: D[] }) {
  const ds = days.filter((d) => d.sleepH != null && d.sleepH > 0);
  // sleep_hours steht in der Zeile des Morgens → Nacht davor. Wochentag = Morgen.
  const byWd = WD.map((w, i) => {
    const x = ds.filter((d) => wdIdx(d.date) === i);
    return {
      wd: `${WD[(i + 6) % 7]}→${w}`,
      'Ø Schlaf h': avg(x.map((d) => d.sleepH)),
      Schlafscore: avg(x.map((d) => d.sleepScore)),
      goalPct: x.length ? Math.round((x.filter((d) => d.sleepH! >= SLEEP_GOAL_H - 1e-9).length / x.length) * 100) : null,
      n: x.length,
    };
  });
  const byMonth = [...new Set(ds.map((d) => d.date.slice(0, 7)))].sort().map((k) => {
    const x = ds.filter((d) => d.date.startsWith(k));
    return {
      label: `${MONTHS_DE[Number(k.slice(5, 7)) - 1]} ${k.slice(2, 4)}`,
      'Nächte ≥ 7:30 %': Math.round((x.filter((d) => d.sleepH! >= SLEEP_GOAL_H - 1e-9).length / x.length) * 100),
      'Ø Schlaf h': avg(x.map((d) => d.sleepH)),
    };
  });
  const buckets: [string, number, number][] = [
    ['< 6 h', 0, 6],
    ['6–7 h', 6, 7],
    ['7–7:30 h', 7, 7.5],
    ['7:30–8 h', 7.5, 8],
    ['≥ 8 h', 8, 99],
  ];
  const bRows = buckets.map(([l, a, b]) => {
    const x = ds.filter((d) => d.sleepH! >= a && d.sleepH! < b);
    return { l, n: x.length, r: avg(x.map((d) => d.readiness)), h: avg(x.map((d) => d.hrv)), p: avg(x.map((d) => d.rhr)) };
  });
  const wkend = ds.filter((d) => [5, 6].includes(wdIdx(d.date)));
  const wk = ds.filter((d) => ![5, 6].includes(wdIdx(d.date)));
  const worst = [...byWd].filter((x) => x.n).sort((a, b) => (a['Ø Schlaf h'] ?? 0) - (b['Ø Schlaf h'] ?? 0))[0];
  const bestW = [...byWd].filter((x) => x.n).sort((a, b) => (b['Ø Schlaf h'] ?? 0) - (a['Ø Schlaf h'] ?? 0))[0];
  const low = bRows[0].n >= 3 ? bRows[0] : bRows[1];
  const high = bRows[4].n >= 3 ? bRows[4] : bRows[3];
  const facts: string[] = [];
  if (worst && bestW) facts.push(`Kürzeste Nacht im Schnitt ${worst.wd} (${fmtHm(worst['Ø Schlaf h'])} h), längste ${bestW.wd} (${fmtHm(bestW['Ø Schlaf h'])} h).`);
  if (wkend.length && wk.length) facts.push(`Wochenende Ø ${fmtHm(avg(wkend.map((d) => d.sleepH)))} h vs. Werktage Ø ${fmtHm(avg(wk.map((d) => d.sleepH)))} h.`);
  if (low.r != null && high.r != null) facts.push(`Readiness nach ${high.l}: Ø ${fmtNum(high.r)} – nach ${low.l}: Ø ${fmtNum(low.r)} (${sgn(high.r - low.r)} Punkte).`);
  facts.push(`${ds.filter((d) => d.sleepH! >= SLEEP_GOAL_H - 1e-9).length} von ${ds.length} Nächten erreichten das Ziel 7:30 h (${fmtNum((ds.filter((d) => d.sleepH! >= SLEEP_GOAL_H - 1e-9).length / Math.max(1, ds.length)) * 100)} %).`);
  return (
    <div className="space-y-4">
      <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-2">
        <Box title="Nach Wochentag (Nacht → Morgen)">
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={byWd} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="wd" tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis yAxisId="h" width={30} tickLine={false} tick={{ fontSize: 10 }} domain={[5, 9]} allowDataOverflow />
                <YAxis yAxisId="p" orientation="right" width={32} tickLine={false} tick={{ fontSize: 10 }} domain={[0, 100]} unit="%" />
                <ReferenceLine yAxisId="h" y={SLEEP_GOAL_H} stroke="#22c55e" strokeDasharray="4 4" />
                <Tooltip formatter={(v: any, n: string) => [n === 'Ø Schlaf h' ? `${fmtHm(v)} h` : n === 'Ziel erreicht' ? `${v} %` : fmtNum(v), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Bar yAxisId="h" dataKey="Ø Schlaf h" fill="#7dd3fc" fillOpacity={0.7} isAnimationActive={false} />
                <Line yAxisId="p" dataKey="goalPct" name="Ziel erreicht" stroke="#e2e8f0" strokeWidth={1.6} dot={{ r: 3 }} isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </Box>
        <Box title="Nächte ≥ 7:30 h je Monat">
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={byMonth} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="label" tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis width={32} tickLine={false} tick={{ fontSize: 10 }} domain={[0, 100]} unit="%" />
                <Tooltip formatter={(v: any, n: string) => [n.includes('%') ? `${v} %` : `${fmtHm(v)} h`, n]} />
                <Bar dataKey="Nächte ≥ 7:30 %" fill="#22c55e" fillOpacity={0.6} isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </Box>
      </div>
      <Box title="Schlafdauer → Morgenwerte">
        <div className="overflow-x-auto">
          <table className="w-full text-xs tnum">
            <thead>
              <tr className="text-ink-dim text-left">
                <th className="py-1 pr-2 font-normal">Schlaf</th>
                <th className="py-1 pr-2 font-normal text-right">Nächte</th>
                <th className="py-1 pr-2 font-normal text-right">Ø Readiness</th>
                <th className="py-1 pr-2 font-normal text-right">Ø HRV</th>
                <th className="py-1 font-normal text-right">Ø RHR</th>
              </tr>
            </thead>
            <tbody>
              {bRows.map((b) => (
                <tr key={b.l} className="border-t border-border">
                  <td className="py-1 pr-2">{b.l}</td>
                  <td className="py-1 pr-2 text-right">{b.n}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(b.r)}</td>
                  <td className="py-1 pr-2 text-right">{fmtNum(b.h)}</td>
                  <td className="py-1 text-right">{fmtNum(b.p, 1)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Box>
      <Facts items={facts} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// 6) Ernährungsbericht
// ---------------------------------------------------------------------------
function NutritionReport({ days }: { days: D[] }) {
  const withN = days.filter((d) => d.x.kcalIn != null && d.x.kcalIn > 0 && d.x.kcalOut != null && d.x.kcalOut > 0);
  const weeks = useMemo(() => {
    const keys = [...new Set(days.map((d) => isoWeek(d.date).key))].sort();
    return keys.map((k) => {
      const w = days.filter((d) => isoWeek(d.date).key === k);
      const n = w.filter((d) => d.x.kcalIn != null && d.x.kcalIn > 0 && d.x.kcalOut != null && d.x.kcalOut > 0);
      return {
        label: `KW ${isoWeek(w[0].date).kw}`,
        'Ø Bilanz kcal': n.length ? Math.round(avg(n.map((d) => d.x.kcalIn! - d.x.kcalOut!))!) : null,
        Wochenlast: sum(w.map((d) => d.essDay)),
        n: n.length,
      };
    });
  }, [days]);
  // Vortageseffekt: Bilanz/KH am Tag t-1 → Readiness/HRV am Tag t
  const pairs = days.slice(1).map((d, i) => ({ prev: days[i], d })).filter((p) => p.prev.x.kcalIn && p.prev.x.kcalOut && p.d.readiness != null);
  const bal = pairs.map((p) => p.prev.x.kcalIn! - p.prev.x.kcalOut!);
  const rBal = pearson(bal, pairs.map((p) => p.d.readiness!));
  const carbPairs = pairs.filter((p) => p.prev.x.carb != null && p.prev.x.carb > 0);
  const rCarb = pearson(carbPairs.map((p) => p.prev.x.carb!), carbPairs.map((p) => p.d.readiness!));
  const bk: [string, number, number][] = [
    ['Defizit > 500', -1e9, -500],
    ['Defizit 0–500', -500, 0],
    ['Überschuss', 0, 1e9],
  ];
  const bRows = bk.map(([l, a, b]) => {
    const x = pairs.filter((p) => {
      const v = p.prev.x.kcalIn! - p.prev.x.kcalOut!;
      return v >= a && v < b;
    });
    return { l, n: x.length, r: avg(x.map((p) => p.d.readiness)), h: avg(x.map((p) => p.d.hrv)) };
  });
  const prot = avg(withN.map((d) => d.x.protein));
  const carb = avg(withN.map((d) => d.x.carb));
  const fat = avg(withN.map((d) => d.x.fat));
  const lastEntry = withN[withN.length - 1]?.date;
  const facts: string[] = [];
  facts.push(`${withN.length} von ${days.length} Tagen mit Ernährungswerten${lastEntry ? `, zuletzt am ${fmtDateShort(lastEntry)}` : ''}.`);
  if (rBal != null) facts.push(`Energiebilanz Vortag ↔ Readiness: r ${fmtNum(rBal, 2)} (${Math.abs(rBal) < 0.2 ? 'kaum Zusammenhang' : rBal > 0 ? 'Defizite drücken die Readiness' : 'gegenläufig'}).`);
  if (rCarb != null) facts.push(`Kohlenhydrate Vortag ↔ Readiness: r ${fmtNum(rCarb, 2)}.`);
  const big = bRows[0];
  const sur = bRows[2];
  if (big.n >= 5 && sur.n >= 5 && big.r != null && sur.r != null) facts.push(`Nach großem Defizit Readiness Ø ${fmtNum(big.r)}, nach Überschuss Ø ${fmtNum(sur.r)}.`);
  return (
    <div className="space-y-4">
      {withN.length < 14 && <p className="text-xs text-orange-300">Wenige Ernährungsdaten im Zeitraum – Aussagen sind unsicher. Abendwerte regelmäßig eintragen verbessert den Bericht.</p>}
      <div className="grid gap-3 grid-cols-2 md:grid-cols-5">
        <Kpi label="Tage mit Daten" value={`${withN.length} / ${days.length}`} />
        <Kpi label="Ø Bilanz" value={withN.length ? `${sgn(avg(withN.map((d) => d.x.kcalIn! - d.x.kcalOut!)), 0)} kcal` : '—'} />
        <Kpi label="Ø Kohlenhydrate" value={carb != null ? `${fmtNum(carb)} g` : '—'} />
        <Kpi label="Ø Protein" value={prot != null ? `${fmtNum(prot)} g` : '—'} />
        <Kpi label="Ø Fett" value={fat != null ? `${fmtNum(fat)} g` : '—'} />
      </div>
      <Box title="Ø Tagesbilanz je Woche · Wochenlast">
        <div style={{ height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={weeks} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
              <CartesianGrid stroke={GRID} vertical={false} />
              <XAxis dataKey="label" tickLine={false} tick={{ fontSize: 10 }} minTickGap={8} />
              <YAxis yAxisId="b" width={44} tickLine={false} tick={{ fontSize: 10 }} />
              <YAxis yAxisId="l" orientation="right" width={40} tickLine={false} tick={{ fontSize: 10 }} />
              <ReferenceLine yAxisId="b" y={0} stroke="#475569" />
              <Tooltip formatter={(v: any, n: string) => [n.includes('kcal') ? `${fmtNum(v)} kcal` : `${fmtNum(v)} ESS`, n]} />
              <Legend wrapperStyle={{ fontSize: 10 }} />
              <Bar yAxisId="b" dataKey="Ø Bilanz kcal" fill="#eab308" fillOpacity={0.6} isAnimationActive={false} />
              <Line yAxisId="l" dataKey="Wochenlast" stroke="#7dd3fc" strokeWidth={1.6} dot={false} isAnimationActive={false} />
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </Box>
      <Box title="Bilanz am Vortag → Morgenwerte">
        <table className="w-full text-xs tnum">
          <thead>
            <tr className="text-ink-dim text-left">
              <th className="py-1 pr-2 font-normal">Vortag</th>
              <th className="py-1 pr-2 font-normal text-right">Tage</th>
              <th className="py-1 pr-2 font-normal text-right">Ø Readiness</th>
              <th className="py-1 font-normal text-right">Ø HRV</th>
            </tr>
          </thead>
          <tbody>
            {bRows.map((b) => (
              <tr key={b.l} className="border-t border-border">
                <td className="py-1 pr-2">{b.l}</td>
                <td className="py-1 pr-2 text-right">{b.n}</td>
                <td className="py-1 pr-2 text-right">{fmtNum(b.r)}</td>
                <td className="py-1 text-right">{fmtNum(b.h)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Box>
      <Facts items={facts} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// 7) Krankheits- & Ausfallbericht
// ---------------------------------------------------------------------------
function OutageReport({ days }: { days: D[] }) {
  const episodes = useMemo(() => {
    const eps: { start: number; end: number }[] = [];
    for (let i = 0; i < days.length; i++) {
      if (!isSick(days[i])) continue;
      const last = eps[eps.length - 1];
      if (last && last.end >= i - 2) last.end = i;
      else eps.push({ start: i, end: i });
    }
    return eps.map((e) => {
      const pre = days.slice(Math.max(0, e.start - 3), e.start);
      const signs: string[] = [];
      for (const d of pre) {
        const off = e.start - days.indexOf(d);
        if (d.rhrDelta != null && d.rhrDelta >= 3) signs.push(`T−${off}: RHR ${sgn(d.rhrDelta, 0)} bpm`);
        if (d.hrv != null && d.hrvLow != null && d.hrv < d.hrvLow) signs.push(`T−${off}: HRV ${fmtNum(d.hrv)} < ${fmtNum(d.hrvLow)}`);
        if (d.readiness != null && d.readiness < 40) signs.push(`T−${off}: Readiness ${fmtNum(d.readiness)}`);
        if (d.befinden != null && d.befinden <= 2) signs.push(`T−${off}: Befinden ${d.befinden}/5`);
      }
      const load7 = sum(days.slice(Math.max(0, e.start - 7), e.start).map((d) => d.essDay));
      const back = days.slice(e.end + 1).findIndex((d) => (d.essDay ?? 0) > 0);
      return { from: days[e.start].date, to: days[e.end].date, len: e.end - e.start + 1, signs, load7, backAfter: back >= 0 ? back + 1 : null };
    });
  }, [days]);
  const gaps = useMemo(() => {
    const m = new Map<string, number>();
    for (const d of days) if (!hasMorning(d) && !isSick(d)) m.set(d.date.slice(0, 7), (m.get(d.date.slice(0, 7)) || 0) + 1);
    return [...m.entries()].sort();
  }, [days]);
  const sickDays = sum(episodes.map((e) => e.len));
  const warned = episodes.filter((e) => e.signs.length).length;
  const facts: string[] = [];
  facts.push(`${episodes.length} Krankheitsphase${episodes.length === 1 ? '' : 'n'} mit zusammen ${sickDays} Tagen im Zeitraum.`);
  if (episodes.length) facts.push(`${warned} von ${episodes.length} kündigten sich in den 3 Tagen davor an (RHR ≥ +3 bpm, HRV unter Normalband, Readiness < 40 oder Befinden ≤ 2).`);
  if (gaps.length) facts.push(`${sum(gaps.map((g) => g[1]))} Tage ohne Morgenwerte (nicht krank) – Lücken im Datensatz.`);
  return (
    <div className="space-y-4">
      <div className="grid gap-3 grid-cols-2 md:grid-cols-4">
        <Kpi label="Krankheitsphasen" value={episodes.length} />
        <Kpi label="Kranktage" value={sickDays} />
        <Kpi label="Mit Vorwarnung" value={episodes.length ? `${warned} / ${episodes.length}` : '—'} />
        <Kpi label="Tage ohne Morgenwerte" value={sum(gaps.map((g) => g[1]))} />
      </div>
      <Box title="Krankheitsphasen">
        {!episodes.length ? (
          <p className="text-xs text-ink-dim">Keine als „krank“ markierten Tage im Zeitraum.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs tnum">
              <thead>
                <tr className="text-ink-dim text-left">
                  <th className="py-1 pr-2 font-normal">Zeitraum</th>
                  <th className="py-1 pr-2 font-normal text-right">Tage</th>
                  <th className="py-1 pr-2 font-normal text-right">Last 7 T davor</th>
                  <th className="py-1 pr-2 font-normal">Warnzeichen (3 Tage davor)</th>
                  <th className="py-1 font-normal text-right">1. Training nach</th>
                </tr>
              </thead>
              <tbody>
                {[...episodes].reverse().map((e) => (
                  <tr key={e.from} className="border-t border-border align-top">
                    <td className="py-1 pr-2 whitespace-nowrap">{fmtDateShort(e.from)}{e.len > 1 ? ` – ${fmtDateShort(e.to).slice(0, 6)}` : ''}</td>
                    <td className="py-1 pr-2 text-right">{e.len}</td>
                    <td className="py-1 pr-2 text-right">{fmtNum(e.load7)} ESS</td>
                    <td className={`py-1 pr-2 ${e.signs.length ? 'text-orange-300' : 'text-ink-dim'}`}>{e.signs.length ? e.signs.join(' · ') : 'keine erkennbar'}</td>
                    <td className="py-1 text-right">{e.backAfter != null ? `${e.backAfter} T` : '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Box>
      <Box title="Tage ohne Morgenwerte je Monat">
        {!gaps.length ? (
          <p className="text-xs text-ink-dim">Keine Lücken.</p>
        ) : (
          <div className="flex flex-wrap gap-2 text-xs tnum">
            {gaps.map(([k, n]) => (
              <span key={k} className="chip border border-border">
                {MONTHS_DE[Number(k.slice(5, 7)) - 1]} {k.slice(2, 4)}: {n}
              </span>
            ))}
          </div>
        )}
      </Box>
      <Facts items={facts} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Container
// ---------------------------------------------------------------------------
type Win = 90 | 180 | 360;

export function Reports() {
  const [sub, setSub] = useState<Sub>('month');
  const [win, setWin] = useState<Win>(360);
  const [raw, setRaw] = useState<any[] | null>(null);
  const [extra, setExtra] = useState<any[] | null>(null);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    fetchChartData('360d', RECOVERY_FETCH_METRICS)
      .then((r) => setRaw(r.data as any[]))
      .catch((e) => setErr(e instanceof ApiError ? e.message : String(e)));
    fetchChartData('360d', EXTRA_METRICS)
      .then((r) => setExtra(r.data as any[]))
      .catch(() => setExtra([]));
  }, []);

  const all: D[] = useMemo(() => {
    if (!raw) return [];
    const today = localIsoDate();
    const xm = new Map<string, any>();
    for (const r of extra || []) xm.set(String(r.date).slice(0, 10), r);
    return buildDays(raw)
      .filter((d) => d.date <= today)
      .map((d) => {
        const r = xm.get(d.date) || {};
        return {
          ...d,
          x: {
            kcalIn: toNum(r.kcal_in),
            kcalOut: toNum(r.kcal_out),
            carb: toNum(r.carb_g),
            protein: toNum(r.protein_g),
            fat: toNum(r.fat_g),
            dist: toNum(r.distance_km_day),
            elev: toNum(r.elev_m_day),
            ctl: toNum(r.fbCTL_obs),
          },
        };
      });
  }, [raw, extra]);
  const winDays = all.slice(-win);
  const needsWin = sub === 'recovery' || sub === 'sleep' || sub === 'nutrition' || sub === 'outage';

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <Chips items={SUBS.map((s) => s.id)} value={sub} onChange={setSub} fmt={(id) => SUBS.find((s) => s.id === id)!.label} />
        {needsWin && <Chips items={[90, 180, 360] as Win[]} value={win} onChange={setWin} fmt={(v) => `${v} T`} />}
      </div>
      {sub === 'month' ? (
        all.length ? <MonthReport days={all} /> : err ? <ErrorBox message={err} /> : <Skeleton className="h-64" />
      ) : (
        <Panel title={`${TITLES[sub]}${needsWin ? ` · ${win} Tage` : ''}`}>
          {err ? (
            <ErrorBox message={err} />
          ) : !all.length ? (
            <Skeleton className="h-64" />
          ) : sub === 'week' ? (
            <WeekReport days={all} />
          ) : sub === 'recovery' ? (
            <RecoveryAfterHard days={winDays} />
          ) : sub === 'year' ? (
            <YearReport days={all} />
          ) : sub === 'sleep' ? (
            <SleepReport days={winDays} />
          ) : sub === 'nutrition' ? (
            <NutritionReport days={winDays} />
          ) : (
            <OutageReport days={winDays} />
          )}
        </Panel>
      )}
    </div>
  );
}
