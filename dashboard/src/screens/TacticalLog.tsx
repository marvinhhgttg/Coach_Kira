import { Fragment, useEffect, useMemo, useState, type ReactNode } from 'react';
import { ApiError, fetchChartData } from '../lib/api';
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
  VOTE_LABEL,
  VOTE_LEVEL,
  VOTE_ORDER,
  type RecoveryDay,
} from './RecoveryPanel';

// ---------------------------------------------------------------------------
// Log: Aktivitäten mit Kontext, Kalender, Filter, Serien, Bestenliste, Ertrag.
// Regelbasiert aus der Timeline.
// ---------------------------------------------------------------------------

type Win = 28 | 60 | 90 | 180 | 360;
const WINS: Win[] = [28, 60, 90, 180, 360];
const WD = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
const MONTHS = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
const SPORT_COLORS: Record<string, string> = { Bike: '#38bdf8', Run: '#f97316', 'Hike/Walk': '#22c55e', Ski: '#a855f7', Sonstige: '#94a3b8' };

const wdIdx = (iso: string) => (new Date(iso + 'T12:00:00').getDay() + 6) % 7;
const isSick = (d: RecoveryDay) => (d.actual.sport || '').trim().toLowerCase() === 'krank';
const normSport = (s: string) => {
  const t = (s || '').trim().toLowerCase();
  if (!t || t === 'off' || t === 'krank') return '';
  if (t.startsWith('run') || t === 'laufen') return 'Run';
  if (t.includes('bike') || t.includes('rad')) return 'Bike';
  if (t.includes('hike') || t === 'walk') return 'Hike/Walk';
  if (t.startsWith('ski')) return 'Ski';
  return 'Sonstige';
};
const zoneGroup = (z: string) => {
  const m = (z || '').toUpperCase().match(/[0-5]/g);
  if (!m) return '';
  const mx = Math.max(...m.map(Number));
  return mx >= 4 ? 'Z4+' : mx === 3 ? 'Z3' : 'Z1–Z2';
};
const loadGroup = (e: number) => (e >= 220 ? 'hoch' : e >= 130 ? 'Training' : 'locker');
const sgn = (v: number | null | undefined, d = 0, unit = '') => (v == null ? '—' : `${v > 0 ? '+' : v < 0 ? '−' : '±'}${fmtNum(Math.abs(v), d)}${unit}`);
const avg = (xs: (number | null | undefined)[]) => {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};

type Act = {
  d: RecoveryDay;
  idx: number; // Index in all
  sport: string;
  zoneG: string;
  ess: number;
  teAe: number;
  teAn: number;
  ertrag: number | null; // TE gesamt je 100 ESS
  demand: (typeof VOTE_ORDER)[number];
  fits: boolean | null;
  hard: boolean;
  next: RecoveryDay | null;
  base: { r: number | null; h: number | null; p: number | null };
};

function buildActs(all: RecoveryDay[]): Act[] {
  const today = localIsoDate();
  const out: Act[] = [];
  all.forEach((d, i) => {
    const ess = d.essDay ?? 0;
    if (ess <= 0 || !(d.activityDone || d.date < today)) return;
    const demand = demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
    const prior = all.slice(Math.max(0, i - 28), i);
    const teAe = d.actual.teAe ?? 0;
    const teAn = d.actual.teAn ?? 0;
    out.push({
      d,
      idx: i,
      sport: normSport(d.actual.sport) || 'Sonstige',
      zoneG: zoneGroup(d.actual.zone),
      ess,
      teAe,
      teAn,
      ertrag: ess >= 30 && teAe + teAn > 0 ? ((teAe + teAn) / ess) * 100 : null,
      demand,
      fits: d.vote ? VOTE_ORDER.indexOf(demand) <= VOTE_ORDER.indexOf(d.vote) : null,
      hard: demand === 'QUALITY' || ess >= 200,
      next: all[i + 1] && all[i + 1].date <= today ? all[i + 1] : null,
      base: { r: avg(prior.map((x) => x.readiness)), h: avg(prior.map((x) => x.hrv)), p: avg(prior.map((x) => x.rhr)) },
    });
  });
  return out;
}

// ---------------------------------------------------------------------------
// UI-Bausteine
// ---------------------------------------------------------------------------
function Chips<T extends string | number>({ items, value, onChange, fmt }: { items: T[]; value: T; onChange: (v: T) => void; fmt?: (v: T) => string }) {
  return (
    <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5">
      {items.map((v) => (
        <button key={String(v)} onClick={() => onChange(v)} className={`px-2 py-1 text-xs rounded tnum whitespace-nowrap ${value === v ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}>
          {fmt ? fmt(v) : String(v)}
        </button>
      ))}
    </div>
  );
}
function Kpi({ label, value, sub, tone }: { label: string; value: ReactNode; sub?: ReactNode; tone?: string }) {
  return (
    <div className="panel-raised p-2.5 min-w-0">
      <div className="text-2xs text-ink-dim">{label}</div>
      <div className={`tnum text-sm font-semibold whitespace-nowrap ${tone || ''}`}>{value}</div>
      {sub != null && <div className="text-2xs text-ink-dim tnum truncate">{sub}</div>}
    </div>
  );
}
function VoteChip({ v }: { v: RecoveryDay['vote'] }) {
  if (!v) return <span className="text-ink-dim">—</span>;
  const lvl = VOTE_LEVEL[v];
  return <span className={`chip border text-2xs ${LEVEL_CELL[lvl]} ${LEVEL_TEXT[lvl]} border-border`}>{VOTE_LABEL[v]}</span>;
}
const tone = (v: number | null, goodHigh = true, t = 3) => (v == null ? '' : (goodHigh ? v : -v) >= -t ? ((goodHigh ? v : -v) > t ? 'text-green-300' : '') : 'text-orange-300');

// ---------------------------------------------------------------------------
// 1) Detail / Nachbetrachtung
// ---------------------------------------------------------------------------
function Detail({ a }: { a: Act }) {
  const d = a.d;
  const n = a.next;
  const dr = n?.readiness != null && a.base.r != null ? n.readiness - a.base.r : null;
  const dh = n?.hrv != null && a.base.h != null ? n.hrv - a.base.h : null;
  const dp = n?.rhr != null && a.base.p != null ? n.rhr - a.base.p : null;
  return (
    <div className="grid gap-3 grid-cols-[minmax(0,1fr)] md:grid-cols-4 text-xs">
      <div>
        <div className="label mb-1">Morgen davor</div>
        <dl className="space-y-0.5 tnum">
          <Row k="Readiness" v={fmtNum(d.readiness)} />
          <Row k="HRV" v={d.hrv != null ? `${fmtNum(d.hrv)} ms${d.hrvLow != null ? ` (Band ${fmtNum(d.hrvLow)}–${fmtNum(d.hrvHigh)})` : ''}` : '—'} />
          <Row k="RHR" v={d.rhr != null ? `${fmtNum(d.rhr)} bpm (${sgn(d.rhrDelta, 1)})` : '—'} />
          <Row k="Schlaf" v={d.sleepH != null ? `${fmtHm(d.sleepH)} h${d.sleepScore != null ? ` · Score ${fmtNum(d.sleepScore)}` : ''}` : '—'} />
          <Row k="Befinden" v={d.befinden != null ? `${d.befinden} / 5` : '—'} />
        </dl>
      </div>
      <div>
        <div className="label mb-1">Votum & Einheit</div>
        <div className="flex items-center gap-2">
          <VoteChip v={d.vote} />
          <span className="text-ink-muted">Anspruch {VOTE_LABEL[a.demand]}</span>
          {a.fits != null && <span className={a.fits ? 'text-green-300' : 'text-orange-300'}>{a.fits ? '✓ passend' : '✗ härter'}</span>}
        </div>
        <ul className="mt-1 text-2xs text-ink-dim space-y-0.5">
          {d.reasons.slice(0, 3).map((r) => (
            <li key={r}>· {r}</li>
          ))}
        </ul>
      </div>
      <div>
        <div className="label mb-1">Folgemorgen vs. 28-T-Schnitt</div>
        {!n ? (
          <span className="text-ink-dim">noch nicht verfügbar</span>
        ) : (
          <dl className="space-y-0.5 tnum">
            <Row k="Readiness" v={<span className={tone(dr, true, 2)}>{fmtNum(n.readiness)} ({sgn(dr)})</span>} />
            <Row k="HRV" v={<span className={tone(dh, true, 1)}>{fmtNum(n.hrv)} ms ({sgn(dh, 1)})</span>} />
            <Row k="RHR" v={<span className={tone(dp, false, 1)}>{fmtNum(n.rhr)} bpm ({sgn(dp, 1)})</span>} />
            <Row k="Schlaf" v={n.sleepH != null ? `${fmtHm(n.sleepH)} h` : '—'} />
          </dl>
        )}
      </div>
      <div>
        <div className="label mb-1">Belastung danach</div>
        <dl className="space-y-0.5 tnum">
          <Row k="ACWR" v={<span className={d.acwrPost != null && d.acwrPost > 1.3 ? 'text-orange-300' : ''}>{fmtNum(d.acwrPost, 2)}</span>} />
          <Row k="Monotonie" v={<span className={d.monotonyPost != null && d.monotonyPost > 2 ? 'text-orange-300' : d.monotonyPost != null && d.monotonyPost > 1.6 ? 'text-yellow-300' : ''}>{fmtNum(d.monotonyPost, 2)}</span>} />
          <Row k="CTL / ATL" v={`${fmtNum(d.ctlPost)} / ${fmtNum(d.atlPost)}`} />
          <Row k="Ertrag" v={a.ertrag != null ? `${fmtNum(a.ertrag, 2)} TE/100 ESS` : '—'} />
        </dl>
      </div>
    </div>
  );
}
function Row({ k, v }: { k: string; v: ReactNode }) {
  return (
    <div className="flex justify-between gap-2">
      <dt className="text-ink-muted">{k}</dt>
      <dd className="text-right">{v}</dd>
    </div>
  );
}

// ---------------------------------------------------------------------------
// 2) Kalender
// ---------------------------------------------------------------------------
function Calendar({ days, acts, onPick, picked }: { days: RecoveryDay[]; acts: Map<string, Act>; onPick: (date: string) => void; picked: string | null }) {
  const months = [...new Set(days.map((d) => d.date.slice(0, 7)))].sort().reverse();
  const byDate = new Map(days.map((d) => [d.date, d]));
  const maxE = Math.max(1, ...[...acts.values()].map((a) => a.ess));
  return (
    <div className="grid gap-4 grid-cols-[minmax(0,1fr)] md:grid-cols-2 xl:grid-cols-3">
      {months.map((k) => {
        const [y, m] = k.split('-').map(Number);
        const first = new Date(y, m - 1, 1, 12);
        const lead = (first.getDay() + 6) % 7;
        const nDays = new Date(y, m, 0).getDate();
        const cells: (string | null)[] = [...Array(lead).fill(null), ...Array.from({ length: nDays }, (_, i) => `${k}-${String(i + 1).padStart(2, '0')}`)];
        const tot = [...acts.values()].filter((a) => a.d.date.startsWith(k)).reduce((s, a) => s + a.ess, 0);
        return (
          <div key={k} className="panel-raised p-3 min-w-0">
            <div className="flex justify-between items-baseline mb-2">
              <span className="text-sm font-semibold">{MONTHS[m - 1]} {y}</span>
              <span className="text-2xs text-ink-dim tnum">{fmtNum(tot)} ESS</span>
            </div>
            <div className="grid grid-cols-7 gap-1 text-[10px] tnum">
              {WD.map((w) => (
                <div key={w} className="text-center text-ink-dim">{w}</div>
              ))}
              {cells.map((iso, i) => {
                if (!iso) return <div key={`e${i}`} />;
                const d = byDate.get(iso);
                const a = acts.get(iso);
                const sick = d && isSick(d);
                const col = a ? SPORT_COLORS[a.sport] || '#94a3b8' : null;
                const alpha = a ? 0.25 + 0.65 * Math.min(1, a.ess / maxE) : 0;
                return (
                  <button
                    key={iso}
                    disabled={!d}
                    onClick={() => onPick(iso)}
                    title={a ? `${fmtDateShort(iso)} · ${fmtNum(a.ess)} ESS ${a.d.actual.sport} ${a.d.actual.zone}${d?.vote ? ` · Votum ${VOTE_LABEL[d.vote]}` : ''}` : d ? `${fmtDateShort(iso)} · ${sick ? 'krank' : 'Ruhe'}` : ''}
                    className={`cal-cell relative h-9 rounded border ${picked === iso ? 'border-accent ring-1 ring-accent' : 'border-border'} ${!d ? 'opacity-30' : ''}`}
                    style={col ? { background: `${col}${Math.round(alpha * 255).toString(16).padStart(2, '0')}` } : undefined}
                  >
                    <span className="absolute top-0.5 left-1 text-ink-dim">{Number(iso.slice(8))}</span>
                    <span className="absolute bottom-0.5 right-1 font-semibold text-ink">{a ? fmtNum(a.ess) : sick ? 'K' : ''}</span>
                    {a && a.fits === false && <span className="absolute top-0.5 right-1 w-1.5 h-1.5 rounded-full bg-ampel-orange" />}
                  </button>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Hauptkomponente
// ---------------------------------------------------------------------------
type SortKey = 'date' | 'ess' | 'teAe' | 'teAn' | 'ertrag';
const SORTS: { k: SortKey; l: string }[] = [
  { k: 'date', l: 'Datum' },
  { k: 'ess', l: 'Last' },
  { k: 'teAe', l: 'TE aerob' },
  { k: 'teAn', l: 'TE anaerob' },
  { k: 'ertrag', l: 'Ertrag' },
];

export function TacticalLog() {
  const [raw, setRaw] = useState<any[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [win, setWin] = useState<Win>(90);
  const [view, setView] = useState<'Tabelle' | 'Kalender'>('Tabelle');
  const [fSport, setFSport] = useState('Alle');
  const [fZone, setFZone] = useState('Alle');
  const [fLoad, setFLoad] = useState('Alle');
  const [onlyMiss, setOnlyMiss] = useState(false);
  const [sort, setSort] = useState<SortKey>('date');
  const [open, setOpen] = useState<string | null>(null);
  const [picked, setPicked] = useState<string | null>(null);

  const load = () => {
    setErr(null);
    setRaw(null);
    fetchChartData('360d', RECOVERY_FETCH_METRICS)
      .then((r) => setRaw(r.data as any[]))
      .catch((e) => setErr(e instanceof ApiError ? e.message : String(e)));
  };
  useEffect(load, []);

  const all = useMemo(() => (raw ? buildDays(raw).filter((d) => d.date <= localIsoDate()) : []), [raw]);
  const allActs = useMemo(() => buildActs(all), [all]);
  const days = all.slice(-win);
  const from = days[0]?.date ?? '';
  const acts = allActs.filter((a) => a.d.date >= from);
  const actMap = useMemo(() => new Map(acts.map((a) => [a.d.date, a])), [acts]);

  // 4) Serien & Abstände (über alle Daten, Stand heute)
  const streaks = useMemo(() => {
    const today = localIsoDate();
    const past = all.filter((d) => d.date <= today);
    const isTrain = (d: RecoveryDay) => (d.essDay ?? 0) > 0 && (d.activityDone || d.date < today);
    const actOf = new Map(allActs.map((a) => [a.d.date, a]));
    // aktuelle Serie: heute zählt, wenn trainiert; sonst ab gestern
    let i = past.length - 1;
    if (i >= 0 && !isTrain(past[i]) && past[i].date === today && !past[i].activityDone) i--;
    let cur = 0;
    for (let k = i; k >= 0 && isTrain(past[k]); k--) cur++;
    let hardRun = 0;
    for (let k = i; k >= 0 && actOf.get(past[k].date)?.hard; k--) hardRun++;
    const lastRest = [...past].reverse().find((d) => !isTrain(d) && !(d.date === today && !d.activityDone));
    const lastQ = [...allActs].reverse().find((a) => a.demand === 'QUALITY');
    const diff = (iso?: string) => (iso ? Math.round((new Date(today + 'T12:00:00').getTime() - new Date(iso + 'T12:00:00').getTime()) / 86400000) : null);
    let best = 0, run = 0, bestEnd = '';
    for (const d of days) {
      run = isTrain(d) ? run + 1 : 0;
      if (run > best) {
        best = run;
        bestEnd = d.date;
      }
    }
    let bestHard = 0, rh = 0;
    for (const d of days) {
      rh = actOf.get(d.date)?.hard ? rh + 1 : 0;
      bestHard = Math.max(bestHard, rh);
    }
    return { cur, hardRun, sinceRest: diff(lastRest?.date), sinceQ: diff(lastQ?.d.date), lastQ, best, bestEnd, bestHard };
  }, [all, allActs, days]);

  const sports = ['Alle', ...Object.keys(SPORT_COLORS).filter((s) => acts.some((a) => a.sport === s))];
  const filtered = acts
    .filter((a) => fSport === 'Alle' || a.sport === fSport)
    .filter((a) => fZone === 'Alle' || a.zoneG === fZone)
    .filter((a) => fLoad === 'Alle' || loadGroup(a.ess) === fLoad)
    .filter((a) => !onlyMiss || a.fits === false)
    .sort((x, y) => {
      if (sort === 'date') return y.d.date.localeCompare(x.d.date);
      const vx = sort === 'ertrag' ? x.ertrag ?? -1 : (x as any)[sort];
      const vy = sort === 'ertrag' ? y.ertrag ?? -1 : (y as any)[sort];
      return vy - vx;
    });

  // 5) Bestenliste je Sport + 6) Ertrag
  const board = useMemo(() => {
    const groups = Object.keys(SPORT_COLORS).filter((s) => acts.some((a) => a.sport === s));
    const top = (xs: Act[], f: (a: Act) => number | null) => xs.reduce<Act | null>((b, a) => (f(a) != null && (!b || f(a)! > f(b)!) ? a : b), null);
    return groups.map((s) => {
      const xs = acts.filter((a) => a.sport === s);
      return { s, n: xs.length, sum: xs.reduce((t, a) => t + a.ess, 0), maxL: top(xs, (a) => a.ess), maxAe: top(xs, (a) => a.teAe), maxAn: top(xs, (a) => a.teAn), bestE: top(xs, (a) => a.ertrag), avgE: avg(xs.map((a) => a.ertrag)) };
    });
  }, [acts]);
  const ertragSorted = acts.filter((a) => a.ertrag != null && a.ess >= 60).sort((a, b) => b.ertrag! - a.ertrag!);
  const jump = (date: string) => {
    setView('Tabelle');
    setFSport('Alle');
    setFZone('Alle');
    setFLoad('Alle');
    setOnlyMiss(false);
    setOpen(date);
    setTimeout(() => document.getElementById(`log-${date}`)?.scrollIntoView({ block: 'center', behavior: 'smooth' }), 50);
  };

  const totals = { n: acts.length, sum: acts.reduce((t, a) => t + a.ess, 0), miss: acts.filter((a) => a.fits === false).length, checked: acts.filter((a) => a.fits != null).length };
  const pickedDay = picked ? all.find((d) => d.date === picked) : null;
  const pickedAct = picked ? actMap.get(picked) : null;

  return (
    <div className="space-y-4">
      <Panel
        title="Log · Aktivitäten"
        right={
          <div className="flex flex-wrap items-center gap-2">
            <Chips items={['Tabelle', 'Kalender'] as const as any} value={view as any} onChange={(v: any) => setView(v)} />
            <Chips items={WINS} value={win} onChange={setWin} fmt={(v) => `${v} T`} />
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={load} disabled={!raw && !err}>
              {!raw && !err ? 'Lade…' : 'Aktualisieren'}
            </button>
          </div>
        }
      >
        {err ? (
          <ErrorBox message={err} onRetry={load} />
        ) : !raw ? (
          <div className="space-y-2">
            {Array.from({ length: 6 }).map((_, i) => (
              <Skeleton key={i} className="h-10" />
            ))}
          </div>
        ) : (
          <div className="space-y-4">
            {/* 4) Serien & Kennzahlen */}
            <div className="grid gap-3 grid-cols-2 md:grid-cols-4 xl:grid-cols-8">
              <Kpi label="Trainingstage am Stück" value={streaks.cur} tone={streaks.cur >= 6 ? 'text-orange-300' : ''} sub="aktuell" />
              <Kpi label="Seit letztem Ruhetag" value={streaks.sinceRest != null ? `${streaks.sinceRest} T` : '—'} tone={(streaks.sinceRest ?? 0) >= 7 ? 'text-orange-300' : ''} />
              <Kpi label="Seit letzter Quality" value={streaks.sinceQ != null ? `${streaks.sinceQ} T` : '—'} sub={streaks.lastQ ? fmtDateShort(streaks.lastQ.d.date).slice(0, 6) : undefined} />
              <Kpi label="Harte Tage am Stück" value={streaks.hardRun} tone={streaks.hardRun >= 2 ? 'text-orange-300' : ''} sub={`max. ${streaks.bestHard} im Zeitraum`} />
              <Kpi label="Längste Serie" value={`${streaks.best} T`} sub={streaks.bestEnd ? `bis ${fmtDateShort(streaks.bestEnd).slice(0, 6)}` : undefined} />
              <Kpi label="Aktivitäten" value={totals.n} sub={`${win} Tage`} />
              <Kpi label="Gesamtlast" value={`${fmtNum(totals.sum)} ESS`} sub={`Ø ${fmtNum(totals.n ? totals.sum / totals.n : null)} / Einheit`} />
              <Kpi label="Härter als Votum" value={`${totals.miss} / ${totals.checked}`} tone={totals.checked && totals.miss / totals.checked > 0.4 ? 'text-orange-300' : ''} />
            </div>

            {view === 'Kalender' ? (
              <>
                <div className="flex flex-wrap gap-3 text-2xs text-ink-muted">
                  {Object.entries(SPORT_COLORS).map(([s, c]) => (
                    <span key={s} className="inline-flex items-center gap-1">
                      <span className="w-2.5 h-2.5 rounded" style={{ background: c }} /> {s}
                    </span>
                  ))}
                  <span>· Farbtiefe = Last · K = krank · <span className="inline-block w-1.5 h-1.5 rounded-full bg-ampel-orange align-middle" /> härter als Votum</span>
                </div>
                {pickedDay && (
                  <div className="panel-raised p-3">
                    <div className="flex justify-between items-baseline mb-2">
                      <span className="text-sm font-semibold tnum">
                        {WD[wdIdx(pickedDay.date)]} {fmtDateShort(pickedDay.date)} ·{' '}
                        {pickedAct ? `${fmtNum(pickedAct.ess)} ESS ${pickedDay.actual.sport} ${pickedDay.actual.zone} · TE ${fmtNum(pickedAct.teAe, 1)}/${fmtNum(pickedAct.teAn, 1)}` : isSick(pickedDay) ? 'krank' : 'Ruhetag'}
                      </span>
                      <button className="btn btn-ghost text-2xs px-2 py-0.5" onClick={() => setPicked(null)}>× schließen</button>
                    </div>
                    {pickedAct ? <Detail a={pickedAct} /> : <p className="text-xs text-ink-dim">Kein Training an diesem Tag.</p>}
                  </div>
                )}
                <Calendar days={days} acts={actMap} onPick={setPicked} picked={picked} />
              </>
            ) : (
              <>
                {/* 3) Filter & Sortierung */}
                <div className="flex flex-wrap items-center gap-2">
                  <Chips items={sports} value={fSport} onChange={setFSport} />
                  <Chips items={['Alle', 'Z1–Z2', 'Z3', 'Z4+']} value={fZone} onChange={setFZone} />
                  <Chips items={['Alle', 'locker', 'Training', 'hoch']} value={fLoad} onChange={setFLoad} fmt={(v) => (v === 'locker' ? 'locker < 130' : v === 'Training' ? '130–219' : v === 'hoch' ? 'hoch ≥ 220' : 'Alle')} />
                  <label className="inline-flex items-center gap-1.5 text-xs text-ink-muted cursor-pointer select-none">
                    <input type="checkbox" checked={onlyMiss} onChange={(e) => setOnlyMiss(e.target.checked)} className="accent-accent" /> nur härter als Votum
                  </label>
                  <span className="ml-auto inline-flex items-center gap-1 text-xs text-ink-dim">
                    Sortierung <Chips items={SORTS.map((s) => s.k)} value={sort} onChange={setSort} fmt={(k) => SORTS.find((s) => s.k === k)!.l} />
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full min-w-[760px] text-xs">
                    <thead>
                      <tr className="text-left border-b border-border">
                        {['Datum', 'Sport', 'Zone', 'Last', 'TE ae/an', 'Ertrag', 'Votum', 'Folgemorgen', ''].map((h) => (
                          <th key={h} className="label px-2 py-2 font-normal">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.map((a) => {
                        const isOpen = open === a.d.date;
                        const dr = a.next?.readiness != null && a.base.r != null ? a.next.readiness - a.base.r : null;
                        return (
                          <Fragment key={a.d.date}>
                            <tr id={`log-${a.d.date}`} className={`border-t border-border cursor-pointer hover:bg-bg-subtle/60 ${isOpen ? 'bg-bg-subtle/60' : ''}`} onClick={() => setOpen(isOpen ? null : a.d.date)}>
                              <td className="px-2 py-1.5 tnum whitespace-nowrap">{WD[wdIdx(a.d.date)]} {fmtDateShort(a.d.date)}</td>
                              <td className="px-2 py-1.5">
                                <span className="inline-block w-2 h-2 rounded-full mr-1.5 align-middle" style={{ background: SPORT_COLORS[a.sport] }} />
                                {a.d.actual.sport}
                              </td>
                              <td className="px-2 py-1.5">{a.d.actual.zone || '—'}</td>
                              <td className={`px-2 py-1.5 tnum font-semibold ${a.ess >= 220 ? 'text-orange-300' : 'text-accent'}`}>{fmtNum(a.ess)}</td>
                              <td className="px-2 py-1.5 tnum">{fmtNum(a.teAe, 1)} / {fmtNum(a.teAn, 1)}</td>
                              <td className="px-2 py-1.5 tnum">{a.ertrag != null ? fmtNum(a.ertrag, 2) : '—'}</td>
                              <td className="px-2 py-1.5 whitespace-nowrap">
                                <VoteChip v={a.d.vote} /> {a.fits === false && <span className="text-orange-300 ml-1">✗</span>}
                                {a.fits === true && <span className="text-green-300 ml-1">✓</span>}
                              </td>
                              <td className={`px-2 py-1.5 tnum ${tone(dr, true, 2)}`}>{a.next ? `Ready ${sgn(dr)}` : '—'}</td>
                              <td className="px-2 py-1.5 text-ink-dim">{isOpen ? '▴' : '▾'}</td>
                            </tr>
                            {isOpen && (
                              <tr className="bg-bg-subtle/40">
                                <td colSpan={9} className="px-3 py-3">
                                  <Detail a={a} />
                                </td>
                              </tr>
                            )}
                          </Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                  {!filtered.length && <p className="text-sm text-ink-muted py-3">Keine Einheiten für diese Filter.</p>}
                </div>
                <p className="text-2xs text-ink-dim">
                  Ertrag = (TE aerob + TE anaerob) je 100 ESS. Folgemorgen = Readiness am nächsten Morgen ggü. deinem Schnitt der 28 Tage davor. Klick auf eine Zeile = Nachbetrachtung.
                </p>
              </>
            )}
          </div>
        )}
      </Panel>

      {raw && !err && (
        <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(0,1.6fr)_minmax(0,1fr)]">
          {/* 5) Bestenliste */}
          <Panel title={`Bestenliste je Sport · ${win} Tage`}>
            <div className="overflow-x-auto">
              <table className="w-full min-w-[620px] text-xs tnum">
                <thead>
                  <tr className="text-left text-ink-dim">
                    {['Sport', 'Einh.', 'Σ Last', 'Höchste Last', 'Max TE aerob', 'Max TE anaerob', 'Ø Ertrag'].map((h) => (
                      <th key={h} className="py-1 pr-2 font-normal">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {board.map((b) => (
                    <tr key={b.s} className="border-t border-border">
                      <td className="py-1.5 pr-2">
                        <span className="inline-block w-2 h-2 rounded-full mr-1.5 align-middle" style={{ background: SPORT_COLORS[b.s] }} />
                        {b.s}
                      </td>
                      <td className="py-1.5 pr-2">{b.n}</td>
                      <td className="py-1.5 pr-2">{fmtNum(b.sum)}</td>
                      {[
                        [b.maxL, (a: Act) => `${fmtNum(a.ess)} ESS`],
                        [b.maxAe, (a: Act) => fmtNum(a.teAe, 1)],
                        [b.maxAn, (a: Act) => fmtNum(a.teAn, 1)],
                      ].map(([a, f]: any, k) => (
                        <td key={k} className="py-1.5 pr-2">
                          {a ? (
                            <button className="hover:text-accent underline decoration-dotted underline-offset-2" onClick={() => jump(a.d.date)}>
                              {f(a)} · {fmtDateShort(a.d.date).slice(0, 6)}
                            </button>
                          ) : (
                            '—'
                          )}
                        </td>
                      ))}
                      <td className="py-1.5 pr-2">{fmtNum(b.avgE, 2)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="mt-1 text-2xs text-ink-dim">Klick auf einen Wert springt zur Einheit in der Tabelle.</p>
          </Panel>

          {/* 6) Ertrag */}
          <Panel title="Ertrag · TE je 100 ESS">
            {ertragSorted.length < 4 ? (
              <p className="text-xs text-ink-dim">Zu wenige Einheiten mit Trainingseffekt.</p>
            ) : (
              <div className="grid grid-cols-2 gap-3 text-xs tnum">
                {[
                  ['Viel Reiz, wenig Last', ertragSorted.slice(0, 5), 'text-green-300'],
                  ['Viel Last, wenig Reiz', ertragSorted.slice(-5).reverse(), 'text-orange-300'],
                ].map(([t, xs, c]: any) => (
                  <div key={t} className="min-w-0">
                    <div className="label mb-1">{t}</div>
                    <ul className="space-y-1">
                      {xs.map((a: Act) => (
                        <li key={a.d.date}>
                          <button className="w-full text-left hover:text-accent" onClick={() => jump(a.d.date)}>
                            <span className={`font-semibold ${c}`}>{fmtNum(a.ertrag, 2)}</span>{' '}
                            <span className="text-ink-muted">
                              {fmtDateShort(a.d.date).slice(0, 6)} · {fmtNum(a.ess)} ESS {a.d.actual.sport} {a.d.actual.zone}
                            </span>
                          </button>
                        </li>
                      ))}
                    </ul>
                  </div>
                ))}
              </div>
            )}
            <p className="mt-2 text-2xs text-ink-dim">Nur Einheiten ab 60 ESS. Ø Ertrag je Sport steht in der Bestenliste.</p>
          </Panel>
        </div>
      )}
    </div>
  );
}
