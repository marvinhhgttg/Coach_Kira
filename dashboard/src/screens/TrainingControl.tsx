import { useEffect, useMemo, useState, type ReactNode } from 'react';
import {
  Bar,
  CartesianGrid,
  Cell,
  ComposedChart,
  Legend,
  Line,
  ReferenceArea,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { ApiError, dedupeByDateKeepLast, fetchChartData, fetchPlanSimulationWithTimeout, toNum, type Range } from '../lib/api';
import { fmtDateShort, fmtNum } from '../lib/format';
import { ErrorBox, Panel, Skeleton } from '../components/UI';

// ---------------------------------------------------------------------------
// Trainingssteuerung (Chart Deck): Monotonie/Strain, Intensitätsverteilung,
// Fitness/Ermüdung/Form, Plan vs. Ist. Alles deterministisch aus der Timeline.
// ---------------------------------------------------------------------------

const TC_METRICS = [
  'coachE_ESS_day',
  'Monotony7',
  'Strain7',
  'Zone',
  'Sport_x',
  'Aerobic_TE',
  'Anaerobic_TE',
  'fbATL_obs',
  'fbCTL_obs',
  'activity_done',
] as const;

const GRID = '#1f2731';
const C = {
  low: '#22c55e',
  mid: '#eab308',
  high: '#ef4444',
  unk: '#64748b',
  ctl: '#22c55e',
  atl: '#f97316',
  mono: '#a855f7',
  strain: '#7dd3fc',
};

type Day = {
  date: string;
  ess: number;
  mono: number | null;
  strain: number | null;
  zone: string;
  sport: string;
  teAe: number | null;
  teAn: number | null;
  atl: number | null;
  ctl: number | null;
  done: boolean;
};

type Win = 90 | 180 | 360;
const WINS: Win[] = [90, 180, 360];

function localIso(d = new Date()) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}
const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const wdOf = (iso: string) => WD[new Date(iso + 'T12:00:00').getDay()];

/** ISO-Kalenderwoche + Montag der Woche. */
function isoWeek(iso: string): { key: string; kw: number; monday: string } {
  const d = new Date(iso + 'T12:00:00');
  const day = (d.getDay() + 6) % 7; // Mo=0
  const mon = new Date(d);
  mon.setDate(d.getDate() - day);
  const th = new Date(mon);
  th.setDate(mon.getDate() + 3);
  const y = th.getFullYear();
  const jan4 = new Date(y, 0, 4, 12);
  const kw = 1 + Math.round(((th.getTime() - jan4.getTime()) / 86400000 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
  return { key: `${y}-W${String(kw).padStart(2, '0')}`, kw, monday: localIso(mon) };
}

/** Zone-Text → Liste Zonenzahlen ("Z2, Z3" → [2,3], "3" → [3], "Off" → []). */
function parseZones(z: string): number[] {
  const s = (z || '').toUpperCase();
  if (!s || s.includes('OFF')) return [];
  const m = s.match(/Z?\s*([0-5])/g) || [];
  return m.map((x) => Number(x.replace(/[^0-5]/g, ''))).filter((n) => Number.isFinite(n));
}

function mean(xs: number[]) {
  return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;
}
function pstdev(xs: number[]) {
  const m = mean(xs);
  return Math.sqrt(mean(xs.map((x) => (x - m) ** 2)));
}
function monotony(xs: number[]): number | null {
  if (xs.length < 7) return null;
  const sd = pstdev(xs);
  return sd > 0 ? mean(xs) / sd : null;
}
const f1 = (v: number | null | undefined) => (v == null || !Number.isFinite(v) ? '—' : v.toFixed(2).replace('.', ','));

// ---------------------------------------------------------------------------
// Karten-Hülle mit Vollbild
// ---------------------------------------------------------------------------
function Card({
  title,
  hint,
  right,
  children,
  height = 260,
}: {
  title: string;
  hint?: string;
  right?: ReactNode;
  children: (h: number | string) => ReactNode;
  height?: number;
}) {
  const [full, setFull] = useState(false);
  useEffect(() => {
    if (!full) return;
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && setFull(false);
    window.addEventListener('keydown', onKey);
    const prev = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      window.removeEventListener('keydown', onKey);
      document.body.style.overflow = prev;
    };
  }, [full]);
  const head = (
    <header className="px-4 py-2.5 border-b border-border flex flex-wrap items-start justify-between gap-2">
      <div className="min-w-0">
        <h3 className="text-sm font-semibold tracking-tight">{title}</h3>
        {hint && <p className="text-2xs text-ink-dim mt-0.5">{hint}</p>}
      </div>
      <div className="flex items-center gap-2">
        {right}
        <button type="button" className="btn btn-ghost text-2xs px-2 py-1" onClick={() => setFull(!full)}>
          {full ? '× Schließen' : '⛶ Vollbild'}
        </button>
      </div>
    </header>
  );
  return (
    <>
      <div className="panel-raised min-w-0">
        {head}
        <div className="px-2 pt-3 pb-2">{children(height)}</div>
      </div>
      {full && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4" onClick={() => setFull(false)}>
          <div className="panel-raised w-full h-full max-w-[1600px] max-h-[92vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
            {head}
            <div className="flex-1 px-2 py-3 min-h-0 overflow-auto">{children('70vh')}</div>
          </div>
        </div>
      )}
    </>
  );
}

function Stat({ label, value, sub, tone }: { label: string; value: ReactNode; sub?: ReactNode; tone?: string }) {
  return (
    <div className="min-w-0">
      <div className="text-2xs text-ink-dim">{label}</div>
      <div className={`tnum text-sm font-semibold ${tone || ''}`}>{value}</div>
      {sub && <div className="text-2xs text-ink-dim tnum">{sub}</div>}
    </div>
  );
}

// ---------------------------------------------------------------------------
// 1) Monotonie & Strain + Hebel
// ---------------------------------------------------------------------------
type PlanDay = { date: string; load: number; sport: string; zone: string };

function monotonyLever(days: Day[], plan: PlanDay[] | null) {
  const today = localIso();
  const past = days.filter((d) => d.date <= today);
  const fut = (plan || []).filter((p) => p.date > today).slice(0, 7);
  if (past.length < 6 || fut.length === 0) return null;
  const base = [...past.slice(-6).map((d) => d.ess), ...fut.map((p) => p.load)];
  const offset = 6; // Index des ersten Plantags in base
  const peak = (arr: number[]) => {
    let mx = -1;
    let at = -1;
    for (let k = offset; k < arr.length; k++) {
      const m = monotony(arr.slice(k - 6, k + 1));
      if (m != null && m > mx) {
        mx = m;
        at = k;
      }
    }
    return { mx, at };
  };
  const b = peak(base);
  let best: { idx: number; mx: number; mode: 'rest' | 'half' } | null = null;
  for (let j = offset; j < base.length; j++) {
    if (base[j] <= 0) continue;
    for (const mode of ['rest', 'half'] as const) {
      const alt = base.slice();
      alt[j] = mode === 'rest' ? 0 : Math.round(base[j] / 2);
      const r = peak(alt);
      if (!best || r.mx < best.mx - 0.005 || (Math.abs(r.mx - best.mx) < 0.005 && mode === 'half')) best = { idx: j, mx: r.mx, mode };
    }
  }
  return {
    peak: b.mx,
    peakDate: fut[b.at - offset]?.date,
    best: best ? { ...best, date: fut[best.idx - offset].date, load: base[best.idx], sport: fut[best.idx - offset].sport } : null,
  };
}

function MonotonyCard({ days, plan }: { days: Day[]; plan: PlanDay[] | null | 'loading' }) {
  const data = days.map((d) => ({ date: d.date, Monotonie: d.mono, Strain: d.strain }));
  const last = days[days.length - 1];
  const avgMono = mean(days.map((d) => d.mono).filter((x): x is number => x != null));
  const above = days.filter((d) => (d.mono ?? 0) > 2.0).length;
  const lever = useMemo(() => (plan === 'loading' ? null : monotonyLever(days, plan)), [days, plan]);
  const tone = (m: number | null | undefined) => (m == null ? '' : m > 2 ? 'text-orange-300' : m > 1.6 ? 'text-yellow-300' : 'text-green-300');
  return (
    <Card title="Monotonie & Strain" hint="Monotonie = Ø/σ der letzten 7 Tageslasten · Strain = Wochenlast × Monotonie (Sheet)">
      {(h) => (
        <>
          <div className="px-2 grid grid-cols-2 sm:grid-cols-4 gap-3 mb-2">
            <Stat label="Monotonie heute" value={f1(last?.mono)} tone={tone(last?.mono)} />
            <Stat label="Strain heute" value={fmtNum(last?.strain ?? null)} />
            <Stat label="Ø Monotonie" value={f1(avgMono)} tone={tone(avgMono)} />
            <Stat label="Tage > 2,0" value={`${above} / ${days.length}`} />
          </div>
          <div style={{ height: h }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={data} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="date" tickFormatter={(v) => fmtDateShort(String(v)).slice(0, 6)} minTickGap={24} tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis yAxisId="m" domain={[0, (max: number) => Math.max(3, Math.ceil(max * 2) / 2)]} width={30} tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis yAxisId="s" orientation="right" width={40} tickLine={false} tick={{ fontSize: 10 }} />
                <Tooltip labelFormatter={(v) => fmtDateShort(String(v))} formatter={(v: any, n: string) => [n === 'Monotonie' ? f1(v) : fmtNum(v), n]} />
                <Bar yAxisId="s" dataKey="Strain" fill={C.strain} fillOpacity={0.25} isAnimationActive={false} />
                <ReferenceLine yAxisId="m" y={1.6} stroke="#eab308" strokeDasharray="4 4" />
                <ReferenceLine yAxisId="m" y={2.0} stroke="#f97316" strokeDasharray="4 4" />
                <Line yAxisId="m" dataKey="Monotonie" stroke={C.mono} strokeWidth={1.8} dot={false} connectNulls isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div className="mx-2 mt-2 rounded border border-border px-3 py-2 text-xs">
            <div className="label mb-1">Hebel · nächste 7 Plantage</div>
            {plan === 'loading' ? (
              <span className="text-ink-dim">Plan wird geladen…</span>
            ) : !lever ? (
              <span className="text-ink-dim">Kein Plan für die nächsten Tage gefunden.</span>
            ) : lever.peak <= 1.6 ? (
              <span className="text-green-300">Laut Plan bleibt die Monotonie im grünen Bereich (Spitze {f1(lever.peak)}).</span>
            ) : lever.best && lever.best.mx < lever.peak - 0.05 ? (
              <span>
                Spitze laut Plan <b className={tone(lever.peak)}>{f1(lever.peak)}</b>
                {lever.peakDate ? ` (${wdOf(lever.peakDate)} ${fmtDateShort(lever.peakDate).slice(0, 6)})` : ''}. Wirksamster Eingriff:{' '}
                <b>
                  {lever.best.mode === 'rest' ? 'Ruhetag' : `Halbieren auf ${Math.round(lever.best.load / 2)} ESS`} am {wdOf(lever.best.date)}{' '}
                  {fmtDateShort(lever.best.date).slice(0, 6)}
                </b>{' '}
                statt {fmtNum(lever.best.load)} ESS {lever.best.sport} → Spitze <b className={tone(lever.best.mx)}>{f1(lever.best.mx)}</b>.
              </span>
            ) : (
              <span>Spitze laut Plan {f1(lever.peak)} – ein einzelner Tag ändert daran wenig; mehr Wechsel zwischen leichten und schweren Tagen nötig.</span>
            )}
          </div>
        </>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// 2) Intensitätsverteilung (Kalenderwochen)
// ---------------------------------------------------------------------------
function IntensityCard({ days }: { days: Day[] }) {
  const weeks = useMemo(() => {
    const map = new Map<string, { key: string; kw: number; monday: string; low: number; mid: number; high: number; unk: number; teAe: number; teAn: number; n: number }>();
    for (const d of days) {
      const w = isoWeek(d.date);
      if (!map.has(w.key)) map.set(w.key, { ...w, low: 0, mid: 0, high: 0, unk: 0, teAe: 0, teAn: 0, n: 0 });
      const o = map.get(w.key)!;
      o.n++;
      o.teAe += d.teAe ?? 0;
      o.teAn += d.teAn ?? 0;
      if (d.ess <= 0) continue;
      const zs = parseZones(d.zone);
      if (!zs.length) {
        o.unk += d.ess;
        continue;
      }
      const share = d.ess / zs.length;
      for (const z of zs) {
        if (z <= 2) o.low += share;
        else if (z === 3) o.mid += share;
        else o.high += share;
      }
    }
    return [...map.values()].map((w) => {
      const tot = w.low + w.mid + w.high;
      return {
        ...w,
        label: `KW ${w.kw}`,
        'Z1–Z2': Math.round(w.low),
        Z3: Math.round(w.mid),
        'Z4–Z5': Math.round(w.high),
        'ohne Zone': Math.round(w.unk),
        lowPct: tot > 0 ? Math.round((w.low / tot) * 100) : null,
        teRatio: w.teAe > 0 ? w.teAn / w.teAe : null,
      };
    });
  }, [days]);
  const last4 = weeks.slice(-4);
  const sum = (k: 'low' | 'mid' | 'high') => last4.reduce((a, w) => a + w[k], 0);
  const tot4 = sum('low') + sum('mid') + sum('high');
  const pct = (k: 'low' | 'mid' | 'high') => (tot4 > 0 ? Math.round((sum(k) / tot4) * 100) : 0);
  const ae4 = last4.reduce((a, w) => a + w.teAe, 0);
  const an4 = last4.reduce((a, w) => a + w.teAn, 0);
  const lowP = pct('low');
  const midP = pct('mid');
  const highP = pct('high');
  const verdict =
    tot4 === 0
      ? 'Keine Zonendaten.'
      : midP >= 20
        ? `Viel „mittleres Grau“: ${midP} % in Z3. Für 80/20 Z3-Anteil zugunsten Z2 bzw. gezielter Z4-Reize senken.`
        : lowP >= 80
          ? highP < 5
            ? `Polarisiert locker (${lowP} % Z1–Z2), aber kaum hohe Intensität (${highP} %). Bei guter Erholung 1 Qualitätsreiz/Woche möglich.`
            : `Gute 80/20-Verteilung (${lowP} % locker, ${highP} % hart).`
          : `Lockerer Anteil ${lowP} % – unter dem 80-%-Ziel.`;
  return (
    <Card title="Intensitätsverteilung" hint="Last je Kalenderwoche nach Zone · Linie = Anteil Z1–Z2 (Ziel ≥ 80 %)">
      {(h) => (
        <>
          <div className="px-2 grid grid-cols-2 sm:grid-cols-4 gap-3 mb-2">
            <Stat label="Z1–Z2 · 4 Wo." value={`${lowP} %`} tone={lowP >= 80 ? 'text-green-300' : 'text-yellow-300'} />
            <Stat label="Z3 · 4 Wo." value={`${midP} %`} tone={midP >= 20 ? 'text-orange-300' : ''} />
            <Stat label="Z4–Z5 · 4 Wo." value={`${highP} %`} />
            <Stat label="TE anaerob : aerob" value={ae4 > 0 ? `1 : ${fmtNum(ae4 / Math.max(an4, 0.01), 0)}` : '—'} sub={`Σ ${fmtNum(an4, 1)} / ${fmtNum(ae4, 1)}`} />
          </div>
          <div style={{ height: h }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={weeks} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="label" minTickGap={10} tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis yAxisId="l" width={40} tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis yAxisId="p" orientation="right" domain={[0, 100]} width={32} tickLine={false} tick={{ fontSize: 10 }} unit="%" />
                <Tooltip
                  formatter={(v: any, n: string) => [n === 'Anteil Z1–Z2' ? `${v} %` : `${fmtNum(v)} ESS`, n]}
                  labelFormatter={(l, p: any) => (p && p[0] ? `${l} · ab ${fmtDateShort(p[0].payload.monday).slice(0, 6)}` : l)}
                />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Bar yAxisId="l" stackId="z" dataKey="Z1–Z2" fill={C.low} fillOpacity={0.75} isAnimationActive={false} />
                <Bar yAxisId="l" stackId="z" dataKey="Z3" fill={C.mid} fillOpacity={0.8} isAnimationActive={false} />
                <Bar yAxisId="l" stackId="z" dataKey="Z4–Z5" fill={C.high} fillOpacity={0.85} isAnimationActive={false} />
                <Bar yAxisId="l" stackId="z" dataKey="ohne Zone" fill={C.unk} fillOpacity={0.5} isAnimationActive={false} />
                <ReferenceLine yAxisId="p" y={80} stroke="#22c55e" strokeDasharray="4 4" />
                <Line yAxisId="p" dataKey="lowPct" name="Anteil Z1–Z2" stroke="#e2e8f0" strokeWidth={1.6} dot={{ r: 2 }} connectNulls isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <p className="px-2 mt-2 text-xs text-ink-muted">{verdict}</p>
        </>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// 3) Fitness / Ermüdung / Form + Rampe
// ---------------------------------------------------------------------------
function phaseOf(formPct: number | null): { label: string; color: string; tone: string } {
  if (formPct == null) return { label: '—', color: C.unk, tone: '' };
  if (formPct > 5) return { label: 'frisch', color: '#38bdf8', tone: 'text-accent' };
  if (formPct >= -10) return { label: 'neutral', color: '#94a3b8', tone: 'text-ink' };
  if (formPct >= -30) return { label: 'produktiv', color: '#22c55e', tone: 'text-green-300' };
  return { label: 'Überlastung', color: '#ef4444', tone: 'text-red-300' };
}

function FitnessCard({ days }: { days: Day[] }) {
  const data = days.map((d) => {
    const formPct = d.ctl && d.atl != null ? ((d.ctl - d.atl) / d.ctl) * 100 : null;
    return { date: d.date, CTL: d.ctl, ATL: d.atl, Form: formPct != null ? Math.round(formPct * 10) / 10 : null };
  });
  const last = data[data.length - 1];
  const ph = phaseOf(last?.Form ?? null);
  const ctl7 = days.length > 7 ? days[days.length - 8].ctl : null;
  const ramp = last?.CTL != null && ctl7 ? ((last.CTL - ctl7) / ctl7) * 100 : null;
  // Wochenrampe (Kalenderwochen, CTL Wochenende ggü. Vorwoche)
  const ramps = useMemo(() => {
    const byW = new Map<string, { label: string; ctl: number | null }>();
    for (const d of days) {
      const w = isoWeek(d.date);
      byW.set(w.key, { label: `KW ${w.kw}`, ctl: d.ctl ?? byW.get(w.key)?.ctl ?? null });
    }
    const arr = [...byW.values()];
    return arr.slice(1).map((w, i) => ({
      label: w.label,
      Rampe: w.ctl != null && arr[i].ctl ? Math.round(((w.ctl - arr[i].ctl!) / arr[i].ctl!) * 1000) / 10 : null,
    }));
  }, [days]);
  const rampTone = (r: number | null) => (r == null ? '' : r > 8 ? 'text-orange-300' : r > 5 ? 'text-yellow-300' : r < -5 ? 'text-accent' : 'text-green-300');
  return (
    <Card title="Fitness · Ermüdung · Form" hint="CTL (Fitness) · ATL (Ermüdung) · Form = (CTL − ATL) / CTL · Bänder = Phasen">
      {(h) => (
        <>
          <div className="px-2 grid grid-cols-2 sm:grid-cols-4 gap-3 mb-2">
            <Stat label="CTL" value={fmtNum(last?.CTL ?? null)} />
            <Stat label="ATL" value={fmtNum(last?.ATL ?? null)} />
            <Stat label="Form" value={last?.Form != null ? `${last.Form > 0 ? '+' : ''}${fmtNum(last.Form, 1)} %` : '—'} sub={ph.label} tone={ph.tone} />
            <Stat label="Rampe 7 T" value={ramp != null ? `${ramp > 0 ? '+' : ''}${fmtNum(ramp, 1)} %` : '—'} sub="Ziel +3–5 %/Woche" tone={rampTone(ramp)} />
          </div>
          <div style={{ height: h }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={data} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={GRID} vertical={false} />
                <XAxis dataKey="date" tickFormatter={(v) => fmtDateShort(String(v)).slice(0, 6)} minTickGap={24} tickLine={false} tick={{ fontSize: 10 }} />
                <YAxis yAxisId="v" width={40} tickLine={false} tick={{ fontSize: 10 }} domain={['auto', 'auto']} />
                <YAxis yAxisId="f" orientation="right" width={38} tickLine={false} tick={{ fontSize: 10 }} unit="%" domain={[-50, 30]} ticks={[-50, -30, -10, 0, 5, 30]} allowDataOverflow />
                <ReferenceArea yAxisId="f" y1={5} y2={30} fill="#38bdf8" fillOpacity={0.06} />
                <ReferenceArea yAxisId="f" y1={-30} y2={-10} fill="#22c55e" fillOpacity={0.07} />
                <ReferenceArea yAxisId="f" y1={-50} y2={-30} fill="#ef4444" fillOpacity={0.08} />
                <ReferenceLine yAxisId="f" y={0} stroke="#475569" />
                <Tooltip labelFormatter={(v) => fmtDateShort(String(v))} formatter={(v: any, n: string) => [n === 'Form' ? `${fmtNum(v, 1)} % (${phaseOf(v).label})` : fmtNum(v), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Bar yAxisId="f" dataKey="Form" name="Form %" fill="#94a3b8" isAnimationActive={false}>
                  {data.map((d) => (
                    <Cell key={d.date} fill={phaseOf(d.Form).color} fillOpacity={0.35} />
                  ))}
                </Bar>
                <Line yAxisId="v" dataKey="CTL" stroke={C.ctl} strokeWidth={2} dot={false} connectNulls isAnimationActive={false} />
                <Line yAxisId="v" dataKey="ATL" stroke={C.atl} strokeWidth={1.4} dot={false} connectNulls isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div className="px-2 mt-3">
            <div className="text-2xs text-ink-dim mb-1">CTL-Rampe je Kalenderwoche (Ziel +3–5 %, &gt; 8 % = zu steil)</div>
            <div style={{ height: 90 }}>
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={ramps} margin={{ top: 4, right: 8, left: 0, bottom: 0 }}>
                  <XAxis dataKey="label" hide />
                  <YAxis width={40} tickLine={false} tick={{ fontSize: 9 }} unit="%" />
                  <ReferenceLine y={0} stroke="#475569" />
                  <ReferenceLine y={5} stroke="#eab308" strokeDasharray="3 3" />
                  <Tooltip formatter={(v: any) => [`${fmtNum(v, 1)} %`, 'CTL-Rampe']} />
                  <Bar dataKey="Rampe" fill="#94a3b8" isAnimationActive={false}>
                    {ramps.map((r) => (
                      <Cell key={r.label} fill={r.Rampe == null ? C.unk : r.Rampe > 8 ? '#f97316' : r.Rampe > 5 ? '#eab308' : r.Rampe < 0 ? '#38bdf8' : '#22c55e'} fillOpacity={0.7} />
                    ))}
                  </Bar>
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>
          <p className="px-2 mt-1 text-2xs text-ink-dim">Form-Phasen: &gt; +5 % frisch · −10…+5 % neutral · −30…−10 % produktiv · &lt; −30 % Überlastung.</p>
        </>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Container
// ---------------------------------------------------------------------------
const RANGE_OF: Record<Win, Range> = { 90: '90d', 180: '180d', 360: '360d' };

export function TrainingControl() {
  const [win, setWin] = useState<Win>(180);
  const [raw, setRaw] = useState<any[] | null>(null);
  const [plan, setPlan] = useState<PlanDay[] | null | 'loading'>('loading');
  const [busy, setBusy] = useState(true);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    let alive = true;
    setBusy(true);
    setErr(null);
    fetchChartData('360d', TC_METRICS)
      .then((r) => alive && setRaw(r.data as any[]))
      .catch((e) => alive && setErr(e instanceof ApiError ? e.message : String(e)))
      .finally(() => alive && setBusy(false));
    fetchPlanSimulationWithTimeout(45000)
      .then((sim: any) => {
        if (!alive) return;
        setPlan(
          sim?.ok
            ? (sim.days || []).map((x: any) => ({ date: String(x.date).slice(0, 10), load: toNum(x.load) ?? 0, sport: String(x.sport || ''), zone: String(x.zone || '') }))
            : null,
        );
      })
      .catch(() => alive && setPlan(null));
    return () => {
      alive = false;
    };
  }, []);

  const allDays: Day[] = useMemo(() => {
    if (!raw) return [];
    const today = localIso();
    return dedupeByDateKeepLast(raw.filter((r) => r && typeof r.date === 'string').map((r) => ({ ...r, date: String(r.date).slice(0, 10) })))
      .sort((a: any, b: any) => (a.date < b.date ? -1 : 1))
      .filter((r: any) => r.date <= today)
      .map((r: any) => {
        return {
          date: r.date,
          ess: toNum(r.coachE_ESS_day) ?? 0,
          mono: toNum(r.Monotony7),
          strain: toNum(r.Strain7),
          zone: String(r.Zone || ''),
          sport: String(r.Sport_x || ''),
          teAe: toNum(r.Aerobic_TE),
          teAn: toNum(r.Anaerobic_TE),
          atl: toNum(r.fbATL_obs),
          ctl: toNum(r.fbCTL_obs),
          done: String(r.activity_done || '').trim().toLowerCase() === 'x',
        };
      });
  }, [raw]);
  const days = allDays.slice(-win);

  return (
    <Panel
      title="Trainingssteuerung"
      right={
        <div className="inline-flex bg-bg rounded border border-border p-0.5">
          {WINS.map((w) => (
            <button
              key={w}
              onClick={() => setWin(w)}
              className={`px-2 py-1 text-xs rounded tnum ${win === w ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
              data-testid={`tc-win-${w}`}
            >
              {w} T
            </button>
          ))}
        </div>
      }
    >
      {err ? (
        <ErrorBox message={err} />
      ) : busy && !raw ? (
        <div className="grid gap-4 lg:grid-cols-2">
          <Skeleton className="h-[320px]" />
          <Skeleton className="h-[320px]" />
        </div>
      ) : (
        <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-2">
          <MonotonyCard days={days} plan={plan} />
          <IntensityCard days={days} />
          <div className="lg:col-span-2 min-w-0">
            <FitnessCard days={days} />
          </div>
        </div>
      )}
    </Panel>
  );
}
