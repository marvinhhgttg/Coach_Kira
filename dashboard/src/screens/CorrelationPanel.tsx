import { useMemo, useState } from 'react';
import {
  CartesianGrid,
  ComposedChart,
  Line,
  ResponsiveContainer,
  Scatter,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { fmtDateShort } from '../lib/format';

// ---------------------------------------------------------------------------
// Korrelationen Readiness / HRV / RHR / Schlaf / Last
// Deterministisch: Pearson r über den gewählten Zeitraum, optional mit 1 Tag Versatz.
// ---------------------------------------------------------------------------

export type CorrDay = {
  date: string;
  readiness: number | null;
  hrv: number | null;
  rhr: number | null;
  sleepH: number | null;
  sleepScore: number | null;
  essDay: number | null;
};

type VarKey = 'readiness' | 'hrv' | 'rhr' | 'sleepH' | 'sleepScore' | 'essDay';

const VARS: { key: VarKey; label: string; short: string; unit: string }[] = [
  { key: 'readiness', label: 'Readiness', short: 'Ready', unit: '' },
  { key: 'hrv', label: 'HRV', short: 'HRV', unit: 'ms' },
  { key: 'rhr', label: 'RHR', short: 'RHR', unit: 'bpm' },
  { key: 'sleepH', label: 'Schlafdauer', short: 'Schlaf h', unit: 'h' },
  { key: 'sleepScore', label: 'Schlafscore', short: 'Score', unit: '' },
  { key: 'essDay', label: 'Last', short: 'Last', unit: 'ESS' },
];
const VAR = Object.fromEntries(VARS.map((v) => [v.key, v])) as Record<VarKey, (typeof VARS)[number]>;

/** Werte ≤ 0 bei Readiness/Schlaf/HRV/RHR = fehlende Messung, nicht „0“. Last 0 = Ruhetag (gültig). */
function val(d: CorrDay | undefined, k: VarKey): number | null {
  if (!d) return null;
  const v = d[k];
  if (v == null || !Number.isFinite(v)) return null;
  if (k !== 'essDay' && v <= 0) return null;
  return v;
}

type Pair = { x: number; y: number; date: string };

function pairs(days: CorrDay[], xk: VarKey, yk: VarKey, lag: 0 | 1): Pair[] {
  const out: Pair[] = [];
  for (let i = lag; i < days.length; i++) {
    const x = val(days[i - lag], xk);
    const y = val(days[i], yk);
    if (x != null && y != null) out.push({ x, y, date: days[i].date });
  }
  return out;
}

function pearson(p: Pair[]): number | null {
  const n = p.length;
  if (n < 8) return null;
  let sx = 0, sy = 0;
  for (const q of p) { sx += q.x; sy += q.y; }
  const mx = sx / n, my = sy / n;
  let cov = 0, vx = 0, vy = 0;
  for (const q of p) {
    const dx = q.x - mx, dy = q.y - my;
    cov += dx * dy; vx += dx * dx; vy += dy * dy;
  }
  if (vx === 0 || vy === 0) return null;
  return cov / Math.sqrt(vx * vy);
}

function regression(p: Pair[]): { a: number; b: number } | null {
  const n = p.length;
  if (n < 3) return null;
  const mx = p.reduce((s, q) => s + q.x, 0) / n;
  const my = p.reduce((s, q) => s + q.y, 0) / n;
  let num = 0, den = 0;
  for (const q of p) { num += (q.x - mx) * (q.y - my); den += (q.x - mx) ** 2; }
  if (den === 0) return null;
  const b = num / den;
  return { a: my - b * mx, b };
}

function strength(r: number): string {
  const a = Math.abs(r);
  if (a >= 0.6) return 'stark';
  if (a >= 0.4) return 'deutlich';
  if (a >= 0.2) return 'schwach';
  return 'kaum';
}

/** Zellfarbe: grün = positiv, rot = negativ, Deckkraft ~ |r|. */
function cellStyle(r: number | null): React.CSSProperties {
  if (r == null) return {};
  const a = Math.min(1, Math.abs(r));
  const alpha = 0.08 + a * 0.62;
  return { background: r >= 0 ? `rgba(34,197,94,${alpha})` : `rgba(239,68,68,${alpha})` };
}

const fmtR = (r: number | null) => (r == null ? '—' : `${r >= 0 ? '+' : '−'}${Math.abs(r).toFixed(2).replace('.', ',')}`);
const fmtV = (v: number, k: VarKey) =>
  k === 'sleepH' ? `${Math.floor(v)}:${String(Math.round((v % 1) * 60)).padStart(2, '0')}` : (Math.round(v * 10) / 10).toString().replace('.', ',');

function trendText(b: number, sx: VarKey, sy: VarKey): string {
  const step = sx === 'essDay' ? 100 : 1;
  const d = b * step;
  const dec = sy === 'sleepH' ? 2 : Math.abs(d) < 1 ? 2 : 1;
  const sign = d >= 0 ? '+' : '−';
  return `+${step} ${VAR[sx].unit || 'Punkt'} → ${sign}${Math.abs(d).toFixed(dec).replace('.', ',')} ${VAR[sy].unit || 'Punkte'}`;
}

export function CorrelationPanel({ days, spanLabel }: { days: CorrDay[]; spanLabel: string }) {
  const [lag, setLag] = useState<0 | 1>(0);
  const [sel, setSel] = useState<[VarKey, VarKey]>(['rhr', 'readiness']);

  // Ohne Versatz: 5 Erholungsgrößen (Last am selben Tag ist meist nach der Messung → wenig sinnvoll).
  // Mit Versatz: Zeile = Vortag (inkl. Last), Spalte = heute.
  const rowsVars: VarKey[] = lag === 0 ? ['readiness', 'hrv', 'rhr', 'sleepH', 'sleepScore'] : ['essDay', 'readiness', 'hrv', 'rhr', 'sleepH', 'sleepScore'];
  const colVars: VarKey[] = ['readiness', 'hrv', 'rhr', 'sleepH', 'sleepScore'];

  const matrix = useMemo(() => {
    const m: Record<string, { r: number | null; n: number }> = {};
    for (const xr of rowsVars)
      for (const yc of colVars) {
        const p = pairs(days, xr, yc, lag);
        m[`${xr}|${yc}`] = { r: pearson(p), n: p.length };
      }
    return m;
  }, [days, lag]);

  const [xk, yk] = sel;
  const selValid = rowsVars.includes(xk) && colVars.includes(yk) && !(lag === 0 && xk === yk);
  const sx: VarKey = selValid ? xk : lag === 0 ? 'rhr' : 'essDay';
  const sy: VarKey = selValid ? yk : lag === 0 ? 'readiness' : 'hrv';

  const pts = useMemo(() => pairs(days, sx, sy, lag), [days, sx, sy, lag]);
  const r = pearson(pts);
  const reg = regression(pts);
  const todayIso = days.length ? days[days.length - 1].date : '';
  const todayPt = pts.find((p) => p.date === todayIso) || null;
  const others = pts.filter((p) => p.date !== todayIso);
  const xs = pts.map((p) => p.x);
  const xMin = xs.length ? Math.min(...xs) : 0;
  const xMax = xs.length ? Math.max(...xs) : 1;
  const trend = reg ? [{ x: xMin, y: reg.a + reg.b * xMin }, { x: xMax, y: reg.a + reg.b * xMax }] : [];

  // strongest pairs (excluding trivial diagonal / Score↔Dauer)
  const top = useMemo(() => {
    const list: { xr: VarKey; yc: VarKey; r: number }[] = [];
    for (const xr of rowsVars)
      for (const yc of colVars) {
        if (xr === yc) continue; // Diagonale bzw. Autokorrelation (HRV gestern ↔ HRV heute) ist trivial
        if (lag === 0 && rowsVars.indexOf(xr) > colVars.indexOf(yc)) continue;
        const c = matrix[`${xr}|${yc}`];
        if (c?.r != null) list.push({ xr, yc, r: c.r });
      }
    return list.sort((a, b) => Math.abs(b.r) - Math.abs(a.r)).slice(0, 3);
  }, [matrix, lag]);

  const sentence =
    r == null
      ? 'Zu wenige gemeinsame Tage für eine Aussage.'
      : `${lag ? `${VAR[sx].label} (Vortag)` : VAR[sx].label} und ${VAR[sy].label}${lag ? ' (heute)' : ''}: ${strength(r)} ${
          r >= 0 ? 'positiv' : 'negativ'
        } (r ${fmtR(r)}, ${pts.length} Tage). ${
          Math.abs(r) < 0.2
            ? 'Praktisch kein Zusammenhang.'
            : r >= 0
              ? `Höhere ${VAR[sx].label} geht mit höherer ${VAR[sy].label} einher.`
              : `Höhere ${VAR[sx].label} geht mit niedrigerer ${VAR[sy].label} einher.`
        }${reg && Math.abs(r) >= 0.2 ? ` Trend: ${trendText(reg.b, sx, sy)}.` : ''}`;

  return (
    <div className="panel-raised p-3 min-w-0">
      <div className="flex flex-wrap items-baseline justify-between gap-2">
        <div className="label">Korrelationen · {spanLabel}</div>
        <div className="flex items-center gap-2">
          <span className="text-2xs text-ink-dim">Versatz</span>
          <div className="flex rounded border border-border p-0.5">
            {([0, 1] as const).map((l) => (
              <button
                key={l}
                onClick={() => setLag(l)}
                className={`px-2 py-0.5 text-2xs rounded whitespace-nowrap ${lag === l ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
                data-testid={`corr-lag-${l}`}
              >
                {l === 0 ? 'gleicher Tag' : 'Vortag → heute'}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="mt-3 grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(0,1fr)_minmax(0,1.2fr)]">
        {/* Matrix */}
        <div className="min-w-0 overflow-x-auto">
          <table className="w-full table-fixed text-2xs tnum border-separate" style={{ borderSpacing: 2 }}>
            <thead>
              <tr>
                <th className="text-left text-ink-dim font-normal pr-1 w-16 leading-tight">{lag ? 'Vortag ↓ heute →' : ''}</th>
                {colVars.map((c) => (
                  <th key={c} className="text-ink-muted font-medium px-1 whitespace-nowrap">{VAR[c].short}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rowsVars.map((xr) => (
                <tr key={xr}>
                  <th className="text-left text-ink-muted font-medium pr-1 whitespace-nowrap">{VAR[xr].short}</th>
                  {colVars.map((yc) => {
                    const diag = lag === 0 && xr === yc;
                    const c = matrix[`${xr}|${yc}`];
                    const active = sx === xr && sy === yc;
                    return (
                      <td key={yc} className="p-0">
                        <button
                          disabled={diag || c?.r == null}
                          onClick={() => setSel([xr, yc])}
                          title={diag ? '' : `${VAR[xr].label}${lag ? ' (Vortag)' : ''} ↔ ${VAR[yc].label}: r ${fmtR(c?.r ?? null)} · ${c?.n ?? 0} Tage`}
                          className={`corr-cell w-full h-8 rounded text-[11px] font-semibold ${diag ? 'bg-bg-subtle text-ink-dim' : 'text-ink'} ${active ? 'ring-2 ring-accent' : ''}`}
                          style={diag ? {} : cellStyle(c?.r ?? null)}
                          data-testid={`corr-${xr}-${yc}`}
                        >
                          {diag ? '·' : fmtR(c?.r ?? null)}
                        </button>
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
          <div className="mt-2 flex items-center gap-2 text-2xs text-ink-dim">
            <span className="inline-block w-3 h-3 rounded" style={{ background: 'rgba(239,68,68,0.7)' }} /> negativ
            <span className="inline-block w-3 h-3 rounded" style={{ background: 'rgba(34,197,94,0.7)' }} /> positiv
            <span>· Farbtiefe = Stärke · Klick = Streudiagramm</span>
          </div>
          {top.length > 0 && (
            <ul className="mt-2 space-y-0.5 text-2xs text-ink-muted">
              {top.map((t) => (
                <li key={`${t.xr}|${t.yc}`} className="tnum">
                  · {VAR[t.xr].label}{lag ? ' (Vortag)' : ''} ↔ {VAR[t.yc].label}: <span className={t.r >= 0 ? 'text-green-300' : 'text-red-300'}>{fmtR(t.r)}</span> ({strength(t.r)})
                </li>
              ))}
            </ul>
          )}
        </div>

        {/* Streudiagramm */}
        <div className="min-w-0">
          <div className="flex items-baseline justify-between gap-2">
            <span className="text-xs font-semibold">
              {VAR[sx].label}{lag ? ' (Vortag)' : ''} → {VAR[sy].label}{lag ? ' (heute)' : ''}
            </span>
            <span className={`text-xs tnum font-semibold ${r == null ? 'text-ink-dim' : r >= 0 ? 'text-green-300' : 'text-red-300'}`}>r {fmtR(r)}</span>
          </div>
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart margin={{ top: 8, right: 12, left: 0, bottom: 4 }}>
                <CartesianGrid stroke="#1f2731" />
                <XAxis type="number" dataKey="x" domain={['auto', 'auto']} tickLine={false} tick={{ fontSize: 10 }} name={VAR[sx].label} allowDuplicatedCategory={false} />
                <YAxis type="number" dataKey="y" domain={['auto', 'auto']} tickLine={false} tick={{ fontSize: 10 }} width={34} name={VAR[sy].label} />
                <Tooltip
                  cursor={{ strokeDasharray: '3 3' }}
                  content={({ payload }) => {
                    const p: any = payload && payload[0] && payload[0].payload;
                    if (!p || p.date == null) return null;
                    return (
                      <div className="panel px-2 py-1 text-2xs tnum">
                        <div className="text-ink-muted">{fmtDateShort(p.date)}</div>
                        <div>{VAR[sx].label}{lag ? ' (Vortag)' : ''}: {fmtV(p.x, sx)}</div>
                        <div>{VAR[sy].label}: {fmtV(p.y, sy)}</div>
                      </div>
                    );
                  }}
                />
                <Scatter data={others} fill="#7dd3fc" fillOpacity={0.45} isAnimationActive={false} />
                {trend.length === 2 && (
                  <Line data={trend} dataKey="y" type="linear" stroke={r != null && r < 0 ? '#ef4444' : '#22c55e'} strokeWidth={2} dot={false} isAnimationActive={false} legendType="none" tooltipType="none" />
                )}
                {todayPt && (
                  <Scatter
                    data={[todayPt]}
                    isAnimationActive={false}
                    shape={(props: any) => <circle cx={props.cx} cy={props.cy} r={6} fill="#f472b6" stroke="#fff" strokeWidth={1.5} />}
                  />
                )}
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <p className="mt-1 text-xs text-ink-muted leading-relaxed">{sentence}</p>
          {pts.length > 0 && pts.length < 30 && (
            <p className="mt-1 text-2xs text-orange-300">Nur {pts.length} Tage – bei so wenigen Punkten schwankt r stark. Für belastbare Aussagen 60 T oder mehr wählen.</p>
          )}
          <p className="mt-1 text-2xs text-ink-dim">
            {todayPt ? '● pink = heute · ' : ''}Pearson r über {spanLabel}; ab ±0,2 schwach, ±0,4 deutlich, ±0,6 stark. Zusammenhang ≠ Ursache – Garmin rechnet Schlaf und HRV teils selbst in die Readiness ein.
          </p>
        </div>
      </div>
    </div>
  );
}
