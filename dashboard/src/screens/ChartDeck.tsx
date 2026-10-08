import { TrainingControl } from './TrainingControl';
import { useEffect, useMemo, useState, useCallback } from 'react';
import {
  CartesianGrid,
  Bar,
  ComposedChart,
  Line,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import {
  fetchChartData,
  ApiError,
  dedupeByDateKeepLast,
  toNum,
  LOAD_METRICS,
  RECOVERY_METRICS,
  type ChartDataResponse,
  type Range,
} from '../lib/api';
import { fmtDateShort, fmtTime } from '../lib/format';
import { ErrorBox, Panel, Skeleton } from '../components/UI';

const RANGES: Range[] = ['7d', '14d', '28d', '60d', '90d', '180d', '360d'];

const METRIC_COLORS: Record<string, string> = {
  load_fb_day: '#7dd3fc',
  fbATL_obs: '#f97316',
  fbCTL_obs: '#22c55e',
  fbACWR_obs: '#a855f7',
  coachE_Smart_Gains: '#e879f9',
  garminEnduranceScore: '#0ea5e9',
  sleep_hours: '#7dd3fc',
  sleep_score_0_100: '#22c55e',
  rhr_bpm: '#ef4444',
  rhr_trend: '#fca5a5',
  Garmin_Training_Readiness: '#a855f7',
};

const METRIC_LABELS: Record<string, string> = {
  load_fb_day: 'Load Day',
  fbATL_obs: 'ATL',
  fbCTL_obs: 'CTL',
  fbACWR_obs: 'ACWR',
  coachE_Smart_Gains: 'KEI',
  garminEnduranceScore: 'Endurance Score',
  sleep_hours: 'Schlaf h',
  sleep_score_0_100: 'Schlafscore',
  rhr_bpm: 'RHR bpm',
  rhr_trend: 'RHR Trend (7d)',
  Garmin_Training_Readiness: 'Readiness',
};

function normaliseData(rows: any[], metrics: readonly string[]) {
  const cleaned = rows
    .filter((r) => r && typeof r.date === 'string' && r.date)
    .map((r) => {
      const out: Record<string, any> = { date: r.date };
      metrics.forEach((m) => {
        out[m] = toNum(r[m]);
      });
      return out;
    });
  return dedupeByDateKeepLast(cleaned);
}

/** Add a 7-day centered-trailing simple moving average of `sourceKey` under `trendKey`. */
function withMovingAverage<T extends Record<string, any>>(
  rows: T[],
  sourceKey: string,
  trendKey: string,
  window = 7,
): T[] {
  return rows.map((row, idx) => {
    const start = Math.max(0, idx - window + 1);
    const slice = rows.slice(start, idx + 1);
    const values = slice
      .map((r) => (typeof r[sourceKey] === 'number' ? (r[sourceKey] as number) : null))
      .filter((v): v is number => v != null && Number.isFinite(v));
    const trend = values.length >= 3 ? values.reduce((a, b) => a + b, 0) / values.length : null;
    return { ...row, [trendKey]: trend };
  });
}

type ChartType = 'line' | 'bar' | 'sleepBars' | 'dualLine' | 'loadBars';

type ChartSectionProps = {
  title: string;
  rows: any[];
  metrics: readonly string[];
  unitHint?: string;
  yDomain?: [number | string, number | string];
  rightDomain?: [number | string, number | string];
  refLine?: { y: number; label: string };
  chartType?: ChartType;
  dualAxis?: boolean;
  fullscreen?: boolean;
  onToggleFullscreen?: () => void;
  trendKeys?: string[];
};

function ChartBody({
  rows,
  metrics,
  yDomain,
  rightDomain,
  refLine,
  chartType = 'line',
  dualAxis = false,
  trendKeys = [],
}: Pick<
  ChartSectionProps,
  'rows' | 'metrics' | 'yDomain' | 'rightDomain' | 'refLine' | 'chartType' | 'dualAxis' | 'trendKeys'
>) {
  const empty = !rows.length || metrics.every((m) => rows.every((r) => r[m] == null));
  if (empty) {
    return (
      <div className="h-full flex items-center justify-center text-sm text-ink-muted">
        Keine Daten im Zeitraum.
      </div>
    );
  }
  return (
    <ResponsiveContainer width="100%" height="100%">
      <ComposedChart data={rows} margin={{ top: 8, right: dualAxis ? 8 : 16, left: 0, bottom: 0 }}>
        <CartesianGrid stroke="#1f2731" vertical={false} />
        <XAxis
          dataKey="date"
          tickFormatter={fmtDateShort}
          tickLine={false}
          axisLine={{ stroke: '#1f2731' }}
          minTickGap={20}
        />
        <YAxis
          yAxisId="left"
          tickLine={false}
          axisLine={{ stroke: '#1f2731' }}
          domain={yDomain || ['auto', 'auto']}
          width={42}
        />
        {dualAxis && (
          <YAxis
            yAxisId="right"
            orientation="right"
            tickLine={false}
            axisLine={{ stroke: '#1f2731' }}
            domain={rightDomain || [0, 100]}
            width={38}
            allowDataOverflow
          />
        )}
        {refLine && (
          <ReferenceLine
            yAxisId="left"
            y={refLine.y}
            stroke="#5b6678"
            strokeDasharray="3 3"
            label={{
              value: refLine.label,
              fill: '#5b6678',
              fontSize: 10,
              position: 'right',
            }}
          />
        )}
        <Tooltip
          labelFormatter={(v) => fmtDateShort(String(v))}
          formatter={(v: any, name: string) => [
            v == null ? '—' : typeof v === 'number' ? v.toFixed(2).replace(/\.00$/, '') : v,
            METRIC_LABELS[name] || name,
          ]}
        />
        {metrics.map((m) =>
          chartType === 'bar' || chartType === 'sleepBars' || chartType === 'loadBars' ? (
            <Bar
              key={m}
              yAxisId={
                dualAxis && (m === 'sleep_score_0_100' || m === 'Garmin_Training_Readiness' || m === 'coachE_Smart_Gains')
                  ? 'right'
                  : 'left'
              }
              dataKey={m}
              fill={METRIC_COLORS[m] || '#7dd3fc'}
              fillOpacity={0.82}
              radius={[3, 3, 0, 0]}
              maxBarSize={16}
              isAnimationActive={false}
            />
          ) : (
            <Line
              key={m}
              yAxisId={dualAxis && (m === 'Garmin_Training_Readiness' || m === 'coachE_Smart_Gains') ? 'right' : 'left'}
              type="monotone"
              dataKey={m}
              stroke={METRIC_COLORS[m] || '#7dd3fc'}
              strokeWidth={1.75}
              dot={false}
              activeDot={{ r: 3 }}
              connectNulls
              isAnimationActive={false}
            />
          ),
        )}
        {trendKeys.map((tk) => (
          <Line
            key={tk}
            yAxisId="left"
            type="monotone"
            dataKey={tk}
            stroke={METRIC_COLORS[tk] || '#fca5a5'}
            strokeWidth={1.6}
            strokeDasharray="5 4"
            dot={false}
            activeDot={false}
            connectNulls
            isAnimationActive={false}
          />
        ))}
      </ComposedChart>
    </ResponsiveContainer>
  );
}

function ChartSection(props: ChartSectionProps) {
  const {
    title,
    metrics,
    unitHint,
    fullscreen = false,
    onToggleFullscreen,
    trendKeys = [],
  } = props;
  const legendMetrics = [...metrics, ...trendKeys];
  return (
    <div className="panel-raised">
      <header className="px-4 py-2.5 border-b border-border flex items-center justify-between gap-3">
        <div className="min-w-0">
          <h3 className="text-sm font-semibold tracking-tight truncate">{title}</h3>
          {unitHint && <p className="text-2xs text-ink-dim mt-0.5">{unitHint}</p>}
        </div>
        <div className="flex flex-wrap items-center gap-2 justify-end">
          {legendMetrics.map((m) => (
            <span key={m} className="inline-flex items-center gap-1.5 text-2xs text-ink-muted">
              <span
                className="inline-block w-2 h-2 rounded-full"
                style={{ backgroundColor: METRIC_COLORS[m] || '#7dd3fc' }}
              />
              {METRIC_LABELS[m] || m}
            </span>
          ))}
          {onToggleFullscreen && (
            <button
              type="button"
              onClick={onToggleFullscreen}
              className="btn btn-ghost text-2xs px-2 py-1"
              title={fullscreen ? 'Vollbild verlassen' : 'Vollbild'}
              aria-label={fullscreen ? 'Vollbild verlassen' : 'Vollbild'}
            >
              {fullscreen ? '× Schließen' : '⛶ Vollbild'}
            </button>
          )}
        </div>
      </header>
      <div className="px-2 pt-3 pb-2 h-[260px]">
        <ChartBody {...props} />
      </div>
    </div>
  );
}

function FullscreenChart({
  title,
  onClose,
  ...bodyProps
}: ChartSectionProps & { onClose: () => void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      window.removeEventListener('keydown', onKey);
      document.body.style.overflow = prevOverflow;
    };
  }, [onClose]);

  const legendMetrics = [...bodyProps.metrics, ...(bodyProps.trendKeys || [])];

  return (
    <div
      className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4"
      role="dialog"
      aria-modal="true"
      aria-label={`${title} — Vollbild`}
      onClick={onClose}
    >
      <div
        className="panel-raised w-full h-full max-w-[1600px] max-h-[92vh] flex flex-col shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="px-4 py-3 border-b border-border flex items-center justify-between gap-3">
          <div className="min-w-0">
            <h3 className="text-base font-semibold tracking-tight truncate">{title}</h3>
            {bodyProps.unitHint && (
              <p className="text-xs text-ink-dim mt-0.5">{bodyProps.unitHint}</p>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-2 justify-end">
            {legendMetrics.map((m) => (
              <span key={m} className="inline-flex items-center gap-1.5 text-xs text-ink-muted">
                <span
                  className="inline-block w-2 h-2 rounded-full"
                  style={{ backgroundColor: METRIC_COLORS[m] || '#7dd3fc' }}
                />
                {METRIC_LABELS[m] || m}
              </span>
            ))}
            <button
              type="button"
              onClick={onClose}
              className="btn btn-ghost text-xs px-3 py-1"
              title="Vollbild verlassen (ESC)"
              aria-label="Vollbild verlassen"
            >
              × Schließen
            </button>
          </div>
        </header>
        <div className="flex-1 px-2 py-3">
          <ChartBody {...bodyProps} />
        </div>
      </div>
    </div>
  );
}

export function ChartDeck() {
  const [range, setRange] = useState<Range>('28d');
  const [load, setLoad] = useState<ChartDataResponse | null>(null);
  const [recovery, setRecovery] = useState<ChartDataResponse | null>(null);
  const [loadErr, setLoadErr] = useState<string | null>(null);
  const [recoveryErr, setRecoveryErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(true);
  const [fullscreenId, setFullscreenId] = useState<string | null>(null);

  // Metric pickers
  const [loadEnabled, setLoadEnabled] = useState<Record<string, boolean>>(
    Object.fromEntries(LOAD_METRICS.map((m) => [m, true]))
  );
  const [recoveryEnabled, setRecoveryEnabled] = useState<Record<string, boolean>>(
    Object.fromEntries(RECOVERY_METRICS.map((m) => [m, true]))
  );

  const load_ = useCallback(async () => {
    setBusy(true);
    setLoadErr(null);
    setRecoveryErr(null);
    const [a, b] = await Promise.allSettled([
      fetchChartData(range, LOAD_METRICS),
      fetchChartData(range, RECOVERY_METRICS),
    ]);
    if (a.status === 'fulfilled') setLoad(a.value);
    else setLoadErr(a.reason instanceof ApiError ? a.reason.message : String(a.reason));
    if (b.status === 'fulfilled') setRecovery(b.value);
    else setRecoveryErr(b.reason instanceof ApiError ? b.reason.message : String(b.reason));
    setBusy(false);
  }, [range]);

  useEffect(() => {
    load_();
  }, [load_]);

  const loadRows = useMemo(
    () => (load ? normaliseData(load.data, LOAD_METRICS) : []),
    [load]
  );
  const recoveryRowsBase = useMemo(
    () => (recovery ? normaliseData(recovery.data, RECOVERY_METRICS) : []),
    [recovery]
  );
  const recoveryRows = useMemo(
    () => withMovingAverage(recoveryRowsBase, 'rhr_bpm', 'rhr_trend', 7),
    [recoveryRowsBase]
  );

  const activeLoadMetrics = LOAD_METRICS.filter((m) => loadEnabled[m]);
  const activeRecoveryMetrics = RECOVERY_METRICS.filter((m) => recoveryEnabled[m]);

  const loadDayMetrics = activeLoadMetrics.filter((m) => m === 'load_fb_day');
  const atlCtlMetrics = activeLoadMetrics.filter((m) => m === 'fbATL_obs' || m === 'fbCTL_obs');
  const acwrMetrics = activeLoadMetrics.filter((m) => m === 'fbACWR_obs' || m === 'coachE_Smart_Gains');

  // Clip KEI to a plottable band so extreme outliers don't blow up the axis.
  // Values below KEI_MIN are shown as gaps (null), not glued to the axis floor.
  const KEI_MIN = -50;
  const acwrChartRows = useMemo(
    () =>
      loadRows.map((r) => ({
        ...r,
        coachE_Smart_Gains:
          typeof r.coachE_Smart_Gains === 'number' && r.coachE_Smart_Gains < KEI_MIN
            ? null
            : r.coachE_Smart_Gains,
      })),
    [loadRows],
  );
  const enduranceMetrics = activeLoadMetrics.filter((m) => m === 'garminEnduranceScore');

  const sleepMetrics = activeRecoveryMetrics.filter(
    (m) => m === 'sleep_hours' || m === 'sleep_score_0_100'
  );
  const readinessMetrics = activeRecoveryMetrics.filter(
    (m) => m === 'rhr_bpm' || m === 'Garmin_Training_Readiness'
  );

  const rhrActive = activeRecoveryMetrics.includes('rhr_bpm');

  type Section = {
    id: string;
    show: boolean;
    props: ChartSectionProps;
  };

  const sections: Section[] = [
    {
      id: 'load-day',
      show: loadDayMetrics.length > 0,
      props: {
        title: 'Load (Day)',
        unitHint: 'ESS pro Tag · Load als Balken',
        rows: loadRows,
        metrics: loadDayMetrics,
        chartType: 'loadBars',
      },
    },
    {
      id: 'atl-ctl',
      show: atlCtlMetrics.length > 0,
      props: {
        title: 'ATL · CTL',
        unitHint: 'Akute vs. chronische Last',
        rows: loadRows,
        metrics: atlCtlMetrics,
      },
    },
    {
      id: 'acwr-kei',
      show: acwrMetrics.length > 0,
      props: {
        title: 'ACWR · KEI',
        unitHint: 'ACWR links · KEI rechts (fest von -50 nach oben offen; Ausreißer < -50 werden ausgeblendet)',
        rows: acwrChartRows,
        metrics: acwrMetrics,
        yDomain: [0, 'auto'],
        rightDomain: [-50, 'auto'],
        refLine: { y: 1.3, label: '1.3' },
        dualAxis: true,
      },
    },
    {
      id: 'sleep',
      show: sleepMetrics.length > 0,
      props: {
        title: 'Schlaf',
        unitHint: 'Balken · linke Skala Stunden, rechte Skala Score',
        rows: recoveryRows,
        metrics: sleepMetrics,
        chartType: 'sleepBars',
        dualAxis: true,
      },
    },
    {
      id: 'readiness-rhr',
      show: readinessMetrics.length > 0,
      props: {
        title: 'Readiness · RHR',
        unitHint: rhrActive
          ? 'Zweite Skala: RHR links, Readiness rechts · RHR-Trend als 7-Tage-Durchschnitt (gestrichelt)'
          : 'Zweite Skala: RHR links, Readiness rechts',
        rows: recoveryRows,
        metrics: readinessMetrics,
        chartType: 'dualLine',
        dualAxis: true,
        trendKeys: rhrActive ? ['rhr_trend'] : [],
      },
    },
    {
      id: 'endurance',
      show: enduranceMetrics.length > 0,
      props: {
        title: 'Garmin Endurance Score',
        unitHint: 'Garmin Ausdauerwert · langfristige Ausdauerentwicklung',
        rows: loadRows,
        metrics: enduranceMetrics,
      },
    },
  ];

  const fullscreenSection = sections.find((s) => s.id === fullscreenId && s.show);

  return (
    <div className="space-y-5">
      <Panel
        title="Chart Deck"
        right={
          <div className="flex flex-wrap items-center gap-2">
            <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5" role="tablist">
              {RANGES.map((r) => (
                <button
                  key={r}
                  onClick={() => setRange(r)}
                  className={`px-1.5 sm:px-2.5 py-1 text-xs rounded tnum ${
                    range === r
                      ? 'bg-bg-subtle text-ink'
                      : 'text-ink-muted hover:text-ink'
                  }`}
                  data-testid={`range-${r}`}
                >
                  {r}
                </button>
              ))}
            </div>
            <button onClick={load_} className="btn btn-ghost text-xs px-2 py-1" disabled={busy}>
              {busy ? 'Lade…' : 'Aktualisieren'}
            </button>
          </div>
        }
      >
        <div className="text-2xs text-ink-dim tnum mb-3">
          Load · {load?.timestamp ? fmtTime(load.timestamp) : '—'} &nbsp;·&nbsp; Recovery ·{' '}
          {recovery?.timestamp ? fmtTime(recovery.timestamp) : '—'}
        </div>

        {/* Metric toggles */}
        <div className="flex flex-wrap gap-x-4 gap-y-2 mb-4 text-xs">
          <span className="label self-center">Load:</span>
          {LOAD_METRICS.map((m) => (
            <label key={m} className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={!!loadEnabled[m]}
                onChange={(e) =>
                  setLoadEnabled((s) => ({ ...s, [m]: e.target.checked }))
                }
                className="accent-accent"
              />
              <span
                className="inline-block w-2 h-2 rounded-full"
                style={{ backgroundColor: METRIC_COLORS[m] }}
              />
              <span className="text-ink-muted">{METRIC_LABELS[m]}</span>
            </label>
          ))}
          <span className="label self-center ml-2">Recovery:</span>
          {RECOVERY_METRICS.map((m) => (
            <label key={m} className="flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={!!recoveryEnabled[m]}
                onChange={(e) =>
                  setRecoveryEnabled((s) => ({ ...s, [m]: e.target.checked }))
                }
                className="accent-accent"
              />
              <span
                className="inline-block w-2 h-2 rounded-full"
                style={{ backgroundColor: METRIC_COLORS[m] }}
              />
              <span className="text-ink-muted">{METRIC_LABELS[m]}</span>
            </label>
          ))}
        </div>

        {loadErr && <div className="mb-3"><ErrorBox message={`Load: ${loadErr}`} onRetry={load_} /></div>}
        {recoveryErr && <div className="mb-3"><ErrorBox message={`Recovery: ${recoveryErr}`} onRetry={load_} /></div>}

        <div className="grid gap-4 lg:grid-cols-2">
          {busy && !load && !recovery ? (
            <>
              <Skeleton className="h-[260px]" />
              <Skeleton className="h-[260px]" />
            </>
          ) : (
            sections
              .filter((s) => s.show)
              .map((s) => (
                <ChartSection
                  key={s.id}
                  {...s.props}
                  onToggleFullscreen={() => setFullscreenId(s.id)}
                />
              ))
          )}
        </div>
      </Panel>

      <TrainingControl />

      {fullscreenSection && (
        <FullscreenChart
          {...fullscreenSection.props}
          onClose={() => setFullscreenId(null)}
        />
      )}
    </div>
  );
}
