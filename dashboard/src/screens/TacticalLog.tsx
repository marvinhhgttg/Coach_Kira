import { useEffect, useMemo, useState } from 'react';
import {
  ACTIVITY_METRICS,
  ApiError,
  fetchChartData,
  fetchRunStatus,
  toNum,
  type ChartDataResponse,
  type Range,
  type RunStatusResponse,
} from '../lib/api';
import { fmtDateShort, fmtRelative, fmtTime, fmtNum } from '../lib/format';
import { ErrorBox, Panel, Skeleton } from '../components/UI';
import { AmpelDot } from '../components/Ampel';

const ACTIVITY_RANGES: Range[] = ['28d', '60d', '90d', '180d', '360d'];

function activityLabel(row: Record<string, any>) {
  return row.Sport_x || row.activity_done || 'Aktivität';
}

function normalizeActivityRows(rows: any[]) {
  return rows
    .filter((r) => r && typeof r.date === 'string')
    .map((r) => ({
      date: r.date,
      load: toNum(r.load_fb_day),
      sport: r.Sport_x || '',
      zone: r.Zone || '',
      activity: r.activity_done || '',
      teAe: toNum(r.Aerobic_TE),
      teAn: toNum(r.Anaerobic_TE),
    }))
    .filter((r) => r.load !== null || r.sport || r.activity || r.teAe !== null || r.teAn !== null)
    .sort((a, b) => b.date.localeCompare(a.date));
}

function flagForLoad(load: number | null) {
  if (load === null) return 'GRAU';
  if (load >= 220) return 'ORANGE';
  if (load >= 130) return 'BLAU';
  if (load > 0) return 'GRÜN';
  return 'GRAU';
}

const FLAG_LEGEND = [
  { ampel: 'GRAU', label: 'Off / keine Last', hint: '0 ESS oder leer' },
  { ampel: 'GRÜN', label: 'Locker', hint: '1–129 ESS' },
  { ampel: 'BLAU', label: 'Training', hint: '130–219 ESS' },
  { ampel: 'ORANGE', label: 'Hohe Last', hint: '≥ 220 ESS' },
];

export function TacticalLog() {
  const [range, setRange] = useState<Range>('90d');
  const [activityData, setActivityData] = useState<ChartDataResponse | null>(null);
  const [runStatus, setRunStatus] = useState<RunStatusResponse | null>(null);
  const [activityErr, setActivityErr] = useState<string | null>(null);
  const [runErr, setRunErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(true);

  async function load(showSpinner = true) {
    if (showSpinner) setBusy(true);
    const [a, b] = await Promise.allSettled([
      fetchChartData(range, ACTIVITY_METRICS),
      fetchRunStatus(),
    ]);
    if (a.status === 'fulfilled') {
      setActivityData(a.value);
      setActivityErr(null);
    } else {
      setActivityErr(a.reason instanceof ApiError ? a.reason.message : String(a.reason));
    }
    if (b.status === 'fulfilled') {
      setRunStatus(b.value);
      setRunErr(null);
    } else {
      setRunErr(b.reason instanceof ApiError ? b.reason.message : String(b.reason));
    }
    setBusy(false);
  }

  useEffect(() => {
    load();
    const id = window.setInterval(() => load(false), 60_000);
    return () => window.clearInterval(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [range]);

  const rows = useMemo(
    () => normalizeActivityRows(activityData?.data || []),
    [activityData]
  );

  const totalLoad = rows.reduce((sum, r) => sum + (r.load || 0), 0);
  const hb = runStatus?.heartbeat;

  return (
    <div className="space-y-5">
      <div className="grid gap-4 lg:grid-cols-3">
        <Panel title="Aktivitäten">
          <div className="tnum text-3xl font-semibold">{rows.length}</div>
          <div className="text-2xs text-ink-dim mt-1">Einträge im Zeitraum</div>
        </Panel>
        <Panel title="Gesamtload">
          <div className="tnum text-3xl font-semibold text-accent">{fmtNum(totalLoad)}</div>
          <div className="text-2xs text-ink-dim mt-1">ESS Summe</div>
        </Panel>
        <Panel title="Durchschnitt">
          <div className="tnum text-3xl font-semibold">
            {fmtNum(rows.length ? totalLoad / rows.length : null)}
          </div>
          <div className="text-2xs text-ink-dim mt-1">ESS pro Eintrag</div>
        </Panel>
      </div>

      <Panel
        title="Tactical Log · geleistete Aktivitäten"
        right={
          <div className="flex flex-wrap items-center gap-2 justify-end">
            <div className="inline-flex bg-bg rounded border border-border p-0.5">
              {ACTIVITY_RANGES.map((r) => (
                <button
                  key={r}
                  onClick={() => setRange(r)}
                  className={`px-2.5 py-1 text-xs rounded tnum ${
                    range === r ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'
                  }`}
                >
                  {r}
                </button>
              ))}
            </div>
            <button
              onClick={() => load()}
              className="btn btn-ghost text-xs px-2 py-1"
              data-testid="button-refresh-activities"
              disabled={busy}
            >
              {busy ? 'Lade…' : 'Aktualisieren'}
            </button>
          </div>
        }
      >
        <div className="text-2xs text-ink-dim tnum mb-3 flex flex-wrap gap-4">
          <span>Quelle · KK_TIMELINE / chartData</span>
          <span>Stand · {activityData?.timestamp ? fmtTime(activityData.timestamp) : '—'}</span>
          {hb && (
            <span className="flex items-center gap-1.5">
              <AmpelDot ampel="BLAU" size={5} />
              Heartbeat {hb.stage || '—'} · {fmtRelative(hb.timestamp)}
            </span>
          )}
        </div>

        <div className="mb-3 grid gap-2 sm:grid-cols-2 xl:grid-cols-4">
          {FLAG_LEGEND.map((f) => (
            <div key={f.label} className="panel-raised px-3 py-2 flex items-center gap-2">
              <AmpelDot ampel={f.ampel} size={8} />
              <div>
                <div className="text-xs text-ink">{f.label}</div>
                <div className="text-2xs text-ink-dim tnum">{f.hint}</div>
              </div>
            </div>
          ))}
        </div>

        {activityErr ? (
          <ErrorBox message={activityErr} onRetry={() => load()} />
        ) : busy && !activityData ? (
          <div className="space-y-2">
            {Array.from({ length: 8 }).map((_, i) => (
              <Skeleton key={i} className="h-12" />
            ))}
          </div>
        ) : !rows.length ? (
          <p className="text-sm text-ink-muted">Keine geleisteten Aktivitäten im Zeitraum.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="min-w-[980px] text-xs">
              <thead>
                <tr className="text-left border-b border-border">
                  {['Datum', 'Sport / Aktivität', 'Zone', 'Load', 'TE AE/AN', 'Flag'].map((h) => (
                    <th key={h} className="label px-2 py-2">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {rows.map((r, i) => (
                  <tr key={`${r.date}-${i}`} className="hover:bg-bg-subtle/50">
                    <td className="px-2 py-2 tnum font-medium">{fmtDateShort(r.date)}</td>
                    <td className="px-2 py-2">
                      <div className="text-ink">{r.sport || activityLabel(r)}</div>
                      {r.activity && <div className="text-2xs text-ink-dim mt-0.5">{r.activity}</div>}
                    </td>
                    <td className="px-2 py-2"><span className="chip border border-border-strong bg-bg-subtle">{r.zone || '—'}</span></td>
                    <td className="px-2 py-2 tnum font-semibold text-accent">{fmtNum(r.load)}</td>
                    <td className="px-2 py-2 tnum">{fmtNum(r.teAe, 1)} / {fmtNum(r.teAn, 1)}</td>
                    <td className="px-2 py-2"><AmpelDot ampel={flagForLoad(r.load)} size={8} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {runErr && <div className="mt-3"><ErrorBox message={`RunStatus: ${runErr}`} onRetry={() => load()} /></div>}
      </Panel>
    </div>
  );
}
