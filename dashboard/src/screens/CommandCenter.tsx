import { useEffect, useRef, useState } from 'react';
import {
  fetchStatus,
  fetchStatusAiAnalysis,
  fetchConsolidatedAnalysis,
  fetchRunStatus,
  runConsolidatedAnalysis,
  runStatusAiAnalysis,
  runSupervisor,
  ApiError,
  type StatusResponse,
  type RunStatusResponse,
} from '../lib/api';
import { fmtDateTime, fmtRelative, fmtTime, fmtNum } from '../lib/format';
import { proxyRunConsolidatedAnalysis, proxyRunStatusAiAnalysis, proxyRunSupervisor } from '../lib/proxy';
import { AmpelChip, AmpelDot } from '../components/Ampel';
import { ArcGauge, ErrorBox, Modal, Panel, Skeleton, Spinner } from '../components/UI';
import { RecoveryPanel } from './RecoveryPanel';

type ToastFn = (kind: 'ok' | 'err' | 'info', text: string) => void;
type StatusItemLike = NonNullable<StatusResponse['summary']>['items'] extends Array<infer T> ? T : never;

interface Props {
  token: string;
  setToken: (t: string) => void;
  toast: ToastFn;
  proxy: { available: boolean; configured: boolean; authenticated: boolean };
  pin: string;
  setPin: (v: string) => void;
  setupPin: string;
  setSetupPin: (v: string) => void;
  onProxySetup: () => void;
  onProxyLogin: () => void;
  onProxyLogout: () => void;
}

const GARMIN_ROUTINES = [
  {
    label: 'Morgens starten',
    skill: '/garmin-morgens',
    desc: 'Morgenroutine: Garmin-Status erfassen, Tagesplan abrufen, Supervisor anstossen.',
  },
  {
    label: 'Nach Aktivität starten',
    skill: '/garmin-nach-aktivitaet',
    desc: 'Nach einer Trainingseinheit: Garmin-Aktivität synchronisieren und Plan neu rechnen.',
  },
  {
    label: 'Abends starten',
    skill: '/garmin-abends',
    desc: 'Abendroutine: Recovery-Daten lesen, Schlaf-Vorbereitung, Plan finalisieren.',
  },
];

function renderInlineMarkdown(text: string) {
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  return parts.map((part, i) => {
    if (part.startsWith('**') && part.endsWith('**')) {
      return <strong key={i} className="text-ink font-semibold">{part.slice(2, -2)}</strong>;
    }
    return <span key={i}>{part}</span>;
  });
}

function AiAnalysisText({ text }: { text: string }) {
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  return (
    <div className="mt-3 space-y-2 text-sm leading-relaxed">
      {lines.map((line, i) => {
        const clean = line.replace(/^[-•]\s*/, '');
        const isBullet = /^[-•]\s*/.test(line);
        return (
          <div key={i} className={isBullet ? 'flex gap-2 text-ink-muted' : 'text-ink'}>
            {isBullet && <span className="mt-2 h-1.5 w-1.5 rounded-full bg-accent shrink-0" />}
            <p className={isBullet ? 'max-w-none' : 'max-w-none font-medium'}>
              {renderInlineMarkdown(clean)}
            </p>
          </div>
        );
      })}
    </div>
  );
}

function BriefingMarkdown({ text }: { text: string }) {
  const lines = text.split('\n').map((line) => line.trim()).filter(Boolean);
  const blocks: Array<{ title: string; body: string[] }> = [];
  let current: { title: string; body: string[] } | null = null;
  for (const line of lines) {
    if (line.startsWith('#')) {
      const title = line.replace(/^#+\s*/, '');
      if (current) blocks.push(current);
      current = { title, body: [] };
    } else {
      if (!current) current = { title: 'Executive Summary', body: [] };
      current.body.push(line);
    }
  }
  if (current) blocks.push(current);
  const [hero, ...rest] = blocks;
  return (
    <div className="mt-4 space-y-3">
      {hero && (
        <div className="rounded-lg border border-accent-strong/30 bg-accent-strong/5 p-4">
          <div className="label mb-2">{hero.title}</div>
          <div className="space-y-2 text-sm leading-relaxed text-ink-muted">
            {hero.body.slice(0, 4).map((line, idx) => (
              <p key={idx}>{renderInlineMarkdown(line.replace(/^[-•]\s*/, ''))}</p>
            ))}
          </div>
        </div>
      )}
      <div className="grid gap-3 lg:grid-cols-2">
        {rest.map((block) => (
          <section key={block.title} className="panel-raised p-4">
            <h3 className="text-sm font-semibold uppercase tracking-wide text-ink mb-3">{block.title}</h3>
            <div className="space-y-2 text-sm leading-relaxed text-ink-muted">
              {block.body.map((line, idx) => {
                const isBullet = /^[-•]\s*/.test(line);
                const clean = line.replace(/^[-•]\s*/, '').replace(/^\|\s?/, '').replace(/\s?\|$/, '');
                if (clean.startsWith('---')) return null;
                if (clean.startsWith('|') && clean.replace(/[|:\-\s]/g, '') === '') return null;
                return (
                  <div key={idx} className={isBullet ? 'flex gap-2' : ''}>
                    {isBullet && <span className="mt-2 h-1.5 w-1.5 rounded-full bg-accent shrink-0" />}
                    <p className="max-w-none">{renderInlineMarkdown(clean)}</p>
                  </div>
                );
              })}
            </div>
          </section>
        ))}
      </div>
    </div>
  );
}

function scoreValue(item: StatusItemLike) {
  const s = typeof item.score === 'number' ? item.score : Number(item.score);
  if (Number.isFinite(s)) return Math.max(0, Math.min(100, s));
  const w = typeof item.wert === 'number' ? item.wert : Number(String(item.wert || '').replace(',', '.').replace('%', ''));
  return Number.isFinite(w) ? Math.max(0, Math.min(100, w)) : 0;
}

function DetailScoreRow({ item }: { item: StatusItemLike }) {
  const score = scoreValue(item);
  const marker = `${score}%`;
  const rawValue = item.wert !== undefined && item.wert !== null && String(item.wert) !== ''
    ? item.wert
    : score;
  return (
    <div className="grid gap-3 py-3 md:grid-cols-[220px_1fr_150px_1.4fr] md:items-center border-b border-border/70 last:border-b-0">
      <div>
        <div className="text-sm font-semibold text-ink">{item.metrik || '—'}</div>
        <div className="text-2xs text-ink-dim tnum mt-0.5">{item.kategorie || 'Score'}</div>
      </div>
      <div>
        <div className="relative h-2.5 rounded-full overflow-hidden bg-bg-subtle">
          <div className="absolute inset-y-0 left-0 bg-ampel-rot" style={{ width: '23%' }} />
          <div className="absolute inset-y-0 left-[23%] bg-ampel-orange" style={{ width: '25%' }} />
          <div className="absolute inset-y-0 left-[48%] bg-ampel-gruen" style={{ width: '25%' }} />
          <div className="absolute inset-y-0 left-[73%] bg-accent-strong" style={{ width: '22%' }} />
          <div className="absolute inset-y-0 right-0 bg-ampel-lila" style={{ width: '5%' }} />
          <div
            className="absolute top-1/2 h-4 w-4 -translate-y-1/2 rounded-full border-2 border-bg shadow"
            style={{ left: `calc(${marker} - 8px)`, backgroundColor: '#38bdf8' }}
          />
        </div>
      </div>
      <div className="tnum">
        <div className="text-lg font-semibold text-accent">{String(rawValue)}</div>
        <div className="text-2xs text-ink-muted">SCORE: {fmtNum(score)}</div>
      </div>
      <p className="text-xs text-ink-muted leading-relaxed max-w-none">{item.text || '—'}</p>
    </div>
  );
}

function DetailScores({ items }: { items: StatusItemLike[] }) {
  const groups = [
    {
      title: 'Training',
      items: items.filter((x) => /ACWR|TE Balance|Training Status|KEI/i.test(x.metrik || '')),
    },
    {
      title: 'Recovery',
      items: items.filter((x) => /RHR|Schlafdauer|Schlafscore|HRV|Training Readiness/i.test(x.metrik || '')),
    },
    {
      title: 'Ernährung',
      items: items.filter((x) => /7-Tage|Protein|Ernährung/i.test(x.metrik || '')),
    },
  ].filter((g) => g.items.length);
  return (
    <div className="space-y-5">
      {groups.map((g) => (
        <section key={g.title}>
          <h3 className="text-base font-semibold uppercase tracking-wide text-ink mb-2">{g.title}</h3>
          <div className="panel-raised px-4">
            {g.items.map((item, i) => <DetailScoreRow key={`${g.title}-${item.metrik}-${i}`} item={item} />)}
          </div>
        </section>
      ))}
    </div>
  );
}

const RADAR_MAP = {
  overall: ['RHR', 'Schlafdauer', 'Schlafscore', 'HRV Status', 'Training Readiness', 'TE Balance', 'ACWR', 'Training Status', '7-Tage-Bilanz', 'KEI', 'Protein-Invest'],
  recovery: ['RHR', 'Schlafdauer', 'Schlafscore', 'HRV Status', 'Training Readiness'],
  training: ['ACWR', 'Training Status', 'Protein-Invest', '7-Tage-Bilanz', 'KEI'],
};

function findRadarItem(items: StatusItemLike[], label: string) {
  const n = label.toLowerCase();
  return items.find((item) => (item.metrik || '').toLowerCase().includes(n));
}

function shortLabel(label: string) {
  return label
    .replace('Training Readiness', 'Readiness')
    .replace('7-Tage-Bilanz', '7d-Bilanz')
    .replace('Protein-Invest', 'Protein');
}

function RadarChart({ title, labels, items }: { title: string; labels: string[]; items: StatusItemLike[] }) {
  const cx = 160;
  const cy = 150;
  const maxR = 92;
  const values = labels.map((label) => ({ label, value: scoreValue(findRadarItem(items, label) || ({} as StatusItemLike)) }));
  const axisPoint = (idx: number, radius = maxR) => {
    const angle = -Math.PI / 2 + (Math.PI * 2 * idx) / labels.length;
    return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
  };
  const valuePoint = (idx: number, value: number) => axisPoint(idx, maxR * Math.max(0, Math.min(100, value)) / 100);
  const polygon = values.map((v, i) => valuePoint(i, v.value)).map((p) => `${p.x},${p.y}`).join(' ');
  return (
    <div className="panel-raised p-4">
      <div className="text-lg font-semibold text-ink mb-2">{title}</div>
      <svg viewBox="0 0 320 300" className="w-full h-[300px]" role="img" aria-label={`${title} Radar`}>
        {[25, 50, 75, 100].map((pct) => {
          const pts = labels.map((_, i) => axisPoint(i, maxR * pct / 100)).map((p) => `${p.x},${p.y}`).join(' ');
          return <polygon key={pct} points={pts} fill="none" stroke="#253140" strokeWidth="1" />;
        })}
        {labels.map((_, i) => {
          const p = axisPoint(i);
          return <line key={i} x1={cx} y1={cy} x2={p.x} y2={p.y} stroke="#1f2731" strokeWidth="1" />;
        })}
        <polygon points={polygon} fill="#60a5fa" fillOpacity="0.12" stroke="#60a5fa" strokeWidth="2.4" />
        {values.map((v, i) => {
          const p = valuePoint(i, v.value);
          const lp = axisPoint(i, maxR + 25);
          return (
            <g key={v.label}>
              <circle cx={p.x} cy={p.y} r="3.8" fill="#60a5fa" stroke="#0b0f14" strokeWidth="1.5" />
              <text x={lp.x} y={lp.y} textAnchor={lp.x < cx - 8 ? 'end' : lp.x > cx + 8 ? 'start' : 'middle'} dominantBaseline="middle" className="fill-ink-muted text-[10px] font-medium">
                {shortLabel(v.label)}
              </text>
            </g>
          );
        })}
        {[0, 25, 50, 75, 100].map((tick) => (
          <text key={tick} x={cx + 4} y={cy - maxR * tick / 100} className="fill-ink-dim text-[9px] tnum">
            {tick}
          </text>
        ))}
      </svg>
    </div>
  );
}

function StatusRadarSection({ items }: { items: StatusItemLike[] }) {
  return (
    <Panel title="AI_REPORT_STATUS · Radar">
      <div className="grid gap-4 xl:grid-cols-3">
        <RadarChart title="Gesamtzustand" labels={RADAR_MAP.overall} items={items} />
        <RadarChart title="Recovery" labels={RADAR_MAP.recovery} items={items} />
        <RadarChart title="Training" labels={RADAR_MAP.training} items={items} />
      </div>
    </Panel>
  );
}

export function CommandCenter({
  token,
  setToken,
  toast,
  proxy,
  pin,
  setPin,
  setupPin,
  setSetupPin,
  onProxySetup,
  onProxyLogin,
  onProxyLogout,
}: Props) {
  const [status, setStatus] = useState<StatusResponse | null>(null);
  const [runStatus, setRunStatus] = useState<RunStatusResponse | null>(null);
  const [statusErr, setStatusErr] = useState<string | null>(null);
  const [runErr, setRunErr] = useState<string | null>(null);
  const [analysis, setAnalysis] = useState<string | null>(null);
  const [analysisTs, setAnalysisTs] = useState<string | null>(null);
  const [analysisBusy, setAnalysisBusy] = useState(false);
  const [consolidated, setConsolidated] = useState<string | null>(null);
  const [consolidatedTs, setConsolidatedTs] = useState<string | null>(null);
  const [consolidatedBusy, setConsolidatedBusy] = useState(false);
  const [statusLoading, setStatusLoading] = useState(true);
  const [supervisorBusy, setSupervisorBusy] = useState(false);
  const [supervisorCooldown, setSupervisorCooldown] = useState(0);
  const [activeRoutine, setActiveRoutine] = useState<typeof GARMIN_ROUTINES[number] | null>(null);

  const cooldownTimer = useRef<number | null>(null);

  async function loadStatus() {
    try {
      const s = await fetchStatus();
      setStatus(s);
      setStatusErr(null);
    } catch (e) {
      setStatusErr(e instanceof ApiError ? e.message : String(e));
    } finally {
      setStatusLoading(false);
    }
  }
  async function loadRunStatus() {
    try {
      const r = await fetchRunStatus();
      setRunStatus(r);
      setRunErr(null);
    } catch (e) {
      setRunErr(e instanceof ApiError ? e.message : String(e));
    }
  }

  async function loadAiAnalysis() {
    try {
      const res = await fetchStatusAiAnalysis();
      if (res.ok) {
        setAnalysis(res.analysis || null);
        setAnalysisTs(res.analysisTimestamp || null);
      }
    } catch {
      // Analyse ist optional; Status-UI darf dadurch nicht fehlschlagen.
    }
  }

  async function loadConsolidatedAnalysis() {
    try {
      const res = await fetchConsolidatedAnalysis();
      if (res.ok) {
        setConsolidated(res.analysis || null);
        setConsolidatedTs(res.analysisTimestamp || null);
      }
    } catch {
      // Optional endpoint.
    }
  }

  useEffect(() => {
    loadStatus();
    loadRunStatus();
    loadAiAnalysis();
    loadConsolidatedAnalysis();
    const id1 = window.setInterval(loadStatus, 60_000);
    const id2 = window.setInterval(loadRunStatus, 10_000);
    return () => {
      window.clearInterval(id1);
      window.clearInterval(id2);
    };
  }, []);

  useEffect(() => {
    if (supervisorCooldown <= 0) return;
    cooldownTimer.current = window.setInterval(() => {
      setSupervisorCooldown((c) => Math.max(0, c - 1));
    }, 1000);
    return () => {
      if (cooldownTimer.current) window.clearInterval(cooldownTimer.current);
    };
  }, [supervisorCooldown]);

  async function onSupervisor() {
    setSupervisorBusy(true);
    try {
      const res = proxy.authenticated ? await proxyRunSupervisor() : await runSupervisor(token.trim());
      if (res.ok && res.queued) {
        toast('ok', `Supervisor gestartet · ${res.runId || 'queued'}`);
        setSupervisorCooldown(60);
        // Refresh sequence per contract
        setTimeout(loadRunStatus, 2_000);
        setTimeout(loadStatus, 90_000);
      } else if ((res as any).error) {
        const errMap: Record<string, string> = {
          unauthorized: 'Token ungültig',
          missing_dashboard_token_config: 'Token nicht im Script gesetzt',
          lock_busy: 'Coach Kira läuft bereits',
        };
        const err = (res as any).error;
        toast('err', errMap[err] || `Fehler: ${err}`);
      } else {
        toast('err', 'Unerwartete Antwort');
      }
    } catch (e) {
      toast('err', e instanceof ApiError ? e.message : String(e));
    } finally {
      setSupervisorBusy(false);
    }
  }

  async function onAiAnalysis() {
    setAnalysisBusy(true);
    try {
      const res = proxy.authenticated ? await proxyRunStatusAiAnalysis() : await runStatusAiAnalysis(token.trim());
      if (res.ok) {
        setAnalysis(res.analysis || null);
        setAnalysisTs(res.timestamp || null);
        toast('ok', 'KI-Statusanalyse aktualisiert.');
      } else {
        toast('err', res.error || 'KI-Analyse fehlgeschlagen.');
      }
    } catch (e) {
      toast('err', e instanceof ApiError ? e.message : String(e));
    } finally {
      setAnalysisBusy(false);
    }
  }

  async function onConsolidatedAnalysis() {
    setConsolidatedBusy(true);
    try {
      const res = proxy.authenticated ? await proxyRunConsolidatedAnalysis() : await runConsolidatedAnalysis(token.trim());
      if (res.ok) {
        setConsolidated(res.analysis || null);
        setConsolidatedTs(res.analysisTimestamp || res.timestamp || null);
        toast('ok', 'Konsolidierte Analyse aktualisiert.');
      } else {
        toast('err', res.error || 'Konsolidierte Analyse fehlgeschlagen.');
      }
    } catch (e) {
      toast('err', e instanceof ApiError ? e.message : String(e));
    } finally {
      setConsolidatedBusy(false);
    }
  }

  const summary = status?.summary;
  const hb = status?.heartbeat;
  const lastQueued = runStatus?.lastQueued;
  const items = summary?.items || [];
  const findMetric = (needle: string) =>
    items.find((x) => (x.metrik || '').toLowerCase().includes(needle.toLowerCase()));
  const readiness = findMetric('Training Readiness')?.score ?? findMetric('Training Readiness')?.wert;
  const kei = findMetric('KEI')?.score ?? findMetric('KEI')?.wert;
  const nutrition = findMetric('7-Tage-Bilanz')?.score ?? findMetric('Ernährung')?.score ?? 100;
  const acwr = findMetric('ACWR')?.wert;
  const toGaugeValue = (v: unknown) => (typeof v === 'number' ? v : Number(v));

  return (
    <div className="space-y-6">
      <RecoveryPanel token={token} toast={toast} proxyAuthenticated={proxy.authenticated} />
      <section className="panel overflow-hidden">
        <div>
          <div className="p-4 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
            <ArcGauge label="Gesamt" value={summary?.gesamtScore} sub={summary?.gesamtAmpel || 'Status'} />
            <ArcGauge label="Training" value={summary?.trainingScore} sub="Training Score" />
            <ArcGauge label="Recovery" value={summary?.recoveryScore} sub="Recovery Score" />
            <ArcGauge
              label="Readiness"
              value={toGaugeValue(readiness)}
              sub="Garmin TR"
              tone={toGaugeValue(readiness) < 50 ? 'orange' : 'auto'}
            />
            <ArcGauge label="KEI" value={toGaugeValue(kei)} sub="Effizienz" />
            <ArcGauge label="Ernährung" value={toGaugeValue(nutrition)} sub="Bilanz / Score" />
            <div className="panel-raised p-4 sm:col-span-2 flex flex-col justify-between">
              <div>
                <div className="label">ACWR / Belastungsfenster</div>
                <div className="mt-3 flex items-baseline gap-3">
                  <div className="tnum text-4xl font-semibold text-accent">
                    {typeof acwr === 'number' ? acwr.toFixed(2).replace('.', ',') : String(acwr || '—')}
                  </div>
                  <span className="chip border border-ampel-gruen/40 bg-ampel-gruen/10 text-green-300">
                    Sweet Spot 0,8–1,3
                  </span>
                </div>
              </div>
              <p className="text-xs text-ink-muted leading-relaxed mt-4">
                Ergänzt die bestehende Hauptanzeige um den wichtigsten Belastungsquotienten für die
                Plan-Simulation.
              </p>
            </div>
          </div>
        </div>
      </section>

      {statusErr && <ErrorBox message={statusErr} onRetry={loadStatus} />}

      {/* Middle row: Heartbeat / RunStatus */}
      <div className="grid gap-4 lg:grid-cols-2">
        <Panel title="Heartbeat">
          {statusLoading ? (
            <Skeleton className="h-20" />
          ) : hb ? (
            <dl className="grid grid-cols-[100px_1fr] gap-y-2 text-sm">
              <dt className="text-ink-muted">Stage</dt>
              <dd className="tnum flex items-center gap-2">
                <AmpelDot ampel="BLAU" size={6} />
                {hb.stage || '—'}
              </dd>
              <dt className="text-ink-muted">RunId</dt>
              <dd className="tnum text-ink-muted">{hb.runId || '—'}</dd>
              <dt className="text-ink-muted">Zeit</dt>
              <dd className="tnum">
                {fmtTime(hb.timestamp)}
                <span className="text-ink-dim ml-2">{fmtRelative(hb.timestamp)}</span>
              </dd>
            </dl>
          ) : (
            <p className="text-sm text-ink-muted">Kein Heartbeat.</p>
          )}
        </Panel>

        <Panel title="Letzter Queue">
          {runErr ? (
            <ErrorBox message={runErr} onRetry={loadRunStatus} />
          ) : lastQueued ? (
            <dl className="grid grid-cols-[100px_1fr] gap-y-2 text-sm">
              <dt className="text-ink-muted">Funktion</dt>
              <dd className="tnum">{lastQueued.functionName || '—'}</dd>
              <dt className="text-ink-muted">RunId</dt>
              <dd className="tnum text-ink-muted">{lastQueued.runId || '—'}</dd>
              <dt className="text-ink-muted">Queued</dt>
              <dd className="tnum">
                {fmtTime(lastQueued.queuedAt)}
                <span className="text-ink-dim ml-2">{fmtRelative(lastQueued.queuedAt)}</span>
              </dd>
            </dl>
          ) : (
            <p className="text-sm text-ink-muted">Keine Queue-Einträge.</p>
          )}
        </Panel>

      </div>

      {/* Action row */}
      <Panel
        title="Aktionen"
        right={
          <span className="text-2xs text-ink-dim tnum">
            {status?.timestamp && <>Status · {fmtTime(status.timestamp)}</>}
          </span>
        }
      >
        <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-4">
          <button
            onClick={onSupervisor}
            disabled={supervisorBusy || supervisorCooldown > 0}
            className="btn btn-primary justify-between text-left"
            data-testid="button-supervisor"
          >
            <span className="flex items-center gap-2">
              {supervisorBusy && <Spinner />}
              Supervisor starten
            </span>
            <span className="text-2xs text-ink-muted tnum">
              {supervisorCooldown > 0 ? `${supervisorCooldown}s` : 'runSupervisor'}
            </span>
          </button>
          {GARMIN_ROUTINES.map((r) => (
            <button
              key={r.skill}
              onClick={() => setActiveRoutine(r)}
              className="btn justify-between text-left"
              data-testid={`button-routine-${r.skill}`}
            >
              <span>{r.label}</span>
              <span className="text-2xs text-ink-muted tnum">Agent</span>
            </button>
          ))}
        </div>
      </Panel>

      {/* Status items, if present */}
      {summary?.items && summary.items.length > 0 && (
        <Panel title="Status-Items">
          <div className="mb-4 panel-raised p-4 border-ampel-lila/25">
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="label">Coach-Kira konsolidierte Analyse</div>
                <p className="text-2xs text-ink-dim tnum mt-1">
                  {consolidatedTs ? `Stand · ${fmtDateTime(consolidatedTs)}` : 'Noch keine konsolidierte Analyse gespeichert.'}
                </p>
              </div>
              <button
                onClick={onConsolidatedAnalysis}
                disabled={consolidatedBusy}
                className="btn btn-primary text-xs px-2 py-1 shrink-0"
                data-testid="button-consolidated-analysis"
              >
                {consolidatedBusy ? 'Analysiere…' : 'Konsolidierte Analyse erstellen'}
              </button>
            </div>
            {consolidated ? (
              <BriefingMarkdown text={consolidated} />
            ) : (
              <p className="mt-3 text-sm text-ink-muted">
                Synthetisiert AI_REPORT_STATUS, Plan Cockpit, Plan-Briefing, Historie und die Perspektiven
                deiner Garmin-Analyse-Skills zu einer konsolidierten Coach-Kira-Lagebeurteilung.
              </p>
            )}
          </div>

          <div className="mb-4 panel-raised p-4 border-accent-strong/20">
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="label">Coach-Kira KI-Analyse</div>
                <p className="text-2xs text-ink-dim tnum mt-1">
                  {analysisTs ? `Stand · ${fmtDateTime(analysisTs)}` : 'Noch keine Perplexity-Analyse gespeichert.'}
                </p>
              </div>
              <button
                onClick={onAiAnalysis}
                disabled={analysisBusy}
                className="btn btn-primary text-xs px-2 py-1 shrink-0"
                data-testid="button-ai-analysis"
              >
                {analysisBusy ? 'Analysiere…' : 'KI-Analyse aktualisieren'}
              </button>
            </div>
            {analysis ? (
              <AiAnalysisText text={analysis} />
            ) : (
              <p className="mt-3 text-sm text-ink-muted">
                Nutzt Perplexity serverseitig über Apps Script und speichert das Ergebnis in
                <span className="tnum text-ink"> AI_STATUS_ANALYSIS</span>.
              </p>
            )}
          </div>
          <DetailScores items={summary.items} />
        </Panel>
      )}

      {summary?.items && summary.items.length > 0 && (
        <StatusRadarSection items={summary.items} />
      )}

      <p className="text-2xs text-ink-dim tnum">
        Zuletzt aktualisiert {status?.timestamp ? fmtDateTime(status.timestamp) : '—'}
      </p>

      <Modal
        open={!!activeRoutine}
        onClose={() => setActiveRoutine(null)}
        title={activeRoutine?.label || ''}
      >
        {activeRoutine && (
          <div className="space-y-4">
            <p className="text-ink leading-relaxed">{activeRoutine.desc}</p>
            <div className="rounded border border-accent-strong/30 bg-accent-strong/5 p-3 text-sm">
              <div className="label mb-2">Agent-Auftrag</div>
              <code className="block tnum text-accent text-base">{activeRoutine.skill}</code>
              <p className="text-2xs text-ink-muted mt-2 leading-relaxed">
                Diese Routine wird über den Computer-Agent gestartet — der Prototyp führt
                keine Browser-Automation aus. Starte den Skill in deiner Computer-Sitzung.
              </p>
            </div>
            <div className="text-2xs text-ink-dim leading-relaxed">
              Erwarteter Abschluss: Formular gesendet · Browser zurück zu Garmin · Supervisor
              läuft über bestehenden Workflow.
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}
