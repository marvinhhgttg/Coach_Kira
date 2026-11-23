import { useEffect, useRef, useState } from 'react';
import { TexteStatusTile } from '../components/TexteStatusTile';
import { AppKopplungTile } from '../components/AppKopplungTile';
import { RechenkernTile } from '../components/RechenkernTile';
import { VotumAbgleichTile } from '../components/VotumAbgleichTile';
import { SollAbgleichTile } from '../components/SollAbgleichTile';
import { IST_V2 } from '../lib/v2guard';
import { GarminNachtTile } from '../components/GarminNachtTile';
import { GarminGesternTile } from '../components/GarminGesternTile';
import { SchlafkontoTile } from '../components/SchlafkontoTile';
import { VorausblickTile } from '../components/VorausblickTile';
import { ErholungsuhrTile } from '../components/ErholungsuhrTile';
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
import { ArcGauge, BandBogen, ErrorBox, KpiKachel, Modal, Panel, Skeleton, Spinner, StorySection, StoryGrid } from '../components/UI';
import { SCORE_SKALA, acwrSkala, farbeVon, nameVon } from '../lib/garminSkala';
import { ACWR_WARN, ACWR_HIGH } from '../kern/votumRegeln';
import { RecoveryPanel } from './RecoveryPanel';
import { AnomalyPanel } from './AnomalyPanel';
import { KiraChat } from './KiraChat';
import { useEngine } from '../lib/engine';
import { runLocalBriefing, fetchBriefingSources, type BriefingSource } from '../lib/briefingLocal';
import { EngineSwitch, SourceBadge } from '../components/EngineSwitch';
import { NutritionTile } from '../components/NutritionTile';

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

function MiniRing({ score }: { score: number | null; farbe?: string }) {
  return (
    <div className="relative flex items-center justify-center shrink-0" style={{ width: 56, height: 56 }}>
      <BandBogen wert={score} baender={SCORE_SKALA.baender} groesse={56} dicke={4} />
      <span className="relative tnum text-sm font-semibold text-ink">{score == null ? '—' : Math.round(score)}</span>
    </div>
  );
}

function DetailScoreRow({ item }: { item: StatusItemLike }) {
  const roh = zahlOderNull(item.score);
  const score = roh != null ? Math.max(0, Math.min(100, roh)) : null;
  const wert = item.wert !== undefined && item.wert !== null && String(item.wert) !== '' ? String(item.wert) : null;
  const farbe = scoreFarbe(score);
  return (
    <div className="grid gap-3 py-3 grid-cols-[56px_1fr] md:grid-cols-[56px_200px_1fr] items-center border-b border-border/70 last:border-b-0">
      <MiniRing score={score} farbe={farbe} />
      <div className="min-w-0">
        <div className="text-sm font-semibold text-ink">{item.metrik || '—'}</div>
        <div className="flex items-baseline gap-2 mt-0.5">
          {wert && <span className="tnum text-base font-semibold text-ink">{wert}</span>}
          <span className="text-2xs font-medium" style={{ color: farbe }}>{score == null ? 'kein Score' : `Score ${Math.round(score)}`}</span>
        </div>
      </div>
      <p className="col-span-2 md:col-span-1 text-xs text-ink-muted leading-relaxed max-w-none">{item.text || '—'}</p>
    </div>
  );
}

// Kiras Einordnung im Stil von „Wichtigste Daten“
const SCORE_HEX = { gruen: '#22a15a', gelb: '#e0a800', orange: '#ef7d00', rot: '#d64545', grau: '#9aa5b1' };
function zahlOderNull(v: unknown): number | null {
  const n = typeof v === 'number' ? v : v == null || v === '' ? NaN : Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}
const wertText = (v: unknown) => { const n = zahlOderNull(v); return n == null ? '—' : Math.round(n); };
function scoreFarbe(v: unknown) {
  const n = zahlOderNull(v);
  // 23.11.2026: Garmin-Farbmuster mit Kiras Score-Schwellen 30/50/70
  return n == null ? SCORE_HEX.grau : farbeVon(SCORE_SKALA, n);
}
const ampelText = (a: unknown) => { const t = String(a || '').trim().toLowerCase(); return t ? t.charAt(0).toUpperCase() + t.slice(1) : 'Status'; };
function acwrText(a: number | null) {
  if (a == null) return 'Belastungsquotient';
  if (a > 1.5) return 'hoch · Korridor 0,8–1,3';
  if (a > 1.3) return 'erhöht · Korridor 0,8–1,3';
  if (a < 0.8) return 'niedrig · Korridor 0,8–1,3';
  return 'im Korridor 0,8–1,3';
}
const ACWR_SK = acwrSkala(ACWR_WARN, ACWR_HIGH);
function AcwrPunkt({ acwr }: { acwr: number | null }) {
  const f = acwr == null ? SCORE_HEX.grau : acwr > 1.5 ? SCORE_HEX.rot : acwr > 1.3 ? SCORE_HEX.orange : acwr < 0.8 ? SCORE_HEX.gelb : SCORE_HEX.gruen;
  return (
    <span className="inline-flex items-center justify-center rounded-full" style={{ width: 44, height: 44, background: `${f}22` }}>
      <span className="rounded-full" style={{ width: 18, height: 18, background: f }} />
    </span>
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
          <h3 className="text-base font-semibold text-ink mb-2">{g.title}</h3>
          <div className="panel px-4">
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
    <Panel title="Status-Radar">
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

  const [engine] = useEngine();
  const [sources, setSources] = useState<Record<string, BriefingSource | null>>({});
  const [analysisLive, setAnalysisLive] = useState<string | null>(null);
  const [consolidatedLive, setConsolidatedLive] = useState<string | null>(null);
  useEffect(() => {
    fetchBriefingSources().then((s) => setSources(s || {}));
  }, []);
  const markSource = (k: 'status' | 'consolidated', local: boolean) =>
    setSources((s) => ({ ...s, [k]: { engine: local ? 'local' : 'cloud', model: local ? 'coach-kira' : 'sonar-pro', timestamp: new Date().toISOString() } }));

  async function onAiAnalysis() {
    setAnalysisBusy(true);
    {
      try {
        setAnalysisLive('');
        const res = await runLocalBriefing('status', (t) => setAnalysisLive(t));
        setAnalysis(res.analysis || null);
        setAnalysisTs(res.timestamp || new Date().toISOString());
        markSource('status', engine === 'local');
        toast('ok', `KI-Statusanalyse ${engine === 'local' ? 'lokal' : 'in der Cloud'} erstellt und gespeichert.`);
      } catch (e) {
        toast('err', e instanceof Error ? e.message : String(e));
      } finally {
        setAnalysisLive(null);
        setAnalysisBusy(false);
      }
      return;
    }
  }

  async function onConsolidatedAnalysis() {
    setConsolidatedBusy(true);
    {
      try {
        setConsolidatedLive('');
        const res = await runLocalBriefing('consolidated', (t) => setConsolidatedLive(t));
        setConsolidated(res.analysis || null);
        setConsolidatedTs(res.timestamp || new Date().toISOString());
        markSource('consolidated', engine === 'local');
        toast('ok', `Konsolidierte Analyse ${engine === 'local' ? 'lokal' : 'in der Cloud'} erstellt und gespeichert.`);
      } catch (e) {
        toast('err', e instanceof Error ? e.message : String(e));
      } finally {
        setConsolidatedLive(null);
        setConsolidatedBusy(false);
      }
      return;
    }
  }

  const [kiOpen, setKiOpen] = useState(false);
  const summary = status?.summary;
  const hb = status?.heartbeat;
  const items = summary?.items || [];
  const findMetric = (needle: string) =>
    items.find((x) => (x.metrik || '').toLowerCase().includes(needle.toLowerCase()));
  const readiness = findMetric('Training Readiness')?.score ?? findMetric('Training Readiness')?.wert;
  const kei = findMetric('KEI')?.score ?? findMetric('KEI')?.wert;
  const nutrition = findMetric('7-Tage-Bilanz')?.score ?? findMetric('Ernährung')?.score ?? 100;
  const acwr = findMetric('ACWR')?.wert;
  const toGaugeValue = (v: unknown) => (typeof v === 'number' ? v : Number(v));

  return (
    <div>
      <RecoveryPanel
        token={token}
        toast={toast}
        proxyAuthenticated={proxy.authenticated}
        fokusRechts={<GarminNachtTile />}
        nachFokus={
          <StorySection id="gestern" titel="Gestern" unter="Was hast du gemacht, und wie schnell bist du davon erholt?">
            <StoryGrid>
              <GarminGesternTile />
              <ErholungsuhrTile />
            </StoryGrid>
          </StorySection>
        }
        nachWoche={
          <>
            <SchlafkontoTile />
            <NutritionTile />
          </>
        }
      />
      <StorySection id="kommt" titel="Was kommt" unter="Fitnesstrend, Bereitschaft der nächsten Tage, Tourfenster und Frühwarnung">
        <VorausblickTile compact />
      </StorySection>
      <StorySection id="kira" titel="Kiras Einordnung" unter="Scores, KI-Analysen und Auffälligkeiten">
      <div className="space-y-4">
      {/* 21.11.2026: gleicher Stil wie „Wichtigste Daten“ (Ring, Wert, Unterzeile); Readiness steht oben schon */}
      <StoryGrid spalten={3}>
        <KpiKachel testid="score-gesamt" skala={SCORE_SKALA} titel={IST_V2 ? 'Supervisor-Score' : 'Gesamt'} ring={zahlOderNull(summary?.gesamtScore)} wert={wertText(summary?.gesamtScore)} farbe={scoreFarbe(summary?.gesamtScore)} unter={IST_V2 ? `Zweitmeinung · ${ampelText(summary?.gesamtAmpel)}` : ampelText(summary?.gesamtAmpel)} />
        <KpiKachel testid="score-training" skala={SCORE_SKALA} titel="Training" ring={zahlOderNull(summary?.trainingScore)} wert={wertText(summary?.trainingScore)} farbe={scoreFarbe(summary?.trainingScore)} unter="Training Score" />
        <KpiKachel testid="score-recovery" skala={SCORE_SKALA} titel="Recovery" ring={zahlOderNull(summary?.recoveryScore)} wert={wertText(summary?.recoveryScore)} farbe={scoreFarbe(summary?.recoveryScore)} unter="Recovery Score" />
        <KpiKachel testid="score-kei" skala={SCORE_SKALA} titel="KEI" ring={zahlOderNull(kei)} wert={wertText(kei)} farbe={scoreFarbe(kei)} unter="Effizienz" />
        <KpiKachel testid="score-ernaehrung" skala={SCORE_SKALA} titel="Ernährung" ring={zahlOderNull(nutrition)} wert={wertText(nutrition)} farbe={scoreFarbe(nutrition)} unter="Bilanz / Score" />
        <KpiKachel
          testid="score-acwr"
          titel="ACWR"
          symbol={zahlOderNull(acwr) == null ? <AcwrPunkt acwr={null} /> : undefined}
          ring={zahlOderNull(acwr)}
          skala={zahlOderNull(acwr) != null ? ACWR_SK : null}
          wert={zahlOderNull(acwr) != null ? (zahlOderNull(acwr) as number).toFixed(2).replace('.', ',') : '—'}
          unter={zahlOderNull(acwr) != null ? `${nameVon(ACWR_SK, zahlOderNull(acwr))} · Korridor 0,8–${String(ACWR_WARN).replace('.', ',')}` : acwrText(null)}
        />
      </StoryGrid>

      {statusErr && <ErrorBox message={statusErr} onRetry={loadStatus} />}

      {/* Status items, if present */}
      {summary?.items && summary.items.length > 0 && (
        <Panel
          title="KI-Lage · Analysen & Detailwerte"
          right={
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => setKiOpen((o) => !o)} data-testid="toggle-ki">
              {kiOpen ? '▴ Einklappen' : '▾ Ausklappen'}
            </button>
          }
        >
          {!kiOpen ? (
            <button className="w-full text-left text-xs text-ink-muted hover:text-ink" onClick={() => setKiOpen(true)}>
              Konsolidierte Analyse {consolidatedTs ? `· Stand ${fmtDateTime(consolidatedTs)}` : '· noch keine'} · KI-Analyse{' '}
              {analysisTs ? `· Stand ${fmtDateTime(analysisTs)}` : '· noch keine'} · {summary.items.length} Detailwerte & Radar – zum Öffnen klicken.
            </button>
          ) : (
          <>
          <div className="mb-3">
            <EngineSwitch />
          </div>
          <div className="mb-4 panel-raised p-4 border-ampel-lila/25">
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="label">Coach-Kira konsolidierte Analyse</div>
                <div className="mt-1"><SourceBadge source={sources.consolidated} /></div>
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
                {consolidatedBusy ? (engine === 'local' ? 'Kira (lokal) schreibt…' : 'Analysiere…') : 'Konsolidierte Analyse erstellen'}
              </button>
            </div>
            {consolidatedLive !== null ? (
              <div className="mt-3 whitespace-pre-wrap text-sm text-ink-muted">{consolidatedLive || (engine === 'local' ? 'Kira (lokal) liest die Daten…' : 'Kira (Cloud) liest die Daten…')}</div>
            ) : consolidated ? (
              <BriefingMarkdown text={consolidated} />
            ) : (
              <p className="mt-3 text-sm text-ink-muted">
                Fasst den Statusbericht, das Plan Cockpit, das Plan-Briefing, die Historie und die Perspektiven
                deiner Garmin-Analysen zu einer gemeinsamen Lagebeurteilung zusammen.
              </p>
            )}
          </div>

          <div className="mb-4 panel-raised p-4 border-accent-strong/20">
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="label">Coach-Kira KI-Analyse</div>
                <div className="mt-1"><SourceBadge source={sources.status} /></div>
                <p className="text-2xs text-ink-dim tnum mt-1">
                  {analysisTs ? `Stand · ${fmtDateTime(analysisTs)}` : 'Noch keine KI-Analyse gespeichert.'}
                </p>
              </div>
              <button
                onClick={onAiAnalysis}
                disabled={analysisBusy}
                className="btn btn-primary text-xs px-2 py-1 shrink-0"
                data-testid="button-ai-analysis"
              >
                {analysisBusy ? (engine === 'local' ? 'Kira (lokal) schreibt…' : 'Analysiere…') : 'KI-Analyse aktualisieren'}
              </button>
            </div>
            {analysisLive !== null ? (
              <div className="mt-3 whitespace-pre-wrap text-sm text-ink-muted">{analysisLive || (engine === 'local' ? 'Kira (lokal) liest die Daten…' : 'Kira (Cloud) liest die Daten…')}</div>
            ) : analysis ? (
              <AiAnalysisText text={analysis} />
            ) : (
              <p className="mt-3 text-sm text-ink-muted">
                Lokal auf deinem Mac oder in der Cloud – das Ergebnis wird im Sheet gespeichert.
              </p>
            )}
          </div>
          <DetailScores items={summary.items} />
          <div className="mt-4">
            <StatusRadarSection items={summary.items} />
          </div>
          </>
          )}
        </Panel>
      )}

      <AnomalyPanel />
      </div>
      </StorySection>
      <StorySection id="chat" titel="Frag Kira" unter="Trainingschat mit deinen Daten">
        <KiraChat />
      </StorySection>

      <StorySection id="texte" titel="Kira-Texte" unter="Status nach dem Absenden – lokal, Cloud-Ersatz, fertig">
        <TexteStatusTile compact />
      </StorySection>

      <StorySection id="technik" titel="Technik" unter="Letzter Lauf der Auswertung">
      <div className="mb-4"><AppKopplungTile /></div>
      <div className="mb-4"><RechenkernTile /></div>
      <div className="mb-4"><VotumAbgleichTile /></div>
      <div className="mb-4"><SollAbgleichTile /></div>
      {/* Letzter Lauf (ganz unten) */}
      <Panel title="Letzter Lauf · Heartbeat">
        {statusLoading ? (
          <Skeleton className="h-16" />
        ) : (
          <dl className="grid gap-x-8 gap-y-2 text-sm sm:grid-cols-2 lg:grid-cols-4">
            <div>
              <dt className="label">{IST_V2 ? 'Letzter Supervisor-Lauf' : 'Letzter KI-Report'}</dt>
              <dd className="tnum mt-1 flex items-center gap-2">
                {summary?.gesamtAmpel && <AmpelDot ampel={summary.gesamtAmpel as any} size={7} />}
                {summary?.gesamtScore != null ? `Score ${summary.gesamtScore}` : '—'}
                {summary?.gesamtAmpel && <span className="text-ink-muted">({summary.gesamtAmpel})</span>}
              </dd>
            </div>
            <div>
              <dt className="label">Zeit</dt>
              <dd className="tnum mt-1">
                {hb?.timestamp ? fmtDateTime(hb.timestamp) : '—'}
                {hb?.timestamp && <span className="text-ink-dim ml-2">{fmtRelative(hb.timestamp)}</span>}
              </dd>
            </div>
            <div>
              <dt className="label">Stage</dt>
              <dd className="tnum mt-1 flex items-center gap-2">
                <AmpelDot ampel="BLAU" size={6} />
                {hb?.stage || '—'}
              </dd>
            </div>
            <div>
              <dt className="label">RunId</dt>
              <dd className="tnum mt-1 text-ink-muted">{hb?.runId || '—'}</dd>
            </div>
          </dl>
        )}
      </Panel>

      <p className="mt-2 text-2xs text-ink-dim tnum">
        Zuletzt aktualisiert {status?.timestamp ? fmtDateTime(status.timestamp) : '—'}
      </p>
      </StorySection>

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
                Starte diese Routine in deiner Perplexity-Computer-Sitzung mit dem oben genannten Auftrag.
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
