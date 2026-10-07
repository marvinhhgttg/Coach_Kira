import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import {
  Area,
  Bar,
  CartesianGrid,
  ComposedChart,
  Line,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { ApiError, dedupeByDateKeepLast, fetchChartData, fetchPlanSimulationWithTimeout, saveWellbeing, toNum } from '../lib/api';
import { proxySaveWellbeing } from '../lib/proxy';
import { fmtDateShort, fmtNum } from '../lib/format';
import { ErrorBox, Panel, Skeleton } from '../components/UI';

// ---------------------------------------------------------------------------
// Konfiguration
// ---------------------------------------------------------------------------
const SLEEP_GOAL_H = 7.5;
const RHR_BASELINE_DAYS = 28;
const RHR_BASELINE_MIN_VALUES = 7;

const RECOVERY_FETCH_METRICS = [
  'Garmin_Training_Readiness',
  'rhr_bpm',
  'sleep_hours',
  'sleep_score_0_100',
  'hrv_status',
  'hrv_threshholds',
  'load_fb_day',
  'Monotony7',
  'befinden_1_5',
  'fbACWR_obs',
  'coachE_ACWR_forecast',
  'coachE_ESS_day',
  'Sport_x',
  'Zone',
  'Target_Aerobic_TE',
  'Target_Anaerobic_TE',
  'activity_done',
  'coachE_ATL_forecast',
  'coachE_CTL_forecast',
  'Aerobic_TE',
  'Anaerobic_TE',
] as const;

const MONOTONY_WARN = 1.6;
const MONOTONY_HIGH = 2.0;
const ACWR_WARN = 1.3;
const ACWR_HIGH = 1.5;

type Level = 'gruen' | 'gelb' | 'orange' | 'rot' | 'grau' | 'leer';
type Vote = 'QUALITY' | 'TRAIN' | 'EASY' | 'REST';

const VOTE_ORDER: Vote[] = ['REST', 'EASY', 'TRAIN', 'QUALITY'];
const VOTE_LABEL: Record<Vote, string> = {
  QUALITY: 'Quality',
  TRAIN: 'Train',
  EASY: 'Easy',
  REST: 'Rest',
};
const VOTE_LEVEL: Record<Vote, Level> = {
  QUALITY: 'gruen',
  TRAIN: 'gruen',
  EASY: 'gelb',
  REST: 'rot',
};
const VOTE_ADVICE: Record<Vote, string> = {
  QUALITY: 'Qualitätseinheit möglich (Schwelle/Intervalle), wenn sie geplant ist.',
  TRAIN: 'Geplantes Training normal umsetzen, Schwerpunkt Z2.',
  EASY: '45–75 min Z1–Z2, keine Intervalle; morgen neu bewerten.',
  REST: 'Ruhetag oder max. 30–45 min sehr locker / Mobility.',
};
const VOTE_STATUS_TEXT: Record<Vote, string> = {
  QUALITY: 'GRÜN — Qualität möglich',
  TRAIN: 'GRÜN — normal trainieren',
  EASY: 'GELB — locker trainieren',
  REST: 'ROT — Erholung priorisieren',
};

const LEVEL_DOT: Record<Level, string> = {
  gruen: 'bg-ampel-gruen',
  gelb: 'bg-ampel-gelb',
  orange: 'bg-ampel-orange',
  rot: 'bg-ampel-rot',
  grau: 'bg-ampel-grau',
  leer: 'bg-transparent border border-border-strong',
};
const LEVEL_CELL: Record<Level, string> = {
  gruen: 'bg-ampel-gruen/10',
  gelb: 'bg-ampel-gelb/10',
  orange: 'bg-ampel-orange/10',
  rot: 'bg-ampel-rot/10',
  grau: 'bg-ampel-grau/10',
  leer: '',
};
const LEVEL_TEXT: Record<Level, string> = {
  gruen: 'text-green-300',
  gelb: 'text-yellow-300',
  orange: 'text-orange-300',
  rot: 'text-red-300',
  grau: 'text-gray-300',
  leer: 'text-ink-dim',
};

// ---------------------------------------------------------------------------
// Datenmodell
// ---------------------------------------------------------------------------
type RecoveryDay = {
  date: string;
  readiness: number | null;
  rhr: number | null;
  rhrBaseline: number | null;
  rhrDelta: number | null;
  sleepH: number | null;
  sleepScore: number | null;
  sleepAvg7: number | null;
  hrv: number | null;
  hrvLow: number | null;
  hrvHigh: number | null;
  load: number | null;
  monotony: number | null;
  acwr: number | null;
  activityDone: boolean;
  actual: { load: number | null; sport: string; zone: string; teAe: number | null; teAn: number | null };
  essDay: number | null;
  atlPost: number | null;
  ctlPost: number | null;
  monotonyPost: number | null;
  acwrPost: number | null;
  befinden: number | null;
  hasMorningData: boolean;
  levels: { readiness: Level; rhr: Level; sleep: Level; hrv: Level; befinden: Level };
  vote: Vote | null;
  reasons: string[];
};

function localIsoDate(d = new Date()) {
  // yyyy-mm-dd in lokaler Zeitzone (Europe/Berlin im Browser)
  return d.toLocaleDateString('sv-SE');
}

function parseThresholds(v: unknown): [number | null, number | null] {
  if (typeof v !== 'string' || !v.includes(';')) return [null, null];
  const [a, b] = v.split(';').map((x) => Number(String(x).trim().replace(',', '.')));
  return [Number.isFinite(a) ? a : null, Number.isFinite(b) ? b : null];
}

function mean(values: number[]) {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : null;
}

function levelReadiness(v: number | null): Level {
  if (v == null) return 'leer';
  if (v >= 75) return 'gruen';
  if (v >= 55) return 'gelb';
  if (v >= 40) return 'orange';
  return 'rot';
}

function levelRhrDelta(d: number | null): Level {
  if (d == null) return 'leer';
  // Ungewöhnlich niedrig ist nicht automatisch gut (Messartefakt / atypisches Signal)
  if (d <= -5) return 'grau';
  if (d <= 1) return 'gruen';
  if (d < 3) return 'gelb';
  if (d < 5) return 'orange';
  return 'rot';
}

function levelSleep(h: number | null): Level {
  if (h == null) return 'leer';
  const ratio = h / SLEEP_GOAL_H;
  if (ratio >= 0.95) return 'gruen';
  if (ratio >= 0.85) return 'gelb';
  if (ratio >= 0.8) return 'orange';
  return 'rot';
}

function levelHrv(v: number | null, low: number | null): Level {
  if (v == null || low == null) return 'leer';
  if (v < low - 2) return 'rot'; // entspricht SG_Flag "H"
  if (v < low) return 'orange';
  return 'gruen';
}

function levelBefinden(v: number | null): Level {
  if (v == null) return 'leer';
  if (v >= 4) return 'gruen';
  if (v >= 3) return 'gelb';
  if (v >= 2) return 'orange';
  return 'rot';
}

function capVote(vote: Vote, max: Vote): Vote {
  return VOTE_ORDER.indexOf(vote) > VOTE_ORDER.indexOf(max) ? max : vote;
}

/**
 * Transparente Regel-Logik (Empfehlung, keine Diagnose):
 * 1. Basis aus Garmin Readiness: ≥75 Quality · 55–74 Train · 40–54 Easy · <40 Rest
 * 2. Ein roter Einzelwert (RHR, Schlaf, HRV) → keine Qualität (max. Train)
 * 3. ΔRHR ≥ +5 bpm und Schlaf < 80 % Ziel → max. Easy
 * 4. HRV-Alarm (unter Normalband − 2, wie SG_Flag „H“) → max. Easy
 * 5. Zwei oder mehr Warnsignale (orange/rot) → max. Easy
 * 6. ΔRHR ≥ +5 bpm und Readiness < 55 → Rest
 * 7. Befinden 3 → max. Train · 2 → max. Easy · 1 → Rest
 * 8. Belastung: Monotony7 > 1,6 oder ACWR > 1,3 → max. Train · Monotony7 > 2,0 oder ACWR > 1,5 → max. Easy
 */
function decideVote(d: Omit<RecoveryDay, 'vote' | 'reasons'>): { vote: Vote | null; reasons: string[] } {
  if (d.readiness == null) return { vote: null, reasons: ['Readiness fehlt (Morgenwerte noch nicht erfasst).'] };
  const reasons: string[] = [];
  let vote: Vote = d.readiness >= 75 ? 'QUALITY' : d.readiness >= 55 ? 'TRAIN' : d.readiness >= 40 ? 'EASY' : 'REST';
  reasons.push(`Readiness ${fmtNum(d.readiness)} → Basis ${VOTE_LABEL[vote]}.`);

  const { rhr, sleep, hrv } = d.levels;
  const warn = [rhr, sleep, hrv].filter((l) => l === 'orange' || l === 'rot').length;

  if ([rhr, sleep, hrv].includes('rot')) {
    const before = vote;
    vote = capVote(vote, 'TRAIN');
    if (vote !== before) reasons.push('Ein Einzelsignal ist rot → keine Qualitätseinheit.');
  }
  const sleepRatio = d.sleepH != null ? d.sleepH / SLEEP_GOAL_H : null;
  if (d.rhrDelta != null && d.rhrDelta >= 5 && sleepRatio != null && sleepRatio < 0.8) {
    vote = capVote(vote, 'EASY');
    reasons.push('RHR ≥ +5 bpm und Schlaf < 80 % des Ziels → maximal Easy.');
  }
  if (hrv === 'rot') {
    vote = capVote(vote, 'EASY');
    reasons.push('HRV unter dem Normalband (HRV-Alarm) → maximal Easy.');
  }
  if (warn >= 2) {
    vote = capVote(vote, 'EASY');
    reasons.push(`${warn} Warnsignale gleichzeitig → maximal Easy.`);
  }
  if (d.rhrDelta != null && d.rhrDelta >= 5 && d.readiness < 55) {
    vote = 'REST';
    reasons.push('RHR ≥ +5 bpm bei niedriger Readiness → Rest.');
  }
  if ((d.monotony != null && d.monotony > MONOTONY_HIGH) || (d.acwr != null && d.acwr > ACWR_HIGH)) {
    const before = vote;
    vote = capVote(vote, 'EASY');
    const why = [
      d.monotony != null && d.monotony > MONOTONY_HIGH ? `Monotonie ${fmtNum(d.monotony, 2)} > ${fmtNum(MONOTONY_HIGH, 1)}` : '',
      d.acwr != null && d.acwr > ACWR_HIGH ? `ACWR ${fmtNum(d.acwr, 2)} > ${fmtNum(ACWR_HIGH, 1)}` : '',
    ].filter(Boolean).join(' und ');
    if (vote !== before) reasons.push(`${why} → maximal Easy.`);
  } else if ((d.monotony != null && d.monotony > MONOTONY_WARN) || (d.acwr != null && d.acwr > ACWR_WARN)) {
    const before = vote;
    vote = capVote(vote, 'TRAIN');
    const why = [
      d.monotony != null && d.monotony > MONOTONY_WARN ? `Monotonie ${fmtNum(d.monotony, 2)} > ${fmtNum(MONOTONY_WARN, 1)}` : '',
      d.acwr != null && d.acwr > ACWR_WARN ? `ACWR ${fmtNum(d.acwr, 2)} > ${fmtNum(ACWR_WARN, 1)}` : '',
    ].filter(Boolean).join(' und ');
    if (vote !== before) reasons.push(`${why} → keine Qualitätseinheit.`);
  }
  if (d.befinden != null) {
    if (d.befinden <= 1) {
      vote = 'REST';
      reasons.push('Befinden 1/5 → Rest.');
    } else if (d.befinden <= 2) {
      vote = capVote(vote, 'EASY');
      reasons.push('Befinden 2/5 → maximal Easy.');
    } else if (d.befinden <= 3) {
      const before = vote;
      vote = capVote(vote, 'TRAIN');
      if (vote !== before) reasons.push('Befinden 3/5 → keine Qualitätseinheit.');
    }
  }
  return { vote, reasons };
}

function buildDays(raw: any[]): RecoveryDay[] {
  const rows = dedupeByDateKeepLast(
    raw.filter((r) => r && typeof r.date === 'string').map((r) => ({ ...r, date: String(r.date).slice(0, 10) })),
  ).sort((a: any, b: any) => (a.date < b.date ? -1 : 1));

  return rows.map((r: any, idx: number) => {
    const readiness = toNum(r.Garmin_Training_Readiness);
    const rhr = toNum(r.rhr_bpm);
    const sleepH = toNum(r.sleep_hours);
    const sleepScore = toNum(r.sleep_score_0_100);
    const hrv = toNum(r.hrv_status);
    const [hrvLow, hrvHigh] = parseThresholds(r.hrv_threshholds);

    const prior = rows
      .slice(Math.max(0, idx - RHR_BASELINE_DAYS), idx)
      .map((p: any) => toNum(p.rhr_bpm))
      .filter((v): v is number => v != null && v > 0);
    const rhrBaseline = prior.length >= RHR_BASELINE_MIN_VALUES ? mean(prior) : null;
    const rhrDelta = rhr != null && rhrBaseline != null ? rhr - rhrBaseline : null;

    const sleepWindow = rows
      .slice(Math.max(0, idx - 6), idx + 1)
      .map((p: any) => toNum(p.sleep_hours))
      .filter((v): v is number => v != null && v > 0);
    const sleepAvg7 = sleepWindow.length >= 3 ? mean(sleepWindow) : null;

    const befindenRaw = toNum(r.befinden_1_5);
    const befinden = befindenRaw != null && befindenRaw >= 1 && befindenRaw <= 5 ? Math.round(befindenRaw) : null;
    const base = {
      date: r.date,
      readiness,
      rhr,
      rhrBaseline,
      rhrDelta,
      sleepH,
      sleepScore,
      sleepAvg7,
      hrv,
      hrvLow,
      hrvHigh,
      load: toNum(r.load_fb_day),
      // Belastung bis GESTERN (heutige Zeile enthält bereits die heutige Einheit)
      monotony: idx > 0 ? toNum(rows[idx - 1].Monotony7) : null,
      acwr: idx > 0 ? toNum(rows[idx - 1].fbACWR_obs) ?? toNum(rows[idx - 1].coachE_ACWR_forecast) : null,
      activityDone: String(r.activity_done || '').trim().toLowerCase() === 'x',
      essDay: toNum(r.coachE_ESS_day),
      atlPost: toNum(r.coachE_ATL_forecast),
      ctlPost: toNum(r.coachE_CTL_forecast),
      monotonyPost: toNum(r.Monotony7),
      acwrPost: toNum(r.fbACWR_obs) ?? toNum(r.coachE_ACWR_forecast),
      actual: {
        teAe: toNum(r.Aerobic_TE),
        teAn: toNum(r.Anaerobic_TE),
        load: toNum(r.coachE_ESS_day),
        sport: String(r.Sport_x || ''),
        zone: String(r.Zone || ''),
      },
      befinden,
      hasMorningData: readiness != null || rhr != null || sleepH != null,
      levels: {
        readiness: levelReadiness(readiness),
        rhr: levelRhrDelta(rhrDelta),
        sleep: levelSleep(sleepH),
        hrv: levelHrv(hrv, hrvLow),
        befinden: levelBefinden(befinden),
      },
    };
    const { vote, reasons } = decideVote(base);
    return { ...base, vote, reasons };
  });
}

// ---------------------------------------------------------------------------
// Plan vs. Votum
// ---------------------------------------------------------------------------
type PlannedSession = {
  date: string;
  load: number;
  sport: string;
  zone: string;
  teAe: number;
  teAn: number;
  locked: boolean;
};

/** Intensität der geplanten Einheit auf die Votum-Skala abbilden. */
function plannedDemand(p: PlannedSession): Vote {
  const zone = (p.zone || '').toUpperCase();
  const sport = (p.sport || '').toLowerCase();
  if (p.load <= 0 || sport === 'off' || zone === 'OFF') return 'REST';
  // Quality = spürbar intensiv: anaerober Reiz, hohe Zonen, sehr hoher aerober Effekt oder sehr hohe Last
  if (p.teAn >= 1 || p.teAe >= 4.5 || /Z4|Z5/.test(zone) || p.load >= 200) return 'QUALITY';
  if (p.teAe >= 3 || /Z3/.test(zone) || p.load >= 120) return 'TRAIN';
  return 'EASY';
}

function planVerdict(p: PlannedSession, vote: Vote): { ok: boolean; text: string } {
  const demand = plannedDemand(p);
  if (VOTE_ORDER.indexOf(demand) <= VOTE_ORDER.indexOf(vote)) {
    return { ok: true, text: `Plan passt zum Votum (${VOTE_LABEL[demand]} ≤ ${VOTE_LABEL[vote]}).` };
  }
  const shorter = Math.max(30, Math.round((p.load * 0.6) / 5) * 5);
  if (vote === 'REST') return { ok: false, text: 'Einheit streichen oder durch 20–30 min Mobility ersetzen.' };
  if (vote === 'EASY')
    return { ok: false, text: `Einheit auf etwa ${shorter} ESS in Z1–Z2 kürzen oder mit einem lockeren Tag tauschen.` };
  return { ok: false, text: 'Intensität rausnehmen: ähnliche Dauer in Z2, Intervalle auf einen besseren Tag schieben.' };
}

/** Heute schon trainiert? activity_done = x in der Timeline oder Plantag gesperrt (Ist-Wert übernommen). */
function isDoneToday(today: RecoveryDay, plan: PlannedSession | null | 'loading'): boolean {
  return today.activityDone || (!!plan && plan !== 'loading' && plan.locked);
}

/** Absolvierte Einheit als Session (für Anspruch-Einstufung). */
function actualSession(today: RecoveryDay): PlannedSession | null {
  const a = today.actual;
  if (a.load == null) return null;
  return {
    date: today.date,
    load: a.load,
    sport: a.sport,
    zone: a.zone,
    teAe: a.teAe ?? 0,
    teAn: a.teAn ?? 0,
    locked: true,
  };
}

function sessionText(p: PlannedSession): string {
  return p.load <= 0 || p.sport.toLowerCase() === 'off'
    ? 'Ruhetag'
    : `${fmtNum(p.load)} ESS ${p.sport || 'Training'}${p.zone ? ` ${p.zone}` : ''}${
        p.teAe || p.teAn ? ` · TE ${fmtNum(p.teAe, 1)}/${fmtNum(p.teAn, 1)}` : ''
      }`;
}

/** Rückblick: lag die absolvierte Einheit im Rahmen des Votums? */
function doneVerdict(today: RecoveryDay): { ok: boolean; demand: Vote } | null {
  const act = actualSession(today);
  if (!act || !today.vote) return null;
  const demand = plannedDemand(act);
  return { ok: VOTE_ORDER.indexOf(demand) <= VOTE_ORDER.indexOf(today.vote), demand };
}

function PlanCheck({ plan, today }: { plan: PlannedSession | null | 'loading'; today: RecoveryDay }) {
  const done = isDoneToday(today, plan);
  const act = actualSession(today);

  // --- Rückblick: Training ist bereits gelaufen ---
  if (done && act) {
    const dv = doneVerdict(today);
    const lvl: Level = !dv ? 'leer' : dv.ok ? 'gruen' : 'orange';
    const planObj = plan && plan !== 'loading' ? plan : null;
    const planDiffers =
      planObj && !planObj.locked && (Math.abs(planObj.load - act.load) >= 10 || plannedDemand(planObj) !== dv?.demand);
    return (
      <div className={`rounded border px-3 py-2 ${dv ? (dv.ok ? 'border-ampel-gruen/40' : 'border-ampel-orange/50') : 'border-border'} ${LEVEL_CELL[lvl]}`}>
        <div className="flex items-baseline justify-between gap-2">
          <span className="label">Training heute · absolviert</span>
          <span className="text-2xs text-ink-dim">
            {dv ? (dv.demand === 'REST' ? 'Ruhe' : `Anspruch ${VOTE_LABEL[dv.demand]}`) : ''}
          </span>
        </div>
        <div className="mt-0.5 text-sm font-semibold tnum">✓ {sessionText(act)}</div>
        {dv && today.vote && (
          <div className={`mt-1 text-xs ${dv.ok ? 'text-green-300' : 'text-orange-300'}`}>
            {dv.ok
              ? `Im Rahmen des Votums (${VOTE_LABEL[dv.demand]} ≤ ${VOTE_LABEL[today.vote]}).`
              : `Intensiver als das Votum (${VOTE_LABEL[dv.demand]} statt ${VOTE_LABEL[today.vote]}) – Erholung heute priorisieren.`}
          </div>
        )}
        {planDiffers && planObj && (
          <div className="mt-1 text-2xs text-ink-dim tnum">Geplant war: {sessionText(planObj)}.</div>
        )}
      </div>
    );
  }

  if (plan === 'loading') {
    return <div className="text-2xs text-ink-dim">Plan für heute wird geladen…</div>;
  }
  if (!plan) {
    return <div className="text-2xs text-ink-dim">Kein Plan für heute gefunden (Plan Cockpit / PlanApp).</div>;
  }
  const verdict = today.vote ? planVerdict(plan, today.vote) : null;
  const lvl: Level = !verdict ? 'leer' : verdict.ok ? 'gruen' : 'orange';
  return (
    <div className={`rounded border px-3 py-2 ${verdict ? (verdict.ok ? 'border-ampel-gruen/40' : 'border-ampel-orange/50') : 'border-border'} ${LEVEL_CELL[lvl]}`}>
      <div className="flex items-baseline justify-between gap-2">
        <span className="label">Plan heute · offen</span>
        <span className="text-2xs text-ink-dim">
          {plannedDemand(plan) === 'REST' ? 'Ruhe' : `Anspruch ${VOTE_LABEL[plannedDemand(plan)]}`}
        </span>
      </div>
      <div className="mt-0.5 text-sm font-semibold tnum">{sessionText(plan)}</div>
      {verdict && (
        <div className={`mt-1 text-xs ${verdict.ok ? 'text-green-300' : 'text-orange-300'}`}>
          {verdict.ok ? '' : `Votum ${VOTE_LABEL[today.vote!]} → `}
          {verdict.text}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Wochenbilanz & Ausblick
// ---------------------------------------------------------------------------
type PlanDaySim = PlannedSession & { day: string };

const ATL_REC = { a: -8.819849, b: 0.83645661, c: 1.24059231 };
const CTL_REC = { a: -3.95042, b: 0.97283824, c: 0.2336734 };

/** Monotonie wie im Sheet: Ø / Standardabweichung (Population) der letzten 7 Tageslasten. */
function monotony7(loads: number[]): number | null {
  const w = loads.slice(-7);
  if (w.length < 7) return null;
  const m = w.reduce((a, b) => a + b, 0) / w.length;
  const sd = Math.sqrt(w.reduce((a, b) => a + (b - m) ** 2, 0) / w.length);
  return sd > 0 ? m / sd : null;
}

function demandOf(load: number | null, sport: string, zone: string, teAe: number | null, teAn: number | null): Vote {
  return plannedDemand({
    date: '',
    load: load ?? 0,
    sport,
    zone,
    teAe: teAe ?? 0,
    teAn: teAn ?? 0,
    locked: false,
  });
}

/** Mini-Balken: Wochenlasten (rollierende 7-Tage-Blöcke bis heute) der letzten 10 Wochen. */
function WeekLoadBars({ days }: { days: RecoveryDay[] }) {
  const weeks: { end: string; start: string; load: number }[] = [];
  for (let k = 0; k < 10; k++) {
    const endIdx = days.length - 1 - k * 7;
    const startIdx = endIdx - 6;
    if (startIdx < 0) break;
    const blk = days.slice(startIdx, endIdx + 1);
    weeks.unshift({ start: blk[0].date, end: blk[blk.length - 1].date, load: blk.reduce((a, d) => a + (d.essDay ?? 0), 0) });
  }
  if (weeks.length < 2) return null;
  const max = Math.max(...weeks.map((w) => w.load), 1);
  const prev = weeks.slice(0, -1);
  const avgPrev = prev.reduce((a, w) => a + w.load, 0) / prev.length;
  const H = 44;
  const avgY = H - (avgPrev / max) * H;
  const last = weeks[weeks.length - 1];
  const diffPct = avgPrev > 0 ? Math.round(((last.load - avgPrev) / avgPrev) * 100) : null;
  return (
    <div className="mt-3 border-t border-border pt-2">
      <div className="flex items-baseline justify-between gap-2">
        <span className="text-2xs text-ink-dim whitespace-nowrap">Wochenlast · {weeks.length} Wochen</span>
        <span className="text-2xs text-ink-dim tnum whitespace-nowrap">
          Ø {fmtNum(avgPrev)} ESS
          {diffPct != null && (
            <span className={`ml-1 ${Math.abs(diffPct) <= 15 ? 'text-ink-muted' : diffPct > 0 ? 'text-orange-300' : 'text-accent'}`}>
              · jetzt {diffPct > 0 ? '+' : ''}{diffPct} %
            </span>
          )}
        </span>
      </div>
      <div className="relative mt-1.5" style={{ height: H }}>
        <div className="absolute inset-0 flex items-end gap-1">
          {weeks.map((w, i) => {
            const isLast = i === weeks.length - 1;
            return (
              <div
                key={w.end}
                className={`flex-1 rounded-sm ${isLast ? 'wk-bar-last bg-accent' : 'wk-bar bg-accent/35'}`}
                style={{ height: `${Math.max(2, (w.load / max) * 100)}%` }}
                title={`${fmtDateShort(w.start).slice(0, 6)}–${fmtDateShort(w.end).slice(0, 6)}: ${fmtNum(w.load)} ESS`}
              />
            );
          })}
        </div>
        <div className="absolute left-0 right-0 border-t border-dashed border-ink-dim/60 pointer-events-none" style={{ top: avgY }} />
      </div>
      <div className="mt-1 flex justify-between text-2xs text-ink-dim tnum">
        <span>{fmtDateShort(weeks[0].end).slice(0, 6)}</span>
        <span>heute</span>
      </div>
    </div>
  );
}

function WeekSummary({ days }: { days: RecoveryDay[] }) {
  const w = days.slice(-7);
  const avg = (xs: (number | null)[]) => {
    const v = xs.filter((x): x is number => x != null);
    return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
  };
  const readinessAvg = avg(w.map((d) => d.readiness));
  const sleepVals = w.map((d) => d.sleepH).filter((x): x is number => x != null);
  const sleepBalance = sleepVals.length ? sleepVals.reduce((a, b) => a + (b - SLEEP_GOAL_H), 0) : null;
  const weekLoad = w.reduce((a, d) => a + (d.essDay ?? 0), 0);
  const votes = VOTE_ORDER.slice()
    .reverse()
    .map((v) => ({ v, n: w.filter((d) => d.vote === v).length }))
    .filter((x) => x.n > 0);

  // Einhaltung: absolvierte Einheit (IST) vs. Votum des Tages – heute nur, wenn schon absolviert
  const checked = w.filter((d) => d.vote && d.essDay != null && (d.activityDone || d.essDay === 0));
  const misses = checked.filter((d) => {
    const dem = demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
    return VOTE_ORDER.indexOf(dem) > VOTE_ORDER.indexOf(d.vote!);
  });
  const ok = checked.length - misses.length;

  return (
    <div className="panel-raised p-3 min-w-0">
      <div className="flex items-baseline justify-between gap-2">
        <div className="label">Wochenbilanz · 7 Tage</div>
        <span className="text-2xs text-ink-dim tnum">
          {w.length ? `${fmtDateShort(w[0].date).slice(0, 6)}–${fmtDateShort(w[w.length - 1].date).slice(0, 6)}` : ''}
        </span>
      </div>
      <dl className="mt-2 grid grid-cols-2 gap-x-3 gap-y-2 text-sm">
        <div>
          <dt className="text-2xs text-ink-dim">Ø Readiness</dt>
          <dd className={`tnum font-semibold ${LEVEL_TEXT[levelReadiness(readinessAvg)]}`}>{readinessAvg != null ? fmtNum(readinessAvg) : '—'}</dd>
        </div>
        <div>
          <dt className="text-2xs text-ink-dim">Schlaf ggü. Ziel</dt>
          <dd className={`tnum font-semibold ${sleepBalance == null ? '' : sleepBalance >= 0 ? 'text-green-300' : sleepBalance > -2 ? 'text-yellow-300' : 'text-orange-300'}`}>
            {sleepBalance != null ? `${sleepBalance >= 0 ? '+' : ''}${fmtHm(sleepBalance)} h` : '—'}
          </dd>
        </div>
        <div>
          <dt className="text-2xs text-ink-dim">Wochenlast</dt>
          <dd className="tnum font-semibold">{fmtNum(weekLoad)} ESS</dd>
        </div>
        <div>
          <dt className="text-2xs text-ink-dim">Voten</dt>
          <dd className="flex flex-wrap gap-1 mt-0.5">
            {votes.map(({ v, n }) => (
              <span key={v} className={`chip border text-2xs ${LEVEL_CELL[VOTE_LEVEL[v]]} ${LEVEL_TEXT[VOTE_LEVEL[v]]} border-border`}>
                {n}× {VOTE_LABEL[v]}
              </span>
            ))}
          </dd>
        </div>
      </dl>
      <div className="mt-3 border-t border-border pt-2 text-xs">
        <div className="flex items-baseline justify-between gap-2">
          <span className="text-ink-muted">Einheiten passend zum Votum</span>
          <span className={`tnum font-semibold ${misses.length === 0 ? 'text-green-300' : 'text-orange-300'}`}>
            {checked.length ? `${ok} / ${checked.length}` : '—'}
          </span>
        </div>
        {misses.length > 0 && (
          <ul className="mt-1 space-y-0.5 text-2xs text-ink-dim">
            {misses.map((d) => (
              <li key={d.date} className="tnum">
                {fmtDateShort(d.date).slice(0, 6)}: {fmtNum(d.essDay)} ESS {d.actual.sport} {d.actual.zone} (
                {VOTE_LABEL[demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn)]}) bei Votum{' '}
                {VOTE_LABEL[d.vote!]}
              </li>
            ))}
          </ul>
        )}
      </div>
      <WeekLoadBars days={days} />
    </div>
  );
}

function Outlook({ days, plan }: { days: RecoveryDay[]; plan: PlanDaySim[] | null | 'loading' }) {
  const todayIso = localIsoDate();
  const today = days.find((d) => d.date === todayIso) || null;

  const rows = useMemo(() => {
    if (!Array.isArray(plan) || !today) return [];
    const future = plan.filter((p) => p.date > todayIso).slice(0, 3);
    if (!future.length || today.atlPost == null || today.ctlPost == null) return [];
    // Ausgangslage = Stand Ende heute (Timeline-Anker), danach Rekursion wie im Plan Cockpit
    let atl = today.atlPost;
    let ctl = today.ctlPost;
    let preAcwr = today.acwrPost ?? (ctl > 0 ? atl / ctl : null);
    let preMono = today.monotonyPost;
    const loads = days.filter((d) => d.date <= todayIso).map((d) => d.essDay ?? 0);
    return future.map((p) => {
      const demand = plannedDemand(p);
      const warnHigh = (preAcwr != null && preAcwr > ACWR_HIGH) || (preMono != null && preMono > MONOTONY_HIGH);
      const warn = (preAcwr != null && preAcwr > ACWR_WARN) || (preMono != null && preMono > MONOTONY_WARN);
      const cap: Vote | null = warnHigh ? 'EASY' : warn ? 'TRAIN' : null;
      const conflict = cap != null && VOTE_ORDER.indexOf(demand) > VOTE_ORDER.indexOf(cap);
      const row = { p, demand, preAcwr, preMono, cap, conflict };
      atl = ATL_REC.a + ATL_REC.b * atl + ATL_REC.c * p.load;
      ctl = CTL_REC.a + CTL_REC.b * ctl + CTL_REC.c * p.load;
      loads.push(p.load);
      preAcwr = ctl > 0 ? atl / ctl : null;
      preMono = monotony7(loads);
      return row;
    });
  }, [plan, days, today, todayIso]);

  return (
    <div className="panel-raised p-3 min-w-0">
      <div className="flex items-baseline justify-between gap-2">
        <div className="label">Ausblick · 3 Tage</div>
        <span className="text-2xs text-ink-dim">Belastung = Prognose Stand Vortag</span>
      </div>
      {plan === 'loading' ? (
        <p className="mt-3 text-xs text-ink-dim">Plan wird geladen…</p>
      ) : !rows.length ? (
        <p className="mt-3 text-xs text-ink-dim">Keine Plandaten für die nächsten Tage.</p>
      ) : (
        <ul className="mt-2 divide-y divide-border">
          {rows.map(({ p, demand, preAcwr, preMono, cap, conflict }) => {
            const lvl: Level = conflict ? 'orange' : cap ? 'gelb' : 'gruen';
            return (
              <li key={p.date} className="py-2 first:pt-1">
                <div className="flex items-baseline justify-between gap-2">
                  <span className="text-sm font-semibold tnum">
                    {p.day} {fmtDateShort(p.date).slice(0, 6)}
                  </span>
                  <span className={`chip border text-2xs ${LEVEL_CELL[VOTE_LEVEL[demand]]} ${LEVEL_TEXT[VOTE_LEVEL[demand]]} border-border`}>
                    {demand === 'REST' ? 'Ruhe' : VOTE_LABEL[demand]}
                  </span>
                </div>
                <div className="text-xs text-ink-muted tnum">
                  {p.load <= 0 || p.sport.toLowerCase() === 'off'
                    ? 'Ruhetag'
                    : `${fmtNum(p.load)} ESS ${p.sport}${p.zone ? ` ${p.zone}` : ''}`}
                </div>
                <div className="mt-0.5 flex items-center gap-1.5 text-2xs tnum">
                  <Dot level={lvl} />
                  <span className="text-ink-dim">
                    ACWR {preAcwr != null ? fmtNum(preAcwr, 2) : '—'} · Monotonie {preMono != null ? fmtNum(preMono, 2) : '—'}
                  </span>
                </div>
                {conflict && (
                  <div className="mt-0.5 text-2xs text-orange-300">
                    {VOTE_LABEL[demand]} geplant, Belastung erlaubt max. {VOTE_LABEL[cap!]} → Einheit entschärfen oder verschieben.
                  </div>
                )}
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Formatierung
// ---------------------------------------------------------------------------
function fmtHm(h: number | null) {
  if (h == null || !Number.isFinite(h)) return '—';
  const sign = h < 0 ? '−' : '';
  const abs = Math.abs(h);
  let hh = Math.floor(abs);
  let mm = Math.round((abs - hh) * 60);
  if (mm === 60) {
    hh += 1;
    mm = 0;
  }
  return `${sign}${hh}:${String(mm).padStart(2, '0')}`;
}

function fmtDelta(v: number | null, digits = 0) {
  if (v == null || !Number.isFinite(v)) return '—';
  const r = Number(v.toFixed(digits));
  if (r === 0) return '±0';
  return `${r > 0 ? '+' : '−'}${fmtNum(Math.abs(r), digits)}`;
}

function Dot({ level }: { level: Level }) {
  return <span className={`inline-block w-2 h-2 rounded-full shrink-0 ${LEVEL_DOT[level]}`} aria-hidden />;
}

function VoteChip({ vote, large = false }: { vote: Vote | null; large?: boolean }) {
  if (!vote) return <span className="text-ink-dim">—</span>;
  const lvl = VOTE_LEVEL[vote];
  const border =
    lvl === 'gruen' ? 'border-ampel-gruen/40' : lvl === 'gelb' ? 'border-ampel-gelb/40' : 'border-ampel-rot/40';
  return (
    <span
      className={`chip border ${border} ${LEVEL_CELL[lvl]} ${LEVEL_TEXT[lvl]} font-semibold ${large ? 'text-sm px-3 py-1' : ''}`}
    >
      {VOTE_LABEL[vote]}
    </span>
  );
}

// ---------------------------------------------------------------------------
// Heute-Karte
// ---------------------------------------------------------------------------
type WellbeingProps = {
  todayIso: string;
  yesterdayIso: string;
  todayValue: number | null;
  yesterdayValue: number | null;
  onSave: (date: string, value: number) => Promise<void>;
};

function WellbeingInput({ todayIso, yesterdayIso, todayValue, yesterdayValue, onSave }: WellbeingProps) {
  const [target, setTarget] = useState<'today' | 'yesterday'>('today');
  const [saving, setSaving] = useState<number | null>(null);
  const date = target === 'today' ? todayIso : yesterdayIso;
  const current = target === 'today' ? todayValue : yesterdayValue;
  return (
    <div className="border-t border-border pt-3">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="label">Befinden · Wie trainierbar fühle ich mich?</div>
        <div className="inline-flex bg-bg rounded border border-border p-0.5">
          {(['today', 'yesterday'] as const).map((t) => (
            <button
              key={t}
              onClick={() => setTarget(t)}
              className={`px-2 py-0.5 text-2xs rounded ${target === t ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
            >
              {t === 'today' ? 'Heute' : 'Gestern'}
            </button>
          ))}
        </div>
      </div>
      <div className="mt-2 grid grid-cols-5 gap-1.5" role="radiogroup" aria-label={`Befinden ${target === 'today' ? 'heute' : 'gestern'}`}>
        {[1, 2, 3, 4, 5].map((v) => {
          const lvl = levelBefinden(v);
          const active = current === v;
          return (
            <button
              key={v}
              role="radio"
              aria-checked={active}
              disabled={saving != null}
              onClick={async () => {
                setSaving(v);
                try {
                  await onSave(date, v);
                } finally {
                  setSaving(null);
                }
              }}
              className={`btn py-1.5 text-sm tnum font-semibold ${active ? `befinden-on ${LEVEL_CELL[lvl]} ${LEVEL_TEXT[lvl]} ring-1 ring-current` : ''}`}
              data-lvl={active ? lvl : undefined}
              data-testid={`befinden-${v}`}
            >
              {saving === v ? '…' : v}
            </button>
          );
        })}
      </div>
      <div className="mt-1 flex justify-between text-2xs text-ink-dim">
        <span>1 = gar nicht</span>
        <span>{current != null ? `gespeichert: ${current}/5` : 'noch nicht erfasst'}</span>
        <span>5 = voll</span>
      </div>
    </div>
  );
}

function TodayCard({
  today,
  yesterday,
  wellbeing,
  plan,
}: {
  today: RecoveryDay | null;
  yesterday: RecoveryDay | null;
  wellbeing: WellbeingProps;
  plan: PlannedSession | null | 'loading';
}) {
  if (!today || !today.hasMorningData) {
    return (
      <div className="panel-raised p-3 sm:p-4 h-full min-w-0 space-y-3">
        <div className="label">Heute · Recovery Brief</div>
        <p className="text-sm text-ink-muted">
          Für heute liegen noch keine Morgenwerte vor. Sobald <span className="font-mono">garmin-morgens</span> gelaufen ist,
          erscheint hier das Tagesvotum.
        </p>
        <WellbeingInput {...wellbeing} />
      </div>
    );
  }
  const vote = today.vote;
  // Nach dem Training: Empfehlung auf Erholung/Morgen umstellen
  const dv = isDoneToday(today, plan) ? doneVerdict(today) : null;
  const doneAdvice = dv
    ? dv.ok
      ? 'Training für heute erledigt. Rest des Tages: Essen, Trinken, Schlaf ab 7:30 h – morgen früh neu bewerten.'
      : 'Einheit war intensiver als empfohlen. Heute nichts mehr nachlegen, früh schlafen; morgen eher locker einplanen.'
    : null;
  const lvl = vote ? VOTE_LEVEL[vote] : 'leer';
  const readinessDiff = today.readiness != null && yesterday?.readiness != null ? today.readiness - yesterday.readiness : null;
  const sleepGap = today.sleepH != null ? today.sleepH - SLEEP_GOAL_H : null;

  const rows: Array<{ label: string; value: string; hint: string; level: Level }> = [
    {
      label: 'Readiness',
      value: today.readiness != null ? `${fmtNum(today.readiness)} / 100` : '—',
      hint: readinessDiff != null ? `${fmtDelta(readinessDiff)} ggü. gestern` : '',
      level: today.levels.readiness,
    },
    {
      label: 'RHR',
      value: today.rhrDelta != null ? `${fmtDelta(today.rhrDelta, 1)} bpm` : today.rhr != null ? `${fmtNum(today.rhr)} bpm` : '—',
      hint:
        today.rhrBaseline != null
          ? `${fmtNum(today.rhr)} bpm vs. Ø ${fmtNum(today.rhrBaseline, 1)} (28 T)`
          : 'Baseline noch nicht verfügbar',
      level: today.levels.rhr,
    },
    {
      label: 'Schlaf',
      value: today.sleepH != null ? `${fmtHm(today.sleepH)} h` : '—',
      hint:
        sleepGap != null
          ? `${sleepGap >= 0 ? '+' : ''}${fmtHm(sleepGap)} h zum Ziel ${fmtHm(SLEEP_GOAL_H)}${
              today.sleepScore != null ? ` · Score ${fmtNum(today.sleepScore)}` : ''
            }`
          : '',
      level: today.levels.sleep,
    },
    {
      label: 'HRV',
      value: today.hrv != null ? `${fmtNum(today.hrv)} ms` : '—',
      hint: today.hrvLow != null ? `Normalband ${fmtNum(today.hrvLow)}–${fmtNum(today.hrvHigh)}` : '',
      level: today.levels.hrv,
    },
    {
      label: 'Belastung',
      value: today.acwr != null ? `ACWR ${fmtNum(today.acwr, 2)}` : '—',
      hint: today.monotony != null ? `Monotonie ${fmtNum(today.monotony, 2)} · Stand gestern` : '',
      level:
        (today.monotony != null && today.monotony > MONOTONY_HIGH) || (today.acwr != null && today.acwr > ACWR_HIGH)
          ? 'orange'
          : (today.monotony != null && today.monotony > MONOTONY_WARN) || (today.acwr != null && today.acwr > ACWR_WARN)
          ? 'gelb'
          : today.acwr == null && today.monotony == null
          ? 'leer'
          : 'gruen',
    },
    {
      label: 'Befinden',
      value: today.befinden != null ? `${today.befinden} / 5` : '—',
      hint: today.befinden == null ? 'unten eintragen' : '',
      level: today.levels.befinden,
    },
  ];

  return (
    <div className={`panel-raised p-3 sm:p-4 h-full min-w-0 border ${lvl === 'rot' ? 'border-ampel-rot/40' : lvl === 'gelb' ? 'border-ampel-gelb/40' : 'border-ampel-gruen/40'}`}>
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="label">Heute · Recovery Brief</div>
          <div className={`mt-1 text-lg font-semibold tracking-tight ${LEVEL_TEXT[lvl]}`}>
            {vote ? VOTE_STATUS_TEXT[vote] : '—'}
          </div>
        </div>
        <VoteChip vote={vote} large />
      </div>

      <dl className="mt-4 space-y-2 text-sm">
        {rows.map((r) => (
          <div key={r.label} className="grid grid-cols-[84px_minmax(0,1fr)] gap-2 items-baseline">
            <dt className="flex items-center gap-2 text-ink-muted">
              <Dot level={r.level} />
              {r.label}
            </dt>
            <dd>
              <span className="tnum font-semibold">{r.value}</span>
              {r.hint && <span className="block sm:inline sm:ml-2 text-2xs text-ink-dim tnum">{r.hint}</span>}
            </dd>
          </div>
        ))}
      </dl>

      <div className="mt-4">
        <PlanCheck plan={plan} today={today} />
      </div>

      <div className="mt-4">
        <WellbeingInput {...wellbeing} />
      </div>

      <div className="mt-4 border-t border-border pt-3 space-y-2 text-sm">
        <div>
          <div className="label mb-1">Begründung</div>
          <ul className="space-y-0.5 text-ink-muted text-xs">
            {today.reasons.map((r) => (
              <li key={r}>· {r}</li>
            ))}
          </ul>
        </div>
        {vote && (
          <div>
            <div className="label mb-1">Empfehlung</div>
            <p className="text-sm">{doneAdvice ?? VOTE_ADVICE[vote]}</p>
          </div>
        )}
        <p className="text-2xs text-ink-dim">
          Regelbasierte Empfehlung aus Garmin-Werten und deinem Befinden, keine Diagnose.
        </p>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// 7-Tage-Strip
// ---------------------------------------------------------------------------
function RecoveryStrip({ days, todayIso }: { days: RecoveryDay[]; todayIso: string }) {
  const rows: Array<{ label: string; render: (d: RecoveryDay) => { text: string; level: Level } }> = [
    { label: 'Readiness', render: (d) => ({ text: d.readiness != null ? fmtNum(d.readiness) : '—', level: d.levels.readiness }) },
    {
      label: 'RHR vs. Baseline',
      render: (d) => ({ text: d.rhrDelta != null ? `${fmtDelta(d.rhrDelta)} bpm` : '—', level: d.levels.rhr }),
    },
    { label: 'Schlaf', render: (d) => ({ text: fmtHm(d.sleepH), level: d.levels.sleep }) },
    { label: 'HRV', render: (d) => ({ text: d.hrv != null ? fmtNum(d.hrv) : '—', level: d.levels.hrv }) },
    {
      label: 'Monotonie (Vortag)',
      render: (d) => ({
        text: d.monotony != null ? fmtNum(d.monotony, 2) : '—',
        level: d.monotony == null ? 'leer' : d.monotony > MONOTONY_HIGH ? 'orange' : d.monotony > MONOTONY_WARN ? 'gelb' : 'gruen',
      }),
    },
    { label: 'Befinden', render: (d) => ({ text: d.befinden != null ? `${d.befinden}/5` : '—', level: d.levels.befinden }) },
  ];

  const scrollRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    // Auf schmalen Bildschirmen mit "Heute" (rechts) beginnen
    const el = scrollRef.current;
    if (el) el.scrollLeft = el.scrollWidth;
  }, [days.length]);

  return (
    <div ref={scrollRef} className="panel-raised overflow-x-auto">
      <table className="w-full text-xs min-w-[640px]">
        <thead>
          <tr className="border-b border-border">
            <th className="label text-left px-3 py-2 sticky left-0 z-10 bg-bg-raised">Metrik</th>
            {days.map((d) => (
              <th key={d.date} className="label text-center px-2 py-2 tnum">
                {d.date === todayIso ? 'Heute' : fmtDateShort(d.date).slice(0, 6)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-border">
          {rows.map((row) => (
            <tr key={row.label}>
              <td className="px-3 py-2 text-ink-muted whitespace-nowrap sticky left-0 z-10 bg-bg-raised">{row.label}</td>
              {days.map((d) => {
                const { text, level } = row.render(d);
                return (
                  <td key={d.date} className={`px-2 py-2 text-center whitespace-nowrap ${LEVEL_CELL[level]}`}>
                    <span className="inline-flex items-center gap-1.5 tnum font-medium">
                      <Dot level={level} />
                      {text}
                    </span>
                  </td>
                );
              })}
            </tr>
          ))}
          <tr>
            <td className="px-3 py-2 text-ink-muted whitespace-nowrap font-semibold sticky left-0 z-10 bg-bg-raised">Tagesvotum</td>
            {days.map((d) => (
              <td key={d.date} className="px-2 py-2 text-center">
                <VoteChip vote={d.vote} />
              </td>
            ))}
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vier-Spuren-Chart (synchronisiert über syncId)
// ---------------------------------------------------------------------------
const GRID = '#1f2731';

function Lane({
  title,
  hint,
  height = 120,
  data,
  children,
  yDomain,
  showX = false,
}: {
  title: string;
  hint?: string;
  height?: number;
  data: any[];
  children: ReactNode;
  yDomain?: [number | string, number | string];
  showX?: boolean;
}) {
  return (
    <div>
      <div className="flex flex-wrap items-baseline justify-between gap-x-3 px-1">
        <span className="text-xs font-semibold whitespace-nowrap">{title}</span>
        {hint && <span className="text-2xs text-ink-dim text-right">{hint}</span>}
      </div>
      <div style={{ height }}>
        <ResponsiveContainer width="100%" height="100%">
          <ComposedChart data={data} syncId="recovery-lanes" margin={{ top: 6, right: 12, left: 0, bottom: 0 }}>
            <CartesianGrid stroke={GRID} vertical={false} />
            <XAxis
              dataKey="date"
              tickFormatter={(v) => fmtDateShort(String(v)).slice(0, 6)}
              tickLine={false}
              axisLine={{ stroke: GRID }}
              minTickGap={16}
              hide={!showX}
            />
            <YAxis tickLine={false} axisLine={{ stroke: GRID }} width={36} domain={yDomain || ['auto', 'auto']} />
            <Tooltip
              labelFormatter={(v) => fmtDateShort(String(v))}
              formatter={(v: any, name: string) => [v == null ? '—' : typeof v === 'number' ? v.toFixed(1).replace(/\.0$/, '') : v, name]}
            />
            {/* Unsichtbarer Balken erzwingt Band-Skala in allen Spuren → Tage liegen exakt übereinander */}
            <Bar dataKey="_align" fill="transparent" isAnimationActive={false} legendType="none" tooltipType="none" />
            {children}
          </ComposedChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function FourLanes({ days }: { days: RecoveryDay[] }) {
  const data = days.map((d) => ({
    date: d.date,
    Readiness: d.readiness,
    'ΔRHR (bpm)': d.rhrDelta != null ? Number(d.rhrDelta.toFixed(1)) : null,
    'Schlaf (h)': d.sleepH,
    'Schlaf Ø7': d.sleepAvg7 != null ? Number(d.sleepAvg7.toFixed(2)) : null,
    HRV: d.hrv,
    'HRV unten': d.hrvLow,
    Befinden: d.befinden,
  }));

  return (
    <div className="space-y-3">
      <Lane title="Readiness" hint="≥75 grün · 55–74 gelb · <55 orange/rot" data={data} yDomain={[0, 100]}>
        <ReferenceLine y={75} stroke="#22c55e" strokeDasharray="3 3" />
        <ReferenceLine y={55} stroke="#eab308" strokeDasharray="3 3" />
        <Area type="monotone" dataKey="Readiness" stroke="#a855f7" fill="#a855f7" fillOpacity={0.18} strokeWidth={1.8} connectNulls isAnimationActive={false} />
      </Lane>
      <Lane title="RHR-Abweichung" hint="bpm relativ zur 28-Tage-Baseline · 0 = neutral" data={data}>
        <ReferenceLine y={0} stroke="#5b6678" />
        <ReferenceLine y={5} stroke="#ef4444" strokeDasharray="3 3" />
        <Bar dataKey="ΔRHR (bpm)" fill="#ef4444" fillOpacity={0.55} maxBarSize={14} radius={[2, 2, 0, 0]} isAnimationActive={false} />
      </Lane>
      <Lane title="Schlaf" hint={`Balken = Stunden · Linie = Ø 7 Tage · Ziel ${fmtHm(SLEEP_GOAL_H)} h`} data={data} yDomain={[0, 'auto']}>
        <ReferenceLine y={SLEEP_GOAL_H} stroke="#22c55e" strokeDasharray="3 3" />
        <Bar dataKey="Schlaf (h)" fill="#7dd3fc" fillOpacity={0.7} maxBarSize={14} radius={[2, 2, 0, 0]} isAnimationActive={false} />
        <Line type="monotone" dataKey="Schlaf Ø7" stroke="#0ea5e9" strokeWidth={1.6} dot={false} connectNulls isAnimationActive={false} />
      </Lane>
      <Lane title="HRV" hint="Nachtwert · gestrichelt = untere Normalgrenze" data={data}>
        <Line type="monotone" dataKey="HRV" stroke="#22c55e" strokeWidth={1.8} dot={{ r: 2 }} connectNulls isAnimationActive={false} />
        <Line type="stepAfter" dataKey="HRV unten" stroke="#f97316" strokeDasharray="4 4" strokeWidth={1.2} dot={false} connectNulls isAnimationActive={false} />
      </Lane>
      <Lane title="Befinden" hint="1–5 · eigene Eingabe · gestrichelt = 3" data={data} yDomain={[1, 5]} showX height={130}>
        <ReferenceLine y={3} stroke="#eab308" strokeDasharray="3 3" />
        <Line type="linear" dataKey="Befinden" stroke="#f472b6" strokeWidth={1.6} dot={{ r: 3.5, fill: '#f472b6' }} connectNulls isAnimationActive={false} />
      </Lane>

      {/* Status-Spur */}
      <div>
        <div className="px-1 text-xs font-semibold mb-1">Tagesvotum</div>
        <div className="grid gap-0.5 pl-9 pr-3" style={{ gridTemplateColumns: `repeat(${days.length}, minmax(0, 1fr))` }}>
          {days.map((d) => {
            const lvl = d.vote ? VOTE_LEVEL[d.vote] : 'leer';
            return (
              <div
                key={d.date}
                title={`${fmtDateShort(d.date)} · ${d.vote ? VOTE_LABEL[d.vote] : 'keine Daten'}`}
                className={`h-5 rounded-sm ${lvl === 'leer' ? 'bg-bg-subtle' : LEVEL_DOT[lvl]} opacity-70`}
              />
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Container
// ---------------------------------------------------------------------------
type ToastFn = (kind: 'ok' | 'err' | 'info', text: string) => void;

export function RecoveryPanel({
  token,
  toast,
  proxyAuthenticated,
}: {
  token: string;
  toast: ToastFn;
  proxyAuthenticated: boolean;
}) {
  const [raw, setRaw] = useState<any[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(true);
  const [span, setSpan] = useState<14 | 28>(14);
  const [planToday, setPlanToday] = useState<PlannedSession | null | 'loading'>('loading');
  const [planDays, setPlanDays] = useState<PlanDaySim[] | null | 'loading'>('loading');

  async function loadPlan() {
    setPlanToday('loading');
    setPlanDays('loading');
    try {
      const sim: any = await fetchPlanSimulationWithTimeout(45000);
      const iso = localIsoDate();
      setPlanDays(
        sim?.ok
          ? (sim.days || []).map((x: any) => ({
              date: String(x.date).slice(0, 10),
              day: String(x.day || ''),
              load: toNum(x.load) ?? 0,
              sport: String(x.sport || ''),
              zone: String(x.zone || ''),
              teAe: toNum(x.te_ae) ?? 0,
              teAn: toNum(x.te_an) ?? 0,
              locked: !!x.locked,
            }))
          : null,
      );
      const d = sim?.ok ? (sim.days || []).find((x: any) => String(x.date).slice(0, 10) === iso) : null;
      setPlanToday(
        d
          ? {
              date: iso,
              load: toNum(d.load) ?? 0,
              sport: String(d.sport || ''),
              zone: String(d.zone || ''),
              teAe: toNum(d.te_ae) ?? 0,
              teAn: toNum(d.te_an) ?? 0,
              locked: !!d.locked,
            }
          : null,
      );
    } catch {
      setPlanToday(null);
      setPlanDays(null);
    }
  }

  async function load() {
    setBusy(true);
    setErr(null);
    loadPlan();
    try {
      const res = await fetchChartData('90d', RECOVERY_FETCH_METRICS);
      setRaw(res.data as any[]);
    } catch (e) {
      setErr(e instanceof ApiError ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  const todayIso = localIsoDate();
  const yDate = new Date();
  yDate.setDate(yDate.getDate() - 1);
  const yesterdayIso = localIsoDate(yDate);

  async function handleSaveWellbeing(date: string, value: number) {
    try {
      const res = proxyAuthenticated ? await proxySaveWellbeing(date, value) : await saveWellbeing(token.trim(), date, value);
      if (!res.ok) throw new ApiError(res.error || 'Speichern fehlgeschlagen');
      setRaw((prev) => {
        const rows = [...(prev || [])];
        const i = rows.findIndex((r) => String(r.date).slice(0, 10) === date);
        if (i >= 0) rows[i] = { ...rows[i], befinden_1_5: value };
        else rows.push({ date, befinden_1_5: value });
        return rows;
      });
      toast('ok', `Befinden ${value}/5 für ${fmtDateShort(date)} gespeichert.`);
    } catch (e) {
      toast('err', `Befinden nicht gespeichert: ${e instanceof ApiError ? e.message : String(e)}`);
    }
  }
  const days = useMemo(() => (raw ? buildDays(raw).filter((d) => d.date <= todayIso) : []), [raw, todayIso]);
  const today = days.length && days[days.length - 1].date === todayIso ? days[days.length - 1] : null;
  const yesterday = today ? days[days.length - 2] || null : days[days.length - 1] || null;
  const stripDays = days.slice(-7);
  const laneDays = days.slice(-span);

  return (
    <Panel
      title="Recovery · Tagesentscheid"
      right={
        <div className="flex items-center gap-2">
          <div className="inline-flex bg-bg rounded border border-border p-0.5">
            {[14, 28].map((n) => (
              <button
                key={n}
                onClick={() => setSpan(n as 14 | 28)}
                className={`px-2.5 py-1 text-xs rounded tnum whitespace-nowrap ${span === n ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
              >
                {n} T
              </button>
            ))}
          </div>
          <button onClick={load} className="btn btn-ghost text-xs px-2 py-1" disabled={busy}>
            {busy ? 'Lade…' : 'Aktualisieren'}
          </button>
        </div>
      }
    >
      {err ? (
        <ErrorBox message={err} onRetry={load} />
      ) : busy && !raw ? (
        <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(300px,1fr)_minmax(0,2fr)]">
          <Skeleton className="h-72" />
          <Skeleton className="h-72" />
        </div>
      ) : (
        <div className="space-y-4">
          <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-[minmax(300px,1fr)_minmax(0,2fr)]">
            <TodayCard
              today={today}
              yesterday={yesterday}
              plan={planToday}
              wellbeing={{
                todayIso,
                yesterdayIso,
                todayValue: days.find((d) => d.date === todayIso)?.befinden ?? null,
                yesterdayValue: days.find((d) => d.date === yesterdayIso)?.befinden ?? null,
                onSave: handleSaveWellbeing,
              }}
            />
            <div className="space-y-2 min-w-0">
              <div className="label">Letzte 7 Tage</div>
              <RecoveryStrip days={stripDays} todayIso={todayIso} />
              <p className="text-2xs text-ink-dim">
                RHR-Baseline = Ø der vorherigen {RHR_BASELINE_DAYS} Tage · Schlafziel {fmtHm(SLEEP_GOAL_H)} h · Werte ≥ 5 bpm
                unter Baseline werden grau (atypisch) statt grün markiert. Belastung (ACWR, Monotonie) jeweils Stand Vortag:
                über {fmtNum(ACWR_WARN, 1)} / {fmtNum(MONOTONY_WARN, 1)} keine Qualität, über {fmtNum(ACWR_HIGH, 1)} / {fmtNum(MONOTONY_HIGH, 1)} max. Easy.
              </p>
              <div className="grid gap-3 sm:grid-cols-2 pt-1">
                <WeekSummary days={days} />
                <Outlook days={days} plan={planDays} />
              </div>
            </div>
          </div>
          <div className="panel-raised p-3">
            <FourLanes days={laneDays} />
          </div>
        </div>
      )}
    </Panel>
  );
}
