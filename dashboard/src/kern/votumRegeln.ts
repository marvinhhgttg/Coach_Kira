/**
 * Coach-Kira-Rechenkern – Regeln für das Tagesvotum (Cockpit-Regel, Entscheidung 20.11.2026).
 * Verschoben aus src/screens/RecoveryPanel.tsx (Etappe 2). Einzige Quelle für Cockpit, Widgets und Server.
 */
import { dedupeByDateKeepLast, toNum, type Range } from '../lib/api';
import { garminDeckel, type GarminSignale } from './garminSignale';
import { fmtNum } from '../lib/format';
import { WINDOWS, rangeFor, type WinKey } from '../lib/windows';
import { assessMonotony, type MonoAssessment } from '../lib/monotony';


// ---------------------------------------------------------------------------
// Konfiguration
// ---------------------------------------------------------------------------
export const SLEEP_GOAL_H = 7.5;
export const RHR_BASELINE_DAYS = 28;
export const RHR_BASELINE_MIN_VALUES = 7;

export type Span = WinKey;
export const SPANS = WINDOWS;
/** Datenfenster je Ansicht: +28 Tage Vorlauf für RHR-Baseline, mindestens 90 Tage (Wochenbalken). */
export function rangeForSpan(n: Span): Range {
  return rangeFor(n, 28, 90);
}

export const RECOVERY_FETCH_METRICS = [
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

export const ACWR_WARN = 1.3;
export const ACWR_HIGH = 1.5;

export type Level = 'blau' | 'gruen' | 'gelb' | 'orange' | 'rot' | 'grau' | 'leer';
export type Vote = 'QUALITY' | 'TRAIN' | 'EASY' | 'REST';

export const VOTE_ORDER: Vote[] = ['REST', 'EASY', 'TRAIN', 'QUALITY'];
export const VOTE_LABEL: Record<Vote, string> = {
  QUALITY: 'Quality',
  TRAIN: 'Train',
  EASY: 'Easy',
  REST: 'Rest',
};
export const VOTE_LEVEL: Record<Vote, Level> = {
  QUALITY: 'blau', // Garmin-Spektrum: Blau über Grün
  TRAIN: 'gruen',
  EASY: 'gelb',
  REST: 'rot',
};
export const VOTE_ADVICE: Record<Vote, string> = {
  QUALITY: 'Qualitätseinheit möglich (Schwelle/Intervalle), wenn sie geplant ist.',
  TRAIN: 'Geplantes Training normal umsetzen, Schwerpunkt Z2.',
  EASY: '45–75 min Z1–Z2, keine Intervalle; morgen neu bewerten.',
  REST: 'Ruhetag oder max. 30–45 min sehr locker / Mobility.',
};
export const VOTE_STATUS_TEXT: Record<Vote, string> = {
  QUALITY: 'BLAU — Qualität möglich',
  TRAIN: 'GRÜN — normal trainieren',
  EASY: 'GELB — locker trainieren',
  REST: 'ROT — Erholung priorisieren',
};

export const LEVEL_DOT: Record<Level, string> = {
  blau: 'bg-ampel-blau',
  gruen: 'bg-ampel-gruen',
  gelb: 'bg-ampel-gelb',
  orange: 'bg-ampel-orange',
  rot: 'bg-ampel-rot',
  grau: 'bg-ampel-grau',
  leer: 'bg-transparent border border-border-strong',
};
export const LEVEL_CELL: Record<Level, string> = {
  blau: 'bg-ampel-blau/10',
  gruen: 'bg-ampel-gruen/10',
  gelb: 'bg-ampel-gelb/10',
  orange: 'bg-ampel-orange/10',
  rot: 'bg-ampel-rot/10',
  grau: 'bg-ampel-grau/10',
  leer: '',
};
export const LEVEL_TEXT: Record<Level, string> = {
  blau: 'text-blue-300',
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
export type RecoveryDay = {
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
  /** Ist-Lasten und Monotonie-Verlauf bis gestern (für die Monotonie-Regel) */
  monoLoads: number[];
  monoHist: (number | null)[];
  monoAssess: MonoAssessment | null;
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
  /** Grund, der das Regel-Votum zuletzt begrenzt hat */
  bind?: VoteGrund | null;
  /** Regel-Votum vor persönlicher Anpassung durch das Erfahrungsmodell (nur heute) */
  ruleVote?: Vote | null;
  /** Hinweise aus dem Erfahrungsmodell (nur heute) */
  erfahrung?: string[];
  /** 27.11.2026: Garmin-Signale (nur heute, live vom Mac) */
  garmin?: GarminSignale | null;
};

export function localIsoDate(d = new Date()) {
  // yyyy-mm-dd in lokaler Zeitzone (Europe/Berlin im Browser)
  return d.toLocaleDateString('sv-SE');
}

export function parseThresholds(v: unknown): [number | null, number | null] {
  if (typeof v !== 'string' || !v.includes(';')) return [null, null];
  const [a, b] = v.split(';').map((x) => Number(String(x).trim().replace(',', '.')));
  return [Number.isFinite(a) ? a : null, Number.isFinite(b) ? b : null];
}

export function mean(values: number[]) {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : null;
}

export function levelReadiness(v: number | null): Level {
  if (v == null) return 'leer';
  if (v >= 75) return 'gruen';
  if (v >= 55) return 'gelb';
  if (v >= 40) return 'orange';
  return 'rot';
}

export function levelRhrDelta(d: number | null): Level {
  if (d == null) return 'leer';
  // Ungewöhnlich niedrig ist nicht automatisch gut (Messartefakt / atypisches Signal)
  if (d <= -5) return 'grau';
  if (d <= 1) return 'gruen';
  if (d < 3) return 'gelb';
  if (d < 5) return 'orange';
  return 'rot';
}

export function levelSleep(h: number | null): Level {
  if (h == null) return 'leer';
  const ratio = h / SLEEP_GOAL_H;
  if (ratio >= 0.95) return 'gruen';
  if (ratio >= 0.85) return 'gelb';
  if (ratio >= 0.8) return 'orange';
  return 'rot';
}

export function levelHrv(v: number | null, low: number | null): Level {
  if (v == null || low == null) return 'leer';
  if (v < low - 2) return 'rot'; // entspricht SG_Flag "H"
  if (v < low) return 'orange';
  return 'gruen';
}

export function levelBefinden(v: number | null): Level {
  if (v == null) return 'leer';
  if (v >= 4) return 'gruen';
  if (v >= 3) return 'gelb';
  if (v >= 2) return 'orange';
  return 'rot';
}

export function capVote(vote: Vote, max: Vote): Vote {
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
 * 9. Garmin (nur heute, kern/garminSignale.ts): Erholungszeit morgens ≥ 60 h → max. Easy · ≥ 36 h → max. Train ·
 *    Trainingszustand Überlastet/Überzogen → max. Easy
 * 8. Belastung: ACWR > 1,3 → max. Train · ACWR > 1,5 → max. Easy
 *    Monotonie (Verlauf, lib/monotony.ts): 2 Tage > 2,0 nach 4–5 fordernden Tagen → Easy · bleibt > 2,5 → Easy + Ruhetag binnen 48 h ·
 *    > 3,0 → Easy + Entlastungstag · > 2,0 mit schlechter Erholung → keine Qualität · 5+ Tage ohne Entlastung → Easy
 */
export type VoteGrund = 'readiness' | 'rot' | 'rhr_schlaf' | 'hrv' | 'warn2' | 'rhr_ready' | 'acwr' | 'mono' | 'befinden' | 'garmin';

export function decideVote(d: Omit<RecoveryDay, 'vote' | 'reasons' | 'bind'>): { vote: Vote | null; reasons: string[]; bind: VoteGrund | null } {
  if (d.readiness == null) return { vote: null, reasons: ['Readiness fehlt (Morgenwerte noch nicht erfasst).'], bind: null };
  const reasons: string[] = [];
  let vote: Vote = d.readiness >= 75 ? 'QUALITY' : d.readiness >= 55 ? 'TRAIN' : d.readiness >= 40 ? 'EASY' : 'REST';
  reasons.push(`Readiness ${fmtNum(d.readiness)} → Basis ${VOTE_LABEL[vote]}.`);
  // bind = Grund, der das Votum zuletzt gesenkt hat (für das Erfahrungsmodell)
  let bind: VoteGrund = 'readiness';
  const senke = (neu: Vote, grund: VoteGrund) => {
    if (VOTE_ORDER.indexOf(neu) < VOTE_ORDER.indexOf(vote)) bind = grund;
    vote = neu;
  };

  const { rhr, sleep, hrv } = d.levels;
  const warn = [rhr, sleep, hrv].filter((l) => l === 'orange' || l === 'rot').length;

  if ([rhr, sleep, hrv].includes('rot')) {
    const before = vote;
    senke(capVote(vote, 'TRAIN'), 'rot');
    if (vote !== before) reasons.push('Ein Einzelsignal ist rot → keine Qualitätseinheit.');
  }
  const sleepRatio = d.sleepH != null ? d.sleepH / SLEEP_GOAL_H : null;
  if (d.rhrDelta != null && d.rhrDelta >= 5 && sleepRatio != null && sleepRatio < 0.8) {
    senke(capVote(vote, 'EASY'), 'rhr_schlaf');
    reasons.push('RHR ≥ +5 bpm und Schlaf < 80 % des Ziels → maximal Easy.');
  }
  if (hrv === 'rot') {
    senke(capVote(vote, 'EASY'), 'hrv');
    reasons.push('HRV unter dem Normalband (HRV-Alarm) → maximal Easy.');
  }
  if (warn >= 2) {
    senke(capVote(vote, 'EASY'), 'warn2');
    reasons.push(`${warn} Warnsignale gleichzeitig → maximal Easy.`);
  }
  if (d.rhrDelta != null && d.rhrDelta >= 5 && d.readiness < 55) {
    senke('REST', 'rhr_ready');
    reasons.push('RHR ≥ +5 bpm bei niedriger Readiness → Rest.');
  }
  // 8a. Belastung über ACWR
  if (d.acwr != null && d.acwr > ACWR_HIGH) {
    const before = vote;
    senke(capVote(vote, 'EASY'), 'acwr');
    if (vote !== before) reasons.push(`ACWR ${fmtNum(d.acwr, 2)} > ${fmtNum(ACWR_HIGH, 1)} → maximal Easy.`);
  } else if (d.acwr != null && d.acwr > ACWR_WARN) {
    const before = vote;
    senke(capVote(vote, 'TRAIN'), 'acwr');
    if (vote !== before) reasons.push(`ACWR ${fmtNum(d.acwr, 2)} > ${fmtNum(ACWR_WARN, 1)} → keine Qualitätseinheit.`);
  }
  // 8b. Monotonie-Regel: Verlaufssignal statt Einzelwert
  const ma = d.monoAssess;
  if (ma && ma.cap) {
    const before = vote;
    senke(capVote(vote, ma.cap), 'mono');
    if (vote !== before || ma.restWithin48h) reasons.push(ma.text);
  } else if (ma && ma.text && ma.stufe && ma.stufe !== 'gut' && ma.stufe !== 'normal') {
    reasons.push(ma.text);
  }
  if (ma && ma.hinweis) reasons.push(ma.hinweis);
  if (d.befinden != null) {
    if (d.befinden <= 1) {
      senke('REST', 'befinden');
      reasons.push('Befinden 1/5 → Rest.');
    } else if (d.befinden <= 2) {
      senke(capVote(vote, 'EASY'), 'befinden');
      reasons.push('Befinden 2/5 → maximal Easy.');
    } else if (d.befinden <= 3) {
      const before = vote;
      senke(capVote(vote, 'TRAIN'), 'befinden');
      if (vote !== before) reasons.push('Befinden 3/5 → keine Qualitätseinheit.');
    }
  }
  // 9. Garmin-Signale (nur heute): Erholungszeit am Morgen und Trainingszustand
  for (const g of garminDeckel(d.garmin)) {
    const before = vote;
    senke(capVote(vote, g.cap), 'garmin');
    if (vote !== before) reasons.push(g.grund);
  }
  return { vote, reasons, bind };
}

export function buildDays(raw: any[]): RecoveryDay[] {
  const rows = dedupeByDateKeepLast(
    raw.filter((r) => r && typeof r.date === 'string').map((r) => ({ ...r, date: String(r.date).slice(0, 10) })),
  ).sort((a: any, b: any) => (a.date < b.date ? -1 : 1));

  return rows.map((r: any, idx: number) => {
    const readiness = toNum(r.Garmin_Training_Readiness);
    const rhr = toNum(r.rhr_bpm);
    const sleepH = toNum(r.sleep_hours);
    const sleepScore = toNum(r.sleep_score_0_100);
    const hrvRaw = toNum(r.hrv_status);
    // Plausibilität: offensichtliche Tippfehler (z. B. 334 statt 34) nicht anzeigen/werten
    const hrv = hrvRaw != null && hrvRaw >= 10 && hrvRaw <= 200 ? hrvRaw : null;
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
      monoLoads: rows.slice(Math.max(0, idx - 28), idx).map((p: any) => toNum(p.load_fb_day) ?? toNum(p.coachE_ESS_day) ?? 0),
      monoHist: rows.slice(Math.max(0, idx - 8), Math.max(0, idx - 1)).map((p: any) => toNum(p.Monotony7)),
      monoAssess: null as MonoAssessment | null,
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
      garmin: (r.__garmin as GarminSignale | undefined) ?? null,
      hasMorningData: readiness != null || rhr != null || sleepH != null,
      levels: {
        readiness: levelReadiness(readiness),
        rhr: levelRhrDelta(rhrDelta),
        sleep: levelSleep(sleepH),
        hrv: levelHrv(hrv, hrvLow),
        befinden: levelBefinden(befinden),
      },
    };
    const recoveryBad =
      [base.levels.rhr, base.levels.sleep, base.levels.hrv].some((l) => l === 'orange' || l === 'rot') ||
      (base.befinden != null && base.befinden <= 2);
    base.monoAssess = assessMonotony({ mono: base.monotony, loads: base.monoLoads, monoHistory: base.monoHist, recoveryBad });
    const { vote, reasons, bind } = decideVote(base);
    return { ...base, vote, reasons, bind };
  });
}
