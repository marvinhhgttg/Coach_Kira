import { useEffect, useMemo, useRef, useState } from 'react';
import {
  ApiError,
  LOAD_METRICS,
  TimeoutError,
  dedupeByDateKeepLast,
  fetchChartData,
  fetchPlan,
  fetchPlanBriefing,
  fetchPlanSimulation,
  fetchPlanSimulationWithTimeout,
  generatePlanBriefing,
  savePlanSimulation,
  toNum,
  type PlanDay,
  type PlanResponse,
} from '../lib/api';
import { fmtDateShort, fmtNum, fmtSigned, fmtTime } from '../lib/format';
import { proxyGeneratePlanBriefing, proxySavePlanSimulation } from '../lib/proxy';
import { ErrorBox, Panel, Skeleton } from '../components/UI';
import {
  Bar,
  CartesianGrid,
  ComposedChart,
  Line,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';

type SimDay = PlanDay & {
  locked?: boolean;
  sim_load: number;
  sport: string;
  phase: string;
  zone: string;
  te_ae: number;
  te_an: number;
  atl: number;
  ctl: number;
  ctl_star: number;
  acwr: number;
  monotony: number;
  kei: number;
  flag: 'OK' | 'WATCH' | 'RISK';
  sg_flags: string;
  hrv_status?: number | null;
  hrv_thresholds?: string | null;
  ai_load?: number | null;
};

type ToastFn = (kind: 'ok' | 'err' | 'info', text: string) => void;
type BriefingRecommendation = {
  date: string;
  day?: string;
  recommended_load_ess: number;
  recommended_zone?: string;
  rationale?: string;
};

interface Props {
  token: string;
  toast: ToastFn;
  proxyAuthenticated: boolean;
}

function normalizeBriefingPayload(
  briefing: string | null | undefined,
  recommendations: BriefingRecommendation[] | null | undefined
): { briefing: string | null; recommendations: BriefingRecommendation[] | null } {
  if (!briefing) return { briefing: null, recommendations: recommendations || null };
  let txt = String(briefing).trim();
  txt = txt.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
  if (txt.startsWith('{')) {
    try {
      const parsed = JSON.parse(txt);
      return {
        briefing: parsed.briefing ? String(parsed.briefing) : txt,
        recommendations: Array.isArray(parsed.recommendations)
          ? parsed.recommendations
          : recommendations || null,
      };
    } catch {
      return { briefing: txt, recommendations: recommendations || null };
    }
  }
  return { briefing: txt, recommendations: recommendations || null };
}

const WEEK_LOAD: Record<string, { load: number; sport: string; zone: string }> = {
  Mo: { load: 130, sport: 'Bike', zone: 'Z2' },
  Di: { load: 130, sport: 'Bike', zone: 'Z2' },
  Mi: { load: 100, sport: 'Run', zone: 'Z2-Z4' },
  Do: { load: 140, sport: 'Bike', zone: 'Z2' },
  Fr: { load: 70, sport: 'Row/HIIT', zone: 'Z2' },
  Sa: { load: 150, sport: 'Run/Bike', zone: 'Z2/Z3' },
  So: { load: 80, sport: 'Hike', zone: 'Z2' },
};

const CURRENT_SIM_PATTERN: Array<{ load: number; sport: string; zone: string; t1: string; t2: string }> = [
  { load: 130, sport: 'Bike', zone: 'Z2', t1: 'Bike ZZ2', t2: 'Pendelfahrt / Grundlage' },
  { load: 0, sport: 'Off', zone: 'Off', t1: 'Off', t2: 'Mobility / Recovery' },
  { load: 185, sport: 'Run', zone: 'Z2', t1: 'Run ZZ2', t2: 'Grundlagenlauf' },
  { load: 220, sport: 'Run', zone: 'Z2', t1: 'Run ZZ2', t2: 'Langer aerober Reiz' },
  { load: 0, sport: 'Off', zone: 'Off', t1: 'Off', t2: 'Recovery' },
  { load: 218, sport: 'Run', zone: 'Z2', t1: 'Run ZZ2', t2: 'Langer Lauf' },
  { load: 186, sport: 'Bike', zone: 'Z2', t1: 'Bike ZZ2', t2: 'Gravel / Race' },
  { load: 193, sport: 'Bike', zone: 'Z2', t1: 'Bike ZZ2', t2: 'Pendelfahrt / Grundlage' },
  { load: 220, sport: 'Run', zone: 'Z2', t1: 'Run ZZ2', t2: 'Belastungsreiz' },
  { load: 0, sport: 'Off', zone: 'Off', t1: 'Off', t2: 'Recovery' },
  { load: 160, sport: 'Bike', zone: 'Z2-Z3', t1: 'Bike Z2-Z3', t2: 'Aerober Aufbau' },
  { load: 130, sport: 'Bike', zone: 'Z2', t1: 'Bike ZZ2', t2: 'Grundlage' },
  { load: 100, sport: 'Run', zone: 'Z2', t1: 'Run ZZ2', t2: 'Locker' },
  { load: 80, sport: 'Hike', zone: 'Z2', t1: 'Hike ZZ2', t2: 'Aktive Erholung' },
];

const DAY_NAMES = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const PLAN_ANCHOR_METRICS = [
  'coachE_ESS_day',
  'coachE_ATL_forecast',
  'coachE_CTL_forecast',
  'coachE_ACWR_forecast',
  'coachE_Smart_Gains',
  'Monotony7',
  'hrv_status',
  'hrv_threshholds',
] as const;

/**
 * SG_Flags — replicates the timeline sheet formula in AO:
 *   =TRIM( IF(hrv_status < (LEFT(hrv_threshholds; ";") - 2); "H "; "") &
 *          IF(Monotony7 > 1,6; "M"; "") )
 * H = HRV alarm (hrv_status more than 2 below the lower threshold band).
 * M = Monotony warning (7-day monotony > 1.6).
 */
function computeSgFlags(input: {
  hrvStatus?: number | null;
  hrvThresholds?: string | null;
  monotony?: number | null;
}): string {
  const parts: string[] = [];
  const { hrvStatus, hrvThresholds, monotony } = input;
  if (
    typeof hrvStatus === 'number' &&
    Number.isFinite(hrvStatus) &&
    typeof hrvThresholds === 'string' &&
    hrvThresholds.includes(';')
  ) {
    const lowerRaw = hrvThresholds.split(';')[0].trim().replace(',', '.');
    const lower = Number(lowerRaw);
    if (Number.isFinite(lower) && hrvStatus < lower - 2) {
      parts.push('H');
    }
  }
  if (typeof monotony === 'number' && Number.isFinite(monotony) && monotony > 1.6) {
    parts.push('M');
  }
  return parts.join(' ');
}

function isoDate(d: Date) {
  return d.toISOString().slice(0, 10);
}

function dayName(iso: string) {
  const d = new Date(`${iso}T12:00:00`);
  return DAY_NAMES[d.getDay()];
}

function addDays(iso: string, n: number) {
  const d = new Date(`${iso}T12:00:00`);
  d.setDate(d.getDate() + n);
  return isoDate(d);
}

function isPastPlan(plan: PlanDay[]) {
  if (!plan.length) return false;
  const last = new Date(`${plan[plan.length - 1].date}T12:00:00`);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return last < today;
}

function normalizeTo14Days(plan: PlanDay[]) {
  const today = isoDate(new Date());
  const shiftToToday = isPastPlan(plan);
  const startDate = shiftToToday ? today : plan[0]?.date || today;
  const out: PlanDay[] = [];
  for (let i = 0; i < 14; i++) {
    const source = plan[i] || plan[i % Math.max(plan.length, 1)] || ({} as PlanDay);
    const date = addDays(startDate, i);
    const day = dayName(date);
    const fallback = WEEK_LOAD[day] || { load: 100, sport: 'Training', zone: 'Z2' };
    const current = CURRENT_SIM_PATTERN[i % CURRENT_SIM_PATTERN.length];
    const selectedLoad = shiftToToday ? current.load : source.recommended_load_ess ?? fallback.load;
    const selectedZone = shiftToToday ? current.zone : source.recommended_zone || fallback.zone;
    out.push({
      ...source,
      index: i + 1,
      date,
      day,
      original_load_ess: shiftToToday ? selectedLoad : source.original_load_ess ?? fallback.load,
      recommended_load_ess: selectedLoad,
      recommended_zone: selectedZone,
      training_1: shiftToToday ? current.t1 : source.training_1 || fallback.sport,
      training_2: shiftToToday ? current.t2 : source.training_2 || 'Alternative Einheit',
      projected_effect: shiftToToday
        ? 'Lokale Simulation analog PlanApp: Slider-Änderungen propagieren ATL, CTL, ACWR und KEI.'
        : source.projected_effect || 'Lokale Simulation auf Basis der aktuellen Lastwerte.',
      weather_recommendation: shiftToToday
        ? 'Wetterhinweise bleiben aus Coach Kira / PlanApp zu übernehmen.'
        : source.weather_recommendation || 'Wetterdaten aus Coach Kira beachten.',
    });
  }
  return out;
}

function phaseForIndex(i: number) {
  if (i < 2) return 'E';
  if (i < 5) return 'A1';
  if (i < 10) return 'A2';
  return 'AT';
}

function estimateTe(load: number, sport: string) {
  if (load <= 0) return { ae: 0, an: 0 };
  const sportBoost = sport.toLowerCase().includes('run') ? 0.1 : sport.toLowerCase().includes('bike') ? 0.05 : 0;
  const ae = Math.min(5, 1.4 + load / 75 + sportBoost);
  const an = load > 180 ? Math.min(1.5, (load - 180) / 60) : 0;
  return { ae, an };
}

function rollingMonotony(loads: number[], idx: number) {
  const start = Math.max(0, idx - 6);
  const window = loads.slice(start, idx + 1);
  if (window.length < 2) return 1;
  const mean = window.reduce((a, b) => a + b, 0) / window.length;
  const variance = window.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / window.length;
  const sd = Math.sqrt(variance);
  if (sd < 1) return 2.5;
  return mean / sd;
}

function rollingMonotonyWithSeed(loadHistorySeed: number[], loads: number[], idx: number) {
  if (!loadHistorySeed.length) return rollingMonotony(loads, idx);
  // loadHistorySeed is expected to include the visible anchor day as its last value.
  // For future rows, append loads from Tag 2 onward and take the last seven values.
  const futureLoads = idx <= 0 ? [] : loads.slice(1, idx + 1);
  const series = loadHistorySeed.concat(futureLoads);
  const window = series.slice(-7);
  if (window.length < 2) return 1;
  const mean = window.reduce((a, b) => a + b, 0) / window.length;
  const variance = window.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / window.length;
  const sd = Math.sqrt(variance);
  if (sd < 1) return 2.5;
  return mean / sd;
}

function simulate(
  base: PlanDay[],
  loads: number[],
  startAtl: number,
  startCtl: number,
  sports: string[],
  zones: string[],
  teAe: number[],
  teAn: number[],
  locks: boolean[],
  ctlHistorySeed: number[],
  loadHistorySeed: number[],
  cfg: Record<string, any> | null
): SimDay[] {
  let atl = startAtl;
  let ctl = startCtl;
  const history = ctlHistorySeed.length >= 7 ? ctlHistorySeed.slice(-7) : Array(7).fill(startCtl);
  return base.map((d, idx) => {
    const load = Math.max(0, loads[idx] ?? 0);
    const hasTimelineAnchor =
      typeof d.timeline_atl === 'number' &&
      typeof d.timeline_ctl === 'number' &&
      Number.isFinite(d.timeline_atl) &&
      Number.isFinite(d.timeline_ctl);

    if (idx === 0 || hasTimelineAnchor) {
      // Der erste sichtbare Tag und jeder vorhandene Timeline-Anker wird 1:1 angezeigt.
      // Diese Werte sind bereits "nach Belastung" und dürfen nicht erneut berechnet werden.
      atl = hasTimelineAnchor ? Number(d.timeline_atl) : startAtl;
      ctl = hasTimelineAnchor ? Number(d.timeline_ctl) : startCtl;
    } else {
      // Rekursive Post-Load-Fortschreibung, geschätzt aus historischen KK_TIMELINE-Werten:
      // ATL_t = -8.819849 + 0.83645661 * ATL_{t-1} + 1.24059231 * Load_t
      // CTL_t = -3.950420 + 0.97283824 * CTL_{t-1} + 0.23367340 * Load_t
      atl = -8.819849 + 0.83645661 * atl + 1.24059231 * load;
      ctl = -3.950420 + 0.97283824 * ctl + 0.23367340 * load;
    }

    const acwr = typeof d.timeline_acwr === 'number' && (idx === 0 || hasTimelineAnchor)
      ? Number(d.timeline_acwr)
      : ctl > 0 ? atl / ctl : 0;
    const monotony = typeof d.timeline_monotony === 'number' && (idx === 0 || hasTimelineAnchor)
      ? Number(d.timeline_monotony)
      : rollingMonotonyWithSeed(loadHistorySeed, loads, idx);
    const ctl7dAgo = history[0] ?? startCtl;
    // GSheet-Analog: coachE_Smart_Gains / KEI
    // = ((coachE_CTL_forecast - CTL_vor_7_Tagen) * 10) / (ACWR * (1 + Monotony7)) * 0.1
    const kei = typeof d.timeline_kei === 'number' && (idx === 0 || hasTimelineAnchor)
      ? Number(d.timeline_kei)
      : acwr > 0 && monotony > -1 ? ((ctl - ctl7dAgo) * 10) / (acwr * (1 + monotony)) * 0.1 : 0;
    // Base CTL history already includes the anchor day. Shift only after simulated future rows.
    if (idx > 0) {
      history.push(ctl);
      history.shift();
    }
    const manualSport = sports[idx] ?? d.training_1 ?? '';
    const manualZone = zones[idx] ?? d.recommended_zone ?? '';
    const manualAe = Number.isFinite(teAe[idx]) ? teAe[idx] : estimateTe(load, manualSport).ae;
    const manualAn = Number.isFinite(teAn[idx]) ? teAn[idx] : estimateTe(load, manualSport).an;
    const flag: SimDay['flag'] = acwr > 1.35 || monotony > 2.4 ? 'RISK' : acwr > 1.25 || monotony > 2.0 ? 'WATCH' : 'OK';
    // SG_Flags: für vergangene/verankerte Tage mit gemessenem HRV wird die Sheet-Formel 1:1 reproduziert;
    // für zukünftige Tage bleibt der H-Anteil leer (kein HRV verfügbar), M kommt aus der simulierten Monotony7.
    const sgFlags = computeSgFlags({
      hrvStatus: d.timeline_hrv_status ?? null,
      hrvThresholds: d.timeline_hrv_thresholds ?? null,
      monotony,
    });
    return {
      ...d,
      locked: !!locks[idx],
      sim_load: load,
      sport: load <= 0 ? 'Off' : manualSport || 'Training',
      phase: phaseForIndex(idx),
      zone: load <= 0 ? 'Off' : manualZone || 'Z2',
      te_ae: load <= 0 ? 0 : manualAe,
      te_an: load <= 0 ? 0 : manualAn,
      atl,
      ctl,
      ctl_star: ctl,
      acwr,
      monotony,
      kei,
      flag,
      sg_flags: sgFlags,
      hrv_status: d.timeline_hrv_status ?? null,
      hrv_thresholds: d.timeline_hrv_thresholds ?? null,
      ai_load: null,
    };
  });
}

function ZoneChip({ zone }: { zone?: string }) {
  const lower = (zone || '').toLowerCase();
  let cls = 'bg-bg-subtle border-border-strong text-ink-muted';
  if (lower.includes('off') || lower.includes('ruhe')) cls = 'bg-ampel-grau/10 border-ampel-grau/40 text-gray-300';
  else if (lower.includes('z2')) cls = 'bg-ampel-gruen/10 border-ampel-gruen/40 text-green-300';
  else if (lower.includes('z4') || lower.includes('schwelle')) cls = 'bg-ampel-orange/10 border-ampel-orange/40 text-orange-300';
  return <span className={`chip border ${cls}`}>{zone || '—'}</span>;
}

export function PlanCockpit({ token, toast, proxyAuthenticated }: Props) {
  const [plan, setPlan] = useState<PlanResponse | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [loads, setLoads] = useState<number[]>([]);
  const [sports, setSports] = useState<string[]>([]);
  const [zones, setZones] = useState<string[]>([]);
  const [teAe, setTeAe] = useState<number[]>([]);
  const [teAn, setTeAn] = useState<number[]>([]);
  const [locks, setLocks] = useState<boolean[]>([]);
  const [startAtl, setStartAtl] = useState(930);
  const [startCtl, setStartCtl] = useState(949);
  const [ctlHistory, setCtlHistory] = useState<number[]>([]);
  const [loadHistory, setLoadHistory] = useState<number[]>([]);
  const [simConfig, setSimConfig] = useState<Record<string, any> | null>(null);
  const [maxEss, setMaxEss] = useState(300);
  const [source, setSource] = useState<'PlanApp' | 'Fallback'>('Fallback');
  const [sourceReason, setSourceReason] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [briefing, setBriefing] = useState<string | null>(null);
  const [briefingTs, setBriefingTs] = useState<string | null>(null);
  const [briefingRecommendations, setBriefingRecommendations] = useState<BriefingRecommendation[] | null>(null);
  const [briefingBusy, setBriefingBusy] = useState(false);

  // Nachladen der PlanApp-Simulation nach Timeout: nur übernehmen, wenn der Nutzer
  // seit dem Laden nichts verändert hat und kein neuerer Ladevorgang läuft.
  const loadSeqRef = useRef(0);
  const userEditedRef = useRef(false);

  async function load() {
    setLoading(true);
    const mySeq = ++loadSeqRef.current;
    userEditedRef.current = false;
    try {
      // Alle Requests parallel: planBriefing blockiert nicht mehr,
      // planSimulationJson bekommt einen harten 8s-Timeout (Apps Script antwortet
      // aktuell mit HTTP 404 nach ~60s, was den Fallback früher ausslösen würde).
      const briefingP = fetchPlanBriefing().catch((): null => null);
      const anchorP = fetchChartData('360d', PLAN_ANCHOR_METRICS as any).catch(
        (e): { data: any[]; timestamp?: string; _err?: unknown } => ({
          data: [],
          _err: e,
        }),
      );
      // Warm (Apps-Script-Cache 60s): 1-3s. Kaltstart kann >12s dauern.
      // Nach 12s wird vorläufig mit Timeline-Ankern gerendert; die Anfrage läuft bis 60s weiter
      // und wird automatisch übernommen, sobald sie eintrifft (siehe unten).
      const simFull: Promise<any> = fetchPlanSimulationWithTimeout(60000).catch(
        (e): { ok: false; _err: unknown } => ({ ok: false, _err: e }),
      );
      const simP: Promise<any> = Promise.race([
        simFull,
        new Promise((resolve) =>
          setTimeout(() => resolve({ ok: false, _pending: true, _err: new TimeoutError() }), 12000),
        ),
      ]);

      const [briefingRes, anchorRes, simRes] = await Promise.all([briefingP, anchorP, simP]);

      if (briefingRes && (briefingRes as any).ok) {
        const normalized = normalizeBriefingPayload(
          (briefingRes as any).briefing,
          (briefingRes as any).recommendations || null,
        );
        setBriefing(normalized.briefing);
        setBriefingTs(
          (briefingRes as any).briefingTimestamp || (briefingRes as any).timestamp || null,
        );
        setBriefingRecommendations(normalized.recommendations);
      }

      const anchorRows = dedupeByDateKeepLast((anchorRes.data as any[]) || []);
      const anchorByDate = new Map(anchorRows.map((row) => [row.date, row]));

      // ---- Happy path: planSimulationJson liefert Tage ----
      const applySim = (sim: any) => {
        const firstDate = sim.days[0]?.date;
        const loadSeed = anchorRows
          .filter((row) => typeof row.date === 'string' && (!firstDate || row.date <= firstDate))
          .slice(-6)
          .map((row) => toNum(row.coachE_ESS_day) ?? 0);
        const normalized: PlanDay[] = sim.days.map((d: any) => ({
          index: d.index,
          date: d.date,
          day: d.day,
          original_load_ess: d.load,
          recommended_load_ess: d.load,
          recommended_zone: d.zone,
          training_1: d.sport,
          training_2: d.zone,
          projected_effect: 'Live aus getSimStartValues / PlanApp-Startwerten.',
          weather_recommendation: '',
          timeline_atl:
            typeof d.atl === 'number' ? d.atl : toNum(anchorByDate.get(d.date)?.coachE_ATL_forecast),
          timeline_ctl:
            typeof d.ctl === 'number' ? d.ctl : toNum(anchorByDate.get(d.date)?.coachE_CTL_forecast),
          timeline_acwr:
            typeof d.acwr === 'number' ? d.acwr : toNum(anchorByDate.get(d.date)?.coachE_ACWR_forecast),
          timeline_kei:
            typeof d.kei === 'number' ? d.kei : toNum(anchorByDate.get(d.date)?.coachE_Smart_Gains),
          timeline_monotony:
            typeof d.monotony === 'number' ? d.monotony : toNum(anchorByDate.get(d.date)?.Monotony7),
          timeline_hrv_status: toNum(anchorByDate.get(d.date)?.hrv_status),
          timeline_hrv_thresholds:
            typeof anchorByDate.get(d.date)?.hrv_threshholds === 'string'
              ? (anchorByDate.get(d.date)?.hrv_threshholds as string)
              : null,
        }));
        setPlan({
          ok: true,
          timestamp: sim.timestamp || new Date().toISOString(),
          rows: normalized.length,
          plan: normalized,
        });
        setLoads(sim.days.map((d: any) => d.load));
        setSports(sim.days.map((d: any) => d.sport || ''));
        setZones(sim.days.map((d: any) => d.zone || ''));
        setTeAe(sim.days.map((d: any) => d.te_ae || 0));
        setTeAn(sim.days.map((d: any) => d.te_an || 0));
        setLocks(sim.days.map((d: any) => !!d.locked));
        setStartAtl(Number(sim.base?.atl || 930));
        setStartCtl(Number(sim.base?.ctl || 949));
        setCtlHistory(sim.base?.ctlHistory || []);
        setLoadHistory(loadSeed);
        setSimConfig(sim.base?.config || null);
        setSource('PlanApp');
        setSourceReason(null);
        setErr(null);
      };

      if ((simRes as any).ok && (simRes as any).days?.length) {
        applySim(simRes);
        return;
      }

      if ((simRes as any)._pending) {
        simFull.then((late: any) => {
          if (loadSeqRef.current !== mySeq) return;
          if (late && late.ok && late.days?.length) {
            if (userEditedRef.current) {
              setSourceReason('PlanApp-Simulation ist inzwischen da, wurde aber nicht übernommen, weil du schon Werte geändert hast. „Aktualisieren“ lädt sie.');
              return;
            }
            applySim(late);
            toast('info', 'PlanApp-Simulation nachgeladen – Live-Daten aktiv.');
          } else {
            setSourceReason('PlanApp-Simulation nicht verfügbar - Timeline-Anker werden genutzt.');
          }
        });
      }

      // ---- Fallback: 14 Tage direkt aus Timeline-Ankerdaten ----
      // Erzeugt KEINE fiktiven WEEK_LOAD-/CURRENT_SIM_PATTERN-Zeilen mehr.
      const simErr = (simRes as any)._err;
      const simReason =
        (simRes as any)._pending
          ? 'PlanApp-Simulation lädt noch (Kaltstart des Apps Script) - vorläufig Timeline-Anker, Live-Daten werden automatisch übernommen.'
          : simErr instanceof TimeoutError
          ? 'Timeout beim Abrufen der PlanApp-Simulation (>60s) - Timeline-Anker werden genutzt.'
          : simErr instanceof ApiError
          ? `PlanApp-Simulation nicht verfügbar (${simErr.message}) - Timeline-Anker werden genutzt.`
          : simErr
          ? 'PlanApp-Simulation nicht verfügbar - Timeline-Anker werden genutzt.'
          : 'PlanApp-Simulation lieferte keine Tage - Timeline-Anker werden genutzt.';

      const today = isoDate(new Date());
      const fallbackPlan: PlanDay[] = [];
      for (let i = 0; i < 14; i++) {
        const date = addDays(today, i);
        const a = anchorByDate.get(date);
        fallbackPlan.push({
          index: i + 1,
          date,
          day: dayName(date),
          original_load_ess: toNum(a?.coachE_ESS_day) ?? 0,
          recommended_load_ess: toNum(a?.coachE_ESS_day) ?? 0,
          recommended_zone: '',
          training_1: '',
          training_2: '',
          projected_effect: 'Timeline-Anker (planSimulationJson nicht verfügbar).',
          weather_recommendation: '',
          timeline_atl: toNum(a?.coachE_ATL_forecast),
          timeline_ctl: toNum(a?.coachE_CTL_forecast),
          timeline_acwr: toNum(a?.coachE_ACWR_forecast),
          timeline_kei: toNum(a?.coachE_Smart_Gains),
          timeline_monotony: toNum(a?.Monotony7),
          timeline_hrv_status: toNum(a?.hrv_status),
          timeline_hrv_thresholds:
            typeof a?.hrv_threshholds === 'string' ? (a.hrv_threshholds as string) : null,
        });
      }
      setPlan({
        ok: true,
        timestamp: new Date().toISOString(),
        rows: fallbackPlan.length,
        plan: fallbackPlan,
      });
      setLoads(fallbackPlan.map((d) => Number(d.recommended_load_ess ?? 0)));
      setSports(fallbackPlan.map(() => ''));
      setZones(fallbackPlan.map(() => ''));
      setTeAe(fallbackPlan.map(() => 0));
      setTeAn(fallbackPlan.map(() => 0));
      setLocks(fallbackPlan.map(() => false));
      // Startwerte aus dem letzten Anker VOR/AM ersten Plantag,
      // damit die Rekursion fortschreibt statt bei Default 930/949 zu starten.
      const firstDate = fallbackPlan[0]?.date;
      const anchorBefore = anchorRows
        .filter((row) => typeof row.date === 'string' && (!firstDate || row.date <= firstDate))
        .slice(-1)[0];
      setStartAtl(toNum(anchorBefore?.coachE_ATL_forecast) ?? 930);
      setStartCtl(toNum(anchorBefore?.coachE_CTL_forecast) ?? 949);
      const loadSeed = anchorRows
        .filter((row) => typeof row.date === 'string' && (!firstDate || row.date <= firstDate))
        .slice(-6)
        .map((row) => toNum(row.coachE_ESS_day) ?? 0);
      setCtlHistory([]);
      setLoadHistory(loadSeed);
      setSimConfig(null);
      setSource('Fallback');
      setSourceReason(simReason);
      setErr(anchorRows.length === 0 ? 'Keine Timeline-Daten geladen.' : null);
    } catch (e) {
      setErr(e instanceof ApiError ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  const sim = useMemo(() => {
    const base = plan?.plan || [];
    const rows = simulate(base, loads, startAtl, startCtl, sports, zones, teAe, teAn, locks, ctlHistory, loadHistory, simConfig);
    if (briefingRecommendations?.length && plan?.plan?.length) {
      const byDate = new Map(briefingRecommendations.map((r) => [r.date, r]));
      return rows.map((row, i) => ({
        ...row,
        ai_load: byDate.get(plan.plan[i]?.date)?.recommended_load_ess ?? null,
      }));
    }
    return rows;
  }, [plan, loads, startAtl, startCtl, sports, zones, teAe, teAn, locks, ctlHistory, loadHistory, simConfig, briefingRecommendations]);

  const keiAvg = sim.length ? sim.reduce((a, b) => a + b.kei, 0) / sim.length : 0;
  const maxAcwr = sim.length ? Math.max(...sim.map((d) => d.acwr)) : 0;
  const finalCtl = sim.length ? sim[sim.length - 1].ctl : startCtl;

  function setLoad(idx: number, value: number) {
    userEditedRef.current = true;
    setLoads((old) => old.map((v, i) => (i === idx ? value : v)));
  }

  function setArrayValue<T>(setter: (fn: (old: T[]) => T[]) => void, idx: number, value: T) {
    userEditedRef.current = true;
    setter((old) => old.map((v, i) => (i === idx ? value : v)));
  }

  function reset() {
    if (!plan?.plan) return;
    setLoads(plan.plan.map((d) => Number(d.recommended_load_ess ?? d.original_load_ess ?? 0)));
  }

  async function saveToTimeline() {
    if (!token.trim() && !proxyAuthenticated) {
      toast('err', 'Bitte zuerst Action-Token im Command Center eingeben.');
      return;
    }
    setSaving(true);
    try {
      const payload = {
        loads,
        teAe,
        teAn,
        sports,
        zones,
        locks,
      };
      const res = proxyAuthenticated
        ? await proxySavePlanSimulation(payload)
        : await savePlanSimulation(token.trim(), payload);
      if (res.ok) toast('ok', res.message || 'Plan in timeline gespeichert.');
      else toast('err', res.error || res.raw || 'Speichern fehlgeschlagen.');
    } catch (e) {
      toast('err', e instanceof ApiError ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  async function onGenerateBriefing() {
    if (!token.trim() && !proxyAuthenticated) {
      toast('err', 'Bitte zuerst Action-Token im Command Center eingeben.');
      return;
    }
    setBriefingBusy(true);
    try {
      const res = proxyAuthenticated
        ? await proxyGeneratePlanBriefing()
        : await generatePlanBriefing(token.trim());
      if (res.ok) {
        const normalized = normalizeBriefingPayload(res.briefing, res.recommendations || null);
        setBriefing(normalized.briefing);
        setBriefingTs(res.briefingTimestamp || res.timestamp || null);
        setBriefingRecommendations(normalized.recommendations);
        toast('ok', 'Plan-Briefing erstellt und gespeichert.');
      } else {
        toast('err', res.error || 'Plan-Briefing fehlgeschlagen.');
      }
    } catch (e) {
      toast('err', e instanceof ApiError ? e.message : String(e));
    } finally {
      setBriefingBusy(false);
    }
  }

  function applyBriefingLoads() {
    if (!briefingRecommendations?.length || !plan?.plan?.length) return;
    const byDate = new Map(briefingRecommendations.map((r) => [r.date, r]));
    setLoads((old) => old.map((v, i) => byDate.get(plan.plan[i]?.date)?.recommended_load_ess ?? v));
    setZones((old) => old.map((v, i) => byDate.get(plan.plan[i]?.date)?.recommended_zone ?? v));
    toast('info', 'ESS-/Zonen-Empfehlungen aus dem Briefing lokal übernommen. Noch nicht gespeichert.');
  }

  return (
    <div className="space-y-5">
      <Panel
        title="Strategische Simulation · 14 Tage"
        right={
          <div className="flex flex-wrap items-center justify-end gap-2">
            <span className="text-2xs text-ink-dim tnum">
              {plan?.timestamp && <>Stand · {fmtTime(plan.timestamp)} · Quelle {source}</>}
            </span>
            <button
              onClick={saveToTimeline}
              disabled={saving}
              className="btn btn-primary text-xs px-2 py-1"
              title="Token-geschützt: schreibt analog PlanApp in timeline"
            >
              {saving ? 'Speichern…' : 'In timeline speichern'}
            </button>
            <button onClick={reset} className="btn btn-ghost text-xs px-2 py-1">
              Reset
            </button>
            <button onClick={load} className="btn btn-ghost text-xs px-2 py-1">
              Aktualisieren
            </button>
          </div>
        }
      >
        {err ? (
          <ErrorBox message={err} onRetry={load} />
        ) : loading ? (
          <div className="space-y-3">
            <Skeleton className="h-24" />
            <Skeleton className="h-96" />
          </div>
        ) : (
          <div className="space-y-4">
            {source === 'Fallback' && sourceReason && (
              <div className="panel-raised p-3 border border-ampel-gelb/40 bg-ampel-gelb/10 text-xs text-yellow-200">
                <strong className="font-semibold">Hinweis:</strong> {sourceReason}
              </div>
            )}
            <div className="grid gap-3 lg:grid-cols-5">
              <div className="panel-raised p-3">
                <div className="label">Start ATL</div>
                <div className="tnum text-2xl font-semibold mt-1">{fmtNum(startAtl)}</div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">Start CTL</div>
                <div className="tnum text-2xl font-semibold mt-1">{fmtNum(startCtl)}</div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">Final CTL</div>
                <div className="tnum text-2xl font-semibold mt-1 text-accent">{fmtNum(finalCtl)}</div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">Max ACWR</div>
                <div className={`tnum text-2xl font-semibold mt-1 ${maxAcwr > 1.3 ? 'text-ampel-orange' : 'text-ampel-gruen'}`}>
                  {fmtNum(maxAcwr, 2)}
                </div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">KEI Ø</div>
                <div className={`tnum text-2xl font-semibold mt-1 ${keiAvg < 0 ? 'text-ampel-rot' : keiAvg < 3 ? 'text-ampel-gelb' : 'text-ampel-gruen'}`}>
                  {fmtSigned(keiAvg, 1)}
                </div>
              </div>
            </div>

            <div className="panel-raised p-4 border-accent-strong/20">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div>
                  <div className="label">Perplexity Sportwissenschaftlerin · Coach Kira Briefing</div>
                  <p className="text-2xs text-ink-dim tnum mt-1">
                    {briefingTs ? `Gespeichert · ${fmtTime(briefingTs)}` : 'Noch kein persistentes Plan-Briefing gespeichert.'}
                  </p>
                </div>
                <div className="flex flex-wrap gap-2">
                  {briefingRecommendations?.length ? (
                    <button onClick={applyBriefingLoads} className="btn btn-ghost text-xs px-2 py-1">
                      ESS übernehmen
                    </button>
                  ) : null}
                  <button
                    onClick={onGenerateBriefing}
                    disabled={briefingBusy}
                    className="btn btn-primary text-xs px-2 py-1"
                  >
                    {briefingBusy ? 'Erstelle Briefing…' : 'Briefing erstellen'}
                  </button>
                </div>
              </div>
              {briefing ? (
                <div className="mt-3 whitespace-pre-wrap text-sm text-ink-muted leading-relaxed">{briefing}</div>
              ) : (
                <p className="mt-3 text-sm text-ink-muted leading-relaxed">
                  Erstellt auf Knopfdruck eine persistente Empfehlung aus Command-Center-Status,
                  Garmin-/Coach-Kira-Werten und sichtbarem 14-Tage-Plan. Die ESS-Empfehlungen können
                  danach lokal in die Slider übernommen und anschließend explizit in die timeline gespeichert werden.
                </p>
              )}
              {briefingRecommendations?.length ? (
                <div className="mt-4 grid gap-2 md:grid-cols-2 xl:grid-cols-4">
                  {briefingRecommendations.slice(0, 14).map((r) => (
                    <div key={r.date} className="rounded border border-border bg-bg-subtle/50 p-2">
                      <div className="tnum text-xs text-ink">{r.day || ''} {r.date}</div>
                      <div className="tnum text-lg font-semibold text-accent">{fmtNum(r.recommended_load_ess)} ESS</div>
                      <div className="text-2xs text-ink-muted">{r.recommended_zone || 'Zone offen'}</div>
                    </div>
                  ))}
                </div>
              ) : null}
            </div>

            <div className="panel-raised p-3">
              <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
                <div>
                  <div className="label">Parameter</div>
                  <p className="text-xs text-ink-muted mt-1">
                    Lokale Simulation: ATL/CTL nach aktueller Last rekursiv aus KK_TIMELINE-Schätzung · KEI=((CTL−CTL vor 7 Tagen)×10)/(ACWR×(1+Monotonie))×0,1
                  </p>
                </div>
                <label className="flex items-center gap-2 text-xs text-ink-muted">
                  MAX ESS
                  <input
                    type="number"
                    value={maxEss}
                    onChange={(e) => setMaxEss(Math.max(50, Number(e.target.value) || 300))}
                    className="input w-20"
                  />
                </label>
              </div>
              <div className="h-[360px]">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={sim}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="day" />
                    <YAxis yAxisId="load" orientation="left" domain={[0, 'auto']} />
                    <YAxis yAxisId="acwr" orientation="right" domain={[0, 1.8]} />
                    <YAxis yAxisId="kei" orientation="right" domain={['auto', 'auto']} hide />
                    <Tooltip />
                    {briefingRecommendations?.length ? (
                      <Bar yAxisId="load" dataKey="ai_load" name="KI ESS" fill="#facc15" fillOpacity={0.28} radius={[3, 3, 0, 0]} maxBarSize={18} />
                    ) : null}
                    <Line yAxisId="load" type="monotone" dataKey="atl" name="ATL" stroke="#f97316" strokeWidth={1.8} dot={false} />
                    <Line yAxisId="load" type="monotone" dataKey="ctl" name="CTL" stroke="#22c55e" strokeWidth={1.8} dot={false} />
                    <Line yAxisId="acwr" type="monotone" dataKey="acwr" name="ACWR" stroke="#a855f7" strokeWidth={1.8} dot={false} />
                    <Line yAxisId="kei" type="monotone" dataKey="kei" name="KEI" stroke="#e879f9" strokeWidth={1.8} dot={false} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
              <div className="mt-3 flex flex-wrap gap-3 text-2xs tnum text-ink-muted">
                <span><span className="inline-block h-2 w-4 rounded bg-yellow-300/30 mr-1" />KI ESS</span>
                <span className="text-orange-300">ATL</span>
                <span className="text-green-300">CTL</span>
                <span className="text-purple-300">ACWR</span>
                <span className="text-fuchsia-300">KEI</span>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="min-w-[1180px] text-xs">
                <thead className="sticky top-0 bg-bg-panel z-10">
                  <tr className="text-left border-b border-border">
                    {['Datum', 'Phase', 'Last (ESS)', 'Sport', 'Zone', 'TE AE/AN', 'ATL', 'CTL', 'ACWR', 'KI', 'SG', 'Flag'].map((h) => (
                      <th key={h} className="label px-2 py-2">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-border">
                  {sim.map((d, i) => (
                    <tr key={`${d.date}-${i}`} className="align-top hover:bg-bg-subtle/40">
                      <td className="px-2 py-3">
                        <div className="tnum font-medium">{d.day} {fmtDateShort(d.date)}</div>
                        <div className="text-2xs text-ink-dim">Tag {i + 1}</div>
                      </td>
                      <td className="px-2 py-3"><span className="chip border border-border-strong bg-bg-subtle">{d.phase}</span></td>
                      <td className="px-2 py-3 w-64">
                        <div className="flex items-center gap-3">
                          <input
                            type="range"
                            min={0}
                            max={maxEss}
                            step={5}
                            value={d.sim_load}
                            disabled={d.locked}
                            onChange={(e) => setLoad(i, Number(e.target.value))}
                            className="w-40 accent-sky-400"
                            data-testid={`slider-load-${i}`}
                          />
                          <span className="tnum w-10 text-right font-semibold">{fmtNum(d.sim_load)}</span>
                        </div>
                      </td>
                      <td className="px-2 py-3">
                        <input
                          value={sports[i] || ''}
                          disabled={d.locked}
                          onChange={(e) => setArrayValue<string>(setSports, i, e.target.value)}
                          className="input w-24"
                          placeholder="Bike"
                        />
                      </td>
                      <td className="px-2 py-3">
                        <input
                          value={zones[i] || ''}
                          disabled={d.locked}
                          onChange={(e) => setArrayValue<string>(setZones, i, e.target.value)}
                          className="input w-20"
                          placeholder="Z2"
                        />
                      </td>
                      <td className="px-2 py-3">
                        <div className="flex items-center gap-1">
                          <input
                            type="number"
                            min={0}
                            max={5}
                            step={0.1}
                            value={teAe[i] ?? 0}
                            disabled={d.locked}
                            onChange={(e) => setArrayValue<number>(setTeAe, i, Number(e.target.value))}
                            className="input w-14"
                          />
                          <span className="text-ink-dim">/</span>
                          <input
                            type="number"
                            min={0}
                            max={5}
                            step={0.1}
                            value={teAn[i] ?? 0}
                            disabled={d.locked}
                            onChange={(e) => setArrayValue<number>(setTeAn, i, Number(e.target.value))}
                            className="input w-14"
                          />
                        </div>
                      </td>
                      <td className="px-2 py-3 tnum">{fmtNum(d.atl)}</td>
                      <td className="px-2 py-3 tnum">{fmtNum(d.ctl)}</td>
                      <td className="px-2 py-3">
                        <div className="flex flex-col gap-1">
                          <span className={`chip border ${d.acwr > 1.3 ? 'border-ampel-orange/40 bg-ampel-orange/10 text-orange-300' : 'border-ampel-gruen/40 bg-ampel-gruen/10 text-green-300'}`}>
                            ACWR {fmtNum(d.acwr, 2)}
                          </span>
                          <span className="chip border border-border-strong bg-bg-subtle text-ink-muted">
                            MONO {fmtNum(d.monotony, 2)}
                          </span>
                        </div>
                      </td>
                      <td className={`px-2 py-3 tnum font-semibold ${d.kei < 0 ? 'text-ampel-rot' : d.kei < 3 ? 'text-ampel-gelb' : d.kei < 8 ? 'text-accent' : 'text-ampel-gruen'}`}>
                        {fmtSigned(d.kei, 1)}
                      </td>
                      <td className="px-2 py-3">
                        {d.sg_flags ? (
                          <span
                            className={`chip border tnum ${
                              d.sg_flags.includes('H') && d.sg_flags.includes('M')
                                ? 'border-ampel-rot/40 bg-ampel-rot/10 text-red-300'
                                : d.sg_flags.includes('H')
                                ? 'border-ampel-orange/40 bg-ampel-orange/10 text-orange-300'
                                : 'border-ampel-gelb/40 bg-ampel-gelb/10 text-yellow-300'
                            }`}
                            title={
                              (d.sg_flags.includes('H') ? 'H: HRV-Alarm (hrv_status < unterer Schwellwert − 2)' : '') +
                              (d.sg_flags.includes('H') && d.sg_flags.includes('M') ? ' · ' : '') +
                              (d.sg_flags.includes('M') ? 'M: Monotony7 > 1,6' : '')
                            }
                          >
                            {d.sg_flags}
                          </span>
                        ) : (
                          <span className="text-ink-dim text-2xs">—</span>
                        )}
                      </td>
                      <td className="px-2 py-3">
                        <span className={`chip border ${d.flag === 'OK' ? 'border-ampel-gruen/40 bg-ampel-gruen/10 text-green-300' : d.flag === 'WATCH' ? 'border-ampel-gelb/40 bg-ampel-gelb/10 text-yellow-300' : 'border-ampel-rot/40 bg-ampel-rot/10 text-red-300'}`}>
                          {d.flag}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="grid gap-3 md:grid-cols-4">
              <div className="panel-raised p-3">
                <div className="label">High Efficiency</div>
                <div className="text-sm text-ampel-gruen mt-1">KEI ≥ 8</div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">Productive</div>
                <div className="text-sm text-accent mt-1">KEI 3 bis &lt; 8</div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">Maintenance</div>
                <div className="text-sm text-ampel-gelb mt-1">KEI 0 bis &lt; 3</div>
              </div>
              <div className="panel-raised p-3">
                <div className="label">Risk</div>
                <div className="text-sm text-ampel-rot mt-1">KEI &lt; 0 oder ACWR hoch</div>
              </div>
            </div>
          </div>
        )}
      </Panel>
    </div>
  );
}
