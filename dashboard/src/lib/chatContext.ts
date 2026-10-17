/**
 * Datenpaket für den Kira-Chat: alles, was die Heute-Seite und das Plan Cockpit wissen,
 * kompakt als JSON. Rein deterministisch (kein Modell).
 */
import { BASE_URL, fetchChartData, fetchPlanSimulationWithTimeout, toNum } from './api';
import { atlStep, ctlStep, esStep, windowMean, ES_MODEL } from './ctlModel';
import { buildDays, localIsoDate, RECOVERY_FETCH_METRICS, VOTE_LABEL, type RecoveryDay } from '../screens/RecoveryPanel';

const EXTRA = ['fbATL_obs', 'fbCTL_obs', 'garminEnduranceScore', 'Trainingszustand'] as const;
const r1 = (v: number | null | undefined, d = 1) => (v == null || !Number.isFinite(v) ? null : Math.round(v * 10 ** d) / 10 ** d);
const mean = (xs: (number | null | undefined)[]) => {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};

export type ChatContext = Record<string, unknown>;

export async function buildChatContext(): Promise<ChatContext> {
  const today = localIsoDate();
  const [chart, sim] = await Promise.all([
    fetchChartData('120d', [...RECOVERY_FETCH_METRICS, ...EXTRA]),
    fetchPlanSimulationWithTimeout(45000).catch(() => null as any),
  ]);
  const raw = ((chart as any).data || []) as any[];
  const days: RecoveryDay[] = buildDays(raw).filter((d) => d.date <= today);
  const byDate = new Map(raw.map((r) => [String(r.date).slice(0, 10), r]));
  const t = days[days.length - 1];
  const rowT = t ? byDate.get(t.date) : null;

  // Serien
  const loads = days.map((d) => d.essDay ?? d.load ?? 0);
  const ctlSer = days.map((d) => toNum(byDate.get(d.date)?.fbCTL_obs));
  const atlT = toNum(rowT?.fbATL_obs);
  const ctlT = toNum(rowT?.fbCTL_obs);
  const ctl7 = ctlSer.length >= 8 ? ctlSer[ctlSer.length - 8] : null;
  const last28 = loads.slice(-28);

  // Erholungs-Index (wie Planungs-Kennzahlen)
  const prior = days.slice(-29, -1);
  const hb = mean(prior.map((d) => d.hrv));
  const rb = mean(prior.map((d) => d.rhr));
  const parts: number[] = [];
  const clamp = (x: number) => Math.max(0, Math.min(100, x));
  if (t?.hrv != null && hb) parts.push(clamp(50 + 5 * ((t.hrv / hb - 1) * 100)));
  if (t?.rhr != null && rb != null) parts.push(clamp(50 - 10 * (t.rhr - rb)));
  if (t?.sleepScore != null) parts.push(clamp(t.sleepScore));
  if (t?.readiness != null) parts.push(clamp(t.readiness));
  const recIdx = parts.length >= 2 ? Math.round(parts.reduce((a, b) => a + b, 0) / parts.length) : null;

  // Plan + Prognose (Formeln wie Plan Cockpit)
  const planDays = (sim && sim.ok ? sim.days || [] : []).filter((x: any) => String(x.date).slice(0, 10) > today).slice(0, 7);
  let atl = atlT ?? 0;
  let ctl = ctlT ?? 0;
  let es = toNum(rowT?.garminEnduranceScore) ?? null;
  const ser = [...loads];
  const plan = planDays.map((x: any) => {
    const load = toNum(x.load) ?? 0;
    ser.push(load);
    atl = atlStep(atl, load);
    ctl = ctlStep(ctl, load, windowMean(ser));
    if (es != null) es = esStep(es, load, windowMean(ser, ES_MODEL.win));
    return {
      datum: String(x.date).slice(0, 10),
      last_ess: load,
      sport: x.sport || '',
      zone: x.zone || '',
      te_aerob: toNum(x.te_ae),
      te_anaerob: toNum(x.te_an),
      fixiert: !!x.locked,
      prognose: { atl: r1(atl, 0), ctl: r1(ctl, 0), acwr: r1(ctl ? atl / ctl : null, 2), form_pct: r1(ctl ? ((ctl - atl) / ctl) * 100 : null), ausdauer: r1(es, 0) },
    };
  });

  return {
    stand: new Date().toISOString(),
    heute: t
      ? {
          datum: t.date,
          tagesentscheid: t.vote ? VOTE_LABEL[t.vote] : null,
          gruende: t.reasons,
          aktivitaet_erledigt: t.activityDone,
          absolviert: t.activityDone ? { last_ess: t.actual.load, sport: t.actual.sport, zone: t.actual.zone, te_aerob: t.actual.teAe, te_anaerob: t.actual.teAn } : null,
          readiness: t.readiness,
          schlaf_h: r1(t.sleepH),
          schlafscore: t.sleepScore,
          hrv_ms: t.hrv,
          hrv_normalbereich: t.hrvLow != null && t.hrvHigh != null ? [t.hrvLow, t.hrvHigh] : null,
          ruhepuls: t.rhr,
          ruhepuls_abweichung_bpm: r1(t.rhrDelta),
          befinden_1_5: t.befinden,
          trainingszustand_garmin: rowT?.Trainingszustand || null,
          ausdauerwert_garmin: toNum(rowT?.garminEnduranceScore),
        }
      : null,
    kennzahlen: {
      atl: atlT,
      ctl: ctlT,
      acwr: r1(atlT != null && ctlT ? atlT / ctlT : null, 2),
      form_pct: r1(atlT != null && ctlT ? ((ctlT - atlT) / ctlT) * 100 : null),
      ctl_rampe_7t_pct: r1(ctlT && ctl7 ? (ctlT / ctl7 - 1) * 100 : null),
      monotonie_7t: r1(t?.monotony ?? null, 2),
      ruhetage_28t: last28.filter((l) => l < 1).length,
      lange_tage_28t: last28.filter((l) => l >= 200).length,
      erholungs_index_0_100: recIdx,
      kei_hinweis: 'KEI = CTL-Zuwachs 7 T / (max(ACWR;0,8) × (1+Monotonie))',
    },
    regeln: {
      votum: 'Monotonie > 1,6 oder ACWR > 1,3 → max. Train; Monotonie > 2,0 oder ACWR > 1,5 → max. Easy',
      ziele: 'CTL-Rampe 3–5 %/Woche; Form vor harter Einheit > −15 %, vor Formspitze +5 bis +15 %; Schlafziel 7,5 h',
    },
    letzte_14_tage: days.slice(-14).map((d) => ({
      datum: d.date,
      last_ess: d.essDay ?? d.load,
      sport: d.actual.sport || null,
      readiness: d.readiness,
      schlaf_h: r1(d.sleepH),
      hrv: d.hrv,
      ruhepuls: d.rhr,
      befinden: d.befinden,
    })),
    plan_naechste_7_tage: plan,
  };
}

export type ChatMsg = { role: 'user' | 'assistant'; content: string; citations?: string[]; ts?: string };

export async function sendChat(messages: ChatMsg[], context: ChatContext, sources: boolean) {
  const res = await fetch(`${BASE_URL}?mode=chat&cb=${Date.now()}`, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ messages: messages.map(({ role, content }) => ({ role, content })), context, sources }),
  });
  const text = await res.text();
  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error('Antwort des Apps Script ist kein JSON (Chat-Patch eingespielt?)');
  }
  if (json && json.status && json.ok === undefined) throw new Error('Der Chat-Patch ist im Apps Script noch nicht aktiv (doPost kennt mode=chat nicht).');
  if (!json.ok) throw new Error(json.error || 'Chat-Fehler');
  return json as { ok: true; answer: string; citations: string[]; used: number; limit: number; model: string };
}
