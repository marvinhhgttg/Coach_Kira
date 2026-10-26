/**
 * Datenpaket für den Kira-Chat: alles, was die Heute-Seite und das Plan Cockpit wissen,
 * kompakt als JSON. Rein deterministisch (kein Modell).
 */
import { BASE_URL, fetchChartData, fetchPlanSimulationWithTimeout, toNum } from './api';
import { fetchGarminTage, teLabelDe, type GarminTage } from './garmin';
import { atlStep, ctlStep, esStep, windowMean, ES_MODEL } from './ctlModel';
import { buildDays, localIsoDate, RECOVERY_FETCH_METRICS, VOTE_LABEL, type RecoveryDay } from '../screens/RecoveryPanel';
import { MONOTONIE_REGEL_TEXT } from './monotony';

const EXTRA = ['fbATL_obs', 'fbCTL_obs', 'garminEnduranceScore', 'Trainingszustand'] as const;
const r1 = (v: number | null | undefined, d = 1) => (v == null || !Number.isFinite(v) ? null : Math.round(v * 10 ** d) / 10 ** d);
const mean = (xs: (number | null | undefined)[]) => {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};

export type ChatContext = Record<string, unknown>;

export async function buildChatContext(): Promise<ChatContext> {
  const today = localIsoDate();
  const [chart, sim, garmin] = await Promise.all([
    fetchChartData('120d', [...RECOVERY_FETCH_METRICS, ...EXTRA]),
    fetchPlanSimulationWithTimeout(45000).catch(() => null as any),
    fetchGarminTage().catch(() => new Map() as GarminTage),
  ]);
  // Garmin Personal API: kompakte Tageswerte (Schlafphasen, Body Battery, Stress, Aktivitäten mit Zonen)
  const gTag = (d: string) => {
    const g = garmin.get(d);
    if (!g) return null;
    const zs = g.summe?.zonen_s;
    const zsum = zs ? zs.reduce((a: number, b: number) => a + (b || 0), 0) : 0;
    return {
      km: g.summe?.km ?? null, hm: g.summe?.hm ?? null,
      bewegung_min: g.summe ? Math.round((g.summe.bewegt_s || g.summe.dauer_s || 0) / 60) : null,
      zonen_pct: zs && zsum ? zs.map((z: number) => Math.round(((z || 0) / zsum) * 100)) : null,
      aktivitaeten: (g.akt || []).map((a) => ({ name: a.name, sport: a.sport, km: a.km, hm: a.hm, min: Math.round((a.bewegt_s || a.dauer_s || 0) / 60), hf: a.hf_avg, te_label: a.te_label ? teLabelDe(a.te_label) : null, watt_np: a.watt_np, bb: a.bb })),
      schlaf_tief_min: g.schlaf?.tief_s != null ? Math.round(g.schlaf.tief_s / 60) : null,
      schlaf_rem_min: g.schlaf?.rem_s != null ? Math.round(g.schlaf.rem_s / 60) : null,
      schlaf_wach_min: g.schlaf?.wach_s != null ? Math.round(g.schlaf.wach_s / 60) : null,
      schlafbedarf_min: g.schlaf?.bedarf_min ?? null, schlafstress: g.schlaf?.stress ?? null,
      atmung_schlaf: g.schlaf?.atem ?? null, spo2_schlaf: g.schlaf?.spo2 ?? null,
      body_battery_min_max: g.tag?.bb_min != null ? [g.tag.bb_min, g.tag.bb_max] : null, body_battery_aufwachen: g.bb_aufwachen ?? null,
      stress_avg: g.tag?.stress_avg ?? null, stress_hoch_min: g.tag?.stress_min?.hoch ?? null,
      schritte: g.tag?.schritte ?? null, intensitaetsmin: g.tag?.int_mod != null || g.tag?.int_vig != null ? (g.tag?.int_mod || 0) + 2 * (g.tag?.int_vig || 0) : null,
      gewicht_kg: g.gewicht_kg ?? null,
    };
  };
  const raw = ((chart as any).data || []) as any[];
  const days: RecoveryDay[] = buildDays(raw).filter((d) => d.date <= today);
  const byDate = new Map(raw.map((r) => [String(r.date).slice(0, 10), r]));
  const t = days[days.length - 1];
  const rowT = t ? byDate.get(t.date) : null;

  // Serien
  const loads = days.map((d) => d.essDay ?? d.load ?? 0);
  const ctlSer = days.map((d) => toNum(byDate.get(d.date)?.fbCTL_obs));
  // Letzter Tag mit ATL/CTL (morgens vor dem Training ist der heutige Wert oft noch leer)
  let li = days.length - 1;
  while (li > 0 && (toNum(byDate.get(days[li].date)?.fbCTL_obs) == null || toNum(byDate.get(days[li].date)?.fbATL_obs) == null)) li--;
  const rowL = li >= 0 && days[li] ? byDate.get(days[li].date) : null;
  const lastLoadDate = li >= 0 && days[li] ? days[li].date : null;
  const atlT = toNum(rowL?.fbATL_obs);
  const ctlT = toNum(rowL?.fbCTL_obs);
  const ctl7 = li >= 7 ? ctlSer[li - 7] : null;
  let esLast: number | null = null;
  for (let k = days.length - 1; k >= 0 && esLast == null; k--) esLast = toNum(byDate.get(days[k].date)?.garminEnduranceScore);
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
  const planDays = (sim && sim.ok ? sim.days || [] : []).filter((x: any) => String(x.date).slice(0, 10) > (lastLoadDate ?? today)).slice(0, 8);
  let atl = atlT ?? 0;
  let ctl = ctlT ?? 0;
  let es = esLast;
  const ser = loads.slice(0, li + 1);
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
          ausdauerwert_garmin: esLast,
          garmin: gTag(t.date),
        }
      : null,
    kennzahlen: {
      stand_belastung: lastLoadDate,
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
      votum: 'ACWR > 1,3 → max. Train; ACWR > 1,5 → max. Easy. Monotonie siehe monotonie_regel (Verlauf statt Einzelwert).',
      monotonie_regel: MONOTONIE_REGEL_TEXT,
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
      garmin: gTag(d.date),
    })),
    plan_naechste_7_tage: plan, // ab dem Tag nach dem letzten Belastungsstand (heute inkl., falls noch offen)
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

// ---------------------------------------------------------------------------
// Lokal: Ollama auf dem Mac (nur erreichbar, wenn das Cockpit auf diesem Mac läuft)
// ---------------------------------------------------------------------------
export const OLLAMA_URL = 'http://localhost:11434';
export const LOCAL_MODEL = 'coach-kira';

export const KIRA_SYSTEM = [
  'Du bist Kira, Sportwissenschaftlerin und persönliche Ausdauer-Trainerin von Marc (Jahrgang 1975; Rennrad, Gravel, Laufen, Bergsport).',
  'Sprich Marc per Du an, antworte auf Deutsch, klar und konkret, ohne Emojis.',
  'Stütze dich zuerst auf die mitgelieferten Cockpit-Daten (JSON unten). Nenne die Zahlen, auf die du dich beziehst.',
  'Erfinde keine Werte. Wenn Daten fehlen oder unklar sind, sag das offen.',
  'Halte dich an die Regeln aus dem Feld "regeln" (Tagesentscheid, Zielbereiche).',
  'Gib am Ende eine klare Empfehlung (z. B. Einheit, Dauer, Zone oder Pause).',
  'Bei Schmerzen, Krankheit oder medizinischen Fragen: auf Arzt bzw. Physiotherapie verweisen.',
].join('\n');

/** Prüft, ob Ollama erreichbar ist und das Modell coach-kira existiert. */
export async function pingLocal(): Promise<{ ok: boolean; hasModel: boolean; models: string[]; error?: string }> {
  try {
    const ctl = new AbortController();
    // großzügig, damit eine Browser-Freigabe (lokales Netzwerk) beantwortet werden kann
    const t = setTimeout(() => ctl.abort(), 15000);
    const res = await fetch(`${OLLAMA_URL}/api/tags`, { signal: ctl.signal, cache: 'no-store' });
    clearTimeout(t);
    const j = await res.json();
    const models: string[] = (j.models || []).map((m: any) => String(m.name || m.model || ''));
    return { ok: true, hasModel: models.some((m) => m === LOCAL_MODEL || m.startsWith(LOCAL_MODEL + ':')), models };
  } catch (e) {
    const name = e instanceof Error ? e.name : '';
    const msg = e instanceof Error ? e.message : String(e);
    const error =
      name === 'AbortError'
        ? 'keine Antwort in 15 s – vermutlich wartet der Browser auf die Freigabe „Lokales Netzwerk“'
        : /Failed to fetch|NetworkError|Load failed/i.test(msg)
          ? 'vom Browser blockiert (Freigabe „Lokales Netzwerk“ oder CORS/OLLAMA_ORIGINS)'
          : msg;
    return { ok: false, hasModel: false, models: [], error };
  }
}

/** Lokale Antwort mit Streaming. onToken erhält den bisher gesammelten Text. */
export async function sendChatLocal(
  messages: ChatMsg[],
  context: ChatContext,
  onToken: (textSoFar: string) => void,
  model = LOCAL_MODEL,
): Promise<{ answer: string; model: string; tokPerS: number | null }> {
  const sys = KIRA_SYSTEM + '\n\nCOCKPIT-DATEN:\n' + JSON.stringify(context);
  const res = await fetch(`${OLLAMA_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model,
      stream: true,
      think: false,
      keep_alive: '30m',
      options: { temperature: 0.3, num_ctx: 16384 },
      messages: [{ role: 'system', content: sys }, ...messages.slice(-12).map(({ role, content }) => ({ role, content }))],
    }),
  });
  if (!res.ok || !res.body) throw new Error(`Ollama HTTP ${res.status}`);
  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let buf = '';
  let answer = '';
  let tokPerS: number | null = null;
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    let nl: number;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).trim();
      buf = buf.slice(nl + 1);
      if (!line) continue;
      const j = JSON.parse(line);
      if (j.error) throw new Error(j.error);
      const piece = j.message?.content || '';
      if (piece) {
        answer += piece;
        onToken(answer);
      }
      if (j.done && j.eval_count && j.eval_duration) tokPerS = Math.round((j.eval_count / (j.eval_duration / 1e9)) * 10) / 10;
    }
  }
  return { answer, model, tokPerS };
}

/** Lokale Frage/Antwort in AI_CHAT_LOG protokollieren (zählt nicht gegen das Tageslimit). */
export async function logLocalChat(question: string, answer: string, model: string, engine: 'local' | 'cloud' = 'local') {
  try {
    await fetch(`${BASE_URL}?mode=chatLog&cb=${Date.now()}`, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ kind: 'chatLog', question, answer, model, engine }),
    });
  } catch {
    // Protokoll ist optional
  }
}
