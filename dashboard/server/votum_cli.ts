/**
 * Etappe 2 (Vorbereitung): Votum-Abgleich Cockpit ↔ Supervisor, nur lesend.
 * Bau: npx esbuild server/votum_cli.ts --bundle --platform=node --format=esm --outfile=server/votum.mjs
 * Aufruf: node server/votum.mjs → JSON auf stdout
 */
import './gsFetch.mjs'; // Restrisiko 3: Token für Lesezugriffe
import { buildDays, RECOVERY_FETCH_METRICS, STUFE_COCKPIT, stufeSupervisor, vergleiche, type RecoveryDay } from '../src/kern/votum';
import { fetchChartData, fetchPlanSimulationWithTimeout } from '../src/lib/api';
import { lastEmpfehlung, lageImRahmen, persoenlicheBaender, type Baender } from '../src/kern/soll';
import { rangeFor } from '../src/lib/windows';
import { signaleAusLive } from '../src/kern/garminSignale';

// 22.11.2026: persönliche Soll-Obergrenzen aus der ganzen Historie (wie Cockpit)
let baenderJetzt: Baender | null = null;
async function baenderLaden(): Promise<Baender | null> {
  if (baenderJetzt) return baenderJetzt;
  try {
    const all: any = await fetchChartData('9999d' as any, RECOVERY_FETCH_METRICS as any);
    baenderJetzt = persoenlicheBaender(buildDays(all?.data || []));
  } catch { baenderJetzt = null; }
  return baenderJetzt;
}

const V2_URL = process.env.CK_V2_SCRIPT_URL ||
  'https://script.google.com/macros/s/AKfycbwfLqLJ56WtQj6CdfufJW3VMclCGyF1XL6EUOBR--Jm3UOlqALIB0Y32wJ6xIZhFw93/exec';
const iso = (d = new Date()) => d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });

const heute = iso();
const [res, st]: any[] = await Promise.all([
  fetchChartData(rangeFor(30, 28, 90), RECOVERY_FETCH_METRICS as any),
  fetch(`${V2_URL}?mode=status&cb=${Date.now()}`, { redirect: 'follow' }).then((r) => r.json()).catch((e) => ({ ok: false, error: String(e) })),
]);
// 27.11.2026: Garmin-Signale (live vom Mac) nur in die heutige Zeile
const garminHeute = (() => { try { return signaleAusLive(JSON.parse(process.env.CK_LIVE || 'null'), heute); } catch { return null; } })();
const rohZeilen = (res?.data || []).map((r: any) => (garminHeute && String(r?.date || '').slice(0, 10) === heute ? { ...r, __garmin: garminHeute } : r));
const days: RecoveryDay[] = buildDays(rohZeilen).filter((d) => d.date <= heute);
const t = days.find((d) => d.date === heute) || null;
// Etappe 2b: Soll-Rahmen aus dem Kern gegen den Plan im Sheet (coachE_ESS_day bzw. Plan-Simulation Tag 1)
const ti = t ? days.indexOf(t) : -1;
const rahmen = t && ti > 0 ? lastEmpfehlung(t, days[ti - 1], await baenderLaden()) : null;
let planSim: any = null;
try {
  const sim: any = await fetchPlanSimulationWithTimeout(45000);
  planSim = sim?.ok ? (sim.days || []).find((x: any) => String(x.date).slice(0, 10) === heute) || null : null;
} catch {}
const planLoad = planSim ? Number(planSim.load) : t?.essDay ?? null;
const istLoad = t?.activityDone ? t.actual?.load ?? t.load ?? null : null;
const rueck = days.slice(-31).map((d, i, a) => {
  if (i === 0 || d.date >= heute) return null;
  const e = lastEmpfehlung(d, a[i - 1], baenderJetzt);
  const ist = d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0;
  return e ? { d: d.date, v: d.vote, min: e.min, max: e.max, ist, lage: lageImRahmen(ist, e) } : null;
}).filter(Boolean) as any[];
const sum = st?.summary || {};
const svStand = st?.heartbeat?.timestamp ? iso(new Date(st.heartbeat.timestamp)) : null;
const svAktuell = svStand === heute && st?.heartbeat?.stage === 'END_OK';
const c = t?.vote ? STUFE_COCKPIT[t.vote] : null;
const s = svAktuell ? stufeSupervisor(sum.gesamtAmpel) : null;
process.stdout.write(JSON.stringify({
  ok: true,
  datum: heute,
  cockpit: t ? { vote: t.vote, stufe: c, bind: t.bind ?? null, gruende: (t.reasons || []).slice(0, 3), readiness: t.readiness, hrv: t.hrv, acwr: t.acwr } : null,
  supervisor: { ampel: sum.gesamtAmpel ?? null, stufe: s, score: sum.gesamtScore ?? null, recovery: sum.recoveryScore ?? null, training: sum.trainingScore ?? null, planStatus: sum.planStatus ?? null, lauf: st?.heartbeat?.timestamp ?? null, aktuell: svAktuell },
  vergleich: vergleiche(c, s),
  soll: {
    rahmen,
    plan: planLoad != null && Number.isFinite(planLoad) ? { load: planLoad, sport: planSim?.sport || '', zone: planSim?.zone || '', gesperrt: !!planSim?.locked, quelle: planSim ? 'planSimulation' : 'coachE_ESS_day' } : null,
    planLage: lageImRahmen(planLoad, rahmen),
    ist: istLoad,
    istLage: istLoad != null ? lageImRahmen(istLoad, rahmen) : null,
  },
  sollRueckblick: { tage: rueck.length, imRahmen: rueck.filter((x) => x.lage === 'im Rahmen').length, darueber: rueck.filter((x) => x.lage === 'darüber').length, darunter: rueck.filter((x) => x.lage === 'darunter').length, verlauf: rueck },
  cockpitVerlauf: days.slice(-30).map((d) => ({ d: d.date, v: d.vote, b: d.bind ?? null })),
}));
