/**
 * Etappe 2 (Vorbereitung): Votum-Abgleich Cockpit ↔ Supervisor, nur lesend.
 * Bau: npx esbuild server/votum_cli.ts --bundle --platform=node --format=esm --outfile=server/votum.mjs
 * Aufruf: node server/votum.mjs → JSON auf stdout
 */
import { buildDays, RECOVERY_FETCH_METRICS, STUFE_COCKPIT, stufeSupervisor, vergleiche, type RecoveryDay } from '../src/kern/votum';
import { fetchChartData } from '../src/lib/api';
import { rangeFor } from '../src/lib/windows';

const V2_URL = process.env.CK_V2_SCRIPT_URL ||
  'https://script.google.com/macros/s/AKfycbwfLqLJ56WtQj6CdfufJW3VMclCGyF1XL6EUOBR--Jm3UOlqALIB0Y32wJ6xIZhFw93/exec';
const iso = (d = new Date()) => d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });

const heute = iso();
const [res, st]: any[] = await Promise.all([
  fetchChartData(rangeFor(30, 28, 90), RECOVERY_FETCH_METRICS as any),
  fetch(`${V2_URL}?mode=status&cb=${Date.now()}`, { redirect: 'follow' }).then((r) => r.json()).catch((e) => ({ ok: false, error: String(e) })),
]);
const days: RecoveryDay[] = buildDays(res?.data || []).filter((d) => d.date <= heute);
const t = days.find((d) => d.date === heute) || null;
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
  cockpitVerlauf: days.slice(-30).map((d) => ({ d: d.date, v: d.vote, b: d.bind ?? null })),
}));
