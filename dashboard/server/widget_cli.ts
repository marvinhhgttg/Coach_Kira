/**
 * Widget-Rechenkern für /api/widget – nutzt dieselben Funktionen wie das Cockpit
 * (Heute: buildDays/Votum, Plan heute aus der Plan-Simulation, Tourform wie in der Tourform-Kachel).
 * Gebaut zu server/widget.mjs:  npx esbuild server/widget_cli.ts --bundle --platform=node --format=esm --outfile=server/widget.mjs
 * Aufruf: node server/widget.mjs  → JSON auf stdout
 */
import { buildDays, RECOVERY_FETCH_METRICS, VOTE_LABEL, VOTE_LEVEL, type RecoveryDay } from '../src/kern/votum';
import { lastEmpfehlung } from '../src/kern/soll';
import { fetchChartData, fetchPlanSimulationWithTimeout } from '../src/lib/api';
import { rangeFor } from '../src/lib/windows';
import { fetchExtras } from '../src/lib/extras';
import { ladeP4, tourform as berechneTourform } from '../src/lib/paket4';
import { zielArt, type Ziel } from '../src/lib/planning';
import { fetchGarminTage } from '../src/lib/garmin';
import { lerneModelle } from '../src/lib/essRechner';
import { eigeneBergtouren, eigeneSkitouren, setzeGarminTage, tourLastPersoenlich } from '../src/lib/tourEss';

const iso = (d = new Date()) => d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
const num = (v: unknown) => {
  const n = typeof v === 'number' ? v : Number(String(v ?? '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
};
const STATUS_TEXT: Record<string, string> = { QUALITY: 'Qualität möglich', TRAIN: 'normal trainieren', EASY: 'locker trainieren', REST: 'Erholung priorisieren' };

async function heuteTeil() {
  const heute = iso();
  // wie Heute-Seite: Span 14 Tage + 28 Tage Vorlauf, mindestens 90 Tage
  const res: any = await fetchChartData(rangeFor(14, 28, 90), RECOVERY_FETCH_METRICS as any);
  const days: RecoveryDay[] = buildDays(res.data || []).filter((d) => d.date <= heute);
  const t = days.length && days[days.length - 1].date === heute ? days[days.length - 1] : null;
  let plan: any = null;
  let naechste: { d: string; ess: number | null; sport: string; zone: string }[] = [];
  try {
    const sim: any = await fetchPlanSimulationWithTimeout(45000);
    plan = sim?.ok ? (sim.days || []).find((x: any) => String(x.date).slice(0, 10) === heute) || null : null;
    // Uhr Seite 3: die nächsten 4 Tage aus der Plan-Simulation (Timeline)
    if (sim?.ok)
      naechste = (sim.days || [])
        .filter((x: any) => String(x.date).slice(0, 10) > heute)
        .slice(0, 4)
        .map((x: any) => ({ d: String(x.date).slice(0, 10), ess: num(x.sim_load ?? x.load), sport: String(x.sport || ''), zone: String(x.zone || '') }));
  } catch {}
  // 22.11.2026: Soll = Soll-Rahmen des Kerns (wie „Empfehlung heute“ im Cockpit), Plan getrennt
  const ti = t ? days.length - 1 : -1;
  const rahmen = t && ti > 0 ? lastEmpfehlung(t, days[ti - 1]) : null;
  const planLoad = plan ? num(plan.load) : t?.essDay ?? null;
  const woche = days.slice(-7).map((d) => ({ d: d.date, ist: d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0 }));
  return {
    datum: heute,
    votum: t?.vote
      ? { stufe: VOTE_LEVEL[t.vote], vote: t.vote, label: VOTE_LABEL[t.vote], text: STATUS_TEXT[t.vote], gruende: t.reasons.slice(0, 3) }
      : { stufe: 'grau', vote: null, label: '–', text: t ? 'Morgenwerte fehlen' : 'noch keine Daten für heute', gruende: t?.reasons?.slice(0, 1) || [] },
    morgen: t
      ? {
          eingetragen: !!t.hasMorningData,
          readiness: t.readiness,
          schlaf: t.sleepScore,
          schlafH: t.sleepH,
          hrv: t.hrv,
          hrvStufe: t.levels.hrv,
          rhr: t.rhr,
          rhrDelta: t.rhrDelta,
          befinden: t.befinden,
        }
      : null,
    last: {
      // sollHeute bleibt für ältere App-/Uhr-Versionen: Obergrenze des Rahmens
      sollHeute: rahmen ? rahmen.max : planLoad,
      sollMin: rahmen ? rahmen.min : null,
      sollMax: rahmen ? rahmen.max : null,
      sollText: rahmen ? `${rahmen.min}–${rahmen.max}` : planLoad != null ? String(Math.round(planLoad)) : null,
      halten: rahmen ? rahmen.halte : null,
      planHeute: planLoad,
      planLage: rahmen && planLoad != null && planLoad > 0 ? (planLoad > rahmen.max ? 'darueber' : planLoad < rahmen.min ? 'darunter' : 'im Rahmen') : null,
      sportHeute: plan ? String(plan.sport || '') : '',
      zoneHeute: plan ? String(plan.zone || '') : '',
      istHeute: t?.activityDone ? t.actual?.load ?? t.load ?? null : 0,
      erledigt: !!t?.activityDone,
      acwr: t?.acwr ?? null,
      acwrBand: t?.acwr == null ? null : t.acwr > 1.5 ? 'hoch' : t.acwr > 1.3 ? 'erhoeht' : t.acwr < 0.8 ? 'niedrig' : 'korridor',
      woche,
    },
    naechste,
  };
}

/** Uhr Seite 4: Form (CTL/ATL/Monotonie, Endurance Score mit 28-Tage-Änderung) und Garmin-Tag. */
async function formTeil() {
  const heute = iso();
  const res: any = await fetchChartData('35d' as any, ['fbCTL_obs', 'fbATL_obs', 'garminEnduranceScore', 'Monotony7'] as any);
  const rows = (res?.data || []).filter((r: any) => String(r.date).slice(0, 10) <= heute);
  const letzter = (k: string) => { for (let i = rows.length - 1; i >= 0; i--) { const v = num(rows[i][k]); if (v != null && v !== 0) return { v, d: String(rows[i].date).slice(0, 10), i }; } return null; };
  const es = letzter('garminEnduranceScore');
  let es28: number | null = null;
  if (es) {
    const ziel = new Date(es.d + 'T12:00:00Z'); ziel.setUTCDate(ziel.getUTCDate() - 28);
    const zIso = ziel.toISOString().slice(0, 10);
    const alt = rows.find((r: any) => String(r.date).slice(0, 10) >= zIso && num(r.garminEnduranceScore));
    es28 = alt ? Math.round(es.v - (num(alt.garminEnduranceScore) as number)) : null;
  }
  let tag: any = null;
  try {
    const g = await fetchGarminTage();
    const t = g.get(heute);
    if (t) tag = { bbAufwachen: t.bb_aufwachen ?? null, bbMax: t.tag?.bb_max ?? null, stress: t.tag?.stress_avg ?? null, schritte: t.tag?.schritte ?? null };
  } catch {}
  return { ctl: letzter('fbCTL_obs')?.v ?? null, atl: letzter('fbATL_obs')?.v ?? null, monotonie: letzter('Monotony7')?.v ?? null, es: es?.v ?? null, es28, tag };
}

async function zielTeil() {
  const heute = iso();
  const items = await fetchExtras<{ ziel?: Ziel }>('ziel', 60);
  const ziele = items
    .map((it) => it.data?.ziel)
    .filter((z): z is Ziel => !!z && !!z.datum && z.datum >= heute)
    .sort((a, b) => a.datum.localeCompare(b.datum));
  const z = ziele[0];
  if (!z) return null;
  const tage = Math.round((new Date(z.datum + 'T12:00:00Z').getTime() - new Date(heute + 'T12:00:00Z').getTime()) / 86400000);
  let prozent: number | null = null;
  try {
    const [p4, garmin] = await Promise.all([ladeP4(), fetchGarminTage()]);
    setzeGarminTage(garmin);
    const modelle = garmin ? lerneModelle(garmin) : [];
    const pers = modelle.length ? tourLastPersoenlich(z, modelle, eigeneSkitouren(garmin), eigeneBergtouren(garmin)) : null;
    const wf0 = berechneTourform(p4, z, pers?.ess);
    prozent = wf0.prozent;
    if (zielArt(z) === 'marsch' && garmin?.size) {
      // wie Tourform-Kachel: Strecke und Zeit am Stück (letzte 8 Wochen)
      const ab = new Date(Date.now() - 56 * 86400000).toLocaleDateString('sv-SE');
      let km = 0, h = 0;
      for (const [d, t] of garmin) {
        if (d < ab) continue;
        km = Math.max(km, (t.akt || []).filter((x: any) => /hiking|walking|mountaineering/i.test(String(x.typ || ''))).reduce((a: number, x: any) => a + (x.km || 0), 0));
        h = Math.max(h, ((t as any).summe?.dauer_s ?? (t.akt || []).reduce((a: number, x: any) => a + (x.dauer_s || 0), 0)) / 3600);
      }
      const strecke = z.distanz_km > 0 ? Math.min(1, km / (z.distanz_km * 0.6)) : null;
      const zeit = z.dauer_h > 0 ? Math.min(1, h / (z.dauer_h * 0.7)) : null;
      prozent = Math.round(100 * (0.2 * wf0.volumen + 0.45 * (strecke ?? wf0.lange) + 0.35 * (zeit ?? 1)));
    }
  } catch {}
  return { name: z.name, datum: z.datum, tage, prio: (z as any).prio ?? null, tourform: prozent, weitere: ziele.length - 1 };
}

const [h, z, f] = await Promise.all([heuteTeil(), zielTeil().catch(() => null), formTeil().catch(() => null)]);
process.stdout.write(JSON.stringify({ ok: true, stand: new Date().toISOString(), ...h, ziel: z, form: f }));
