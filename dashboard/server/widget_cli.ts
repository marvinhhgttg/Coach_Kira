/**
 * Widget-Rechenkern für /api/widget – nutzt dieselben Funktionen wie das Cockpit
 * (Heute: buildDays/Votum, Plan heute aus der Plan-Simulation, Tourform wie in der Tourform-Kachel).
 * Gebaut zu server/widget.mjs:  npx esbuild server/widget_cli.ts --bundle --platform=node --format=esm --outfile=server/widget.mjs
 * Aufruf: node server/widget.mjs  → JSON auf stdout
 */
import { buildDays, RECOVERY_FETCH_METRICS, VOTE_LABEL, VOTE_LEVEL, type RecoveryDay } from '../src/kern/votum';
import { lastEmpfehlung, persoenlicheBaender, type Baender } from '../src/kern/soll';
import { fetchChartData, fetchPlanSimulationWithTimeout } from '../src/lib/api';
import { rangeFor } from '../src/lib/windows';
import { fetchExtras } from '../src/lib/extras';
import { ladeP4, tourform as berechneTourform } from '../src/lib/paket4';
import { zielArt, type Ziel } from '../src/lib/planning';
import { fetchGarminTage } from '../src/lib/garmin';
import { lerneModelle } from '../src/lib/essRechner';
import { eigeneBergtouren, eigeneSkitouren, setzeGarminTage, tourLastPersoenlich } from '../src/lib/tourEss';

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

const iso = (d = new Date()) => d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
const num = (v: unknown) => {
  const n = typeof v === 'number' ? v : Number(String(v ?? '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
};
const STATUS_TEXT: Record<string, string> = { QUALITY: 'Qualität möglich', TRAIN: 'normal trainieren', EASY: 'locker trainieren', REST: 'Erholung priorisieren' };

let tageMerk: RecoveryDay[] = [];
// 23.11.2026: neuester Garmin-Wert eines Tages: nach Aktivität (fb…_obs) vor Morgenwert (garmin…)
const neuWert = (r: any, post: string, morgen: string): number | null => {
  const p = r ? num(r[post]) : null;
  if (p != null && p > 0) return p;
  const m = r ? num(r[morgen]) : null;
  return m != null && m > 0 ? m : null;
};

async function heuteTeil() {
  const heute = iso();
  // wie Heute-Seite: Span 14 Tage + 28 Tage Vorlauf, mindestens 90 Tage
  // fb_TR_obs = laufend synchronisierte Trainingsbereitschaft (nach Aktivität), Garmin_Training_Readiness = Morgenwert
  const res: any = await fetchChartData(rangeFor(14, 28, 90), [...RECOVERY_FETCH_METRICS, 'fb_TR_obs', 'fbACWR_obs', 'garminACWR'] as any);
  const days: RecoveryDay[] = buildDays(res.data || []).filter((d) => d.date <= heute);
  const acwrJetzt = (() => { const r = (res.data || []).find((x: any) => String(x.date).slice(0, 10) === heute); return neuWert(r, 'fbACWR_obs', 'garminACWR'); })();
  const trJetzt = (() => { const r = (res.data || []).find((x: any) => String(x.date).slice(0, 10) === heute); const v = r ? num(r.fb_TR_obs) : null; return v != null && v > 0 ? v : null; })();
  tageMerk = days;
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
        .slice(0, 7)
        .map((x: any) => ({ d: String(x.date).slice(0, 10), ess: num(x.sim_load ?? x.load), sport: String(x.sport || ''), zone: String(x.zone || '') }));
  } catch {}
  // 22.11.2026: Soll = Soll-Rahmen des Kerns (wie „Empfehlung heute“ im Cockpit), Plan getrennt
  const ti = t ? days.length - 1 : -1;
  const rahmen = t && ti > 0 ? lastEmpfehlung(t, days[ti - 1], await baenderLaden()) : null;
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
          // 22.11.2026: Anzeige immer der neueste Wert (nach Aktivität), Votum bleibt beim Morgenwert
          readiness: trJetzt ?? t.readiness,
          readinessMorgen: t.readiness,
          readinessQuelle: trJetzt != null && trJetzt !== t.readiness ? 'aktuell' : 'morgens',
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
      // 23.11.2026: Anzeige neuester Garmin-Wert von heute, Votum-Basis (gestern nach Aktivität) getrennt
      acwr: acwrJetzt ?? t?.acwr ?? null,
      acwrVotum: t?.acwr ?? null,
      acwrBand: (() => { const v = acwrJetzt ?? t?.acwr; return v == null ? null : v > 1.5 ? 'hoch' : v > 1.3 ? 'erhoeht' : v < 0.8 ? 'niedrig' : 'korridor'; })(),
      woche,
    },
    naechste,
    _rahmen: rahmen,
  };
}

/** Uhr Seite 4: Form (CTL/ATL/Monotonie, Endurance Score mit 28-Tage-Änderung) und Garmin-Tag. */
async function formTeil() {
  const heute = iso();
  const res: any = await fetchChartData('35d' as any, ['fbCTL_obs', 'fbATL_obs', 'garminCTL', 'garminATL', 'garminEnduranceScore', 'Monotony7'] as any);
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
  const letzterNeu = (post: string, morgen: string) => { for (let i = rows.length - 1; i >= 0; i--) { const v = neuWert(rows[i], post, morgen); if (v != null) return v; } return null; };
  return { ctl: letzterNeu('fbCTL_obs', 'garminCTL'), atl: letzterNeu('fbATL_obs', 'garminATL'), monotonie: letzter('Monotony7')?.v ?? null, es: es?.v ?? null, es28, tag };
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


// ─── 22.11.2026: Uhr-Diagramme, Datenfeld-Raten (nur mit ?mehr=1 bzw. /api/widget/feld) ───
const GRUPPE = (typ: string) =>
  /cycling|biking|virtual_ride/.test(typ) ? 'rad' : /running/.test(typ) ? 'lauf' : /hiking|mountaineering|walking/.test(typ) ? 'berg' : /ski/.test(typ) ? 'ski' : 'sonst';
/** ESS je Stunde in Zone 1–5 aus deinen Aktivitäten (nicht-negative kleinste Quadrate: Last ≈ Σ Stunden_z × Rate_z). */
function zonenRaten(xs: { h: number[]; last: number }[]): number[] | null {
  if (xs.length < 12) return null;
  // Koordinatenabstieg, leicht zu einem Startwert hingezogen (Zone 5 kommt selten vor)
  const prior = [20, 45, 80, 130, 190], lam = 0.5;
  const r = prior.slice();
  for (let it = 0; it < 300; it++) {
    for (let z = 0; z < 5; z++) {
      let num0 = lam * prior[z], den = lam;
      for (const x of xs) {
        const rest = x.last - x.h.reduce((a, h, k) => (k === z ? a : a + h * r[k]), 0);
        num0 += x.h[z] * rest; den += x.h[z] * x.h[z];
      }
      r[z] = Math.max(0, num0 / den);
    }
  }
  for (let z = 1; z < 5; z++) if (r[z] < r[z - 1]) r[z] = r[z - 1];
  return r.map((v) => Math.round(v));
}
async function feldTeil(rahmen: any, t: RecoveryDay | null) {
  const g = await fetchGarminTage();
  const pro: Record<string, { h: number[]; last: number }[]> = { rad: [], lauf: [], berg: [], ski: [], sonst: [], alle: [] };
  for (const [, tag] of g) for (const a of tag.akt || []) {
    const zs = (a as any).zonen_s as number[] | undefined;
    if (!zs || zs.length < 5 || !(a as any).last || ((a as any).dauer_s || 0) < 900) continue;
    const x = { h: zs.slice(0, 5).map((v) => (v || 0) / 3600), last: (a as any).last as number };
    pro[GRUPPE(String(a.typ || ''))].push(x);
    pro.alle.push(x);
  }
  const alle = zonenRaten(pro.alle) || [20, 45, 80, 130, 190];
  const raten: Record<string, number[]> = {};
  for (const k of ['rad', 'lauf', 'berg', 'ski', 'sonst']) raten[k] = zonenRaten(pro[k]) || alle;
  return {
    datum: iso(), vote: t?.vote ?? null, label: t?.vote ? VOTE_LABEL[t.vote] : '–',
    min: rahmen ? rahmen.min : null, max: rahmen ? rahmen.max : null, raten,
    n: Object.fromEntries(Object.entries(pro).map(([k, v]) => [k, v.length])),
  };
}
const LV: Record<string, string> = { gruen: 'g', gelb: 'y', orange: 'o', rot: 'r', blau: 'b' };
async function uhrTeil(rahmen: any) {
  const heute = iso();
  const days = tageMerk;
  const t = days.length && days[days.length - 1].date === heute ? days[days.length - 1] : null;
  const b = await baenderLaden();
  // Last 14 Tage mit Rahmen je Tag
  const l14: any = { von: null as string | null, ist: [] as number[], lo: [] as number[], hi: [] as number[] };
  const ab = Math.max(1, days.length - 14);
  for (let i = ab; i < days.length; i++) {
    const d = days[i];
    const e = lastEmpfehlung(d, days[i - 1], b);
    if (!l14.von) l14.von = d.date;
    l14.ist.push(Math.round(d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0));
    l14.lo.push(e ? e.min : -1);
    l14.hi.push(e ? e.max : -1);
  }
  // Morgenwerte 7 Tage
  const m = days.slice(-7);
  const m7 = {
    rd: m.map((d) => d.readiness ?? -1), hrv: m.map((d) => d.hrv ?? -1), rhr: m.map((d) => d.rhr ?? -1), sl: m.map((d) => d.sleepScore ?? -1),
    rdS: m.map((d) => LV[d.levels.readiness] || 'x').join(''), hrvS: m.map((d) => LV[d.levels.hrv] || 'x').join(''),
    rhrS: m.map((d) => LV[d.levels.rhr] || 'x').join(''), slS: m.map((d) => LV[d.levels.sleep] || 'x').join(''),
  };
  // Form: CTL/ATL 42 Tage, Ausdauerwert 90 Tage (jeder 3. Tag)
  let form: any = null;
  try {
    const r: any = await fetchChartData('100d' as any, ['fbCTL_obs', 'fbATL_obs', 'garminCTL', 'garminATL', 'garminEnduranceScore'] as any);
    const rows = (r?.data || []).filter((x: any) => String(x.date).slice(0, 10) <= heute);
    const reihe = (k: string, n: number) => rows.slice(-n).map((x: any) => { const v = num(x[k]); return v == null || v <= 0 ? -1 : Math.round(v); });
    const es = reihe('garminEnduranceScore', 90).filter((_: number, i: number, a: number[]) => (a.length - 1 - i) % 3 === 0);
    const reihe2 = (post: string, morgen: string, n: number) => rows.slice(-n).map((x: any) => { const v = neuWert(x, post, morgen); return v == null ? -1 : Math.round(v); });
    form = { ctl: reihe2('fbCTL_obs', 'garminCTL', 42), atl: reihe2('fbATL_obs', 'garminATL', 42), es };
  } catch {}
  // Letzte Einheit
  let letzte: any = null;
  try {
    const g = await fetchGarminTage();
    const ds = [...g.keys()].filter((d) => d <= heute).sort();
    for (let i = ds.length - 1; i >= 0 && !letzte; i--) {
      const akt = (g.get(ds[i])?.akt || []).slice().sort((x: any, y: any) => String(y.start || '').localeCompare(String(x.start || '')));
      // 22.11.2026: Kurz-/Testaufzeichnungen (unter 10 min oder ohne Last) überspringen
      const a: any = akt.find((x: any) => (x.dauer_s || 0) >= 600 && (x.last || 0) > 0);
      if (!a) continue;
      const di = days.findIndex((d) => d.date === ds[i]);
      const e = di > 0 ? lastEmpfehlung(days[di], days[di - 1], b) : null;
      const tagLast = (g.get(ds[i])?.akt || []).reduce((s2: number, x: any) => s2 + (x.last || 0), 0);
      // 22.11.2026: alle Einheiten des Tages (chronologisch) + Summen für die Uhr-Seite „Trainingstag“
      const echte = akt.filter((x: any) => (x.dauer_s || 0) >= 600 && (x.last || 0) > 0).reverse();
      const zSum = [0, 0, 0, 0, 0];
      for (const x of echte) ((x.zonen_s || []) as number[]).slice(0, 5).forEach((v, k) => { zSum[k] += v || 0; });
      letzte = {
        einheiten: echte.slice(0, 4).map((x: any) => ({ name: String(x.name || '').slice(0, 28), sport: String(x.typ || ''), min: Math.round((x.dauer_s || 0) / 60), ess: Math.round(x.last || 0) })),
        tagMin: Math.round(echte.reduce((s2: number, x: any) => s2 + (x.dauer_s || 0), 0) / 60),
        tagKm: Math.round(echte.reduce((s2: number, x: any) => s2 + (x.km || 0), 0) * 10) / 10,
        tagZ: zSum.map((v) => Math.round(v / 60)),
        d: ds[i], name: String(a.name || '').slice(0, 40), sport: String(a.typ || ''), min: Math.round((a.dauer_s || 0) / 60),
        ess: Math.round(a.last || 0), tagEss: Math.round(tagLast), te: a.te_ae ?? null, tean: a.te_an ?? null, label: a.te_label ?? null,
        z: ((a.zonen_s || []) as number[]).slice(0, 5).map((v) => Math.round((v || 0) / 60)), km: a.km ?? null, hm: a.hm ?? null,
        lo: e ? e.min : null, hi: e ? e.max : null,
        lage: e ? (tagLast > e.max ? 'darueber' : tagLast < e.min ? 'darunter' : 'im Rahmen') : null,
      };
    }
  } catch {}
  let feld: any = null;
  try { feld = await feldTeil(rahmen, t); } catch {}
  return { l14, m7, form, letzte, feld };
}

const [h, z, f] = await Promise.all([heuteTeil(), zielTeil().catch(() => null), formTeil().catch(() => null)]);
let u: any = null;
try { u = await uhrTeil((h as any)._rahmen ?? null); } catch {}
const { _rahmen: _r, ...hRest } = h as any;
process.stdout.write(JSON.stringify({ ok: true, stand: new Date().toISOString(), ...hRest, ziel: z, form: f, uhr: u }));
