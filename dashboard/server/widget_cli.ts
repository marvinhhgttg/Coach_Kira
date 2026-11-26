/**
 * Widget-Rechenkern für /api/widget – nutzt dieselben Funktionen wie das Cockpit
 * (Heute: buildDays/Votum, Plan heute aus der Plan-Simulation, Tourform wie in der Tourform-Kachel).
 * Gebaut zu server/widget.mjs:  npx esbuild server/widget_cli.ts --bundle --platform=node --format=esm --outfile=server/widget.mjs
 * Aufruf: node server/widget.mjs  → JSON auf stdout
 */
import './gsFetch.mjs'; // Restrisiko 3: Token für Lesezugriffe
import { buildDays, RECOVERY_FETCH_METRICS, VOTE_LABEL, VOTE_LEVEL, type RecoveryDay } from '../src/kern/votum';
import { datenstand, datenstandKurz } from '../src/kern/datenstand';
import { lastEmpfehlung, persoenlicheBaender, type Baender } from '../src/kern/soll';
import { fetchChartData, fetchPlanSimulationWithTimeout } from '../src/lib/api';
import { rangeFor } from '../src/lib/windows';
import { fetchExtras } from '../src/lib/extras';
import { ladeP4, tourform as berechneTourform } from '../src/lib/paket4';
import { zielArt, type Ziel } from '../src/lib/planning';
import { fetchGarminTage } from '../src/lib/garmin';
import { lerneModelle, vorwaerts } from '../src/lib/essRechner';
import { atlStep, ctlStep, windowMean } from '../src/kern/formeln';
import { eigeneBergtouren, eigeneSkitouren, setzeGarminTage, tourLastPersoenlich } from '../src/lib/tourEss';
import { ladeP3, schlafKonto } from '../src/lib/paket3';

// 22.11.2026: persönliche Soll-Obergrenzen aus der ganzen Historie (wie Cockpit)
let baenderJetzt: Baender | null = null;
let alleTage: RecoveryDay[] = [];
async function baenderLaden(): Promise<Baender | null> {
  if (baenderJetzt) return baenderJetzt;
  try {
    const all: any = await fetchChartData('9999d' as any, RECOVERY_FETCH_METRICS as any);
    alleTage = buildDays(all?.data || []);
    baenderJetzt = persoenlicheBaender(alleTage);
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
let sperreMerk: boolean[] = [];
// 24.11.2026: Garmin-Livewerte vom Mac (vom Server per Umgebung übergeben, nur heute und höchstens 3 h alt)
const LIVE: any = (() => {
  try {
    const w = JSON.parse(process.env.CK_LIVE || 'null');
    if (!w || w.datum !== iso() || Date.now() - Number(w.at) > 3 * 3600 * 1000) return null;
    return w;
  } catch { return null; }
})();
const liveSumme = (): number => (LIVE && Array.isArray(LIVE.akt) ? Math.round(LIVE.akt.reduce((s: number, x: any) => s + (Number(x.last) > 0 ? Number(x.last) : 0), 0)) : 0);
let vorlaeufig = false;
/** Stufe 1: fehlen die Morgenwerte im Sheet, setzt Kira die Garmin-Livewerte für heute ein (vorläufig, Sheet hat Vorrang). */
function mitLiveMorgen(rows: any[], heute: string): any[] {
  if (!LIVE) return rows;
  if (process.env.CK_TEST_OHNE_MORGEN === '1') rows = rows.map((r: any) => (String(r?.date || '').slice(0, 10) === heute ? { ...r, Garmin_Training_Readiness: null, rhr_bpm: null, sleep_hours: null, sleep_score_0_100: null, hrv_status: null, hrv_threshholds: null } : r));
  const hatMorgen = (r: any) => r && [r.Garmin_Training_Readiness, r.rhr_bpm, r.sleep_hours].some((v) => num(v) != null && (num(v) as number) > 0);
  const i = rows.findIndex((r: any) => String(r?.date || '').slice(0, 10) === heute);
  if (i >= 0 && hatMorgen(rows[i])) return rows;
  const liveMorgen = [LIVE.tr, LIVE.rhr, LIVE.schlafH].some((v: any) => v != null);
  if (!liveMorgen) return rows;
  const z: any = { ...(i >= 0 ? rows[i] : { date: heute }) };
  if (LIVE.tr != null) z.Garmin_Training_Readiness = LIVE.tr;
  if (LIVE.rhr != null) z.rhr_bpm = LIVE.rhr;
  if (LIVE.schlafH != null) z.sleep_hours = LIVE.schlafH;
  if (LIVE.schlafScore != null) z.sleep_score_0_100 = LIVE.schlafScore;
  if (LIVE.hrv != null) z.hrv_status = LIVE.hrv;
  if (LIVE.hrvLo != null && LIVE.hrvHi != null) z.hrv_threshholds = `${LIVE.hrvLo};${LIVE.hrvHi}`;
  vorlaeufig = true;
  const out = rows.slice();
  if (i >= 0) out[i] = z; else out.push(z);
  return out;
}
let plan7Merk: { start: string; ess: number[]; atl: number[]; ctl: number[]; acwr: number[] } | null = null;
// 24.11.2026: Auto-Plancheck – verlässt die ACWR-Prognose den Korridor (> 1,50), schlägt Kira kleinere Lasten vor
let vorschlagMerk: any = null;
const ACWR_GRENZE = 1.5;
// 23.11.2026: neuester Garmin-Wert eines Tages: nach Aktivität (fb…_obs) vor Morgenwert (garmin…)
const neuWert = (r: any, post: string, morgen: string): number | null => {
  const p = r ? num(r[post]) : null;
  if (p != null && p > 0) return p;
  const m = r ? num(r[morgen]) : null;
  return m != null && m > 0 ? m : null;
};

/** 24.11.2026: „Licht aus“ wie Schlafkonto-Kachel (Aufstehzeit − Bedarf − 15 min) – nur abends gerechnet. */
async function lichtAus(): Promise<{ lichtAus: string; bedarfH: number; aufstehen: string } | null> {
  try {
    const h = Number(new Date().toLocaleString('en-GB', { timeZone: 'Europe/Berlin', hour: '2-digit', hour12: false }));
    if (!(h >= 17) && process.env.CK_TEST_ABEND !== '1') return null;
    const t = await ladeP3();
    const k = schlafKonto(t, t.length ? t[t.length - 1].load ?? null : null);
    const morgen = new Date(Date.now() + 86400000);
    const wd = new Date(morgen.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' }) + 'T12:00:00Z').getUTCDay();
    const auf = /^\d{2}:\d{2}$/.test(process.env.CK_AUFSTEHEN || '') ? String(process.env.CK_AUFSTEHEN) : wd === 0 || wd === 6 ? '07:30' : '06:15';
    const [ah, am] = auf.split(':').map(Number);
    const m = ((Math.round(ah * 60 + am - k.bedarfH * 60 - 15) % 1440) + 1440) % 1440;
    return { lichtAus: `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`, bedarfH: Math.round(k.bedarfH * 100) / 100, aufstehen: auf };
  } catch {
    return null;
  }
}

async function heuteTeil() {
  const heute = iso();
  // wie Heute-Seite: Span 14 Tage + 28 Tage Vorlauf, mindestens 90 Tage
  // fb_TR_obs = laufend synchronisierte Trainingsbereitschaft (nach Aktivität), Garmin_Training_Readiness = Morgenwert
  const res: any = await fetchChartData(rangeFor(14, 28, 90), [...RECOVERY_FETCH_METRICS, 'fb_TR_obs', 'fbACWR_obs', 'garminACWR'] as any);
  const days: RecoveryDay[] = buildDays(mitLiveMorgen(res.data || [], heute)).filter((d) => d.date <= heute);
  const acwrJetzt = (() => { const r = (res.data || []).find((x: any) => String(x.date).slice(0, 10) === heute); return neuWert(r, 'fbACWR_obs', 'garminACWR'); })();
  const trJetzt = (() => { const r = (res.data || []).find((x: any) => String(x.date).slice(0, 10) === heute); const v = r ? num(r.fb_TR_obs) : null; return v != null && v > 0 ? v : null; })();
  tageMerk = days;
  const t = days.length && days[days.length - 1].date === heute ? days[days.length - 1] : null;
  let plan: any = null;
  let naechste: { d: string; ess: number | null; sport: string; zone: string }[] = [];
  let sperre: boolean[] = [];
  try {
    const sim: any = await fetchPlanSimulationWithTimeout(45000);
    plan = sim?.ok ? (sim.days || []).find((x: any) => String(x.date).slice(0, 10) === heute) || null : null;
    // Uhr Seite 3: die nächsten 4 Tage aus der Plan-Simulation (Timeline)
    if (sim?.ok)
      naechste = (sim.days || [])
        .filter((x: any) => String(x.date).slice(0, 10) > heute)
        .slice(0, 7)
        .map((x: any) => ({ d: String(x.date).slice(0, 10), ess: num(x.sim_load ?? x.load), sport: String(x.sport || ''), zone: String(x.zone || '') }));
    if (sim?.ok) sperre = (sim.days || []).filter((x: any) => String(x.date).slice(0, 10) > heute).slice(0, 7).map((x: any) => !!x.locked);
    sperreMerk = sperre;
  } catch {}
  // 23.11.2026: Plan der nächsten 7 Tage mit ATL/CTL-Prognose (Kiras Formel ab Stand Ende heute, wie Ausblick im Cockpit)
  try {
    if (t && t.atlPost != null && t.ctlPost != null && naechste.length) {
      let atl = t.atlPost, ctl = t.ctlPost;
      const loads = days.map((d) => d.essDay ?? 0);
      // Stufe 2: heute schon trainiert → Start bei Garmins Live-ATL/CTL nach der Einheit, heutige Last aus den Live-Einheiten
      const ls = liveSumme();
      if (LIVE && ls > 0 && LIVE.atl != null && LIVE.ctl != null) {
        atl = Number(LIVE.atl); ctl = Number(LIVE.ctl);
        if (loads.length) loads[loads.length - 1] = Math.max(loads[loads.length - 1], ls);
      }
      const sim7 = (ess: number[]) => {
        let a = atl, c = ctl;
        const ls = loads.slice();
        const o = { atl: [] as number[], ctl: [] as number[], acwr: [] as number[] };
        for (const l of ess) {
          ls.push(l);
          a = atlStep(a, l);
          c = ctlStep(c, l, windowMean(ls));
          o.atl.push(Math.round(a)); o.ctl.push(Math.round(c)); o.acwr.push(c > 0 ? Math.round((a / c) * 100) / 100 : 0);
        }
        return o;
      };
      const ess7 = naechste.slice(0, 7).map((x) => Math.round(x.ess ?? 0));
      const r7 = sim7(ess7);
      const p7 = { start: naechste[0].d, ess: ess7, atl: r7.atl, ctl: r7.ctl, acwr: r7.acwr };
      plan7Merk = p7;
      try {
        const maxAlt = Math.max(...r7.acwr);
        if (maxAlt > ACWR_GRENZE) {
          const neu = ess7.slice();
          let r = r7, n = 0;
          while (Math.max(...r.acwr) > ACWR_GRENZE && n < 80) {
            const k = r.acwr.findIndex((a) => a > ACWR_GRENZE);
            let best = -1;
            for (let i = 0; i <= k; i++) if (!sperre[i] && neu[i] > 0 && (best < 0 || neu[i] > neu[best])) best = i;
            if (best < 0) break;
            neu[best] = Math.max(0, neu[best] - 10);
            r = sim7(neu);
            n++;
          }
          const aenderungen = neu.map((v, i) => ({ d: naechste[i].d, sport: naechste[i].sport, zone: naechste[i].zone, alt: ess7[i], neu: v })).filter((x) => x.neu !== x.alt);
          if (aenderungen.length) {
            const kAlt = r7.acwr.indexOf(maxAlt);
            vorschlagMerk = {
              id: aenderungen.map((x) => `${x.d}:${x.alt}>${x.neu}`).join(','),
              grenze: ACWR_GRENZE, maxAlt, tagAlt: naechste[kAlt].d, maxNeu: Math.max(...r.acwr), erreicht: Math.max(...r.acwr) <= ACWR_GRENZE,
              aenderungen, essNeu: neu, acwrNeu: r.acwr,
            };
          }
        }
      } catch {}
    }
  } catch {}
  // 22.11.2026: Soll = Soll-Rahmen des Kerns (wie „Empfehlung heute“ im Cockpit), Plan getrennt
  const ti = t ? days.length - 1 : -1;
  const rahmen = t && ti > 0 ? lastEmpfehlung(t, days[ti - 1], await baenderLaden()) : null;
  const planLoad = plan ? num(plan.load) : t?.essDay ?? null;
  const woche = days.slice(-7).map((d) => ({ d: d.date, ist: d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0 }));
  // 23.11.2026: Wochen-Ring – Ist Mo..heute und Plan-Budget Mo..So (vergangene Tage Ist, ab heute Plan)
  const montag = (() => { const d = new Date(heute + 'T12:00:00Z'); const wd = (d.getUTCDay() + 6) % 7; d.setUTCDate(d.getUTCDate() - wd); return d.toISOString().slice(0, 10); })();
  const sonntag = (() => { const d = new Date(montag + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + 6); return d.toISOString().slice(0, 10); })();
  const istTag = (d: RecoveryDay) => (d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0);
  const wIstVorher = days.filter((d) => d.date >= montag && d.date < heute).reduce((a, d) => a + istTag(d), 0);
  const wIstHeute = t ? istTag(t) : 0;
  const wPlanRest = naechste.filter((x) => x.d <= sonntag).reduce((a, x) => a + (x.ess || 0), 0);
  const wocheTeil = { kw: montag, istVorher: Math.round(wIstVorher), ist: Math.round(wIstVorher + wIstHeute), budgetBasis: Math.round(wIstVorher + Math.max(wIstHeute, planLoad || 0) + wPlanRest) };
  // 26.11.2026: Datenstand des Votums (Kern) – gleiche Aussage wie im Cockpit
  const dsKern = datenstand(t, heute);
  return {
    datum: heute,
    woche: wocheTeil,
    daten: { stufe: dsKern.stufe, label: dsKern.label, kurz: datenstandKurz(dsKern), fehlt: dsKern.fehlt, optional: dsKern.optional },
    votum: t?.vote
      ? { stufe: VOTE_LEVEL[t.vote], vote: t.vote, label: VOTE_LABEL[t.vote], text: (vorlaeufig ? 'Vorläufig (Garmin live): ' : '') + STATUS_TEXT[t.vote], gruende: t.reasons.slice(0, 3), vorlaeufig, quelle: vorlaeufig ? 'garmin live' : 'sheet' }
      : { stufe: 'grau', vote: null, label: '–', text: t ? 'Morgenwerte fehlen' : 'noch keine Daten für heute', gruende: t?.reasons?.slice(0, 1) || [] },
    morgen: t
      ? {
          eingetragen: !!t.hasMorningData,
          vorlaeufig,
          // 22.11.2026: Anzeige immer der neueste Wert (nach Aktivität), Votum bleibt beim Morgenwert
          readiness: trJetzt ?? t.readiness,
          readinessMorgen: t.readiness,
          readinessQuelle: trJetzt != null && trJetzt !== t.readiness ? 'aktuell' : 'morgens',
          schlaf: t.sleepScore,
          schlafH: t.sleepH,
          hrv: t.hrv,
          hrvStufe: t.levels.hrv,
          rhr: t.rhr,
          rhrStufe: t.levels.rhr,
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
    planVorschlag: vorschlagMerk,
    schlaf: await lichtAus(),
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
    // 23.11.2026: ESS heute schon gespeicherter Einheiten – das Datenfeld rechnet Tagesstand = vorher + laufend
    // 24.11.2026: für die Datenfeld-Einstellungen (Zeile frei wählbar)
    acwr: (t as any)?.acwr != null ? Math.round((t as any).acwr * 100) / 100 : null, bereit: (t as any)?.readiness ?? null,
    vorher: Math.round(((g.get(iso())?.akt || []) as any[]).reduce((a: number, x: any) => a + ((x.last || 0) > 0 ? x.last : 0), 0)),
    n: Object.fromEntries(Object.entries(pro).map(([k, v]) => [k, v.length])),
  };
}
const LV: Record<string, string> = { gruen: 'g', gelb: 'y', orange: 'o', rot: 'r', blau: 'b' };
// 23.11.2026: Garmin-Farbmuster (R rot · O orange · G grün · b blau · v lila). Höher-besser-Werte: Kiras Grün = Garmin Blau, Gelb = Grün
const LV_G: Record<string, string> = { gruen: 'b', gelb: 'G', orange: 'O', rot: 'R', grau: 'v' };
const LV_HRV: Record<string, string> = { gruen: 'G', gelb: 'O', orange: 'O', rot: 'R' };
async function uhrTeil(rahmen: any) {
  const heute = iso();
  const days = tageMerk;
  const t = days.length && days[days.length - 1].date === heute ? days[days.length - 1] : null;
  const b = await baenderLaden();
  // Last 14 Tage mit Rahmen je Tag
  const l14: any = { von: null as string | null, ist: [] as number[], lo: [] as number[], hi: [] as number[] };
  const ab = Math.max(1, days.length - 28);   // 24.11.2026: 28 Tage, Uhr kürzt nach Einstellung
  for (let i = ab; i < days.length; i++) {
    const d = days[i];
    const e = lastEmpfehlung(d, days[i - 1], b);
    if (!l14.von) l14.von = d.date;
    l14.ist.push(Math.round(d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0));
    l14.lo.push(e ? e.min : -1);
    l14.hi.push(e ? e.max : -1);
  }
  // Morgenwerte 7 Tage
  const m = days.slice(-14);   // 24.11.2026: 14 Tage, Uhr kürzt nach Einstellung
  const m7 = {
    rd: m.map((d) => d.readiness ?? -1), hrv: m.map((d) => d.hrv ?? -1), rhr: m.map((d) => d.rhr ?? -1), sl: m.map((d) => d.sleepScore ?? -1),
    // 23.11.2026: Bereitschaft im Garmin-Muster (R <25 · O <50 · G <75 · b <95 · v Optimal)
    rdS: m.map((d) => (d.readiness == null ? 'x' : d.readiness >= 95 ? 'v' : d.readiness >= 75 ? 'b' : d.readiness >= 50 ? 'G' : d.readiness >= 25 ? 'O' : 'R')).join(''), hrvS: m.map((d) => LV_HRV[d.levels.hrv] || 'x').join(''),
    rhrS: m.map((d) => LV_G[d.levels.rhr] || 'x').join(''), slS: m.map((d) => LV_G[d.levels.sleep] || 'x').join(''),
  };
  // Form: CTL/ATL 42 Tage, Ausdauerwert 90 Tage (jeder 3. Tag)
  let form: any = null;
  try {
    const r: any = await fetchChartData('100d' as any, ['fbCTL_obs', 'fbATL_obs', 'garminCTL', 'garminATL', 'garminEnduranceScore'] as any);
    const rows = (r?.data || []).filter((x: any) => String(x.date).slice(0, 10) <= heute);
    const reihe = (k: string, n: number) => rows.slice(-n).map((x: any) => { const v = num(x[k]); return v == null || v <= 0 ? -1 : Math.round(v); });
    const es = reihe('garminEnduranceScore', 90).filter((_: number, i: number, a: number[]) => (a.length - 1 - i) % 3 === 0);
    const reihe2 = (post: string, morgen: string, n: number) => rows.slice(-n).map((x: any) => { const v = neuWert(x, post, morgen); return v == null ? -1 : Math.round(v); });
    form = { ctl: reihe2('fbCTL_obs', 'garminCTL', 90), atl: reihe2('fbATL_obs', 'garminATL', 90), es };   // 24.11.2026: 90 Tage
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
  return { l14, m7, form, letzte, feld, plan7: plan7Merk };
}

// ─── 25.11.2026: zehn neue Uhr-Funktionen (Seiten 10–17, Tour-Rechner, Plan ändern) ───
const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const tagK = (d: string) => { const x = new Date(d + 'T12:00:00Z'); return `${WT[x.getUTCDay()]} ${d.slice(8, 10)}.${d.slice(5, 7)}.`; };
const plusTage = (d: string, n: number) => { const x = new Date(d + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); };
const r1 = (x: number) => Math.round(x * 10) / 10;
const istT = (d: RecoveryDay) => (d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0);

async function extraTeil(h: any, z: any, u: any) {
  const heute = iso();
  const days = tageMerk;
  const t = days.length && days[days.length - 1].date === heute ? days[days.length - 1] : null;
  const rahmen = h._rahmen;
  const out: any = {};
  let g: any = null;
  try { g = await fetchGarminTage(); } catch {}
  // 1 Vorausblick (Mac, lokal berechnet)
  let pg: any = null;
  try {
    const it = (await fetchExtras<any>('prognose', 1))[0];
    pg = it?.data || null;
    if (pg) {
      const fw = pg.fruehwarnung || {}, vk = pg.vorausblick || {};
      out.vb = {
        stufe: fw.stufe || null, prog: (() => { const t = String(fw.prognose || ''); const i = t.indexOf('Stellschraube'); return (i >= 0 ? t.slice(i) : t).slice(0, 260); })(), seit: String(fw.verlauf?.text || '').slice(0, 160),
        acwr: vk.acwr_max ? { w: vk.acwr_max.wert, d: tagK(vk.acwr_max.datum) } : null,
        qt: vk.qualitaetstag?.tag || null, tf: pg.tourfenster?.[0] ? { tag: pg.tourfenster[0].tag, s: pg.tourfenster[0].score } : null,
        ctl: vk.ctl_heute ?? null, ctl28: vk.ctl_28 ?? null, trend: vk.trend || null, um: String(pg.erstellt || '').slice(11, 16), am: String(it.datum || ''),
        sig: (fw.signale || []).slice(0, 4).map((x: any) => ({ n: x.name, s: x.schwere })),
      };
    }
  } catch {}
  // 2 Was geht heute noch? (Restbudget in Minuten je Sport, aus deinen Zonen-Raten)
  try {
    const ist = Math.max(t ? istT(t) : 0, liveSumme());
    const raten = u?.feld?.raten || {};
    if (rahmen && rahmen.max > 0) {
      const rest = Math.max(0, Math.round(rahmen.max - ist)), restH = Math.max(0, Math.round((rahmen.halte || 0) - ist));
      const opt = (lab: string, rate: number | undefined) => (rate && rate > 0 ? { s: lab, min: Math.round((rest / rate) * 60), minH: Math.round((restH / rate) * 60) } : null);
      out.rest = {
        ess: rest, halte: restH, ist: Math.round(ist), max: rahmen.max, min: rahmen.min,
        opts: [opt('Bike Z2', raten.rad?.[1]), opt('Bike Z1', raten.rad?.[0]), opt('Lauf Z2', raten.lauf?.[1]), opt('Wandern', raten.berg ? (raten.berg[0] + raten.berg[1]) / 2 : undefined)].filter(Boolean),
      };
    }
  } catch {}
  // 4 Zonenverteilung 7/28 Tage (Minuten je Zone aus Garmin)
  try {
    if (g) {
      const summe = (n: number) => {
        const z5 = [0, 0, 0, 0, 0];
        const ab = plusTage(heute, -(n - 1));
        for (const [d, tg] of g) if (d >= ab && d <= heute) for (const a of tg.akt || []) ((a.zonen_s || []) as number[]).slice(0, 5).forEach((v, k) => { z5[k] += v || 0; });
        const ges = z5.reduce((a, b) => a + b, 0);
        return { p: z5.map((v) => (ges > 0 ? Math.round((v / ges) * 100) : 0)), h: r1(ges / 3600) };
      };
      const a = summe(7), b = summe(28);
      out.zonen = { p7: a.p, h7: a.h, p28: b.p, h28: b.h, grau7: a.p[2], hart7: a.p[3] + a.p[4], grau28: b.p[2], hart28: b.p[3] + b.p[4] };
    }
  } catch {}
  // 5 Erholungstrend 14 Tage
  try {
    const xs = days.slice(-14);
    const schuld = xs.slice(-7).reduce((a, d) => a + (d.sleepH != null ? Math.max(0, 7.5 - d.sleepH) : 0), 0);
    out.erh = {
      sl: xs.map((d) => (d.sleepH != null ? Math.round(d.sleepH * 10) : -1)), sc: xs.map((d) => d.sleepScore ?? -1),
      hrv: xs.map((d) => d.hrv ?? -1), lo: t?.hrvLow ?? xs.map((d) => d.hrvLow).filter((v) => v != null).pop() ?? null, hi: t?.hrvHigh ?? xs.map((d) => d.hrvHigh).filter((v) => v != null).pop() ?? null,
      bb: xs.map((d) => { const v = g?.get(d.date)?.bb_aufwachen; return v != null ? Math.round(v) : -1; }), schuld: r1(schuld),
    };
  } catch {}
  // 6 Wochenbilanz Mo–So
  try {
    const mo = (() => { const d = new Date(heute + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); return d.toISOString().slice(0, 10); })();
    const nx: any[] = h.naechste || [];
    const tage = [];
    for (let i = 0; i < 7; i++) {
      const d = plusTage(mo, i);
      const r = days.find((x) => x.date === d);
      const p = nx.find((x) => x.d === d);
      tage.push({ t: WT[new Date(d + 'T12:00:00Z').getUTCDay()].slice(0, 1), ist: d <= heute && r ? Math.round(istT(r)) : -1, plan: d > heute && p ? Math.round(p.ess || 0) : d === heute ? Math.round(h.last?.planHeute ?? 0) : -1, heute: d === heute });
    }
    const halte = rahmen?.halte || null;
    const entl = nx.find((x) => halte && (x.ess || 0) <= halte * 0.4);
    const mono = t?.monotonyPost ?? t?.monotony ?? null;
    const sum7 = days.slice(-7).reduce((a, d) => a + istT(d), 0);
    out.wb = { tage, mono: mono != null ? r1(mono) : null, strain: mono != null ? Math.round(mono * sum7) : null, entl: entl ? tagK(entl.d) : null, ohne: (() => { let n = 0; for (let i = days.length - 1; i >= 0 && halte && istT(days[i]) > halte * 0.4; i--) n++; return n; })() };
  } catch {}
  // 7 Bereitschaft morgen: Prognose vom Mac + persönliche Erfahrung (Readiness-Änderung über Nacht je Lastverhältnis)
  try {
    const pts: { q: number; d1: number }[] = [];
    const all = alleTage.length ? alleTage : days;
    for (let i = 1; i + 1 < all.length; i++) {
      const d = all[i], n = all[i + 1];
      if (d.readiness == null || n.readiness == null) continue;
      const e = lastEmpfehlung(d, all[i - 1]);
      if (!e || !(e.halte > 0)) continue;
      pts.push({ q: istT(d) / e.halte, d1: n.readiness - d.readiness });
    }
    const bin = (q: number) => { const g2 = pts.filter((p) => Math.abs(p.q - q) <= 0.35); return g2.length >= 6 ? { d1: Math.round(g2.reduce((a, p) => a + p.d1, 0) / g2.length), n: g2.length } : null; };
    const rd = t?.readiness ?? null;
    const halte = rahmen?.halte || null;
    const szen = [] as any[];
    const kand = [[0, 'Ruhe'], [rahmen?.max, 'Obergrenze'], [h.last?.planHeute, 'Plan']] as [number | null | undefined, string][];
    for (const [ess, lab] of kand) {
      if (ess == null || !halte || rd == null) continue;
      if (szen.some((x) => x.ess === Math.round(ess))) continue;
      const b = bin(ess / halte);
      if (b) szen.push({ lab, ess: Math.round(ess), rd: Math.max(1, Math.min(100, rd + b.d1)), d1: b.d1, n: b.n });
    }
    const pt = pg?.vorausblick?.tage?.[1];
    out.bm = { heute: rd, prog: pt?.bereitschaft ?? null, progTag: pt?.datum ? tagK(pt.datum) : null, mae: pg?.vorausblick?.modell?.mae ?? null, szen };
  } catch {}
  // 9 Ziel-Countdown
  try {
    if (z) {
      const ctlR = (u?.form?.ctl || []) as number[];
      const ist = [] as number[];
      for (let i = ctlR.length - 1; i >= 0 && ist.length < 9; i -= 7) ist.unshift(ctlR[i]);
      const wo = (pg?.vorausblick?.wochen || []) as any[];
      const prog = wo.filter((w) => w.bis <= z.datum || w.von <= z.datum).map((w) => w.ctl_ende);
      out.zc = { name: String(z.name || '').slice(0, 30), tage: z.tage, tf: z.tourform, datum: tagK(z.datum), ist, prog, ctlEnde: prog.length ? prog[prog.length - 1] : null };
    }
  } catch {}
  // 10 Treffsicherheit 28 Tage: im Rahmen / darüber / darunter und Readiness am Folgemorgen
  try {
    const b = await baenderLaden();
    const grp: Record<string, number[]> = { drin: [], ueber: [], unter: [] };
    const xs = days.slice(-29);
    for (let i = 1; i < xs.length - 1; i++) {
      const d = xs[i], n = xs[i + 1];
      const e = lastEmpfehlung(d, xs[i - 1], b);
      if (!e || !d.vote) continue;
      const v = istT(d);
      const k = v > e.max ? 'ueber' : v < e.min ? 'unter' : 'drin';
      grp[k].push(n.readiness != null && d.readiness != null ? n.readiness - d.readiness : NaN);
    }
    const m = (a: number[]) => { const f = a.filter((x) => Number.isFinite(x)); return f.length ? Math.round(f.reduce((p, q) => p + q, 0) / f.length) : null; };
    const nGes = grp.drin.length + grp.ueber.length + grp.unter.length;
    out.tr = { n: nGes, drin: grp.drin.length, ueber: grp.ueber.length, unter: grp.unter.length, pDrin: nGes ? Math.round((grp.drin.length / nGes) * 100) : null, dDrin: m(grp.drin), dUeber: m(grp.ueber), dUnter: m(grp.unter) };
  } catch {}
  // 8 Plan ändern: die nächsten 7 Tage mit Sperre
  out.pl = (h.naechste || []).slice(0, 7).map((x: any, i: number) => ({ d: x.d, t: tagK(x.d), ess: Math.round(x.ess || 0), s: x.sport || '', z: x.zone || '', lock: !!sperreMerk[i] }));
  return out;
}

// 3 Tour-Rechner: node widget.mjs --tour '{"teile":[{"s":"gravel","km":17.3,"hm":650},…]}'
async function tourRechnen(eingabe: any) {
  const g = await fetchGarminTage();
  const modelle = lerneModelle(g);
  const akt: any[] = [];
  for (const [d, tg] of g) for (const a of tg.akt || []) akt.push({ d, ...a });
  const PROFIL: Record<string, string> = { gravel: 'bike_gravel', rad: 'bike_road', lauf: 'run', berg: 'hike' };
  const TYP: Record<string, RegExp> = { gravel: /gravel|mountain_bik/, rad: /road_biking|^cycling$|gravel|mountain_bik/, lauf: /running/, berg: /hiking|mountaineering/, abfahrt: /cycling|biking/ };
  const med = (x: number[]) => { const s = [...x].sort((a, b) => a - b); return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2; };
  const teile = [] as any[];
  for (const tl of (eingabe?.teile || []).slice(0, 6)) {
    const s = String(tl.s || ''), km = Math.max(0, Number(tl.km) || 0), hm = Math.max(0, Number(tl.hm) || 0);
    if (s === 'abfahrt') {
      const xs = akt.filter((a) => TYP.abfahrt.test(String(a.typ || '')) && a.km >= 5 && (a.hm || 0) / a.km < 5 && (a.last || 0) > 0 && (a.hf_avg || 200) < 118 && a.dauer_s > 0 && a.km / (a.dauer_s / 3600) >= 21);
      const pk = xs.length >= 2 ? med(xs.map((a) => a.last / a.km)) : 1.2;
      const ess = Math.round(pk * km);
      teile.push({ s, km, hm, ess, lo: Math.round(ess * 0.75), hi: Math.round(ess * 1.3), q: xs.length >= 2 ? `${xs.length} eigene Abfahrten` : 'Faustwert', ref: '' });
      continue;
    }
    const re = TYP[s] || /x^/;
    const vgl = akt.filter((a) => re.test(String(a.typ || '')) && (a.last || 0) > 0 && a.km > 0 && a.km >= km * 0.5 && a.km <= km * 1.8 && (hm < 100 ? (a.hm || 0) < 250 : (a.hm || 0) >= hm * 0.65 && (a.hm || 0) <= hm * 1.5));
    if (vgl.length >= 2) {
      const sk = vgl.map((a) => a.last * (hm >= 100 && a.hm > 50 ? 0.2 * (km / a.km) + 0.8 * (hm / a.hm) : km / a.km));
      const ess = Math.round(med(sk));
      const best = vgl.slice().sort((a, b) => Math.abs(a.hm - hm) + 20 * Math.abs(a.km - km) - (Math.abs(b.hm - hm) + 20 * Math.abs(b.km - km)))[0];
      teile.push({ s, km, hm, ess, lo: Math.round(Math.min(...sk)), hi: Math.round(Math.max(...sk)), q: `${vgl.length} eigene Touren`, ref: `${String(best.name || '').slice(0, 22)} ${best.last}` });
    } else {
      const m = modelle.find((x) => x.id === PROFIL[s]);
      const r = m ? vorwaerts(m, { km, hm }) : null;
      teile.push({ s, km, hm, ess: r ? Math.round(r.ess) : null, lo: r ? Math.round(r.lo) : null, hi: r ? Math.round(r.hi) : null, q: r ? 'ESS-Rechner' : 'keine Daten', ref: '' });
    }
  }
  const ok = teile.filter((x) => x.ess != null);
  return { ok: true, ess: ok.reduce((a, x) => a + x.ess, 0), lo: ok.reduce((a, x) => a + x.lo, 0), hi: ok.reduce((a, x) => a + x.hi, 0), teile };
}

if (process.argv[2] === '--tour') {
  try { process.stdout.write(JSON.stringify(await tourRechnen(JSON.parse(process.argv[3] || '{}')))); }
  catch (e: any) { process.stdout.write(JSON.stringify({ ok: false, error: String(e?.message || e) })); }
} else {
const [h, z, f] = await Promise.all([heuteTeil(), zielTeil().catch(() => null), formTeil().catch(() => null)]);
let u: any = null;
try { u = await uhrTeil((h as any)._rahmen ?? null); } catch {}
try { if (u) u.x = await extraTeil(h, z, u); } catch {}
const { _rahmen: _r, ...hRest } = h as any;
process.stdout.write(JSON.stringify({ ok: true, stand: new Date().toISOString(), ...hRest, ziel: z, form: f, uhr: u }));
}
