/**
 * Tourlast über den persönlichen ESS-Rechner (gelernte Modelle aus Garmin-Tagesdetails)
 * statt Faustregel. Fehlt ein passendes Modell (< 8 Einheiten), bleibt die Faustregel.
 */
import { zielArt, type Ziel } from './planning';
import { vorwaerts, type Modell, type ProfilId } from './essRechner';
import type { GarminTage } from './garmin';
import SKI_API from '../data/skitouren_garmin.json';
import MARSCH_API from '../data/lange_maersche_garmin.json';

export interface TourTeil { profil: ProfilId; min: number; km?: number; hm: number; ess: number; lo: number; hi: number; n: number }
export interface TourEss { ess: number; lo: number; hi: number; teile: TourTeil[]; hinweise: string[] }

const MIN_N = 8;

/** Eigene Skitouren (Garmin backcountry_skiing) – Grundlage für Skitour-Schätzungen. */
export interface Skitour { datum: string; min: number; km?: number; hm: number; hf: number; ess: number; vergleichbar: boolean; name?: string }
/**
 * Eigene Skitouren: Garmin-API-Historie (2020–2025, einmalig abgerufen) + neue Touren aus den Garmin-Tagesdetails.
 * „vergleichbar“ = Garmin-Last auf demselben Maßstab wie heute (12/2022–12/2023 lag sie etwa doppelt so hoch).
 */
export function eigeneSkitouren(tage: GarminTage | null | undefined): Skitour[] {
  const m = new Map<string, Skitour>();
  for (const t of (SKI_API as any).touren as Skitour[]) m.set(t.datum + '|' + t.min, t);
  const daten = new Set([...m.values()].map((t) => t.datum));
  if (tage)
    for (const [datum, t] of tage)
      for (const a of t.akt || [])
        if (/backcountry_ski|skitour|ski_touring/i.test(String(a.typ || '')) && a.dauer_s && a.dauer_s >= 2700 && a.hf_avg && (a.last || 0) >= 40 && !daten.has(datum))
          m.set(datum + '|' + a.dauer_s, { datum, name: a.name, min: a.dauer_s / 60, km: a.km, hm: a.hm || 0, hf: a.hf_avg, ess: a.last!, vergleichbar: true });
  return [...m.values()].sort((a, b) => a.datum.localeCompare(b.datum));
}
/** Eigene Bergtouren mit der aktuellen Uhr (Garmin-Tagesdetails ab 26.08.25): ≥ 700 Hm und ≥ 2 h, Wandern/Bergsteigen. */
export interface Bergtour { datum: string; min: number; km?: number; hm: number; hf?: number; ess: number; name?: string }
export function eigeneBergtouren(tage: GarminTage | null | undefined): Bergtour[] {
  const out: Bergtour[] = [];
  if (!tage) return out;
  for (const [datum, t] of tage)
    for (const a of t.akt || [])
      if (/^(hiking|mountaineering)$/i.test(String(a.typ || '')) && (a.hm || 0) >= 700 && (a.dauer_s || 0) >= 7200 && (a.last || 0) > 0)
        out.push({ datum, name: a.name, min: a.dauer_s! / 60, km: a.km, hm: a.hm!, hf: a.hf_avg, ess: a.last! });
  return out.sort((a, b) => a.datum.localeCompare(b.datum));
}
/** Lange Märsche: Garmin-API-Historie + neue aus den Tagesdetails (≥ 4 h, höchstens 40 Hm/km, Gehen/Wandern). */
export interface Marsch { datum: string; min: number; km?: number; hm: number; hf?: number; ess: number; name?: string }
let tageCache: GarminTage | null = null;
/** Garmin-Tagesdetails für neue Märsche bereitstellen (vom Aufrufer, einmal je Laden). */
export function setzeGarminTage(t: GarminTage | null) { tageCache = t; }
export function eigeneMaersche(tage: GarminTage | null | undefined): Marsch[] {
  const m = new Map<string, Marsch>();
  for (const t of (MARSCH_API as any).touren as Marsch[]) m.set(t.datum, t);
  if (tage)
    for (const [datum, t] of tage)
      for (const a of t.akt || [])
        if (/^(hiking|walking|speed_walking|mountaineering)$/i.test(String(a.typ || '')) && (a.dauer_s || 0) >= 14400 && a.km && (a.hm || 0) / a.km <= 40 && a.last != null && !m.has(datum))
          m.set(datum, { datum, name: a.name, min: a.dauer_s! / 60, km: a.km, hm: a.hm || 0, hf: a.hf_avg, ess: a.last });
  return [...m.values()].sort((a, b) => a.datum.localeCompare(b.datum));
}
const median = (x: number[]) => {
  const s = [...x].sort((a, b) => a - b);
  return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
};
const de = (iso: string) => `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(2, 4)}`;
const finde = (modelle: Modell[], ...ids: ProfilId[]) => {
  for (const id of ids) {
    const m = modelle.find((x) => x.id === id && x.n >= MIN_N);
    if (m) return m;
  }
  return null;
};

export function tourLastPersoenlich(z: Ziel, modelle: Modell[], skitouren: Skitour[] = [], bergtouren: Bergtour[] = []): TourEss | null {
  skitouren = [...skitouren];
  if (!z || !(z.dauer_h > 0) || !modelle?.length) return null;
  const hinweise: string[] = [];
  const km = z.distanz_km > 0 ? z.distanz_km : undefined;
  const hm = Math.max(0, z.hm || 0);
  const min = z.dauer_h * 60;
  let plan: { m: Modell; min: number; km?: number; hm: number }[] = [];
  const art = zielArt(z);
  if (art === 'marsch') {
    // Garmin-Last beim langen Gehen: aus deinen eigenen langen Märschen (≥ 4 h, flach) ab 2024 – ESS je Stunde × Dauer
    const ms = eigeneMaersche(tageCache).filter((t) => t.datum >= '2024-01-01');
    if (ms.length >= 3) {
      const ph = ms.map((t) => t.ess / (t.min / 60)).sort((a, b) => a - b);
      const med = median(ph);
      const ess = Math.round(med * z.dauer_h), lo = Math.round(ph[0] * z.dauer_h), hi = Math.round(ph[ph.length - 1] * z.dauer_h);
      const lang = ms.reduce((a, t) => (t.min > a.min ? t : a), ms[0]);
      hinweise.push(
        `Marsch aus ${ms.length} eigenen langen Touren seit 2024 (≥ 4 h, flach): Garmin vergab ${ph[0].toLocaleString('de-DE', { maximumFractionDigits: 1 })}–${ph[ph.length - 1].toLocaleString('de-DE', { maximumFractionDigits: 1 })} ESS je Stunde – ` +
          `z. B. ${de(lang.datum)}: ${(Math.round(lang.min / 6) / 10).toLocaleString('de-DE')} h, ${Math.round(lang.km || 0)} km, ${lang.hm} Hm → ${lang.ess} ESS. ` +
          'Garmin misst die Last über die Intensität, nicht über die Dauer – für einen Marsch zählen deshalb Strecke und Zeit am Stück.',
      );
      return { ess, lo, hi, teile: [{ profil: 'hike', min: Math.round(min), km, hm: Math.round(hm), ess, lo, hi, n: ms.length }], hinweise };
    }
    const m = finde(modelle, 'hike');
    if (!m) return null;
    const e = vorwaerts(m, { min, km, hm });
    if (!e) return null;
    hinweise.push('Marsch über dein Wandermodell geschätzt – noch zu wenige eigene lange Märsche für einen direkten Abgleich.');
    return { ess: Math.round(e.ess), lo: Math.round(e.lo), hi: Math.round(e.hi), teile: [{ profil: 'hike', min: Math.round(min), km, hm: Math.round(hm), ess: Math.round(e.ess), lo: Math.round(e.lo), hi: Math.round(e.hi), n: m.n }], hinweise };
  }
  if (art === 'lauf') {
    const m = finde(modelle, 'run');
    if (!m) return null;
    plan = [{ m, min, km, hm }];
  } else if (art === 'rad') {
    const m = finde(modelle, 'bike_road', 'bike_gravel');
    if (!m) return null;
    plan = [{ m, min, km, hm }];
  } else if (art === 'bergtour' || art === 'skitour') {
    const m = finde(modelle, 'hike');
    if (!m) return null;
    if (art === 'skitour') {
      const alleSki = skitouren;
      skitouren = skitouren.filter((t) => t.vergleichbar);
      const ausgeschl = alleSki.length - skitouren.length;
      if (!skitouren.length) {
        hinweise.push('Skitour über dein Wandermodell geschätzt – eigene Skitouren fehlen, die Herzfrequenz beim Skitourengehen liegt meist höher (Schätzung eher zu niedrig).');
        plan = [{ m, min, km, hm }];
      } else {
        // Wandermodell mit deiner Skitour-Herzfrequenz, kalibriert an deinen echten Skitouren
        const hf = median(skitouren.map((t) => t.hf));
        const ratios = skitouren
          .map((t) => ({ t, e: vorwaerts(m, { min: t.min, km: t.km, hm: t.hm, hf: t.hf }) }))
          .filter((x) => x.e && x.e.ess > 0)
          .map((x) => x.t.ess / x.e!.ess);
        const k = ratios.length ? Math.min(1.5, Math.max(0.7, median(ratios))) : 1;
        const e = vorwaerts(m, { min, km, hm, hf });
        if (!e) return null;
        const f = (v: number) => Math.round(v * k);
        const jahre = `${skitouren[0].datum.slice(0, 4)}–${skitouren[skitouren.length - 1].datum.slice(0, 4)}`;
        const letzte = skitouren[skitouren.length - 1];
        hinweise.push(
          `Skitour mit deiner Skitour-Herzfrequenz (Ø HF ${Math.round(hf)}) gerechnet und an ${skitouren.length} eigene${skitouren.length === 1 ? 'r' : 'n'} Skitour${skitouren.length === 1 ? '' : 'en'} aus Garmin abgeglichen (${jahre}, Faktor ${k.toLocaleString('de-DE', { maximumFractionDigits: 2 })}; zuletzt ${de(letzte.datum)}: ${(Math.round(letzte.min / 6) / 10).toLocaleString('de-DE')} h, ${letzte.hm} Hm, ${letzte.ess} ESS).` +
            (ausgeschl ? ` ${ausgeschl} Touren von 12/2022–12/2023 nicht verwendet: Garmin-Last dort etwa doppelt so hoch, anderer Maßstab.` : '') +
            (skitouren.length < 3 ? ' Wenige eigene Skitouren – Schätzung wird mit jeder Tour genauer.' : ''),
        );
        const teile: TourTeil[] = [{ profil: 'hike', min: Math.round(min), km, hm: Math.round(hm), ess: f(e.ess), lo: f(e.lo), hi: f(e.hi), n: skitouren.length }];
        return { ess: f(e.ess), lo: f(e.lo), hi: f(e.hi), teile, hinweise };
      }
    } else {
      // Bergtour: Wandermodell, abgeglichen an deinen eigenen Bergtouren mit der aktuellen Uhr (ab 3 Touren)
      const e = vorwaerts(m, { min, km, hm });
      if (!e) return null;
      const ratios = bergtouren
        .map((t) => ({ t, p: vorwaerts(m, { min: t.min, km: t.km, hm: t.hm }) }))
        .filter((x) => x.p && x.p.ess > 0)
        .map((x) => x.t.ess / x.p!.ess);
      const k = ratios.length >= 3 ? Math.min(1.5, Math.max(0.7, median(ratios))) : 1;
      if (ratios.length >= 3) {
        const letzte = bergtouren[bergtouren.length - 1];
        hinweise.push(
          `Bergtour an ${ratios.length} eigenen Bergtouren seit 08/2025 abgeglichen (≥ 700 Hm, Faktor ${k.toLocaleString('de-DE', { maximumFractionDigits: 2 })}; zuletzt ${de(letzte.datum)}: ${(Math.round(letzte.min / 6) / 10).toLocaleString('de-DE')} h, ${letzte.hm} Hm, ${letzte.ess} ESS). ` +
            'Ältere Bergtouren (2020–07/2025) hatten bei gleicher Tour rund 30 % weniger Garmin-Last und zählen nicht.',
        );
      } else hinweise.push('Bergtour über dein Wandermodell – noch zu wenige eigene Bergtouren ab 700 Hm für einen Abgleich.');
      const f = (v: number) => Math.round(v * k);
      return { ess: f(e.ess), lo: f(e.lo), hi: f(e.hi), teile: [{ profil: 'hike', min: Math.round(min), km, hm: Math.round(hm), ess: f(e.ess), lo: f(e.lo), hi: f(e.hi), n: m.n }], hinweise: [...hinweise, ...e.hinweise.filter((h) => !hinweise.includes(h))] };
    }
  } else if (art === 'bike_hike') {
    const rad = finde(modelle, 'bike_gravel', 'bike_road');
    const geh = finde(modelle, 'hike');
    if (!rad || !geh) return null;
    const h = (v?: number) => (v == null ? undefined : v / 2);
    plan = [
      { m: rad, min: min / 2, km: h(km), hm: hm / 2 },
      { m: geh, min: min / 2, km: h(km), hm: hm / 2 },
    ];
    hinweise.push('Bike & Hike: je halbe Zeit und halbe Höhenmeter auf Rad und zu Fuß angenommen.');
  } else return null;

  const teile: TourTeil[] = [];
  for (const p of plan) {
    const e = vorwaerts(p.m, { min: p.min, km: p.km, hm: p.hm });
    if (!e || !Number.isFinite(e.ess)) return null;
    teile.push({ profil: p.m.id, min: Math.round(p.min), km: p.km, hm: Math.round(p.hm), ess: e.ess, lo: e.lo, hi: e.hi, n: p.m.n });
    for (const t of e.hinweise) if (!hinweise.includes(t)) hinweise.push(t);
  }
  const sum = (k: 'ess' | 'lo' | 'hi') => Math.round(teile.reduce((a, t) => a + t[k], 0));
  return { ess: sum('ess'), lo: sum('lo'), hi: sum('hi'), teile, hinweise };
}
