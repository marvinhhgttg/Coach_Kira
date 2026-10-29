/**
 * Tourlast über den persönlichen ESS-Rechner (gelernte Modelle aus Garmin-Tagesdetails)
 * statt Faustregel. Fehlt ein passendes Modell (< 8 Einheiten), bleibt die Faustregel.
 */
import type { Ziel } from './planning';
import { vorwaerts, type Modell, type ProfilId } from './essRechner';

export interface TourTeil { profil: ProfilId; min: number; km?: number; hm: number; ess: number; lo: number; hi: number; n: number }
export interface TourEss { ess: number; lo: number; hi: number; teile: TourTeil[]; hinweise: string[] }

const MIN_N = 8;
const finde = (modelle: Modell[], ...ids: ProfilId[]) => {
  for (const id of ids) {
    const m = modelle.find((x) => x.id === id && x.n >= MIN_N);
    if (m) return m;
  }
  return null;
};

export function tourLastPersoenlich(z: Ziel, modelle: Modell[]): TourEss | null {
  if (!z || !(z.dauer_h > 0) || !modelle?.length) return null;
  const hinweise: string[] = [];
  const km = z.distanz_km > 0 ? z.distanz_km : undefined;
  const hm = Math.max(0, z.hm || 0);
  const min = z.dauer_h * 60;
  let plan: { m: Modell; min: number; km?: number; hm: number }[] = [];
  if (z.art === 'lauf') {
    const m = finde(modelle, 'run');
    if (!m) return null;
    plan = [{ m, min, km, hm }];
  } else if (z.art === 'rad') {
    const m = finde(modelle, 'bike_road', 'bike_gravel');
    if (!m) return null;
    plan = [{ m, min, km, hm }];
  } else if (z.art === 'bergtour' || z.art === 'skitour') {
    const m = finde(modelle, 'hike');
    if (!m) return null;
    if (z.art === 'skitour') hinweise.push('Skitour über dein Wandermodell geschätzt (eigene Skitouren fehlen).');
    plan = [{ m, min, km, hm }];
  } else if (z.art === 'bike_hike') {
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
