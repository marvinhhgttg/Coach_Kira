/** KI-Extras: Speicher AI_KIRA_EXTRAS, Kalender-Belegung, phys_baseline, Wetter (Open-Meteo). */
import { BASE_URL } from './api';

export type ExtraTyp = 'treffer_tag' | 'treffer_woche' | 'selbsttest' | 'plausi' | 'ernaehrung' | 'alltag' | 'workout' | 'ziel' | 'saison' | 'autopilot' | 'umplanung' | 'lastfaktoren' | 'prognose' | 'ereignis' | 'garmin_api' | 'garmin_tag' | 'strategie' | 'zeitfenster';

export interface ExtraItem<T = any> {
  datum: string;
  typ: ExtraTyp;
  text: string;
  data: T | null;
  quelle: string;
  gespeichert: string;
}

async function getJson(url: string): Promise<any> {
  const res = await fetch(url, { redirect: 'follow', cache: 'no-store' });
  const t = await res.text();
  try {
    return JSON.parse(t);
  } catch {
    throw new Error('Apps Script antwortet nicht mit JSON (Patch „KI-Extras“ eingespielt?)');
  }
}

export async function fetchExtras<T = any>(typ: ExtraTyp | ExtraTyp[], n = 30): Promise<ExtraItem<T>[]> {
  const t = Array.isArray(typ) ? typ.join(',') : typ;
  const j = await getJson(`${BASE_URL}?mode=extras&typ=${encodeURIComponent(t)}&n=${n}&cb=${Date.now()}`);
  if (!j || !j.ok) throw new Error((j && j.error) || 'Extras nicht lesbar');
  return (j.items || []) as ExtraItem<T>[];
}

export async function saveExtra(typ: ExtraTyp, datum: string, text: string, data: unknown, quelle = 'cockpit:lokal'): Promise<void> {
  const res = await fetch(`${BASE_URL}?mode=saveExtra&cb=${Date.now()}`, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ kind: 'saveExtra', typ, datum, text, data, quelle }),
  });
  const t = await res.text();
  let j: any;
  try {
    j = JSON.parse(t);
  } catch {
    throw new Error('Speichern fehlgeschlagen (Patch „KI-Extras“ eingespielt?)');
  }
  if (!j.ok) throw new Error(j.error || 'Speichern fehlgeschlagen');
}

export interface CalendarDay {
  date: string;
  termine: number;
  ganztaegig: number;
  belegt_h: number;
  frueh: boolean;
  abends: boolean;
  von: string;
  bis: string;
  reise: boolean;
  frei: boolean;
  gesundheit: boolean;
}

export async function fetchCalendarDays(days = 14): Promise<CalendarDay[]> {
  return (await fetchCalendarDaysGefiltert(days)).days;
}

/** Alltag-Check: nur Marcs eigene Kalender, ganztägige Termine zählen nicht. */
export const ALLTAG_KALENDER = ['Marc', 'Marc (geteilt)', 'Natalia & Marc'];

export async function fetchCalendarDaysGefiltert(
  days = 14,
  opt: { kalender?: string[]; ohneGanztag?: boolean } = {},
): Promise<{ days: CalendarDay[]; kalender: string[] | null; filterAktiv: boolean }> {
  const q = new URLSearchParams({ mode: 'calendarDays', days: String(days), cb: String(Date.now()) });
  if (opt.kalender?.length) q.set('kal', opt.kalender.join('|'));
  if (opt.ohneGanztag) q.set('ganztag', '0');
  const j = await getJson(`${BASE_URL}?${q.toString()}`);
  if (!j || !j.ok) throw new Error((j && j.error) || 'Kalender nicht lesbar');
  const f = j.filter;
  let ds: CalendarDay[] = j.days || [];
  // Apps Script ohne Filter-Patch: ganztägige wenigstens aus der Zählung nehmen
  if (!f && opt.ohneGanztag) ds = ds.map((d) => ({ ...d, termine: Math.max(0, d.termine - (d.ganztaegig || 0)), ganztaegig: 0 }));
  return { days: ds, kalender: f && Array.isArray(f.kalender) ? f.kalender : null, filterAktiv: !!f };
}

export interface PhysBaseline {
  weightKg: number;
  maxHr: number;
  lthr: number;
  ftp: number;
  rhr: number;
  restingKcal: number;
}

const PHYS_FALLBACK: PhysBaseline = { weightKg: 76, maxHr: 187, lthr: 170, ftp: 256, rhr: 48, restingKcal: 2028 };
let physCache: PhysBaseline | null = null;

export async function fetchPhysBaseline(): Promise<PhysBaseline> {
  if (physCache) return physCache;
  try {
    const j = await getJson(`${BASE_URL}?mode=physBaseline&cb=${Date.now()}`);
    const b = (j && j.baseline) || {};
    const pick = (prefix: string, fb: number) => {
      const k = Object.keys(b).find((x) => x.toLowerCase().startsWith(prefix.toLowerCase()));
      const v = k ? Number(String(b[k]).replace(',', '.')) : NaN;
      return Number.isFinite(v) && v > 0 ? v : fb;
    };
    physCache = {
      weightKg: pick('Weight', PHYS_FALLBACK.weightKg),
      maxHr: pick('Max_HF', PHYS_FALLBACK.maxHr),
      lthr: pick('Laktatschwelle', PHYS_FALLBACK.lthr),
      ftp: pick('FTP', PHYS_FALLBACK.ftp),
      rhr: pick('RHR_default', PHYS_FALLBACK.rhr),
      restingKcal: pick('Resting_Calories', PHYS_FALLBACK.restingKcal),
    };
  } catch {
    physCache = { ...PHYS_FALLBACK };
  }
  return physCache;
}

export interface WeatherDay {
  date: string;
  tmax: number;
  tmin: number;
  regen_mm: number;
  regen_pct: number;
  wind_kmh: number;
  code: number;
  text: string;
}

const WMO: Record<number, string> = {
  0: 'klar', 1: 'überwiegend klar', 2: 'teils bewölkt', 3: 'bedeckt', 45: 'Nebel', 48: 'Reifnebel',
  51: 'Niesel', 53: 'Niesel', 55: 'starker Niesel', 61: 'leichter Regen', 63: 'Regen', 65: 'starker Regen',
  66: 'gefrierender Regen', 67: 'gefrierender Regen', 71: 'leichter Schnee', 73: 'Schnee', 75: 'starker Schnee',
  77: 'Schneegriesel', 80: 'Schauer', 81: 'Schauer', 82: 'heftige Schauer', 85: 'Schneeschauer', 86: 'Schneeschauer',
  95: 'Gewitter', 96: 'Gewitter mit Hagel', 99: 'Gewitter mit Hagel',
};

/** Wetter für Kadeltshofen (Open-Meteo, ohne Schlüssel) */
export async function fetchWeather(days = 14): Promise<WeatherDay[]> {
  const url =
    'https://api.open-meteo.com/v1/forecast?latitude=48.31&longitude=10.12' +
    '&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,wind_speed_10m_max,weather_code' +
    `&timezone=Europe%2FBerlin&forecast_days=${Math.min(16, days)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Wetter HTTP ${res.status}`);
  const j = await res.json();
  const d = j.daily || {};
  return (d.time || []).map((date: string, i: number) => ({
    date,
    tmax: Math.round(d.temperature_2m_max?.[i] ?? 0),
    tmin: Math.round(d.temperature_2m_min?.[i] ?? 0),
    regen_mm: Math.round((d.precipitation_sum?.[i] ?? 0) * 10) / 10,
    regen_pct: Math.round(d.precipitation_probability_max?.[i] ?? 0),
    wind_kmh: Math.round(d.wind_speed_10m_max?.[i] ?? 0),
    code: d.weather_code?.[i] ?? 0,
    text: WMO[d.weather_code?.[i] ?? 0] || '',
  }));
}

/** Herzfrequenz-Zonen nach Laktatschwelle (Friel-Schema, gerundet) */
export function hrZones(lthr: number, sport: 'run' | 'bike' | 'other' = 'run') {
  const f = sport === 'bike'
    ? [[0, 0.81], [0.81, 0.89], [0.9, 0.93], [0.94, 0.99], [1.0, 1.06]]
    : [[0, 0.85], [0.85, 0.89], [0.9, 0.94], [0.95, 0.99], [1.0, 1.06]];
  return f.map(([a, b], i) => ({ zone: `Z${i + 1}`, low: Math.round(lthr * a) || 100, high: Math.round(lthr * b) }));
}

/** Eintrag löschen (z. B. ein Ziel) – braucht den Patch „Zukunftsplanung“ */
export async function deleteExtra(typ: ExtraTyp, datum: string): Promise<void> {
  const res = await fetch(`${BASE_URL}?mode=saveExtra&cb=${Date.now()}`, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ kind: 'saveExtra', typ, datum, loeschen: true }),
  });
  const j = JSON.parse(await res.text());
  if (!j.ok) throw new Error(j.error || 'Löschen fehlgeschlagen');
}

export interface LoadFactors {
  load: { ActivityType: string; Zone: string; Factor: number }[];
  elev: { ActivityType: string; ElevFactor_per_100m: number }[];
}
export async function fetchLoadFactors(): Promise<LoadFactors> {
  const j = await getJson(`${BASE_URL}?mode=loadFactors&cb=${Date.now()}`);
  if (!j || !j.ok) throw new Error((j && j.error) || 'Lastfaktoren nicht lesbar');
  return { load: j.load || [], elev: j.elev || [] };
}
