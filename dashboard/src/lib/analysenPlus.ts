/**
 * Weitere Analysen aus garmin_tag (+ Timeline, Kalender-Historie):
 * Hitze-Faustregel, Wind beim Rad, Entkopplung über die Saison, Einflüsse auf die Entkopplung,
 * Alltag und Erholung, Form bei Spitzenleistungen. Rein deterministisch.
 */
import type { GarminAkt, GarminTage } from './garmin';
import { BASE_URL } from './api';
import type { TlZeile } from './analysenErholung';
import {
  LAUF, RAD, RAD_ALLE, cohenD, eintraege, flachGeeignet, hart, median, minuten, mittel, plusTage, quantil,
  startMin, tageSeit, theilSen, vBewegt, wochentag, type E,
} from './stat';
import type { Sport } from './analysenTraining';

// ─── Hilfen ─────────────────────────────────────────────────────────────────
/** Kleinste Quadrate (Normalgleichungen, Gauß). X ohne Konstante, gibt [a, b1..bk] und R². */
export function ols(X: number[][], y: number[]): { beta: number[]; r2: number; n: number } | null {
  const n = y.length, k = (X[0]?.length || 0) + 1;
  if (n < k + 3) return null;
  const A = Array.from({ length: k }, () => new Array(k + 1).fill(0));
  for (let i = 0; i < n; i++) {
    const row = [1, ...X[i]];
    for (let r = 0; r < k; r++) { for (let c = 0; c < k; c++) A[r][c] += row[r] * row[c]; A[r][k] += row[r] * y[i]; }
  }
  for (let c = 0; c < k; c++) {
    let p = c;
    for (let r = c + 1; r < k; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
    if (Math.abs(A[p][c]) < 1e-9) return null;
    [A[c], A[p]] = [A[p], A[c]];
    for (let r = 0; r < k; r++) if (r !== c) { const f = A[r][c] / A[c][c]; for (let j = c; j <= k; j++) A[r][j] -= f * A[c][j]; }
  }
  const beta = A.map((r, i) => r[k] / r[i]);
  const ym = y.reduce((a, b) => a + b, 0) / n;
  let ssr = 0, sst = 0;
  for (let i = 0; i < n; i++) { const f = beta[0] + X[i].reduce((a, x, j) => a + x * beta[j + 1], 0); ssr += (y[i] - f) ** 2; sst += (y[i] - ym) ** 2; }
  return { beta, r2: sst > 0 ? 1 - ssr / sst : 0, n };
}
const hpk = (a: GarminAkt) => (a.hm || 0) / Math.max(a.km || 1, 1);
const indoor = (a: GarminAkt) => (a.typ || '').startsWith('indoor') || a.typ === 'virtual_ride';
/** Belastbarkeit einer Aussage nach Stichprobe. */
export const evidenz = (n: number, d?: number | null) =>
  n < 15 ? 'Tendenz (zu wenige Daten)' : d != null && Math.abs(d) < 0.2 ? 'kein klarer Effekt' : n < 40 ? 'Tendenz' : 'belastbar';

// ─── 1 Hitze-Faustregel ─────────────────────────────────────────────────────
/**
 * HF = a + b·Intensität + c·Wärme (°C über 15) + d·Kälte (°C unter 15) + e·Zeit.
 * Intensität: Laufen = km/h, Rad = Normalized Power (sonst km/h). Zeit fängt den Formtrend ab.
 */
export function hitzeRegel(tage: GarminTage | null, sport: Sport) {
  const es = eintraege(tage).filter((e) => flachGeeignet(e.a, sport) && e.a.wetter_c != null);
  const mitNp = sport === 'rad' && es.filter((e) => e.a.watt_np).length >= es.length * 0.7;
  const sel = mitNp ? es.filter((e) => e.a.watt_np) : es;
  const pts = sel.map((e) => ({
    datum: e.datum, temp: e.a.wetter_c!, feuchte: e.a.feuchte ?? null, hf: e.a.hf_avg!,
    x: mitNp ? e.a.watt_np! : vBewegt(e.a), name: e.a.name || '',
  }));
  const fit = ols(pts.map((p) => [p.x, Math.max(0, p.temp - 15), Math.min(0, p.temp - 15), tageSeit(p.datum) / 100]), pts.map((p) => p.hf));
  const warm = fit ? fit.beta[2] : null; // Schläge je °C über 15
  const kalt = fit ? fit.beta[3] : null; // Schläge je °C unter 15 (positiv = Kälte senkt HF)
  const proX = fit ? fit.beta[1] : null; // Schläge je km/h bzw. je W
  // Schwüle: Restfehler gegen Luftfeuchte an warmen Tagen (≥ 20 °C)
  let schwuele: number | null = null;
  if (fit) {
    const rs = pts.filter((p) => p.temp >= 20 && p.feuchte != null).map((p) => {
      const f = fit.beta[0] + fit.beta[1] * p.x + fit.beta[2] * Math.max(0, p.temp - 15) + fit.beta[3] * Math.min(0, p.temp - 15) + fit.beta[4] * tageSeit(p.datum) / 100;
      return { t: p.feuchte!, y: p.hf - f };
    });
    const ts = rs.length >= 8 ? theilSen(rs) : null;
    schwuele = ts ? ts.b * 20 : null; // Schläge je +20 % Luftfeuchte
  }
  const regel = [20, 25, 30].map((t) => {
    const mehr = warm != null ? warm * (t - 15) : null;
    // gleiche HF halten: Intensität um mehr/proX senken
    const weniger = mehr != null && proX && proX > 0 ? mehr / proX : null;
    return { temp: t, mehrSchlaege: mehr, wenigerX: weniger, n: pts.filter((p) => p.temp >= t - 2.5 && p.temp < t + 2.5).length };
  });
  const typX = median(pts.map((p) => p.x));
  return { n: pts.length, mitNp, warm, kalt, proX, r2: fit?.r2 ?? null, schwuele, regel, typX, pts, nWarm: pts.filter((p) => p.temp >= 20).length };
}

// ─── 2 Wind beim Rad ────────────────────────────────────────────────────────
/** Watt je km/h (NP/Tempo) relativ zum Formstand ±30 Tage, gegen Wind; danach windbereinigter Trend. */
export function windRad(tage: GarminTage | null) {
  const es = eintraege(tage).filter((e) => RAD.has(e.a.typ || '') && flachGeeignet(e.a, 'rad') && e.a.watt_np && e.a.wind_kmh != null);
  const pts = es.map((e) => ({ datum: e.datum, t: tageSeit(e.datum), wind: e.a.wind_kmh!, y: e.a.watt_np! / vBewegt(e.a), name: e.a.name || '', typ: e.a.typ || '' }));
  const rel = pts.map((p) => {
    const nb = pts.filter((q) => Math.abs(q.t - p.t) <= 30).map((q) => q.y);
    const m = median(nb) || p.y;
    return { ...p, relPct: ((p.y - m) / m) * 100 };
  });
  const tsWind = theilSen(rel.map((p) => ({ t: p.wind, y: p.relPct })));
  const proWind = tsWind ? tsWind.b : null; // % mehr Watt je km/h Wind
  const windMed = median(pts.map((p) => p.wind)) ?? 0;
  const korr = pts.map((p) => ({ ...p, yk: proWind != null ? p.y / (1 + (proWind * (p.wind - windMed)) / 100) : p.y }));
  const trRoh = theilSen(korr.map((p) => ({ t: p.t, y: p.y })));
  const trKorr = theilSen(korr.map((p) => ({ t: p.t, y: p.yk })));
  const basis = median(korr.map((p) => p.y)) || 1;
  // Streuung vorher/nachher (Robustheit): Abweichung vom Trend
  const streu = (k: 'y' | 'yk', tr: { a: number; b: number } | null) => tr ? median(korr.map((p) => Math.abs(p[k] - (tr.a + tr.b * p.t)) / basis * 100)) : null;
  const bins = [{ lab: 'bis 10 km/h', lo: 0, hi: 10 }, { lab: '10–15', lo: 10, hi: 15 }, { lab: '15–20', lo: 15, hi: 20 }, { lab: '20–25', lo: 20, hi: 25 }, { lab: 'über 25', lo: 25, hi: 200 }]
    .map((b) => { const xs = rel.filter((p) => p.wind >= b.lo && p.wind < b.hi); return { ...b, n: xs.length, relPct: median(xs.map((p) => p.relPct)) }; });
  return {
    n: pts.length, proWind, windMed, bins, punkte: korr,
    trendRohJe90: trRoh ? (trRoh.b * 90 / basis) * 100 : null, trendKorrJe90: trKorr ? (trKorr.b * 90 / basis) * 100 : null,
    streuRoh: streu('y', trRoh), streuKorr: streu('yk', trKorr),
  };
}

// ─── 3 Entkopplung über die Saison ──────────────────────────────────────────
export const DAUER_BINS = [{ lab: '60–90 min', lo: 60, hi: 90 }, { lab: '90–120 min', lo: 90, hi: 120 }, { lab: 'über 2 h', lo: 120, hi: 9999 }];
export interface EntPunkt { datum: string; a: GarminAkt; pct: number; min: number }
export function entPunkte(tage: GarminTage | null, sport: Sport): EntPunkt[] {
  return eintraege(tage)
    .filter((e) => (sport === 'lauf' ? LAUF : RAD_ALLE).has(e.a.typ || '') && minuten(e.a) >= 60 && e.a.ent && Number.isFinite(e.a.ent.pct))
    .filter((e) => hpk(e.a) <= (sport === 'lauf' ? 5 : 4) || indoor(e.a))
    .map((e) => ({ datum: e.datum, a: e.a, pct: e.a.ent!.pct, min: minuten(e.a) }));
}
const quartal = (d: string) => `${d.slice(0, 4)}-Q${Math.floor((Number(d.slice(5, 7)) - 1) / 3) + 1}`;
export function entSaison(tage: GarminTage | null, sport: Sport) {
  const ps = entPunkte(tage, sport);
  // Dauer und Zeit gemeinsam: pct = a + b·Stunden + c·Zeit
  const fit = ols(ps.map((p) => [p.min / 60, tageSeit(p.datum) / 90]), ps.map((p) => p.pct));
  const qs = [...new Set(ps.map((p) => quartal(p.datum)))].sort();
  const tabelle = qs.map((q) => {
    const xs = ps.filter((p) => quartal(p.datum) === q);
    return { q, lab: q.replace('-Q', ' Q'), n: xs.length, alle: median(xs.map((p) => p.pct)), b: DAUER_BINS.map((b) => median(xs.filter((p) => p.min >= b.lo && p.min < b.hi).map((p) => p.pct))) };
  });
  // Gleitender Median (8 Einheiten) für die Linie
  const linie = ps.map((p, i) => ({ ts: new Date(p.datum + 'T12:00:00').getTime(), y: p.pct, m: median(ps.slice(Math.max(0, i - 7), i + 1).map((q) => q.pct)), min: p.min, name: p.a.name }));
  return { n: ps.length, proStunde: fit ? fit.beta[1] : null, trendJe90: fit ? fit.beta[2] : null, tabelle, linie };
}

// ─── 4 Was die Entkopplung beeinflusst ──────────────────────────────────────
export interface EntFaktor { id: string; lab: string; niedrig: string; hoch: string; schwelle: number; n1: number; n2: number; m1: number | null; m2: number | null; d: number | null; einheit: string }
export function entFaktoren(tage: GarminTage | null, tl: Map<string, TlZeile> | null, sport: Sport | 'alle') {
  const ps = sport === 'alle' ? [...entPunkte(tage, 'lauf'), ...entPunkte(tage, 'rad')] : entPunkte(tage, sport);
  if (!tage) return { n: 0, faktoren: [] as EntFaktor[] };
  // Entkopplung je Sportart zentrieren (Rad und Laufen haben verschiedene Niveaus), Dauer herausrechnen
  const fitD = ols(ps.map((p) => [p.min / 60, LAUF.has(p.a.typ || '') ? 1 : 0]), ps.map((p) => p.pct));
  const rest = (p: EntPunkt) => fitD ? p.pct - (fitD.beta[0] + fitD.beta[1] * p.min / 60 + fitD.beta[2] * (LAUF.has(p.a.typ || '') ? 1 : 0)) : p.pct;
  const ds = [...tage.keys()].sort();
  const hrvRel = (d: string) => {
    const h = tage.get(d)?.hrv?.nacht;
    const b = median(ds.filter((x) => x < d && x >= plusTage(d, -28)).map((x) => tage.get(x)?.hrv?.nacht || NaN).filter(Number.isFinite));
    return h && b ? ((h - b) / b) * 100 : null;
  };
  const zeilen = ps.map((p) => {
    const g = tage.get(p.datum), v = tage.get(plusTage(p.datum, -1));
    const vorLast = [1, 2, 3].reduce((a, k) => a + (tage.get(plusTage(p.datum, -k))?.summe?.last || 0), 0);
    return {
      y: rest(p), pct: p.pct,
      x: {
        carb: tl?.get(plusTage(p.datum, -1))?.carb ?? null,
        kcal: tl?.get(plusTage(p.datum, -1))?.kcal_in ?? null,
        schlaf: g?.schlaf?.score ?? null,
        schlafh: g?.schlaf?.dauer_s ? g.schlaf.dauer_s / 3600 : null,
        hrv: hrvRel(p.datum),
        bb: g?.bb_aufwachen ?? null,
        temp: p.a.wetter_c ?? null,
        start: startMin(p.a) != null ? startMin(p.a)! / 60 : null,
        vorlast: vorLast,
        stress: v?.tag?.stress_avg ?? null,
      } as Record<string, number | null>,
    };
  });
  const DEF: { id: string; lab: string; niedrig: string; hoch: string; einheit: string; fest?: number }[] = [
    { id: 'carb', lab: 'Kohlenhydrate am Vortag', niedrig: 'wenig KH', hoch: 'viele KH', einheit: 'g' },
    { id: 'kcal', lab: 'Kalorien am Vortag', niedrig: 'wenig gegessen', hoch: 'viel gegessen', einheit: 'kcal' },
    { id: 'schlaf', lab: 'Schlafscore der Nacht davor', niedrig: 'schlechter Schlaf', hoch: 'guter Schlaf', einheit: '' },
    { id: 'schlafh', lab: 'Schlafdauer der Nacht davor', niedrig: 'kurz geschlafen', hoch: 'lang geschlafen', einheit: 'h' },
    { id: 'hrv', lab: 'HRV am Morgen (zur 28-Tage-Basis)', niedrig: 'HRV niedrig', hoch: 'HRV hoch', einheit: '%' },
    { id: 'bb', lab: 'Body Battery beim Aufwachen', niedrig: 'wenig Energie', hoch: 'viel Energie', einheit: '' },
    { id: 'temp', lab: 'Temperatur', niedrig: 'kühl', hoch: 'warm', einheit: '°C', fest: 18 },
    { id: 'start', lab: 'Startzeit', niedrig: 'vormittags', hoch: 'ab 14 Uhr', einheit: 'Uhr', fest: 14 },
    { id: 'vorlast', lab: 'Trainingslast der 3 Tage davor', niedrig: 'ausgeruht', hoch: 'vorbelastet', einheit: '' },
    { id: 'stress', lab: 'Stress am Vortag', niedrig: 'wenig Stress', hoch: 'viel Stress', einheit: '' },
  ];
  const faktoren: EntFaktor[] = [];
  for (const f of DEF) {
    const xs = zeilen.filter((z) => z.x[f.id] != null);
    if (xs.length < 10) continue;
    const s = f.fest ?? median(xs.map((z) => z.x[f.id]!))!;
    const lo = xs.filter((z) => z.x[f.id]! < s || (f.fest == null && z.x[f.id]! === s));
    const hi = xs.filter((z) => !lo.includes(z));
    if (lo.length < 4 || hi.length < 4) continue;
    faktoren.push({
      id: f.id, lab: f.lab, niedrig: f.niedrig, hoch: f.hoch, schwelle: s, einheit: f.einheit,
      n1: lo.length, n2: hi.length, m1: median(lo.map((z) => z.y)), m2: median(hi.map((z) => z.y)),
      d: cohenD(hi.map((z) => z.y), lo.map((z) => z.y)),
    });
  }
  faktoren.sort((a, b) => Math.abs(b.d ?? 0) - Math.abs(a.d ?? 0));
  return { n: ps.length, faktoren };
}

// ─── 5 Alltag und Erholung (Kalender-Historie) ──────────────────────────────
export interface KalTag { date: string; termine: number; belegt_h: number; abends: boolean; frueh: boolean }
/** Vergangene Kalendertage (nur Marcs drei Kalender, ohne ganztägige). Apps-Script-Modus „calendarHistorie“. */
export async function fetchKalenderHistorie(von: string, bis: string, kalender: string[]): Promise<KalTag[]> {
  const KEY = 'ck_kal_hist_v1';
  try {
    const c = JSON.parse(localStorage.getItem(KEY) || 'null');
    if (c && c.von === von && c.bis >= plusTage(bis, -1) && Date.now() - c.at < 12 * 3600000) return c.days;
  } catch { /* ignore */ }
  const q = new URLSearchParams({ mode: 'calendarHistorie', von, bis, kal: kalender.join('|'), ganztag: '0', cb: String(Date.now()) });
  let txt = '';
  try { txt = await (await fetch(`${BASE_URL}?${q.toString()}`, { redirect: 'follow', cache: 'no-store' })).text(); }
  catch { throw new Error('Die Kalender-Historie ist gerade nicht verfügbar'); } // unbekannter Modus → Google-Fehlerseite ohne CORS
  let j: any;
  try { j = JSON.parse(txt); } catch { throw new Error('Die Kalender-Historie ist gerade nicht verfügbar'); }
  if (!j?.ok || !Array.isArray(j.days)) throw new Error(j?.error && /mode|unbekannt|unknown/i.test(j.error) ? 'Die Kalender-Historie ist gerade nicht verfügbar' : (j?.error || 'Kalender nicht lesbar'));
  try { localStorage.setItem(KEY, JSON.stringify({ at: Date.now(), von, bis, days: j.days })); } catch { /* ignore */ }
  return j.days;
}
export const BELEGT_BINS = [{ lab: 'frei', lo: -1, hi: 0.01 }, { lab: 'bis 3 h', lo: 0.01, hi: 3 }, { lab: '3–6 h', lo: 3, hi: 6 }, { lab: '6–9 h', lo: 6, hi: 9 }, { lab: 'über 9 h', lo: 9, hi: 99 }];
export function alltagErholung(tage: GarminTage | null, kal: KalTag[] | null) {
  if (!tage || !kal) return null;
  const ds = [...tage.keys()].sort();
  const hrvRel = new Map<string, number>();
  for (const d of ds) {
    const h = tage.get(d)?.hrv?.nacht;
    const b = median(ds.filter((x) => x < d && x >= plusTage(d, -28)).map((x) => tage.get(x)?.hrv?.nacht || NaN).filter(Number.isFinite));
    if (h && b) hrvRel.set(d, ((h - b) / b) * 100);
  }
  // Kalendertag d → Nacht/Morgen d+1; Trainingslast am Tag d herausrechnen
  const zeilen = kal.map((k) => {
    const n = tage.get(plusTage(k.date, 1)), g = tage.get(k.date);
    return {
      ...k, wt: wochentag(k.date), last: g?.summe?.last || 0, stressTag: g?.tag?.stress_avg ?? null,
      hrv: hrvRel.get(plusTage(k.date, 1)) ?? null, schlaf: n?.schlaf?.score ?? null, bb: n?.bb_aufwachen ?? null,
      schlafh: n?.schlaf?.dauer_s ? n.schlaf.dauer_s / 3600 : null,
    };
  }).filter((z) => z.hrv != null || z.schlaf != null);
  const bereinigt = (key: 'hrv' | 'schlaf' | 'bb') => {
    const xs = zeilen.filter((z) => z[key] != null);
    const fit = ols(xs.map((z) => [z.belegt_h, z.last / 100, z.wt >= 5 ? 1 : 0]), xs.map((z) => z[key] as number));
    return fit ? { proStunde: fit.beta[1], n: fit.n } : null;
  };
  const bins = BELEGT_BINS.map((b) => {
    const xs = zeilen.filter((z) => z.belegt_h > b.lo && z.belegt_h <= b.hi || (b.lo < 0 && z.belegt_h === 0));
    return { lab: b.lab, n: xs.length, hrv: median(xs.map((z) => z.hrv!).filter((v) => v != null)), schlaf: median(xs.map((z) => z.schlaf!).filter((v) => v != null)), bb: median(xs.map((z) => z.bb!).filter((v) => v != null)), stress: median(xs.map((z) => z.stressTag!).filter((v) => v != null)) };
  }).filter((b) => b.n);
  const abends = zeilen.filter((z) => z.abends), nicht = zeilen.filter((z) => !z.abends && z.termine > 0);
  const nachAbend = {
    n1: abends.length, n2: nicht.length,
    hrv1: median(abends.map((z) => z.hrv!).filter((v) => v != null)), hrv2: median(nicht.map((z) => z.hrv!).filter((v) => v != null)),
    schlaf1: median(abends.map((z) => z.schlaf!).filter((v) => v != null)), schlaf2: median(nicht.map((z) => z.schlaf!).filter((v) => v != null)),
  };
  return { n: zeilen.length, von: kal[0]?.date, bins, hrv: bereinigt('hrv'), schlaf: bereinigt('schlaf'), bb: bereinigt('bb'), nachAbend, mitTerminen: zeilen.filter((z) => z.termine > 0).length };
}

// ─── 6 Form bei Spitzenleistungen ───────────────────────────────────────────
/** Leistungswert je Einheit, mit dem „Spitze“ bestimmt wird. */
const spitzenWert = (a: GarminAkt, sport: Sport): number | null => {
  // Rad: nur harte Einheiten, sonst bestimmt die Absicht (Pendeln, locker) den Wert statt der Form
  if (sport === 'rad') return RAD_ALLE.has(a.typ || '') && a.watt_20 && hart(a) ? a.watt_20 : null;
  // Laufen: Tempo je Herzschlag auf flachen Läufen ab 5 km (höher = besser), ab 20 min
  if (!LAUF.has(a.typ || '') || !flachGeeignet(a, 'lauf') || (a.km || 0) < 5) return null;
  return (vBewegt(a) / a.hf_avg!) * 1000;
};
export function formBeiSpitzen(tage: GarminTage | null, tl: Map<string, TlZeile> | null, sport: Sport) {
  if (!tage) return null;
  const es = eintraege(tage).map((e) => ({ ...e, w: spitzenWert(e.a, sport) })).filter((e): e is E & { w: number } => e.w != null);
  if (es.length < 10) return { n: es.length, spitzen: [], profil: null, basis: null } as const;
  // Spitze = obere 15 % relativ zum gleitenden 90-Tage-Median (sonst gewinnen nur neue Messwerte)
  const rel = es.map((e) => {
    const nb = es.filter((q) => Math.abs(tageSeit(q.datum) - tageSeit(e.datum)) <= 45).map((q) => q.w);
    const m = median(nb) || e.w;
    return { ...e, rel: ((e.w - m) / m) * 100 };
  });
  const grenze = quantil(rel.map((r) => r.rel), 0.85)!;
  const tagesLast = (d: string) => tage.get(d)?.summe?.last || 0;
  const profilFuer = (d: string) => {
    const l7 = [1, 2, 3, 4, 5, 6, 7].reduce((a, k) => a + tagesLast(plusTage(d, -k)), 0);
    const l28 = Array.from({ length: 28 }, (_, k) => tagesLast(plusTage(d, -(k + 8)))).reduce((a, b) => a + b, 0) / 4; // Wochen davor
    let ruhe = 0;
    for (let k = 1; k <= 14; k++) { const akts = tage.get(plusTage(d, -k))?.akt || []; if (akts.some((a) => hart(a))) break; ruhe++; }
    const t = tl?.get(plusTage(d, -1));
    const acwr = t?.atl && t?.ctl ? t.atl / t.ctl : null;
    return { l7rel: l28 > 0 ? ((l7 - l28) / l28) * 100 : null, ohneHart: ruhe, gestern: tagesLast(plusTage(d, -1)), acwr };
  };
  const mit = rel.map((r) => ({ ...r, p: profilFuer(r.datum) }));
  const sp = mit.filter((r) => r.rel >= grenze), rest = mit.filter((r) => r.rel < grenze);
  const zus = (xs: typeof mit) => ({
    n: xs.length,
    l7rel: median(xs.map((x) => x.p.l7rel!).filter((v) => v != null)),
    ohneHart: median(xs.map((x) => x.p.ohneHart)),
    gestern: median(xs.map((x) => x.p.gestern)),
    acwr: median(xs.map((x) => x.p.acwr!).filter((v) => v != null)),
    gesternFrei: xs.length ? (xs.filter((x) => x.p.gestern < 20).length / xs.length) * 100 : null,
  });
  return {
    n: es.length, grenze,
    spitzen: sp.slice(-12).reverse().map((r) => ({ datum: r.datum, name: r.a.name || '', w: r.w, rel: r.rel, ...r.p })),
    profil: zus(sp), basis: zus(rest),
  };
}
