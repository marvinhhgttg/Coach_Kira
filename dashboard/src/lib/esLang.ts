/**
 * Endurance Score und lange Einheiten (21.11.2026).
 * Garmin/Firstbeat: VO2max ist die Basis, dazu die Trainingshistorie aus zwei Fenstern (2–3 Monate und
 * die letzten 2 Wochen); die längsten Aktivitäten zählen besonders.
 * Hier: wie stark lange Einheiten bei Marc über ihre ESS-Last hinaus wirken (Residuum des Fortschreibungsmodells
 * ES_MODEL in den Tagen 0–3 nach der Einheit), dazu Status und konkrete Empfehlung. Reine Funktionen ohne Netz.
 */
import { ES_MODEL } from './ctlModel';
import type { P3Tag } from './paket3';
import type { GarminTage } from './garmin';

export type LangTag = { date: string; h: number; sport: string | null; name: string | null };
export type LangKlasse = { lab: string; von: number; bis: number; n: number; extra: number | null };
export type LangEffekt = { klassen: LangKlasse[]; bonus4: number; bonus4Roh: number | null; n4: number; basisTage: number };
export type LangStatus = {
  heute: string;
  laengste14: LangTag | null;
  laengste28h: number;
  n2_28: number;
  n2_84: number;
  n4_84: number;
  letzte2: LangTag | null;
  letzte4: LangTag | null;
};
export type LangEmpfehlung = {
  wocheDauerH: number;
  wocheFaellig: string;
  wocheUeberfaellig: boolean;
  sehrLangProMonat: number;
  sehrLangFaellig: string;
  sehrLangUeberfaellig: boolean;
  fenster14Leer: boolean;
  /** true, solange die längste Einheit der letzten 4 Wochen unter 3 h liegt: erst aufbauen, dann ab 4 h */
  sehrLangAufbau: boolean;
  punkte: string[];
};

const SPORT_DE: Record<string, string> = {
  Bike: 'Rad', Hike: 'Wandern', Ski: 'Skitour', Run: 'Lauf', Walk: 'Gehen', multi_sport: 'Mehrsport', Swim: 'Schwimmen',
};
export const sportDe = (s: string | null | undefined) => (s ? SPORT_DE[s] || s : '—');

const addTage = (iso: string, n: number) => {
  const d = new Date(iso + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};
const diffTage = (a: string, b: string) => Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);

/** Längste Aktivität je Tag (Dauer in Stunden, aus den Garmin-Tagesdetails). */
export function laengsteProTag(g: GarminTage): Map<string, LangTag> {
  const m = new Map<string, LangTag>();
  for (const [date, tag] of g) {
    let best: LangTag | null = null;
    for (const a of tag.akt || []) {
      const h = (a.dauer_s || a.bewegt_s || 0) / 3600;
      if (h > 0 && (!best || h > best.h)) best = { date, h, sport: a.sport || a.typ || null, name: a.name || null };
    }
    if (best) m.set(date, best);
  }
  return m;
}

/** Zusatzwirkung langer Einheiten über ihre ESS-Last hinaus (Punkte Endurance Score, Tage 0–3). */
export function langEffekt(t: P3Tag[], lang: Map<string, LangTag>): LangEffekt {
  const n = t.length;
  const res: (number | null)[] = new Array(n).fill(null);
  for (let i = 1; i < n; i++) {
    const e0 = t[i - 1].es, e1 = t[i].es;
    if (e0 == null || e1 == null) continue;
    let s = 0, k = 0;
    for (let m = Math.max(0, i - 56); m < i; m++) { s += t[m].load ?? 0; k++; }
    const m56 = k ? s / k : 0;
    res[i] = e1 - (ES_MODEL.a + ES_MODEL.b * e0 + ES_MODEL.c * (t[i].load ?? 0) + ES_MODEL.d * m56);
  }
  const kum = (i: number): number | null => {
    let s = 0, k = 0;
    for (let j = i; j <= i + 3 && j < n; j++) if (res[j] != null) { s += res[j] as number; k++; }
    return k >= 2 ? s : null;
  };
  const alle: number[] = [];
  const proTag: { h: number; v: number }[] = [];
  for (let i = 0; i < n; i++) {
    const v = kum(i);
    if (v == null) continue;
    alle.push(v);
    proTag.push({ h: lang.get(t[i].date)?.h ?? 0, v });
  }
  const basis = alle.length ? alle.reduce((a, b) => a + b, 0) / alle.length : 0;
  const def: [string, number, number][] = [['unter 2 h', 0.01, 2], ['2–4 h', 2, 4], ['ab 4 h', 4, 99]];
  const klassen = def.map(([lab, von, bis]) => {
    const v = proTag.filter((x) => x.h >= von && x.h < bis).map((x) => x.v);
    return { lab, von, bis, n: v.length, extra: v.length ? v.reduce((a, b) => a + b, 0) / v.length - basis : null };
  });
  const k4 = klassen[2];
  // Wenige Fälle: zur Null hin schrumpfen (n/(n+3)), negative Werte nicht als Bonus werten
  const bonus4 = k4.extra != null && k4.extra > 0 ? (k4.extra * k4.n) / (k4.n + 3) : 0;
  return { klassen, bonus4, bonus4Roh: k4.extra, n4: k4.n, basisTage: alle.length };
}

export function langStatus(lang: Map<string, LangTag>, heute: string): LangStatus {
  const tage = [...lang.values()].filter((x) => x.date <= heute).sort((a, b) => (a.date < b.date ? -1 : 1));
  const im = (n: number) => tage.filter((x) => diffTage(x.date, heute) < n);
  const l14 = im(14).reduce<LangTag | null>((b, x) => (!b || x.h > b.h ? x : b), null);
  const l28 = im(28).reduce((b, x) => Math.max(b, x.h), 0);
  const letzte = (h: number) => [...tage].reverse().find((x) => x.h >= h) || null;
  return {
    heute,
    laengste14: l14,
    laengste28h: l28,
    n2_28: im(28).filter((x) => x.h >= 2).length,
    n2_84: im(84).filter((x) => x.h >= 2).length,
    n4_84: im(84).filter((x) => x.h >= 4).length,
    letzte2: letzte(2),
    letzte4: letzte(4),
  };
}

/**
 * Empfehlung: 1 lange Einheit pro Woche (ab 2 h, Z1–Z2), dazu 2 sehr lange (ab 4 h) pro Monat im Abstand
 * von höchstens 14 Tagen, damit Garmins 2-Wochen-Fenster nie leer ist. Dauer der Wocheneinheit in kleinen
 * Schritten steigern (+30 min gegenüber der längsten der letzten 4 Wochen, 2–4 h).
 */
export function langEmpfehlung(s: LangStatus): LangEmpfehlung {
  const wocheDauerH = Math.min(4, Math.max(2, Math.round((s.laengste28h + 0.5) * 2) / 2));
  const wocheFaellig = s.letzte2 ? addTage(s.letzte2.date, 7) : s.heute;
  const sehrLangFaellig = s.letzte4 ? addTage(s.letzte4.date, 14) : s.heute;
  const fenster14Leer = !s.laengste14 || s.laengste14.h < 2;
  const sehrLangAufbau = s.laengste28h < 3;
  const punkte: string[] = [
    `Jede Woche 1 lange Einheit von etwa ${String(wocheDauerH).replace('.', ',')} h, ruhig in Z1–Z2 (Rad, Wandern, Skitour).`,
    sehrLangAufbau
      ? 'Sobald die Wocheneinheit bei etwa 3 h liegt: 2-mal im Monat eine sehr lange ab 4 h, höchstens 14 Tage auseinander. Ruhige Wanderungen und Skitouren gehen auch früher.'
      : '2-mal im Monat eine sehr lange Einheit ab 4 h (Tour, lange Ausfahrt), höchstens 14 Tage auseinander.',
    'Dauer langsam steigern: pro Woche höchstens 30 min länger als die längste der letzten 4 Wochen.',
    'Nur an Tagen mit Tagesvotum Train oder Easy. Bei Recover oder niedriger Readiness lieber verschieben als kürzer und härter fahren.',
  ];
  return {
    wocheDauerH,
    wocheFaellig,
    wocheUeberfaellig: wocheFaellig < s.heute,
    sehrLangProMonat: 2,
    sehrLangFaellig,
    sehrLangUeberfaellig: !sehrLangAufbau && sehrLangFaellig < s.heute,
    fenster14Leer,
    sehrLangAufbau,
    punkte,
  };
}
