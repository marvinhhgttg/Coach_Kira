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

export type LangTag = { date: string; h: number; sport: string | null; name: string | null; /** Anteil Z1+Z2 an der Zonenzeit (0–1) */ z12: number | null };
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
      const z = a.zonen_s || [];
      const zs = z.reduce((x, y) => x + (y || 0), 0);
      const z12 = z.length >= 2 && zs > 0 ? ((z[0] || 0) + (z[1] || 0)) / zs : null;
      if (h > 0 && (!best || h > best.h)) best = { date, h, sport: a.sport || a.typ || null, name: a.name || null, z12 };
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
export type LangEinstellung = { tage: number[]; sehrProMonat: number };
export const LANG_STANDARD: LangEinstellung = { tage: [6], sehrProMonat: 2 };
export const WT_KURZ = ['', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
/** Einstellung aus den Geräte-Einstellungen: ck_lang_tage "6" bzw. "3,6" (1 = Mo … 7 = So, leer = keine), ck_lang_sehr "0"–"4". */
export function parseLangEinstellung(tage?: string | null, sehr?: string | null): LangEinstellung {
  const t = tage == null ? LANG_STANDARD.tage : String(tage).split(',').map(Number).filter((x) => x >= 1 && x <= 7);
  const n = sehr == null || sehr === '' ? LANG_STANDARD.sehrProMonat : Math.max(0, Math.min(4, Math.round(Number(sehr)) || 0));
  return { tage: [...new Set(t)].sort((a, b) => a - b), sehrProMonat: n };
}
const wochentag = (iso: string) => ((new Date(iso + 'T00:00:00Z').getUTCDay() + 6) % 7) + 1; // 1 = Mo
const montag = (iso: string) => addTage(iso, 1 - wochentag(iso));

export function langEmpfehlung(s: LangStatus, e: LangEinstellung = LANG_STANDARD): LangEmpfehlung {
  const wocheDauerH = Math.min(4, Math.max(2, Math.round((s.laengste28h + 0.5) * 2) / 2));
  const proWoche = e.tage.length;
  const wocheFaellig = s.letzte2 ? addTage(s.letzte2.date, proWoche > 1 ? Math.max(3, Math.round(7 / proWoche)) : 7) : s.heute;
  const abstandSehr = e.sehrProMonat > 0 ? Math.max(7, Math.round(30 / e.sehrProMonat)) : 9999;
  const sehrLangFaellig = s.letzte4 ? addTage(s.letzte4.date, Math.min(abstandSehr, 365)) : s.heute;
  const tageTxt = proWoche ? e.tage.map((x) => WT_KURZ[x]).join(', ') : 'keine Tage gewählt';
  const fenster14Leer = !s.laengste14 || s.laengste14.h < 2;
  const sehrLangAufbau = s.laengste28h < 3;
  const sehrTxt = e.sehrProMonat === 0
    ? 'Sehr lange Einheiten ab 4 h sind abgewählt – sie bringen bei dir am meisten zusätzlich.'
    : sehrLangAufbau
      ? `Sobald die längste Einheit bei etwa 3 h liegt: ${e.sehrProMonat}-mal im Monat eine sehr lange ab 4 h (etwa alle ${abstandSehr} Tage). Ruhige Wanderungen und Skitouren gehen auch früher.`
      : `${e.sehrProMonat}-mal im Monat eine sehr lange Einheit ab 4 h (Tour, lange Ausfahrt), etwa alle ${abstandSehr} Tage.`;
  const punkte: string[] = [
    proWoche
      ? `${proWoche === 1 ? 'Jede Woche 1 lange Einheit' : `Jede Woche ${proWoche} lange Einheiten`} (${tageTxt}), die längste etwa ${String(wocheDauerH).replace('.', ',')} h, weitere ab 2 h – ruhig in Z1–Z2 (Rad, Wandern, Skitour).`
      : 'Keine Wochentage für lange Einheiten gewählt – ohne lange Einheiten trägt nur die mittlere Last den Score.',
    sehrTxt,
    'Dauer langsam steigern: pro Woche höchstens 30 min länger als die längste der letzten 4 Wochen.',
    'Nur an Tagen mit Tagesvotum Train oder Easy. Bei Recover oder niedriger Readiness lieber verschieben als kürzer und härter fahren.',
  ];
  return {
    wocheDauerH,
    wocheFaellig,
    wocheUeberfaellig: wocheFaellig < s.heute,
    sehrLangProMonat: e.sehrProMonat,
    sehrLangFaellig,
    sehrLangUeberfaellig: e.sehrProMonat > 0 && !sehrLangAufbau && sehrLangFaellig < s.heute,
    fenster14Leer,
    sehrLangAufbau,
    punkte,
  };
}

/** ESS pro Stunde ruhiger langer Einheiten (ab 1,5 h, mindestens 85 % in Z1–Z2): Median und Spanne (25–75 %) je Sport. */
export type EssRate = { med: number; p25: number; p75: number; n: number };
export function essProStunde(t: P3Tag[], lang: Map<string, LangTag>): { gesamt: number | null; sport: Record<string, number>; spanne: Record<string, EssRate> } {
  const proSport: Record<string, number[]> = {};
  const alle: number[] = [];
  for (const d of t) {
    const l = lang.get(d.date);
    if (!l || l.h < 1.5 || !d.load || d.load <= 0) continue;
    const skitour = sportDe(l.sport) === 'Skitour';
    if (!skitour && l.z12 != null && l.z12 < 0.85) continue; // nur ruhige Einheiten (Skitouren sind bergauf nie ganz ruhig)
    const v = d.load / l.h;
    if (v < 10 && sportDe(l.sport) === 'Rad') continue; // Aufzeichnung ohne Puls o. Ä.
    alle.push(v);
    (proSport[sportDe(l.sport)] ||= []).push(v);
  }
  const q = (a: number[], p: number) => { const b = [...a].sort((x, y) => x - y); return b.length ? b[Math.min(b.length - 1, Math.floor(p * (b.length - 1) + 0.5))] : 0; };
  const sport: Record<string, number> = {};
  const spanne: Record<string, EssRate> = {};
  for (const [k, v] of Object.entries(proSport)) {
    if (v.length >= 3) spanne[k] = { med: Math.round(q(v, 0.5)), p25: Math.round(q(v, 0.25)), p75: Math.round(q(v, 0.75)), n: v.length };
    else if (k === 'Skitour' && v.length >= 1) { const m = q(v, 0.5); spanne[k] = { med: Math.round(m), p25: Math.round(m * 0.8), p75: Math.round(m * 1.2), n: v.length }; } // wenige Skitouren: ±20 %
    else continue;
    sport[k] = spanne[k].med;
  }
  return { gesamt: alle.length ? Math.round(q(alle, 0.5)) : null, sport, spanne };
}

/** Wirkung der Last selbst auf den Score (Fortschreibungsmodell ES_MODEL, Höchstwert gegenüber einem Ruhetag) – etwa 0,17 Punkte je ESS. */
export function lastWirkung(ess: number): number {
  let e = 0, mx = 0;
  for (let k = 0; k < 70; k++) {
    e = ES_MODEL.b * e + (k === 0 ? ES_MODEL.c * ess : 0) + (k >= 1 && k <= ES_MODEL.win ? (ES_MODEL.d * ess) / ES_MODEL.win : 0);
    if (e > mx) mx = e;
  }
  return mx;
}
/** Gesamter erwarteter Einfluss einer langen Einheit: Dauer-Zuschlag aus der Historie + Wirkung der Last. */
export function esGesamt(ess: number, h: number, eff: LangEffekt): { gesamt: number; dauer: number; last: number } {
  const dauer = effektFuer(h, eff);
  const last = Math.round(lastWirkung(Math.max(0, ess)));
  return { gesamt: dauer + last, dauer, last };
}
/** ESS-Spanne für eine Dauer und Sportart (aus essProStunde); null, wenn zu wenig Daten. */
export function essSpanne(h: number, sport: string | null | undefined, ep: { spanne: Record<string, EssRate> }): { von: number; bis: number; richt: number } | null {
  const r = ep.spanne[sportDe(sport)];
  if (!r) return null;
  return { von: Math.round(r.p25 * h), bis: Math.round(r.p75 * h), richt: Math.round(r.med * h) };
}

export type LangSlot = {
  date: string; wt: string; art: 'lang' | 'sehr_lang' | 'ziel' | 'taper';
  /** Sport-Hinweis aus der Zieltour (z. B. Skitour) */
  sport?: string | null;
  hinweis?: string | null;
  dauerH: number; vonH: number; bisH: number;
  /** Dauer-Zuschlag aus der Historie (Punkte) */
  esEffekt: number;
  /** ESS-Spanne Rad für den Richtwert (falls genug Daten) und Gesamtwirkung beim Richtwert-ESS */
  essRad?: { von: number; bis: number; richt: number } | null;
  esGes?: number | null;
};

/** Zusatzwirkung (Score-Punkte) je Dauer, vorsichtig geschrumpft wie bonus4. */
export function effektFuer(h: number, eff: LangEffekt): number {
  const k = eff.klassen.find((x) => h >= x.von && h < x.bis);
  if (!k || k.extra == null || k.extra <= 0) return 0;
  return Math.round((k.extra * k.n) / (k.n + 3));
}

/** Zieltour (aus dem Zielplaner) – nur was der Hebel braucht. */
export type ZielKurz = { datum: string; name: string; art: string; dauer_h: number; hm: number; distanz_km: number; prioritaet: string };
export type ZielBedarf = {
  ziel: ZielKurz; wochen: number;
  /** so lang sollte die längste Einheit vor der Tour werden (70 % der Tourdauer, mindestens 2 h) */
  zielH: number;
  aufbauBis: string; taperAb: string; sport: string;
  /** erlaubte Steigerung pro Woche (Rad/Lauf 30 min, Gehen/Skitour 1 h) */
  schrittH: number; noetigProWoche: number; machbar: boolean; hmZiel: number;
};
const ZIEL_SPORT: Record<string, string> = { skitour: 'Skitour', bergtour: 'Bergtour', marsch: 'Wandern', bike_hike: 'Rad oder Wandern', rad: 'Rad', lauf: 'Lauf' };
const zielArtKurz = (z: ZielKurz) => (z.art === 'sonstiges' && /marsch|wander|walk|hike|trekking|pilger/i.test(z.name) ? 'marsch' : z.art);
/** Bedarf je kommender Zieltour (bis 30 Wochen voraus), nächste zuerst. */
export function zielBedarf(ziele: ZielKurz[], s: LangStatus): ZielBedarf[] {
  return ziele
    .filter((z) => z && z.datum >= s.heute && (z.dauer_h || 0) > 0 && diffTage(s.heute, z.datum) <= 210)
    .sort((a, b) => (a.datum < b.datum ? -1 : 1))
    .map((z) => {
      const art = zielArtKurz(z);
      const geh = art === 'marsch' || art === 'skitour' || art === 'bergtour';
      const zielH = Math.max(2, Math.round(z.dauer_h * 0.7 * 4) / 4);
      const aufbauBis = addTage(z.datum, art === 'marsch' ? -21 : -14);
      const taperAb = addTage(z.datum, art === 'marsch' ? -14 : -10);
      const schrittH = geh ? 1 : 0.5;
      const wochenAufbau = Math.max(1, Math.round(diffTage(s.heute, aufbauBis) / 7));
      const noetigProWoche = Math.max(0, (zielH - s.laengste28h) / wochenAufbau);
      return {
        ziel: z, wochen: Math.max(0, Math.round(diffTage(s.heute, z.datum) / 7)), zielH, aufbauBis, taperAb,
        sport: ZIEL_SPORT[art] || 'Rad', schrittH, noetigProWoche: Math.round(noetigProWoche * 100) / 100,
        machbar: diffTage(s.heute, aufbauBis) >= 0 && noetigProWoche <= schrittH + 1e-9, hmZiel: Math.round((z.hm || 0) * 0.7),
      };
    });
}

/**
 * Termine im Planfenster (ab heute, tageVoraus Tage) an den gewählten Wochentagen.
 * Je Woche ist der letzte gewählte Tag der Haupttag (Dauer wächst um 30 min pro Woche, ab der Empfehlung),
 * weitere gewählte Tage 2 h. Ein Haupttag wird sehr lang (4–5 h), wenn die nächste sehr lange fällig ist
 * und der Aufbau (längste ≥ 3 h) erreicht ist.
 */
export function langPlan(e: LangEinstellung, s: LangStatus, eff: LangEffekt, tageVoraus = 14, ep?: { spanne: Record<string, EssRate> }, bedarf: ZielBedarf[] = []): LangSlot[] {
  const emp = langEmpfehlung(s, e);
  const out: LangSlot[] = [];
  if (!e.tage.length) return out;
  const haupt = Math.max(...e.tage);
  const abstand = e.sehrProMonat > 0 ? Math.max(7, Math.round(30 / e.sehrProMonat)) : 0;
  let naechsteSehr = e.sehrProMonat > 0 && !emp.sehrLangAufbau ? emp.sehrLangFaellig : null;
  const ersteWoche = montag(s.heute);
  const zb = bedarf.find((b) => b.ziel.datum >= s.heute) || null; // nächste Zieltour steuert Länge, Sport und Tapering
  const esFuer = (h: number, sport: string | null) => {
    const roh = sport === 'Skitour' ? ['Ski', 'Hike'] : sport === 'Bergtour' || sport === 'Wandern' ? ['Hike'] : sport === 'Lauf' ? ['Run'] : ['Bike'];
    const sp = ep ? roh.map((r) => essSpanne(h, r, ep)).find((x) => x) || null : null;
    return { essRad: sp, esGes: sp ? esGesamt(sp.richt, h, eff).gesamt : null };
  };
  for (let i = 0; i < tageVoraus; i++) {
    const d = addTage(s.heute, i);
    const wt = wochentag(d);
    if (zb && d === zb.ziel.datum) {
      const h = zb.ziel.dauer_h;
      out.push({ date: d, wt: WT_KURZ[wt], art: 'ziel', sport: zb.sport, hinweis: zb.ziel.name, dauerH: h, vonH: h, bisH: h, esEffekt: effektFuer(h, eff), ...esFuer(h, zb.sport) });
      continue;
    }
    if (!e.tage.includes(wt)) continue;
    const woche = Math.round(diffTage(ersteWoche, montag(d)) / 7);
    let art: LangSlot['art'] = 'lang';
    let dauerH = 2;
    let sport: string | null = null;
    let hinweis: string | null = null;
    if (zb && d >= zb.taperAb && d < zb.ziel.datum) {
      art = 'taper';
      dauerH = Math.min(2, Math.max(1.5, zb.zielH * 0.5));
      sport = zb.sport;
      hinweis = `Tapering vor ${zb.ziel.name}`;
    } else if (zb && d > zb.aufbauBis && d < zb.taperAb && wt === haupt) {
      // zwischen Aufbauende und Tapering: Form halten, nicht weiter steigern
      dauerH = Math.max(2, Math.round(zb.zielH * 0.8 * 4) / 4);
      sport = zb.sport;
      hinweis = `Form halten vor ${zb.ziel.name}`;
    } else if (wt === haupt) {
      const schritt = zb && d <= zb.aufbauBis ? zb.schrittH : 0.5;
      // im Aufbau für eine Zieltour ist deren Ziellänge der Deckel (nicht darüber hinaus steigern)
      const deckel = zb && d <= zb.aufbauBis ? zb.zielH : emp.sehrLangAufbau ? 3.5 : 4;
      const start = Math.min(deckel, Math.max(2, Math.round((s.laengste28h + schritt) * 4) / 4));
      dauerH = Math.min(deckel, start + schritt * woche);
      if (zb && d <= zb.aufbauBis) {
        sport = zb.sport;
        hinweis = `Aufbau für ${zb.ziel.name} (Ziel ${String(zb.zielH).replace('.', ',')} h${zb.hmZiel ? `, ${zb.hmZiel} Hm` : ''})${zb.sport === 'Skitour' ? ' – ohne Schnee: Bergtour mit Höhenmetern' : ''}`;
      }
      if (naechsteSehr && d >= naechsteSehr && !(zb && d <= zb.ziel.datum && diffTage(d, zb.ziel.datum) <= 28)) {
        art = 'sehr_lang';
        dauerH = Math.max(4.5, dauerH);
        naechsteSehr = addTage(d, abstand);
      }
    }
    const vonH = art === 'sehr_lang' ? 4 : Math.max(art === 'taper' ? 1 : 2, dauerH - 0.5);
    const bisH = art === 'sehr_lang' ? Math.max(5.5, dauerH + 0.5) : dauerH + 0.5;
    out.push({ date: d, wt: WT_KURZ[wt], art, sport, hinweis, dauerH, vonH, bisH, esEffekt: effektFuer(dauerH, eff), ...esFuer(dauerH, sport) });
  }
  return out;
}

/** Datenpaket für das Plan-Briefing (Kira legt Dauer und Einfluss fest). */
export type LangKontext = {
  einstellung: LangEinstellung; status: LangStatus; effekt: LangEffekt; ziele?: ZielBedarf[];
  essProStunde: { gesamt: number | null; sport: Record<string, number>; spanne: Record<string, EssRate> }; termine: LangSlot[];
};
export function langKontext(t: P3Tag[], g: GarminTage, e: LangEinstellung, heute: string, ziele: ZielKurz[] = []): LangKontext {
  const lang = laengsteProTag(g);
  const effekt = langEffekt(t, lang);
  const status = langStatus(lang, heute);
  const ep = essProStunde(t, lang);
  const zb = zielBedarf(ziele, status);
  return { einstellung: e, status, effekt, ziele: zb, essProStunde: ep, termine: langPlan(e, status, effekt, 14, ep, zb) };
}
