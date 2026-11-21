/**
 * Zielplan Tag für Tag (21.11.2026).
 * Ablauf: 1. Der Rechenkern legt den Rahmen fest (Einstiegsfenster, Tapering, Ziellänge der langen Einheit,
 * ESS pro Stunde, aktuelle Fitness). 2. Kira (lokal) entscheidet Einstieg, Phasen, Wochenlast und Wochenmuster.
 * 3. Der Kern rollt das Tag für Tag aus (Steigerung, Entlastungswochen, lange Einheit, Zieltag, Erholung)
 * und prüft ACWR und harte Nachbartage der langen Einheit. Reine Funktionen ohne Netz.
 */
import { atlStep, ctlStep, windowMean } from './ctlModel';
import type { Ziel } from './planning';
import type { EssRate, LangKontext } from './esLang';

export const ZP_SPORTARTEN = ['Rad', 'Lauf', 'Wandern', 'Bergtour', 'Skitour', 'Skike', 'Langlauf', 'Kraft', 'Mobility', 'Off'] as const;
export type ZpZone = 'Off' | 'Z1' | 'Z2' | 'Z3' | 'Z4' | 'Z5';
export const ZONE_INTENSITAET: Record<ZpZone, string> = { Off: 'Ruhe', Z1: 'sehr locker', Z2: 'locker', Z3: 'mittel', Z4: 'hart', Z5: 'sehr hart' };

export type ZpRahmen = {
  klasse: 'kurz' | 'mittel' | 'lang' | 'ultra';
  klasseText: string;
  /** Einstieg frühestens / spätestens (YYYY-MM-DD) und Vorschlag des Kerns */
  fruehestens: string; spaetestens: string; vorschlag: string; knapp: boolean;
  taperTage: number; taperAb: string;
  zielSport: string; zielH: number; schrittH: number; laengsteH: number;
  tourLast: number; wochenlastJetzt: number; ctlJetzt: number; atlJetzt: number;
  langWt: number; // 1 = Mo … 7 = So
  sportarten: string[];
  raten: Record<string, number>;
};
export type ZpPhase = { name: string; von: string; bis: string; wochen_ess: number; fokus: string; woche: ZpMuster[] };
export type ZpMuster = { wt: number; sport: string; zone: ZpZone; anteil: number; training: string; lang?: boolean };
export type ZpTag = {
  date: string; wt: number; woche: number; phase: string; sport: string; zone: ZpZone; intensitaet: string;
  ess: number; dauerH: number | null; training: string; art?: 'lang' | 'ziel' | 'taper' | 'erholung' | 'entlastung'; acwr?: number;
};
export type Zielplan = {
  ziel: { name: string; datum: string };
  erstellt: string; engine: string; einstieg: string; begruendung: string;
  rahmen: ZpRahmen; phasen: Omit<ZpPhase, 'woche'>[]; tage: ZpTag[];
  korrekturen: string[]; tapering: string; hinweise: string[]; maxAcwr: number; ctlStart: number; /** CTL am Vortag des Ziels */ ctlZiel: number;
};

const DAY = 86400000;
const d2 = (s: string) => new Date(s + 'T12:00:00');
export const addT = (s: string, n: number) => { const x = d2(s); x.setDate(x.getDate() + n); return x.toLocaleDateString('sv-SE'); };
export const diffT = (a: string, b: string) => Math.round((d2(b).getTime() - d2(a).getTime()) / DAY);
const wtVon = (s: string) => ((d2(s).getDay() + 6) % 7) + 1;
const montagVon = (s: string) => addT(s, 1 - wtVon(s));
const rund5 = (x: number) => Math.max(0, Math.round(x / 5) * 5);
export const WT_NAME = ['', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];

const ZIEL_SPORT: Record<string, string> = { skitour: 'Skitour', bergtour: 'Bergtour', marsch: 'Wandern', bike_hike: 'Bergtour', rad: 'Rad', lauf: 'Lauf' };
function zielSportVon(z: Ziel) {
  if (z.art === 'sonstiges') return /marsch|wander|walk|hike|trek|pilger/i.test(z.name) ? 'Wandern' : 'Rad';
  return ZIEL_SPORT[z.art] || 'Rad';
}
/** ESS pro Stunde (ruhig, Z1–Z2) je Sportart aus Marcs Historie; Bergtour = Wandern. */
function ratenAus(sp: Record<string, EssRate> | undefined): Record<string, number> {
  const r: Record<string, number> = { Rad: 50, Lauf: 75, Wandern: 10, Skitour: 50, Skike: 60, Langlauf: 60, Kraft: 30, Mobility: 10 };
  for (const k of ['Rad', 'Lauf', 'Wandern', 'Skitour']) if (sp?.[k]?.med) r[k] = sp[k].med;
  if (sp?.Skike?.med && (sp.Skike.n || 0) >= 8) r.Skike = sp.Skike.med;
  r.Bergtour = Math.max(r.Wandern, 12);
  return r;
}
const ZONE_FAKTOR: Record<ZpZone, number> = { Off: 0, Z1: 0.8, Z2: 1, Z3: 1.5, Z4: 2, Z5: 2.4 };

/** Rahmen aus Ziel, Fitness und Langzeitkontext. hist = Tageslasten (älteste zuerst, bis gestern). */
export function zielRahmen(z: Ziel, tourLast: number, hist: { date: string; load: number | null }[], lk: LangKontext | null, heute: string): ZpRahmen {
  const h = z.dauer_h || 0;
  const schwer = (z.hm || 0) >= 1500 || (z.distanz_km || 0) >= 40;
  let klasse: ZpRahmen['klasse'] = h >= 10 ? 'ultra' : h >= 7 ? 'lang' : h >= 4 ? 'mittel' : 'kurz';
  if (schwer && klasse === 'kurz') klasse = 'mittel';
  if (schwer && klasse === 'mittel' && h >= 6) klasse = 'lang';
  const WOCHEN: Record<ZpRahmen['klasse'], [number, number]> = { kurz: [4, 6], mittel: [6, 10], lang: [10, 14], ultra: [14, 20] };
  let [minW, maxW] = WOCHEN[klasse];
  if (z.prioritaet === 'B') { minW = Math.max(3, Math.round(minW * 0.75)); maxW = Math.round(maxW * 0.75); }
  if (z.prioritaet === 'C') { minW = Math.max(2, Math.round(minW * 0.5)); maxW = Math.max(minW + 1, Math.round(maxW * 0.5)); }
  const morgen = addT(heute, 1);
  const fr = addT(z.datum, -maxW * 7), sp = addT(z.datum, -minW * 7);
  const fruehestens = fr < morgen ? morgen : montagVon(fr) < morgen ? morgen : montagVon(fr);
  const spaetestens = sp < morgen ? morgen : sp;
  const mitte = addT(z.datum, -Math.round(((minW + maxW) / 2) * 7));
  const vorschlag = mitte < morgen ? morgen : montagVon(mitte) < morgen ? morgen : montagVon(mitte);
  const TAPER: Record<string, Record<ZpRahmen['klasse'], number>> = {
    A: { kurz: 5, mittel: 7, lang: 10, ultra: 14 }, B: { kurz: 4, mittel: 5, lang: 7, ultra: 7 }, C: { kurz: 2, mittel: 2, lang: 3, ultra: 3 },
  };
  const taperTage = TAPER[z.prioritaet || 'A'][klasse];
  const zielSport = zielSportVon(z);
  const geh = ['Wandern', 'Bergtour', 'Skitour'].includes(zielSport);
  // Ziellänge der längsten Vorbereitungseinheit: 70 % bis 6 h Tourdauer, darüber höchstens 60 % (Ultra-Märsche)
  const zielH = Math.max(2, Math.round(Math.min(h * 0.7, h > 8 ? Math.max(5.5, h * 0.6) : h * 0.7) * 4) / 4);
  // Fitness: ATL/CTL über die ganze Historie fortschreiben
  const loads = hist.map((x) => x.load ?? 0);
  let atl = 0, ctl = 0;
  loads.forEach((l, i) => { atl = atlStep(atl, l); ctl = ctlStep(ctl, l, windowMean(loads.slice(Math.max(0, i - 27), i + 1))); });
  const wochenlastJetzt = Math.round(loads.slice(-28).reduce((a, b) => a + b, 0) / 4);
  const monat = d2(z.datum).getMonth() + 1;
  const winter = monat >= 11 || monat <= 3;
  const sportarten = winter
    ? ['Rad', 'Lauf', 'Wandern', 'Bergtour', 'Skitour', 'Skike', 'Langlauf', 'Kraft', 'Mobility', 'Off']
    : ['Rad', 'Lauf', 'Wandern', 'Bergtour', 'Skike', 'Kraft', 'Mobility', 'Off'];
  const langWt = lk?.einstellung?.tage?.length ? Math.max(...lk.einstellung.tage) : 6;
  return {
    klasse,
    klasseText: `${{ kurz: 'kurze Tour', mittel: 'mittlere Tour', lang: 'lange Tour', ultra: 'Ultra' }[klasse]} (${String(h).replace('.', ',')} h${z.hm ? `, ${z.hm} Hm` : ''}${z.distanz_km ? `, ${z.distanz_km} km` : ''}${schwer && h < 7 ? ' – wegen Höhenmetern/Distanz eine Stufe höher' : ''})`,
    fruehestens, spaetestens, vorschlag, knapp: sp < morgen,
    taperTage, taperAb: addT(z.datum, -taperTage),
    zielSport, zielH, schrittH: geh ? 1 : 0.5, laengsteH: Math.round((lk?.status?.laengste28h || 0) * 4) / 4,
    tourLast: Math.round(tourLast), wochenlastJetzt, ctlJetzt: Math.round(ctl), atlJetzt: Math.round(atl),
    langWt, sportarten, raten: ratenAus(lk?.essProStunde?.spanne),
  };
}

/** Prompt für Kira (lokal). */
export function zielplanPrompt(z: Ziel, r: ZpRahmen, extra: string): { role: 'system' | 'user'; content: string }[] {
  const raten = Object.entries(r.raten).filter(([k]) => r.sportarten.includes(k)).map(([k, v]) => `${k} ${v}`).join(', ');
  const user = [
    `ZIEL: ${z.name} am ${z.datum} (${WT_NAME[wtVon(z.datum)]}), ${z.dauer_h} h, ${z.hm} Hm, ${z.distanz_km} km, Priorität ${z.prioritaet}. ${z.notiz || ''}`.trim(),
    `Einordnung: ${r.klasseText}. Zielsport: ${r.zielSport}. Geschätzte Tourlast am Zieltag: ${r.tourLast} ESS.`,
    `HEUTE: Fitness CTL ${r.ctlJetzt}, ATL ${r.atlJetzt}, Wochenlast der letzten 4 Wochen Ø ${r.wochenlastJetzt} ESS. Längste Einheit der letzten 4 Wochen ${r.laengsteH} h.`,
    `RAHMEN DES RECHENKERNS (verbindlich):`,
    `- Einstieg zwischen ${r.fruehestens} und ${r.spaetestens}${r.knapp ? ' (knapp: Ziel ist schon nah, sofort einsteigen)' : ''}; Vorschlag des Kerns ${r.vorschlag}. Du entscheidest und begründest kurz (Qualität, Intensität, Länge des Ziels, aktuelle Form).`,
    `- Tapering ab ${r.taperAb} (${r.taperTage} Tage). Danach Zieltag und 2 Tage Erholung – die legt der Kern selbst an.`,
    `- Lange Einheit wöchentlich am ${WT_NAME[r.langWt]} in ${r.zielSport === 'Skitour' ? 'Skitour (ohne Schnee: Bergtour)' : r.zielSport}; ihre Dauer steigert der Kern von ${r.laengsteH} h in Schritten von höchstens ${r.schrittH} h pro Woche bis ${r.zielH} h. Markiere diesen Tag mit "lang": true.`,
    `- Wochenlast: Grundlage etwa auf heutigem Niveau (Ø ${r.wochenlastJetzt} ESS), Aufbau und Spezifisch bis etwa +20 %, Tapering etwa 60 %. Steigerung höchstens 8 % pro Woche; jede 4. Woche Entlastung (macht der Kern). ACWR muss unter 1,3 bleiben. Lange Geh-Einheiten bringen wenig ESS (Wandern ${r.raten.Wandern}/h) – die Wochenlast kommt dann aus Rad und Lauf.`,
    `- Vortag und Folgetag der langen Einheit locker (Z1–Z2) oder frei. Mindestens 1 Ruhetag pro Woche. Höchstens 2 harte Tage (Z4–Z5) pro Woche.`,
    `- Erlaubte Sportarten: ${r.sportarten.join(', ')}. ESS pro Stunde ruhig bei Marc: ${raten}.`,
    extra ? `KONTEXT:\n${extra}` : '',
    '',
    'AUFGABE: Lege den Einstieg fest und teile die Zeit bis zum Tapering in 2–4 Phasen (z. B. Grundlage, Aufbau, Spezifisch) plus die Phase Tapering. ' +
      'Je Phase: Wochenlast (ESS) am Phasenanfang, Fokus und ein Wochenmuster mit genau 7 Tagen (wt 1=Mo … 7=So): Sportart, Zone, Anteil an der Wochenlast in Prozent (Summe 100, Ruhetag 0), Training in max. 6 Wörtern. ' +
      'Spezifisch heißt: Sport und Gelände wie am Zieltag (Höhenmeter, Dauer, Rucksack), Back-to-Back-Tage sind erlaubt.',
    'Antworte nur mit JSON: {"einstieg":"YYYY-MM-DD","begruendung":"2 Sätze","phasen":[{"name":"Grundlage","von":"YYYY-MM-DD","bis":"YYYY-MM-DD","wochen_ess":500,"fokus":"1 Satz","woche":[{"wt":1,"sport":"Rad","zone":"Z2","anteil":15,"training":"Pendeln locker"}, … 7 Einträge, einer mit "lang": true]}],"tapering":"1 Satz","hinweise":["…"]}',
  ].filter(Boolean).join('\n');
  return [
    { role: 'system', content: 'Du bist Coach Kira, Sportwissenschaftlerin und Marcs Trainerin. Deutsch, du-Form, keine Emojis. Du planst konservativ und konkret. Antworte nur mit JSON.' },
    { role: 'user', content: user },
  ];
}

const ZONEN: ZpZone[] = ['Off', 'Z1', 'Z2', 'Z3', 'Z4', 'Z5'];
const zoneVon = (v: unknown): ZpZone => { const s = String(v || '').toUpperCase().replace('OFF', 'Off'); return (ZONEN as string[]).includes(s) ? (s as ZpZone) : 'Z2'; };
const sportVon = (v: unknown, erlaubt: string[]): string => {
  const s = String(v || '').trim();
  const m = erlaubt.find((x) => x.toLowerCase() === s.toLowerCase()) || erlaubt.find((x) => s.toLowerCase().includes(x.toLowerCase()));
  if (m) return m;
  if (/bike|gravel|rad/i.test(s)) return 'Rad';
  if (/run|lauf/i.test(s)) return 'Lauf';
  if (/hike|wander|walk/i.test(s)) return 'Wandern';
  if (/ruhe|rest|frei|off/i.test(s)) return 'Off';
  return 'Rad';
};

/** Kiras Antwort lesen und auf den Rahmen zurechtrücken. */
export function zielplanLesen(j: any, r: ZpRahmen, zielDatum: string): { einstieg: string; begruendung: string; phasen: ZpPhase[]; tapering: string; hinweise: string[]; korrekturen: string[] } {
  const korr: string[] = [];
  let einstieg = String(j?.einstieg || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(einstieg)) { einstieg = r.vorschlag; korr.push(`Kein Einstieg von Kira – Vorschlag des Kerns ${einstieg}.`); }
  if (einstieg < r.fruehestens) { korr.push(`Einstieg ${einstieg} vor dem Rahmen – auf ${r.fruehestens} gesetzt.`); einstieg = r.fruehestens; }
  if (einstieg > r.spaetestens) { korr.push(`Einstieg ${einstieg} zu spät für das Ziel – auf ${r.spaetestens} gesetzt.`); einstieg = r.spaetestens; }
  const roh: any[] = Array.isArray(j?.phasen) ? j.phasen : Array.isArray(j?.phases) ? j.phases : [];
  let phasen: ZpPhase[] = roh.map((p) => {
    const woche: ZpMuster[] = [];
    const w: any[] = Array.isArray(p?.woche) ? p.woche : [];
    for (let t = 1; t <= 7; t++) {
      const x = w.find((y) => Number(y?.wt) === t) || w[t - 1] || {};
      const sport = sportVon(x.sport, r.sportarten);
      const zone = sport === 'Off' ? 'Off' : zoneVon(x.zone);
      const anteil = sport === 'Off' || zone === 'Off' ? 0 : Math.max(0, Number(x.anteil) || 0);
      woche.push({ wt: t, sport, zone, anteil, training: String(x.training || '').slice(0, 60), lang: !!x.lang });
    }
    return { name: String(p?.name || 'Phase').slice(0, 30), von: String(p?.von || '').slice(0, 10), bis: String(p?.bis || '').slice(0, 10), wochen_ess: Math.max(0, Math.round(Number(p?.wochen_ess) || 0)), fokus: String(p?.fokus || '').slice(0, 160), woche };
  }).filter((p) => p.woche.some((x) => x.anteil > 0));
  if (!phasen.length) throw new Error('Kira hat keine verwertbaren Phasen geliefert – bitte noch einmal versuchen.');
  // Phasen auf Einstieg … Tag vor dem Ziel ziehen; Tapering-Phase als letzte erkennen
  const taperIdx = phasen.findIndex((p) => /taper/i.test(p.name));
  const aufbau = taperIdx >= 0 ? phasen.filter((_, i) => i !== taperIdx) : phasen;
  const taper = taperIdx >= 0 ? phasen[taperIdx] : null;
  aufbau.sort((a, b) => (a.von < b.von ? -1 : 1));
  const ende = addT(r.taperAb, -1);
  // gleichmäßig verteilen, wenn Kiras Daten nicht passen
  const passt = aufbau.every((p, i) => /^\d{4}-\d{2}-\d{2}$/.test(p.von) && /^\d{4}-\d{2}-\d{2}$/.test(p.bis) && p.von <= p.bis && (i === 0 || p.von > aufbau[i - 1].bis));
  const n = diffT(einstieg, ende) + 1;
  aufbau.forEach((p, i) => {
    if (!passt) { p.von = addT(einstieg, Math.round((n * i) / aufbau.length)); p.bis = addT(einstieg, Math.round((n * (i + 1)) / aufbau.length) - 1); }
    if (i === 0) p.von = einstieg;
    if (i === aufbau.length - 1) p.bis = ende;
    if (i > 0) p.von = addT(aufbau[i - 1].bis, 1);
  });
  if (!passt) korr.push('Phasengrenzen gleichmäßig verteilt (Kiras Daten passten nicht zusammen).');
  const letzte = aufbau[aufbau.length - 1];
  const tp: ZpPhase = taper || {
    name: 'Tapering', von: '', bis: '', wochen_ess: Math.round((letzte?.wochen_ess || r.wochenlastJetzt) * 0.6), fokus: 'Umfang runter, kurze lockere Reize, frisch zum Ziel',
    woche: (letzte?.woche || []).map((x) => ({ ...x, zone: x.zone === 'Z5' || x.zone === 'Z4' ? 'Z3' : x.zone })),
  };
  if (!taper) korr.push('Tapering-Phase ergänzt (60 % der letzten Wochenlast).');
  tp.von = r.taperAb; tp.bis = addT(zielDatum, -1); tp.name = 'Tapering';
  phasen = [...aufbau, tp];
  // Lange Einheit: genau ein Tag, bevorzugt Marcs Wochentag
  for (const p of phasen) {
    const l = p.woche.filter((x) => x.lang);
    if (l.length !== 1) {
      p.woche.forEach((x) => (x.lang = false));
      const am = p.woche[r.langWt - 1];
      if (am.anteil <= 0) { am.sport = r.zielSport; am.zone = 'Z2'; am.anteil = 25; am.training = 'Lange Einheit ruhig'; }
      am.lang = true;
    }
  }
  return {
    einstieg, begruendung: String(j?.begruendung || '').slice(0, 400), phasen,
    tapering: String(j?.tapering || '').slice(0, 200), hinweise: (Array.isArray(j?.hinweise) ? j.hinweise : []).map((x: unknown) => String(x).slice(0, 200)).slice(0, 5), korrekturen: korr,
  };
}

const HART = (z: ZpZone) => z === 'Z3' || z === 'Z4' || z === 'Z5';

/** Tag für Tag ausrollen und prüfen. hist wie bei zielRahmen. */
export function zielplanAusrollen(z: Ziel, r: ZpRahmen, k: ReturnType<typeof zielplanLesen>, hist: { date: string; load: number | null }[], engine: string): Zielplan {
  const korr = [...k.korrekturen];
  const ende = addT(z.datum, 2);
  const tage: ZpTag[] = [];
  const phaseAm = (d: string) => k.phasen.find((p) => d >= p.von && d <= p.bis) || k.phasen[k.phasen.length - 1];
  const naechste = (p: ZpPhase) => k.phasen[k.phasen.indexOf(p) + 1];
  const mo0 = montagVon(k.einstieg);
  // Dauer der langen Einheit je Woche
  let langH = Math.max(r.zielSport === 'Rad' || r.zielSport === 'Lauf' ? 1.5 : 2, r.laengsteH);
  const langJeWoche = new Map<number, number>();
  const wochen = Math.ceil((diffT(mo0, ende) + 1) / 7);
  const aufbauEnde = addT(r.taperAb, -1);
  for (let w = 0; w < wochen; w++) {
    const mo = addT(mo0, 7 * w);
    const entl = w % 4 === 3 && mo < r.taperAb;
    if (mo >= r.taperAb) langJeWoche.set(w, Math.min(2, Math.max(1.5, r.zielH * 0.5)));
    else if (entl) langJeWoche.set(w, Math.max(1.5, langH - r.schrittH));
    else {
      // so steigern, dass zielH bis zum Aufbauende erreicht wird, höchstens schrittH pro Woche
      const restW = Math.max(1, Math.round(diffT(mo, aufbauEnde) / 7));
      const schritt = Math.min(r.schrittH, Math.max(0.25, (r.zielH - langH) / restW));
      langH = Math.min(r.zielH, w === 0 ? langH : langH + schritt);
      langJeWoche.set(w, Math.round(langH * 4) / 4);
    }
  }
  for (let d = k.einstieg; d <= ende; d = addT(d, 1)) {
    const w = Math.floor(diffT(mo0, d) / 7);
    const wt = wtVon(d);
    if (d === z.datum) {
      tage.push({ date: d, wt, woche: w, phase: 'Ziel', sport: r.zielSport, zone: 'Z2', intensitaet: 'Zieltour, gleichmäßig', ess: r.tourLast, dauerH: z.dauer_h, training: z.name, art: 'ziel' });
      continue;
    }
    if (d > z.datum) {
      const t1 = diffT(z.datum, d) === 1;
      tage.push({ date: d, wt, woche: w, phase: 'Erholung', sport: t1 ? 'Off' : 'Rad', zone: t1 ? 'Off' : 'Z1', intensitaet: t1 ? 'Ruhe' : 'sehr locker', ess: t1 ? 0 : 25, dauerH: t1 ? null : 0.5, training: t1 ? 'Ruhetag, Beine hoch' : 'Locker ausrollen', art: 'erholung' });
      continue;
    }
    const p = phaseAm(d);
    const nx = naechste(p);
    const mo = addT(mo0, 7 * w);
    const entl = w % 4 === 3 && mo < r.taperAb;
    // Wochenlast: linear von dieser zur nächsten Phase, Entlastung ×0,7, im Tapering fallend
    const anteilPhase = Math.min(1, Math.max(0, diffT(p.von, d) / Math.max(1, diffT(p.von, p.bis) + 1)));
    let wl = p.wochen_ess || r.wochenlastJetzt;
    if (nx && !/taper/i.test(p.name) && !/taper/i.test(nx.name)) wl = wl + (nx.wochen_ess - wl) * anteilPhase;
    if (/taper/i.test(p.name)) wl = wl * (1 - 0.4 * Math.min(1, diffT(r.taperAb, d) / Math.max(1, r.taperTage)));
    if (entl) wl *= 0.7;
    const m = p.woche[wt - 1];
    const summe = p.woche.reduce((a, x) => a + (x.lang ? 0 : x.anteil), 0) || 1;
    const lh = langJeWoche.get(w) || 2;
    let sport = m.sport, zone = m.zone, training = m.training, dauerH: number | null = null, ess = 0, art: ZpTag['art'];
    const tageBisZiel = diffT(d, z.datum);
    if (tageBisZiel === 1) { sport = 'Off'; zone = 'Off'; training = 'Ruhetag vor dem Ziel'; }
    else if (tageBisZiel === 2) { sport = 'Rad'; zone = 'Z1'; training = 'Kurz locker, Beine wach'; dauerH = 0.75; ess = rund5(0.75 * r.raten.Rad * ZONE_FAKTOR.Z1); }
    else if (m.lang) {
      sport = r.zielSport === 'Skitour' && sport !== 'Skitour' && sport !== 'Bergtour' ? 'Skitour' : sport === 'Off' ? r.zielSport : sport;
      zone = HART(zone) ? 'Z2' : zone === 'Off' ? 'Z2' : zone;
      dauerH = lh;
      ess = rund5(lh * (r.raten[sport] || 30) * ZONE_FAKTOR[zone]);
      training = `${training && !/^lange/i.test(training) ? training + ' – ' : ''}lange Einheit ${String(lh).replace('.', ',')} h`;
      art = 'lang';
    } else if (m.anteil > 0 && zone !== 'Off') {
      const langEss = (() => { const l = p.woche.find((x) => x.lang); return l ? lh * (r.raten[l.sport] || 30) * ZONE_FAKTOR[l.zone === 'Off' ? 'Z2' : l.zone] : 0; })();
      // Lange Einheiten im Gehen (Wandern/Bergtour) sind ESS-arm: Restlast trotzdem nicht über die Wochenlast hinaus
      const rest = Math.max(wl * 0.4, wl - langEss);
      ess = rund5((rest * m.anteil) / summe);
      const rate = (r.raten[sport] || 40) * ZONE_FAKTOR[zone];
      // Nebeneinheiten bleiben alltagstauglich: Gehen höchstens 2,5 h, Kraft/Mobility 1 h, sonst 2,5 h
      const deckel = sport === 'Kraft' || sport === 'Mobility' ? 1 : 2.5;
      if (rate > 0 && ess / rate > deckel) ess = rund5(deckel * rate);
      dauerH = rate > 0 && ess > 0 ? Math.round((ess / rate) * 4) / 4 : null;
    } else { sport = 'Off'; zone = 'Off'; }
    if (!art) art = /taper/i.test(p.name) ? 'taper' : entl ? 'entlastung' : undefined;
    tage.push({ date: d, wt, woche: w, phase: p.name, sport, zone, intensitaet: ZONE_INTENSITAET[zone], ess, dauerH, training, art });
  }
  // Regel: Vortag/Folgetag der langen Einheit und des Ziels nicht hart
  tage.forEach((t, i) => {
    if (t.art !== 'lang' && t.art !== 'ziel') return;
    for (const j of [i - 1, i + 1]) {
      const x = tage[j];
      if (x && HART(x.zone) && x.art !== 'lang') {
        korr.push(`${WT_NAME[x.wt]} ${x.date.slice(8, 10)}.${x.date.slice(5, 7)}.: ${x.zone} neben der langen Einheit auf Z2 gesenkt.`);
        x.ess = rund5(x.ess * 0.7); x.zone = 'Z2'; x.intensitaet = ZONE_INTENSITAET.Z2; x.training = `${x.training ? x.training + ' – ' : ''}locker statt hart`;
        x.dauerH = x.dauerH != null ? Math.round(((x.ess / ((r.raten[x.sport] || 40) * ZONE_FAKTOR.Z2)) || 0) * 4) / 4 || null : null;
      }
    }
  });
  // ACWR prüfen: ab heute fortschreiben; Wochen mit ACWR > 1,3 dämpfen (lange Einheit und Ziel bleiben)
  const basis = hist.map((x) => x.load ?? 0);
  const heute = hist.length ? addT(hist[hist.length - 1].date, 1) : k.einstieg;
  const sim = () => {
    let atl = 0, ctl = 0;
    const reihe = [...basis];
    basis.forEach((l, i) => { atl = atlStep(atl, l); ctl = ctlStep(ctl, l, windowMean(basis.slice(Math.max(0, i - 27), i + 1))); });
    const ctl0 = ctl;
    const byDate = new Map(tage.map((t) => [t.date, t]));
    let mx = 0, ctlZ = ctl;
    for (let d = heute; d <= ende; d = addT(d, 1)) {
      const t = byDate.get(d);
      const l = t ? t.ess : Math.round(reihe.slice(-28).reduce((a, b) => a + b, 0) / 28); // vor dem Einstieg: wie bisher
      reihe.push(l);
      atl = atlStep(atl, l);
      ctl = ctlStep(ctl, l, windowMean(reihe.slice(-28)));
      const ac = ctl > 0 ? atl / ctl : 0;
      if (t) t.acwr = Math.round(ac * 100) / 100;
      if (t && t.art !== 'ziel' && t.art !== 'erholung') mx = Math.max(mx, ac);
      if (d === addT(z.datum, -1)) ctlZ = ctl;
    }
    return { mx, ctl0, ctlEnd: ctlZ };
  };
  let s = sim();
  for (let it = 0; it < 8 && s.mx > 1.3; it++) {
    const wochenHoch = new Set(tage.filter((t) => (t.acwr || 0) > 1.3 && t.art !== 'ziel' && t.art !== 'erholung').map((t) => t.woche));
    for (const t of tage) if (wochenHoch.has(t.woche) && t.art !== 'lang' && t.art !== 'ziel' && t.ess > 0) { t.ess = rund5(t.ess * 0.9); if (t.dauerH) t.dauerH = Math.max(0.25, Math.round(t.dauerH * 0.9 * 4) / 4); }
    if (it === 0) korr.push(`ACWR über 1,3 in Woche ${[...wochenHoch].map((w) => w + 1).join(', ')} – Nebeneinheiten dieser Wochen gekürzt.`);
    s = sim();
  }
  if (s.ctl0 > 0 && s.ctlEnd < s.ctl0 * 0.85) korr.push(`Fitness (CTL) fällt bis zum Ziel von ${Math.round(s.ctl0)} auf ${Math.round(s.ctlEnd)} – Wochenlast im Aufbau eher höher ansetzen.`);
  if (s.mx > 1.3) korr.push(`ACWR bleibt an einzelnen Tagen über 1,3 (max. ${String(Math.round(s.mx * 100) / 100).replace('.', ',')}) – dort locker bleiben.`);
  const zielTag = tage.find((t) => t.art === 'ziel');
  return {
    ziel: { name: z.name, datum: z.datum }, erstellt: new Date().toISOString(), engine, einstieg: k.einstieg, begruendung: k.begruendung,
    rahmen: r, phasen: k.phasen.map(({ woche: _w, ...p }) => p), tage, korrekturen: korr, tapering: k.tapering, hinweise: k.hinweise,
    maxAcwr: Math.round(s.mx * 100) / 100, ctlStart: Math.round(s.ctl0), ctlZiel: Math.round(zielTag ? s.ctlEnd : 0),
  };
}

/** Wochensummen für die Tabelle. */
export function zielplanWochen(p: Zielplan): { woche: number; mo: string; phase: string; ess: number; entl: boolean; tage: ZpTag[] }[] {
  const m = new Map<number, ZpTag[]>();
  for (const t of p.tage) (m.get(t.woche) || m.set(t.woche, []).get(t.woche)!).push(t);
  return [...m.entries()].sort((a, b) => a[0] - b[0]).map(([w, ts]) => ({
    woche: w, mo: montagVon(ts[0].date), phase: ts.find((t) => t.art !== 'erholung')?.phase || ts[0].phase,
    ess: ts.reduce((a, t) => a + t.ess, 0), entl: ts.some((t) => t.art === 'entlastung'), tage: ts,
  }));
}

/** Zeilen für den Wochenplan-Prompt: Vorgaben eines Zielplans für die nächsten Tage. */
export function zielplanVorgabe(p: Zielplan | null | undefined, heute: string, tage = 14): string {
  if (!p) return '';
  const bis = addT(heute, tage - 1);
  const xs = p.tage.filter((t) => t.date >= heute && t.date <= bis);
  if (!xs.length) return '';
  return `ZIELPLAN „${p.ziel.name}“ (${p.ziel.datum}) – Vorgabe für diese Tage, nur bei Tagesvotum/ACWR abweichen:\n` +
    xs.map((t) => `- ${t.date} ${WT_NAME[t.wt]}: ${t.phase}, ${t.sport} ${t.zone}, ${t.ess} ESS${t.dauerH ? `, ca. ${t.dauerH} h` : ''}${t.training ? ' – ' + t.training : ''}`).join('\n');
}

export function zielplanCsv(p: Zielplan): string {
  const kopf = 'Datum;Tag;Woche;Phase;Sportart;Zone;Intensität;ESS;Dauer h;Einheit';
  return [kopf, ...p.tage.map((t) => [t.date, WT_NAME[t.wt], t.woche + 1, t.phase, t.sport, t.zone, t.intensitaet, t.ess, t.dauerH != null ? String(t.dauerH).replace('.', ',') : '', `"${(t.training || '').replace(/"/g, "'")}"`].join(';'))].join('\n');
}
