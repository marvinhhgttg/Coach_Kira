/**
 * Coach Kira Paket 3 – reines TypeScript, keine React-Imports.
 */

// ─── Datentypen ──────────────────────────────────────────────────────────────

export interface P3Tag {
  date: string;
  load: number | null;
  sleep: number | null;
  sleepScore: number | null;
  readiness: number | null;
  hrv: number | null;
  atl: number | null;
  ctl: number | null;
  es: number | null;
  zone: string;
  sport: string;
}

// ─── Parse-Hilfe ─────────────────────────────────────────────────────────────

function parseNum(v: unknown): number | null {
  if (v === '' || v == null) return null;
  const s = String(v).trim();
  if (s === '' || s.toLowerCase() === 'nan') return null;
  const n = Number(s.replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

// ─── fetchChartData import ──────────────────────────────────────────────────

import { fetchChartData } from './api.js';
import { esStep, ES_MODEL, atlStep, ctlStep, CTL_MODEL } from './ctlModel.js';
import { zonenKlasse } from './recoveryPrice.js';

// ─── P3Tag laden (Promise-Cache) ────────────────────────────────────────────

let p3Cache: P3Tag[] | null = null;
let p3Promise: Promise<P3Tag[]> | null = null;

export async function ladeP3(): Promise<P3Tag[]> {
  if (p3Cache) return p3Cache;
  if (!p3Promise) {
    p3Promise = (async () => {
      const r: any = await fetchChartData('9999d', [
        'load_fb_day','sleep_hours','sleep_score_0_100','Garmin_Training_Readiness',
        'hrv_status','fbATL_obs','fbCTL_obs','garminEnduranceScore','Zone','Sport_x',
      ]);
      const rows: Array<Record<string, unknown>> = (r && r.data) || [];
      const todayStr = new Date().toLocaleDateString('sv-SE');

      const seen = new Map<string, P3Tag>();
      for (const row of rows) {
        const d = String(row.date || '').split('T')[0];
        if (d === '' || d > todayStr) continue;
        if (seen.has(d)) continue;
        seen.set(d, {
          date: d,
          load: parseNum(row.load_fb_day),
          sleep: parseNum(row.sleep_hours),
          sleepScore: parseNum(row.sleep_score_0_100),
          readiness: parseNum(row.Garmin_Training_Readiness),
          hrv: parseNum(row.hrv_status),
          atl: parseNum(row.fbATL_obs),
          ctl: parseNum(row.fbCTL_obs),
          es: parseNum(row.garminEnduranceScore),
          zone: String(row.Zone || ''),
          sport: String(row.Sport_x || ''),
        });
      }
      const result = Array.from(seen.values()).sort((a, b) => a.date.localeCompare(b.date));
      p3Cache = result;
      return result;
    })();
  }
  return p3Promise;
}

// ─── schlafKonto ─────────────────────────────────────────────────────────────

export function schlafKonto(
  t: P3Tag[],
  loadHeute: number | null,
): { schuld: number | null; quartile: Array<{ bis: number; readiness: number; n: number }>; jetztQuartil: number | null; bedarfH: number; zuschlagLast: number; zuschlagSchuld: number } {
  interface SchuldEintrag { index: number; schuld: number | null }
  const eintraege: SchuldEintrag[] = [];

  for (let i = 6; i < t.length; i++) {
    let sum = 0, cnt = 0;
    for (let j = i - 6; j <= i; j++) {
      const s = t[j].sleep;
      if (s === null) continue;
      sum += 7.5 - s;
      cnt++;
    }
    const schuld = cnt >= 5 ? sum : null;
    eintraege.push({ index: i, schuld });
  }

  const letzerSchuld = eintraege.length > 0 ? eintraege[eintraege.length - 1].schuld : null;

  type QEintrag = { Schuld: number; readiness: number };
  const qDaten: QEintrag[] = [];
  for (const e of eintraege) {
    if (e.schuld === null) continue;
    const r = t[e.index].readiness;
    if (r === null) continue;
    qDaten.push({ Schuld: e.schuld, readiness: r });
  }

  const schuldWerte = qDaten.map(d => d.Schuld).sort((a, b) => a - b);

  function quartilGrenze(sorted: number[], q: number): number {
    if (sorted.length === 0) return 0;
    const idx = q * (sorted.length - 1);
    const lo = Math.floor(idx);
    const hi = Math.ceil(idx);
    if (lo === hi) return sorted[lo];
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
  }

  const q1 = quartilGrenze(schuldWerte, 0.25);
  const q2 = quartilGrenze(schuldWerte, 0.5);
  const q3 = quartilGrenze(schuldWerte, 0.75);
  const maxSchuld = schuldWerte.length > 0 ? schuldWerte[schuldWerte.length - 1] : 0;
  const grenzen = [q1, q2, q3, maxSchuld];

  function quartilName(val: number): number {
    if (val <= q1) return 0;
    if (val <= q2) return 1;
    if (val <= q3) return 2;
    return 3;
  }

  const quartile: Array<{ bis: number; readiness: number; n: number }> = [];
  for (let q = 0; q < 4; q++) {
    const loLimit = q === 0 ? -Infinity : grenzen[q - 1];
    const hiLimit = grenzen[q];
    const passthrough: QEintrag[] = [];
    for (const d of qDaten) {
      if (q < 3 && d.Schuld >= loLimit && d.Schuld <= hiLimit) passthrough.push(d);
      else if (q === 3 && d.Schuld >= loLimit) passthrough.push(d);
    }
    const avgReadiness = passthrough.length > 0 ? passthrough.reduce((a, d) => a + d.readiness, 0) / passthrough.length : 0;
    quartile.push({ bis: hiLimit, readiness: avgReadiness, n: passthrough.length });
  }

  const jetzterWert = letzerSchuld;
  const jetztQuartil: number | null = jetzterWert === null ? null : quartilName(jetzterWert);

  const zuschlagLast = Math.max(0, (loadHeute ?? 0) - 100) / 100 * 0.5;
  const schuldBetrag = Math.min(Math.max(jetzterWert ?? 0, 0), 4) / 4 * 0.5;
  const bedarfH = Math.min(7.5 + zuschlagLast + schuldBetrag, 9.5);

  return {
    schuld: letzerSchuld,
    quartile,
    jetztQuartil,
    bedarfH,
    zuschlagLast,
    zuschlagSchuld: schuldBetrag,
  };
}

// ─── wochenrhythmus ──────────────────────────────────────────────────────────

export type Fenster = 'alle' | '90' | 'sommer' | 'winter';

export function wochenrhythmus(
  t: P3Tag[],
  f: Fenster,
): { zeilen: Array<{ key: 'sleep'|'readiness'|'hrv'|'load'; label: string; werte: Array<number|null>; n: number[] }>; besteTage: string[]; schwaechsterTag: string } {
  function weekdayIdx(d: string): number {
    const parts = d.split('-');
    const dt = new Date(Date.UTC(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2])));
    return (dt.getUTCDay() + 6) % 7; // Mon=0 .. Sun=6
  }

  const tage = t.filter(tag => {
    if (f === '90') return true;
    const parts = tag.date.split('-');
    const month = parseInt(parts[1]);
    if (f === 'sommer') return month >= 4 && month <= 9;
    if (f === 'winter') return month >= 10 || month <= 3;
    return true;
  });

  const filtered = f === '90' && tage.length > 0 ? tage.slice(-90) : tage;

  type WDData = { sleep: (number|null)[]; readiness: (number|null)[]; hrv: (number|null)[]; load: (number|null)[] };
  const wd: WDData[] = Array.from({ length: 7 }, () => ({
    sleep: [], readiness: [], hrv: [], load: []
  }));

  for (const tag of filtered) {
    const wi = weekdayIdx(tag.date);
    wd[wi].sleep.push(tag.sleep);
    wd[wi].readiness.push(tag.readiness);
    wd[wi].hrv.push(tag.hrv);
    wd[wi].load.push(tag.load);
  }

  const zeilen = [
    { key: 'sleep' as const, label: 'Schlaf h', werte: wd.map(w => avgOrNull(w.sleep)), n: wd.map(w => nonNullCount(w.sleep)) },
    { key: 'readiness' as const, label: 'Readiness', werte: wd.map(w => avgOrNull(w.readiness)), n: wd.map(w => nonNullCount(w.readiness)) },
    { key: 'hrv' as const, label: 'HRV', werte: wd.map(w => avgOrNull(w.hrv)), n: wd.map(w => nonNullCount(w.hrv)) },
    { key: 'load' as const, label: 'Last', werte: wd.map(w => avgOrNull(w.load)), n: wd.map(w => nonNullCount(w.load)) },
  ];

  const kuerzel = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];

  const readinessAvgs: Array<{ idx: number; avg: number }> = wd.map((w, i) => ({
    idx: i, avg: nonNullCount(w.readiness) > 0 ? (w.readiness.reduce((a: number, b: number | null) => a + (b ?? 0), 0) / nonNullCount(w.readiness)) : -1
  })).filter(r => r.avg >= 0);

  readinessAvgs.sort((a, b) => b.avg - a.avg);
  const besteTage = readinessAvgs.slice(0, 2).map(r => kuerzel[r.idx]);

  const schwächste = readinessAvgs.length > 0 ? readinessAvgs[readinessAvgs.length - 1] : null;
  const schwaechsterTag = schwächste ? kuerzel[schwächste.idx] : 'Mo';

  return { zeilen, besteTage, schwaechsterTag };
}

function avgOrNull(arr: (number | null)[]): number | null {
  const vals = arr.filter((v): v is number => v !== null);
  return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
}

function nonNullCount(arr: (number | null)[]): number {
  return arr.filter(v => v !== null).length;
}

// ─── formZone / Rampe ────────────────────────────────────────────────────────

export type FormZone = 'Transition' | 'Frisch' | 'Grau' | 'Optimal' | 'Hohes Risiko';

export function formZone(ctl: number, atl: number): { formPct: number; zone: FormZone } {
  const formPct = (ctl - atl) / ctl * 100;
  let zone: FormZone;
  if (formPct > 25) zone = 'Transition';
  else if (formPct > 5) zone = 'Frisch';
  else if (formPct > -10) zone = 'Grau';
  else if (formPct > -30) zone = 'Optimal';
  else zone = 'Hohes Risiko';
  return { formPct, zone };
}

export function rampe(t: P3Tag[]): { proWoche: number | null; pct: number | null; urteil: 'abbauend'|'aufbauend'|'ambitioniert'|'hoch'|null } {
  if (t.length < 8) return { proWoche: null, pct: null, urteil: null };
  const last = t[t.length - 1];
  const old7 = t[t.length - 8];
  if (last.ctl === null || old7.ctl === null) return { proWoche: null, pct: null, urteil: null };
  const ctlAlt = old7.ctl;
  const diff = last.ctl - ctlAlt;
  const pct = (diff / ctlAlt) * 100;
  let urteil: 'abbauend'|'aufbauend'|'ambitioniert'|'hoch';
  if (diff < 0) urteil = 'abbauend';
  else if (pct <= 5) urteil = 'aufbauend';
  else if (pct <= 8) urteil = 'ambitioniert';
  else urteil = 'hoch';
  return { proWoche: diff, pct, urteil };
}

// ─── esBacktest / esBreakEven / esZiel ──────────────────────────────────────

export function esBacktest(t: P3Tag[]): { n: number; maeModell: number; maeNaiv: number; horizont: 28 } {
  const n = t.length;
  const startIdxMin = Math.floor(n * 0.25);

  let maeModellSum = 0, maeNaivSum = 0, count = 0;

  for (let s = startIdxMin; s < n - 28; s++) {
    const esStart = t[s].es;
    const esZiel7w = t[s + 28].es;
    if (esStart === null || esZiel7w === null) continue;

    let currentEs = esStart;
    for (let k = 1; k <= 28; k++) {
      const loadK = t[s + k].load ?? 0;
      const startMean = Math.max(0, s + k - 56);
      const endMean = s + k;
      const lasten: number[] = [];
      for (let m = startMean; m < endMean && m < t.length; m++) {
        lasten.push(t[m].load ?? 0);
      }
      const m56 = lasten.length > 0 ? lasten.reduce((a, b) => a + b, 0) / lasten.length : 0;
      currentEs = esStep(currentEs, loadK, m56);
    }

    const prognose = currentEs;
    maeModellSum += Math.abs(prognose - esZiel7w);
    maeNaivSum += Math.abs(esStart - esZiel7w);
    count++;
  }

  return {
    n: count,
    maeModell: count > 0 ? maeModellSum / count : 0,
    maeNaiv: count > 0 ? maeNaivSum / count : 0,
    horizont: 28,
  };
}

export function esBreakEven(es: number): number {
  return (es * (1 - ES_MODEL.b) - ES_MODEL.a) / (ES_MODEL.c + ES_MODEL.d);
}

export function esZiel(
  t: P3Tag[],
  zielEs: number,
  zielDatum: string,
): { tage: number; ladeProTag: number | null; machbar: boolean } {
  if (t.length === 0) return { tage: 0, ladeProTag: null, machbar: false };
  const lastTag = t[t.length - 1];
  const ctlStart = lastTag.ctl ?? 0;
  const atlStart = lastTag.atl ?? 0;

  const zielDt = new Date(zielDatum + 'T00:00:00Z');
  const lastDt = new Date(lastTag.date + 'T00:00:00Z');
  const diffMs = zielDt.getTime() - lastDt.getTime();
  const tage = Math.round(diffMs / (24 * 60 * 60 * 1000));

  function forwardEs(L: number): number {
    let ctlVal = ctlStart;
    let atlVal = atlStart;
    const histLoads: number[] = [];
    for (let i = 0; i < t.length && i < 56; i++) {
      histLoads.push(t[i].load ?? 0);
    }
    let esVal = lastTag.es ?? 0;

    for (let day = 0; day < tage; day++) {
      // Update ATL
      atlVal = atlStep(atlVal, L);

      // mean28 for CTL
      const m28 = histLoads.length > 0 ? histLoads.slice(-Math.min(histLoads.length, 28)).reduce((s, v) => s + v, 0) / Math.min(histLoads.length, 28) : L;
      // Update CTL
      ctlVal = ctlStep(ctlVal, L, m28);

      // mean56 for ES
      const histLen = Math.min(histLoads.length, 56);
      const m56val = histLen > 0 ? histLoads.slice(-histLen).reduce((s, v) => s + v, 0) / histLen : L;
      // Update ES
      esVal = esStep(esVal, L, m56val);

      histLoads.push(L);
    }
    return esVal;
  }

  let lo = 0, hi = 400;
  for (let iter = 0; iter < 200; iter++) {
    const mid = (lo + hi) / 2;
    const esMid = forwardEs(mid);
    if (esMid > zielEs) hi = mid;
    else lo = mid;
  }

  const finalL = (lo + hi) / 2;
  return {
    tage,
    ladeProTag: finalL <= 250 ? Math.round(finalL * 100) / 100 : null,
    machbar: finalL <= 250,
  };
}

// ─── intensitaetProWoche ─────────────────────────────────────────────────────

export type WochenTyp = 'polarisiert' | 'pyramidal' | 'schwellenlastig' | 'Grundlage' | 'Ruhewoche';

function parseDateParts(d: string): { y: number; m: number; day: number } {
  const p = d.split('-');
  return { y: parseInt(p[0]), m: parseInt(p[1]), day: parseInt(p[2]) };
}

function mondayOfWeek(d: string): string {
  const parts = parseDateParts(d);
  const dt = new Date(Date.UTC(parts.y, parts.m - 1, parts.day));
  const dayOfWeek = (dt.getUTCDay() + 6) % 7;
  const mon = new Date(dt);
  mon.setUTCDate(dt.getUTCDate() - dayOfWeek);
  return mon.toISOString().slice(0, 10);
}

export function intensitaetProWoche(
  t: P3Tag[],
  wochen = 12,
): Array<{ woche: string; last: number; locker: number; grau: number; hart: number; typ: WochenTyp }> {
  if (t.length === 0) return [];

  const weekMap = new Map<string, P3Tag[]>();
  for (const tag of t) {
    const m = mondayOfWeek(tag.date);
    if (!weekMap.has(m)) weekMap.set(m, []);
    weekMap.get(m)!.push(tag);
  }

  const sortedWeeks = Array.from(weekMap.keys()).sort().slice(-wochen);

  const ergebnis: Array<{ woche: string; last: number; locker: number; grau: number; hart: number; typ: WochenTyp }> = [];

  const summe = (tage: P3Tag[]) => tage.reduce((a, x) => a + (x.load && x.load > 0 ? x.load : 0), 0);
  const wochenLast = sortedWeeks.map((wk) => summe(weekMap.get(wk)!)).sort((a, b) => a - b);
  const median = wochenLast.length ? wochenLast[Math.floor(wochenLast.length / 2)] : 0;
  for (const wk of sortedWeeks) {
    const tageWoche = weekMap.get(wk)!;
    let totalLoad = 0, lockerL = 0, grauL = 0, hartL = 0;
    for (const tag of tageWoche) {
      if (tag.load == null || tag.load <= 0) continue;
      totalLoad += tag.load;
      const k = zonenKlasse(tag.zone);
      if (k === 'locker') lockerL += tag.load;
      else if (k === 'grau') grauL += tag.load;
      else if (k === 'hart') hartL += tag.load;
    }
    const zonenL = lockerL + grauL + hartL;
    const lockerPct = zonenL > 0 ? lockerL / zonenL : 0;
    const grauPct = zonenL > 0 ? grauL / zonenL : 0;
    const hartPct = zonenL > 0 ? hartL / zonenL : 0;
    let typ: WochenTyp;
    if (totalLoad < 0.6 * median) typ = 'Ruhewoche';
    else if (hartPct >= 0.1 && grauPct < 0.15) typ = 'polarisiert';
    else if (grauPct >= 0.25) typ = 'schwellenlastig';
    else if (grauPct >= 0.1 && grauPct > hartPct) typ = 'pyramidal';
    else typ = 'Grundlage';
    ergebnis.push({ woche: wk, last: Math.round(totalLoad), locker: Math.round(lockerPct * 1000) / 1000, grau: Math.round(grauPct * 1000) / 1000, hart: Math.round(hartPct * 1000) / 1000, typ });
  }

  return ergebnis;
}
