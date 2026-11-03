/**
 * Erholungs- und Ernährungsanalysen aus garmin_tag + Timeline (Sheet): Erholungsfaktoren, Schlafrhythmus,
 * Atem/SpO2-Frühwarnung, Stress-Bilanz, Energiebilanz und Eiweiß. Rein deterministisch.
 */
import { useEffect, useState } from 'react';
import { fetchChartData } from './api';
import type { GarminTage } from './garmin';
import { cohenD, eintraege, hart, median, minuten, mittel, plusTage, quantil, schlaegeProKm, sd, startMin, tageSeit, wochentag, wochenStart, zuNachtMin, WT } from './stat';

export interface TlZeile { date: string; kcal_in: number | null; carb: number | null; protein: number | null; fett: number | null; defizit: number | null; ready: number | null; last: number | null; atl: number | null; ctl: number | null }
const z = (x: unknown) => { const n = Number(x); return x === '' || x == null || !Number.isFinite(n) ? null : n; };
let tlCache: Promise<Map<string, TlZeile>> | null = null;
export function fetchTimeline(): Promise<Map<string, TlZeile>> {
  if (!tlCache) {
    tlCache = fetchChartData('9999d', ['kcal_in', 'carb_g', 'protein_g', 'fat_g', 'deficit', 'Garmin_Training_Readiness', 'load_fb_day', 'fbATL_obs', 'fbCTL_obs'])
      .then((r: any) => {
        const m = new Map<string, TlZeile>();
        for (const x of (r && r.data) || []) {
          const d = String(x.date || '').slice(0, 10);
          if (!d) continue;
          m.set(d, { date: d, kcal_in: z(x.kcal_in), carb: z(x.carb_g), protein: z(x.protein_g), fett: z(x.fat_g), defizit: z(x.deficit), ready: z(x.Garmin_Training_Readiness), last: z(x.load_fb_day), atl: z(x.fbATL_obs), ctl: z(x.fbCTL_obs) });
        }
        return m;
      })
      .catch(() => { tlCache = null; return new Map(); });
  }
  return tlCache;
}
export function useTimeline() {
  const [tl, setTl] = useState<Map<string, TlZeile> | null>(null);
  useEffect(() => { let ok = true; fetchTimeline().then((m) => ok && setTl(m)); return () => { ok = false; }; }, []);
  return tl;
}

// ─── 7 Was dich am besten erholt ────────────────────────────────────────────
export interface Faktor { id: string; lab: string; gut: string; schlecht: string; n1: number; n2: number; hrv1: number | null; hrv2: number | null; bb1: number | null; bb2: number | null; d: number | null }
export function erholungsFaktoren(tage: GarminTage | null, tl: Map<string, TlZeile> | null) {
  if (!tage) return { faktoren: [] as Faktor[], n: 0 };
  // Ziel je Morgen d: HRV relativ zur eigenen 28-Tage-Basis (in %) und Body Battery beim Aufwachen
  const ds = [...tage.keys()].sort();
  const hrvRel = new Map<string, number>();
  for (const d of ds) {
    const h = tage.get(d)?.hrv?.nacht;
    if (!h) continue;
    const basis = median(ds.filter((x) => x < d && x >= plusTage(d, -28)).map((x) => tage.get(x)?.hrv?.nacht || NaN).filter(Number.isFinite));
    if (basis) hrvRel.set(d, ((h - basis) / basis) * 100);
  }
  type Zeile = { d: string; hrv: number | null; bb: number | null; x: Record<string, number | null> };
  const zeilen: Zeile[] = [];
  for (const d of ds) {
    const g = tage.get(d)!, v = tage.get(plusTage(d, -1));
    if (!v) continue;
    const hrv = hrvRel.get(d) ?? null, bb = g.bb_aufwachen ?? null;
    if (hrv == null && bb == null) continue;
    const akts = (v.akt || []).filter((a) => minuten(a) >= 15);
    const spaet = akts.map((a) => startMin(a) ?? 0).filter((m) => m >= 18 * 60);
    const t = tl?.get(plusTage(d, -1));
    zeilen.push({
      d, hrv, bb,
      x: {
        bett: zuNachtMin(g.schlaf?.start),
        dauer: g.schlaf?.dauer_s ? g.schlaf.dauer_s / 3600 : null,
        spaet: akts.length ? (spaet.length ? 1 : 0) : null,
        last: akts.reduce((a, x) => a + (x.last || 0), 0),
        stress: v.tag?.stress_avg ?? null,
        kcal: t?.kcal_in ?? null, carb: t?.carb ?? null, protein: t?.protein ?? null, fett: t?.fett ?? null,
        schritte: v.tag?.schritte ?? null,
      },
    });
  }
  const DEF: { id: string; lab: string; schwelle: 'median' | number; gut: string; schlecht: string }[] = [
    { id: 'bett', lab: 'Zubettgehen', schwelle: 'median', gut: 'früh ins Bett', schlecht: 'spät ins Bett' },
    { id: 'dauer', lab: 'Schlafdauer', schwelle: 'median', gut: 'langer Schlaf', schlecht: 'kurzer Schlaf' },
    { id: 'spaet', lab: 'Training am Abend', schwelle: 0.5, gut: 'Training vor 18 Uhr', schlecht: 'Training ab 18 Uhr' },
    { id: 'last', lab: 'Trainingslast am Vortag', schwelle: 'median', gut: 'wenig Last', schlecht: 'viel Last' },
    { id: 'stress', lab: 'Stress am Vortag', schwelle: 'median', gut: 'wenig Stress', schlecht: 'viel Stress' },
    { id: 'kcal', lab: 'Kalorien am Vortag', schwelle: 'median', gut: 'viel gegessen', schlecht: 'wenig gegessen' },
    { id: 'carb', lab: 'Kohlenhydrate am Vortag', schwelle: 'median', gut: 'viele Kohlenhydrate', schlecht: 'wenig Kohlenhydrate' },
    { id: 'protein', lab: 'Eiweiß am Vortag', schwelle: 'median', gut: 'viel Eiweiß', schlecht: 'wenig Eiweiß' },
    { id: 'fett', lab: 'Fett am Vortag', schwelle: 'median', gut: 'wenig Fett', schlecht: 'viel Fett' },
    { id: 'schritte', lab: 'Schritte am Vortag', schwelle: 'median', gut: 'viele Schritte', schlecht: 'wenig Schritte' },
  ];
  // „gut“ = Seite, die in den Daten tatsächlich besser erholt; Richtung wird erst nach der Rechnung festgelegt.
  const faktoren: Faktor[] = [];
  for (const f of DEF) {
    const xs = zeilen.filter((r) => r.x[f.id] != null);
    if (xs.length < 12) continue;
    const s = f.schwelle === 'median' ? median(xs.map((r) => r.x[f.id]!))! : f.schwelle;
    const lo = xs.filter((r) => r.x[f.id]! <= s && !(f.schwelle === 'median' && r.x[f.id]! === s && false));
    const hi = xs.filter((r) => r.x[f.id]! > s);
    if (lo.length < 5 || hi.length < 5) continue;
    const h = (rs: Zeile[]) => rs.map((r) => r.hrv).filter((v): v is number => v != null);
    const b = (rs: Zeile[]) => rs.map((r) => r.bb).filter((v): v is number => v != null);
    const dH = cohenD(h(lo), h(hi)), dB = cohenD(b(lo), b(hi));
    const d = dH != null && dB != null ? (dH + dB) / 2 : dH ?? dB;
    // Beschriftung: niedrige Seite zuerst (z. B. früh ins Bett = kleiner Wert)
    const niedrigLab = { bett: 'früh ins Bett', dauer: 'kurzer Schlaf', spaet: 'Training vor 18 Uhr', last: 'wenig Last', stress: 'wenig Stress', kcal: 'wenig gegessen', carb: 'wenig Kohlenhydrate', protein: 'wenig Eiweiß', fett: 'wenig Fett', schritte: 'wenig Schritte' }[f.id]!;
    const hochLab = { bett: 'spät ins Bett', dauer: 'langer Schlaf', spaet: 'Training ab 18 Uhr', last: 'viel Last', stress: 'viel Stress', kcal: 'viel gegessen', carb: 'viele Kohlenhydrate', protein: 'viel Eiweiß', fett: 'viel Fett', schritte: 'viele Schritte' }[f.id]!;
    const besserNiedrig = (d ?? 0) >= 0;
    const G = besserNiedrig ? lo : hi, S = besserNiedrig ? hi : lo;
    faktoren.push({
      id: f.id, lab: f.lab, gut: besserNiedrig ? niedrigLab : hochLab, schlecht: besserNiedrig ? hochLab : niedrigLab,
      n1: G.length, n2: S.length, hrv1: mittel(h(G)), hrv2: mittel(h(S)), bb1: mittel(b(G)), bb2: mittel(b(S)), d: d == null ? null : Math.abs(d),
    });
  }
  faktoren.sort((a, b) => (b.d ?? 0) - (a.d ?? 0));
  return { faktoren, n: zeilen.length };
}

// ─── 8 Schlafrhythmus ───────────────────────────────────────────────────────
export function schlafrhythmus(tage: GarminTage | null, tl: Map<string, TlZeile> | null) {
  if (!tage) return null;
  const ds = [...tage.keys()].sort();
  const naechte = ds.map((d) => {
    const s = tage.get(d)?.schlaf;
    const bett = zuNachtMin(s?.start), auf = zuNachtMin(s?.ende);
    if (bett == null || auf == null || !s?.dauer_s) return null;
    return { d, bett, auf, mitte: (bett + auf) / 2, score: s.score ?? null, dauer: s.dauer_s / 3600, ready: tl?.get(d)?.ready ?? null, wochenende: wochentag(d) >= 5 };
  }).filter((x): x is NonNullable<typeof x> => !!x && x.bett > -6 * 60 && x.bett < 4 * 60);
  if (naechte.length < 10) return { naechte, n: naechte.length } as any;
  // Abweichung vom eigenen 14-Tage-Median der Schlafmitte
  const mitAbw = naechte.map((n) => {
    const vor = naechte.filter((x) => x.d < n.d && x.d >= plusTage(n.d, -14)).map((x) => x.mitte);
    const m = median(vor);
    return { ...n, abw: m == null ? null : Math.abs(n.mitte - m) };
  });
  const reg = mitAbw.filter((x) => x.abw != null && x.abw <= 30), unreg = mitAbw.filter((x) => x.abw != null && x.abw > 60);
  const letzte28 = naechte.filter((n) => n.d >= plusTage(ds[ds.length - 1], -28));
  const wt = letzte28.filter((n) => !n.wochenende), we = letzte28.filter((n) => n.wochenende);
  return {
    n: naechte.length, naechte: mitAbw,
    bettSd: sd(letzte28.map((n) => n.bett)), aufSd: sd(letzte28.map((n) => n.auf)),
    bettMedian: median(letzte28.map((n) => n.bett)), aufMedian: median(letzte28.map((n) => n.auf)),
    jetlag: wt.length && we.length ? (median(we.map((n) => n.mitte))! - median(wt.map((n) => n.mitte))!) : null,
    regel: { n: reg.length, score: mittel(reg.map((x) => x.score!).filter(Number.isFinite)), ready: mittel(reg.map((x) => x.ready!).filter(Number.isFinite)) },
    unregel: { n: unreg.length, score: mittel(unreg.map((x) => x.score!).filter(Number.isFinite)), ready: mittel(unreg.map((x) => x.ready!).filter(Number.isFinite)) },
  };
}

// ─── 9 Atem/SpO2-Frühwarnung ────────────────────────────────────────────────
export interface WarnTag { d: string; atem: number | null; spo2: number | null; rhr: number | null; hrv: number | null; signale: string[]; stufe: 'gruen' | 'gelb' | 'rot' }
export function fruehwarnung(tage: GarminTage | null): { tage: WarnTag[]; episoden: { von: string; bis: string; stufe: string; signale: string[] }[] } {
  if (!tage) return { tage: [], episoden: [] };
  const ds = [...tage.keys()].sort();
  const val = (d: string) => {
    const g = tage.get(d);
    return { atem: g?.schlaf?.atem ?? g?.tag?.atem_avg ?? null, spo2: g?.schlaf?.spo2 ?? g?.tag?.spo2_avg ?? null, rhr: g?.tag?.rhr ?? null, hrv: g?.hrv?.nacht ?? null };
  };
  const out: WarnTag[] = [];
  for (const d of ds) {
    const v = val(d);
    const basis = ds.filter((x) => x < d && x >= plusTage(d, -28)).map(val);
    if (basis.length < 10) continue;
    const b = (k: keyof typeof v) => median(basis.map((x) => x[k]).filter((x): x is number => x != null));
    const sig: string[] = [];
    const bA = b('atem'), bS = b('spo2'), bR = b('rhr'), bH = b('hrv');
    if (v.atem != null && bA != null && v.atem >= bA + 1) sig.push(`Atmung ${v.atem}/min (sonst ${bA})`);
    if (v.spo2 != null && bS != null && v.spo2 <= bS - 2) sig.push(`SpO2 ${v.spo2} % (sonst ${bS})`);
    if (v.rhr != null && bR != null && v.rhr >= bR + 4) sig.push(`Ruhepuls ${v.rhr} (sonst ${bR})`);
    if (v.hrv != null && bH != null && v.hrv <= bH * 0.85) sig.push(`HRV ${v.hrv} ms (sonst ${bH})`);
    out.push({ d, ...v, signale: sig, stufe: sig.length >= 3 ? 'rot' : sig.length === 2 ? 'gelb' : 'gruen' });
  }
  // Ein einzelner auffälliger Morgen ist oft nur eine kurze Nacht oder ein Glas Wein – rot erst, wenn es sich am Folgetag bestätigt
  for (let i = 0; i < out.length; i++) {
    if (out[i].stufe !== 'rot') continue;
    const vor = out[i - 1], nach = out[i + 1];
    const bestaetigt = (vor && vor.signale.length >= 2 && plusTage(vor.d, 1) === out[i].d) || (nach && nach.signale.length >= 2 && plusTage(out[i].d, 1) === nach.d);
    if (!bestaetigt && nach) out[i] = { ...out[i], stufe: 'gelb' };
  }
  const episoden: { von: string; bis: string; stufe: string; signale: string[] }[] = [];
  for (const t of out) {
    if (t.stufe === 'gruen') continue;
    const l = episoden[episoden.length - 1];
    if (l && plusTage(l.bis, 2) >= t.d) { l.bis = t.d; if (t.stufe === 'rot') l.stufe = 'rot'; l.signale = [...new Set([...l.signale, ...t.signale.map((s) => s.split(' ')[0])])]; }
    else episoden.push({ von: t.d, bis: t.d, stufe: t.stufe, signale: t.signale.map((s) => s.split(' ')[0]) });
  }
  return { tage: out, episoden: episoden.reverse() };
}

// ─── 10 Stress-Bilanz ───────────────────────────────────────────────────────
export function stressBilanz(tage: GarminTage | null) {
  if (!tage) return null;
  const xs = [...tage.entries()].map(([d, g]) => {
    const s = g.tag;
    if (!s?.stress_avg) return null;
    const akts = (g.akt || []).filter((a) => minuten(a) >= 20);
    return { d, wt: wochentag(d), avg: s.stress_avg, hoch: (s.stress_min?.hoch || 0) + (s.stress_min?.mittel || 0), ruhe: s.stress_min?.ruhe || 0, training: akts.length > 0, hartTag: akts.some(hart), bbVerbraucht: s.bb_verbraucht ?? null };
  }).filter((x): x is NonNullable<typeof x> => !!x);
  const proWt = WT.map((lab, i) => {
    const t = xs.filter((x) => x.wt === i);
    return { lab, n: t.length, avg: mittel(t.map((x) => x.avg)), hoch: mittel(t.map((x) => x.hoch)), ruhe: mittel(t.map((x) => x.ruhe)) };
  });
  const gr = (f: (x: (typeof xs)[number]) => boolean) => { const t = xs.filter(f); return { n: t.length, avg: mittel(t.map((x) => x.avg)), hoch: mittel(t.map((x) => x.hoch)) }; };
  const werktagRuhe = gr((x) => x.wt < 5 && !x.training), werktagTraining = gr((x) => x.wt < 5 && x.training), wochenende = gr((x) => x.wt >= 5), hartTage = gr((x) => x.hartTag);
  const monate = new Map<string, number[]>();
  for (const x of xs) monate.set(x.d.slice(0, 7), [...(monate.get(x.d.slice(0, 7)) || []), x.avg]);
  const reihe = [...monate.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([m, v]) => ({ monat: m, avg: mittel(v) }));
  return { n: xs.length, proWt, werktagRuhe, werktagTraining, wochenende, hartTage, reihe };
}

// ─── 11 Energiebilanz ───────────────────────────────────────────────────────
export function energiebilanz(tage: GarminTage | null, tl: Map<string, TlZeile> | null) {
  if (!tage || !tl) return null;
  const ds = [...tl.keys()].sort();
  const zeilen = ds.map((d) => {
    const t = tl.get(d)!, g = tage.get(d);
    const verbrauch = g?.tag?.kcal_gesamt ?? null;
    if (t.kcal_in == null || verbrauch == null || t.kcal_in < 800) return null;
    return { d, in: t.kcal_in, aus: verbrauch, saldo: t.kcal_in - verbrauch, carb: t.carb, protein: t.protein, fett: t.fett };
  }).filter((x): x is NonNullable<typeof x> => !!x);
  const wo = new Map<string, typeof zeilen>();
  for (const z of zeilen) wo.set(wochenStart(z.d), [...(wo.get(wochenStart(z.d)) || []), z]);
  const wochen = [...wo.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([w, xs]) => ({ woche: w, n: xs.length, saldo: xs.reduce((a, x) => a + x.saldo, 0), saldoTag: mittel(xs.map((x) => x.saldo)), in: mittel(xs.map((x) => x.in)), aus: mittel(xs.map((x) => x.aus)) }));
  // Kohlenhydrate vor harten Einheiten: Vortag + Tag (bis zur Einheit)
  const harte = eintraege(tage).filter((e) => hart(e.a) && (LAUF_RAD.has(e.a.typ || '')));
  const kh = harte.map((e) => {
    const v = tl.get(plusTage(e.datum, -1));
    const eff = schlaegeProKm(e.a);
    const nb = eintraege(tage).filter((q) => q.a.typ === e.a.typ && Math.abs(tageSeit(q.datum) - tageSeit(e.datum)) <= 45 && q.a.hf_avg).map((q) => schlaegeProKm(q.a)!).filter(Number.isFinite);
    const m = median(nb);
    const morgen = tl.get(plusTage(e.datum, 1))?.ready ?? null;
    return { d: e.datum, carb: v?.carb ?? null, effRel: eff && m ? ((eff - m) / m) * 100 : null, ready: morgen };
  }).filter((x) => x.carb != null);
  const sch = median(kh.map((x) => x.carb!));
  const wenig = kh.filter((x) => sch != null && x.carb! < (sch as number)), viel = kh.filter((x) => sch != null && x.carb! >= (sch as number));
  const g = (xs: typeof kh) => ({ n: xs.length, carb: median(xs.map((x) => x.carb!)), effRel: median(xs.map((x) => x.effRel!).filter(Number.isFinite)), ready: mittel(xs.map((x) => x.ready!).filter(Number.isFinite)) });
  const letzte4 = wochen.slice(-4);
  return { n: zeilen.length, zeilen, wochen, khVorHart: { wenig: g(wenig), viel: g(viel), schwelle: sch }, saldo4: letzte4.length ? mittel(letzte4.map((w) => w.saldoTag!)) : null };
}
const LAUF_RAD = new Set(['running', 'road_biking', 'cycling', 'gravel_cycling', 'mountain_biking', 'indoor_cycling', 'virtual_ride']);

// ─── 12 Eiweiß und Erholung ─────────────────────────────────────────────────
export function eiweiss(tage: GarminTage | null, tl: Map<string, TlZeile> | null) {
  if (!tage || !tl) return null;
  const ds = [...tl.keys()].sort();
  const rows = ds.map((d) => {
    const t = tl.get(d)!, n = tl.get(plusTage(d, 1)), gn = tage.get(plusTage(d, 1));
    if (t.protein == null || t.protein < 20) return null;
    return { d, protein: t.protein, last: t.last ?? 0, ready: n?.ready ?? null, hrv: gn?.hrv?.nacht ?? null, bb: gn?.bb_aufwachen ?? null };
  }).filter((x): x is NonNullable<typeof x> => !!x && (x.ready != null || x.hrv != null));
  if (rows.length < 15) return { n: rows.length, terzile: [], hochLast: [] } as any;
  const q1 = quantil(rows.map((r) => r.protein), 1 / 3)!, q2 = quantil(rows.map((r) => r.protein), 2 / 3)!;
  const terz = (xs: typeof rows) => [
    { lab: 'wenig', g: `< ${Math.round(q1)} g`, xs: xs.filter((r) => r.protein < q1) },
    { lab: 'mittel', g: `${Math.round(q1)}–${Math.round(q2)} g`, xs: xs.filter((r) => r.protein >= q1 && r.protein < q2) },
    { lab: 'viel', g: `ab ${Math.round(q2)} g`, xs: xs.filter((r) => r.protein >= q2) },
  ].map((t) => ({ lab: `${t.lab} (${t.g})`, kurz: t.lab, n: t.xs.length, protein: median(t.xs.map((r) => r.protein)), ready: mittel(t.xs.map((r) => r.ready!).filter(Number.isFinite)), hrv: mittel(t.xs.map((r) => r.hrv!).filter(Number.isFinite)), bb: mittel(t.xs.map((r) => r.bb!).filter(Number.isFinite)) }));
  const mLast = median(rows.map((r) => r.last))!;
  return { n: rows.length, terzile: terz(rows), hochLast: terz(rows.filter((r) => r.last >= mLast)), mLast, proteinMedian: median(rows.map((r) => r.protein)) };
}
