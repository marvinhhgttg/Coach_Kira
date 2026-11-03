/**
 * Trainingsanalysen aus garmin_tag: Entkopplung, Temperatur, Tageszeit, Rad-Leistungskurve,
 * Trainingswirkung, Laufökonomie und Hausrunden. Rein deterministisch.
 */
import type { GarminAkt, GarminTage } from './garmin';
import {
  BERG, LAUF, RAD, RAD_ALLE, eintraege, flachGeeignet, hart, median, minuten, mittel, plusTage, schlaegeProKm,
  startMin, tageSeit, theilSen, vBewegt, wochenStart, type E,
} from './stat';

export type Sport = 'lauf' | 'rad';

// ─── 1 Aerobe Entkopplung ───────────────────────────────────────────────────
/** Vom Mac aus den Runden berechnet: Effizienz (Tempo bzw. Leistung je Herzschlag) 1. vs. 2. Hälfte. */
export interface Entkopplung { pct: number; basis: 'tempo' | 'watt'; hf1: number; hf2: number; x1: number; x2: number; runden: number }
type AktMitEnt = GarminAkt & { ent?: Entkopplung | null };

export function entkopplung(tage: GarminTage | null, sport: Sport) {
  const es = eintraege(tage).filter((e) => (sport === 'lauf' ? LAUF : RAD_ALLE).has(e.a.typ || '') && minuten(e.a) >= 60);
  const mitEnt = es.filter((e) => (e.a as AktMitEnt).ent && Number.isFinite((e.a as AktMitEnt).ent!.pct));
  const flach = mitEnt.filter((e) => (e.a.hm || 0) / Math.max(e.a.km || 1, 1) <= (sport === 'lauf' ? 5 : 4) || (e.a.typ || '').startsWith('indoor') || e.a.typ === 'virtual_ride');
  const pts = flach.map((e) => ({ datum: e.datum, t: tageSeit(e.datum), y: (e.a as AktMitEnt).ent!.pct, a: e.a as AktMitEnt }));
  const ts = theilSen(pts);
  const letzte = pts.slice(-6).map((p) => p.y);
  return {
    kandidaten: es.length, mitDaten: mitEnt.length, punkte: pts,
    trendJe90: ts ? ts.b * 90 : null,
    median: median(pts.map((p) => p.y)), aktuell: median(letzte),
    unter5: pts.filter((p) => p.y < 5).length,
  };
}

// ─── 2 Hitze und Kälte ──────────────────────────────────────────────────────
export const TEMP_BINS = [
  { lab: 'unter 5 °C', lo: -50, hi: 5 }, { lab: '5–10 °C', lo: 5, hi: 10 }, { lab: '10–15 °C', lo: 10, hi: 15 },
  { lab: '15–20 °C', lo: 15, hi: 20 }, { lab: '20–25 °C', lo: 20, hi: 25 }, { lab: 'über 25 °C', lo: 25, hi: 60 },
];
/** Orte im Raum Pfaffenhofen/Ulm – dort passt die Wetterstation (Open-Meteo, Kadeltshofen). */
export const HEIM_ORTE = /Pfaffenhofen|Ulm|Kadeltshofen|Holzheim|Senden|Breitenthal|Erbach|Wei(ß|ss)enhorn|Nersingen|Roggenburg|Vöhringen|Illertissen/i;
/** Lufttemperatur zur Startstunde aus Open-Meteo (Schlüssel „YYYY-MM-DDTHH“), Browser-Cache 7 Tage. */
export async function fetchWetterStunden(von: string, bis: string): Promise<Map<string, number>> {
  const KEY = 'ck_wetter_stunden_v1';
  try {
    const c = JSON.parse(localStorage.getItem(KEY) || 'null');
    if (c && c.von <= von && c.bis >= plusTage(bis, -6) && Date.now() - c.at < 7 * 86400000) return new Map(Object.entries(c.t));
  } catch { /* ignore */ }
  const url = `https://archive-api.open-meteo.com/v1/archive?latitude=48.33&longitude=10.12&start_date=${von}&end_date=${bis}&hourly=temperature_2m&timezone=Europe%2FBerlin`;
  const r = await fetch(url);
  const j = await r.json();
  if (!r.ok || j?.error) throw new Error(j?.reason || `Open-Meteo HTTP ${r.status}`);
  const m = new Map<string, number>();
  const ts: string[] = j?.hourly?.time || [], vs: (number | null)[] = j?.hourly?.temperature_2m || [];
  ts.forEach((t, i) => { if (vs[i] != null) m.set(t.slice(0, 13), vs[i] as number); });
  try { localStorage.setItem(KEY, JSON.stringify({ at: Date.now(), von, bis, t: Object.fromEntries(m) })); } catch { /* ignore */ }
  return m;
}
const tempFuer = (e: E, wetter: Map<string, number> | null) => {
  const sm = startMin(e.a);
  if (wetter && sm != null && HEIM_ORTE.test(e.a.ort || '')) {
    const h = Math.min(23, Math.floor((sm + minuten(e.a) / 2) / 60)); // Mitte der Einheit
    const t = wetter.get(`${e.datum}T${String(h).padStart(2, '0')}`);
    if (t != null) return { t, quelle: 'wetter' as const };
  }
  return null;
};
export function temperatur(tage: GarminTage | null, sport: Sport, wetter: Map<string, number> | null) {
  const es = eintraege(tage).filter((e) => flachGeeignet(e.a, sport) && tempFuer(e, wetter));
  const pts = es.map((e) => ({ datum: e.datum, temp: tempFuer(e, wetter)!.t, y: schlaegeProKm(e.a)!, hf: e.a.hf_avg!, v: vBewegt(e.a) }));
  // Formtrend herausrechnen: Abweichung vom eigenen gleitenden Median (± 30 Tage)
  const rel = pts.map((p) => {
    const nb = pts.filter((q) => Math.abs(tageSeit(q.datum) - tageSeit(p.datum)) <= 30).map((q) => q.y);
    const m = median(nb) || p.y;
    return { ...p, relPct: ((p.y - m) / m) * 100 };
  });
  const bins = TEMP_BINS.map((b) => {
    const xs = rel.filter((p) => p.temp >= b.lo && p.temp < b.hi);
    return { ...b, n: xs.length, schlaege: median(xs.map((p) => p.y)), relPct: median(xs.map((p) => p.relPct)), hf: median(xs.map((p) => p.hf)), v: median(xs.map((p) => p.v)) };
  });
  const ts = theilSen(rel.map((p) => ({ t: p.temp, y: p.relPct })));
  return { n: pts.length, punkte: rel, bins, proGrad: ts ? ts.b : null, optimum: bins.filter((b) => b.n >= 4).sort((a, b) => (a.relPct ?? 99) - (b.relPct ?? 99))[0] || null };
}

// ─── 3 Tageszeit ────────────────────────────────────────────────────────────
export const ZEITEN = [
  { id: 'frueh', lab: 'früh (vor 9 Uhr)', lo: 0, hi: 9 * 60 }, { id: 'vormittag', lab: 'Vormittag (9–12)', lo: 9 * 60, hi: 12 * 60 },
  { id: 'mittag', lab: 'Mittag (12–15)', lo: 12 * 60, hi: 15 * 60 }, { id: 'nachmittag', lab: 'Nachmittag (15–18)', lo: 15 * 60, hi: 18 * 60 },
  { id: 'abend', lab: 'Abend (ab 18 Uhr)', lo: 18 * 60, hi: 24 * 60 },
];
export function tageszeit(tage: GarminTage | null, sport: Sport) {
  const es = eintraege(tage).filter((e) => flachGeeignet(e.a, sport) && startMin(e.a) != null);
  const pts = es.map((e) => {
    const nb = es.filter((q) => Math.abs(tageSeit(q.datum) - tageSeit(e.datum)) <= 30).map((q) => schlaegeProKm(q.a)!);
    const m = median(nb)!;
    const y = schlaegeProKm(e.a)!;
    return { datum: e.datum, min: startMin(e.a)!, y, relPct: ((y - m) / m) * 100 };
  });
  const effizienz = ZEITEN.map((z) => {
    const xs = pts.filter((p) => p.min >= z.lo && p.min < z.hi);
    return { ...z, n: xs.length, relPct: median(xs.map((p) => p.relPct)), schlaege: median(xs.map((p) => p.y)) };
  });
  // Schlaf danach: Nacht nach dem Tag (garmin_tag des Folgetags)
  const gruppen: Record<'abend' | 'frueher' | 'ruhe', { score: number[]; hrv: number[]; tief: number[]; bb: number[] }> = {
    abend: { score: [], hrv: [], tief: [], bb: [] }, frueher: { score: [], hrv: [], tief: [], bb: [] }, ruhe: { score: [], hrv: [], tief: [], bb: [] },
  };
  if (tage) for (const [d, g] of tage) {
    const n = tage.get(plusTage(d, 1));
    if (!n?.schlaf?.score) continue;
    const akts = (g.akt || []).filter((a) => minuten(a) >= 20);
    const k = !akts.length ? 'ruhe' : akts.some((a) => (startMin(a) ?? 0) >= 18 * 60) ? 'abend' : 'frueher';
    gruppen[k].score.push(n.schlaf.score);
    if (n.hrv?.nacht) gruppen[k].hrv.push(n.hrv.nacht);
    if (n.schlaf.tief_s) gruppen[k].tief.push(n.schlaf.tief_s / 60);
    if (n.bb_aufwachen) gruppen[k].bb.push(n.bb_aufwachen);
  }
  const schlaf = (Object.keys(gruppen) as (keyof typeof gruppen)[]).map((k) => ({
    id: k, lab: k === 'abend' ? 'nach Abend-Einheit' : k === 'frueher' ? 'nach Einheit tagsüber' : 'Ruhetag',
    n: gruppen[k].score.length, score: mittel(gruppen[k].score), hrv: mittel(gruppen[k].hrv), tief: mittel(gruppen[k].tief), bb: mittel(gruppen[k].bb),
  }));
  return { n: pts.length, effizienz, schlaf };
}

// ─── 4 Rad-Leistungskurve ───────────────────────────────────────────────────
export function leistungskurve(tage: GarminTage | null) {
  const es = eintraege(tage).filter((e) => RAD_ALLE.has(e.a.typ || ''));
  const monate = new Map<string, { w20: number[]; np: number[]; tss: number; if: number[]; n: number }>();
  for (const e of es) {
    const k = e.datum.slice(0, 7);
    const m = monate.get(k) || { w20: [], np: [], tss: 0, if: [], n: 0 };
    m.n++;
    if (e.a.watt_20) m.w20.push(e.a.watt_20);
    if (e.a.watt_np && minuten(e.a) >= 45) m.np.push(e.a.watt_np);
    if (e.a.tss) m.tss += e.a.tss;
    if (e.a.if) m.if.push(e.a.if);
    monate.set(k, m);
  }
  const ks = [...monate.keys()].sort();
  // FTP-Schätzung: 95 % der besten 20 Minuten der letzten 90 Tage, je Monatsende
  const reihe = ks.map((k) => {
    const ende = k + '-28';
    const best90 = Math.max(0, ...es.filter((e) => e.a.watt_20 && e.datum <= ende && e.datum >= plusTage(ende, -90)).map((e) => e.a.watt_20!));
    const m = monate.get(k)!;
    return {
      monat: k, n: m.n, best20: m.w20.length ? Math.max(...m.w20) : null, npMedian: median(m.np),
      ftp: best90 > 0 ? Math.round(best90 * 0.95) : null, tss: Math.round(m.tss), ifMedian: median(m.if),
    };
  });
  // TSS je Woche (letzte 16 Wochen)
  const wo = new Map<string, number>();
  for (const e of es) if (e.a.tss) wo.set(wochenStart(e.datum), (wo.get(wochenStart(e.datum)) || 0) + e.a.tss);
  const wochen = [...wo.entries()].sort((a, b) => a[0].localeCompare(b[0])).slice(-16).map(([w, t]) => ({ woche: w, tss: Math.round(t) }));
  const best = es.filter((e) => e.a.watt_20).sort((a, b) => b.a.watt_20! - a.a.watt_20!)[0];
  const ftpJetzt = reihe.length ? reihe[reihe.length - 1].ftp : null;
  const ftpStart = reihe.find((r) => r.ftp)?.ftp ?? null;
  return { reihe, wochen, best: best ? { datum: best.datum, w: best.a.watt_20!, name: best.a.name } : null, ftpJetzt, ftpStart, n: es.length };
}

// ─── 5 Trainingswirkung ─────────────────────────────────────────────────────
export const TE_LABELS: Record<string, { lab: string; farbe: string }> = {
  RECOVERY: { lab: 'Erholung', farbe: '#94a3b8' }, AEROBIC_BASE: { lab: 'Grundlage', farbe: '#38bdf8' }, TEMPO: { lab: 'Tempo', farbe: '#22c55e' },
  LACTATE_THRESHOLD: { lab: 'Schwelle', farbe: '#eab308' }, THRESHOLD: { lab: 'Schwelle', farbe: '#eab308' }, VO2MAX: { lab: 'VO2max', farbe: '#f97316' },
  ANAEROBIC_CAPACITY: { lab: 'Anaerob', farbe: '#ef4444' }, SPRINT: { lab: 'Sprint', farbe: '#a855f7' }, NO_BENEFIT: { lab: 'ohne Effekt', farbe: '#475569' },
};
export function trainingswirkung(tage: GarminTage | null, wochenAnzahl = 16) {
  const es = eintraege(tage).filter((e) => minuten(e.a) >= 15 && (e.a.te_ae != null || e.a.te_label));
  const wo = new Map<string, { labels: Record<string, number>; ae: number; an: number; qualitaet: number; n: number }>();
  for (const e of es) {
    const w = wochenStart(e.datum);
    const x = wo.get(w) || { labels: {}, ae: 0, an: 0, qualitaet: 0, n: 0 };
    const lab = TE_LABELS[e.a.te_label || ''] ? TE_LABELS[e.a.te_label || ''].lab : 'sonstige';
    x.labels[lab] = (x.labels[lab] || 0) + 1;
    x.ae += e.a.te_ae || 0; x.an += e.a.te_an || 0; x.n++;
    if (hart(e.a)) x.qualitaet++;
    wo.set(w, x);
  }
  const wochen = [...wo.entries()].sort((a, b) => a[0].localeCompare(b[0])).slice(-wochenAnzahl).map(([w, x]) => ({ woche: w, ...x }));
  const letzte4 = wochen.slice(-4);
  const summe = (f: (x: (typeof wochen)[number]) => number) => letzte4.reduce((a, x) => a + f(x), 0);
  const labelsGesamt: Record<string, number> = {};
  for (const w of letzte4) for (const [k, v] of Object.entries(w.labels)) labelsGesamt[k] = (labelsGesamt[k] || 0) + v;
  const qualiJeWoche = letzte4.length ? summe((x) => x.qualitaet) / letzte4.length : null;
  const anJeWoche = letzte4.length ? summe((x) => x.an) / letzte4.length : null;
  const urteil = qualiJeWoche == null ? '' : qualiJeWoche < 1 ? 'zu wenig Qualitätsreize' : qualiJeWoche > 3 ? 'viele harte Einheiten – auf Erholung achten' : 'ausgewogen';
  return { wochen, labelsGesamt, qualiJeWoche, anJeWoche, urteil, n: es.length };
}

// ─── 6 Laufökonomie ─────────────────────────────────────────────────────────
export function laufoekonomie(tage: GarminTage | null) {
  const es = eintraege(tage).filter((e) => flachGeeignet(e.a, 'lauf') && e.a.kadenz && e.a.kadenz > 120 && e.a.kadenz < 220);
  const pts = es.map((e) => {
    const v = vBewegt(e.a);
    return { datum: e.datum, t: tageSeit(e.datum), kadenz: e.a.kadenz!, schritt: (v * 1000 / 60) / e.a.kadenz!, tempo: 60 / v, hf: e.a.hf_avg!, schlaege: schlaegeProKm(e.a)! };
  });
  // Kadenz bei gleichem Tempo: Residuum einer robusten Geraden Kadenz ~ Tempo
  const kv = theilSen(pts.map((p) => ({ t: p.tempo, y: p.kadenz })));
  const mitRes = pts.map((p) => ({ ...p, kadenzBereinigt: kv ? p.kadenz - (kv.a + kv.b * p.tempo) : 0 }));
  const monate = new Map<string, typeof mitRes>();
  for (const p of mitRes) monate.set(p.datum.slice(0, 7), [...(monate.get(p.datum.slice(0, 7)) || []), p]);
  const reihe = [...monate.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([m, xs]) => ({
    monat: m, n: xs.length, kadenz: median(xs.map((x) => x.kadenz)), schritt: median(xs.map((x) => x.schritt)), tempo: median(xs.map((x) => x.tempo)),
    hf: median(xs.map((x) => x.hf)), kadenzBereinigt: median(xs.map((x) => x.kadenzBereinigt)),
  }));
  const trK = theilSen(mitRes.map((p) => ({ t: p.t, y: p.kadenzBereinigt })));
  const trS = theilSen(mitRes.map((p) => ({ t: p.t, y: p.schritt })));
  // Zusammenhang Kadenz ↔ Effizienz: schneller Schritt = weniger Herzschläge je km?
  const hoch = mitRes.filter((p) => p.kadenzBereinigt > 2).map((p) => p.schlaege), tief = mitRes.filter((p) => p.kadenzBereinigt < -2).map((p) => p.schlaege);
  return {
    n: pts.length, reihe, kadenzProTempo: kv ? kv.b : null,
    kadenzTrend90: trK ? trK.b * 90 : null, schrittTrend90: trS ? trS.b * 90 : null,
    effHoch: median(hoch), effTief: median(tief), nHoch: hoch.length, nTief: tief.length,
  };
}

// ─── 15 Hausrunden ──────────────────────────────────────────────────────────
export interface Runde {
  id: string; sport: string; ort: string; km: number; hm: number; n: number;
  laeufe: { datum: string; a: GarminAkt; v: number; schlaege: number | null }[];
  tempoTrendPct: number | null; effTrendPct: number | null; bestV: number; letzteV: number; medianV: number;
}
const sportGruppe = (t: string) => (LAUF.has(t) ? 'Laufen' : RAD.has(t) ? (t === 'gravel_cycling' || t === 'mountain_biking' ? 'Gravel' : 'Rennrad') : BERG.has(t) || t === 'walking' ? 'Wandern' : null);
export function hausrunden(tage: GarminTage | null, minN = 4): Runde[] {
  const es = eintraege(tage).filter((e) => e.a.km && e.a.km >= 3 && sportGruppe(e.a.typ || '') && e.a.bewegt_s);
  const rest = [...es];
  const out: Runde[] = [];
  while (rest.length) {
    const s = rest.shift()!;
    const g = sportGruppe(s.a.typ || '')!, ort = s.a.ort || '?';
    const passt = (e: E) => sportGruppe(e.a.typ || '') === g && (e.a.ort || '?') === ort
      && Math.abs((e.a.km || 0) - (s.a.km || 0)) / (s.a.km || 1) <= 0.06
      && Math.abs((e.a.hm || 0) - (s.a.hm || 0)) <= Math.max(25, 0.2 * (s.a.hm || 0));
    const gruppe = [s, ...rest.filter(passt)];
    for (const e of gruppe.slice(1)) rest.splice(rest.indexOf(e), 1);
    if (gruppe.length < minN) continue;
    const laeufe = gruppe.map((e) => ({ datum: e.datum, a: e.a, v: vBewegt(e.a), schlaege: e.a.hf_avg ? schlaegeProKm(e.a) : null }));
    const mV = median(laeufe.map((l) => l.v))!;
    const tv = theilSen(laeufe.map((l) => ({ t: tageSeit(l.datum), y: l.v })));
    const effs = laeufe.filter((l) => l.schlaege != null);
    const te = theilSen(effs.map((l) => ({ t: tageSeit(l.datum), y: l.schlaege! })));
    const mE = median(effs.map((l) => l.schlaege!));
    const tage0 = tageSeit(laeufe[0].datum), tage1 = tageSeit(laeufe[laeufe.length - 1].datum), span = tage1 - tage0;
    out.push({
      id: `${g}-${ort}-${Math.round(median(laeufe.map((l) => l.a.km!))! * 10)}`, sport: g, ort,
      km: median(laeufe.map((l) => l.a.km!))!, hm: median(laeufe.map((l) => l.a.hm || 0))!, n: laeufe.length, laeufe,
      tempoTrendPct: tv && span >= 30 ? ((tv.b * span) / mV) * 100 : null,
      effTrendPct: te && mE && span >= 30 && effs.length >= 3 ? ((te.b * span) / mE) * 100 : null,
      bestV: Math.max(...laeufe.map((l) => l.v)), letzteV: laeufe[laeufe.length - 1].v, medianV: mV,
    });
  }
  return out.sort((a, b) => b.n - a.n);
}
