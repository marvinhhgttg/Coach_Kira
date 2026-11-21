/**
 * Jahresbericht (Saison-Steckbrief als PDF): Kennzahlen, Diagramme und Analysen als druckfertiges HTML.
 * Rein deterministisch – nur Zahlen aus garmin_tag und der Zeitreihe (Ausdauerwert, CTL, ATL).
 */
import type { GarminAkt, GarminTage } from './garmin';
import { TE_LABEL_DE } from './garmin';
import { rekorde, GRUPPEN } from './rekorde';
import { effPunkte, nachGelaende, trend, wertVon, GRENZEN_STANDARD } from './effizienz';
import { entkopplung, leistungskurve } from './analysenTraining';
import {
  BERG, LAUF, RAD, RAD_ALLE, WT, de, dauerTxt, eintraege, kurz, median, mittel, monatName, plusTage, startMin, vz, wochenStart, wochentag, type E,
} from './stat';

export interface Reihe { d: string; v: number }
export interface Zeitreihen { es: Reihe[]; ctl: Reihe[]; atl: Reihe[] }

const esc = (s: unknown) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));
const GLAB: Record<string, string> = Object.fromEntries(GRUPPEN.map((g) => [g.id, g.label]));

// ─── Sportgruppen und Farben ──────────────────────────────────────────────
export const JB_GRUPPEN = [
  { id: 'lauf', lab: 'Laufen', farbe: '#ef7d00', typen: LAUF },
  { id: 'rad', lab: 'Rad draußen', farbe: '#1976d2', typen: RAD },
  { id: 'indoor', lab: 'Rad indoor', farbe: '#64b5f6', typen: new Set(['indoor_cycling', 'virtual_ride']) },
  { id: 'berg', lab: 'Wandern & Berg', farbe: '#22a15a', typen: new Set([...BERG, 'walking']) },
  { id: 'ski', lab: 'Ski & Skike', farbe: '#7c3aed', typen: new Set(['skate_skiing_ws', 'backcountry_skiing', 'cross_country_skiing_ws', 'resort_skiing_snowboarding_ws']) },
  { id: 'rudern', lab: 'Rudern', farbe: '#0ea5a4', typen: new Set(['indoor_rowing', 'rowing_v2']) },
];
const SONST = { id: 'sonst', lab: 'Sonstiges', farbe: '#9aa5b1' };
const gruppeVon = (a: GarminAkt) => JB_GRUPPEN.find((g) => g.typen.has(a.typ || '')) || SONST;
const ZONEN_FARBE = ['#9aa5b1', '#1976d2', '#22a15a', '#ef7d00', '#d64545'];

// ─── SVG-Bausteine ────────────────────────────────────────────────────────
const W = 700;
const ts = (d: string) => new Date(d + 'T12:00:00').getTime();
function niceTicks(lo: number, hi: number, n = 4): number[] {
  if (!(hi > lo)) return [lo];
  const roh = (hi - lo) / n, p = Math.pow(10, Math.floor(Math.log10(roh)));
  const schritt = [1, 2, 2.5, 5, 10].map((f) => f * p).find((s) => s >= roh) || roh;
  const out: number[] = [];
  for (let v = Math.ceil(lo / schritt) * schritt; v <= hi + 1e-9; v += schritt) out.push(Math.round(v * 1000) / 1000 || 0);
  return out;
}
function monatsTicks(von: string, bis: string): { t: number; lab: string }[] {
  const out: { t: number; lab: string }[] = [];
  const d = new Date(von.slice(0, 7) + '-01T12:00:00');
  const ende = ts(bis);
  const anzahl = Math.max(1, Math.round((ende - d.getTime()) / (30.4 * 86400000)));
  const jede = anzahl > 14 ? 2 : 1;
  let i = 0;
  while (d.getTime() <= ende) {
    if (d.getTime() >= ts(von) - 15 * 86400000 && i % jede === 0) out.push({ t: d.getTime(), lab: d.toLocaleDateString('de-DE', { month: 'short' }) + (d.getMonth() === 0 || out.length === 0 ? ' ' + String(d.getFullYear()).slice(2) : '') });
    d.setMonth(d.getMonth() + 1); i++;
  }
  return out;
}

/** Zeitreihen-Liniendiagramm mit Achsen, Gitter und optionalen Markierungen. */
export function svgZeit(serien: { name: string; farbe: string; xs: Reihe[]; breite?: number; flaeche?: boolean; strich?: string }[], opt: { hoehe?: number; einheit?: string; von?: string; bis?: string; marken?: { d: string; v: number; text: string; farbe?: string }[]; null0?: boolean; dez?: number; breite?: number } = {}) {
  const W = opt.breite ?? 700;
  const H = opt.hoehe ?? 150, L = 44, R = 12, T = 12, B = 22;
  const alle = serien.flatMap((s) => s.xs);
  if (alle.length < 2) return '<div class="leer">zu wenige Daten</div>';
  const von = opt.von || alle.reduce((a, x) => (x.d < a ? x.d : a), alle[0].d);
  const bis = opt.bis || alle.reduce((a, x) => (x.d > a ? x.d : a), alle[0].d);
  let lo = Math.min(...alle.map((x) => x.v)), hi = Math.max(...alle.map((x) => x.v));
  if (opt.null0) lo = Math.min(0, lo);
  const pad = (hi - lo) * 0.08 || 1; lo -= opt.null0 && lo === 0 ? 0 : pad; hi += pad;
  const t0 = ts(von), t1 = ts(bis) || t0 + 1;
  const x = (d: string) => L + ((ts(d) - t0) / Math.max(1, t1 - t0)) * (W - L - R);
  const xt = (t: number) => L + ((t - t0) / Math.max(1, t1 - t0)) * (W - L - R);
  const y = (v: number) => T + (1 - (v - lo) / (hi - lo)) * (H - T - B);
  const ticks = niceTicks(lo, hi);
  let s = `<svg class="ch" viewBox="0 0 ${W} ${H}" width="100%">`;
  for (const v of ticks) s += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#e8e8e8"/><text x="${L - 6}" y="${y(v) + 3}" text-anchor="end">${de(v, opt.dez ?? 0)}</text>`;
  for (const m of monatsTicks(von, bis)) if (xt(m.t) >= L - 1) s += `<line x1="${xt(m.t)}" x2="${xt(m.t)}" y1="${H - B}" y2="${H - B + 4}" stroke="#bbb"/><text x="${xt(m.t)}" y="${H - 6}" text-anchor="middle">${esc(m.lab)}</text>`;
  s += `<line x1="${L}" x2="${W - R}" y1="${H - B}" y2="${H - B}" stroke="#bbb"/>`;
  for (const se of serien) {
    const xs = [...se.xs].sort((a, b) => a.d.localeCompare(b.d));
    if (xs.length < 2) continue;
    const pts = xs.map((p) => `${x(p.d).toFixed(1)},${y(p.v).toFixed(1)}`).join(' ');
    if (se.flaeche) s += `<polygon points="${x(xs[0].d)},${H - B} ${pts} ${x(xs[xs.length - 1].d)},${H - B}" fill="${se.farbe}" fill-opacity="0.12"/>`;
    s += `<polyline points="${pts}" fill="none" stroke="${se.farbe}" stroke-width="${se.breite ?? 2}" stroke-linejoin="round"${se.strich ? ` stroke-dasharray="${se.strich}"` : ''}/>`;
  }
  for (const m of opt.marken || []) {
    const mx = x(m.d), my = y(m.v), rechts = mx > W - 130;
    s += `<circle cx="${mx}" cy="${my}" r="3.5" fill="${m.farbe || '#111'}"/><text x="${mx + (rechts ? -6 : 6)}" y="${my - 6}" text-anchor="${rechts ? 'end' : 'start'}" class="mk">${esc(m.text)}</text>`;
  }
  if (opt.einheit) s += `<text x="${L}" y="9" class="ax">${esc(opt.einheit)}</text>`;
  if (serien.length > 1) {
    let lx = W - R;
    for (const se of [...serien].reverse()) { const w = se.name.length * 5.4 + 18; lx -= w; s += `<rect x="${lx}" y="2" width="10" height="3" fill="${se.farbe}"/><text x="${lx + 13}" y="8">${esc(se.name)}</text>`; }
  }
  return s + '</svg>';
}

/** Gestapelte Monatsbalken. */
function svgStapel(spalten: { lab: string; teile: { v: number; farbe: string }[] }[], opt: { hoehe?: number; einheit?: string; summeText?: (v: number) => string } = {}) {
  const H = opt.hoehe ?? 170, L = 34, T = 16, B = 22;
  const max = Math.max(1, ...spalten.map((c) => c.teile.reduce((a, t) => a + t.v, 0)));
  const ticks = niceTicks(0, max), top = Math.max(max, ticks[ticks.length - 1]);
  const bw = (W - L - 6) / Math.max(1, spalten.length);
  const y = (v: number) => T + (1 - v / top) * (H - T - B);
  let s = `<svg class="ch" viewBox="0 0 ${W} ${H}" width="100%">`;
  for (const v of ticks) s += `<line x1="${L}" x2="${W - 6}" y1="${y(v)}" y2="${y(v)}" stroke="#eee"/><text x="${L - 5}" y="${y(v) + 3}" text-anchor="end">${de(v)}</text>`;
  spalten.forEach((c, i) => {
    let acc = 0;
    const x0 = L + i * bw + bw * 0.14, w = bw * 0.72;
    for (const t of c.teile) { if (t.v <= 0) continue; s += `<rect x="${x0}" y="${y(acc + t.v)}" width="${w}" height="${y(acc) - y(acc + t.v)}" fill="${t.farbe}"/>`; acc += t.v; }
    if (acc > 0) s += `<text x="${x0 + w / 2}" y="${y(acc) - 3}" text-anchor="middle" class="mk">${esc(opt.summeText ? opt.summeText(acc) : de(acc))}</text>`;
    s += `<text x="${x0 + w / 2}" y="${H - 6}" text-anchor="middle">${esc(c.lab)}</text>`;
  });
  if (opt.einheit) s += `<text x="${L}" y="9" class="ax">${esc(opt.einheit)}</text>`;
  return s + '</svg>';
}

/** Kalender-Heatmap der Tageslast (Spalten = Wochen, Zeilen = Mo–So). */
function svgKalender(tage: Map<string, number>, von: string, bis: string) {
  const start = wochenStart(von);
  const wochen = Math.ceil((ts(bis) - ts(start)) / (7 * 86400000)) + 1;
  const L = 22, T = 14, gap = 2;
  const zelle = Math.min(13, (W - L - 4) / wochen - gap);
  const H = T + 7 * (zelle + gap) + 18;
  const werte = [...tage.values()].filter((v) => v > 0).sort((a, b) => a - b);
  const q = (p: number) => werte.length ? werte[Math.min(werte.length - 1, Math.floor(p * werte.length))] : 1;
  const stufen = [q(0.25), q(0.5), q(0.75), q(0.92)];
  const farben = ['#ebeef2', '#cfe3f7', '#8fc0ec', '#3d8fd9', '#1565c0', '#0b3d7a'];
  const farbe = (v: number) => (v <= 0 ? farben[0] : v < stufen[0] ? farben[1] : v < stufen[1] ? farben[2] : v < stufen[2] ? farben[3] : v < stufen[3] ? farben[4] : farben[5]);
  let s = `<svg class="ch" viewBox="0 0 ${W} ${H}" width="100%">`;
  ['Mo', '', 'Mi', '', 'Fr', '', 'So'].forEach((l, i) => { if (l) s += `<text x="0" y="${T + i * (zelle + gap) + zelle - 2}">${l}</text>`; });
  let letzterMonat = '';
  for (let w = 0; w < wochen; w++) {
    const ws = plusTage(start, w * 7);
    const x = L + w * (zelle + gap);
    const m = ws.slice(0, 7);
    if (m !== letzterMonat && plusTage(ws, 6) >= von) { s += `<text x="${x}" y="9">${esc(new Date(ws + 'T12:00:00').toLocaleDateString('de-DE', { month: 'short' }))}</text>`; letzterMonat = m; }
    for (let i = 0; i < 7; i++) {
      const d = plusTage(ws, i);
      if (d < von || d > bis) continue;
      s += `<rect x="${x}" y="${T + i * (zelle + gap)}" width="${zelle}" height="${zelle}" rx="2" fill="${farbe(tage.get(d) || 0)}"/>`;
    }
  }
  const ly = H - 6;
  s += `<text x="${W - 168}" y="${ly}">weniger</text>`;
  farben.forEach((f, i) => { s += `<rect x="${W - 128 + i * 14}" y="${ly - 9}" width="11" height="11" rx="2" fill="${f}"/>`; });
  s += `<text x="${W - 40}" y="${ly}">mehr</text>`;
  return s + '</svg>';
}

/** Horizontale Balken mit Beschriftung. */
function svgHBalken(xs: { lab: string; v: number; farbe?: string; text?: string }[], opt: { breite?: number } = {}) {
  const BW = opt.breite ?? 340, zeile = 18, L = 96, H = xs.length * zeile + 4;
  const max = Math.max(1, ...xs.map((x) => x.v));
  let s = `<svg class="ch" viewBox="0 0 ${BW} ${H}" width="100%">`;
  xs.forEach((x, i) => {
    const w = (x.v / max) * (BW - L - 60);
    s += `<text x="${L - 6}" y="${i * zeile + 13}" text-anchor="end">${esc(x.lab)}</text><rect x="${L}" y="${i * zeile + 3}" width="${Math.max(0, w)}" height="${zeile - 6}" rx="2" fill="${x.farbe || '#1976d2'}"/><text x="${L + w + 5}" y="${i * zeile + 13}" class="mk">${esc(x.text ?? de(x.v))}</text>`;
  });
  return s + '</svg>';
}

/** Ring (Donut) mit Legende. */
function svgDonut(xs: { lab: string; v: number; farbe: string }[], mitte: string, unter: string) {
  const S = 340, H = 170, cx = 85, cy = 85, r = 62, sw = 24, u = 2 * Math.PI * r;
  const sum = xs.reduce((a, x) => a + x.v, 0) || 1;
  let acc = 0, s = `<svg class="ch" viewBox="0 0 ${S} ${H}" width="100%"><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#eee" stroke-width="${sw}"/>`;
  for (const x of xs) {
    const l = (x.v / sum) * u;
    s += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${x.farbe}" stroke-width="${sw}" stroke-dasharray="${l} ${u - l}" stroke-dashoffset="${-acc}" transform="rotate(-90 ${cx} ${cy})"/>`;
    acc += l;
  }
  s += `<text x="${cx}" y="${cy + 2}" text-anchor="middle" class="big">${esc(mitte)}</text><text x="${cx}" y="${cy + 18}" text-anchor="middle">${esc(unter)}</text>`;
  xs.forEach((x, i) => { s += `<rect x="180" y="${18 + i * 20}" width="11" height="11" rx="2" fill="${x.farbe}"/><text x="197" y="${27 + i * 20}">${esc(x.lab)} · ${de((x.v / sum) * 100)} %</text>`; });
  return s + '</svg>';
}

/** Punktwolke über die Zeit (Entkopplung). */
function svgPunkte(serien: { name: string; farbe: string; xs: Reihe[] }[], opt: { von: string; bis: string; hoehe?: number; einheit?: string; linie?: number }) {
  const H = opt.hoehe ?? 140, L = 44, R = 12, T = 12, B = 22;
  const alle = serien.flatMap((s) => s.xs);
  if (!alle.length) return '<div class="leer">keine Daten</div>';
  let lo = Math.min(0, ...alle.map((p) => p.v)), hi = Math.max(opt.linie ?? 0, ...alle.map((p) => p.v));
  lo -= 1; hi += 1;
  const t0 = ts(opt.von), t1 = ts(opt.bis);
  const x = (d: string) => L + ((ts(d) - t0) / Math.max(1, t1 - t0)) * (W - L - R);
  const xt = (t: number) => L + ((t - t0) / Math.max(1, t1 - t0)) * (W - L - R);
  const y = (v: number) => T + (1 - (v - lo) / (hi - lo)) * (H - T - B);
  let s = `<svg class="ch" viewBox="0 0 ${W} ${H}" width="100%">`;
  for (const v of niceTicks(lo, hi)) s += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#eee"/><text x="${L - 6}" y="${y(v) + 3}" text-anchor="end">${de(v)}</text>`;
  if (opt.linie != null) s += `<line x1="${L}" x2="${W - R}" y1="${y(opt.linie)}" y2="${y(opt.linie)}" stroke="#22a15a" stroke-dasharray="4 3"/><text x="${W - R}" y="${y(opt.linie) - 3}" text-anchor="end" class="mk" fill="#22a15a">${de(opt.linie)} % Grenze</text>`;
  for (const m of monatsTicks(opt.von, opt.bis)) if (xt(m.t) >= L - 1) s += `<text x="${xt(m.t)}" y="${H - 6}" text-anchor="middle">${esc(m.lab)}</text>`;
  s += `<line x1="${L}" x2="${W - R}" y1="${H - B}" y2="${H - B}" stroke="#bbb"/>`;
  for (const se of serien) for (const p of se.xs) s += `<circle cx="${x(p.d)}" cy="${y(p.v)}" r="3.6" fill="${se.farbe}" fill-opacity="0.85"/>`;
  if (opt.einheit) s += `<text x="${L}" y="9" class="ax">${esc(opt.einheit)}</text>`;
  let lx = W - R;
  for (const se of [...serien].reverse()) { const w = se.name.length * 5.4 + 18; lx -= w; s += `<circle cx="${lx + 4}" cy="5" r="4" fill="${se.farbe}"/><text x="${lx + 12}" y="8">${esc(se.name)}</text>`; }
  return s + '</svg>';
}

// ─── Kennzahlen ───────────────────────────────────────────────────────────
const monate = (von: string, bis: string) => {
  const out: string[] = [];
  const d = new Date(von.slice(0, 7) + '-15T12:00:00');
  while (d.toLocaleDateString('sv-SE').slice(0, 7) <= bis.slice(0, 7)) { out.push(d.toLocaleDateString('sv-SE').slice(0, 7)); d.setMonth(d.getMonth() + 1); }
  return out;
};
const monatsMedian = (xs: Reihe[], ms: string[]) => ms.map((m) => ({ d: m + '-15', v: median(xs.filter((x) => x.d.startsWith(m)).map((x) => x.v)) })).filter((x): x is Reihe => x.v != null);
const tagLast = (g: { akt?: GarminAkt[]; summe?: { last?: number } } | undefined) => g?.summe?.last ?? (g?.akt || []).reduce((a, x) => a + (x.last || 0), 0);

export function jahresberichtDaten(tage: GarminTage, z: Zeitreihen, von: string, bis: string) {
  const t: GarminTage = new Map([...tage].filter(([d]) => d >= von && d <= bis));
  const ds = [...t.keys()].sort();
  const anfang = ds[0] || von, ende = ds[ds.length - 1] || bis;
  const es = eintraege(t);
  const sum = (xs: E[], f: (a: GarminAkt) => number | undefined) => xs.reduce((a, e) => a + (f(e.a) || 0), 0);
  const ges = { n: es.length, km: sum(es, (a) => a.km), hm: sum(es, (a) => a.hm), s: sum(es, (a) => a.dauer_s), kcal: sum(es, (a) => a.kcal), last: sum(es, (a) => a.last) };
  const aktiv = ds.filter((d) => (t.get(d)?.akt || []).length > 0);
  const kalTage = Math.round((ts(ende) - ts(anfang)) / 86400000) + 1;
  // Serien
  let serie = 0, best = 0, bestBis = '', pause = 0, bestPause = 0;
  for (let d = anfang; d <= ende; d = plusTage(d, 1)) {
    if ((t.get(d)?.akt || []).length) { serie++; pause = 0; if (serie > best) { best = serie; bestBis = d; } } else { serie = 0; pause++; bestPause = Math.max(bestPause, pause); }
  }
  const schritte = ds.reduce((a, d) => a + (t.get(d)?.tag?.schritte || 0), 0);
  const etagen = ds.reduce((a, d) => a + (t.get(d)?.tag?.etagen || 0), 0);
  // Sportarten
  const gr = [...JB_GRUPPEN, SONST].map((g) => {
    const xs = es.filter((e) => gruppeVon(e.a).id === g.id);
    return { ...g, n: xs.length, km: sum(xs, (a) => a.km), hm: sum(xs, (a) => a.hm), s: sum(xs, (a) => a.dauer_s), last: sum(xs, (a) => a.last) };
  }).filter((g) => g.n);
  // Monate
  const ms = monate(anfang, ende);
  const monat = ms.map((m) => {
    const xs = es.filter((e) => e.datum.startsWith(m));
    return { m, n: xs.length, s: sum(xs, (a) => a.dauer_s), km: sum(xs, (a) => a.km), hm: sum(xs, (a) => a.hm), last: sum(xs, (a) => a.last), teile: gr.map((g) => ({ v: sum(xs.filter((e) => gruppeVon(e.a).id === g.id), (a) => a.dauer_s) / 3600, farbe: g.farbe })) };
  });
  // Wochen
  const wo = new Map<string, { s: number; last: number; hm: number; km: number }>();
  for (const e of es) { const w = wochenStart(e.datum); const x = wo.get(w) || { s: 0, last: 0, hm: 0, km: 0 }; x.s += e.a.dauer_s || 0; x.last += e.a.last || 0; x.hm += e.a.hm || 0; x.km += e.a.km || 0; wo.set(w, x); }
  const wochen = [...wo.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const wocheMax = (f: (x: { s: number; last: number; hm: number; km: number }) => number) => wochen.reduce((p, w) => (f(w[1]) > f(p[1]) ? w : p), wochen[0] || ['', { s: 0, last: 0, hm: 0, km: 0 }]);
  const anzWochen = Math.max(1, kalTage / 7);
  // Tage
  const tagesLast = new Map(ds.map((d) => [d, tagLast(t.get(d))]));
  const tagMax = (f: (d: string) => number) => ds.reduce((p, d) => (f(d) > f(p) ? d : p), ds[0] || anfang);
  const tagS = (d: string) => (t.get(d)?.akt || []).reduce((a, x) => a + (x.dauer_s || 0), 0);
  const tagHm = (d: string) => (t.get(d)?.akt || []).reduce((a, x) => a + (x.hm || 0), 0);
  // Wochentag und Tageszeit
  const wt = WT.map((lab, i) => ({ lab, v: sum(es.filter((e) => wochentag(e.datum) === i), (a) => a.dauer_s) / 3600 }));
  const TZ = [{ lab: 'früh (bis 9)', lo: 0, hi: 9 }, { lab: 'vormittags', lo: 9, hi: 12 }, { lab: 'mittags', lo: 12, hi: 14 }, { lab: 'nachmittags', lo: 14, hi: 18 }, { lab: 'abends (ab 18)', lo: 18, hi: 24 }];
  const tz = TZ.map((b) => ({ lab: b.lab, v: es.filter((e) => { const m = startMin(e.a); return m != null && m / 60 >= b.lo && m / 60 < b.hi; }).length }));
  // Intensität
  const zonen = [0, 0, 0, 0, 0];
  for (const e of es) (e.a.zonen_s || []).forEach((v, i) => { if (i < 5) zonen[i] += v || 0; });
  const zSum = zonen.reduce((a, b) => a + b, 0) || 1;
  const te: Record<string, number> = {};
  for (const e of es) if (e.a.te_label && e.a.te_label !== 'UNKNOWN') { const l = TE_LABEL_DE[e.a.te_label] || e.a.te_label; te[l] = (te[l] || 0) + 1; }
  // Ausdauerwert, CTL, ATL
  const inR = (xs: Reihe[]) => xs.filter((x) => x.d >= von && x.d <= bis && Number.isFinite(x.v) && x.v > 0).sort((a, b) => a.d.localeCompare(b.d));
  const esR = inR(z.es), ctl = inR(z.ctl), atl = inR(z.atl);
  const ext = (xs: Reihe[]) => xs.length ? { a: xs[0], b: xs[xs.length - 1], max: xs.reduce((p, x) => (x.v > p.v ? x : p)), min: xs.reduce((p, x) => (x.v < p.v ? x : p)) } : null;
  // VO2max (Laufen) je Monat
  const vo2 = monatsMedian(es.filter((e) => LAUF.has(e.a.typ || '') && e.a.vo2max).map((e) => ({ d: e.datum, v: e.a.vo2max! })), ms);
  const vo2Rad = monatsMedian(es.filter((e) => RAD_ALLE.has(e.a.typ || '') && e.a.vo2max).map((e) => ({ d: e.datum, v: e.a.vo2max! })), ms);
  // Erholung
  const reihe = (f: (d: string) => number | undefined) => ds.map((d) => ({ d, v: f(d) })).filter((x): x is Reihe => x.v != null && x.v > 0);
  const rhr = reihe((d) => t.get(d)?.tag?.rhr), hrv = reihe((d) => t.get(d)?.hrv?.nacht);
  const schlafH = reihe((d) => (t.get(d)?.schlaf?.dauer_s ? t.get(d)!.schlaf!.dauer_s! / 3600 : undefined)), schlafS = reihe((d) => t.get(d)?.schlaf?.score);
  const bb = reihe((d) => t.get(d)?.bb_aufwachen), stress = reihe((d) => t.get(d)?.tag?.stress_avg);
  const ersteLetzte = (xs: Reihe[]) => ({ a: median(xs.slice(0, 30).map((x) => x.v)), b: median(xs.slice(-30).map((x) => x.v)) });
  // Analysen
  const entL = entkopplung(t, 'lauf', bis), entR = entkopplung(t, 'rad', bis);
  const effLp = nachGelaende(effPunkte(t, 'lauf'), 'flach', GRENZEN_STANDARD.lauf), effRp = nachGelaende(effPunkte(t, 'rad'), 'flach', GRENZEN_STANDARD.rad);
  const effL = trend(effLp, 'hf'), effR = trend(effRp, 'hf');
  const effMon = (ps: typeof effLp) => monatsMedian(ps.filter((p) => wertVon(p, 'hf') != null).map((p) => ({ d: p.datum, v: wertVon(p, 'hf')! })), ms);
  const lk = leistungskurve(t);
  // Highlights
  const laengste = [...es].sort((a, b) => (b.a.dauer_s || 0) - (a.a.dauer_s || 0)).slice(0, 5);
  const hoechste = es.filter((e) => (e.a.hm || 0) > 0).sort((a, b) => (b.a.hm || 0) - (a.a.hm || 0)).slice(0, 5);
  const haerteste = es.filter((e) => e.a.last).sort((a, b) => (b.a.last || 0) - (a.a.last || 0)).slice(0, 5);
  const rek = rekorde(t, bis);
  return {
    von: anfang, bis: ende, kalTage, ges, aktiv: aktiv.length, serie: { tage: best, bis: bestBis }, bestPause, schritte, etagen, gr, monat, wochen, anzWochen,
    wocheZeit: wocheMax((x) => x.s), wocheLast: wocheMax((x) => x.last), wocheHm: wocheMax((x) => x.hm),
    tagesLast, tagZeit: tagMax(tagS), tagZeitS: tagS(tagMax(tagS)), tagHm: tagMax(tagHm), tagHmV: tagHm(tagMax(tagHm)),
    wt, tz, zonen, zSum, te, esR, ctl, atl, esX: ext(esR), ctlX: ext(ctl), atlX: ext(atl), vo2, vo2Rad,
    rhr, hrv, schlafH, schlafS, bb, stress, ms,
    rhrX: ersteLetzte(rhr), hrvX: ersteLetzte(hrv), schlafHM: mittel(schlafH.map((x) => x.v)), schlafSM: mittel(schlafS.map((x) => x.v)), bbM: mittel(bb.map((x) => x.v)), stressM: mittel(stress.map((x) => x.v)),
    entL, entR, effL, effR, effLMon: effMon(effLp), effRMon: effMon(effRp), lk, laengste, hoechste, haerteste, rek,
  };
}
export type JbDaten = ReturnType<typeof jahresberichtDaten>;

/** Zusätzliche Fakten für Kiras Fazit (Ausdauerwert, Highlights, Intensität). */
export function jahresberichtFakten(d: JbDaten): string[] {
  const z12 = ((d.zonen[0] + d.zonen[1]) / d.zSum) * 100, z45 = ((d.zonen[3] + d.zonen[4]) / d.zSum) * 100;
  return [
    d.esX ? `Ausdauerwert (Garmin Endurance Score) ${de(d.esX.a.v)} → ${de(d.esX.b.v)} (${vz(d.esX.b.v - d.esX.a.v)}), Höchstwert ${de(d.esX.max.v)} am ${kurz(d.esX.max.d)}, Tiefstwert ${de(d.esX.min.v)} am ${kurz(d.esX.min.d)}.` : '',
    `Aktive Tage ${d.aktiv} von ${d.kalTage}, längste Serie ${d.serie.tage} Tage am Stück, Ø ${de(d.ges.s / 3600 / d.anzWochen, 1)} h je Woche.`,
    `Intensität (Zeit in HF-Zonen): Z1–Z2 ${de(z12)} %, Z3 ${de((d.zonen[2] / d.zSum) * 100)} %, Z4–Z5 ${de(z45)} %.`,
    d.entL.punkte.length ? `Aerobe Entkopplung Laufen: Median ${de(d.entL.median, 1)} % aus ${d.entL.punkte.length} gleichmäßigen Läufen ab 60 min.` : '',
  ].filter(Boolean);
}

// ─── HTML ─────────────────────────────────────────────────────────────────
export const JB_CSS = `
.ch{display:block}.ch text{font-size:9.5px;fill:#555}.ch .mk{font-size:9px;fill:#222;font-weight:600}.ch .ax{font-size:9px;fill:#888}.ch .big{font-size:22px;font-weight:700;fill:#111}
.hero{background:linear-gradient(135deg,#0b1f3a,#1565c0);color:#fff;border-radius:10px;padding:22px 24px 18px;margin-bottom:14px}
.hero h1{color:#fff;font-size:30px;margin:0;letter-spacing:-.01em}.hero .sub{color:#cfe0f5;margin:2px 0 14px}
.hg{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.hg div{background:rgba(255,255,255,.1);border-radius:8px;padding:8px 10px}
.hg b{display:block;font-size:22px;font-variant-numeric:tabular-nums;color:#fff}.hg span{font-size:10px;color:#cfe0f5}
.ff{margin-top:12px;font-size:11px;color:#e3edf9}.ff b{color:#fff}
.zw{display:grid;grid-template-columns:1fr 1fr;gap:14px}.box{border:1px solid #e6e6e6;border-radius:8px;padding:8px 10px;break-inside:avoid}
.box h3{font-size:11px;margin:0 0 4px;color:#333}.note{color:#666;font-size:10px;margin:3px 0 0}
.leer{color:#999;font-size:10px;padding:8px 0}.seite{break-before:page}.grid.g3{grid-template-columns:repeat(3,1fr)}
.pill{display:inline-block;border-radius:10px;padding:1px 7px;font-size:9.5px;font-weight:600}.up{background:#e3f4ea;color:#15803d}.down{background:#fde8e8;color:#b91c1c}.flat{background:#eef1f5;color:#555}
.box td:not(:nth-child(2)){white-space:nowrap}.legende{font-size:9.5px;color:#555;margin:2px 0 4px}.legende i{display:inline-block;width:9px;height:9px;border-radius:2px;margin:0 4px 0 10px;vertical-align:-1px}
@media print{.hero{-webkit-print-color-adjust:exact;print-color-adjust:exact}.box,.hg div{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
`;

const pill = (v: number | null | undefined, gutWennNegativ: boolean, txt: string) => {
  if (v == null || !Number.isFinite(v)) return '';
  const gut = gutWennNegativ ? v < 0 : v > 0;
  return `<span class="pill ${Math.abs(v) < 0.05 ? 'flat' : gut ? 'up' : 'down'}">${esc(txt)}</span>`;
};

export function jahresberichtHtml(d: JbDaten, opt: { titel: string; fazit: string | null; name?: string }) {
  const k = (lab: string, v: string, sub = '') => `<div class="k"><span>${esc(lab)}</span><b>${esc(v)}</b><span>${sub}</span></div>`;
  const h = d.ges.s / 3600;
  const everest = d.ges.hm / 8849, erde = (d.ges.km / 40075) * 100;
  const z12 = ((d.zonen[0] + d.zonen[1]) / d.zSum) * 100, z3 = (d.zonen[2] / d.zSum) * 100, z45 = ((d.zonen[3] + d.zonen[4]) / d.zSum) * 100;
  const teSort = Object.entries(d.te).sort((a, b) => b[1] - a[1]);
  const besterMonat = d.monat.reduce((p, m) => (m.s > p.s ? m : p), d.monat[0]);
  const legende = d.gr.map((g) => `<i style="background:${g.farbe}"></i>${esc(g.lab)}`).join('');

  // Seite 1 – Überblick
  let html = `
<div class="hero">
  <h1>${esc(opt.titel)}</h1>
  <div class="sub">${esc(opt.name || 'Marc Reichpietsch')} · ${kurz(d.von)} bis ${kurz(d.bis)} · ${d.kalTage} Tage</div>
  <div class="hg">
    <div><span>Einheiten</span><b>${de(d.ges.n)}</b><span>an ${d.aktiv} aktiven Tagen</span></div>
    <div><span>Trainingszeit</span><b>${de(h)} h</b><span>Ø ${de(h / d.anzWochen, 1)} h je Woche</span></div>
    <div><span>Distanz</span><b>${de(d.ges.km)} km</b><span>Ø ${de(d.ges.km / d.anzWochen)} km je Woche</span></div>
    <div><span>Höhenmeter</span><b>${de(d.ges.hm)}</b><span>Ø ${de(d.ges.hm / d.anzWochen)} hm je Woche</span></div>
    <div><span>Ausdauerwert</span><b>${d.esX ? de(d.esX.b.v) : '–'}</b><span>${d.esX ? `Start ${de(d.esX.a.v)} · max ${de(d.esX.max.v)}` : 'keine Daten'}</span></div>
    <div><span>Fitness (CTL)</span><b>${d.ctlX ? `${de(d.ctlX.a.v)} → ${de(d.ctlX.b.v)}` : '–'}</b><span>${d.ctlX ? `Höchstwert ${de(d.ctlX.max.v)} am ${kurz(d.ctlX.max.d)}` : ''}</span></div>
    <div><span>Längste Serie</span><b>${d.serie.tage} Tage</b><span>${d.serie.bis ? `bis ${kurz(d.serie.bis)}` : ''} am Stück aktiv</span></div>
    <div><span>Aktive Kalorien</span><b>${de(d.ges.kcal)}</b><span>kcal in Einheiten</span></div>
  </div>
  <div class="ff">Das sind <b>${de(everest, 1)}× Mount Everest</b> an Höhenmetern und <b>${de(erde, 1)} %</b> des Erdumfangs. Dazu ${de(d.schritte)} Schritte und ${de(d.etagen)} Etagen im Alltag.</div>
</div>
${opt.fazit ? `<h2>Kiras Fazit</h2><div class="faz">${esc(opt.fazit)}</div>` : ''}
<h2>Ausdauerwert und Fitness</h2>
<div class="box"><h3>Garmin-Ausdauerwert ${d.esX ? pill(d.esX.b.v - d.esX.a.v, false, `${vz(d.esX.b.v - d.esX.a.v)} im Zeitraum`) : ''}</h3>
${svgZeit([{ name: 'Ausdauerwert', farbe: '#1565c0', xs: d.esR, flaeche: true, breite: 2.2 }], { hoehe: 170, von: d.von, bis: d.bis, marken: d.esX ? [{ d: d.esX.max.d, v: d.esX.max.v, text: `max ${de(d.esX.max.v)}`, farbe: '#15803d' }, { d: d.esX.min.d, v: d.esX.min.v, text: `min ${de(d.esX.min.v)}`, farbe: '#b91c1c' }, { d: d.esX.b.d, v: d.esX.b.v, text: `heute ${de(d.esX.b.v)}`, farbe: '#1565c0' }] : [] })}
<p class="note">Der Ausdauerwert fasst laut Garmin VO2max, Dauer und Intensität der letzten Wochen zusammen. Er reagiert träge, deshalb zählt die Richtung über Monate.</p></div>
<div class="box" style="margin-top:10px"><h3>Fitness (CTL) und Ermüdung (ATL)</h3>
${svgZeit([{ name: 'Fitness CTL', farbe: '#22a15a', xs: d.ctl, breite: 2.2 }, { name: 'Ermüdung ATL', farbe: '#ef7d00', xs: d.atl, breite: 1.3, strich: '3 2' }], { hoehe: 150, von: d.von, bis: d.bis, marken: d.ctlX ? [{ d: d.ctlX.max.d, v: d.ctlX.max.v, text: `CTL max ${de(d.ctlX.max.v)}`, farbe: '#15803d' }] : [] })}
</div>`;

  // Seite 2 – Umfang
  html += `
<div class="seite"></div>
<h2>Umfang über das Jahr</h2>
<div class="box"><h3>Trainingsstunden je Monat nach Sportart</h3><div class="legende">${legende}</div>
${svgStapel(d.monat.map((m) => ({ lab: monatName(m.m), teile: m.teile })), { einheit: 'Stunden', summeText: (v) => de(v, 0) })}
<p class="note">Stärkster Monat: ${besterMonat ? `${monatName(besterMonat.m)} mit ${dauerTxt(besterMonat.s)}, ${de(besterMonat.km)} km und ${de(besterMonat.hm)} hm` : '–'}.</p></div>
<div class="box" style="margin-top:10px"><h3>Jeder Tag im Überblick (Trainingslast)</h3>
${svgKalender(d.tagesLast, d.von, d.bis)}
<p class="note">${d.aktiv} aktive Tage, längste Pause ${d.bestPause} Tage. Dunkler = höhere Garmin-Trainingslast an diesem Tag.</p></div>
<div class="zw" style="margin-top:10px">
  <div class="box"><h3>Anteil der Trainingszeit</h3>${svgDonut(d.gr.map((g) => ({ lab: g.lab, v: g.s, farbe: g.farbe })), `${de(h)} h`, 'gesamt')}</div>
  <div class="box"><h3>Wann du trainierst</h3>
    ${svgHBalken(d.wt.map((w) => ({ lab: w.lab, v: w.v, text: `${de(w.v)} h` })))}
    ${svgHBalken(d.tz.map((w) => ({ lab: w.lab, v: w.v, farbe: '#7c3aed', text: `${w.v}×` })))}
  </div>
</div>
<h2>Nach Sportart</h2>
<table><tr><th>Sport</th><th>Einheiten</th><th>Zeit</th><th>km</th><th>hm</th><th>Last</th><th>Ø Dauer</th></tr>${d.gr.map((g) => `<tr><td><span style="color:${g.farbe}">■</span> ${esc(g.lab)}</td><td>${g.n}</td><td>${dauerTxt(g.s)}</td><td>${de(g.km)}</td><td>${de(g.hm)}</td><td>${de(g.last)}</td><td>${dauerTxt(g.s / g.n)}</td></tr>`).join('')}
<tr><td><b>Gesamt</b></td><td><b>${d.ges.n}</b></td><td><b>${dauerTxt(d.ges.s)}</b></td><td><b>${de(d.ges.km)}</b></td><td><b>${de(d.ges.hm)}</b></td><td><b>${de(d.ges.last)}</b></td><td><b>${dauerTxt(d.ges.s / Math.max(1, d.ges.n))}</b></td></tr></table>`;

  // Seite 3 – Intensität und Erholung
  html += `
<div class="seite"></div>
<h2>Intensität</h2>
<div class="zw">
  <div class="box"><h3>Zeit in Herzfrequenzzonen</h3>
    ${svgHBalken(d.zonen.map((v, i) => ({ lab: `Zone ${i + 1}`, v: v / 3600, farbe: ZONEN_FARBE[i], text: `${de(v / 3600)} h · ${de((v / d.zSum) * 100)} %` })))}
    <p class="note">Locker (Z1–Z2) ${de(z12)} %, Mitte (Z3) ${de(z3)} %, hart (Z4–Z5) ${de(z45)} %.</p></div>
  <div class="box"><h3>Trainingswirkung laut Garmin</h3>
    ${teSort.length ? svgHBalken(teSort.slice(0, 7).map(([l, n]) => ({ lab: l, v: n, farbe: '#0ea5a4', text: `${n}×` }))) : '<div class="leer">keine Daten</div>'}
    <p class="note">Häufigste Wirkung: ${teSort[0] ? `${esc(teSort[0][0])} (${teSort[0][1]}×)` : '–'}.</p></div>
</div>
<h2>Erholung und Körper</h2>
<div class="grid">${k('Ruhepuls', `${de(d.rhrX.a)} → ${de(d.rhrX.b)}`, 'erster vs. letzter Monat ' + pill(d.rhrX.b != null && d.rhrX.a != null ? d.rhrX.b - d.rhrX.a : null, true, `${vz(d.rhrX.b != null && d.rhrX.a != null ? d.rhrX.b - d.rhrX.a : null)} bpm`))}
${k('HRV nachts', `${de(d.hrvX.a)} → ${de(d.hrvX.b)} ms`, pill(d.hrvX.b != null && d.hrvX.a != null ? d.hrvX.b - d.hrvX.a : null, false, `${vz(d.hrvX.b != null && d.hrvX.a != null ? d.hrvX.b - d.hrvX.a : null)} ms`))}
${k('Schlaf Ø', `${de(d.schlafHM, 1)} h`, `Score Ø ${de(d.schlafSM)}`)}
${k('Body Battery morgens', `Ø ${de(d.bbM)}`, `Stress Ø ${de(d.stressM)}`)}</div>
<div class="zw" style="margin-top:10px">
  <div class="box"><h3>Ruhepuls und HRV (Monatsmedian)</h3>
    ${svgZeit([{ name: 'Ruhepuls', farbe: '#d64545', xs: monatsMedian(d.rhr, d.ms) }], { hoehe: 120, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'bpm' })}
    ${svgZeit([{ name: 'HRV', farbe: '#7c3aed', xs: monatsMedian(d.hrv, d.ms) }], { hoehe: 120, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'ms' })}
  </div>
  <div class="box"><h3>Schlaf (Monatsmittel)</h3>
    ${svgZeit([{ name: 'Schlafdauer', farbe: '#1976d2', xs: monatsMedian(d.schlafH, d.ms) }], { hoehe: 120, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'Stunden', dez: 1 })}
    ${svgZeit([{ name: 'Schlafscore', farbe: '#0ea5a4', xs: monatsMedian(d.schlafS, d.ms) }], { hoehe: 120, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'Score' })}
  </div>
</div>`;

  // Seite 4 – Analysen
  const entPkt = [
    { name: 'Laufen', farbe: '#ef7d00', xs: d.entL.punkte.map((p) => ({ d: p.datum, v: p.y })) },
    { name: 'Rad', farbe: '#1976d2', xs: d.entR.punkte.map((p) => ({ d: p.datum, v: p.y })) },
  ].filter((s) => s.xs.length);
  const ftp = d.lk.reihe.filter((r) => r.ftp).map((r) => ({ d: r.monat + '-15', v: r.ftp! }));
  const best20 = d.lk.reihe.filter((r) => r.best20).map((r) => ({ d: r.monat + '-15', v: r.best20! }));
  html += `
<div class="seite"></div>
<h2>Leistung im Detail</h2>
<div class="grid">${k('Effizienz Laufen flach', `${vz(d.effL.prozent, 1)} %`, 'Herzschläge je km, − = besser ' + pill(d.effL.prozent, true, d.effL.prozent == null ? '' : d.effL.prozent < 0 ? 'effizienter' : 'weniger effizient'))}
${k('Effizienz Rad flach', `${vz(d.effR.prozent, 1)} %`, 'Herzschläge je km, − = besser ' + pill(d.effR.prozent, true, d.effR.prozent == null ? '' : d.effR.prozent < 0 ? 'effizienter' : 'weniger effizient'))}
${k('Entkopplung Laufen', d.entL.punkte.length ? `${de(d.entL.median, 1)} %` : '–', `${d.entL.punkte.length} lange Läufe, unter 5 % = stabil`)}
${k('Entkopplung Rad', d.entR.punkte.length ? `${de(d.entR.median, 1)} %` : '–', `${d.entR.punkte.length} lange Fahrten`)}</div>
<div class="box" style="margin-top:10px"><h3>Aerobe Entkopplung langer, gleichmäßiger Einheiten</h3>
${svgPunkte(entPkt, { von: d.von, bis: d.bis, einheit: 'Entkopplung %', linie: 5 })}
<p class="note">Wie stark der Puls in der zweiten Hälfte bei gleichem Tempo bzw. gleicher Leistung steigt. Nur flache, gleichmäßige Einheiten ab 60 min.${d.entL.luecken.length ? ` Laufen: ${d.entL.luecken.map((l) => `${kurz(l.von)} bis ${l.offen ? 'heute' : kurz(l.bis)} keine auswertbaren Läufe`).join('; ')}.` : ''}</p></div>
<div class="zw" style="margin-top:10px">
  <div class="box"><h3>Herzschläge je km auf flachen Strecken</h3>
    ${svgZeit([{ name: 'Laufen', farbe: '#ef7d00', xs: d.effLMon }], { hoehe: 125, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'Laufen · Schläge/km' })}
    ${svgZeit([{ name: 'Rad', farbe: '#1976d2', xs: d.effRMon }], { hoehe: 125, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'Rad · Schläge/km' })}
    <p class="note">Monatsmedian. Weniger Schläge für denselben Kilometer heißt effizienter.</p></div>
  <div class="box"><h3>${d.vo2.length >= 2 || d.vo2Rad.length >= 2 ? 'Rad-Leistung und VO2max' : 'Rad-Leistung'}</h3>
    ${ftp.length >= 2 ? svgZeit([{ name: 'FTP-Schätzung', farbe: '#1565c0', xs: ftp }], { hoehe: 125, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'FTP-Schätzung (Watt, 95 % der besten 20 min)' }) : '<div class="leer">zu wenige Wattdaten</div>'}
    ${d.vo2.length >= 2 || d.vo2Rad.length >= 2 ? svgZeit([...(d.vo2.length >= 2 ? [{ name: 'VO2max Laufen', farbe: '#ef7d00', xs: d.vo2 }] : []), ...(d.vo2Rad.length >= 2 ? [{ name: 'VO2max Rad', farbe: '#1976d2', xs: d.vo2Rad }] : [])], { hoehe: 125, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'VO2max (Garmin)' }) : best20.length >= 2 ? svgZeit([{ name: 'Beste 20 min', farbe: '#7c3aed', xs: best20 }], { hoehe: 125, breite: 340, von: d.ms[0] + '-15', bis: d.ms[d.ms.length - 1] + '-15', einheit: 'Beste 20 Minuten je Monat (Watt)' }) : ''}
    ${d.lk.reihe.some((r) => r.ftp) ? `<p class="note">FTP-Schätzung zuletzt ${de(d.lk.reihe.filter((r) => r.ftp).slice(-1)[0].ftp)} W, Höchstwert ${de(Math.max(...d.lk.reihe.map((r) => r.ftp || 0)))} W.</p>` : ''}
  </div>
</div>`;

  // Seite 5 – Highlights und Rekorde
  const zeile = (e: E, extra: string) => `<tr><td>${kurz(e.datum)}</td><td>${esc(e.a.name || '')}</td><td>${de(e.a.km, 1)}</td><td>${de(e.a.hm)}</td><td>${dauerTxt(e.a.dauer_s)}</td><td>${extra}</td></tr>`;
  html += `
<div class="seite"></div>
<h2>Highlights</h2>
<div class="grid">
${k('Härteste Woche', d.wocheLast[0] ? `${de(d.wocheLast[1].last)} Last` : '–', d.wocheLast[0] ? `ab ${kurz(d.wocheLast[0])} · ${dauerTxt(d.wocheLast[1].s)}` : '')}
${k('Längste Woche', d.wocheZeit[0] ? dauerTxt(d.wocheZeit[1].s) : '–', d.wocheZeit[0] ? `ab ${kurz(d.wocheZeit[0])} · ${de(d.wocheZeit[1].km)} km` : '')}
${k('Meiste Höhenmeter an einem Tag', `${de(d.tagHmV)} hm`, d.tagHm ? kurz(d.tagHm) : '')}
${k('Längster Trainingstag', dauerTxt(d.tagZeitS), d.tagZeit ? kurz(d.tagZeit) : '')}
</div>
<div class="zw" style="margin-top:10px">
  <div class="box"><h3>Längste Einheiten</h3><table><tr><th>Datum</th><th>Einheit</th><th>km</th><th>hm</th><th>Dauer</th><th>Last</th></tr>${d.laengste.map((e) => zeile(e, de(e.a.last))).join('')}</table></div>
  <div class="box"><h3>Härteste Einheiten (Trainingslast)</h3><table><tr><th>Datum</th><th>Einheit</th><th>km</th><th>hm</th><th>Dauer</th><th>Last</th></tr>${d.haerteste.map((e) => zeile(e, `<b>${de(e.a.last)}</b>`)).join('')}</table></div>
</div>
<h2>Größte Touren nach Höhenmetern</h2>
<table><tr><th>Datum</th><th>Tour</th><th>km</th><th>hm</th><th>Dauer</th><th>Ø HF</th></tr>${d.hoechste.map((e) => zeile(e, de(e.a.hf_avg))).join('')}</table>
<h2>Rekorde</h2>
<table><tr><th>Bereich</th><th>Rekord</th><th>Wert</th><th>Datum</th><th>Einheit</th></tr>${d.rek.filter((r) => r.gruppe !== 'koerper').map((r) => `<tr><td>${esc(GLAB[r.gruppe] || r.gruppe)}</td><td>${esc(r.titel)}</td><td>${esc(r.text)}${r.neu ? ' <span class="neu">neu</span>' : ''}</td><td>${r.bis ? kurz(r.datum) + '–' + kurz(r.bis) : kurz(r.datum)}</td><td>${esc(r.detail || '')}</td></tr>`).join('')}</table>`;
  return html;
}
