/**
 * Rekorde aus den Garmin-Tagesdetails (garmin_tag): Einzel-Einheiten, Wochen/Monate, Serien und Körperwerte.
 * Rein deterministisch – mit Mindestumfängen, damit kurze Schnipsel (Ampelsprint, GPS-Sprung) keine Rekorde setzen.
 */
import type { GarminAkt, GarminTage } from './garmin';

export type RekordGruppe = 'lauf' | 'rad' | 'berg' | 'winter' | 'indoor' | 'gesamt' | 'koerper';

export const GRUPPEN: { id: RekordGruppe; label: string; typen?: string[] }[] = [
  { id: 'lauf', label: 'Laufen', typen: ['running'] },
  { id: 'rad', label: 'Rad draußen', typen: ['road_biking', 'cycling', 'gravel_cycling', 'mountain_biking'] },
  { id: 'berg', label: 'Wandern, Trail & Berg', typen: ['hiking', 'mountaineering', 'walking', 'trail_running'] },
  { id: 'winter', label: 'Ski & Skike', typen: ['backcountry_skiing', 'skate_skiing_ws', 'cross_country_skiing_ws', 'resort_skiing'] },
  { id: 'indoor', label: 'Indoor', typen: ['indoor_cycling', 'virtual_ride', 'indoor_rowing', 'treadmill_running'] },
  { id: 'gesamt', label: 'Wochen & Serien' },
  { id: 'koerper', label: 'Körper & Erholung' },
];

export interface Rekord {
  gruppe: RekordGruppe;
  titel: string;
  wert: number;
  text: string;          // formatierter Wert
  datum: string;         // Tag bzw. Wochenbeginn
  bis?: string;          // Ende bei Wochen/Serien
  detail?: string;       // Einordnung (km, hm, HF …)
  aktId?: number;        // Garmin-Aktivität
  neu: boolean;          // in den letzten 30 Tagen aufgestellt
  zweiter?: string;      // bisher zweitbester Wert (Abstand)
}

const de = (v: number, d = 0) => v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
export const tempoTxt = (minProKm: number) => {
  const s = Math.round(minProKm * 60);
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
};
const dauerTxt = (s: number) => {
  const m = Math.round(s / 60);
  return m >= 60 ? `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')} h` : `${m} min`;
};
const kurz = (d: string) => new Date(d + 'T12:00:00').toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: '2-digit' });

type Eintrag = { datum: string; a: GarminAkt };
const RAD_INDOOR = new Set(['indoor_cycling', 'virtual_ride']);

/** Tempo aus Bewegungszeit (Pausen zählen nicht). */
const vBewegt = (a: GarminAkt) => (a.km && a.bewegt_s ? a.km / (a.bewegt_s / 3600) : a.v_kmh || 0);

interface Def {
  titel: string;
  wert: (a: GarminAkt) => number | null;
  fmt: (w: number) => string;
  detail?: (a: GarminAkt) => string;
  /** true = kleiner ist besser */
  min?: boolean;
}

function besteAus(es: Eintrag[], def: Def, gruppe: RekordGruppe, seit: string): Rekord | null {
  const xs = es
    .map((e) => ({ e, w: def.wert(e.a) }))
    .filter((x): x is { e: Eintrag; w: number } => x.w != null && Number.isFinite(x.w) && x.w > 0);
  if (!xs.length) return null;
  xs.sort((p, q) => (def.min ? p.w - q.w : q.w - p.w));
  const b = xs[0], z = xs[1];
  return {
    gruppe, titel: def.titel, wert: b.w, text: def.fmt(b.w), datum: b.e.datum,
    detail: [def.detail?.(b.e.a), b.e.a.name].filter(Boolean).join(' · '),
    aktId: b.e.a.id, neu: b.e.datum >= seit,
    zweiter: z ? `zweitbester: ${def.fmt(z.w)} (${kurz(z.e.datum)})` : undefined,
  };
}

const kmHm = (a: GarminAkt) => `${de(a.km || 0, 1)} km · ${de(a.hm || 0)} hm · ${dauerTxt(a.bewegt_s || a.dauer_s || 0)}`;

function sportDefs(g: RekordGruppe): Def[] {
  const dist: Def = { titel: 'Längste Distanz', wert: (a) => a.km || null, fmt: (w) => `${de(w, 1)} km`, detail: (a) => `${de(a.hm || 0)} hm · ${dauerTxt(a.bewegt_s || a.dauer_s || 0)}` };
  const hm: Def = { titel: 'Meiste Höhenmeter', wert: (a) => a.hm || null, fmt: (w) => `${de(w)} hm`, detail: (a) => `${de(a.km || 0, 1)} km · ${dauerTxt(a.bewegt_s || a.dauer_s || 0)}` };
  const dauer: Def = { titel: 'Längste Einheit', wert: (a) => a.dauer_s || null, fmt: dauerTxt, detail: kmHm };
  const last: Def = { titel: 'Höchste Garmin-Trainingsbelastung', wert: (a) => a.last || null, fmt: (w) => `${de(w)}`, detail: kmHm };
  const kcal: Def = { titel: 'Meiste Kalorien', wert: (a) => a.kcal || null, fmt: (w) => `${de(w)} kcal`, detail: kmHm };
  const hoehe: Def = { titel: 'Höchster Punkt', wert: (a) => a.hoehe_max || null, fmt: (w) => `${de(w)} m`, detail: kmHm };
  if (g === 'lauf') {
    const pace = (minKm: number, titel: string): Def => ({
      titel, min: true,
      wert: (a) => ((a.km || 0) >= minKm && vBewegt(a) > 5 && vBewegt(a) < 22 ? 60 / vBewegt(a) : null),
      fmt: (w) => `${tempoTxt(w)} /km`,
      detail: (a) => `${de(a.km || 0, 1)} km in ${dauerTxt(a.bewegt_s || a.dauer_s || 0)} · Ø HF ${a.hf_avg ?? '–'}`,
    });
    return [
      pace(5, 'Schnellster Lauf ab 5 km'), pace(10, 'Schnellster Lauf ab 10 km'), pace(21, 'Schnellster Lauf ab 21 km'),
      dist, hm, dauer,
      { titel: 'Höchste Ø Laufleistung (ab 20 min)', wert: (a) => ((a.bewegt_s || 0) >= 1200 ? a.watt || null : null), fmt: (w) => `${de(w)} W`, detail: kmHm },
      last, kcal,
    ];
  }
  if (g === 'rad') {
    return [
      { titel: 'Schnellster Schnitt ab 30 km', wert: (a) => ((a.km || 0) >= 30 && vBewegt(a) < 50 ? vBewegt(a) : null), fmt: (w) => `${de(w, 1)} km/h`, detail: (a) => `${de(a.km || 0, 1)} km · ${de(a.hm || 0)} hm · Ø HF ${a.hf_avg ?? '–'}` },
      { titel: 'Höchste Normalized Power (ab 45 min)', wert: (a) => ((a.bewegt_s || 0) >= 2700 ? a.watt_np || null : null), fmt: (w) => `${de(w)} W`, detail: kmHm },
      { titel: 'Beste 20 Minuten', wert: (a) => a.watt_20 || null, fmt: (w) => `${de(w)} W`, detail: kmHm },
      dist, hm,
      { titel: 'Steilste Tour (ab 20 km)', wert: (a) => ((a.km || 0) >= 20 ? (a.hm || 0) / (a.km || 1) : null), fmt: (w) => `${de(w, 1)} hm/km`, detail: kmHm },
      dauer, last, kcal,
    ];
  }
  if (g === 'berg') {
    return [hm, dist, dauer, hoehe,
      { titel: 'Schnellster Aufstieg (ab 500 hm)', wert: (a) => ((a.hm || 0) >= 500 && a.bewegt_s ? (a.hm || 0) / (a.bewegt_s / 3600) : null), fmt: (w) => `${de(w)} hm/h`, detail: kmHm },
      last, kcal];
  }
  if (g === 'winter') return [dist, hm, dauer, hoehe, last, kcal];
  if (g === 'indoor') {
    return [dauer,
      { titel: 'Rolle: höchste Ø Leistung (ab 20 min)', wert: (a) => (RAD_INDOOR.has(a.typ || '') && (a.bewegt_s || a.dauer_s || 0) >= 1200 ? a.watt || null : null), fmt: (w) => `${de(w)} W`, detail: (a) => dauerTxt(a.bewegt_s || a.dauer_s || 0) },
      { titel: 'Rolle: beste 20 Minuten', wert: (a) => (RAD_INDOOR.has(a.typ || '') ? a.watt_20 || null : null), fmt: (w) => `${de(w)} W`, detail: (a) => dauerTxt(a.bewegt_s || a.dauer_s || 0) },
      { titel: 'Rudern: höchste Ø Leistung (ab 10 min)', wert: (a) => (a.typ === 'indoor_rowing' && (a.bewegt_s || a.dauer_s || 0) >= 600 ? a.watt || null : null), fmt: (w) => `${de(w)} W`, detail: (a) => `${de(a.km || 0, 1)} km · ${dauerTxt(a.bewegt_s || a.dauer_s || 0)}` },
      { titel: 'Rudern: längste Distanz', wert: (a) => (a.typ === 'indoor_rowing' ? a.km || null : null), fmt: (w) => `${de(w, 1)} km`, detail: (a) => dauerTxt(a.bewegt_s || a.dauer_s || 0) },
      last, kcal];
  }
  return [];
}

/** Montag der Woche (ISO). */
function wochenStart(d: string): string {
  const x = new Date(d + 'T12:00:00');
  const wd = (x.getDay() + 6) % 7;
  x.setDate(x.getDate() - wd);
  return x.toLocaleDateString('sv-SE');
}
const plusTage = (d: string, n: number) => { const x = new Date(d + 'T12:00:00'); x.setDate(x.getDate() + n); return x.toLocaleDateString('sv-SE'); };

function perioden(tage: GarminTage, seit: string, heute: string): Rekord[] {
  const woche = new Map<string, { km: number; hm: number; s: number; last: number; n: number }>();
  const monat = new Map<string, { km: number; hm: number; s: number; last: number }>();
  const aktiv: string[] = [];
  for (const [d, g] of tage) {
    const akts = (g.akt || []).filter((a) => (a.dauer_s || 0) >= 900);
    if (!akts.length) continue;
    aktiv.push(d);
    const w = woche.get(wochenStart(d)) || { km: 0, hm: 0, s: 0, last: 0, n: 0 };
    const m = monat.get(d.slice(0, 7)) || { km: 0, hm: 0, s: 0, last: 0 };
    for (const a of akts) {
      w.km += a.km || 0; w.hm += a.hm || 0; w.s += a.dauer_s || 0; w.last += a.last || 0; w.n++;
      m.km += a.km || 0; m.hm += a.hm || 0; m.s += a.dauer_s || 0; m.last += a.last || 0;
    }
    woche.set(wochenStart(d), w); monat.set(d.slice(0, 7), m);
  }
  const out: Rekord[] = [];
  const wochen = [...woche.entries()];
  const top = (titel: string, f: (x: { km: number; hm: number; s: number; last: number }) => number, fmt: (v: number) => string, detail: (x: any) => string) => {
    const xs = wochen.map(([d, x]) => ({ d, x, v: f(x) })).sort((a, b) => b.v - a.v);
    if (!xs.length || xs[0].v <= 0) return;
    out.push({ gruppe: 'gesamt', titel, wert: xs[0].v, text: fmt(xs[0].v), datum: xs[0].d, bis: plusTage(xs[0].d, 6), detail: detail(xs[0].x), neu: plusTage(xs[0].d, 6) >= seit, zweiter: xs[1] ? `zweitbeste Woche: ${fmt(xs[1].v)} (KW ab ${kurz(xs[1].d)})` : undefined });
  };
  const wd = (x: any) => `${x.n} Einheiten · ${de(x.km)} km · ${de(x.hm)} hm · ${dauerTxt(x.s)}`;
  top('Meiste Kilometer in einer Woche', (x) => x.km, (v) => `${de(v)} km`, wd);
  top('Meiste Höhenmeter in einer Woche', (x) => x.hm, (v) => `${de(v)} hm`, wd);
  top('Meiste Trainingszeit in einer Woche', (x) => x.s, dauerTxt, wd);
  top('Höchste Wochenbelastung', (x) => x.last, (v) => de(v), wd);
  const ms = [...monat.entries()].sort((a, b) => b[1].hm - a[1].hm);
  if (ms.length) {
    const [mk, mx] = ms[0];
    const name = new Date(mk + '-15T12:00:00').toLocaleDateString('de-DE', { month: 'long', year: 'numeric' });
    out.push({ gruppe: 'gesamt', titel: 'Höhenmeter-Monat', wert: mx.hm, text: `${de(mx.hm)} hm`, datum: mk + '-01', detail: `${name} · ${de(mx.km)} km · ${dauerTxt(mx.s)}`, neu: mk === heute.slice(0, 7) || mk === seit.slice(0, 7), zweiter: ms[1] ? `zweitbester Monat: ${de(ms[1][1].hm)} hm` : undefined });
  }
  // längste Serie aktiver Tage (≥ 15 min)
  aktiv.sort();
  let best = { n: 0, von: '', bis: '' }, cur = { n: 0, von: '', bis: '' };
  for (const d of aktiv) {
    if (cur.n && plusTage(cur.bis, 1) === d) { cur.n++; cur.bis = d; } else cur = { n: 1, von: d, bis: d };
    if (cur.n > best.n) best = { ...cur };
  }
  if (best.n > 1) out.push({ gruppe: 'gesamt', titel: 'Längste Serie aktiver Tage', wert: best.n, text: `${best.n} Tage`, datum: best.von, bis: best.bis, detail: 'jeden Tag mindestens 15 min Sport', neu: best.bis >= seit });
  return out;
}

function koerper(tage: GarminTage, seit: string): Rekord[] {
  const out: Rekord[] = [];
  const xs = [...tage.entries()];
  const top = (titel: string, f: (g: any) => number | null | undefined, fmt: (v: number) => string, min = false, detail?: (g: any) => string) => {
    const ys = xs.map(([d, g]) => ({ d, g, v: f(g) })).filter((y): y is { d: string; g: any; v: number } => typeof y.v === 'number' && y.v > 0).sort((a, b) => (min ? a.v - b.v : b.v - a.v));
    if (!ys.length) return;
    out.push({ gruppe: 'koerper', titel, wert: ys[0].v, text: fmt(ys[0].v), datum: ys[0].d, detail: detail?.(ys[0].g), neu: ys[0].d >= seit, zweiter: ys[1] ? `danach: ${fmt(ys[1].v)} (${kurz(ys[1].d)})` : undefined });
  };
  top('Niedrigster Ruhepuls', (g) => g.tag?.rhr, (v) => `${v} bpm`, true);
  top('Höchste nächtliche HRV', (g) => g.hrv?.nacht, (v) => `${v} ms`, false, (g) => g.hrv?.status ? `Status ${String(g.hrv.status).toLowerCase()}` : '');
  top('Bester Schlafscore', (g) => g.schlaf?.score, (v) => `${v}`, false, (g) => g.schlaf?.dauer_s ? `${dauerTxt(g.schlaf.dauer_s)} Schlaf` : '');
  top('Längster Tiefschlaf', (g) => g.schlaf?.tief_s, dauerTxt, false);
  top('Höchste Body Battery am Morgen', (g) => g.bb_aufwachen, (v) => `${v}`, false);
  top('Meiste Schritte an einem Tag', (g) => g.tag?.schritte, (v) => de(v), false);
  return out;
}

const VMAX: Partial<Record<RekordGruppe, number>> = { lauf: 22, rad: 60, berg: 12, winter: 40 };
function plausibel(a: GarminAkt, g: RekordGruppe): boolean {
  const vmax = VMAX[g];
  if (!vmax || !a.km) return true;
  const s = a.bewegt_s || a.dauer_s || 0;
  return s > 0 && a.km / (s / 3600) <= vmax;
}

export function rekorde(tage: GarminTage | null, heute = new Date().toLocaleDateString('sv-SE')): Rekord[] {
  if (!tage || !tage.size) return [];
  const seit = plusTage(heute, -30);
  const es: Eintrag[] = [];
  for (const [d, g] of tage) for (const a of g.akt || []) es.push({ datum: d, a });
  const out: Rekord[] = [];
  for (const g of GRUPPEN) {
    if (!g.typen) continue;
    // Plausibilität: GPS-Sprünge (z. B. Störungen) dürfen keine Rekorde setzen
    const sel = es.filter((e) => g.typen!.includes(e.a.typ || '') && plausibel(e.a, g.id));
    if (!sel.length) continue;
    for (const def of sportDefs(g.id)) {
      const r = besteAus(sel, def, g.id, seit);
      if (r) out.push(r);
    }
  }
  out.push(...perioden(tage, seit, heute), ...koerper(tage, seit));
  return out;
}

export function zeitraum(tage: GarminTage | null): { von: string; bis: string } | null {
  if (!tage || !tage.size) return null;
  const ds = [...tage.keys()].sort();
  return { von: ds[0], bis: ds[ds.length - 1] };
}
