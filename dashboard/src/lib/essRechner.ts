/**
 * ESS-Rechner: lernt je Sportprofil aus Marcs Garmin-Aktivitäten (garmin_tag, Garmin Personal API),
 * wie die Last (ESS = Garmin Trainingsbelastung) von Dauer, HF, Tempo, Leistung und Höhenmetern abhängt.
 * Rechnet in beide Richtungen: Einheit → ESS und Ziel-ESS (+ Dauer oder Distanz) → Tempo, HF, Leistung.
 *
 * Modelle (gewichtete kleinste Quadrate, neuere Einheiten zählen mehr, Halbwertszeit 365 Tage):
 *   Basis:  ln ESS = a + b·ln(Minuten) + c·HF                (alle Profile)
 *   Rad:    ln ESS = a + b·ln(Minuten) + c·HF + d·NP          (wenn NP bekannt)
 *   Lauf:   ln ESS = a + b·ln(Minuten) + c·HF + d·v           (wenn Tempo bekannt)
 *   Tempo:  v = a + b·HF + c·hm/km   (Lauf, Rad draußen, Rudern)
 *   NP:     NP = a + b·HF             (Rad)
 *   Wandern: Gehzeit h = km/v_flach + hm/hm_pro_h
 */
import type { GarminAkt, GarminTage } from './garmin';

export type ProfilId = 'run' | 'bike_road' | 'bike_gravel' | 'bike_indoor' | 'hike' | 'row';

export const PROFILE: { id: ProfilId; label: string; sport: string; typen: string[]; tempo: 'pace' | 'kmh' | 'row' | 'none' }[] = [
  { id: 'run', label: 'Laufen', sport: 'Run', typen: ['running', 'treadmill_running', 'trail_running'], tempo: 'pace' },
  { id: 'bike_road', label: 'Rennrad', sport: 'Bike', typen: ['road_biking', 'cycling'], tempo: 'kmh' },
  { id: 'bike_gravel', label: 'Gravel/MTB', sport: 'Bike', typen: ['gravel_cycling', 'mountain_biking'], tempo: 'kmh' },
  { id: 'bike_indoor', label: 'Rad indoor', sport: 'Bike', typen: ['indoor_cycling', 'virtual_ride'], tempo: 'none' },
  { id: 'hike', label: 'Wandern/Bergtour', sport: 'Hike', typen: ['hiking', 'walking', 'mountaineering'], tempo: 'kmh' },
  { id: 'row', label: 'Rudern indoor', sport: 'Row', typen: ['indoor_rowing', 'rowing_v2'], tempo: 'row' },
];

type Probe = { datum: string; name?: string; min: number; km: number; hm: number; hf: number; np?: number; v: number; ess: number; w: number };

export type Modell = {
  id: ProfilId;
  n: number;
  base: number[]; // a, b, c
  ext?: number[]; // a, b, c, d
  extKey?: 'np' | 'v';
  vLink?: number[]; // a, b(hf), c(hm/km)
  npLink?: number[]; // a, b(hf)
  hike?: { vFlach: number; hmProH: number };
  /** ohne HF: ln ESS = a + b·ln min + c·v + d·hm/km  bzw. Rad: a + b·ln min + c·NP */
  ohneHfV?: number[];
  ohneHfNp?: number[];
  fehlerOhneHf?: number;
  /** Wandern: HF = a + b·hm/h + c·v */
  hfLink?: number[];
  fehler: number; // mittlerer relativer Fehler Basis
  fehlerExt?: number;
  hf: { p5: number; p25: number; p50: number; p75: number; p95: number };
  min: { p5: number; p50: number; p95: number };
  hpk: { p50: number; p90: number };
  proben: Probe[];
};

const LN2 = Math.log(2);

function quantil(xs: number[], q: number) {
  if (!xs.length) return NaN;
  const s = [...xs].sort((a, b) => a - b);
  const i = (s.length - 1) * q;
  const lo = Math.floor(i);
  return s[lo] + (s[Math.min(s.length - 1, lo + 1)] - s[lo]) * (i - lo);
}

/** gewichtete kleinste Quadrate, X ohne Einsen-Spalte (wird ergänzt) */
function wls(X: number[][], y: number[], w: number[], ridge = 1e-6): number[] | null {
  const k = X[0].length + 1;
  const A = Array.from({ length: k }, () => new Array(k).fill(0));
  const b = new Array(k).fill(0);
  for (let i = 0; i < y.length; i++) {
    const r = [1, ...X[i]];
    for (let p = 0; p < k; p++) {
      b[p] += w[i] * r[p] * y[i];
      for (let q = 0; q < k; q++) A[p][q] += w[i] * r[p] * r[q];
    }
  }
  for (let p = 1; p < k; p++) A[p][p] += ridge;
  // Gauß
  for (let c = 0; c < k; c++) {
    let piv = c;
    for (let r = c + 1; r < k; r++) if (Math.abs(A[r][c]) > Math.abs(A[piv][c])) piv = r;
    if (Math.abs(A[piv][c]) < 1e-12) return null;
    [A[c], A[piv]] = [A[piv], A[c]];
    [b[c], b[piv]] = [b[piv], b[c]];
    for (let r = 0; r < k; r++) {
      if (r === c) continue;
      const f = A[r][c] / A[c][c];
      for (let q = c; q < k; q++) A[r][q] -= f * A[c][q];
      b[r] -= f * b[c];
    }
  }
  return b.map((v, i) => v / A[i][i]);
}

const dot = (beta: number[], x: number[]) => beta[0] + x.reduce((s, v, i) => s + v * beta[i + 1], 0);

function plausibel(id: ProfilId, a: GarminAkt): boolean {
  const v = a.v_kmh || 0;
  if (id === 'run') return v >= 5 && v <= 20;
  if (id === 'bike_road' || id === 'bike_gravel') return v >= 8 && v <= 45;
  if (id === 'hike') return v > 0.8 && v <= 8 && (a.km || 0) < 60;
  if (id === 'row') return v >= 6 && v <= 25;
  return true;
}

export function lerneModelle(tage: GarminTage, heute = new Date()): Modell[] {
  const t0 = heute.getTime();
  const out: Modell[] = [];
  for (const p of PROFILE) {
    const proben: Probe[] = [];
    tage.forEach((tag, datum) => {
      for (const a of tag.akt || []) {
        if (!a.typ || !p.typen.includes(a.typ)) continue;
        const min = (a.dauer_s || 0) / 60;
        if (min < 10 || !a.last || a.last < 3 || !a.hf_avg || a.hf_avg < 60) continue;
        if (!plausibel(p.id, a)) continue;
        const alter = (t0 - new Date(datum + 'T12:00:00').getTime()) / 864e5;
        proben.push({
          datum,
          name: a.name,
          min,
          km: a.km || 0,
          hm: a.hm || 0,
          hf: a.hf_avg,
          np: a.watt_np || undefined,
          v: a.v_kmh || 0,
          ess: a.last,
          w: Math.exp((-LN2 * Math.max(0, alter)) / 365),
        });
      }
    });
    if (proben.length < 8) continue;
    const y = proben.map((s) => Math.log(s.ess));
    const w = proben.map((s) => s.w);
    let base: number[] | null;
    if (p.id === 'hike') {
      // Wandern: Last ∝ Dauer (b = 1), größere Touren zählen stärker (viele kurze Spaziergänge mit ~0 Last)
      const hb = wls(proben.map((s) => [s.hf]), proben.map((s) => Math.log(s.ess) - Math.log(s.min)), proben.map((s) => s.w * Math.sqrt(s.ess)));
      base = hb ? [hb[0], 1, hb[1]] : null;
    } else base = wls(proben.map((s) => [Math.log(s.min), s.hf]), y, w);
    if (!base || base[2] <= 0) continue;
    const relErr = (pred: (s: Probe) => number, ps: Probe[]) => ps.reduce((sum, s) => sum + Math.abs(pred(s) - s.ess) / s.ess, 0) / ps.length;
    const m: Modell = {
      id: p.id,
      n: proben.length,
      base,
      fehler: relErr((s) => Math.exp(dot(base!, [Math.log(s.min), s.hf])), p.id === 'hike' && proben.filter((s) => s.ess >= 30).length >= 5 ? proben.filter((s) => s.ess >= 30) : proben),
      hf: { p5: quantil(proben.map((s) => s.hf), 0.05), p25: quantil(proben.map((s) => s.hf), 0.25), p50: quantil(proben.map((s) => s.hf), 0.5), p75: quantil(proben.map((s) => s.hf), 0.75), p95: quantil(proben.map((s) => s.hf), 0.95) },
      min: { p5: quantil(proben.map((s) => s.min), 0.05), p50: quantil(proben.map((s) => s.min), 0.5), p95: quantil(proben.map((s) => s.min), 0.95) },
      hpk: { p50: quantil(proben.filter((s) => s.km > 0.5).map((s) => s.hm / s.km), 0.5) || 0, p90: quantil(proben.filter((s) => s.km > 0.5).map((s) => s.hm / s.km), 0.9) || 0 },
      proben,
    };
    if (p.sport === 'Bike') {
      const pn = proben.filter((s) => s.np && s.np > 50);
      if (pn.length >= 8) {
        const ext = wls(pn.map((s) => [Math.log(s.min), s.hf, s.np!]), pn.map((s) => Math.log(s.ess)), pn.map((s) => s.w));
        const npLink = wls(pn.map((s) => [s.hf]), pn.map((s) => s.np!), pn.map((s) => s.w));
        if (ext) {
          m.ext = ext;
          m.extKey = 'np';
          m.fehlerExt = relErr((s) => Math.exp(dot(ext, [Math.log(s.min), s.hf, s.np!])), pn);
        }
        if (npLink && npLink[1] > 0) m.npLink = npLink;
      }
    }
    if (p.id === 'run') {
      const ext = wls(proben.map((s) => [Math.log(s.min), s.hf, s.v]), y, w);
      if (ext) {
        m.ext = ext;
        m.extKey = 'v';
        m.fehlerExt = relErr((s) => Math.exp(dot(ext, [Math.log(s.min), s.hf, s.v])), proben);
      }
    }
    if (p.id === 'run' || p.id === 'bike_road' || p.id === 'bike_gravel' || p.id === 'row') {
      const pv = proben.filter((s) => s.km > 0.5);
      const vl = wls(pv.map((s) => [s.hf, s.hm / s.km]), pv.map((s) => s.v), pv.map((s) => s.w));
      if (vl && vl[1] > 0) m.vLink = vl;
    }
    if (m.vLink) {
      const pv = proben.filter((s) => s.km > 0.5);
      const o = wls(pv.map((s) => [Math.log(s.min), s.v, s.hm / s.km]), pv.map((s) => Math.log(s.ess)), pv.map((s) => s.w));
      if (o) {
        m.ohneHfV = o;
        m.fehlerOhneHf = relErr((s) => Math.exp(dot(o, [Math.log(s.min), s.v, s.km > 0.5 ? s.hm / s.km : 0])), pv);
      }
    }
    if (p.sport === 'Bike') {
      const pn = proben.filter((s) => s.np && s.np > 50);
      if (pn.length >= 8) {
        const o = wls(pn.map((s) => [Math.log(s.min), s.np!]), pn.map((s) => Math.log(s.ess)), pn.map((s) => s.w));
        if (o) m.ohneHfNp = o;
      }
    }
    if (p.id === 'hike') {
      const hl = wls(proben.map((s) => [s.hm / (s.min / 60), s.v]), proben.map((s) => s.hf), w);
      if (hl) m.hfLink = hl;
      // Gehzeit h = km·x + hm·y (ohne Achsabschnitt), robust: nur bewegte Zeit
      const ph = proben.filter((s) => s.km > 1);
      let sxx = 0, sxy = 0, syy = 0, sx = 0, sy = 0;
      for (const s of ph) {
        const h = s.min / 60;
        sxx += s.w * s.km * s.km; sxy += s.w * s.km * s.hm; syy += s.w * s.hm * s.hm; sx += s.w * s.km * h; sy += s.w * s.hm * h;
      }
      const det = sxx * syy - sxy * sxy;
      let x = det ? (sx * syy - sy * sxy) / det : 0.25;
      let yv = det ? (sxx * sy - sxy * sx) / det : 1 / 400;
      if (!(x > 0.08 && x < 0.5)) x = 0.22; // Fallback ~4,5 km/h
      if (!(yv > 1 / 1200 && yv < 1 / 150)) yv = 1 / 400;
      m.hike = { vFlach: 1 / x, hmProH: 1 / yv };
    }
    out.push(m);
  }
  return out;
}

// ---------- Rechnen ----------

export type Werte = { min?: number; km?: number; hm?: number; hf?: number; np?: number; v?: number };
export type Ergebnis = {
  ess: number;
  lo: number;
  hi: number;
  werte: Required<Pick<Werte, 'min' | 'hf'>> & Werte;
  geschaetzt: (keyof Werte)[];
  modell: 'basis' | 'np' | 'v';
  hinweise: string[];
};

const essBase = (m: Modell, min: number, hf: number) => Math.exp(dot(m.base, [Math.log(min), hf]));
const hfAusEss = (m: Modell, ess: number, min: number) => (Math.log(ess) - m.base[0] - m.base[1] * Math.log(min)) / m.base[2];
const vAusHf = (m: Modell, hf: number, hpk: number) => (m.vLink ? dot(m.vLink, [hf, hpk]) : undefined);
const hfAusV = (m: Modell, v: number, hpk: number) => (m.vLink ? (v - m.vLink[0] - m.vLink[2] * hpk) / m.vLink[1] : undefined);
const npAusHf = (m: Modell, hf: number) => (m.npLink ? dot(m.npLink, [hf]) : undefined);
const hfAusNp = (m: Modell, np: number) => (m.npLink ? (np - m.npLink[0]) / m.npLink[1] : undefined);
const hikeMin = (m: Modell, km: number, hm: number) => (m.hike ? (km / m.hike.vFlach + hm / m.hike.hmProH) * 60 : undefined);

function pruefe(m: Modell, w: { min: number; hf: number }): string[] {
  const h: string[] = [];
  if (w.hf > m.hf.p95 + 3) h.push(`Ø HF ${Math.round(w.hf)} liegt über deinem üblichen Bereich (bis ${Math.round(m.hf.p95)}) – Schätzung unsicher.`);
  if (w.hf < m.hf.p5 - 3) h.push(`Ø HF ${Math.round(w.hf)} liegt unter deinem üblichen Bereich (ab ${Math.round(m.hf.p5)}).`);
  if (w.min > m.min.p95 * 1.3) h.push(`So lange Einheiten (${Math.round(w.min)} min) gibt es in deinen Daten kaum – Schätzung unsicher.`);
  if (w.min < m.min.p5 * 0.7) h.push(`Sehr kurze Einheit – Schätzung unsicher.`);
  return h;
}

/** Einheit → ESS. Fehlende Werte werden aus den gelernten Zusammenhängen ergänzt. */
export function vorwaerts(m: Modell, ein: Werte): Ergebnis | null {
  const g: (keyof Werte)[] = [];
  const w: Werte = { ...ein };
  if (w.km && w.hm == null) {
    w.hm = Math.round(w.km * m.hpk.p50);
    g.push('hm');
  }
  const hpk = w.km && w.km > 0.5 ? (w.hm || 0) / w.km : m.hpk.p50;
  if (!w.v && w.km && w.min) w.v = w.km / (w.min / 60);
  if (!w.min && m.id === 'hike' && w.km) {
    w.min = hikeMin(m, w.km, w.hm || 0);
    g.push('min');
  }
  const hfGegeben = !!w.hf;
  let essOhneHf: number | undefined;
  let fehlerOhneHf: number | undefined;
  if (!w.hf) {
    let hf: number | undefined;
    if (w.min && w.np && m.ohneHfNp) {
      essOhneHf = Math.exp(dot(m.ohneHfNp, [Math.log(w.min), w.np]));
      fehlerOhneHf = (m.fehlerExt ?? m.fehler) * 1.4;
    } else if (w.min && w.v && m.ohneHfV) {
      essOhneHf = Math.exp(dot(m.ohneHfV, [Math.log(w.min), w.v, hpk]));
      fehlerOhneHf = m.fehlerOhneHf ?? m.fehler;
    }
    if (essOhneHf && w.min) hf = hfAusEss(m, essOhneHf, w.min);
    else if (m.id === 'hike' && m.hfLink && w.min && w.km) hf = dot(m.hfLink, [(w.hm || 0) / (w.min / 60), w.km / (w.min / 60)]);
    else if (w.np) hf = hfAusNp(m, w.np);
    w.hf = hf != null && isFinite(hf) ? hf : m.hf.p50;
    g.push('hf');
  }
  if (!w.v && m.id !== 'bike_indoor') {
    const v = m.id === 'hike' ? undefined : vAusHf(m, w.hf, hpk);
    if (v) {
      w.v = v;
      g.push('v');
    }
  }
  if (!w.min && w.km && w.v) {
    w.min = (w.km / w.v) * 60;
    g.push('min');
  }
  if (!w.min) return null;
  if (!w.km && w.v) {
    w.km = (w.v * w.min) / 60;
    g.push('km');
  }
  if (!w.np && m.npLink) {
    w.np = npAusHf(m, w.hf);
    g.push('np');
  }
  let ess = essBase(m, w.min, w.hf);
  let typ: Ergebnis['modell'] = 'basis';
  let fehler = hfGegeben ? m.fehler : m.fehler * 1.5;
  if (essOhneHf) {
    ess = essOhneHf;
    typ = ein.np ? 'np' : 'v';
    fehler = fehlerOhneHf ?? fehler;
  } else if (!hfGegeben) {
    // nichts weiter
  } else if (m.ext && m.extKey === 'np' && ein.np) {
    ess = Math.exp(dot(m.ext, [Math.log(w.min), w.hf, ein.np]));
    typ = 'np';
    fehler = m.fehlerExt ?? fehler;
  } else if (m.ext && m.extKey === 'v' && w.v && !g.includes('v')) {
    ess = Math.exp(dot(m.ext, [Math.log(w.min), w.hf, w.v]));
    typ = 'v';
    fehler = m.fehlerExt ?? fehler;
  }
  const hinweise = pruefe(m, w as { min: number; hf: number });
  if (!hfGegeben) hinweise.push('Mit deiner Ø HF wird die Schätzung deutlich genauer – Garmin rechnet die Last aus der Herzfrequenz.');
  return { ess, lo: ess * (1 - fehler), hi: ess * (1 + fehler), werte: w as Ergebnis['werte'], geschaetzt: g, modell: typ, hinweise };
}

export type Loesung = { min: number; km?: number; hm: number; hf: number; v?: number; np?: number; ess: number; charakter: string; hinweise: string[] };

function charakter(m: Modell, hf: number) {
  if (hf < m.hf.p25) return 'locker';
  if (hf < m.hf.p75) return 'mittel';
  return 'zügig';
}

/** Ziel-ESS + Dauer → HF, Tempo, Distanz, NP */
export function ausDauer(m: Modell, ess: number, min: number, hpk: number): Loesung {
  const hf = hfAusEss(m, ess, min);
  let v = m.id === 'hike' || m.id === 'bike_indoor' ? undefined : vAusHf(m, hf, hpk);
  let km = v ? (v * min) / 60 : undefined;
  if (m.id === 'hike' && m.hike) {
    // Gehzeit = km/vf + km·hpk/hmh → km
    km = (min / 60) / (1 / m.hike.vFlach + hpk / m.hike.hmProH);
    v = km / (min / 60);
  }
  const hm = km ? Math.round(km * hpk) : 0;
  return { min, km, hm, hf, v, np: npAusHf(m, hf), ess, charakter: charakter(m, hf), hinweise: pruefe(m, { min, hf }) };
}

/** Ziel-ESS + Distanz → Dauer, Tempo, HF */
export function ausDistanz(m: Modell, ess: number, km: number, hm: number): Loesung | null {
  const hpk = km > 0.5 ? hm / km : 0;
  if (m.id === 'hike') {
    const min = hikeMin(m, km, hm);
    if (!min) return null;
    const hf = hfAusEss(m, ess, min);
    return { min, km, hm, hf, v: km / (min / 60), ess, charakter: charakter(m, hf), hinweise: pruefe(m, { min, hf }) };
  }
  if (!m.vLink) return null;
  // ESS(min) mit HF aus Tempo km/(min/60): Nullstelle per Bisektion
  const f = (min: number) => {
    const v = km / (min / 60);
    const hf = hfAusV(m, v, hpk)!;
    return { e: essBase(m, min, hf), hf, v };
  };
  let lo = 5, hi = 600;
  const flo = f(lo).e - ess, fhi = f(hi).e - ess;
  if (Math.sign(flo) === Math.sign(fhi)) return null;
  for (let i = 0; i < 60; i++) {
    const mid = (lo + hi) / 2;
    const fm = f(mid).e - ess;
    if (Math.sign(fm) === Math.sign(flo)) lo = mid;
    else hi = mid;
  }
  const min = (lo + hi) / 2;
  const r = f(min);
  return { min, km, hm, hf: r.hf, v: r.v, np: npAusHf(m, r.hf), ess, charakter: charakter(m, r.hf), hinweise: pruefe(m, { min, hf: r.hf }) };
}

/** Variationen für ein Ziel-ESS: kurz & zügig bis lang & locker, plus eine Variante mit mehr Höhenmetern. */
export function variationen(m: Modell, ess: number, hpk: number): (Loesung & { titel: string })[] {
  const out: (Loesung & { titel: string })[] = [];
  // Dauer so wählen, dass die HF bei p90 / p70 / p50 / p30 / p10 deines Bereichs liegt
  const ziele: [string, number][] = [
    ['Kurz & zügig', m.hf.p95 - 0.15 * (m.hf.p95 - m.hf.p75)],
    ['Kompakt', m.hf.p75],
    ['Typisch', m.hf.p50],
    ['Länger & ruhig', m.hf.p25],
    ['Lang & locker', m.hf.p5 + 0.3 * (m.hf.p25 - m.hf.p5)],
  ];
  for (const [titel, hf] of ziele) {
    // ln ess = a + b ln min + c hf → min
    const min = Math.exp((Math.log(ess) - m.base[0] - m.base[2] * hf) / m.base[1]);
    if (!isFinite(min) || min < 10 || min > 720) continue;
    out.push({ titel, ...ausDauer(m, ess, min, hpk) });
  }
  if (m.id !== 'bike_indoor' && m.id !== 'row' && m.hpk.p90 > hpk * 1.5 + 3) {
    const min = Math.exp((Math.log(ess) - m.base[0] - m.base[2] * m.hf.p50) / m.base[1]);
    if (isFinite(min) && min >= 10 && min <= 720) out.push({ titel: 'Mit Höhenmetern', ...ausDauer(m, ess, min, m.hpk.p90) });
  }
  // gleiche Dauer doppelt vermeiden
  return out.filter((x, i, a) => a.findIndex((y) => Math.abs(y.min - x.min) < 4) === i);
}

/** Ähnlichste echte Einheiten (Dauer, HF, hm/km) */
export function aehnliche(m: Modell, w: { min: number; hf: number; km?: number; hm?: number }, n = 3): Probe[] {
  const hpk = w.km && w.km > 0.5 ? (w.hm || 0) / w.km : m.hpk.p50;
  const sd = (xs: number[]) => {
    const mu = xs.reduce((s, v) => s + v, 0) / xs.length;
    return Math.sqrt(xs.reduce((s, v) => s + (v - mu) ** 2, 0) / xs.length) || 1;
  };
  const sMin = sd(m.proben.map((p) => Math.log(p.min)));
  const sHf = sd(m.proben.map((p) => p.hf));
  const sHpk = sd(m.proben.map((p) => (p.km > 0.5 ? p.hm / p.km : 0)));
  return [...m.proben]
    .map((p) => ({
      p,
      d: ((Math.log(p.min) - Math.log(w.min)) / sMin) ** 2 + ((p.hf - w.hf) / sHf) ** 2 + 0.5 * (((p.km > 0.5 ? p.hm / p.km : 0) - hpk) / sHpk) ** 2,
    }))
    .sort((a, b) => a.d - b.d)
    .slice(0, n)
    .map((x) => x.p);
}

// ---------- Format ----------
export function fmtTempo(id: ProfilId, v?: number): string {
  if (!v || !isFinite(v) || v <= 0) return '–';
  const p = PROFILE.find((x) => x.id === id)!;
  if (p.tempo === 'pace') {
    const s = Math.round(3600 / v);
    return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')} min/km`;
  }
  if (p.tempo === 'row') {
    const s = Math.round(1800 / v);
    return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')} /500 m`;
  }
  if (p.tempo === 'kmh') return `${v.toLocaleString('de-DE', { maximumFractionDigits: 1 })} km/h`;
  return '–';
}
export function parseTempo(id: ProfilId, s: string): number | undefined {
  const t = s.trim().replace(',', '.');
  if (!t) return undefined;
  const p = PROFILE.find((x) => x.id === id)!;
  const mm = t.match(/^(\d+):(\d{1,2})$/);
  if (p.tempo === 'pace') {
    const sec = mm ? Number(mm[1]) * 60 + Number(mm[2]) : Number(t) * 60;
    return sec > 0 ? 3600 / sec : undefined;
  }
  if (p.tempo === 'row') {
    const sec = mm ? Number(mm[1]) * 60 + Number(mm[2]) : Number(t) * 60;
    return sec > 0 ? 1800 / sec : undefined;
  }
  const v = Number(t);
  return v > 0 ? v : undefined;
}
export function fmtMin(min?: number): string {
  if (!min || !isFinite(min)) return '–';
  const m = Math.round(min);
  return m >= 60 ? `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')} h` : `${m} min`;
}
export function parseMin(s: string): number | undefined {
  const t = s.trim().replace(',', '.');
  if (!t) return undefined;
  const mm = t.match(/^(\d+):(\d{1,2})$/);
  if (mm) return Number(mm[1]) * 60 + Number(mm[2]);
  const v = Number(t);
  return v > 0 ? v : undefined;
}
