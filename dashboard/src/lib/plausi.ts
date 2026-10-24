export interface Befund {
  feld: string;
  wert: string;
  erwartet: string;
  grund: string;
  schwere: 'warnung' | 'hinweis';
}

/** Alle Feldschlüssel, die in plausiCheck geprüft werden (ohne 'Makros'). */
export const PLAUSI_FELDER = [
  'sleep_hours',
  'sleep_score_0_100',
  'rhr_bpm',
  'hrv_status',
  'hrv_threshholds',
  'Garmin_Training_Readiness',
  'garminATL',
  'garminCTL',
  'fbATL_obs',
  'fbCTL_obs',
  'kcal_in',
  'kcal_out',
  'carb_g',
  'protein_g',
  'fat_g',
  'load_fb_day',
  'Aerobic_TE',
  'Anaerobic_TE',
];

export const FELD_LABEL: Record<string, string> = {
  sleep_hours: 'Schlaf (h)',
  sleep_score_0_100: 'Schlafscore',
  rhr_bpm: 'Ruhepuls',
  hrv_status: 'HRV',
  hrv_threshholds: 'HRV-Normalband',
  Garmin_Training_Readiness: 'Trainingsbereitschaft',
  garminATL: 'Garmin ATL',
  garminCTL: 'Garmin CTL',
  fbATL_obs: 'ATL',
  fbCTL_obs: 'CTL',
  kcal_in: 'kcal gegessen',
  kcal_out: 'kcal verbraucht',
  carb_g: 'Kohlenhydrate',
  protein_g: 'Protein',
  fat_g: 'Fett',
  load_fb_day: 'Belastung',
  Aerobic_TE: 'TE aerob',
  Anaerobic_TE: 'TE anaerob',
};

/** Zahlen-Parser: '', null, undefined → null; '7,5' → 7.5. */
export function num(v: unknown): number | null {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = String(v).trim().replace(',', '.');
  if (s === '') return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function median(values: number[]): number {
  const s = [...values].sort((a, b) => a - b);
  const n = s.length;
  if (!n) return 0;
  const m = Math.floor(n / 2);
  return n % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

function fmt(v: number | null): string {
  return v == null ? '–' : v.toFixed(1).replace('.', ',');
}
const f0 = (v: number) => v.toFixed(0);
const f1 = (v: number) => v.toFixed(1);

/** Exakte Portierung von plausi_check aus kira_extras.py (Mac-Dienst). */
export function plausiCheck(row: Record<string, unknown>, hist: Record<string, unknown>[]): Befund[] {
  const med = (feld: string, min = 5): number | null => {
    const vals = hist.map((h) => num(h[feld])).filter((x): x is number => x !== null);
    return vals.length < min ? null : median(vals);
  };
  const out: Befund[] = [];
  const add = (feld: string, wert: unknown, erwartet: string, grund: string, schwere: 'warnung' | 'hinweis') =>
    out.push({ feld, wert: typeof wert === 'number' ? fmt(wert).replace(/,0$/, '') : String(wert), erwartet, grund, schwere });
  const prev = hist.length ? hist[hist.length - 1] : null;

  const sh = num(row.sleep_hours);
  if (sh !== null) {
    const m = med('sleep_hours');
    if (sh < 3 || sh > 12) add('sleep_hours', sh, `ca. ${fmt(m)} h`, 'außerhalb 3–12 Stunden', 'warnung');
    else if (m !== null && Math.abs(sh - m) > 3) add('sleep_hours', sh, `ca. ${fmt(m)} h`, `Ausreißer – ${f1(sh)} statt ca. ${fmt(m)}`, 'hinweis');
    const ss = num(row.sleep_score_0_100);
    if (ss !== null) {
      if (ss < 0 || ss > 100) add('sleep_score_0_100', ss, '0–100', 'außerhalb gültigen Bereichs', 'warnung');
      else {
        if (ss >= 80 && sh < 5) add('sleep_score_0_100', ss, `höher bei ${f1(sh)}h Schlaf`, 'Score passt nicht zur Schlafdauer', 'hinweis');
        if (ss < 40 && sh > 8) add('sleep_score_0_100', ss, `niedriger bei ${f1(sh)}h Schlaf`, 'Score passt nicht zur Schlafdauer', 'hinweis');
      }
    }
  }

  const rhr = num(row.rhr_bpm);
  if (rhr !== null) {
    const m = med('rhr_bpm');
    if (rhr < 30 || rhr > 100) add('rhr_bpm', rhr, `ca. ${fmt(m)} bpm`, 'außerhalb 30–100 bpm', 'warnung');
    else if (m !== null && Math.abs(rhr - m) > 8) add('rhr_bpm', rhr, `ca. ${fmt(m)} bpm`, `Ausreißer – ${f0(rhr)} statt ca. ${fmt(m)}`, 'hinweis');
  }

  const hrv = num(row.hrv_status);
  if (hrv !== null) {
    const m = med('hrv_status');
    if (hrv < 10 || hrv > 200) add('hrv_status', hrv, `ca. ${fmt(m)} ms`, 'außerhalb 10–200 ms', 'warnung');
    else if (m !== null && (hrv < 0.6 * m || hrv > 1.5 * m)) add('hrv_status', hrv, `ca. ${fmt(m)} ms`, `Ausreißer – ${f0(hrv)} statt ca. ${fmt(m)}`, 'hinweis');
  }

  if (row.hrv_threshholds !== undefined && row.hrv_threshholds !== null) {
    const s = String(row.hrv_threshholds).trim();
    const parts = s.includes(';') ? s.split(';').map((x) => x.trim()) : [];
    let ok = false;
    if (parts.length === 2) {
      const a = Number(parts[0].replace(',', '.'));
      const b = Number(parts[1].replace(',', '.'));
      ok = Number.isFinite(a) && Number.isFinite(b) && parts[0] !== '' && parts[1] !== '' && a < b;
    }
    if (!ok) add('hrv_threshholds', s, 'Format a;b mit Zahlen a<b', 'Normalband unten;oben', 'warnung');
  }

  const tr = num(row.Garmin_Training_Readiness);
  if (tr !== null) {
    const m = med('Garmin_Training_Readiness');
    if (tr < 0 || tr > 100) add('Garmin_Training_Readiness', tr, '0–100', 'außerhalb 0–100', 'warnung');
    else if (m !== null && Math.abs(tr - m) > 35) add('Garmin_Training_Readiness', tr, `ca. ${fmt(m)}`, `Ausreißer – ${f0(tr)} statt ca. ${fmt(m)}`, 'hinweis');
  }

  for (const af of ['garminATL', 'garminCTL', 'fbATL_obs', 'fbCTL_obs']) {
    const val = num(row[af]);
    if (val === null) continue;
    if (val <= 0) add(af, val, '> 0', `${af} ≤ 0`, 'warnung');
    else if (prev) {
      const pv = num(prev[af]);
      if (pv !== null && pv > 0) {
        const rel = Math.abs(val - pv) / pv;
        if (rel > 0.25) add(af, val, `ca. ${fmt(pv)} (Vortag)`, `Sprung > 25 % zum Vortag (${f0(rel * 100)}%)`, 'hinweis');
      }
    }
  }

  const ki = num(row.kcal_in);
  if (ki !== null && (ki < 800 || ki > 6500)) add('kcal_in', ki, '800–6500', `${f0(ki)} kcal außerhalb Bereich`, 'warnung');
  const ko = num(row.kcal_out);
  if (ko !== null && (ko < 1200 || ko > 8000)) add('kcal_out', ko, '1200–8000', `${f0(ko)} kcal außerhalb Bereich`, 'warnung');

  const cg = num(row.carb_g), pg = num(row.protein_g), fg = num(row.fat_g);
  if (cg !== null && pg !== null && fg !== null && ki !== null) {
    const k = 4 * cg + 4 * pg + 9 * fg;
    if (ki > 0 && Math.abs(k - ki) / ki > 0.3) add('Makros', k, `${f0(ki)} kcal`, `Makros ergeben ${f0(k)} kcal, eingetragen ${f0(ki)}`, 'hinweis');
  }
  if (pg !== null && pg > 400) add('protein_g', pg, '< 400', `${f0(pg)} g zu viel`, 'warnung');
  if (cg !== null && cg > 1000) add('carb_g', cg, '< 1000', `${f0(cg)} g zu viel`, 'warnung');
  if (fg !== null && fg > 400) add('fat_g', fg, '< 400', `${f0(fg)} g zu viel`, 'warnung');

  const lf = num(row.load_fb_day);
  if (lf !== null) {
    const m = med('load_fb_day');
    if (lf < 0 || lf > 800) add('load_fb_day', lf, `0–800 (ca. ${m ? fmt(m) : 'mind. 5 Werte'})`, `${f0(lf)} außerhalb Bereich`, 'warnung');
    else if (m !== null && m > 20 && lf > 2.5 * m) add('load_fb_day', lf, `ca. ${fmt(m)}`, `${f0(lf)} > 2,5× Median`, 'hinweis');
  }

  for (const te of ['Aerobic_TE', 'Anaerobic_TE']) {
    const v = num(row[te]);
    if (v !== null && (v < 0 || v > 5)) add(te, v, '0–5', `${f1(v)} außerhalb Bereich`, 'warnung');
  }
  return out;
}
