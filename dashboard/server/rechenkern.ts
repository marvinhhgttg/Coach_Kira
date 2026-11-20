// Rechenkern Etappe 1 – die 12 Formelspalten aus dem Blatt `timeline`, 1:1 portiert.
// Entwurf: lokales Modell coach-kira (Ollama). Geprüft und korrigiert in der Cloud:
//  - 7-/28-Tage-Fenster begannen bei Index 2 statt 0 (Sheet-Zeile 2 = Index 0)
//  - SG_Flags: leere Eingaben ergeben "" wie im Sheet (IFERROR → ""); leere Monotonie setzt "M" (Sheets-Textvergleich)
//  - Strain7: Fenster endet gestern, ctl_star: Startwert = CTL-Prognose (Sheet-Zeile 10)
// Datenmodell: rows[i] = Sheet-Zeile i+2, Spaltennamen wie in KK_TIMELINE.

export type Zelle = number | string | null;
export type Zeile = Record<string, Zelle>;
export type Konstanten = Record<string, Zelle>;

const leer = (v: Zelle | undefined): boolean => v === null || v === undefined || v === '';
const zahl = (v: Zelle | undefined): number => {
  if (leer(v)) return NaN;
  if (typeof v === 'number') return v;
  return Number(String(v).replace(',', '.'));
};
// In Sheets-Produkten und -Summen zählt eine leere Zelle als 0
const nullWenigstens = (v: Zelle | undefined): number => (leer(v) ? 0 : zahl(v));

export const KERN_SPALTEN = [
  'fKEI', 'load_star', 'ctl_star', 'ctl_gap', 'coachE_Smart_Gains', 'SG_Flags',
  'Monotony7', 'Strain7', 'coachE_ATL_morning', 'coachE_CTL_morning', 'TE_Balance_Trend',
] as const;

function monotonie(w: number[]): number | null {
  if (w.length < 2) return null;
  const mu = w.reduce((s, x) => s + x, 0) / w.length;
  const sd = Math.sqrt(w.reduce((s, x) => s + (x - mu) ** 2, 0) / w.length);
  return sd === 0 ? null : mu / sd;
}

/** Rechnet eine Zeile. Bezüge auf andere Spalten und Zeilen lesen die Sheet-Werte (Schattenbetrieb). */
export function rechneZeile(rows: Zeile[], i: number, k: Konstanten): Record<string, Zelle> {
  const r = rows[i];
  const vor = (n: number): Zeile | undefined => rows[i - n];
  const c = (name: string) => zahl(k[name]);

  // 1 fKEI = MIN(1,15; MAX(0,85; 1 + 0,05 * AN))
  const fKEI = Math.min(1.15, Math.max(0.85, 1 + 0.05 * nullWenigstens(r.coachE_Smart_Gains)));

  // 2 load_star = AF * AJ
  const load_star = nullWenigstens(r.coachE_ESS_day) * nullWenigstens(r.fKEI);

  // 3 ctl_star = (1-α)·AL[r-1] + α·(S·AK + B + wH·Q[r+1] + wS·R[r+1])
  const nach = rows[i + 1];
  const ctl_star = leer(vor(1)?.ctl_star ?? null) && i === 8
    ? nullWenigstens(r.coachE_CTL_forecast) // Sheet-Zeile 10: Startwert
    : (1 - c('alpha_CTL')) * nullWenigstens(vor(1)?.ctl_star ?? null) +
    c('alpha_CTL') * (
      c('S_CTL') * nullWenigstens(r.load_star) + c('B_CTL') +
      c('wH_CTL') * nullWenigstens(nach?.sleep_hours ?? null) +
      c('wS_CTL') * nullWenigstens(nach?.sleep_score_0_100 ?? null)
    );

  // 4 ctl_gap = AH[Sheet-Zeile 10] − AL  (feste Zeile 10 wie im Sheet)
  const ctl_gap = nullWenigstens(rows[8]?.coachE_CTL_forecast ?? null) - nullWenigstens(r.ctl_star);

  // 5 Smart Gains
  let coachE_Smart_Gains: Zelle = null;
  const ah7 = vor(7)?.coachE_CTL_forecast;
  if (!leer(r.coachE_CTL_forecast) && !leer(ah7 ?? null) && !leer(r.Monotony7)) {
    coachE_Smart_Gains =
      ((zahl(r.coachE_CTL_forecast) - zahl(ah7 ?? null)) * 10) /
      (Math.max(nullWenigstens(r.coachE_ACWR_forecast), 0.8) * (1 + zahl(r.Monotony7))) * 0.1;
  }

  // 6 SG_Flags
  let SG_Flags = '';
  try {
    const u = String(r.hrv_threshholds ?? '');
    const p = u.indexOf(';');
    if (p < 0) throw new Error('kein ;');
    const unten = Number(u.slice(0, p).replace(',', '.'));
    if (!Number.isFinite(unten)) throw new Error('keine Zahl');
    let f = nullWenigstens(r.hrv_status) < unten - 2 ? 'H ' : '';
    // Sheets-Eigenheit: Monotony7 = "" (Text) ist im Vergleich immer größer als jede Zahl → "M"
    f += leer(r.Monotony7) || zahl(r.Monotony7) > 1.6 ? 'M' : '';
    SG_Flags = f.trim();
  } catch {
    SG_Flags = '';
  }

  // 7/8 Monotony7 und Strain7 über die letzten bis zu 7 Tage (inkl. heute)
  const fenster7: number[] = [];
  for (let j = Math.max(0, i - 6); j <= i; j++) {
    const v = rows[j]?.coachE_ESS_day;
    if (!leer(v ?? null) && Number.isFinite(zahl(v ?? null))) fenster7.push(zahl(v ?? null));
  }
  const Monotony7 = monotonie(fenster7);
  // Strain7 im Sheet: ROWS($AF$2:…ROW()) = ROW()−1 → Fenster endet gestern (bewusst 1:1 übernommen)
  const fensterGestern: number[] = [];
  for (let j = Math.max(0, i - 7); j <= i - 1; j++) {
    const v = rows[j]?.coachE_ESS_day;
    if (!leer(v ?? null) && Number.isFinite(zahl(v ?? null))) fensterGestern.push(zahl(v ?? null));
  }
  const mono2 = monotonie(fensterGestern);
  const Strain7: Zelle = mono2 === null ? null : mono2 * fensterGestern.reduce((s, x) => s + x, 0);

  // 9/10 Morgenwerte aus der Prognose von vorgestern
  const v2 = vor(2);
  const coachE_ATL_morning: Zelle = !v2 || leer(v2.coachE_ESS_day)
    ? null
    : -3.442725 + 0.815463593 * nullWenigstens(v2.coachE_ATL_forecast) + 1.088519637 * zahl(v2.coachE_ESS_day);
  const coachE_CTL_morning: Zelle = !v2 || leer(v2.coachE_CTL_forecast)
    ? null
    : -3.95042 + 0.97283824 * zahl(v2.coachE_CTL_forecast);

  // 11 TE_Balance_Trend über 28 Tage
  let summe = 0;
  let gewichtet = 0;
  for (let j = Math.max(0, i - 27); j <= i; j++) {
    const z = rows[j];
    const af = nullWenigstens(z?.coachE_ESS_day ?? null);
    summe += af;
    const treffer = nullWenigstens(z?.Target_Anaerobic_TE ?? null) >= 1 || nullWenigstens(z?.Target_Aerobic_TE ?? null) >= 4;
    if (treffer) gewichtet += af;
  }
  const TE_Balance_Trend = summe === 0 ? 0 : gewichtet / summe;

  return { fKEI, load_star, ctl_star, ctl_gap, coachE_Smart_Gains, SG_Flags, Monotony7, Strain7, coachE_ATL_morning, coachE_CTL_morning, TE_Balance_Trend };
}
