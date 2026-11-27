// Rechenkern Etappe 1c – Prognose- und Beobachtungsformeln (ATL/CTL/ACWR) aus `timeline`, 1:1 portiert.
// Entwurf: lokales Modell coach-kira (Ollama), in der Cloud geprüft und nur gestrafft.
// Bezüge auf andere Zeilen lesen die Sheet-Werte (Schattenbetrieb), nicht eigene Ergebnisse.

import { atlStep, ctlStep } from '../src/kern/formeln';

type Zelle = number | string | null;
type Zeile = Record<string, Zelle>;

const leer = (v: Zelle | undefined) => v === null || v === undefined || v === '';
const n0 = (v: Zelle | undefined) => (leer(v) ? 0 : Number(v));

function schnitt(rows: Zeile[], i: number, spalte: string): number | null {
  const w: number[] = [];
  for (let k = Math.max(0, i - 27); k <= i; k++) {
    const v = rows[k]?.[spalte];
    if (!leer(v ?? null) && Number.isFinite(Number(v))) w.push(Number(v));
  }
  return w.length ? w.reduce((s, x) => s + x, 0) / w.length : null; // alle leer → Fehler → null
}

/** ATL: −8,819849 + 0,83645661·ATL[r−1] + 1,24059231·Last[r] */
const atl = (vorher: Zelle | undefined, last: Zelle) => atlStep(n0(vorher), Number(last));
/** CTL: 0,7·(−3,95042 + 0,97283824·CTL[r−1] + 0,2336734·Last[r]) + 0,3·7,1·Ø28(Last) */
function ctl(rows: Zeile[], i: number, vorher: Zelle | undefined, spalteLast: string): number | null {
  const s = schnitt(rows, i, spalteLast);
  if (s === null) return null;
  return ctlStep(n0(vorher), Number(rows[i][spalteLast]), s);
}
const quote = (a: Zelle, b: Zelle) => (n0(b) === 0 ? null : n0(a) / n0(b));

export const PROGNOSE_SPALTEN = ['coachE_ATL_forecast', 'coachE_CTL_forecast', 'coachE_ACWR_forecast', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs'] as const;

export function rechnePrognose(rows: Zeile[], i: number): Record<string, Zelle> {
  const r = rows[i];
  const v = rows[i - 1];
  const af = r.coachE_ESS_day;
  const e = r.load_fb_day;
  return {
    coachE_ATL_forecast: leer(af) ? null : atl(v?.coachE_ATL_forecast, af),
    coachE_CTL_forecast: leer(af) ? null : ctl(rows, i, v?.coachE_CTL_forecast, 'coachE_ESS_day'),
    coachE_ACWR_forecast: leer(af) ? null : quote(r.coachE_ATL_forecast, r.coachE_CTL_forecast),
    fbATL_obs: leer(e) ? null : atl(v?.fbATL_obs, e),
    fbCTL_obs: leer(e) ? null : ctl(rows, i, v?.fbCTL_obs, 'load_fb_day'),
    fbACWR_obs: leer(e) ? null : quote(r.fbATL_obs, r.fbCTL_obs),
  };
}

/**
 * Etappe 1d – Plan durchrechnen (Kette): ab Zeile `start` werden ATL/CTL/ACWR aus den eigenen
 * Ergebnissen fortgeschrieben. `ess` überschreibt optional das geplante ESS je Datum (Was-wäre-wenn).
 * Die Zeilen davor sind Vergangenheit und bleiben unverändert.
 */
export function simuliere(rows: Zeile[], start: number, ess: Record<string, number> = {}) {
  const kopie: Zeile[] = rows.map((r) => ({ ...r }));
  const aus: { date: Zelle; ess: Zelle; atl: number | null; ctl: number | null; acwr: number | null }[] = [];
  for (let i = start; i < kopie.length; i++) {
    const r = kopie[i];
    const d = String(r.date);
    if (d in ess) r.coachE_ESS_day = ess[d];
    if (leer(r.coachE_ESS_day)) break; // Ende des Plans
    const p = rechnePrognose(kopie, i);
    r.coachE_ATL_forecast = p.coachE_ATL_forecast;
    r.coachE_CTL_forecast = p.coachE_CTL_forecast;
    r.coachE_ACWR_forecast = quote(r.coachE_ATL_forecast, r.coachE_CTL_forecast);
    aus.push({ date: r.date, ess: r.coachE_ESS_day, atl: r.coachE_ATL_forecast as number, ctl: r.coachE_CTL_forecast as number, acwr: r.coachE_ACWR_forecast as number });
  }
  return aus;
}
