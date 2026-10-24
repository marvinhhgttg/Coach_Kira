export const PLATZHALTER: Record<string, number[]> = { kcal_in: [2100] };

export function istPlatzhalter(feld: string, wert: unknown): boolean {
  const v = PLATZHALTER[feld]; if (!v) return false;
  const n = typeof wert === 'number' ? wert : parseFloat(String(wert).replace(',', '.'));
  return Number.isFinite(n) && v.includes(n);
}

export type Urteil = 'belastbar' | 'eingeschränkt' | 'zu wenig Daten';

export interface FeldQualitaet {
  key: string; label: string; quote30: number; quote90: number;
  quoteAlle: number; platzhalter: number; laengsteSerie: number; urteil: Urteil;
}

const TEXT_F = new Set(['Sport_x', 'Zone', 'hrv_threshholds']);
function istZahlenfeld(k: string): boolean { return !TEXT_F.has(k); }

function addDays(ymd: string, d: number): string {
  const ms = new Date(ymd + 'T12:00:00Z').getTime() + d * 864e5;
  return new Date(ms).toISOString().slice(0, 10);
}

function toNum(v: unknown): number | null {
  if (v == null || String(v).trim() === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function maxRun(nums: (number | null)[]): number {
  let best = 0, run = 1;
  if (!nums.length) return 0;
  for (let i = 1; i < nums.length; i++) {
    if (nums[i] != null && nums[i - 1] != null && nums[i] === nums[i - 1]) run++;
    else { best = Math.max(best, run); run = 1; }
  }
  return Math.max(best, run);
}

export function pruefeQualitaet(
  rows: Array<Record<string, unknown>>,
  felder: ReadonlyArray<{ key: string; label: string }>,
  heute: string,
): { felder: FeldQualitaet[]; doppelteDaten: string[]; fehlendeTage: string[]; tage: number } {
  if (!felder.length) return { felder: [], doppelteDaten: [], fehlendeTage: [], tage: 0 };

  const seen = new Set<string>();
  const dupCnt = new Map<string, number>();
  const firstIdx = new Map<string, number>();

  for (let i = 0; i < rows.length; i++) {
    const d = String(rows[i].date).slice(0, 10);
    if (d > heute) continue;
    dupCnt.set(d, (dupCnt.get(d) || 0) + 1);
    if (!seen.has(d)) { seen.add(d); firstIdx.set(d, i); }
  }

  const doppelteDaten = Array.from(dupCnt.entries()).filter(([, c]) => c > 1).map(([d]) => d).sort();
  const datums = Array.from(seen).sort();

  if (!datums.length) return { felder: [], doppelteDaten, fehlendeTage: [], tage: 0 };

  const er = datums[0], le = datums[datums.length - 1];
  const erMs = new Date(er + 'T12:00:00Z').getTime();
  const leMs = new Date(le + 'T12:00:00Z').getTime();
  const tage = Math.max(1, Math.round((leMs - erMs) / 864e5) + 1);
  const datumsSet = new Set(datums);
  const fehlendeTage: string[] = [];

  for (let i = 0; i < tage && fehlendeTage.length < 50; i++) {
    const d = addDays(er, i);
    if (!datumsSet.has(d)) fehlendeTage.push(d);
  }

  const qualitaeten: FeldQualitaet[] = [];

  for (const feld of felder) {
    const isNum = istZahlenfeld(feld.key);
    let gez30 = 0, gez90 = 0, gesAlle = 0, platzhalterZ = 0;
    const numVals: (number | null)[] = [];

    for (let i = 0; i < tage; i++) {
      const d = addDays(er, i);
      const di = datums.indexOf(d);
      let w: unknown;
      if (di >= 0) w = rows[firstIdx.get(datums[di])!][feld.key] ?? undefined;
      else w = undefined;

      const isPlatz = istPlatzhalter(feld.key, w);
      if (isPlatz) platzhalterZ++;

      const filled = !isPlatz && w !== undefined && w !== null && String(w).trim() !== '';
      if (filled) gesAlle++;
      const in30 = i >= tage - 30 && i < tage;
      if (in30 && filled) gez30++;
      const in90 = i >= tage - 90 && i < tage;
      if (in90 && filled) gez90++;

      numVals.push(isNum && !isPlatz ? toNum(w) : null);
    }

    const q30 = Math.min(30, tage), q90 = Math.min(90, tage);
    const quote30 = gez30 / q30, quote90 = gez90 / q90, quoteAlle = gesAlle / tage;
    let laengsteSerie = isNum ? maxRun(numVals) : 0;

    let urteil: Urteil = 'zu wenig Daten';
    if (quote90 >= 0.8) urteil = 'belastbar';
    else if (quote90 >= 0.4) urteil = 'eingeschränkt';

    if (isNum && feld.key !== 'elev_m_day' && feld.key !== 'load_fb_day'
        && feld.key !== 'befinden_1_5' && laengsteSerie >= 10 && urteil === 'belastbar')
      urteil = 'eingeschränkt';

    qualitaeten.push({ key: feld.key, label: feld.label, quote30, quote90, quoteAlle, platzhalter: platzhalterZ, laengsteSerie, urteil });
  }

  return { felder: qualitaeten, doppelteDaten, fehlendeTage, tage };
}
