// ─── Typen (Export 1) ─────────────────────────────────────────────

export interface P4Tag {
  date: string;
  load: number | null;
  sport: string;
  zone: string;
  sleep: number | null;
  readiness: number | null;
  hrv: number | null;
  rhr: number | null;
  ctl: number | null;
  es: number | null;
  km: number | null;
  hm: number | null;
}

// ─── Export 2 (BlockKennzahlen) ──────────────────────────────────

export interface BlockKennzahlen {
  last: number;
  wochenLast: number;
  ruhetage: number;
  laengste: number;
  sportMix: Record<string, number>;
  zonen: { locker: number; grau: number; hart: number };
  schlaf: number | null;
  readiness: number | null;
}

// ─── Export 3 (Block) ────────────────────────────────────────────

export interface Block {
  start: string;
  ende: string;
  esDelta: number | null;
  ctlDelta: number | null;
  k: BlockKennzahlen;
}

export interface RezeptbuchResult {
  top: Block[];
  schwach: Block[];
  alle: number;
  muster: string[];
}

// ─── Export 4 (ZeitraumVergleich) ────────────────────────────────

export interface VergleichsZeile {
  key: string;
  label: string;
  a: number | null;
  b: number | null;
  einheit: string;
  besserWennHoeher: boolean | null;
}

// ─── Export 5 (Rueckblick) ──────────────────────────────────────

export interface Rueckblick {
  titel: string;
  tage: number;
  summeLast: number;
  einheiten: number;
  km: number;
  hm: number;
  topSport: { name: string; anteil: number };
  besterTag: { date: string; load: number };
  laengsteSerie: { tage: number; bis: string };
  besterSchlafMonat: { monat: string; schlaf: number } | null;
  besterReadinessMonat: { monat: string; readiness: number } | null;
  esMax: { date: string; es: number } | null;
  esStart: number | null;
  esEnde: number | null;
  ctlStart: number | null;
  ctlEnde: number | null;
  besterBlock: Block | null;
  ruhigsterMonat: { monat: string; last: number } | null;
}

// ─── Export 6 (Tourform) ────────────────────────────────────────

import { fetchChartData } from './api';
import { zonenKlasse, sportGruppe } from './recoveryPrice';
import { zielBereitschaft, tourLast as tpl, Ziel } from './planning';

export interface Tourform {
  prozent: number;
  volumen: number;
  lange: number;
  hoehe: number | null;
  tourLast: number;
  wochenLastIst: number;
  wochenLastZiel: number;
  langeTage: number;
  langeZiel: number;
  maxHm: number | null;
  hmZiel: number;
  hinweise: string[];
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────

/** Parse a possibly comma-separated numeric value to number|null */
function toNum(v: unknown): number | null {
  if (v === null || v === undefined || v === '' || String(v).trim() === '') return null;
  const s = String(v).trim();
  if (s.toLowerCase() === 'nan' || s.toLowerCase() === '#ref!') return null;
  const n = Number(s.replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

/** CSV row → Record<string, string> */
function parseCsvLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') { inQuotes = false; } else { current += c; }
    } else {
      if (c === '"') { inQuotes = true; }
      else if (c === ',') { result.push(current); current = ''; }
      else { current += c; }
    }
  }
  result.push(current);
  return result;
}

function parseCsvRows(csv: string): { headers: string[]; rows: string[][] } {
  const lines = csv.split(/\r?\n/).filter((l) => l.trim() !== '');
  if (lines.length === 0) return { headers: [], rows: [] };
  const headers = parseCsvLine(lines[0]);
  const dataRows = lines.slice(1).map(parseCsvLine);
  return { headers, rows: dataRows };
}

// ─── Export 1: ladeP4() ──────────────────────────────────────────

const METRIC_NAMES: Record<string, string> = {
  'load_fb_day': 'load_fb_day',
  'Sport_x': 'Sport_x',
  'Zone': 'Zone',
  'sleep_hours': 'sleep_hours',
  'Garmin_Training_Readiness': 'Garmin_Training_Readiness',
  'hrv_status': 'hrv_status',
  'rhr_bpm': 'rhr_bpm',
  'fbCTL_obs': 'fbCTL_obs',
  'garminEnduranceScore': 'garminEnduranceScore',
  'distance_km_day': 'distance_km_day',
  'elev_m_day': 'elev_m_day',
};

const CSV_KEYS = ['load_fb_day','Sport_x','Zone','sleep_hours','Garmin_Training_Readiness','hrv_status','rhr_bpm','fbCTL_obs','garminEnduranceScore','distance_km_day','elev_m_day'];
const CSV_IDX: Record<string, number> = {};

function buildCsvIndex(headers: string[]): void {
  for (let i = 0; i < headers.length; i++) {
    CSV_IDX[headers[i]] = i;
  }
}

let p4Cache: P4Tag[] | null = null;
let p4Promise: Promise<P4Tag[]> | null = null;

function collectCsvTags(csv: string, csvKeys: string[]): P4Tag[] {
  const { headers, rows } = parseCsvRows(csv);
  if (headers.length === 0) return [];
  buildCsvIndex(headers);

  const today = new Date().toISOString().split('T')[0];
  const tags: P4Tag[] = [];
  const seenDates = new Set<string>();

  for (const row of rows) {
    const dateRaw = row[CSV_IDX['date'] as number];
    if (!dateRaw) continue;
    const date = String(dateRaw).split('T')[0];

    // Filter: up to is_today == 1, then date <= today
    const isTodayRaw = row[CSV_IDX['is_today'] as number];
    if (String(isTodayRaw).trim() === '1') {
      // Include this row, but stop here — only include dates <= today too
    }

    if (date > today) continue; // filter out future dates

    if (seenDates.has(date)) continue; // keep first occurrence
    seenDates.add(date);

    const load = toNum(row[CSV_IDX['load_fb_day'] as number]);
    const sportRaw = String(row[CSV_IDX['Sport_x'] as number] || '').trim();
    const zoneRaw = String(row[CSV_IDX['Zone'] as number] || '').trim();
    const sleep = toNum(row[CSV_IDX['sleep_hours'] as number]);
    const readiness = toNum(row[CSV_IDX['Garmin_Training_Readiness'] as number]);
    const hrv = toNum(row[CSV_IDX['hrv_status'] as number]);
    const rhr = toNum(row[CSV_IDX['rhr_bpm'] as number]);
    const ctl = toNum(row[CSV_IDX['fbCTL_obs'] as number]);
    const es = toNum(row[CSV_IDX['garminEnduranceScore'] as number]);
    const km = toNum(row[CSV_IDX['distance_km_day'] as number]);
    const hm = toNum(row[CSV_IDX['elev_m_day'] as number]);

    tags.push({ date, load, sport: sportRaw, zone: zoneRaw, sleep, readiness, hrv, rhr, ctl, es, km, hm });
  }

  tags.sort((a, b) => a.date.localeCompare(b.date));
  return tags;
}

// Build cache from CSV text (for testing)
function parseP4FromCsv(csv: string): P4Tag[] {
  const tags = collectCsvTags(csv, CSV_KEYS);
  p4Cache = tags;
  return tags;
}

export async function ladeP4(): Promise<P4Tag[]> {
  if (p4Cache) return p4Cache;
  if (!p4Promise) {
    p4Promise = (async () => {
      try {
        const r: any = await fetchChartData('9999d' as any, CSV_KEYS);
        const data = (r && r.data) as Array<Record<string, unknown>> | undefined;
        if (!data || !Array.isArray(data)) { p4Cache = []; return []; }

        const tags: P4Tag[] = [];
        const seenDates = new Map<string, boolean>();

        for (const row of data) {
          const date = String(row.date || '').split('T')[0];
          if (date === '' || date > new Date().toISOString().split('T')[0]) continue;
          if (seenDates.has(date)) continue;
          seenDates.set(date, true);

          const load = toNum(row.load_fb_day);
          const sportRaw = String(row.Sport_x || '').trim();
          const zoneRaw = String(row.Zone || '').trim();
          const sleep = toNum(row.sleep_hours);
          const readiness = toNum(row.Garmin_Training_Readiness);
          const hrv = toNum(row.hrv_status);
          const rhr = toNum(row.rhr_bpm);
          const ctl = toNum(row.fbCTL_obs);
          const es = toNum(row.garminEnduranceScore);
          const km = toNum(row.distance_km_day);
          const hm = toNum(row.elev_m_day);

          tags.push({ date, load, sport: sportRaw, zone: zoneRaw, sleep, readiness, hrv, rhr, ctl, es, km, hm });
        }

        tags.sort((a, b) => a.date.localeCompare(b.date));
        p4Cache = tags;
        return tags;
      } catch {
        p4Cache = [];
        return [];
      } finally {
        p4Promise = null;
      }
    })();
  }
  return p4Promise;
}

// ─── Export 2: kennzahlen() ──────────────────────────────────────

export function kennzahlen(t: P4Tag[]): BlockKennzahlen {
  const n = t.length;
  if (n === 0) return { last: 0, wochenLast: 0, ruhetage: 0, laengste: 0, sportMix: {}, zonen: { locker: 0, grau: 0, hart: 0 }, schlaf: null, readiness: null };

  // last = Σ load (only non-null)
  let sumLoad = 0;
  for (const tag of t) {
    if (tag.load != null) sumLoad += tag.load;
  }

  // wochenLast = last / Tage * 7
  const daysWithLoadCount = t.filter((tag) => tag.load != null).length;
  const last = sumLoad;
  const wochenLast = daysWithLoadCount > 0 ? (last / n) * 7 : 0;

  // ruhetage = days with load <= 20 or null
  let ruhetage = 0;
  for (const tag of t) {
    if (tag.load == null || tag.load <= 20) ruhetage++;
  }

  // laengste = max load (only non-null)
  let laengste = 0;
  for (const tag of t) {
    if (tag.load != null && tag.load > laengste) laengste = tag.load;
  }

  // sportMix: relative share per sportGruppe (only days where load > 0)
  const sportGroups = new Map<string, number>(); // sportGroup → total load
  let totalSportLoad = 0;
  for (const tag of t) {
    if (tag.load != null && tag.load > 0) {
      const grp = sportGruppe(tag.sport);
      sportGroups.set(grp, (sportGroups.get(grp) || 0) + tag.load);
      totalSportLoad += tag.load;
    }
  }
  const sportMix: Record<string, number> = {};
  for (const [grp, load] of sportGroups) {
    sportMix[grp] = totalSportLoad > 0 ? load / totalSportLoad : 0;
  }

  // zonen: relative share per zonenKlasse (excluding 'off')
  const zoneLoads = new Map<string, number>(); // locker, grau, hart
  let totalZoneLoad = 0;
  for (const tag of t) {
    if (tag.load != null && tag.load > 0) {
      const kl = zonenKlasse(tag.zone);
      if (kl !== 'off') {
        zoneLoads.set(kl, (zoneLoads.get(kl) || 0) + tag.load);
        totalZoneLoad += tag.load;
      }
    }
  }
  const zonen = { locker: 0, grau: 0, hart: 0 };
  for (const [kl, load] of zoneLoads) {
    if (kl === 'locker') zonen.locker = totalZoneLoad > 0 ? load / totalZoneLoad : 0;
    else if (kl === 'grau') zonen.grau = totalZoneLoad > 0 ? load / totalZoneLoad : 0;
    else if (kl === 'hart') zonen.hart = totalZoneLoad > 0 ? load / totalZoneLoad : 0;
  }

  // schlaf/readiness = average (only non-null values)
  let sleepSum = 0, sleepCount = 0;
  let readinessSum = 0, readinessCount = 0;
  for (const tag of t) {
    if (tag.sleep != null) { sleepSum += tag.sleep; sleepCount++; }
    if (tag.readiness != null) { readinessSum += tag.readiness; readinessCount++; }
  }
  const schlaf = sleepCount > 0 ? sleepSum / sleepCount : null;
  const readiness = readinessCount > 0 ? readinessSum / readinessCount : null;

  return { last, wochenLast, ruhetage, laengste, sportMix, zonen, schlaf, readiness };
}

// ─── Export 3: rezeptbuch() ──────────────────────────────────────

function blocksOverlap(b1: Block, b2: Block): boolean {
  return b1.start <= b2.ende && b2.start <= b1.ende;
}

function selectNonOverlapping(blocks: Block[], count: number, descending: boolean): Block[] {
  const sorted = [...blocks].sort((a, b) => descending ? (b.esDelta || 0) - (a.esDelta || 0) : (a.esDelta || 0) - (b.esDelta || 0));
  const result: Block[] = [];
  for (const block of sorted) {
    if (result.length >= count) break;
    let overlaps = false;
    for (const chosen of result) {
      if (blocksOverlap(chosen, block)) { overlaps = true; break; }
    }
    if (!overlaps) result.push(block);
  }
  return result;
}

export function rezeptbuch(t: P4Tag[]): RezeptbuchResult {
  const n = t.length;
  if (n < 28) return { top: [], schwach: [], alle: 0, muster: [] };

  // Create blocks of 28 days starting every 7 days
  const rawBlocks: Block[] = [];
  for (let i = 0; i <= n - 28; i += 7) {
    const slice = t.slice(i, i + 28);
    if (slice.length < 28) continue;

    const start = slice[0].date;
    const ende = slice[slice.length - 1].date;
    const k = kennzahlen(slice);

    // esDelta: es(last) - es(first), only if both present
    const firstTag = slice[0];
    const lastTag = slice[slice.length - 1];
    let esDelta: number | null = null;
    if (firstTag.es != null && lastTag.es != null) {
      esDelta = lastTag.es - firstTag.es;
    }

    let ctlDelta: number | null = null;
    if (firstTag.ctl != null && lastTag.ctl != null) {
      ctlDelta = lastTag.ctl - firstTag.ctl;
    }

    if (esDelta === null) continue; // nur Blöcke mit esDelta != null

    rawBlocks.push({ start, ende, esDelta, ctlDelta, k });
  }

  // Sort by esDelta descending
  rawBlocks.sort((a, b) => (b.esDelta || 0) - (a.esDelta || 0));

  // top = 5 best non-overlapping
  const top = selectNonOverlapping(rawBlocks, 5, true);

  // schwach = 3 worst non-overlapping (sort ascending)
  const schwachBlocks = [...rawBlocks].sort((a, b) => (a.esDelta || 0) - (b.esDelta || 0));
  const schwach: Block[] = [];
  for (const block of schwachBlocks) {
    if (schwach.length >= 3) break;
    let overlaps = false;
    for (const chosen of schwach) {
      if (blocksOverlap(chosen, block)) { overlaps = true; break; }
    }
    if (!overlaps) schwach.push(block);
  }

  // muster: compare mean of top vs mean of all blocks
  const allMean = computeBlockMeans(rawBlocks);
  const topMean = computeBlockMeans(top.length > 0 ? top : [rawBlocks[0]]);

  const muster: string[] = [];

  // --- helper: integer, comma-separator ---
  const fmtInt = (n: number): string => Math.round(n).toLocaleString('de-DE');
  // --- helper: 1 decimal, comma-separator ---
  const fmtOne = (n: number): string => n.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

  // --- scalar checks (≥15 % relative deviation) ---
  interface ScalarEntry { label: string; topVal: number; allVal: number; fmtTop: string; fmtAll: string; unit: string }
  const scalars: ScalarEntry[] = [
    { label: 'Wochenlast', topVal: topMean.wochenLast, allVal: allMean.wochenLast, fmtTop: fmtInt(topMean.wochenLast), fmtAll: fmtInt(allMean.wochenLast), unit: ' ESS' },
    { label: 'Ruhetage',   topVal: topMean.ruhetage,  allVal: allMean.ruhetage,  fmtTop: fmtOne(topMean.ruhetage),  fmtAll: fmtOne(allMean.ruhetage),  unit: '' },
    { label: 'längste Einheit', topVal: topMean.laengste, allVal: allMean.laengste, fmtTop: fmtInt(topMean.laengste), fmtAll: fmtInt(allMean.laengste), unit: ' ESS' },
    { label: 'Schlaf',    topVal: topMean.schlaf ?? 0, allVal: allMean.schlaf ?? 0, fmtTop: fmtOne(topMean.schlaf ?? 0), fmtAll: fmtOne(allMean.schlaf ?? 0), unit: ' h' },
    { label: 'Readiness', topVal: topMean.readiness ?? 0, allVal: allMean.readiness ?? 0, fmtTop: fmtInt(topMean.readiness ?? 0), fmtAll: fmtInt(allMean.readiness ?? 0), unit: '' },
  ];

  for (const s of scalars) {
    if (s.allVal === 0) continue;
    const relDiff = ((s.topVal - s.allVal) / Math.abs(s.allVal)) * 100;
    if (Math.abs(relDiff) < 15) continue;
    const direction = s.topVal > s.allVal ? 'mehr' : 'weniger';
    const allWithUnit = s.unit ? ` ${s.fmtAll}${s.unit}` : s.fmtAll;
    const topWithUnit = s.unit ? ` ${s.fmtTop}${s.unit}` : s.fmtTop;
    muster.push(`Deine besten Blöcke hatten ${direction} ${s.label} (${topWithUnit.trim()} statt ${allWithUnit.trim()}).`);
  }

  // --- zone / sport checks (≥5 pp absolute difference, integer percent) ---
  interface ZoneEntry { label: string; topPct: number; allPct: number }
  const zoneEntries: ZoneEntry[] = [
    { label: 'Grauzone',       topPct: Math.round((topMean.zonen.grau  ?? 0) * 100), allPct: Math.round((allMean.zonen.grau  ?? 0) * 100) },
    { label: 'harte Anteile',  topPct: Math.round((topMean.zonen.hart ?? 0) * 100), allPct: Math.round((allMean.zonen.hart ?? 0) * 100) },
    { label: 'Laufen',          topPct: Math.round((topMean.sportMix['Laufen'] ?? 0) * 100), allPct: Math.round((allMean.sportMix['Laufen'] ?? 0) * 100) },
    { label: 'Rad',             topPct: Math.round((topMean.sportMix['Rad']    ?? 0) * 100), allPct: Math.round((allMean.sportMix['Rad']    ?? 0) * 100) },
  ];

  for (const z of zoneEntries) {
    const diff = z.topPct - z.allPct; // percentage points
    if (Math.abs(diff) < 5) continue;
    const direction = z.topPct > z.allPct ? 'mehr' : 'weniger';
    muster.push(`Deine besten Blöcke hatten ${direction} ${z.label} (${z.topPct} % statt ${z.allPct} %).`);
  }

  return { top, schwach, alle: rawBlocks.length, muster };
}

export interface BlockMeans {
  wochenLast: number;
  ruhetage: number;
  laengste: number;
  schlaf: number | null;
  readiness: number | null;
  zonen: { locker: number; grau: number; hart: number };
  sportMix: Record<string, number>;
  esDeltaAvg: number;
}

function computeBlockMeans(blocks: Block[]): BlockMeans {
  if (blocks.length === 0) return {} as BlockMeans;

  let sumWochenLast = 0, sumRuhetage = 0, sumLaengste = 0;
  let sumSchlaf = 0, countSchlaf = 0;
  let sumReadiness = 0, countReadiness = 0;
  let esDeltaSum = 0;

  const zonenSums = { locker: 0, grau: 0, hart: 0 };
  const sportMixSums = new Map<string, number>();

  for (const b of blocks) {
    const k = b.k;
    sumWochenLast += k.wochenLast;
    sumRuhetage += k.ruhetage;
    sumLaengste += k.laengste;

    if (k.schlaf != null) { sumSchlaf += k.schlaf; countSchlaf++; }
    if (k.readiness != null) { sumReadiness += k.readiness; countReadiness++; }
    if (b.esDelta != null) esDeltaSum += b.esDelta;

    zonenSums.locker += k.zonen.locker;
    zonenSums.grau += k.zonen.grau;
    zonenSums.hart += k.zonen.hart;

    for (const [sport, frac] of Object.entries(k.sportMix)) {
      sportMixSums.set(sport, (sportMixSums.get(sport) || 0) + frac);
    }
  }

  const n = blocks.length;
  return {
    wochenLast: sumWochenLast / n,
    ruhetage: sumRuhetage / n,
    laengste: sumLaengste / n,
    schlaf: countSchlaf > 0 ? sumSchlaf / countSchlaf : null,
    readiness: countReadiness > 0 ? sumReadiness / countReadiness : null,
    esDeltaAvg: esDeltaSum / n,
    zonen: {
      locker: zonenSums.locker / n,
      grau: zonenSums.grau / n,
      hart: zonenSums.hart / n,
    },
    sportMix: Object.fromEntries([...sportMixSums].map(([k, v]) => [k, v / n])) as Record<string, number>,
  };
}

// ─── Export 4: zeitraumVergleich() ──────────────────────────────

export function zeitraumVergleich(t: P4Tag[], a: { von: string; bis: string }, b: { von: string; bis: string }): VergleichsZeile[] {
  const filterRange = (tags: P4Tag[], von: string, bis: string): P4Tag[] => {
    return tags.filter((tag) => tag.date >= von && tag.date <= bis);
  };

  const computeRangeStats = (tags: P4Tag[]) => {
    const n = tags.length;
    if (n === 0) return null;

    let sumLoad = 0, maxLoad = 0;
    let einheiten = 0, ruhTage = 0;
    let sportLoads = new Map<string, number>();
    let zoneSums = { locker: 0, grau: 0, hart: 0 };
    let zoneTotal = 0;
    let sleepSum = 0, sleepCount = 0;
    let hrvSum = 0, hrvCount = 0, rhrSum = 0, rhrCount = 0;
    let esDelta: number | null = null, ctlDelta: number | null = null;
    let kmTotal = 0, hmTotal = 0;

    for (const tag of tags) {
      if (tag.load != null) { sumLoad += tag.load; }
      if (tag.load != null && tag.load > maxLoad) maxLoad = tag.load;
      if (tag.load != null && tag.load > 20) einheiten++;
      if (tag.load == null || tag.load <= 20) ruhTage++;

      if (tag.load != null && tag.load > 0) {
        const grp = sportGruppe(tag.sport);
        sportLoads.set(grp, (sportLoads.get(grp) || 0) + tag.load);
      }

      if (tag.load != null && tag.load > 0) {
        const kl = zonenKlasse(tag.zone);
        if (kl !== 'off') {
          zoneSums[kl] += tag.load;
          zoneTotal += tag.load;
        }
      }

      if (tag.sleep != null) { sleepSum += tag.sleep; sleepCount++; }
      if (tag.hrv != null) { hrvSum += tag.hrv; hrvCount++; }
      if (tag.rhr != null) { rhrSum += tag.rhr; rhrCount++; }
      if (tag.es != null) { /* track for delta */ }

      kmTotal += tag.km || 0;
      hmTotal += tag.hm || 0;

      // Track first/last CTL and ES for delta
    }

    const lastPerTag = n > 0 ? sumLoad / n : 0;
    const wochenLastCalc = lastPerTag * 7;

    // esDelta, ctlDelta from first and last tag
    const firstTag = tags[0];
    const lastTag = tags[n - 1];
    let esDeltaCalc: number | null = null;
    if (firstTag.es != null && lastTag.es != null) esDeltaCalc = lastTag.es - firstTag.es;
    let ctlDeltaCalc: number | null = null;
    if (firstTag.ctl != null && lastTag.ctl != null) ctlDeltaCalc = lastTag.ctl - firstTag.ctl;

    const anteileGrau = zoneTotal > 0 ? zoneSums.grau / zoneTotal : 0;
    const anteileHart = zoneTotal > 0 ? zoneSums.hart / zoneTotal : 0;

    const øSchlaf = sleepCount > 0 ? sleepSum / sleepCount : null;
    const øHRV = hrvCount > 0 ? hrvSum / hrvCount : null;
    const øRHR = rhrCount > 0 ? rhrSum / rhrCount : null;

    return {
      lastPerTag, wochenLast: wochenLastCalc, einheiten, ruhTage, maxLoad,
      anteileGrau, anteileHart, øSchlaf, øHRV, øRHR,
      esDelta: esDeltaCalc, ctlDelta: ctlDeltaCalc, kmTotal, hmTotal
    };
  };

  const statsA = computeRangeStats(filterRange(t, a.von, a.bis));
  const statsB = computeRangeStats(filterRange(t, b.von, b.bis));

  if (!statsA || !statsB) return [];

  const rows: VergleichsZeile[] = [
    { key: 'lastPerTag', label: 'Ø Last/Tag', a: statsA.lastPerTag, b: statsB.lastPerTag, einheit: 'ESS', besserWennHoeher: null },
    { key: 'wochenLast', label: 'Wochenlast', a: statsA.wochenLast, b: statsB.wochenLast, einheit: 'ESS', besserWennHoeher: null },
    { key: 'einheiten', label: 'Einheiten (Tage load>20)', a: statsA.einheiten, b: statsB.einheiten, einheit: 'Tage', besserWennHoeher: true },
    { key: 'ruhetage', label: 'Ruhetage (load≤20/null)', a: statsA.ruhTage, b: statsB.ruhTage, einheit: 'Tage', besserWennHoeher: true },
    { key: 'laengste', label: 'Längste Einheit', a: statsA.maxLoad, b: statsB.maxLoad, einheit: 'ESS', besserWennHoeher: null },
    { key: 'anteilGrau', label: 'Anteil Grauzone', a: statsA.anteileGrau == null ? null : Math.round(statsA.anteileGrau * 100), b: statsB.anteileGrau == null ? null : Math.round(statsB.anteileGrau * 100), einheit: '%', besserWennHoeher: null },
    { key: 'anteilHart', label: 'Anteil hart', a: statsA.anteileHart == null ? null : Math.round(statsA.anteileHart * 100), b: statsB.anteileHart == null ? null : Math.round(statsB.anteileHart * 100), einheit: '%', besserWennHoeher: null },
    { key: 'schlaf', label: 'Ø Schlaf', a: statsA.øSchlaf, b: statsB.øSchlaf, einheit: 'h', besserWennHoeher: true },
    { key: 'readinessAvg', label: 'Ø Readiness', a: (() => {
      const tagsAB = filterRange(t, a.von, a.bis);
      const tagsBB = filterRange(t, b.von, b.bis);
      let sA = 0, cA = 0;
      for (const t of tagsAB) { if (t.readiness != null) sA += t.readiness; cA++; }
      let sB = 0, cB = 0;
      for (const t of tagsBB) { if (t.readiness != null) sB += t.readiness; cB++; }
      return { a: cA > 0 ? sA / cA : null, b: cB > 0 ? sB / cB : null };
    })().a, b: (() => {
      const tagsAB = filterRange(t, a.von, a.bis);
      const tagsBB = filterRange(t, b.von, b.bis);
      let sA = 0, cA = 0;
      for (const t of tagsAB) { if (t.readiness != null) sA += t.readiness; cA++; }
      let sB = 0, cB = 0;
      for (const t of tagsBB) { if (t.readiness != null) sB += t.readiness; cB++; }
      return { a: cA > 0 ? sA / cA : null, b: cB > 0 ? sB / cB : null };
    })().b, einheit: '', besserWennHoeher: true },
    { key: 'hrvAvg', label: 'Ø HRV', a: statsA.øHRV, b: statsB.øHRV, einheit: 'ms', besserWennHoeher: true },
    { key: 'rhrAvg', label: 'Ø Ruhepuls', a: statsA.øRHR, b: statsB.øRHR, einheit: 'bpm', besserWennHoeher: false },
    { key: 'ctlDelta', label: 'CTL Änderung', a: statsA.ctlDelta, b: statsB.ctlDelta, einheit: '', besserWennHoeher: null },
    { key: 'esDelta', label: 'Endurance Score Änderung', a: statsA.esDelta, b: statsB.esDelta, einheit: '', besserWennHoeher: null },
    { key: 'kmTotal', label: 'km gesamt', a: statsA.kmTotal, b: statsB.kmTotal, einheit: 'km', besserWennHoeher: true },
    { key: 'hmTotal', label: 'Höhenmeter gesamt', a: statsA.hmTotal, b: statsB.hmTotal, einheit: 'm', besserWennHoeher: true },
  ];

  return rows;
}

// ─── Export 5: jahresRueckblick() ────────────────────────────────

export function jahresRueckblick(t: P4Tag[], jahr: number | 'letzte365'): Rueckblick {
  let filtered: P4Tag[];

  if (jahr === 'letzte365') {
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 365);
    const cutoffStr = cutoff.toISOString().split('T')[0];
    filtered = t.filter((tag) => tag.date >= cutoffStr);
  } else {
    const jahrStart = `${jahr}-01-01`;
    const jahrEnde = `${jahr}-12-31`;
    filtered = t.filter((tag) => tag.date >= jahrStart && tag.date <= jahrEnde);
  }

  if (filtered.length === 0) {
    return {
      titel: typeof jahr === 'number' ? `Kira Wrapped ${jahr}` : 'Kira Wrapped – letzte 365 Tage',
      tage: 0, summeLast: 0, einheiten: 0, km: 0, hm: 0,
      topSport: { name: '', anteil: 0 }, besterTag: { date: '', load: 0 },
      laengsteSerie: { tage: 0, bis: '' },
      besterSchlafMonat: null, besterReadinessMonat: null,
      esMax: null, esStart: null, esEnde: null, ctlStart: null, ctlEnde: null,
      besterBlock: null, ruhigsterMonat: null,
    };
  }

  // titel
  const titel = typeof jahr === 'number' ? `Kira Wrapped ${jahr}` : 'Kira Wrapped – letzte 365 Tage';

  // SummeLast, Einheiten, km, hm
  let summeLast = 0, einheiten = 0;
  let kmTotal = 0, hmTotal = 0;

  for (const tag of filtered) {
    if (tag.load != null) summeLast += tag.load;
    if (tag.load != null && tag.load > 20) einheiten++;
    kmTotal += tag.km || 0;
    hmTotal += tag.hm || 0;
  }

  // topSport: sportGruppe with highest share of load > 0
  const sportLoads = new Map<string, number>();
  let totalLoadForTopSport = 0;
  for (const tag of filtered) {
    if (tag.load != null && tag.load > 0) {
      const grp = sportGruppe(tag.sport);
      sportLoads.set(grp, (sportLoads.get(grp) || 0) + tag.load);
      totalLoadForTopSport += tag.load;
    }
  }
  let topSportName = '', topSportAnteil = 0;
  for (const [name, load] of sportLoads) {
    if (load > topSportAnteil * totalLoadForTopSport) { // fix: compare load values
      const anteil = totalLoadForTopSport > 0 ? load / totalLoadForTopSport : 0;
      if (anteil > topSportAnteil) {
        topSportName = name;
        topSportAnteil = anteil;
      }
    }
  }
  // Recompute correctly
  const topSportEntries = Array.from(sportLoads.entries()).sort((a, b) => b[1] - a[1]);
  const bestSport = topSportEntries.length > 0 ? {
    name: topSportEntries[0][0],
    anteil: totalLoadForTopSport > 0 ? topSportEntries[0][1] / totalLoadForTopSport : 0
  } : { name: '', anteil: 0 };

  // besterTag
  let bestLoad = -1, bestDate = '';
  for (const tag of filtered) {
    if (tag.load != null && tag.load > bestLoad) {
      bestLoad = tag.load;
      bestDate = tag.date;
    }
  }

  // laengsteSerie: longest consecutive run of days with load > 20
  let serieTage = 0, maxSerie = 0;
  let besteSerieBis = '';
  for (let i = 0; i < filtered.length; i++) {
    const tag = filtered[i];
    if (tag.load != null && tag.load > 20) {
      serieTage++;
    } else {
      if (serieTage > maxSerie) { maxSerie = serieTage; besteSerieBis = filtered[i - 1]?.date || bestDate; }
      serieTage = 0;
    }
  }
  // check last run
  if (serieTage > maxSerie) { maxSerie = serieTage; besteSerieBis = filtered[filtered.length - 1].date; }

  // monthly stats
  const months = new Map<string, { schlafSum: number; schlafCount: number; readinessSum: number; readinessCount: number; last: number; daysInMonth: Set<string> }>();

  for (const tag of filtered) {
    const monat = tag.date.substring(0, 7); // YYYY-MM
    if (!months.has(monat)) months.set(monat, { schlafSum: 0, schlafCount: 0, readinessSum: 0, readinessCount: 0, last: 0, daysInMonth: new Set() });
    const m = months.get(monat)!;
    m.daysInMonth.add(tag.date);
    if (tag.sleep != null) { m.schlafSum += tag.sleep; m.schlafCount++; }
    if (tag.readiness != null) { m.readinessSum += tag.readiness; m.readinessCount++; }
    if (tag.load != null) m.last += tag.load;
  }

  // besterSchlafMonat: only months with >= 15 days data
  let bestSchlafMonat: { monat: string; schlaf: number } | null = null;
  let bestSchlafVal = -1;
  for (const [monat, data] of months) {
    if (data.daysInMonth.size >= 15 && data.schlafCount > 0) {
      const avg = data.schlafSum / data.schlafCount;
      if (avg > bestSchlafVal) { bestSchlafVal = avg; bestSchlafMonat = { monat, schlaf: avg }; }
    }
  }

  let bestReadinessMonat: { monat: string; readiness: number } | null = null;
  let bestReadinessVal = -1;
  for (const [monat, data] of months) {
    if (data.daysInMonth.size >= 15 && data.readinessCount > 0) {
      const avg = data.readinessSum / data.readinessCount;
      if (avg > bestReadinessVal) { bestReadinessVal = avg; bestReadinessMonat = { monat, readiness: avg }; }
    }
  }

  // esMax, esStart/Ende, ctlStart/Ende (first and last date in range)
  let esMaxDate = '', esMaxVal = -1;
  for (const tag of filtered) {
    if (tag.es != null && tag.es > esMaxVal) {
      esMaxVal = tag.es;
      esMaxDate = tag.date;
    }
  }

  const sortedFiltered = [...filtered].sort((a, b) => a.date.localeCompare(b.date));
  const esStart = sortedFiltered[0]?.es ?? null;
  const esEnde = sortedFiltered[sortedFiltered.length - 1]?.es ?? null;
  const ctlStart = sortedFiltered[0]?.ctl ?? null;
  const ctlEnde = sortedFiltered[sortedFiltered.length - 1]?.ctl ?? null;

  // ruhigsterMonat: lowest monthly last, only months with >= 15 days
  let ruhigster: { monat: string; last: number } | null = null;
  for (const [monat, data] of months) {
    if (data.daysInMonth.size >= 15) {
      if (ruhigster == null || data.last < ruhigster.last) {
        ruhigster = { monat, last: data.last };
      }
    }
  }

  // besterBlock = rezeptbuch(filtered).top[0]
  const rb = rezeptbuch(sortedFiltered);
  const besterBlock = rb.top.length > 0 ? rb.top[0] : null;

  return {
    titel,
    tage: filtered.length,
    summeLast,
    einheiten,
    km: kmTotal,
    hm: hmTotal,
    topSport: bestSport,
    besterTag: { date: bestDate, load: bestLoad },
    laengsteSerie: { tage: maxSerie, bis: besteSerieBis },
    besterSchlafMonat: bestSchlafMonat,
    besterReadinessMonat: bestReadinessMonat,
    esMax: esMaxVal > 0 ? { date: esMaxDate, es: esMaxVal } : null,
    esStart, esEnde, ctlStart, ctlEnde,
    besterBlock, ruhigsterMonat: ruhigster,
  };
}

// ─── Export 6: tourform() ────────────────────────────────────────

export function tourform(t: P4Tag[], z: Ziel): Tourform {
  // Sort by date
  const sorted = [...t].sort((a, b) => a.date.localeCompare(b.date));

  // Get last 8 weeks of data (56 days)
  const last8 = sorted.slice(-56);

  // wochenLastIst: average weekly load of last 8 weeks
  let weeksCount = 0;
  if (last8.length >= 7) {
    const dayCount = last8.length;
    weeksCount = Math.max(1, dayCount / 7); // at least 1 week
  }

  let wochenLastIst = 0;
  if (weeksCount > 0) {
    let weekSum = 0;
    for (const tag of last8) {
      if (tag.load != null) weekSum += tag.load;
    }
    wochenLastIst = weekSum / weeksCount; // avg weekly load
  }

  // tl = tourLast(z) using planning.tourLast
  const tl = tpl(z, undefined);

  // wochenLastZiel = 2 * tl
  const wochenLastZiel = 2 * tl;

  // volumen = min(1, ø wochenLast der letzten 8 Wochen / (2 * tl))
  const volumen = Math.min(1, wochenLastIst / (2 * tl));

  // langeTage: Tage der letzten 10 Wochen (70 days) mit load >= 0.6 * tl
  const last10 = sorted.slice(-70);
  const langeZielWert = z.prioritaet === 'A' ? 6 : 4;
  let langeTage = 0;
  const threshold = 0.6 * tl;

  for (const tag of last10) {
    if (tag.load != null && tag.load >= threshold) {
      langeTage++;
    }
  }

  // länge = min(1, langeTage / langeZiel)
  const lange = Math.min(1, langeTage / langeZielWert);

  // maxHm: höchster hm-Wert der letzten 10 Wochen
  let maxHm: number | null = null;
  for (const tag of last10) {
    if (tag.hm != null && tag.hm > 0) {
      if (maxHm == null || tag.hm > maxHm) {
        maxHm = tag.hm;
      }
    }
  }

  // hmZiel = 0.7 * z.hm
  const hmZiel = 0.7 * z.hm;

  // hoehe: null if no hm data, else min(1, maxHm / hmZiel)
  let hoehe: number | null = null;
  if (maxHm != null) {
    hoehe = hmZiel > 0 ? Math.min(1, maxHm / hmZiel) : null;
  }

  // prozent = round(100 * (hoehe == null ? 0.6*volumen + 0.4*lange : 0.5*volumen + 0.35*lange + 0.15*hoehe))
  let prozent = 0;
  if (hoehe == null) {
    prozent = Math.round(100 * (0.6 * volumen + 0.4 * lange));
  } else {
    prozent = Math.round(100 * (0.5 * volumen + 0.35 * lange + 0.15 * hoehe));
  }

  // hinweise: deutsche Sätze für schwächste Komponente
  const hinweise: string[] = [];

  if (hoehe == null) {
    // volume and long comparison
    const volScore = 0.6 * volumen;
    const langScore = 0.4 * lange;

    if (volumen < 0.7) {
      hinweise.push(`Wochenlast auf etwa ${Math.round(wochenLastZiel)} ESS steigern (aktuell ca. ${Math.round(wochenLastIst)}).`);
    }
    if (lange < 0.7) {
      const needed = Math.max(0, langeZielWert - langeTage);
      hinweise.push(`Noch ${needed} lange Tage ab ${Math.round(0.6 * tl)} ESS in den nächsten Wochen.`);
    }
    hinweise.push('Höhenmeter werden kaum erfasst – Wert nicht berücksichtigt.');
  } else {
    const volScore = 0.5 * volumen;
    const langScore = 0.35 * lange;
    const hoeheScore = 0.15 * hoehe;

    // Find weakest component
    const comps: { name: string; score: number; value: number }[] = [
      { name: 'Volumen', score: volScore, value: volumen },
      { name: 'Lange Tage', score: langScore, value: lange },
      { name: 'Höhe', score: hoeheScore, value: hoehe },
    ];

    comps.sort((a, b) => a.score - b.score);
    const weakest = comps[0];

    if (weakest.name === 'Volumen') {
      hinweise.push(`Wochenlast auf etwa ${Math.round(wochenLastZiel)} ESS steigern (aktuell ca. ${Math.round(wochenLastIst)}).`);
    } else if (weakest.name === 'Lange Tage') {
      const needed = Math.max(0, langeZielWert - langeTage);
      hinweise.push(`Noch ${needed} lange Tage ab ${Math.round(0.6 * tl)} ESS in den nächsten Wochen.`);
    } else if (weakest.name === 'Höhe') {
      hinweise.push(`Höhenmeter: zuletzt höchstens ${maxHm} Hm an einem Tag – Ziel etwa ${Math.round(hmZiel)}.`);
    }
  }

  return {
    prozent, volumen, lange, hoehe, tourLast: tl, wochenLastIst, wochenLastZiel,
    langeTage, langeZiel: langeZielWert, maxHm, hmZiel, hinweise,
  };
}

// ─── Format-Hilfe für deutsche Zahlen ────────────────────────────

function formatNumber(n: number): string {
  if (n == null) return '—';
  // Use German locale: thousands separator ., decimal comma
  return n.toLocaleString('de-DE', { maximumFractionDigits: 1 });
}
