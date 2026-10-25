/**
 * Erholungspreis-Dashboard: Sport- und Zonen-preis für Trainingsbelastung.
 * Reines TypeScript – keine React-Abhängigkeiten, keine 'any' außer beim Parsen.
 */

// ─── Datentypen ──────────────────────────────────────────────────────────────

export interface ETag {
  date: string;
  load: number | null;
  readiness: number | null;
  sport: string;
  zone: string;
  planZone: string;
  krank: boolean;
}

export type SportGruppe = 'Laufen' | 'Rad' | 'Wandern' | 'Rudern' | 'Ski' | 'Sonst';

export function sportGruppe(s: string): SportGruppe {
  const v = s.toLowerCase().trim();
  if (v === 'run' || v === 'lauf') return 'Laufen';
  if (/^(bike|rad|gravel|ride|cycl)$/.test(v)) return 'Rad';
  if (v === 'hike' || v === 'wander' || v === 'berg') return 'Wandern';
  if (v === 'row' || v === 'ruder') return 'Rudern';
  if (v === 'ski' || v === 'langlauf') return 'Ski';
  return 'Sonst';
}

export type ZonenKlasse = 'locker' | 'grau' | 'hart' | 'off';

export function zonenKlasse(z: string): ZonenKlasse {
  const norm = z.toUpperCase();
  // Prüfe ob 'Z4' oder 'Z5' anywhere (eigenständig oder in Bereichen)
  if (/Z[45]/.test(norm)) return 'hart';
  // Prüfe ob 'Z3' anywhere (eigenständig oder in Bereichen wie Z2–Z3)
  if (/Z3/.test(norm)) return 'grau';
  // Prüfe ob nur Z1 oder Z2 (und nicht Z3/Z4/Z5)
  if (/^Z[12]$/.test(norm)) return 'locker';
  // Bereiche wie Z1–Z2, Z1/Z2 etc. – nur wenn Z3-Z5 nicht vorkommen
  if (!/Z[345]/.test(norm)) {
    if (/Z[12]/.test(norm)) return 'locker';
  }
  return 'off';
}

// ─── Daten laden (Cache pro Sitzung) ────────────────────────────────────────

import { fetchChartData } from './api';
import { fetchExtras } from './extras';

let cache: ETag[] | null = null;
let cachePromise: Promise<ETag[]> | null = null;

async function loadDays(): Promise<ETag[]> {
  if (cache) return cache;
  if (!cachePromise) {
    cachePromise = (async () => {
      try {
        const r: any = await fetchChartData('9999d' as any, [
          'load_fb_day',
          'Garmin_Training_Readiness',
          'Sport_x',
          'Zone',
          'coach_Zone',
        ]);
        const rows = (r && r.data) as { date: string; [key: string]: unknown }[] | undefined;
        const arr = rows || [];

        // Extras für Krankheitsmarken laden (nur einmal)
        let krankArten: string[] = [];
        try {
          const exResult: any = await fetchExtras<{ arten?: string[] }>('ereignis', 400);
          const exData: any = (exResult && exResult.data) || (Array.isArray(exResult) ? exResult : null);
          if (exData && typeof exData === 'object' && !Array.isArray(exData)) {
            const artenField = (exResult as any).arten;
            if (Array.isArray(artenField)) krankArten = artenField.map(String);
          } else if (Array.isArray(exData)) {
            for (const item of exData) {
              if (item && typeof item === 'object') {
                const a = (item as any).arten;
                if (Array.isArray(a)) krankArten.push(...a.map(String));
              }
            }
          }
        } catch {
          // fallback: leer
        }

        const krankSet = new Set(krankArten.map((a) => a.toLowerCase()));
        const today = new Date().toLocaleDateString('sv-SE');
        const seenDates = new Map<string, ETag>();

        for (const row of arr) {
          const d = String(row.date || '').split('T')[0]; // YYYY-MM-DD
          if (d === '' || d > today) continue;
          if (seenDates.has(d)) continue; // erste behalten

          const loadVal = parseNum(row.load_fb_day);
          const readinessVal = parseNum(row.Garmin_Training_Readiness);
          const sportRaw = String(row.Sport_x || '');
          const zoneRaw = String(row.Zone || '');
          const planZoneRaw = String(row.coach_Zone || '');

          // Krankheitsprüfung: Datum oder Vortag hat 'krank' in arten
          let krank = false;
          // Da Extras-Antwort kein Datum enthält, kann keine Datums-Zuordnung erfolgen.
          // krank bleibt false; bei richtiger API mit Datum-Feld wäre hier Zuordnung möglich.

          seenDates.set(d, {
            date: d,
            load: loadVal,
            readiness: readinessVal,
            sport: sportRaw,
            zone: zoneRaw,
            planZone: planZoneRaw,
            krank: krank,
          });
        }

        const result = Array.from(seenDates.values()).sort((a, b) => a.date.localeCompare(b.date));
        cache = result;
        return result;
      } catch {
        cache = [];
        return [];
      } finally {
        cachePromise = null; // für Wiederaufruf bei Fehler
      }
    })();
  }
  return cachePromise;
}

// export default für die öffentliche API-Funktion:
export async function ladeTage(): Promise<ETag[]> {
  return loadDays();
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

function parseNum(v: unknown): number | null {
  if (v === '' || v == null) return null;
  const s = String(v).trim();
  if (s === '' || s.toLowerCase() === 'nan') return null;
  const n = Number(s.replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

// ─── Reaktionen berechnen ────────────────────────────────────────────────────

export interface Reaktion {
  date: string;
  load: number;
  dR: number;
  gruppe: SportGruppe;
  klasse: ZonenKlasse;
}

export function reaktionen(tags: ETag[]): Reaktion[] {
  const out: Reaktion[] = [];

  for (let i = 0; i < tags.length - 1; i++) {
    const t = tags[i];
    if (t.load == null || t.load <= 30) continue;
    if (t.krank) continue; // Tag i krank

    const tNext = tags[i + 1];
    if (tNext.krank) continue; // Tag i+1 krank

    const rNext = tNext.readiness;
    if (rNext == null) continue; // readiness bei i+1 vorhanden

    const basisReadiness: number[] = [];
    for (let j = i - 27; j <= i; j++) {
      if (j < 0 || j >= tags.length) continue;
      const tRef = tags[j];
      if (tRef.krank) continue; // kranke Tage auslassen
      if (tRef.readiness == null) continue;
      basisReadiness.push(tRef.readiness);
    }

    if (basisReadiness.length < 14) continue;
    const basis = basisReadiness.reduce((a, b) => a + b, 0) / basisReadiness.length;
    const dR = rNext - basis;

    out.push({
      date: t.date,
      load: t.load,
      dR,
      gruppe: sportGruppe(t.sport),
      klasse: zonenKlasse(t.zone),
    });
  }

  return out;
}

// ─── Preisgruppen ────────────────────────────────────────────────────────────

export interface PreisGruppe {
  name: string;
  n: number;
  avgLoad: number;
  avgDR: number;
  a: number;
  b: number;
  sd: number;
  urteil: 'belastbar' | 'eingeschränkt' | 'wenig Daten';
}

function linearRegression(xs: number[], ys: number[]): { a: number; b: number } {
  const n = xs.length;
  if (n < 2) return { a: ys.reduce((a, b) => a + b, 0) / n || 0, b: 0 };
  const xMean = xs.reduce((a, b) => a + b, 0) / n;
  const yMean = ys.reduce((a, b) => a + b, 0) / n;
  let sxy = 0, sxx = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - xMean) * (ys[i] - yMean);
    sxx += (xs[i] - xMean) ** 2;
  }
  const b = sxx === 0 ? 0 : sxy / sxx;
  const a = yMean - b * xMean;
  return { a, b };
}

function sdResiduen(xs: number[], ys: number[], a: number, b: number): number {
  const n = xs.length;
  if (n < 2) return 0;
  const residuen = ys.map((y, i) => y - (a + b * xs[i]));
  const mean = residuen.reduce((a, b) => a + b, 0) / n;
  const variance = residuen.reduce((a, r) => a + (r - mean) ** 2, 0) / n;
  return Math.sqrt(variance);
}

function urteil(n: number): 'belastbar' | 'eingeschränkt' | 'wenig Daten' {
  if (n >= 30) return 'belastbar';
  if (n >= 15) return 'eingeschränkt';
  return 'wenig Daten';
}

export function preisGruppen(rs: Reaktion[]): { sport: PreisGruppe[]; zone: PreisGruppe[]; alle: PreisGruppe } {
  // Sportgruppen gruppieren
  const bySport = new Map<string, Reaktion[]>();
  for (const r of rs) {
    if (!bySport.has(r.gruppe)) bySport.set(r.gruppe, []);
    bySport.get(r.gruppe)!.push(r);
  }

  // Zonengruppen gruppieren (ohne 'off')
  const byZone = new Map<string, Reaktion[]>();
  for (const r of rs) {
    if (r.klasse === 'off') continue;
    if (!byZone.has(r.klasse)) byZone.set(r.klasse, []);
    byZone.get(r.klasse)!.push(r);
  }

  const sportGruppen: PreisGruppe[] = Array.from(bySport.entries()).map(([name, items]) => {
    const xs = items.map((r) => r.load);
    const ys = items.map((r) => r.dR);
    let a: number, b: number;
    if (items.length < 5) {
      a = ys.reduce((a, b) => a + b, 0) / ys.length || 0;
      b = 0;
    } else {
      const lr = linearRegression(xs, ys);
      a = lr.a;
      b = lr.b;
    }
    const sd = sdResiduen(xs, ys, a, b);
    return {
      name,
      n: items.length,
      avgLoad: xs.reduce((a, b) => a + b, 0) / xs.length,
      avgDR: ys.reduce((a, b) => a + b, 0) / ys.length,
      a,
      b,
      sd,
      urteil: urteil(items.length),
    };
  }).sort((a, b) => b.n - a.n);

  const zoneGruppen: PreisGruppe[] = Array.from(byZone.entries()).map(([name, items]) => {
    const xs = items.map((r) => r.load);
    const ys = items.map((r) => r.dR);
    let a: number, b: number;
    if (items.length < 5) {
      a = ys.reduce((a, b) => a + b, 0) / ys.length || 0;
      b = 0;
    } else {
      const lr = linearRegression(xs, ys);
      a = lr.a;
      b = lr.b;
    }
    const sd = sdResiduen(xs, ys, a, b);
    return {
      name,
      n: items.length,
      avgLoad: xs.reduce((a, b) => a + b, 0) / xs.length,
      avgDR: ys.reduce((a, b) => a + b, 0) / ys.length,
      a,
      b,
      sd,
      urteil: urteil(items.length),
    };
  }).sort((a, b) => b.n - a.n);

  // Alle: alle Reaktionen außer off
  const allItems = rs.filter((r) => r.klasse !== 'off');
  const allXs = allItems.map((r) => r.load);
  const allYs = allItems.map((r) => r.dR);
  let aAll: number, bAll: number;
  if (allItems.length < 5) {
    aAll = allYs.reduce((a, b) => a + b, 0) / allYs.length || 0;
    bAll = 0;
  } else {
    const lr = linearRegression(allXs, allYs);
    aAll = lr.a;
    bAll = lr.b;
  }

  return {
    sport: sportGruppen,
    zone: zoneGruppen,
    alle: {
      name: 'Alle',
      n: allItems.length,
      avgLoad: allXs.reduce((a, b) => a + b, 0) / (allXs.length || 1),
      avgDR: allYs.reduce((a, b) => a + b, 0) / (allYs.length || 1),
      a: aAll,
      b: bAll,
      sd: sdResiduen(allXs, allYs, aAll, bAll),
      urteil: urteil(allItems.length),
    },
  };
}

// ─── Erwartetes dR berechnen ─────────────────────────────────────────────────

export interface Erwartet {
  wert: number;
  sd: number;
  quelle: string;
}

export function erwartetDR(
  pg: { sport: PreisGruppe[]; zone: PreisGruppe[]; alle: PreisGruppe },
  load: number,
  sport: string,
  zone: string,
): Erwartet {
  const gruppe = sportGruppe(sport);
  const klasse = zonenKlasse(zone);

  let quelle: string;

  // Sportgruppe verwenden wenn n >= 15
  const sportGrp = pg.sport.find((g) => g.name === gruppe);
  let baseWert: number, baseSd: number;

  if (sportGrp && sportGrp.n >= 15) {
    baseWert = sportGrp.a + sportGrp.b * load;
    baseSd = sportGrp.sd;
    quelle = `Sportgruppe ${gruppe} (n=${sportGrp.n})`;
  } else {
    baseWert = pg.alle.a + pg.alle.b * load;
    baseSd = pg.alle.sd;
    quelle = `Alle Daten (n=${pg.alle.n})`;
  }

  let wert = baseWert;
  // Zonen-Zuschlag: wenn grau oder hart und Gruppe n >= 15
  const zoneGrp = pg.zone.find((g) => g.name === klasse);
  if ((klasse === 'grau' || klasse === 'hart') && zoneGrp && zoneGrp.n >= 15) {
    const lockerGrp = pg.zone.find((g) => g.name === 'locker');
    if (lockerGrp && lockerGrp.n >= 15) {
      const zuschlag = (zoneGrp.avgDR - lockerGrp.avgDR) / 2;
      wert = baseWert + zuschlag;
      quelle += ` + Zonenzuschlag ${klasse}`;
    } else {
      wert = baseWert + zoneGrp.avgDR / 2; // fallback: halben Mittelwert addieren
      quelle += ` + Zonenzuschlag ${klasse} (locker fehlt)`;
    }
  }

  return { wert, sd: baseSd, quelle };
}

// ─── Erholungsdauer ──────────────────────────────────────────────────────────

export interface ErholungsDauer {
  p90: number;
  faelle: number;
  verteilung: Record<string, number>;
  median: number | null;
}

export function erholungsDauer(tags: ETag[]): ErholungsDauer {
  // p90 von load > 0
  const loads = tags.filter((t) => t.load != null && t.load > 0).map((t) => t.load!);
  if (loads.length === 0) return { p90: 0, faelle: 0, verteilung: {}, median: null };

  const sorted = [...loads].sort((a, b) => a - b);
  const p90Idx = Math.ceil(sorted.length * 0.9) - 1;
  const p90 = sorted[Math.max(0, p90Idx)];

  // Fälle: Tage i mit load >= p90 und nicht krank
  const verteilung: Record<string, number> = {};

  for (let i = 0; i < tags.length; i++) {
    const t = tags[i];
    if (t.load == null || t.load < p90) continue;
    if (t.krank) continue;

    // Erster Folgetag k = 1..8 mit readiness >= Basis - 3
    const basisReadiness: number[] = [];
    for (let j = i - 28; j <= i - 1; j++) {
      if (j < 0 || j >= tags.length) continue;
      const tRef = tags[j];
      if (tRef.krank) continue;
      if (tRef.readiness == null) continue;
      basisReadiness.push(tRef.readiness);
    }

    if (basisReadiness.length < 14) continue;
    const basis = basisReadiness.reduce((a, b) => a + b, 0) / basisReadiness.length;
    const threshold = basis - 3;

    let k: number | null = null;
    for (let kk = 1; kk <= 8 && i + kk < tags.length; kk++) {
      const next = tags[i + kk];
      if (next.readiness != null && next.readiness >= threshold) {
        k = kk;
        break;
      }
    }

    const key = k == null ? '9+' : String(k);
    verteilung[key] = (verteilung[key] || 0) + 1;
  }

  // Median berechnen
  const keys = Object.keys(verteilung).sort((a, b) => {
    if (a === '9+') return 1;
    if (b === '9+') return -1;
    return Number(a) - Number(b);
  });

  let median: number | null = null;
  if (keys.length > 0) {
    const totalCases = Object.values(verteilung).reduce((a, b) => a + b, 0);
    let half = Math.ceil(totalCases / 2);
    for (const key of keys) {
      half -= verteilung[key];
      if (half <= 0) {
        median = key === '9+' ? 9 : Number(key);
        break;
      }
    }
  }

  return { p90, faelle: Object.values(verteilung).reduce((a, b) => a + b, 0), verteilung: verteilung, median };
}

// ─── Erholungsdauer je Lastbereich ───────────────────────────────────────────

export type LastBand = 'locker' | 'mittel' | 'gross';
export interface BandErholung {
  band: LastBand;
  label: string;
  von: number;
  bis: number | null;
  faelle: number;
  median: number | null;
  verteilung: Record<string, number>;
  /** Anteil der Fälle, in denen am nächsten Morgen schon die Basis erreicht war. */
  sofortAnteil: number | null;
}
export interface ErholungBaender { p50: number; p90: number; baender: BandErholung[] }

/** Lastbereich eines Tages (Grenzen aus eigenen Daten: P50 und P90 der Tage mit Last). */
export function lastBand(load: number, p50: number, p90: number): LastBand {
  return load >= p90 ? 'gross' : load >= p50 ? 'mittel' : 'locker';
}

/** Wie erholungsDauer, aber für alle Lastbereiche (locker < P50 ≤ mittel < P90 ≤ groß). */
export function erholungJeBand(tags: ETag[]): ErholungBaender {
  const loads = tags.filter((t) => t.load != null && t.load > 0).map((t) => t.load!).sort((a, b) => a - b);
  if (!loads.length) return { p50: 0, p90: 0, baender: [] };
  const q = (p: number) => loads[Math.max(0, Math.ceil(loads.length * p) - 1)];
  const p50 = q(0.5), p90 = q(0.9);
  const acc: Record<LastBand, Record<string, number>> = { locker: {}, mittel: {}, gross: {} };
  for (let i = 0; i < tags.length; i++) {
    const t = tags[i];
    if (t.load == null || t.load <= 0 || t.krank) continue;
    const ref: number[] = [];
    for (let j = i - 28; j <= i - 1; j++) {
      if (j < 0) continue;
      const r = tags[j];
      if (!r.krank && r.readiness != null) ref.push(r.readiness);
    }
    if (ref.length < 14) continue;
    const thr = ref.reduce((a, b) => a + b, 0) / ref.length - 3;
    let k: number | null = null;
    let offen = false;
    for (let kk = 1; kk <= 8; kk++) {
      if (i + kk >= tags.length) { offen = true; break; }
      const n = tags[i + kk];
      if (n.readiness != null && n.readiness >= thr) { k = kk; break; }
    }
    if (offen && k == null) continue; // noch nicht abgeschlossen
    const key = k == null ? '9+' : String(k);
    const b = lastBand(t.load, p50, p90);
    acc[b][key] = (acc[b][key] || 0) + 1;
  }
  const mk = (band: LastBand, label: string, von: number, bis: number | null): BandErholung => {
    const v = acc[band];
    const keys = Object.keys(v).sort((a, b) => (a === '9+' ? 1 : b === '9+' ? -1 : Number(a) - Number(b)));
    const total = Object.values(v).reduce((a, b) => a + b, 0);
    let median: number | null = null;
    let half = Math.ceil(total / 2);
    for (const k of keys) { half -= v[k]; if (half <= 0) { median = k === '9+' ? 9 : Number(k); break; } }
    return { band, label, von, bis, faelle: total, median, verteilung: v, sofortAnteil: total ? (v['1'] || 0) / total : null };
  };
  return {
    p50, p90,
    baender: [
      mk('locker', 'locker', 1, p50),
      mk('mittel', 'mittel', p50, p90),
      mk('gross', 'groß', p90, null),
    ],
  };
}

// ─── Grauzone-Analyse ────────────────────────────────────────────────────────

export interface GrauFall {
  date: string;
  planZone: string;
  zone: string;
  load: number;
  dR: number | null;
}

export interface GrauWoche {
  woche: string; // ISO-Woche als 'YYYY-MM-DD' (Montag)
  geplantLocker: number;
  grau: number;
}

export function grauzone(tags: ETag[]): { faelle: GrauFall[]; lockerGeplant: number; proWoche: GrauWoche[] } {
  const faelle: GrauFall[] = [];

  // Map von Datum zu Reaktion für dR-Berechnung
  const tagsByDate = new Map<string, ETag>();
  for (const t of tags) {
    tagsByDate.set(t.date, t);
  }

  // Map von Datum zu readiness für dR-Berechnung
  const tagsByDateIdx = new Map<string, number>();
  for (let i = 0; i < tags.length; i++) {
    tagsByDateIdx.set(tags[i].date, i);
  }

  // dR-Funktion (analog zu reaktionen)
  function calcDR(date: string): number | null {
    const idx = tagsByDateIdx.get(date);
    if (idx == null || idx >= tags.length - 1) return null;
    const t = tags[idx];
    if (t.load == null || t.load <= 30) return null;
    const tNext = tags[idx + 1];
    if (!tNext || tNext.krank) return null;
    const rNext = tNext.readiness;
    if (rNext == null) return null;

    const basisReadiness: number[] = [];
    for (let j = idx - 27; j <= idx; j++) {
      if (j < 0 || j >= tags.length) continue;
      const tRef = tags[j];
      if (tRef.krank) continue;
      if (tRef.readiness == null) continue;
      basisReadiness.push(tRef.readiness);
    }
    if (basisReadiness.length < 14) return null;
    const basis = basisReadiness.reduce((a, b) => a + b, 0) / basisReadiness.length;
    return rNext - basis;
  }

  let lockerGeplant = 0;

  // ProWoche: letzte 8 Wochen
  const proWocheMap = new Map<string, GrauWoche>();

  for (let w = 7; w >= 0; w--) {
    const montag = new Date();
    montag.setDate(montag.getDate() - montag.getDay() + 1 - w * 7);
    const montagStr = montag.toLocaleDateString('sv-SE');

    proWocheMap.set(montagStr, {
      woche: montagStr,
      geplantLocker: 0,
      grau: 0,
    });
  }

  // Vorherige Woche berechnen für Zuordnung
  function dateInWeek(dateStr: string): string | null {
    const parts = dateStr.split('-');
    if (parts.length !== 3) return null;
    const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    if (isNaN(d.getTime())) return null;
    const dayOfWeek = d.getDay(); // 0=So, 1=Mo, ...
    const montagOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    const montag = new Date(d);
    montag.setDate(montag.getDate() + montagOffset);
    return montag.toLocaleDateString('sv-SE');
  }

  for (const t of tags) {
    const klasse = zonenKlasse(t.planZone);

    // Prüfe: planZone-Klasse locker und load > 30
    if (klasse === 'locker' && t.load != null && t.load > 30) {
      lockerGeplant++;
      const zKl = zonenKlasse(t.zone);
      const istGrau = zKl === 'grau' || zKl === 'hart';
      if (istGrau) faelle.push({ date: t.date, planZone: t.planZone, zone: t.zone, load: t.load, dR: calcDR(t.date) });
      const woek = dateInWeek(t.date);
      const w = woek ? proWocheMap.get(woek) : undefined;
      if (w) {
        w.geplantLocker++;
        if (istGrau) w.grau++;
      }
    }
  }

  // Sortiere proWoche nach Datum absteigend (neueste zuerst)
  const proWoche = Array.from(proWocheMap.values()).sort((a, b) => a.woche.localeCompare(b.woche));

  return { faelle, lockerGeplant, proWoche };
}
