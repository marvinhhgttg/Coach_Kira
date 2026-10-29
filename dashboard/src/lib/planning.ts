/**
 * Zukunftsplanung: Saisonrahmen, Tour-/Zielanforderung, Lastfaktor-Kalibrierung.
 * Alles deterministisch – das lokale Modell formuliert nur Fokus, Blöcke und Begründungen.
 */

export interface Ziel {
  datum: string; // YYYY-MM-DD
  name: string;
  art: 'bergtour' | 'bike_hike' | 'rad' | 'lauf' | 'skitour' | 'sonstiges';
  hm: number;
  dauer_h: number;
  distanz_km: number;
  prioritaet: 'A' | 'B' | 'C';
  notiz?: string;
}

export interface Saisonwoche {
  start: string; // Montag
  kw: number;
  phase: 'Aufbau' | 'Entlastung' | 'Tapering' | 'Ziel' | 'Erholung';
  ziel_ess: number;
  sportmix: string;
  fokus?: string;
  ziel?: string;
}

const iso = (d: Date) => d.toLocaleDateString('sv-SE');
export function mondayOf(d: Date) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12);
  const wd = (x.getDay() + 6) % 7;
  x.setDate(x.getDate() - wd);
  return x;
}
function isoWeek(d: Date) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t.getTime() - y0.getTime()) / 86400000 + 1) / 7);
}

/** Sportmix nach Monat (Kadeltshofen, Allgäu/Alpen in Reichweite) */
export function sportmixFuerMonat(m: number): string {
  if (m === 11) return 'Rad (Pendeln/Indoor), Lauf, Kraft/Mobility, erste Skike-Einheiten';
  if (m === 12 || m === 1 || m === 2) return 'Skitour, Langlauf/Skike, Lauf, Rad indoor, Kraft';
  if (m === 3) return 'Skitour (Frühjahr), Lauf, Rad draußen wieder aufbauen';
  if (m === 4 || m === 5) return 'Rad/Gravel, Lauf, erste Bergwanderungen';
  if (m >= 6 && m <= 9) return 'Bergtouren, Bike & Hike, Gravel/Rennrad, Lauf';
  return 'Gravel/Rad, Lauf, Bergwandern, Übergang zu Skike';
}

/**
 * Saisonrahmen: 3:1-Rhythmus, Steigerung pro Aufbauwoche, Tapering vor A-/B-Zielen, Erholung danach.
 * startWeekly = aktuelle Wochenlast (Ø der letzten 4 Wochen).
 */
export function buildSeason(opts: { startWeekly: number; wochen: number; rampePct: number; ziele: Ziel[]; maxWeekly?: number; heute?: Date }): Saisonwoche[] {
  const heute = opts.heute || new Date();
  const mon0 = mondayOf(heute);
  if (heute.getDay() !== 1) mon0.setDate(mon0.getDate() + 7); // ab nächstem Montag
  const max = opts.maxWeekly || Math.round(opts.startWeekly * 1.6);
  const out: Saisonwoche[] = [];
  let level = opts.startWeekly;
  let buildCount = 0;
  for (let w = 0; w < opts.wochen; w++) {
    const start = new Date(mon0);
    start.setDate(start.getDate() + 7 * w);
    const end = new Date(start);
    end.setDate(end.getDate() + 6);
    const s = iso(start), e = iso(end);
    const zielHier = opts.ziele.find((z) => z.datum >= s && z.datum <= e && z.prioritaet !== 'C');
    const zielNaechste = opts.ziele.find((z) => {
      const d1 = new Date(start); d1.setDate(d1.getDate() + 7);
      return z.prioritaet === 'A' && z.datum >= iso(d1) && z.datum <= iso(new Date(d1.getTime() + 6 * 86400000));
    });
    const zielVorher = opts.ziele.find((z) => {
      const d0 = new Date(start); d0.setDate(d0.getDate() - 7);
      return z.prioritaet !== 'C' && z.datum >= iso(d0) && z.datum < s;
    });
    let phase: Saisonwoche['phase'];
    let ess: number;
    if (zielHier) { phase = 'Ziel'; ess = level * (zielHier.prioritaet === 'A' ? 0.7 : 0.85); }
    else if (zielNaechste) { phase = 'Tapering'; ess = level * 0.75; }
    else if (zielVorher) { phase = 'Erholung'; ess = level * 0.65; buildCount = 0; }
    else if (buildCount === 3) { phase = 'Entlastung'; ess = level * 0.65; buildCount = 0; }
    else { level = Math.min(max, w === 0 ? level : level * (1 + opts.rampePct / 100)); phase = 'Aufbau'; ess = level; buildCount++; }
    out.push({
      start: s,
      kw: isoWeek(start),
      phase,
      ziel_ess: Math.round(ess / 10) * 10,
      sportmix: sportmixFuerMonat(start.getMonth() + 1),
      ziel: zielHier ? `${zielHier.name} (${zielHier.prioritaet})` : undefined,
    });
  }
  return out;
}

/** Last einer Tour grob schätzen: Dauer × Faktor Z2 + Höhenmeter × Höhenfaktor */
export function tourLast(z: Ziel, faktoren?: { load: any[]; elev: any[] }) {
  const typ = z.art === 'rad' ? 'Bike' : z.art === 'lauf' ? 'Run' : z.art === 'skitour' ? 'Skitour' : 'Other';
  const f = (faktoren?.load || []).find((x) => String(x.ActivityType) === typ && /Z2|Default/.test(String(x.Zone)));
  const perH = Number(f?.Factor) || (typ === 'Bike' ? 85 : typ === 'Run' ? 120 : 50);
  const ef = (faktoren?.elev || []).find((x) => String(x.ActivityType) === (typ === 'Skitour' ? 'Other' : typ));
  const per100 = Number(ef?.ElevFactor_per_100m) || (typ === 'Run' ? 10 : typ === 'Bike' ? 4 : 6);
  const bikeHike = z.art === 'bike_hike' ? 0.5 * 85 * z.dauer_h * 0.5 : 0; // halbe Zeit auf dem Rad
  const base = z.art === 'bike_hike' ? 50 * z.dauer_h * 0.5 + bikeHike : perH * z.dauer_h;
  return Math.round(base + (z.hm / 100) * per100);
}

/** Tourlast aus dem persönlichen ESS-Rechner (lib/tourEss.ts) – ersetzt die Faustregel, wenn vorhanden. */
export interface TourLastPersoenlich { ess: number; lo: number; hi: number }

export interface Bereitschaft {
  tourLast: number;
  tourLastQuelle: 'ess-rechner' | 'faustregel';
  tourLastFaust: number; // Faustregel zum Vergleich
  tourLastLo?: number;
  tourLastHi?: number;
  ctl: number | null; // Ø Wochenlast der letzten 4 Wochen
  ctlZiel: number; // 2 × Tourlast
  maxTag: number; // höchste Tageslast der letzten 8 Wochen
  maxHm: number; // meiste Höhenmeter an einem Tag (8 Wochen)
  langeTageZiel: number;
  hmZiel: number;
  wochen: number;
  ampel: 'gruen' | 'gelb' | 'rot';
  punkte: string[];
}

/** Abgleich Anforderung vs. aktueller Stand (Faustregeln, transparent) */
export function zielBereitschaft(
  z: Ziel,
  hist: { date: string; load: number | null; elev: number | null; ctl: number | null }[],
  faktoren?: { load: any[]; elev: any[] },
  pers?: TourLastPersoenlich | null,
  maxDauerH?: number | null, // längster Tag (Summe Aktivitätszeit) der letzten 8 Wochen laut Garmin
): Bereitschaft {
  const faust = tourLast(z, faktoren);
  const tl = pers && pers.ess > 0 ? pers.ess : faust;
  const last8 = hist.slice(-56);
  const maxTag = Math.max(0, ...last8.map((r) => r.load ?? 0));
  const maxHm = Math.max(0, ...last8.map((r) => r.elev ?? 0));
  // Die CTL des Sheets hat eine eigene Skala; als Fitnessmaß dient deshalb die Ø-Wochenlast der letzten 4 Wochen.
  // Faustregel: Ø-Wochenlast ≥ 2 × Tourlast → der lange Tag ist gut verkraftbar.
  const w4 = hist.slice(-28).map((r) => r.load ?? 0);
  const ctl = w4.length ? Math.round((w4.reduce((a, b) => a + b, 0) / w4.length) * 7) : null;
  const ctlZiel = Math.round(tl * 2);
  const langeTageZiel = Math.round(tl * 0.7);
  const hmZiel = Math.round(z.hm * 0.7);
  const wochen = Math.max(0, Math.round((new Date(z.datum + 'T12:00:00').getTime() - Date.now()) / (7 * 86400000)));
  const punkte: string[] = [];
  let rot = 0, gelb = 0;
  const chk = (ok: boolean, nah: boolean, txt: string) => { punkte.push(txt); if (!ok) (nah ? gelb++ : rot++); };
  if (ctl != null) chk(ctl >= ctlZiel, ctl >= ctlZiel * 0.8, `Ø Wochenlast (4 Wochen) ${ctl} ESS – Ziel ca. ${ctlZiel} (2 × Tourlast)`);
  chk(maxTag >= langeTageZiel, maxTag >= langeTageZiel * 0.75, `Längster Tag der letzten 8 Wochen ${maxTag} ESS – vor der Tour mind. einmal ca. ${langeTageZiel} ESS`);
  const hmTage = last8.filter((r) => (r.elev ?? 0) > 0).length;
  if (z.hm > 0 && hmTage === 0)
    punkte.push(`Höhenmeter: in den letzten 8 Wochen nicht erfasst (Spalte elev_m_day leer) – vor der Tour mind. einmal ca. ${hmZiel} Hm am Stück einplanen`);
  else if (z.hm > 0) chk(maxHm >= hmZiel, maxHm >= hmZiel * 0.7, `Meiste Höhenmeter an einem Tag ${maxHm} Hm – vor der Tour mind. einmal ca. ${hmZiel} Hm`);
  if (maxDauerH != null && z.dauer_h > 0) {
    const dz = Math.round(z.dauer_h * 0.7 * 10) / 10;
    const f = (h: number) => h.toLocaleString('de-DE', { maximumFractionDigits: 1 });
    chk(maxDauerH >= dz, maxDauerH >= dz * 0.75, `Längster Tag der letzten 8 Wochen ${f(maxDauerH)} h unterwegs – vor der Tour mind. einmal ca. ${f(dz)} h am Stück`);
  }
  const ampel = rot && wochen < 6 ? 'rot' : rot || gelb ? 'gelb' : 'gruen';
  return {
    tourLast: tl,
    tourLastQuelle: pers && pers.ess > 0 ? 'ess-rechner' : 'faustregel',
    tourLastFaust: faust,
    ...(pers && pers.ess > 0 ? { tourLastLo: pers.lo, tourLastHi: pers.hi } : {}),
    ctl, ctlZiel, maxTag, maxHm, langeTageZiel, hmZiel, wochen, ampel, punkte };
}

export interface FaktorVorschlag {
  typ: string;
  zone: string;
  faktor: number;
  n: number;
  verhaeltnis: number; // Median Ist/Soll
  vorschlag: number;
}

/** Normalisiert Sport aus der timeline auf die ActivityType-Namen von KK_LOAD_FACTORS */
export function sportTyp(s: string): string {
  const t = s.toLowerCase();
  if (/run|lauf/.test(t)) return 'Run';
  if (/bike|rad|gravel|ride|commute/.test(t)) return 'Bike';
  if (/hike|wander|walk/.test(t)) return 'Hike (flach)';
  if (/row|rudern/.test(t)) return 'Row';
  if (/skitour/.test(t)) return 'Skitour';
  return 'Other';
}
export function zoneKey(z: string): string {
  const m = String(z || '').toUpperCase().match(/Z([1-5])/g);
  if (!m) return 'Default';
  return m[m.length - 1]; // höchste genannte Zone bestimmt den Faktor
}

/**
 * Kalibrierung: Ist-Last (Garmin) gegen Soll-Last (Plan-Schnappschuss am Morgen) je Sport/Zone.
 * Nur Tage mit Soll ≥ 40 und Ist ≥ 20; Vorschlag erst ab 4 Tagen und > 8 % Abweichung.
 */
export function kalibriereFaktoren(
  reviews: { date: string; loadIst: number | null; loadSoll: number | null }[],
  timeline: Record<string, { sport: string; zone: string }>,
  faktoren: { ActivityType: string; Zone: string; Factor: number }[],
): FaktorVorschlag[] {
  const groups = new Map<string, number[]>();
  for (const r0 of reviews) {
    const r = { ...r0, loadIst: Number(r0.loadIst), loadSoll: Number(r0.loadSoll) };
    if (!Number.isFinite(r.loadIst) || !Number.isFinite(r.loadSoll) || r.loadSoll < 40 || r.loadIst < 20) continue;
    const t = timeline[r.date];
    if (!t || !t.sport || /off|rest|ruhe/i.test(t.sport)) continue;
    const key = `${sportTyp(t.sport)}|${zoneKey(t.zone)}`;
    const ratio = r.loadIst / r.loadSoll;
    if (ratio < 0.3 || ratio > 3) continue; // anderer Sport als geplant o. Ä.
    groups.set(key, [...(groups.get(key) || []), ratio]);
  }
  const out: FaktorVorschlag[] = [];
  for (const [key, xs] of groups) {
    const [typ, zone] = key.split('|');
    const s = [...xs].sort((a, b) => a - b);
    const med = s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
    const f = faktoren.find((x) => x.ActivityType === typ && x.Zone === zone) || faktoren.find((x) => x.ActivityType === typ && x.Zone === 'Default');
    if (!f) continue;
    const vorschlag = Math.round((Number(f.Factor) * med) / 5) * 5;
    out.push({ typ, zone: f.Zone, faktor: Number(f.Factor), n: xs.length, verhaeltnis: Math.round(med * 100) / 100, vorschlag });
  }
  return out.sort((a, b) => b.n - a.n);
}
