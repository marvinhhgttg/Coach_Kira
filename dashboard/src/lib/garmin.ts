/**
 * Garmin Personal API – Tagesdetails (AI_KIRA_EXTRAS, typ 'garmin_tag', geschrieben vom Mac-Dienst kira_garmin.py).
 * Je Tag: Tageswerte, Schlaf, HRV, alle Aktivitäten mit Details und Tagessumme.
 */
import { useEffect, useState } from 'react';
import { fetchExtras } from './extras';

export interface GarminAkt {
  id?: number; name?: string; typ?: string; sport?: string; start?: string;
  dauer_s?: number; bewegt_s?: number; km?: number; hm?: number; hm_ab?: number;
  hf_avg?: number; hf_max?: number; zonen_s?: number[];
  te_ae?: number; te_an?: number; te_label?: string; last?: number; kcal?: number; bb?: number;
  temp_min?: number; temp_max?: number; v_kmh?: number; watt?: number; watt_np?: number; watt_20?: number;
  kadenz?: number; vo2max?: number; tss?: number; if?: number; ort?: string; schritte?: number;
  int_mod?: number; int_vig?: number; atem?: number; stress_diff?: number; runden?: number; hoehe_max?: number; pr?: boolean;
}
export interface GarminSumme {
  n?: number; km?: number; hm?: number; dauer_s?: number; bewegt_s?: number; last?: number; kcal?: number; bb?: number; zonen_s?: number[];
}
export interface GarminTagWerte {
  schritte?: number; schritte_ziel?: number; etagen?: number; kcal_gesamt?: number; kcal_aktiv?: number; kcal_ruhe?: number; kcal_gegessen?: number;
  hf_min?: number; hf_max?: number; rhr?: number; int_mod?: number; int_vig?: number;
  stress_avg?: number; stress_max?: number; stress_q?: string;
  stress_min?: { ruhe?: number; niedrig?: number; mittel?: number; hoch?: number; aktiv?: number };
  bb_min?: number; bb_max?: number; bb_geladen?: number; bb_verbraucht?: number; bb_ende?: number;
  atem_avg?: number; atem_min?: number; atem_max?: number; spo2_avg?: number; spo2_min?: number; wasser_ml?: number; jetlag?: string;
}
export interface GarminSchlaf {
  dauer_s?: number; tief_s?: number; leicht_s?: number; rem_s?: number; wach_s?: number; nap_s?: number; score?: number;
  wachphasen?: number; stress?: number; hf?: number; atem?: number; spo2?: number; start?: string; ende?: string; bedarf_min?: number; feedback?: string;
}
export interface GarminHrv { nacht?: number; woche?: number; max5?: number; status?: string; lo?: number; hi?: number }
export interface GarminTag {
  v: number; tag?: GarminTagWerte; schlaf?: GarminSchlaf; hrv?: GarminHrv; akt?: GarminAkt[]; summe?: GarminSumme;
  gewicht_kg?: number; fitnessalter?: number; bb_aufwachen?: number;
}
export type GarminTage = Map<string, GarminTag>;

let cache: { at: number; p: Promise<GarminTage> } | null = null;
const CACHE_MS = 10 * 60 * 1000;

/** Alle Garmin-Tage (bis 1000). Fehler -> leere Map (Patch evtl. nicht aktiv). */
export function fetchGarminTage(): Promise<GarminTage> {
  if (cache && Date.now() - cache.at < CACHE_MS) return cache.p;
  const p = fetchExtras<GarminTag>('garmin_tag', 1000)
    .then((items) => {
      const m: GarminTage = new Map();
      for (const it of items) if (it.data && typeof it.data === 'object') m.set(it.datum, it.data);
      return m;
    })
    .catch(() => new Map() as GarminTage);
  cache = { at: Date.now(), p };
  return p;
}

/** Hook: null solange geladen wird. */
export function useGarminTage(): { tage: GarminTage | null } {
  const [tage, setTage] = useState<GarminTage | null>(null);
  useEffect(() => {
    let alive = true;
    fetchGarminTage().then((m) => alive && setTage(m));
    return () => { alive = false; };
  }, []);
  return { tage };
}

// ---------------------------------------------------------------------------
// Hilfsfunktionen
// ---------------------------------------------------------------------------
export const ZONEN_FARBEN = ['#94a3b8', '#38bdf8', '#22c55e', '#f97316', '#ef4444'];

/** "1:37 h" bzw. "45 min"; null -> "—". */
export function fmtDauer(s?: number | null): string {
  if (s == null || !Number.isFinite(s) || s <= 0) return '—';
  const min = Math.round(s / 60);
  if (min < 60) return `${min} min`;
  return `${Math.floor(min / 60)}:${String(min % 60).padStart(2, '0')} h`;
}
/** Stunden mit einer Nachkommastelle, z. B. "7,3 h". */
export function fmtStunden(s?: number | null): string {
  if (s == null || !Number.isFinite(s) || s <= 0) return '—';
  return `${(s / 3600).toFixed(1).replace('.', ',')} h`;
}
/** Höhenmeter je km (nur ab 1 km). */
export function hmProKm(km?: number | null, hm?: number | null): number | null {
  if (km == null || hm == null || km < 1) return null;
  return Math.round((hm / km) * 10) / 10;
}
/** Last je Stunde (ab 10 min). */
export function lastProStunde(last?: number | null, sekunden?: number | null): number | null {
  if (last == null || sekunden == null || sekunden < 600) return null;
  return Math.round(last / (sekunden / 3600));
}
/** Zonenanteile in % (Summe 100) oder null. */
export function zonenProzent(z?: number[] | null): number[] | null {
  if (!z || z.length < 5) return null;
  const sum = z.slice(0, 5).reduce((a, b) => a + (b || 0), 0);
  if (sum <= 0) return null;
  const raw = z.slice(0, 5).map((v) => ((v || 0) / sum) * 100);
  const r = raw.map(Math.round);
  const diff = 100 - r.reduce((a, b) => a + b, 0);
  if (diff) { const i = raw.indexOf(Math.max(...raw)); r[i] += diff; }
  return r;
}
/** Dominante Zone 1..5. */
export function dominanteZone(z?: number[] | null): number | null {
  if (!z || z.length < 5) return null;
  let best = -1, idx = -1;
  z.slice(0, 5).forEach((v, i) => { if ((v || 0) > best) { best = v || 0; idx = i; } });
  return best > 0 ? idx + 1 : null;
}
/** Zonensekunden addieren. */
export function zonenSumme(list: (number[] | undefined | null)[]): number[] | null {
  const out = [0, 0, 0, 0, 0];
  let any = false;
  for (const z of list) if (z && z.length >= 5) { any = true; for (let i = 0; i < 5; i++) out[i] += z[i] || 0; }
  return any ? out : null;
}

export const TE_LABEL_DE: Record<string, string> = {
  AEROBIC_BASE: 'Grundlage', RECOVERY: 'Erholung', TEMPO: 'Tempo', LACTATE_THRESHOLD: 'Schwelle',
  VO2MAX: 'VO2max', ANAEROBIC_CAPACITY: 'Anaerob', SPEED: 'Sprint', NO_BENEFIT: 'ohne Effekt',
};
/** Typische Zone eines TE-Labels. */
export const TE_LABEL_ZONE: Record<string, number> = {
  RECOVERY: 1, AEROBIC_BASE: 2, TEMPO: 3, LACTATE_THRESHOLD: 4, VO2MAX: 5, ANAEROBIC_CAPACITY: 5, SPEED: 5,
};
export function teLabelDe(x?: string | null): string {
  if (!x) return '—';
  return TE_LABEL_DE[x] ?? x.replace(/_/g, ' ').toLowerCase();
}
/** Höchste Zonenziffer aus einem Eintrag wie "Z2, Z3". */
export function zoneAusText(z?: string | null): number | null {
  const m = (z || '').match(/[1-5]/g);
  return m ? Math.max(...m.map(Number)) : null;
}

export const HRV_STATUS_DE: Record<string, string> = {
  BALANCED: 'ausgeglichen', UNBALANCED: 'unausgeglichen', LOW: 'niedrig', POOR: 'schwach', NONE: '—',
};
export const STRESS_Q_DE: Record<string, string> = {
  calm: 'ruhig', balanced: 'ausgeglichen', stressful: 'stressig', very_stressful: 'sehr stressig', unknown: '—',
};

/** Summe über Garmin-Tage in einem Datumsbereich [von, bis]. */
export function garminSumme(tage: GarminTage | null, von: string, bis: string) {
  let km = 0, hm = 0, sek = 0, n = 0, tageMit = 0;
  if (tage) for (const [d, g] of tage) {
    if (d < von || d > bis || !g.summe) continue;
    tageMit++;
    km += g.summe.km || 0; hm += g.summe.hm || 0; sek += g.summe.bewegt_s || g.summe.dauer_s || 0; n += g.summe.n || 0;
  }
  return { km, hm, sek, n, tageMit };
}
