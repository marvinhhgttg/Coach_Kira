/**
 * Coach-Kira-Rechenkern – Soll-Rahmen für heute (Etappe 2b, 20.11.2026).
 * Unverändert aus src/screens/RecoveryPanel.tsx hierher verschoben (Cockpit „Empfehlung heute“):
 * Band je Votum × „Fitness halten“-Last, Obergrenze gedeckelt durch ACWR ≤ 1,3.
 * Das Tages-Soll im Sheet (coachE_ESS_day) ist ein Planwert (Plan/KI/Sperren) und keine Formel –
 * der Kern rechnet daher den Rahmen und prüft den Plan dagegen (lageImRahmen).
 */
import { atlStep, ctlStep } from './formeln';
import type { RecoveryDay, Vote } from './votumRegeln';

export const BAND: Record<Vote, [number, number]> = { REST: [0, 0.25], EASY: [0.3, 0.8], TRAIN: [0.7, 1.25], QUALITY: [1.0, 1.7] };
export type Empfehlung = {
  min: number; max: number; halte: number; acwrGrenze: number | null; gedeckelt: boolean;
  /** 22.11.2026: Obergrenze aus deiner Historie angehoben (Faktor × Fitness halten) */
  persoenlich?: { faktor: number; standard: number; n: number; d1: number } | null;
};

/** Persönliche Obergrenzen je Votum aus der Historie: bis zu welcher Last (× Fitness halten) war die Readiness am Folgemorgen im Mittel nicht niedriger? */
export type Baender = Partial<Record<Vote, { faktor: number; n: number; d1: number }>>;
export const BAND_MIN_N = 8;
export const BAND_FENSTER = 0.4;

/** Last zum Halten der Fitness für einen Tag (aus CTL/ATL von gestern). */
function halteLast(yesterday: RecoveryDay | null): number | null {
  const ctl0 = yesterday?.ctlPost ?? null;
  if (ctl0 == null || ctl0 <= 0) return null;
  const m0 = ctl0 / 7.04;
  for (let L = 0; L <= 500; L += 1) if (ctlStep(ctl0, L, (27 * m0 + L) / 28) >= ctl0) return L;
  return null;
}

export function persoenlicheBaender(days: RecoveryDay[]): Baender {
  const pts: { v: Vote; q: number; d1: number }[] = [];
  for (let i = 1; i + 1 < days.length; i++) {
    const d = days[i], n = days[i + 1];
    if (!d.vote || d.readiness == null || n.readiness == null || !d.activityDone) continue;
    const ist = d.actual?.load ?? d.load ?? 0;
    const h = halteLast(days[i - 1]);
    if (!h || ist <= 0) continue;
    pts.push({ v: d.vote, q: ist / h, d1: n.readiness - d.readiness });
  }
  const out: Baender = {};
  // Nur Easy und Train anheben – Rest bleibt Ruhe, Quality ist ohnehin hart gemeint
  for (const v of ['EASY', 'TRAIN'] as Vote[]) {
    const xs = pts.filter((p) => p.v === v);
    let best: { faktor: number; n: number; d1: number } | null = null;
    for (let edge = 0.8; edge <= 2.4 + 1e-9; edge += 0.2) {
      const g = xs.filter((p) => p.q > edge - BAND_FENSTER && p.q <= edge);
      if (g.length < BAND_MIN_N) continue;
      const d1 = g.reduce((a, p) => a + p.d1, 0) / g.length;
      if (d1 < 0) break;
      best = { faktor: Math.round(edge * 10) / 10, n: g.length, d1: Math.round(d1 * 10) / 10 };
    }
    if (best) out[v] = best;
  }
  return out;
}

export function lastEmpfehlung(today: RecoveryDay, yesterday: RecoveryDay | null, baender?: Baender | null): Empfehlung | null {
  const vote = today.vote;
  const atl0 = yesterday?.atlPost ?? null;
  const ctl0 = yesterday?.ctlPost ?? null;
  if (!vote || atl0 == null || ctl0 == null || ctl0 <= 0) return null;
  const m0 = ctl0 / 7.04; // Ø-Tageslast der letzten 28 Tage (CTL ≈ 7,04 × Ø28)
  const ctlNach = (L: number) => ctlStep(ctl0, L, (27 * m0 + L) / 28);
  const acwrNach = (L: number) => atlStep(atl0, L) / ctlNach(L);
  let halte = 0;
  for (let L = 0; L <= 500; L += 1) if (ctlNach(L) >= ctl0) { halte = L; break; }
  let acwrGrenze: number | null = null;
  for (let L = 0; L <= 600; L += 1) if (acwrNach(L) > 1.3) { acwrGrenze = L - 1; break; }
  const r5 = (x: number) => Math.max(0, Math.round(x / 5) * 5);
  let [lo, hi] = BAND[vote].map((f) => r5(f * halte));
  // persönliche Obergrenze nur anheben, nie senken
  const pb = baender?.[vote];
  const persoenlich = pb && pb.faktor > BAND[vote][1] ? { faktor: pb.faktor, standard: hi, n: pb.n, d1: pb.d1 } : null;
  if (persoenlich) hi = r5(pb!.faktor * halte);
  let gedeckelt = false;
  if (acwrGrenze != null && hi > acwrGrenze) { hi = r5(Math.max(0, acwrGrenze)); gedeckelt = true; if (lo > hi) lo = r5(hi * 0.6); }
  return { min: lo, max: hi, halte: r5(halte), acwrGrenze: acwrGrenze != null ? r5(acwrGrenze) : null, gedeckelt, persoenlich };
}

/** Liegt ein Lastwert im Soll-Rahmen? */
export function lageImRahmen(load: number | null | undefined, e: Empfehlung | null): 'im Rahmen' | 'darüber' | 'darunter' | null {
  if (load == null || !e) return null;
  return load > e.max ? 'darüber' : load < e.min ? 'darunter' : 'im Rahmen';
}
