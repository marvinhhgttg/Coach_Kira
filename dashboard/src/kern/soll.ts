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
export type Empfehlung = { min: number; max: number; halte: number; acwrGrenze: number | null; gedeckelt: boolean };

export function lastEmpfehlung(today: RecoveryDay, yesterday: RecoveryDay | null): Empfehlung | null {
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
  let gedeckelt = false;
  if (acwrGrenze != null && hi > acwrGrenze) { hi = r5(Math.max(0, acwrGrenze)); gedeckelt = true; if (lo > hi) lo = r5(hi * 0.6); }
  return { min: lo, max: hi, halte: r5(halte), acwrGrenze: acwrGrenze != null ? r5(acwrGrenze) : null, gedeckelt };
}

/** Liegt ein Lastwert im Soll-Rahmen? */
export function lageImRahmen(load: number | null | undefined, e: Empfehlung | null): 'im Rahmen' | 'darüber' | 'darunter' | null {
  if (load == null || !e) return null;
  return load > e.max ? 'darüber' : load < e.min ? 'darunter' : 'im Rahmen';
}
