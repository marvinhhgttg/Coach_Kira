/**
 * Coach-Kira-Rechenkern – Tagesvotum (Etappe 2, Vorbereitung).
 * Einstiegspunkt für Cockpit, Widgets und Server. Die Regeln liegen heute noch in
 * src/screens/RecoveryPanel.tsx (decideVote/buildDays); Etappe 2 verschiebt sie hierher,
 * ohne dass sich ein Aufrufer ändert.
 */
export { buildDays, RECOVERY_FETCH_METRICS, VOTE_LABEL, VOTE_LEVEL, VOTE_ORDER } from '../screens/RecoveryPanel';
export type { RecoveryDay, Vote, VoteGrund } from '../screens/RecoveryPanel';

/** Gemeinsame Stufenleiter: 0 rot · 1 gelb · 2 grün · 3 blau/lila */
export const STUFE_COCKPIT: Record<string, number> = { REST: 0, EASY: 1, TRAIN: 2, QUALITY: 3 };
export function stufeSupervisor(ampel: string | null | undefined): number | null {
  const a = String(ampel || '').toUpperCase().replace('Ü', 'UE');
  if (a === 'ROT') return 0;
  if (a === 'GELB' || a === 'ORANGE') return 1;
  if (a === 'GRUEN' || a === 'GRÜN') return 2;
  if (a === 'LILA' || a === 'BLAU') return 3;
  return null;
}
export const STUFE_NAME = ['rot', 'gelb', 'grün', 'blau'];

/** Vergleich: gleich | Cockpit strenger | Supervisor strenger | offen */
export function vergleiche(cockpit: number | null, supervisor: number | null) {
  if (cockpit == null || supervisor == null) return { urteil: 'offen' as const, diff: null };
  const diff = cockpit - supervisor;
  return { urteil: diff === 0 ? ('gleich' as const) : diff < 0 ? ('cockpit_strenger' as const) : ('supervisor_strenger' as const), diff };
}
