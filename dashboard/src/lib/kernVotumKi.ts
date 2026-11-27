/**
 * Umschaltung KI-Texte auf das Kern-Votum (20.11.2026).
 * Maßgeblich für Briefing, Analyse (Supervisor-Text) und Telegram ist das Tagesvotum der Cockpit-Regel
 * (buildDays). Rückfall auf die Supervisor-Ampel, wenn
 *   - der Schalter „votumQuelle“ auf 'ampel' steht (Technik › Kira-Texte), oder
 *   - für heute kein Tagesvotum vorliegt (z. B. Morgenwerte noch nicht erfasst).
 * Reine Funktion ohne Netz – die Rohdaten liefert der Aufrufer (chartData mit RECOVERY_FETCH_METRICS).
 */
import { buildDays, VOTE_LABEL } from '../screens/RecoveryPanel';

export type VotumQuelle = 'kern' | 'ampel';
export type KernVotumKi = {
  quelle: VotumQuelle;
  datum: string;
  vote: string | null;
  label: string | null;
  gruende: string[];
  rueckfall: string | null;
};

export const heuteBerlin = (d = new Date()) => d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });

export function kernVotumAusRohdaten(raw: any[], schalter: VotumQuelle | string | null | undefined, heute = heuteBerlin()): KernVotumKi {
  const leer = (grund: string): KernVotumKi => ({ quelle: 'ampel', datum: heute, vote: null, label: null, gruende: [], rueckfall: grund });
  if (schalter === 'ampel') return leer('Schalter steht auf Supervisor-Ampel');
  let t: any = null;
  try {
    t = buildDays(raw || []).find((d) => d.date === heute) || null;
  } catch {
    t = null;
  }
  if (!t || !t.vote) return leer('Tagesvotum für heute fehlt');
  return {
    quelle: 'kern',
    datum: heute,
    vote: t.vote,
    label: (VOTE_LABEL as any)[t.vote] || t.vote,
    gruende: (t.reasons || []).slice(0, 4).map(String),
    rueckfall: null,
  };
}
