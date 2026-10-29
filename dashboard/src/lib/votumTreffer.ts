/**
 * „Votum befolgt“: rückwirkend für die ganze Historie.
 * Tagesvotum = dieselbe feste Regel wie im Recovery Brief (buildDays/decideVote, Stand Vortag).
 * Befolgt = tatsächlicher Anspruch der Einheit (demandOf) ≤ Votum.
 * Reaktion am Folgemorgen = gleiche Formel wie die Treffsicherheit des Mac-Dienstes (kira_extras.treffer_fuer):
 *   Basis = Mittel der 7 Tage bis einschließlich Tag D, Score = ΔReadiness/10 + ΔHRV/5 − ΔRuhepuls/3 + ΔSchlafscore/10.
 * Kein Sprachmodell beteiligt.
 */
import { demandOf, VOTE_ORDER, type RecoveryDay } from '../screens/RecoveryPanel';
import type { Vote } from '../screens/RecoveryPanel';

export type VotumTag = {
  datum: string;
  votum: Vote;
  anspruch: Vote;
  last: number;
  befolgt: boolean;
  /** härter als das Votum erlaubt hätte */
  haerter: boolean;
  score: number;
  delta: { readiness: number | null; hrv: number | null; rhr: number | null; schlaf: number | null };
};

type Feld = 'readiness' | 'hrv' | 'rhr' | 'sleepScore';
const wert = (d: RecoveryDay | undefined, f: Feld): number | null => {
  if (!d) return null;
  const v = (d as any)[f]; // HRV: plausibilisierter Wert (10–200)
  return typeof v === 'number' && Number.isFinite(v) && v > 0 ? v : null;
};

export function votumTreffer(days: RecoveryDay[]): VotumTag[] {
  const byDate = new Map(days.map((d) => [d.date, d]));
  const plus1 = (iso: string) => {
    const t = new Date(iso + 'T12:00:00Z');
    t.setUTCDate(t.getUTCDate() + 1);
    return t.toISOString().slice(0, 10);
  };
  const out: VotumTag[] = [];
  days.forEach((d, idx) => {
    if (!d.vote) return;
    const n = byDate.get(plus1(d.date));
    if (!n || wert(n, 'readiness') == null) return;
    const last = d.actual.load ?? d.load;
    if (last == null) return;
    const anspruch = demandOf(last, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
    const befolgt = VOTE_ORDER.indexOf(anspruch) <= VOTE_ORDER.indexOf(d.vote);
    const fenster = days.slice(Math.max(0, idx - 6), idx + 1);
    const delta = (f: Feld): number | null => {
      const w = wert(n, f);
      const b = fenster.map((x) => wert(x, f)).filter((v): v is number => v != null);
      if (w == null || !b.length) return null;
      return Math.round((w - b.reduce((a, c) => a + c, 0) / b.length) * 10) / 10;
    };
    const dl = { readiness: delta('readiness'), hrv: delta('hrv'), rhr: delta('rhr'), schlaf: delta('sleepScore') };
    const score = Math.round(((dl.readiness ?? 0) / 10 + (dl.hrv ?? 0) / 5 - (dl.rhr ?? 0) / 3 + (dl.schlaf ?? 0) / 10) * 100) / 100;
    out.push({ datum: d.date, votum: d.vote, anspruch, last, befolgt, haerter: !befolgt, score, delta: dl });
  });
  return out.sort((a, b) => (a.datum < b.datum ? 1 : -1));
}

export type VotumBilanz = { n: number; befolgtPct: number | null; scoreBefolgt: number | null; scoreHaerter: number | null; nBefolgt: number; nHaerter: number };

export function votumBilanz(tage: VotumTag[]): VotumBilanz {
  const b = tage.filter((t) => t.befolgt);
  const h = tage.filter((t) => !t.befolgt);
  const avg = (xs: VotumTag[]) => (xs.length ? Math.round((xs.reduce((a, t) => a + t.score, 0) / xs.length) * 100) / 100 : null);
  return {
    n: tage.length,
    befolgtPct: tage.length ? Math.round((b.length / tage.length) * 100) : null,
    scoreBefolgt: avg(b),
    scoreHaerter: avg(h),
    nBefolgt: b.length,
    nHaerter: h.length,
  };
}
