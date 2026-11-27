/**
 * 23.11.2026: Trainingsbereitschaft im Garmin-Farbmuster (Anzeige überall gleich: Cockpit, Uhr, Zifferblätter, Widgets).
 * Schlecht 1–24 rot · Niedrig 25–49 orange · Mäßig 50–74 grün · Hoch 75–94 blau · Optimal 95–100 lila.
 * Nur Darstellung – das Votum rechnet weiter mit Kiras Schwellen (levelReadiness).
 */
export const TR_BAENDER = [
  { von: 0, bis: 25, farbe: '#D9433A', name: 'Schlecht' },
  { von: 25, bis: 50, farbe: '#E5822F', name: 'Niedrig' },
  { von: 50, bis: 75, farbe: '#5DB75B', name: 'Mäßig' },
  { von: 75, bis: 95, farbe: '#3A78CF', name: 'Hoch' },
  { von: 95, bis: 101, farbe: '#6E45E9', name: 'Optimal' },
] as const;

export function trBand(v: number | null | undefined) {
  if (v == null || !Number.isFinite(v)) return null;
  return TR_BAENDER.find((b) => v >= b.von && v < b.bis) ?? TR_BAENDER[TR_BAENDER.length - 1];
}
export const trFarbe = (v: number | null | undefined) => trBand(v)?.farbe ?? '#9aa5b1';
export const trName = (v: number | null | undefined) => trBand(v)?.name ?? '';
