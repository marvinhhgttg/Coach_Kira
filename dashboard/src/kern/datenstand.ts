/**
 * Coach-Kira-Rechenkern – Datenstand des Tagesvotums (26.11.2026).
 * Ein Score und die Verlässlichkeit seiner Daten sind zwei verschiedene Aussagen:
 *   vollstaendig   – alle Eingaben für das Votum liegen für heute vor
 *   eingeschraenkt – das Votum steht, aber ein Signal fehlt (Regel greift dann nicht)
 *   offen          – Readiness fehlt → kein Votum möglich (nicht entscheidungsfähig)
 * Cockpit, Widgets und Uhr nutzen nur diese Funktion.
 */
import type { RecoveryDay } from './votumRegeln';

export type DatenStufe = 'vollstaendig' | 'eingeschraenkt' | 'offen';

export type Datenstand = {
  stufe: DatenStufe;
  /** kurze Anzeige, z. B. „Daten vollständig“ */
  label: string;
  /** fehlende Signale, die das Votum beeinflussen würden */
  fehlt: string[];
  /** optionale Eingaben, die fehlen (ändern die Stufe nicht) */
  optional: string[];
  /** Datum der Morgenwerte (yyyy-mm-dd) oder null */
  datum: string | null;
};

export const DATEN_LABEL: Record<DatenStufe, string> = {
  vollstaendig: 'Daten vollständig',
  eingeschraenkt: 'Daten eingeschränkt',
  offen: 'Nicht entscheidungsfähig',
};

/** today = Tageszeile aus buildDays, heuteIso = lokales Datum */
export function datenstand(today: RecoveryDay | null | undefined, heuteIso: string): Datenstand {
  const fehlt: string[] = [];
  const optional: string[] = [];
  const istHeute = !!today && today.date === heuteIso;
  if (!today || !istHeute || today.readiness == null) {
    return {
      stufe: 'offen',
      label: DATEN_LABEL.offen,
      fehlt: [!today || !istHeute ? 'Morgenwerte von heute' : 'Trainingsbereitschaft'],
      optional,
      datum: today?.date ?? null,
    };
  }
  if (today.hrv == null) fehlt.push('HRV');
  else if (today.hrvLow == null || today.hrvHigh == null) fehlt.push('HRV-Normalband');
  if (today.rhr == null) fehlt.push('Ruhepuls');
  else if (today.rhrDelta == null) fehlt.push('Ruhepuls-Basis');
  if (today.sleepH == null) fehlt.push('Schlafdauer');
  if (today.acwr == null) fehlt.push('ACWR (Belastungsverlauf)');
  if (today.befinden == null) optional.push('Befinden');
  const stufe: DatenStufe = fehlt.length ? 'eingeschraenkt' : 'vollstaendig';
  return { stufe, label: DATEN_LABEL[stufe], fehlt, optional, datum: today.date };
}

/** „HRV fehlt“ / „Morgenwerte von heute fehlen“ / „HRV, Ruhepuls fehlen“ */
export function fehltText(liste: string[]): string {
  const mehr = liste.length > 1 || /werte/i.test(liste[0] || '');
  return `${liste.join(', ')} ${mehr ? 'fehlen' : 'fehlt'}`;
}

/** Einzeiler für Widgets/Uhr, z. B. „Daten eingeschränkt: HRV fehlt“ */
export function datenstandKurz(d: Datenstand): string {
  if (d.stufe === 'vollstaendig') return d.optional.length ? `${d.label} · ${d.optional.join(', ')} offen` : d.label;
  return `${d.label}: ${fehltText(d.fehlt)}`;
}
