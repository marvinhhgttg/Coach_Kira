/**
 * Coach-Kira-Rechenkern – Garmin-Signale im Tagesvotum (27.11.2026).
 *
 * Garmins tägliche Trainingsvorschläge entstehen auf der Uhr und sind über keine bekannte Schnittstelle
 * abrufbar. Ins Votum gehen deshalb die Grundlagen dieser Vorschläge ein, die der Mac live liest:
 *   - Erholungszeit am Morgen (erste Trainingsbereitschaft des Tages, nicht der Wert nach dem Training)
 *   - Trainingszustand (Garmin-Kürzel, z. B. STRAINED_1, OVERREACHING_1)
 *   - Belastungsfokus (4-Wochen-Last niedrig aerob / hoch aerob / anaerob gegen Garmins Zielbereich)
 * Regeln (nur heute, nur senken – die Timeline bleibt die Wahrheit, die Historie bleibt unverändert):
 *   G1 Erholungszeit ≥ 60 h → maximal Easy · ≥ 36 h → keine Qualität (maximal Train)
 *   G2 Trainingszustand „Überlastet“ (STRAINED) oder „Überzogen“ (OVERREACHING) → maximal Easy
 *   Belastungsfokus senkt nie, er gibt den Schwerpunkt im Rahmen des Votums vor.
 * Der Garmin-Tagesvorschlag (von Marc auf der Uhr/im Cockpit eingetragen) ist eine Zweitmeinung:
 * Kira zeigt den Abgleich, das Votum ändert er nicht.
 */
import type { Vote } from './votumRegeln';

export type GarminFokus = {
  lo: number | null; hi: number | null; an: number | null;
  loMin: number | null; loMax: number | null; hiMin: number | null; hiMax: number | null; anMin: number | null; anMax: number | null;
  phrase?: string | null;
};
export type GarminVorschlag = { datum: string; typ: VorschlagTyp; min: number; zone?: string | null; at?: number };
export type GarminSignale = {
  erholH?: number | null;
  status?: string | null;
  fokus?: GarminFokus | null;
  vorschlag?: GarminVorschlag | null;
};

export const VORSCHLAG_TYPEN = ['ruhe', 'erholung', 'basis', 'lang', 'tempo', 'schwelle', 'vo2max', 'anaerob'] as const;
export type VorschlagTyp = (typeof VORSCHLAG_TYPEN)[number];
export const VORSCHLAG_LABEL: Record<VorschlagTyp, string> = {
  ruhe: 'Ruhe', erholung: 'Erholung', basis: 'Basis', lang: 'Lang', tempo: 'Tempo', schwelle: 'Schwelle', vo2max: 'VO2max', anaerob: 'Anaerob',
};
/** Anspruch des Garmin-Vorschlags in Kiras Votumsstufen */
export const VORSCHLAG_VOTUM: Record<VorschlagTyp, Vote> = {
  ruhe: 'REST', erholung: 'EASY', basis: 'TRAIN', lang: 'TRAIN', tempo: 'QUALITY', schwelle: 'QUALITY', vo2max: 'QUALITY', anaerob: 'QUALITY',
};
/** Standardzone je Typ, falls keine angegeben ist */
export const VORSCHLAG_ZONE: Record<VorschlagTyp, string> = {
  ruhe: '', erholung: 'Z2', basis: 'Z2', lang: 'Z2', tempo: 'Z3', schwelle: 'Z4', vo2max: 'Z5', anaerob: 'Z5',
};

const ORDER: Vote[] = ['REST', 'EASY', 'TRAIN', 'QUALITY'];
const LABEL: Record<Vote, string> = { QUALITY: 'Quality', TRAIN: 'Train', EASY: 'Easy', REST: 'Rest' };
const fmt = (v: number, d = 0) => v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });

export const ERHOL_EASY_H = 60;
export const ERHOL_TRAIN_H = 36;

/** Garmin-Trainingszustand auf Deutsch (Kürzel ohne _1/_2) */
export function statusLabel(s: string | null | undefined): string | null {
  const k = String(s || '').toUpperCase().replace(/_\d+$/, '');
  const m: Record<string, string> = {
    PRODUCTIVE: 'Produktiv', MAINTAINING: 'Erhaltend', RECOVERY: 'Erholung', UNPRODUCTIVE: 'Unproduktiv', DETRAINING: 'Leistungsverlust',
    PEAKING: 'Leistungshoch', STRAINED: 'Überlastet', OVERREACHING: 'Überzogen', NO_STATUS: 'Kein Status',
  };
  return k ? m[k] || k : null;
}

/** Regelstufe G1/G2: höchstzulässiges Votum aus den Garmin-Signalen. */
export function garminDeckel(s: GarminSignale | null | undefined): { cap: Vote; grund: string }[] {
  const out: { cap: Vote; grund: string }[] = [];
  if (!s) return out;
  const e = s.erholH;
  if (e != null && Number.isFinite(e)) {
    if (e >= ERHOL_EASY_H) out.push({ cap: 'EASY', grund: `Garmin-Erholungszeit ${fmt(e)} h (≥ ${ERHOL_EASY_H} h) → maximal Easy.` });
    else if (e >= ERHOL_TRAIN_H) out.push({ cap: 'TRAIN', grund: `Garmin-Erholungszeit ${fmt(e)} h (≥ ${ERHOL_TRAIN_H} h) → keine Qualitätseinheit.` });
  }
  const k = String(s.status || '').toUpperCase();
  if (/^(STRAINED|OVERREACHING)/.test(k)) out.push({ cap: 'EASY', grund: `Garmin-Trainingszustand „${statusLabel(k)}“ → maximal Easy.` });
  return out;
}

/** Belastungsfokus: welcher Bereich liegt unter Garmins Zielbereich – und was passt davon zum Votum? */
export function fokusHinweis(vote: Vote | null, f: GarminFokus | null | undefined): { mangel: ('lo' | 'hi' | 'an')[]; text: string | null } {
  if (!f) return { mangel: [], text: null };
  const mangel: ('lo' | 'hi' | 'an')[] = [];
  const unter = (v: number | null, min: number | null) => v != null && min != null && min > 0 && v < min;
  if (unter(f.lo, f.loMin)) mangel.push('lo');
  if (unter(f.hi, f.hiMin)) mangel.push('hi');
  if (unter(f.an, f.anMin)) mangel.push('an');
  const name = { lo: 'niedrig aerob (Z1–Z2)', hi: 'hoch aerob (Z3–Z4)', an: 'anaerob (Z5, Sprints)' } as const;
  if (!mangel.length) return { mangel, text: 'Garmin-Belastungsfokus ausgeglichen.' };
  // Was erlaubt das Votum heute?
  const erlaubt = (m: 'lo' | 'hi' | 'an') => (m === 'lo' ? vote !== 'REST' : m === 'hi' ? vote === 'TRAIN' || vote === 'QUALITY' : vote === 'QUALITY');
  const jetzt = mangel.filter(erlaubt);
  const spaeter = mangel.filter((m) => !erlaubt(m));
  const teile: string[] = [];
  if (jetzt.length) teile.push(`Schwerpunkt heute: ${jetzt.map((m) => name[m]).join(' und ')} – dort fehlt dir laut Garmin Belastung`);
  if (spaeter.length) teile.push(`${spaeter.map((m) => name[m]).join(' und ')} fehlt auch, passt aber nicht zum Votum ${vote ? LABEL[vote] : '–'}`);
  return { mangel, text: teile.join('; ') + '.' };
}

/** Zweitmeinung: Garmin-Tagesvorschlag gegen Kiras Votum (ändert das Votum nicht). */
export function vorschlagAbgleich(vote: Vote | null, v: GarminVorschlag | null | undefined, ess: number | null, rahmen?: { min: number; max: number } | null) {
  if (!v || !VORSCHLAG_VOTUM[v.typ]) return null;
  const anspruch = VORSCHLAG_VOTUM[v.typ];
  const zone = v.zone || VORSCHLAG_ZONE[v.typ];
  const titel = v.typ === 'ruhe' ? 'Ruhe' : `${v.min} min ${VORSCHLAG_LABEL[v.typ]}${zone ? ` ${zone}` : ''}`;
  let urteil: 'gleich' | 'garmin_mehr' | 'garmin_weniger' | 'offen' = 'offen';
  let text = `Garmin schlägt vor: ${titel}${ess != null && v.typ !== 'ruhe' ? ` (≈ ${ess} ESS)` : ''}`;
  if (vote) {
    const d = ORDER.indexOf(anspruch) - ORDER.indexOf(vote);
    urteil = d === 0 ? 'gleich' : d > 0 ? 'garmin_mehr' : 'garmin_weniger';
    text += urteil === 'gleich' ? ` – passt zu Kiras Votum ${LABEL[vote]}.`
      : urteil === 'garmin_mehr' ? ` – mehr als Kiras Votum ${LABEL[vote]}; Kira bleibt bei ${LABEL[vote]}.`
      : ` – vorsichtiger als Kiras Votum ${LABEL[vote]}; Garmins Vorschlag ist heute die sichere Wahl.`;
  } else text += '.';
  const lage = rahmen && ess != null && v.typ !== 'ruhe' ? (ess > rahmen.max ? 'darüber' : ess < rahmen.min ? 'darunter' : 'im Rahmen') : null;
  return { titel, anspruch, urteil, ess, lage, text };
}

/** ESS des Vorschlags aus deinen eigenen Zonenraten (ESS je Stunde in Z1…Z5 aus der Garmin-Historie). */
export function vorschlagEss(v: GarminVorschlag | null | undefined, raten: number[] | null | undefined): number | null {
  if (!v || v.typ === 'ruhe') return v ? 0 : null;
  if (!raten || raten.length < 5) return null;
  const z = Number(String(v.zone || VORSCHLAG_ZONE[v.typ]).replace(/[^1-5]/g, '').slice(0, 1)) || 2;
  return Math.round((v.min / 60) * raten[z - 1] / 5) * 5;
}

/** Live-Objekt vom Server (/api/garmin/live) → Signale für das heutige Votum (nur heute). */
export function signaleAusLive(live: any, heute: string): GarminSignale | null {
  if (!live || live.datum !== heute) return null;
  const erholH = live.erholHMorgen ?? null; // Morgenwert – der Wert nach dem Training würde das Votum nachträglich ändern
  const s: GarminSignale = { erholH: erholH != null ? Number(erholH) : null, status: live.status ?? null, fokus: live.fokus ?? null, vorschlag: live.vorschlag && live.vorschlag.datum === heute ? live.vorschlag : null };
  return s.erholH != null || s.status || s.fokus || s.vorschlag ? s : null;
}
