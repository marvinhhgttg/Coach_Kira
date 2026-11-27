/**
 * Tourform je Zieltour – gemeinsame Rechnung für die Tourform-Kachel und den Wochenplan (21.11.2026).
 * Ergänzt den Endurance-Score-Hebel: „lange Tage“ zählen jetzt auch über die Dauer (Garmin), nicht nur über ESS –
 * ruhige lange Einheiten haben bei Marc wenig ESS pro Stunde und fielen sonst durch.
 */
import { tourform as berechneTourform, type P4Tag, type Tourform } from './paket4';
import { zielArt, type Ziel } from './planning';
import type { GarminTage } from './garmin';
import { lerneModelle } from './essRechner';
import { eigeneBergtouren, eigeneSkitouren, setzeGarminTage, tourLastPersoenlich } from './tourEss';

export type TourformZiel = {
  ziel: Ziel;
  wochen: number;
  marsch: boolean;
  persoenlich: boolean;
  wf: Tourform & { langeDauer: number; langeSchwelleH: number; langeNurDauer: boolean };
};

const tagMs = 86400000;

export function tourformFuerZiele(p4: P4Tag[], ziele: Ziel[], garmin: GarminTage | null, heute = new Date()): TourformZiel[] {
  if (garmin) setzeGarminTage(garmin);
  const modelle = garmin && garmin.size ? lerneModelle(garmin) : [];
  // Für Märsche: längste Strecke zu Fuß und längste Zeit unterwegs an einem Tag (letzte 8 Wochen)
  let fuss: { km: number; h: number } | null = null;
  const laengste = new Map<string, number>(); // Tag → längste Aktivität (h)
  if (garmin && garmin.size) {
    const ab = new Date(heute.getTime() - 56 * tagMs).toLocaleDateString('sv-SE');
    let km = 0, h = 0;
    for (const [d, t] of garmin) {
      const lh = Math.max(0, ...(t.akt || []).map((a) => (a.dauer_s || a.bewegt_s || 0) / 3600));
      if (lh > 0) laengste.set(d, lh);
      if (d < ab) continue;
      km = Math.max(km, (t.akt || []).filter((x) => /hiking|walking|mountaineering/i.test(String(x.typ || ''))).reduce((a, x) => a + (x.km || 0), 0));
      h = Math.max(h, (t.summe?.dauer_s ?? (t.akt || []).reduce((a, x) => a + (x.dauer_s || 0), 0)) / 3600);
    }
    fuss = { km, h };
  }
  const ab70 = new Date(heute.getTime() - 70 * tagMs).toLocaleDateString('sv-SE');
  const h0 = new Date(heute); h0.setHours(0, 0, 0, 0);
  return ziele.map((z) => {
    const pers = modelle.length ? tourLastPersoenlich(z, modelle, eigeneSkitouren(garmin), eigeneBergtouren(garmin)) : null;
    const wf0 = berechneTourform(p4, z, pers?.ess);
    // Lange Tage auch über die Dauer: längste Einheit mindestens halb so lang wie die Tour (mindestens 1,5 h)
    const langeSchwelleH = Math.max(1.5, (z.dauer_h || 0) * 0.5);
    let langeDauer = 0;
    for (const [d, h] of laengste) if (d >= ab70 && h >= langeSchwelleH) langeDauer++;
    // Touren ab 3 h: nur die Dauer zählt (kurze, intensive Tage mit viel ESS bereiten nicht auf 4–6 h vor);
    // kürzere Touren: ESS oder Dauer
    const nurDauer = (z.dauer_h || 0) >= 3;
    const langeBeide = nurDauer ? langeDauer : Math.max(wf0.langeTage, langeDauer);
    const lange = Math.min(1, langeBeide / wf0.langeZiel);
    const prozentNeu = wf0.hoehe == null ? Math.round(100 * (0.6 * wf0.volumen + 0.4 * lange)) : Math.round(100 * (0.5 * wf0.volumen + 0.35 * lange + 0.15 * wf0.hoehe));
    const hTxt = String(Math.round(langeSchwelleH * 4) / 4).replace('.', ',');
    const langeHinweis = `Noch ${Math.max(0, wf0.langeZiel - langeBeide)} lange Tage (${nurDauer ? `ab ${hTxt} h` : `ab ${Math.round(0.6 * wf0.tourLast)} ESS oder ab ${hTxt} h`}) in den nächsten Wochen – Termine im Endurance-Score-Hebel.`;
    let hinweise = wf0.hinweise.filter((x) => !(x.startsWith('Noch ') && x.includes('lange Tage')));
    if (lange < 0.7) hinweise = [langeHinweis, ...hinweise];
    const wfL = { ...wf0, lange, langeTage: langeBeide, prozent: prozentNeu, hinweise, langeDauer, langeSchwelleH, langeNurDauer: nurDauer };
    const marsch = zielArt(z) === 'marsch' && !!fuss;
    const strecke = marsch && z.distanz_km > 0 ? Math.min(1, fuss!.km / (z.distanz_km * 0.6)) : null;
    const zeit = marsch && z.dauer_h > 0 ? Math.min(1, fuss!.h / (z.dauer_h * 0.7)) : null;
    const wf = marsch
      ? {
          ...wfL,
          lange: strecke ?? wfL.lange,
          hoehe: zeit,
          prozent: Math.round(100 * (0.2 * wfL.volumen + 0.45 * (strecke ?? wfL.lange) + 0.35 * (zeit ?? 1))),
          hinweise: [
            `Strecke: zuletzt höchstens ${Math.round(fuss!.km)} km zu Fuß an einem Tag – Ziel vor dem Marsch etwa ${Math.round(z.distanz_km * 0.6)} km.`,
            `Zeit: zuletzt höchstens ${(Math.round(fuss!.h * 10) / 10).toLocaleString('de-DE')} h unterwegs – Ziel etwa ${(Math.round(z.dauer_h * 0.7 * 10) / 10).toLocaleString('de-DE')} h am Stück.`,
          ],
        }
      : wfL;
    const wochen = Math.max(0, Math.round((new Date(z.datum + 'T12:00:00').getTime() - h0.getTime()) / (7 * tagMs)));
    return { ziel: z, wochen, marsch, persoenlich: !!pers, wf };
  });
}

/** Kurzfassung für Kira (Wochenplan-Kontext). */
export function tourformText(list: TourformZiel[]): string {
  if (!list.length) return '';
  return list
    .map(({ ziel: z, wochen, marsch, wf }) => {
      const teile = marsch
        ? `Wochenumfang ${Math.round(wf.volumen * 100)} %, Strecke ${Math.round(wf.lange * 100)} %, Zeit ${wf.hoehe == null ? '–' : Math.round(wf.hoehe * 100) + ' %'}`
        : `Wochenumfang ${Math.round(wf.volumen * 100)} %, lange Tage ${wf.langeTage}/${wf.langeZiel}, Höhe ${wf.hoehe == null ? 'nicht erfasst' : Math.round(wf.hoehe * 100) + ' %'}`;
      return `TOURFORM ${z.name} am ${z.datum} (${wochen} Wochen, ${z.dauer_h} h, ${z.hm} Hm, Prio ${z.prioritaet}): ${wf.prozent} % – ${teile}. ${wf.hinweise.join(' ')}`;
    })
    .join('\n');
}
