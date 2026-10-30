/**
 * Zeitfenster (Dienstreise, Urlaub …): speichern, Kalender-Vorschläge, Saisonrahmen nachführen, Rückblick.
 * Gespeichert im Sheet (AI_KIRA_EXTRAS, typ 'zeitfenster', je Fenster ein Eintrag, datum = von).
 * Der Saisonrahmen wird nach jeder Änderung deterministisch neu gelegt (gleiche Parameter, Fokus-Texte bleiben,
 * solange die Phase gleich bleibt) – so lesen Cockpit, Autopilot und Vorausblick am Mac denselben Stand.
 */
import { BASE_URL } from './api';
import { deleteExtra, fetchExtras, saveExtra, type ExtraItem } from './extras';
import { buildSeason, type FensterArt, type Saisonwoche, type Ziel, type Zeitfenster } from './planning';

export type { Zeitfenster, FensterArt };

const heute = () => new Date().toLocaleDateString('sv-SE');

export async function ladeFenster(): Promise<Zeitfenster[]> {
  const xs = await fetchExtras<{ zeitfenster: Zeitfenster }>('zeitfenster', 500);
  return xs
    .map((x) => x.data?.zeitfenster)
    .filter((f): f is Zeitfenster => !!f && /^\d{4}-\d{2}-\d{2}$/.test(f.von) && /^\d{4}-\d{2}-\d{2}$/.test(f.bis))
    .sort((a, b) => a.von.localeCompare(b.von));
}

const ART_TEXT: Record<FensterArt, string> = { entlastung: 'Entlastung passt', pause: 'kein Training möglich', aktiv: 'Aktivurlaub' };

export async function speichereFenster(f: Zeitfenster): Promise<void> {
  await saveExtra('zeitfenster', f.von, `${f.titel} (${ART_TEXT[f.art]}): ${f.von} bis ${f.bis}`, { zeitfenster: f }, 'cockpit');
}

export async function loescheFenster(f: Zeitfenster): Promise<void> {
  await deleteExtra('zeitfenster', f.von);
}

/** Kalender-Vorschlag (nur Zeitraum und Stichwort, keine Termintitel). */
export interface KalVorschlag { von: string; bis: string; typ: 'reise' | 'frei'; stichworte: string[] }
export async function kalenderVorschlaege(von = heute(), tage = 200): Promise<KalVorschlag[]> {
  const b = new Date(von + 'T12:00:00');
  b.setDate(b.getDate() + tage);
  const res = await fetch(`${BASE_URL}?mode=calendarFenster&von=${von}&bis=${b.toLocaleDateString('sv-SE')}&cb=${Date.now()}`, { redirect: 'follow' }).catch(() => {
    throw new Error('Kalender-Vorschläge noch nicht verfügbar – Apps-Script-Patch „Zeitfenster“ fehlt.');
  });
  let j: any = null;
  try { j = JSON.parse(await res.text()); } catch { throw new Error('Apps-Script-Patch „Zeitfenster“ fehlt (Kalender-Abruf)'); }
  if (!j || !j.ok) throw new Error(/mode|unbekannt|unknown/i.test(String(j?.error)) ? 'Apps-Script-Patch „Zeitfenster“ fehlt (Kalender-Abruf)' : (j && j.error) || 'Kalender nicht lesbar');
  return (j.fenster || []) as KalVorschlag[];
}

/** Vorschlag schon als Zeitfenster erfasst? (Überschneidung) */
export function vorschlagErfasst(v: { von: string; bis: string }, fenster: Zeitfenster[]) {
  return fenster.some((f) => f.von <= v.bis && f.bis >= v.von);
}

interface SaisonDaten { wochen: Saisonwoche[]; parameter: { wochen: number; rampePct: number; startWeekly: number; maxWeekly: number }; fenster_stand?: string }

/**
 * Saisonrahmen an die aktuellen Zeitfenster anpassen (ohne KI): gleiche Parameter, neu gelegte Phasen.
 * Fokus-Text bleibt, wenn die Phase der Woche gleich bleibt; sonst kurzer fester Fokus.
 */
export async function saisonAnpassen(fenster: Zeitfenster[]): Promise<{ geaendert: number; wochen: number } | null> {
  const [sx, zx] = await Promise.all([fetchExtras<SaisonDaten>(['saison'], 1), fetchExtras<{ ziel: Ziel }>(['ziel'], 60)]);
  const alt: ExtraItem<SaisonDaten> | undefined = sx[0];
  if (!alt?.data?.parameter || !Array.isArray(alt.data.wochen)) return null;
  const ziele = zx.map((x) => x.data?.ziel).filter((z): z is Ziel => !!z && z.datum >= heute());
  const p = alt.data.parameter;
  const neu = buildSeason({ startWeekly: p.startWeekly, wochen: p.wochen, rampePct: p.rampePct, maxWeekly: p.maxWeekly, ziele, fenster });
  const altMap = new Map(alt.data.wochen.map((w) => [w.start, w]));
  let geaendert = 0;
  const ws = neu.map((w) => {
    const a = altMap.get(w.start);
    if (a && a.phase === w.phase && (a.fenster || '') === (w.fenster || '')) return { ...w, fokus: a.fokus || '' };
    geaendert++;
    const fokus =
      w.phase === 'Pause' ? `${w.fenster || 'Zeitfenster'}: nur kurze lockere Einheiten, wenn möglich`
      : w.phase === 'Aktivurlaub' ? `${w.fenster || 'Aktivurlaub'}: lange Tage, dazwischen lockere Tage`
      : w.phase === 'Entlastung' && w.fenster ? `${w.fenster}: locker, Schlaf und Erholung im Blick`
      : a?.fokus || '';
    return { ...w, fokus };
  });
  const txt = `Saisonrahmen ${ws.length} Wochen ab ${ws[0].start}: ${ws.filter((w) => w.phase === 'Aufbau').length} Aufbau-, ${ws.filter((w) => w.phase === 'Entlastung').length} Entlastungswochen, ${ws.filter((w) => w.fenster).length} Wochen nach Zeitfenstern gelegt, Wochenlast ${ws[0].ziel_ess} → ${Math.max(...ws.map((w) => w.ziel_ess))} ESS.`;
  await saveExtra('saison', heute(), txt, { wochen: ws, parameter: p, fenster_stand: new Date().toISOString() }, 'cockpit');
  window.dispatchEvent(new CustomEvent('ck-saison-neu'));
  return { geaendert, wochen: ws.length };
}

/** Rückblick: War das Zeitfenster tatsächlich eine Entlastung, und wie hat der Körper reagiert? */
export interface FensterRueckblick {
  f: Zeitfenster;
  lastFenster: number | null; // Ø Tageslast im Fenster
  lastVorher: number | null; // Ø Tageslast der 28 Tage davor
  anteil: number | null; // lastFenster / lastVorher
  readyVorher: number | null; readyIm: number | null; readyDanach: number | null; // Ø Bereitschaft: 7 T davor, im Fenster, 7 T danach
  rhrVorher: number | null; rhrDanach: number | null;
  schlafVorher: number | null; schlafIm: number | null;
  urteil: string;
}
export interface IstZeile { date: string; load: number | null; ready: number | null; rhr: number | null; schlaf: number | null }

const mean = (xs: (number | null)[]) => {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};
const add = (d: string, n: number) => {
  const x = new Date(d + 'T12:00:00');
  x.setDate(x.getDate() + n);
  return x.toLocaleDateString('sv-SE');
};

export function rueckblick(f: Zeitfenster, ist: IstZeile[]): FensterRueckblick | null {
  const by = new Map(ist.map((r) => [r.date, r]));
  const range = (a: string, b: string) => {
    const out: IstZeile[] = [];
    for (let d = a; d <= b; d = add(d, 1)) { const r = by.get(d); if (r) out.push(r); }
    return out;
  };
  const im = range(f.von, f.bis);
  if (!im.length) return null;
  const vor28 = range(add(f.von, -28), add(f.von, -1));
  const vor7 = range(add(f.von, -7), add(f.von, -1));
  const nach7 = range(add(f.bis, 1), add(f.bis, 7));
  const lastFenster = mean(im.map((r) => r.load ?? 0));
  const lastVorher = mean(vor28.map((r) => r.load ?? 0));
  const anteil = lastFenster != null && lastVorher ? lastFenster / lastVorher : null;
  const readyVorher = mean(vor7.map((r) => r.ready)), readyIm = mean(im.map((r) => r.ready)), readyDanach = mean(nach7.map((r) => r.ready));
  const rhrVorher = mean(vor7.map((r) => r.rhr)), rhrDanach = mean(nach7.map((r) => r.rhr));
  const schlafVorher = mean(vor7.map((r) => r.schlaf)), schlafIm = mean(im.map((r) => r.schlaf));
  const teile: string[] = [];
  if (anteil != null) {
    if (f.art === 'aktiv') teile.push(anteil >= 1.1 ? 'wie geplant mehr Last' : 'weniger Last als für einen Aktivurlaub erwartet');
    else teile.push(anteil <= 0.7 ? 'echte Entlastung' : anteil <= 0.9 ? 'nur leicht entlastet' : 'keine Entlastung – Last wie sonst');
  }
  if (readyDanach != null && readyVorher != null) {
    const d = readyDanach - readyVorher;
    teile.push(d >= 5 ? `danach erholt (Bereitschaft ${d > 0 ? '+' : ''}${Math.round(d)})` : d <= -5 ? `danach schlechter erholt (Bereitschaft ${Math.round(d)})` : 'Bereitschaft danach etwa gleich');
  } else if (nach7.length === 0) teile.push('Woche danach noch offen');
  if (schlafIm != null && schlafVorher != null && schlafIm < schlafVorher - 5) teile.push(`Schlaf im Fenster schlechter (${Math.round(schlafIm)} statt ${Math.round(schlafVorher)})`);
  const r = (x: number | null) => (x == null ? null : Math.round(x * 10) / 10);
  return {
    f, lastFenster: r(lastFenster), lastVorher: r(lastVorher), anteil: anteil == null ? null : Math.round(anteil * 100) / 100,
    readyVorher: r(readyVorher), readyIm: r(readyIm), readyDanach: r(readyDanach), rhrVorher: r(rhrVorher), rhrDanach: r(rhrDanach),
    schlafVorher: r(schlafVorher), schlafIm: r(schlafIm), urteil: teile.join(' · '),
  };
}
