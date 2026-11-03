import { useEffect, useState } from 'react';
import { localChat } from '../lib/localLLM';
import { useEngine } from '../lib/engine';
import { fremdeZahlen } from '../lib/fazitCheck';
export { fremdeZahlen };

/**
 * KI-Fazit zu einem Bericht – lokal mit coach-kira (bzw. Cloud laut Schalter), pro Gerät zwischengespeichert.
 * Schutz gegen Fehldeutungen:
 *  - Fazit ist an Titel, Sportart, Filter und alle Fakten gebunden (Hash). Ändert sich etwas, gilt das alte als veraltet.
 *  - Richtung jeder Kennzahl ist fest vorgegeben (RICHTUNG), das Modell darf sie nicht selbst erraten.
 *  - Bei dünner Datenlage keine Trainingsempfehlung.
 *  - Nachprüfung: Zahlen im Fazit, die nicht in den Fakten stehen, werden markiert.
 */
const KEY = 'ck_kira_fazit_v2';
const PROMPT_VERSION = 'p2';
type Eintrag = { text: string; ts: string; warn?: string[] };
function readCache(): Record<string, Eintrag> {
  try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch { return {}; }
}
function writeCache(c: Record<string, Eintrag>) {
  try {
    const keys = Object.keys(c);
    if (keys.length > 60) for (const k of keys.slice(0, keys.length - 60)) delete c[k];
    localStorage.setItem(KEY, JSON.stringify(c));
  } catch { /* ignore */ }
}
function hash(s: string) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return String(h);
}
const idFuer = (title: string, facts: string[], data: unknown, evidenz: string) =>
  title + ':' + hash(JSON.stringify({ v: PROMPT_VERSION, title, facts, data: data ?? null, evidenz }));

/** Feste Richtung der Kennzahlen – gilt für alle Berichte. */
export const RICHTUNG = [
  'Puls je km / Schläge je km: NIEDRIGER = effizienter (besser).',
  'Tempo je Herzschlag, Leistung je Herzschlag: HÖHER = besser.',
  'Aerobe Entkopplung in %: NIEDRIGER = besser; unter 5 % = stabile Grundlage. Positive Abweichung = stärker entkoppelt = schlechter.',
  'Watt je km/h (Rad, Normalized Power je Tempo): NIEDRIGER = weniger Aufwand fürs gleiche Tempo = effizienter. Ein Anstieg ist KEINE Effizienzsteigerung; er kann auch an Rad/Untergrund (Gravel vs. Rennrad) liegen.',
  'Garmin-Laufleistung (Laufwatt): Seit August 2026 zeigt Garmin bei gleichem Tempo mehr Watt (Messbruch). Aus Laufwatt-Trends über diesen Zeitpunkt hinweg KEINE Schlüsse auf Laufökonomie ziehen.',
  'HRV relativ zur Basis, Schlafscore, Body Battery, Trainingsbereitschaft: HÖHER = besser erholt.',
  'Stress (Garmin): NIEDRIGER = besser.',
  'Effektstärke „kaum“ oder „leicht“ bzw. Unterschiede unter 1 Prozentpunkt: kein belastbarer Effekt.',
  'CTL/ATL: Eine sinkende CTL ist eine sinkende modellierte Belastung, kein Leistungsverlust. Eine laufende (unvollständige) Woche nicht als Ruhewoche bewerten.',
];

/** Zuletzt erstelltes Fazit zu genau diesen Fakten (für Druck-Berichte). */
export function fazitAusCache(title: string, facts: string[], data?: unknown, evidenz: 'schwach' | 'mittel' | 'gut' = 'mittel'): string | null {
  return readCache()[idFuer(title, facts, data, evidenz)]?.text ?? null;
}

export function KiraFazit({ title, facts, data, onText, evidenz = 'mittel', regeln }: {
  title: string; facts: string[]; data?: unknown; onText?: (t: string) => void;
  /** schwach = keine Trainingsempfehlung, nur Beobachtung */
  evidenz?: 'schwach' | 'mittel' | 'gut';
  /** zusätzliche, berichtsspezifische Richtungs- und Deutungsregeln */
  regeln?: string[];
}) {
  const [engine] = useEngine();
  const id = idFuer(title, facts, data, evidenz);
  const [eintrag, setEintrag] = useState<Eintrag | null>(() => readCache()[id] ?? null);
  const [live, setLive] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);
  // Sportart/Filter/Daten geändert → passendes Fazit laden oder leer (nie das alte stehen lassen)
  useEffect(() => { setEintrag(readCache()[id] ?? null); setLive(null); setErr(null); }, [id]);

  async function run() {
    setBusy(true);
    setErr(null);
    setLive('');
    const quelle = facts.join('\n') + '\n' + JSON.stringify(data ?? {});
    try {
      const out = await localChat(
        [
          {
            role: 'system',
            content: [
              'Du bist Coach Kira, Sportwissenschaftlerin. Du sprichst Marc direkt per Du an. Antworte auf Deutsch, ohne Überschrift, in 3 bis 5 Sätzen.',
              'Regeln (verbindlich):',
              '1. Nutze nur Zahlen, die wörtlich in den Fakten stehen. Rechne nichts Neues aus, runde nicht um.',
              '2. Die Richtung jeder Kennzahl steht unten fest. Deute sie NIE anders herum. Vertausche keine Gruppen: Nenne immer, welche Gruppe welchen Wert hat.',
              '3. Effekte „kaum“/„leicht“, Unterschiede unter 1 Prozentpunkt oder Gruppen mit weniger als 5 Einheiten sind Rauschen: als „kein klarer Unterschied“ benennen, keine Empfehlung daraus.',
              '4. Keine Empfehlung zu mehr Intensität, Kadenzänderung oder festen Regeln, wenn die Fakten sie nicht klar stützen.',
              '5. Zusammenhang ist keine Ursache – formuliere vorsichtig („geht einher mit“).',
              evidenz === 'schwach'
                ? '6. Die Datenlage dieses Berichts ist dünn: Gib KEINE Trainingsempfehlung. Schließe mit „Empfehlung: noch keine – zu wenige Daten.“'
                : '6. Schließe mit genau einer Zeile „Empfehlung: …“ – nur wenn ein Befund klar und belastbar ist, sonst „Empfehlung: keine – die Unterschiede sind zu klein.“',
              '',
              'Richtung der Kennzahlen:',
              ...RICHTUNG.map((r) => '- ' + r),
              ...(regeln?.length ? ['', 'Für diesen Bericht zusätzlich:', ...regeln.map((r) => '- ' + r)] : []),
            ].join('\n'),
          },
          {
            role: 'user',
            content: `Bericht: ${title}\n\nRegelbasierte Fakten:\n- ${facts.join('\n- ') || 'keine'}\n\nKennzahlen (JSON):\n${JSON.stringify(data ?? {}, null, 1)}\n\nAufgabe: Was zeigen die Zahlen (Richtung beachten), wie belastbar ist das, und was folgt daraus?`,
          },
        ],
        { numCtx: 8192, onToken: setLive, temperature: 0.2 },
      );
      const warn = fremdeZahlen(out, quelle);
      // harte Deutungsfehler, die der Modelltest gezeigt hat
      if (/Lauf(ö|oe)konomie[^.]{0,60}(schlechter|verschlechter|gesunken|abgenommen)|(schlechter|verschlechter)[^.]{0,40}Lauf(ö|oe)konomie/i.test(out)) warn.push('„Laufökonomie schlechter“ ist wegen des Garmin-Messbruchs nicht belegt');
      const e: Eintrag = { text: out, ts: new Date().toISOString(), warn: warn.length ? warn : undefined };
      setEintrag(e);
      setLive(null);
      onText?.(out);
      const c = readCache();
      c[id] = e;
      writeCache(c);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
      setLive(null);
    } finally {
      setBusy(false);
    }
  }

  const text = live ?? eintrag?.text ?? null;
  return (
    <div className="mt-2 rounded border border-border px-3 py-2">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="label">Kira-Fazit{evidenz === 'schwach' && <span className="ml-2 text-2xs text-amber-400 normal-case">dünne Datenlage – nur Beobachtung</span>}</div>
        <button className="btn btn-ghost text-xs px-2 py-1" onClick={run} disabled={busy}>
          {busy ? `Kira (${engine === 'local' ? 'lokal' : 'Cloud'}) schreibt…` : text ? 'Neu erstellen' : `Fazit erstellen (${engine === 'local' ? 'lokal' : 'Cloud'})`}
        </button>
      </div>
      {err && <p className="mt-1 text-xs text-red-300">{err}</p>}
      {text ? (
        <>
          <p className="mt-1 text-sm text-ink-muted leading-relaxed whitespace-pre-wrap">{text}</p>
          {!busy && eintrag && (
            <p className="mt-1 text-2xs text-ink-dim">
              Zu genau diesen Zahlen erstellt am {new Date(eintrag.ts).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })}. Ändern sich Daten oder Filter, wird es neu erstellt.
            </p>
          )}
          {!busy && eintrag?.warn && (
            <p className="mt-1 text-2xs text-amber-400">Prüfung: {eintrag.warn.map((w) => (/^[-+−]?\d/.test(w) ? `Zahl ${w} steht nicht in den Daten` : w)).join(' · ')}. Dieses Fazit bitte nicht als Handlungsanweisung nutzen.</p>
          )}
        </>
      ) : (
        !busy && <p className="mt-1 text-2xs text-ink-dim">Kira deutet die Zahlen dieses Berichts – je nach Schalter lokal auf deinem Mac oder in der Cloud. Feste Regeln zur Richtung jeder Kennzahl; bei dünner Datenlage ohne Empfehlung.</p>
      )}
    </div>
  );
}
