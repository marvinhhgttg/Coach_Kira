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
const PROMPT_VERSION = 'p4';
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

/** Feste Richtung der Kennzahlen – nur die zum Bericht passenden werden mitgegeben (Stichwort-Filter). */
export const RICHTUNG_ALLE: [RegExp, string][] = [
  [/Schläge\/km|Schläge je km|Puls je km/i, 'Puls je km / Schläge je km: NIEDRIGER = weniger Pulsaufwand je km. Das ist ein Verlaufsindikator, keine direkte Messung der Laufökonomie; Gelände, Tempo und Wetter beeinflussen ihn.'],
  [/je Herzschlag/i, 'Tempo je Herzschlag, Leistung je Herzschlag: HÖHER = besser.'],
  [/Entkopplung/i, 'Aerobe Entkopplung in %: Wie stark das Verhältnis HF zu Tempo/Leistung in der zweiten Hälfte einer langen Einheit gegenüber der ersten Hälfte abdriftet. Es ist KEIN Zeitanteil. NIEDRIGER = besser; unter 5 % = stabile Grundlage. Positive Abweichung = stärker entkoppelt.'],
  [/Watt je km\/h|W je km\/h/i, 'Watt je km/h (Rad, Normalized Power je Tempo): NIEDRIGER = weniger Aufwand fürs gleiche Tempo. Ein Anstieg ist KEINE Effizienzsteigerung; er kann auch an Rad/Untergrund (Gravel vs. Rennrad) liegen.'],
  [/Laufleistung|Laufwatt|Messbruch/i, 'Garmin-Laufleistung (Laufwatt): Seit August 2026 zeigt Garmin bei gleichem Tempo mehr Watt (Messbruch). Aus Laufwatt-Trends über diesen Zeitpunkt hinweg KEINE Schlüsse auf Laufökonomie ziehen.'],
  [/HRV|Schlafscore|Body Battery|Bereitschaft/i, 'HRV relativ zur Basis, Schlafscore, Body Battery, Trainingsbereitschaft: HÖHER = eher erholt – jeweils ein Messwert, kein vollständiges Bild. Garmin rechnet Schlaf und HRV selbst in die Trainingsbereitschaft ein; Zusammenhänge zwischen diesen Größen sind deshalb teils algorithmisch und keine unabhängige Bestätigung.'],
  [/Stress/i, 'Stress (Garmin): NIEDRIGER = besser.'],
  [/kaum|leicht|Effekt/i, 'Effektstärke „kaum“ oder „leicht“: kein klarer Unterschied – das heißt „im Modell nicht nachweisbar“, NICHT „kein Einfluss vorhanden“.'],
  [/CTL|ATL|Form|Woche/i, 'CTL/ATL/Form sind Modellwerte aus der Trainingslast, keine gemessenen Körperzustände. Sinkende CTL = sinkende modellierte Belastung, kein Leistungsverlust. Eine laufende (unvollständige) Woche nicht als Ruhewoche bewerten.'],
  [/Zone|Z1|Z3|Grauzone/i, 'Zonenanteile: immer sagen, ob Anteil an der Last oder an der Zeit. Pyramidale Verteilung ist nicht schlechter als polarisierte; 80/20 ist eine Orientierung, kein Test.'],
  [/FTP/i, 'FTP-Schätzung aus 95 % der besten 20 min ist kein FTP-Test; keine Prozent-Leistungszuwächse daraus ableiten.'],
  [/Kohlenhydrat|Kalorien|Eiweiß|Energie|KH/i, 'Ernährung: Nur ein Teil der Tage ist erfasst; erfasste Tage können systematisch anders sein. Nie ableiten, dass weniger essen die Erholung verbessert.'],
];
export const RICHTUNG = RICHTUNG_ALLE.map(([, t]) => t);
const richtungFuer = (text: string) => RICHTUNG_ALLE.filter(([re]) => re.test(text)).map(([, t]) => t);

/** Wörter/Behauptungen, die ein Fazit ohne Nachweis nie enthalten darf. */
const VERBOTEN: [RegExp, string][] = [
  [/signifikant/i, '„signifikant“ ohne statistischen Test'],
  [/\bbeweis|\bbewiesen|\bbelegt eindeutig|garantiert/i, 'Beweis-Sprache bei Beobachtungsdaten'],
  [/verursach|ist die Ursache|führt zu (besserer|schlechterer)/i, 'Ursachen-Sprache bei Zusammenhängen'],
  [/kein(en)? Einfluss (hat|haben|besteht)|hat keinen Einfluss/i, '„kein Einfluss“ statt „im Modell nicht nachweisbar“'],
];

/** Zuletzt erstelltes Fazit zu genau diesen Fakten (für Druck-Berichte). */
export function fazitAusCache(title: string, facts: string[], data?: unknown, evidenz: 'schwach' | 'mittel' | 'gut' = 'mittel'): string | null {
  return readCache()[idFuer(title, facts, data, evidenz)]?.text ?? null;
}

export function KiraFazit({ title, facts, data, onText, evidenz = 'mittel', regeln, evidenzText }: {
  title: string; facts: string[]; data?: unknown; onText?: (t: string) => void;
  /** schwach = keine Trainingsempfehlung, nur Beobachtung */
  evidenz?: 'schwach' | 'mittel' | 'gut';
  /** zusätzliche, berichtsspezifische Richtungs- und Deutungsregeln */
  regeln?: string[];
  /** fest berechnete Belastbarkeit (wie in der Kachel angezeigt) – wird wörtlich vorgegeben */
  evidenzText?: string;
}) {
  const [engine] = useEngine();
  const id = idFuer(title, [...facts, evidenzText || ''], data, evidenz);
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
    const quelle = facts.join('\n') + '\n' + (evidenzText || '') + '\n' + JSON.stringify(data ?? {});
    const passend = richtungFuer(title + ' ' + quelle);
    const belastbarkeit = evidenzText || (evidenz === 'schwach' ? 'dünne Datenlage – nur Beobachtung' : evidenz === 'gut' ? 'belastbar' : 'mittel – Tendenzen, keine Gewissheit');
    try {
      const out = await localChat(
        [
          {
            role: 'system',
            content: [
              'Du bist Coach Kira, Sportwissenschaftlerin. Du sprichst Marc direkt per Du an. Antworte auf Deutsch, ohne Überschrift, in 3 bis 5 Sätzen.',
              'Regeln (verbindlich):',
              '0. Du ERKLÄRST nur. Was belastbar ist, ist bereits fest berechnet und steht unter „Belastbarkeit“. Beurteile Belastbarkeit und Gruppengrößen nie selbst und widersprich dieser Angabe nicht.',
              '1. Nutze nur Zahlen, die wörtlich in den Fakten stehen. Rechne nichts Neues aus, runde nicht um, extrapoliere nicht.',
              '2. Die Richtung jeder Kennzahl steht unten fest. Deute sie NIE anders herum. Vertausche keine Gruppen: Nenne immer, welche Gruppe welchen Wert hat.',
              '3. Unsicherheit benennen, wie sie in den Fakten markiert ist: Zeilen mit „(wenige …)“, „(zu wenige …)“ oder „Extrapolation“ sind unsicher – weder belegt noch ausgeschlossen, nie „Rauschen“. „kaum“/„leicht“ = kein klarer Unterschied. „nicht auswertbar“ = Messbruch.',
              '3b. Nenne zu jeder Zahl die Anzahl (n) aus DERSELBEN Zeile. Vermische keine Stichproben (z. B. Schlafnächte und Einheiten, Gesamtzahl und Zahl am Rand).',
              '3c. Beantworte die Frage DIESES Berichts. Erwähne keine Kennzahlen, die in den Fakten nicht vorkommen.',
              '4. Aus Beobachtungen werden keine Handlungsregeln: keine Empfehlung zu mehr Intensität, Kadenz- oder Ernährungsänderung oder festen Regeln, wenn die Fakten sie nicht klar stützen. „Im Modell nicht nachweisbar“ ist nicht „kein Einfluss“.',
              '5. Zusammenhang ist keine Ursache – formuliere „war verbunden mit“. Verboten: „signifikant“, „beweist“, „verursacht“.',
              evidenz === 'schwach'
                ? '6. Die Datenlage dieses Berichts ist dünn: Gib KEINE Trainingsempfehlung. Schließe mit „Empfehlung: noch keine – zu wenige Daten.“'
                : '6. Schließe mit genau einer Zeile „Empfehlung: …“ – nur wenn ein Befund klar und belastbar ist. Sonst „Empfehlung: keine – weil …“ mit dem konkreten Grund (zu wenige Daten / Unterschied unter der Rauschgrenze / Messbruch / gemischtes Bild). Große Änderungen (über 5 %) nie als „zu klein“ abtun.',
              '',
              ...(passend.length ? ['Richtung und Bedeutung der Kennzahlen:', ...passend.map((r) => '- ' + r)] : []),
              ...(regeln?.length ? ['', 'Für diesen Bericht zusätzlich:', ...regeln.map((r) => '- ' + r)] : []),
            ].join('\n'),
          },
          {
            role: 'user',
            content: `Bericht: ${title}\nBelastbarkeit (fest berechnet, so übernehmen): ${belastbarkeit}\n\nRegelbasierte Fakten:\n- ${facts.join('\n- ') || 'keine'}\n\nKennzahlen (JSON):\n${JSON.stringify(data ?? {}, null, 1)}\n\nAufgabe: Was zeigen die Zahlen (Richtung beachten), wie belastbar ist das, und was folgt daraus?`,
          },
        ],
        { numCtx: 8192, onToken: setLive, temperature: 0.2 },
      );
      const warn = fremdeZahlen(out, quelle);
      for (const [re, msg] of VERBOTEN) if (re.test(out)) warn.push(msg);
      // Gruppengrößen-Behauptungen („unter 5 Einheiten“) gegen die Fakten prüfen
      for (const m of out.matchAll(/(?:unter|weniger als)\s+(\d+)\s+(Einheiten|Fälle|Spitzen|Tage|Nächte|Läufe)/gi)) {
        if (!quelle.includes(m[0])) warn.push(`„${m[0]}“ steht so nicht in den Daten`);
      }
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
