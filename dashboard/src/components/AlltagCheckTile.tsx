import { useMemo, useState } from 'react';
import { fetchCalendarDays, fetchWeather, saveExtra, type CalendarDay, type WeatherDay } from '../lib/extras';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { fmtDateShort } from '../lib/format';

export interface AlltagSwap { a: string; b: string; grund: string }
interface Hinweis { datum: string; text: string }
interface KiErgebnis { fazit: string; tausch: AlltagSwap[]; hinweise: Hinweis[] }

const SCHEMA = {
  type: 'object',
  properties: {
    fazit: { type: 'string' },
    tausch: {
      type: 'array',
      items: { type: 'object', properties: { a: { type: 'string' }, b: { type: 'string' }, grund: { type: 'string' } }, required: ['a', 'b', 'grund'] },
    },
    hinweise: {
      type: 'array',
      items: { type: 'object', properties: { datum: { type: 'string' }, text: { type: 'string' } }, required: ['datum', 'text'] },
    },
  },
  required: ['fazit', 'tausch', 'hinweise'],
};

const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const wd = (d: string) => WT[new Date(d + 'T12:00:00').getDay()];
const localToday = () => new Date().toLocaleDateString('sv-SE');

function calText(c?: CalendarDay): string {
  if (!c) return '–';
  const p: string[] = [];
  if (c.reise) p.push('Reise');
  if (c.frei) p.push('Urlaub/frei');
  if (c.gesundheit) p.push('Arzt/Physio');
  const n = c.termine;
  if (n) p.push(`${n} Termin${n === 1 ? '' : 'e'}`);
  if (c.belegt_h > 0) p.push(`${String(c.belegt_h).replace('.', ',')} h belegt`);
  if (c.frueh) p.push('früh');
  if (c.abends) p.push('abends');
  return p.length ? p.join(' · ') : 'frei';
}

function wetText(w?: WeatherDay): string {
  if (!w) return '–';
  return `${w.tmin}–${w.tmax}° · ${String(w.regen_mm).replace('.', ',')} mm (${w.regen_pct} %) · Wind ${w.wind_kmh} · ${w.text}`;
}

export function AlltagCheckTile({
  dates,
  loads,
  sports,
  zones,
  locks,
  onApply,
}: {
  dates: string[];
  loads: number[];
  sports: string[];
  zones: string[];
  locks: boolean[];
  onApply: (swaps: AlltagSwap[]) => void;
}) {
  const [cal, setCal] = useState<Record<string, CalendarDay> | null>(null);
  const [wet, setWet] = useState<Record<string, WeatherDay> | null>(null);
  const [notes, setNotes] = useState<string[]>([]);
  const [phase, setPhase] = useState<'idle' | 'laden' | 'ki'>('idle');
  const [err, setErr] = useState<string | null>(null);
  const [res, setRes] = useState<KiErgebnis | null>(null);
  const [sel, setSel] = useState<boolean[]>([]);
  const [info, setInfo] = useState<string | null>(null);

  const rows = useMemo(() => {
    const t = localToday();
    return dates
      .map((d, i) => ({ d, i, load: Math.round(loads[i] || 0), sport: sports[i] || '', zone: zones[i] || '', fest: !!locks[i] }))
      .filter((r) => r.d >= t)
      .slice(0, 14);
  }, [dates, loads, sports, zones, locks]);

  async function run() {
    setErr(null);
    setRes(null);
    setInfo(null);
    setPhase('laden');
    const n: string[] = [];
    const [c, w] = await Promise.allSettled([fetchCalendarDays(14), fetchWeather(14)]);
    const calMap: Record<string, CalendarDay> = {};
    const wetMap: Record<string, WeatherDay> = {};
    if (c.status === 'fulfilled') c.value.forEach((x) => (calMap[x.date] = x));
    else n.push(/JSON/.test(String(c.reason)) ? 'Kalender nicht verfügbar (Patch „KI-Extras“ noch nicht aktiv)' : 'Kalender nicht verfügbar');
    if (w.status === 'fulfilled') w.value.forEach((x) => (wetMap[x.date] = x));
    else n.push('Wetter nicht verfügbar');
    setCal(c.status === 'fulfilled' ? calMap : null);
    setWet(w.status === 'fulfilled' ? wetMap : null);
    setNotes(n);
    if (c.status !== 'fulfilled' && w.status !== 'fulfilled') {
      setPhase('idle');
      setErr('Weder Kalender noch Wetter verfügbar – nichts zu prüfen.');
      return;
    }

    setPhase('ki');
    try {
      const table = rows
        .map((r) =>
          `${r.d} ${wd(r.d)} | Plan: ${r.load} ESS ${r.sport || '-'} ${r.zone || ''}${r.fest ? ' FEST' : ''} | Kalender: ${
            c.status === 'fulfilled' ? calText(calMap[r.d]) : 'unbekannt'
          } | Wetter: ${w.status === 'fulfilled' ? wetText(wetMap[r.d]) : 'unbekannt'}`,
        )
        .join('\n');
      const user = [
        'Trainingsplan der nächsten Tage mit Kalender-Belegung und Wetter (Kadeltshofen):',
        table,
        '',
        'Regeln:',
        '- Schlage höchstens 3 Tausche zwischen zwei Tagen vor (a und b = Datum YYYY-MM-DD aus der Tabelle).',
        '- Niemals einen FEST-Tag tauschen.',
        '- Lange oder intensive Einheiten auf Tage mit wenig Terminen und gutem Wetter (Rad/Lauf draußen).',
        '- An Reisetagen und sehr vollen Tagen (> 9 h belegt) eher Ruhetag oder kurze Einheit.',
        '- Regen > 5 mm oder Wind > 40 km/h: Rad draußen vermeiden; als Hinweis Indoor oder Lauf vorschlagen.',
        '- Wenn alles passt: tausch leer lassen.',
        '- fazit: 1–2 Sätze. hinweise: nur echte Auffälligkeiten, je max. 1 Satz.',
        '',
        'Antworte exakt in diesem JSON-Format (Schlüssel genau so, tausch und hinweise sind Listen):',
        '{"fazit":"…","tausch":[{"a":"2026-10-24","b":"2026-10-25","grund":"…"}],"hinweise":[{"datum":"2026-10-25","text":"…"}]}',
      ].join('\n');
      const text = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Du prüfst, ob der Trainingsplan zu Kalender und Wetter passt. Antworte nur mit JSON, Deutsch, du-Form, keine Emojis.' },
          { role: 'user', content: user },
        ],
        { temperature: 0.2, numCtx: 12288 },
      );
      const raw = parseJsonLoose<any>(text) || {};
      const list = (v: any): any[] => (Array.isArray(v) ? v : v && typeof v === 'object' ? (Object.values(v).find((x) => Array.isArray(x)) as any[]) || [v] : []);
      const dt = (v: any) => {
        const m = String(v ?? '').match(/(\d{4}-\d{2}-\d{2})/);
        return m ? m[1] : '';
      };
      const j: KiErgebnis = {
        fazit: String(raw.fazit ?? raw.zusammenfassung ?? raw.summary ?? ''),
        tausch: list(raw.tausch ?? raw.tausche ?? raw.swaps ?? raw.vorschlaege)
          .map((x: any) => ({ a: dt(x?.a ?? x?.von ?? x?.tag1 ?? x?.from), b: dt(x?.b ?? x?.nach ?? x?.tag2 ?? x?.to), grund: String(x?.grund ?? x?.begruendung ?? x?.reason ?? '') }))
          .filter((x) => x.a && x.b),
        hinweise: list(raw.hinweise ?? raw.notes ?? raw.hinweis)
          .map((x: any) => (typeof x === 'string' ? { datum: dt(x), text: x } : { datum: dt(x?.datum ?? x?.date), text: String(x?.text ?? x?.hinweis ?? '') }))
          .filter((x) => x.text),
      };
      const byDate = new Map(rows.map((r) => [r.d, r]));
      const tausch = j.tausch.filter((s) => {
        const a = byDate.get(String(s.a));
        const b = byDate.get(String(s.b));
        return a && b && a.d !== b.d && !a.fest && !b.fest;
      });
      const verworfen = j.tausch.length - tausch.length;
      const r: KiErgebnis = { fazit: j.fazit || 'Kira hat kein Fazit geliefert.', tausch, hinweise: j.hinweise };
      setRes(r);
      setSel(tausch.map(() => true));
      if (verworfen > 0) setNotes((x) => [...x, `${verworfen} Vorschlag/Vorschläge verworfen (gesperrter oder unbekannter Tag).`]);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setPhase('idle');
    }
  }

  async function onSave() {
    if (!res) return;
    try {
      const txt = [res.fazit, ...res.tausch.map((s) => `- ${fmtDateShort(s.a)} ↔ ${fmtDateShort(s.b)}: ${s.grund}`), ...res.hinweise.map((h) => `- ${fmtDateShort(h.datum)}: ${h.text}`)].join('\n');
      await saveExtra('alltag', localToday(), txt, res);
      setInfo('Gespeichert.');
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div className="panel-raised p-4">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Alltag-Check</div>
          <p className="text-2xs text-ink-dim mt-1">Passt der Plan zu deinen Terminen und zum Wetter? Kalender ohne Termintitel, Wetter für Kadeltshofen.</p>
        </div>
        <button className="btn btn-primary text-xs px-2 py-1" onClick={run} disabled={phase !== 'idle' || !rows.length}>
          {phase === 'laden' ? 'Lade Kalender und Wetter…' : phase === 'ki' ? 'Kira prüft…' : 'Prüfen'}
        </button>
      </div>
      {notes.map((n, i) => (
        <p key={i} className="mt-1 text-2xs text-orange-300">{n}</p>
      ))}
      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}

      {(cal || wet) && (
        <div className="mt-3 overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="border-b border-border text-ink-dim">
                <th className="py-1 pr-3 font-normal">Tag</th>
                <th className="py-1 pr-3 font-normal">Plan</th>
                <th className="py-1 pr-3 font-normal">Kalender</th>
                <th className="py-1 font-normal">Wetter</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.d} className="border-b border-border/60">
                  <td className="py-1 pr-3 tnum whitespace-nowrap">{wd(r.d)} {fmtDateShort(r.d)}</td>
                  <td className="py-1 pr-3 whitespace-nowrap">
                    <span className="tnum">{r.load} ESS</span> <span className="text-ink-muted">{r.sport} {r.zone}</span>
                    {r.fest && <span className="ml-1 chip border border-border-strong text-2xs text-ink-muted">fest</span>}
                  </td>
                  <td className="py-1 pr-3 text-ink-muted">{cal ? calText(cal[r.d]) : '–'}</td>
                  <td className="py-1 text-ink-muted">{wet ? wetText(wet[r.d]) : '–'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {res && (
        <div className="mt-3 flex flex-col gap-2">
          <p className="text-sm text-ink">{res.fazit}</p>
          {res.tausch.length > 0 ? (
            <div className="flex flex-col gap-1">
              {res.tausch.map((s, i) => (
                <label key={i} className="flex items-start gap-2 text-xs">
                  <input type="checkbox" checked={!!sel[i]} onChange={(e) => setSel((x) => x.map((v, k) => (k === i ? e.target.checked : v)))} />
                  <span>
                    <span className="tnum text-ink">{wd(s.a)} {fmtDateShort(s.a)} ↔ {wd(s.b)} {fmtDateShort(s.b)}</span>
                    <span className="text-ink-muted"> – {s.grund}</span>
                  </span>
                </label>
              ))}
            </div>
          ) : (
            <p className="text-xs text-ink-muted">Kein Tausch nötig.</p>
          )}
          {res.hinweise.length > 0 && (
            <ul className="text-xs text-ink-muted list-disc pl-5">
              {res.hinweise.map((h, i) => (
                <li key={i}>{h.datum ? `${fmtDateShort(h.datum)}: ` : ''}{h.text}</li>
              ))}
            </ul>
          )}
          <div className="flex flex-wrap gap-2">
            {res.tausch.length > 0 && (
              <button className="btn btn-primary text-xs px-2 py-1" onClick={() => onApply(res.tausch.filter((_, i) => sel[i]))} disabled={!sel.some(Boolean)}>
                Ausgewählte Tausche übernehmen
              </button>
            )}
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={onSave}>Speichern</button>
            {info && <span className="text-2xs text-ink-dim self-center">{info}</span>}
          </div>
          <p className="text-2xs text-ink-dim">Übernommene Tausche landen in den Slidern; gespeichert wird erst mit „In timeline speichern“.</p>
        </div>
      )}
    </div>
  );
}
