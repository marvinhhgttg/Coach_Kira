import { useCallback, useEffect, useMemo, useState } from 'react';
import type { PlanDay } from '../lib/api';
import { fetchPhysBaseline, hrZones, saveExtra, type PhysBaseline } from '../lib/extras';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { downloadFit, type FitIntensity, type FitSport, type FitWorkout } from '../lib/fit';
import { fmtDateShort } from '../lib/format';

type Zone = 'Z1' | 'Z2' | 'Z3' | 'Z4' | 'Z5' | 'offen';
interface KiStep { name: string; minuten: number; art: string; zone: Zone; hinweis?: string }
interface KiBlock { wiederholungen: number; schritte: KiStep[] }
interface KiWorkout { name: string; beschreibung?: string; bloecke: KiBlock[] }

const SCHEMA = {
  type: 'object',
  properties: {
    name: { type: 'string' },
    beschreibung: { type: 'string' },
    bloecke: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          wiederholungen: { type: 'integer' },
          schritte: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string' },
                minuten: { type: 'number' },
                art: { type: 'string', enum: ['warmup', 'active', 'interval', 'recovery', 'cooldown', 'rest'] },
                zone: { type: 'string', enum: ['Z1', 'Z2', 'Z3', 'Z4', 'Z5', 'offen'] },
                hinweis: { type: 'string' },
              },
              required: ['name', 'minuten', 'art', 'zone'],
            },
          },
        },
        required: ['wiederholungen', 'schritte'],
      },
    },
  },
  required: ['name', 'bloecke'],
};

function detectSport(day: PlanDay): FitSport {
  const t = `${day.training_1 || ''}`.toLowerCase();
  if (/bike|rad|gravel|ride|commute|cycl/.test(t)) return 'cycling';
  if (/run|lauf|hiit/.test(t)) return 'running';
  if (/hike|wander|walk|spazier/.test(t)) return 'hiking';
  return 'training';
}
const SPORT_LABEL: Record<FitSport, string> = {
  cycling: 'Rad', running: 'Laufen', hiking: 'Wandern', walking: 'Gehen', training: 'Training', generic: 'Training',
};
const ART: Record<string, FitIntensity> = {
  warmup: 'warmup', active: 'active', interval: 'interval', recovery: 'recovery', cooldown: 'cooldown', rest: 'rest',
};

function blockMinutes(b: KiBlock) {
  return b.schritte.reduce((s, x) => s + (Number(x.minuten) || 0), 0) * Math.max(1, b.wiederholungen || 1);
}

export function WorkoutPanel({ day, onClose }: { day: PlanDay; onClose: () => void }) {
  const [phys, setPhys] = useState<PhysBaseline | null>(null);
  const [ki, setKi] = useState<KiWorkout | null>(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const sport = detectSport(day);
  const zones = useMemo(() => (phys ? hrZones(phys.lthr, sport === 'cycling' ? 'bike' : 'run') : []), [phys, sport]);
  const zoneHr = (z: Zone) => zones.find((x) => x.zone === z) || null;

  const generate = useCallback(async () => {
    setBusy(true);
    setErr(null);
    setInfo(null);
    try {
      const p = phys || (await fetchPhysBaseline());
      if (!phys) setPhys(p);
      const user = [
        `Datum: ${day.date} (${day.day || ''})`,
        `Training: ${day.training_1 || '-'}${day.training_2 ? ` (Alternative: ${day.training_2})` : ''}`,
        `Zone laut Plan: ${day.recommended_zone || 'offen'}`,
        `Ziel-Last: ${Math.round(Number(day.recommended_load_ess ?? 0))} ESS`,
        `Erwartete Wirkung: ${day.projected_effect || '-'}`,
        `Wetter: ${day.weather_recommendation || '-'}`,
        `Laktatschwelle ${p.lthr} bpm, Max-HF ${p.maxHr} bpm, FTP ${p.ftp} W.`,
        '',
        'Baue ein strukturiertes Workout:',
        '- Dauer passend zur Last: grob 1 ESS ≈ 0,4–0,6 min bei Z2, intensiver = kürzer.',
        '- Immer Aufwärmen (warmup) und Ausrollen (cooldown).',
        '- Intervalle als eigener Block mit wiederholungen > 1 (Belastung + Erholung).',
        '- Schrittnamen max. 15 Zeichen, hinweis max. 30 Zeichen, Deutsch.',
        '- Bei 0 ESS: kurze Mobility-Einheit, 20 min, Zone offen.',
      ].join('\n');
      const text = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Du baust strukturierte Garmin-Workouts. Antworte nur mit JSON. Deutsch, keine Emojis.' },
          { role: 'user', content: user },
        ],
        { format: SCHEMA, temperature: 0.3, numCtx: 8192 },
      );
      const j = parseJsonLoose<KiWorkout>(text);
      if (!j || !Array.isArray(j.bloecke) || !j.bloecke.length) throw new Error('Kira hat kein gültiges Workout geliefert.');
      j.bloecke = j.bloecke
        .map((b) => ({ wiederholungen: Math.max(1, Math.round(Number(b.wiederholungen) || 1)), schritte: (b.schritte || []).filter((s) => Number(s.minuten) > 0) }))
        .filter((b) => b.schritte.length);
      setKi(j);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, [day, phys]);

  useEffect(() => {
    generate();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [day.date]);

  const fit: FitWorkout | null = useMemo(() => {
    if (!ki) return null;
    return {
      name: ki.name || `Kira ${day.date}`,
      sport,
      blocks: ki.bloecke.map((b) => ({
        repeat: b.wiederholungen,
        steps: b.schritte.map((s) => {
          const hr = zoneHr(s.zone);
          return {
            name: String(s.name || '').slice(0, 15),
            minutes: Number(s.minuten) || 0,
            intensity: ART[s.art] || 'active',
            hrLow: hr ? hr.low : null,
            hrHigh: hr ? hr.high : null,
            notes: String(s.hinweis || '').slice(0, 30),
          };
        }),
      })),
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ki, zones, sport, day.date]);

  const total = ki ? ki.bloecke.reduce((s, b) => s + blockMinutes(b), 0) : 0;

  const asText = () => {
    if (!ki) return '';
    const lines = [`${ki.name} · ${fmtDateShort(day.date)} · ${SPORT_LABEL[sport]} · ${Math.round(total)} min`];
    if (ki.beschreibung) lines.push(ki.beschreibung);
    ki.bloecke.forEach((b) => {
      const pre = b.wiederholungen > 1 ? `${b.wiederholungen} × ` : '';
      b.schritte.forEach((s, i) => {
        const hr = zoneHr(s.zone);
        lines.push(
          `${i === 0 ? pre : b.wiederholungen > 1 ? '    ' : ''}- ${s.name}: ${s.minuten} min, ${s.zone}${hr ? ` (${hr.low}–${hr.high} bpm)` : ''}${s.hinweis ? ` – ${s.hinweis}` : ''}`,
        );
      });
    });
    return lines.join('\n');
  };

  async function onSave() {
    try {
      await saveExtra('workout', day.date, asText(), { workout: ki, sport, quelle: 'cockpit' });
      setInfo('Gespeichert.');
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div className="mt-3 rounded border border-accent-strong/30 bg-bg-subtle/40 p-3">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="text-sm font-semibold text-ink">
          Workout · {day.day} {fmtDateShort(day.date)} · {SPORT_LABEL[sport]}
        </div>
        <button className="btn btn-ghost text-xs px-2 py-1" onClick={onClose}>Schließen</button>
      </div>
      {busy && <p className="mt-2 text-xs text-ink-muted">Kira baut das Workout…</p>}
      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}
      {ki && !busy && (
        <>
          <div className="mt-2 text-sm text-ink">{ki.name}</div>
          {ki.beschreibung && <p className="text-xs text-ink-muted">{ki.beschreibung}</p>}
          <div className="mt-1 text-2xs text-ink-dim tnum">Gesamtdauer {Math.round(total)} min</div>
          <div className="mt-2 flex flex-col gap-2">
            {ki.bloecke.map((b, bi) => (
              <div key={bi} className="rounded border border-border p-2">
                {b.wiederholungen > 1 && <div className="text-2xs text-accent mb-1">{b.wiederholungen} × wiederholen</div>}
                {b.schritte.map((s, si) => {
                  const hr = zoneHr(s.zone);
                  return (
                    <div key={si} className="flex flex-wrap gap-x-3 text-xs">
                      <span className="text-ink min-w-[8rem]">{s.name}</span>
                      <span className="tnum text-ink-muted">{s.minuten} min</span>
                      <span className="tnum text-ink-muted">{s.zone}{hr ? ` · ${hr.low}–${hr.high} bpm` : ''}</span>
                      {s.hinweis && <span className="text-ink-dim">{s.hinweis}</span>}
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
          <div className="mt-3 flex flex-wrap gap-2">
            <button className="btn btn-primary text-xs px-2 py-1" onClick={() => fit && downloadFit(fit, `kira-${day.date}`)}>FIT herunterladen</button>
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => navigator.clipboard.writeText(asText()).then(() => setInfo('Kopiert.'))}>Als Text kopieren</button>
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={onSave}>Speichern</button>
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={generate}>Neu</button>
            {info && <span className="text-2xs text-ink-dim self-center">{info}</span>}
          </div>
          <p className="mt-2 text-2xs text-ink-dim">
            FIT-Datei in Garmin Connect unter Training › Workouts › Importieren hochladen und dann an die Uhr senden. Herzfrequenzbereiche aus deiner Laktatschwelle ({phys?.lthr} bpm).
          </p>
        </>
      )}
    </div>
  );
}
