import { useCallback, useEffect, useMemo, useState } from 'react';
import type { PlanDay } from '../lib/api';
import { fetchPhysBaseline, hrZones, saveExtra, type PhysBaseline } from '../lib/extras';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { downloadFit, type FitIntensity, type FitSport, type FitWorkout } from '../lib/fit';
import { fmtDateShort } from '../lib/format';
import { fetchGarminTage } from '../lib/garmin';
import { lerneModelle, type Modell } from '../lib/essRechner';
import { anZielAnpassen, dauerHinweis, profilFuer, workoutWerte } from '../lib/workoutEss';

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


// Antwort tolerant lesen: das lokale Modell hält sich nicht immer an die Schlüsselnamen
function arr(v: any): any[] {
  if (Array.isArray(v)) return v;
  if (v && typeof v === 'object') {
    const inner = Object.values(v).find((x) => Array.isArray(x));
    if (inner) return inner as any[];
    return [v];
  }
  return [];
}
function pick(o: any, keys: string[]): any {
  for (const k of keys) if (o && o[k] !== undefined && o[k] !== null) return o[k];
  return undefined;
}
const ZONEN = ['Z1', 'Z2', 'Z3', 'Z4', 'Z5'];
function normZone(v: any): Zone {
  const m = String(v ?? '').toUpperCase().match(/Z?\s*([1-5])/);
  return m && ZONEN.includes('Z' + m[1]) ? (('Z' + m[1]) as Zone) : 'offen';
}
function normArt(v: any): string {
  const t = String(v ?? '').toLowerCase();
  if (/warm|aufw|einroll|einlauf/.test(t)) return 'warmup';
  if (/cool|ausroll|auslauf|abw/.test(t)) return 'cooldown';
  if (/interv|tempo|belast|work|hart/.test(t)) return 'interval';
  if (/recov|erhol|locker|pause/.test(t)) return 'recovery';
  if (/rest|ruhe/.test(t)) return 'rest';
  return 'active';
}
function normMin(v: any): number {
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object') {
    const val = Number(v.value ?? v.wert ?? v.minuten ?? v.minutes ?? 0);
    return /h|stunde/i.test(String(v.unit ?? v.einheit ?? '')) ? val * 60 : /s(ec|ek)/i.test(String(v.unit ?? '')) ? val / 60 : val;
  }
  const t = String(v ?? '');
  const h = t.match(/(\d+(?:[.,]\d+)?)\s*h/);
  const m = t.match(/(\d+(?:[.,]\d+)?)/);
  if (h) return Number(h[1].replace(',', '.')) * 60;
  return m ? Number(m[1].replace(',', '.')) : 0;
}
function normalizeWorkout(raw: any): KiWorkout | null {
  if (!raw || typeof raw !== 'object') return null;
  const root = raw.workout && typeof raw.workout === 'object' ? raw.workout : raw;
  let blocks = arr(pick(root, ['bloecke', 'blöcke', 'blocks', 'block', 'abschnitte', 'phasen']));
  // nur Schritte ohne Blöcke geliefert
  if (!blocks.length) {
    const steps = arr(pick(root, ['schritte', 'steps', 'intervalle', 'einheiten']));
    if (steps.length) blocks = [{ wiederholungen: 1, schritte: steps }];
  }
  const out: KiBlock[] = blocks
    .map((b: any) => {
      const st = arr(pick(b, ['schritte', 'steps', 'intervalle'])).length ? arr(pick(b, ['schritte', 'steps', 'intervalle'])) : pick(b, ['minuten', 'minutes', 'dauer', 'duration']) !== undefined ? [b] : [];
      return {
        wiederholungen: Math.max(1, Math.round(Number(pick(b, ['wiederholungen', 'repeat', 'repeats', 'wdh', 'anzahl'])) || 1)),
        schritte: st
          .map((x: any) => ({
            name: String(pick(x, ['name', 'titel', 'bezeichnung']) || 'Schritt').slice(0, 15),
            minuten: Math.round(normMin(pick(x, ['minuten', 'minutes', 'dauer', 'duration', 'min'])) * 10) / 10,
            art: normArt(`${pick(x, ['art', 'typ', 'type', 'intensity']) ?? ''} ${pick(x, ['name', 'titel']) ?? ''}`),
            zone: normZone(pick(x, ['zone', 'hf_zone', 'hr_zone', 'intensitaet']) ?? x?.target?.zone ?? x?.ziel?.zone),
            hinweis: String(pick(x, ['hinweis', 'notiz', 'note', 'notes', 'beschreibung']) || '').slice(0, 30),
          }))
          .filter((x: KiStep) => x.minuten > 0),
      };
    })
    .filter((b: KiBlock) => b.schritte.length);
  if (!out.length) return null;
  return { name: String(pick(root, ['name', 'titel', 'title']) || 'Kira-Workout'), beschreibung: String(pick(root, ['beschreibung', 'description', 'ziel']) || ''), bloecke: out };
}

export function WorkoutPanel({ day, onClose }: { day: PlanDay; onClose: () => void }) {
  const [phys, setPhys] = useState<PhysBaseline | null>(null);
  const [ki, setKi] = useState<KiWorkout | null>(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [modell, setModell] = useState<Modell | null>(null);
  const [anpassung, setAnpassung] = useState<{ vorherMin: number; vorherEss: number } | null>(null);
  const sport = detectSport(day);
  const zielEss = Math.round(Number(day.recommended_load_ess ?? 0));
  const zones = useMemo(() => (phys ? hrZones(phys.lthr, sport === 'cycling' ? 'bike' : 'run') : []), [phys, sport]);
  const zoneHr = (z: Zone) => zones.find((x) => x.zone === z) || null;

  const generate = useCallback(async () => {
    setBusy(true);
    setErr(null);
    setInfo(null);
    try {
      const p = phys || (await fetchPhysBaseline());
      if (!phys) setPhys(p);
      const zonesNow = hrZones(p.lthr, sport === 'cycling' ? 'bike' : 'run');
      // persönliches ESS-Modell (Garmin-Einheiten) für das Verhältnis Last ↔ Dauer
      let m = modell;
      if (!m) {
        const pid = profilFuer(sport, `${day.training_1 || ''}`);
        if (pid) {
          const tage = await fetchGarminTage();
          m = lerneModelle(tage).find((x) => x.id === pid && x.n >= 8) || null;
          setModell(m);
        }
      }
      const dh = m && zielEss > 0 ? dauerHinweis(m, zielEss, zonesNow, String(day.recommended_zone || 'Z2')) : null;
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
        dh
          ? `- GESAMTDAUER inkl. Aufwärmen, Pausen und Ausrollen: etwa ${dh.min} min (±5). Das ergibt bei Ø HF um ${dh.hf} bpm deine Ziel-Last von ${zielEss} ESS – gemessen an deinen eigenen Garmin-Einheiten. Mehr Intensität = kürzer, nicht länger.`
          : '- Dauer passend zur Last: grob 1 ESS ≈ 0,4 min bei Z2, intensiver = kürzer.',
        '- Immer Aufwärmen (warmup) und Ausrollen (cooldown).',
        '- Intervalle als eigener Block mit wiederholungen > 1 (Belastung + Erholung).',
        '- Schrittnamen max. 15 Zeichen, hinweis max. 30 Zeichen, Deutsch.',
        '- Bei 0 ESS: kurze Mobility-Einheit, 20 min, Zone offen.',
        '',
        'Antworte exakt in diesem JSON-Format (Schlüssel genau so):',
        '{"name":"Gravel Z2","beschreibung":"1 Satz","bloecke":[{"wiederholungen":1,"schritte":[{"name":"Einrollen","minuten":15,"art":"warmup","zone":"Z1","hinweis":""}]},{"wiederholungen":4,"schritte":[{"name":"Tempo","minuten":8,"art":"interval","zone":"Z3","hinweis":""},{"name":"Locker","minuten":4,"art":"recovery","zone":"Z2","hinweis":""}]},{"wiederholungen":1,"schritte":[{"name":"Ausrollen","minuten":10,"art":"cooldown","zone":"Z1","hinweis":""}]}]}',
        'Erlaubt: art = warmup|active|interval|recovery|cooldown|rest, zone = Z1|Z2|Z3|Z4|Z5|offen.',
      ].join('\n');
      const text = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Du baust strukturierte Garmin-Workouts. Antworte nur mit JSON. Deutsch, keine Emojis.' },
          { role: 'user', content: user },
        ],
        { temperature: 0.3, numCtx: 8192 },
      );
      const j = normalizeWorkout(parseJsonLoose<any>(text));
      if (!j) {
        console.warn('Workout-Antwort nicht lesbar:', text);
        throw new Error('Kira hat kein gültiges Workout geliefert. Bitte „Neu“ versuchen.');
      }
      // deterministisch an die Ziel-Last anpassen (Kira hält Dauerangaben nicht immer ein)
      if (m && zielEss > 0) {
        const r = anZielAnpassen(j, zonesNow, m, zielEss);
        setAnpassung(r.angepasst ? { vorherMin: Math.round(r.vorher.min), vorherEss: Math.round(r.vorher.ess) } : null);
        setKi(r.workout);
      } else {
        setAnpassung(null);
        setKi(j);
      }
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, [day, phys, modell, sport, zielEss]);

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
  const essSchaetzung = ki && modell && zones.length ? workoutWerte(ki, zones, modell) : null;

  const asText = () => {
    if (!ki) return '';
    const lines = [`${ki.name} · ${fmtDateShort(day.date)} · ${SPORT_LABEL[sport]} · ${Math.round(total)} min${essSchaetzung ? ` · ≈ ${Math.round(essSchaetzung.ess)} ESS (Ziel ${zielEss})` : ''}`];
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
          <div className="mt-1 text-2xs text-ink-dim tnum" data-testid="workout-ess">
            Gesamtdauer {Math.round(total)} min
            {essSchaetzung && <> · geschätzte Last ≈ {Math.round(essSchaetzung.ess)} ESS bei Ø HF {Math.round(essSchaetzung.hf)} (Ziel {zielEss} ESS)</>}
          </div>
          {anpassung && (
            <div className="text-2xs text-ink-dim tnum">
              An deine Ziel-Last angepasst: Kiras Vorschlag hatte {anpassung.vorherMin} min ≈ {anpassung.vorherEss} ESS. Grundlage: dein ESS-Modell aus {modell?.n} eigenen Garmin-Einheiten.
            </div>
          )}
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
