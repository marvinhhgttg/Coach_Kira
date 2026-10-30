import { useState, useEffect } from 'react';
import { fetchExtras, saveExtra, deleteExtra } from '../lib/extras';
import type { ExtraItem } from '../lib/extras';

export const EREIGNIS_ARTEN = [
  { v: 'krank', l: 'krank' },
  { v: 'reise', l: 'Reise' },
  { v: 'hoehe', l: 'Höhe' },
  { v: 'hitze', l: 'Hitze' },
  { v: 'stress', l: 'Job-Stress' },
  { v: 'alkohol', l: 'Alkohol' },
  { v: 'feiertag', l: 'Feiertag' },
  { v: 'sonstiges', l: 'Sonstiges' },
] as const;

export interface EreignisDaten {
  arten: string[];
  notiz?: string;
}

function fmtDatum(ymd: string): string {
  const p = ymd.split('-');
  return `${p[2]}.${p[1]}`;
}

export function EreignisTile() {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [events, setEvents] = useState<ExtraItem<EreignisDaten>[]>([]);
  const [datum, setDatum] = useState(() => new Date().toISOString().slice(0, 10));
  const [selectedArten, setSelectedArten] = useState<string[]>([]);
  const [notiz, setNotiz] = useState('');

  function laden() {
    setLoading(true);
    fetchExtras<EreignisDaten>('ereignis', 60)
      .then(r => {
        // Absteigend nach datum sortieren (Extras kommen oft in beliebiger Reihenfolge)
        const sorted = r.sort((a, b) => (b.datum > a.datum ? 1 : b.datum < a.datum ? -1 : 0));
        setEvents(sorted);
        setLoading(false);
      })
      .catch(e => {
        const msg = e.message || '';
        if (/Typ|JSON|Failed to fetch|Load failed/i.test(msg)) {
          setError('Patch „Ereignis-Marker" noch nicht aktiv.');
        } else {
          setError(msg);
        }
        setLoading(false);
      });
  }

  useEffect(() => { laden(); }, []);

  const labelMap = new Map<string, string>(EREIGNIS_ARTEN.map((a) => [a.v, a.l]));
  const text = selectedArten.map(v => labelMap.get(v) ?? v).join(', ');
  const submitText = text + (notiz ? ' – ' + notiz : '');

  function handleSubmit() {
    if (selectedArten.length === 0) return;
    saveExtra('ereignis', datum, submitText, { arten: selectedArten, notiz })
      .then(() => {
        setDatum(new Date().toISOString().slice(0, 10));
        setSelectedArten([]);
        setNotiz('');
        laden();
      })
      .catch(e => {
        const msg = e.message || '';
        if (/Typ|JSON|Failed to fetch|Load failed/i.test(msg)) {
          setError('Patch „Ereignis-Marker" noch nicht aktiv.');
        } else {
          setError(msg);
        }
      });
  }

  function toggleArt(art: string) {
    setSelectedArten(prev =>
      prev.includes(art) ? prev.filter(a => a !== art) : [...prev, art]
    );
  }

  function loescheEvent(d: string) {
    deleteExtra('ereignis', d).then(laden);
  }

  if (loading) return <div className="panel-raised p-4">Lade Ereignisse …</div>;

  return (
    <div className="panel-raised p-4">
      <div>
        <div className="label">Coach Kira · Ereignisse</div>
        <p className="text-2xs text-ink-dim mt-1">
          Markiere besondere Tage. Krankheitstage nimmt Kira aus den Lernmodellen heraus; nach einigen Episoden lässt sich die Frühwarnung prüfen.
        </p>
      </div>

      <form className="mt-3" onSubmit={e => { e.preventDefault(); handleSubmit(); }}>
        <div className="flex flex-col gap-3">
          <div>
            <label className="text-xs text-ink-dim block mb-1">Datum</label>
            <input type="date" value={datum} onChange={e => setDatum(e.target.value)} className="input text-xs w-full" />
          </div>

          <div>
            <label className="text-xs text-ink-dim block mb-1">Art</label>
            <div className="flex flex-wrap gap-1">
              {EREIGNIS_ARTEN.map(a => (
                <button
                  key={a.v}
                  type="button"
                  onClick={() => toggleArt(a.v)}
                  className={selectedArten.includes(a.v) ? 'chip border border-accent text-ink' : 'chip border border-border text-ink-muted'}
                >
                  {a.l}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label className="text-xs text-ink-dim block mb-1">Notiz (optional)</label>
            <input type="text" value={notiz} maxLength={120} onChange={e => setNotiz(e.target.value)} placeholder="Freitext …" className="input text-xs w-full" />
          </div>

          <button type="submit" disabled={selectedArten.length === 0} className="btn btn-primary text-xs px-2 py-1 self-start disabled:opacity-40">
            Speichern
          </button>
        </div>
      </form>

      {error && <p className="mt-3 text-orange-300 text-xs">{error}</p>}

      <div className="mt-4">
        {events.length === 0 && !error && <p className="text-2xs text-ink-muted">Noch keine Ereignisse.</p>}
        {events.map((ev, i) => (
          <div key={i} className="flex items-center justify-between py-1 border-b border-border last:border-none">
            <div>
              <span className="text-xs tnum">{fmtDatum(ev.datum)}</span>{' '}
              <span className="text-xs text-ink">{ev.text}</span>
            </div>
            <button type="button" onClick={() => loescheEvent(ev.datum)} className="btn btn-ghost text-2xs">
              Löschen
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}
