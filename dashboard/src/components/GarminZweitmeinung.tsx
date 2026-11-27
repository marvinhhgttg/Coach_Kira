/**
 * 27.11.2026: Garmin im Tagesvotum – Signale (Erholungszeit, Trainingszustand, Belastungsfokus) und
 * Garmins Tagesvorschlag als Zweitmeinung. Die Rechnung macht der Server (server/widget_cli.ts über
 * kern/garminSignale.ts) – hier nur Anzeige und Eintrag des Vorschlags (Garmin liefert ihn nicht aus).
 */
import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { VORSCHLAG_LABEL, VORSCHLAG_TYPEN, VORSCHLAG_ZONE, type VorschlagTyp } from '../kern/garminSignale';
import type { Vote } from '../kern/votumRegeln';

type Garmin = {
  erholH: number | null; status: string | null; statusText: string | null; deckel: string[]; fokus: string | null;
  vorschlag: { typ: VorschlagTyp; min: number; zone: string | null; ess: number | null; text: string | null; urteil: string | null } | null;
} | null;

async function j(r: Response) {
  const t = await r.text();
  try { return JSON.parse(t); } catch { return { ok: false, error: `HTTP ${r.status}` }; }
}

export default function GarminZweitmeinung({ vote }: { vote: Vote | null }) {
  const [g, setG] = useState<Garmin>(null);
  const [offen, setOffen] = useState(false);
  const [typ, setTyp] = useState<VorschlagTyp>('erholung');
  const [min, setMin] = useState('30');
  const [zone, setZone] = useState('');
  const [msg, setMsg] = useState<string | null>(null);
  const laden = useCallback(async () => {
    try {
      const r = await j(await fetch(`${SERVER_BASE}/api/widget/abgleich?cb=${Date.now()}`, { cache: 'no-store' }));
      if (r.ok) setG(r.garmin ?? null);
    } catch {}
  }, []);
  useEffect(() => { laden(); }, [laden, vote]);

  async function speichern(loeschen = false) {
    setMsg('Speichere…');
    const body = loeschen ? { typ: 'loeschen' } : { typ, min: Number(min), zone: zone || null };
    const r = await j(await fetch(`${SERVER_BASE}/api/garmin/vorschlag`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }));
    if (!r.ok) { setMsg(r.error === 'not_authenticated' ? 'Bitte im Cockpit anmelden.' : `Nicht gespeichert: ${r.error || 'Fehler'}`); return; }
    setMsg(loeschen ? 'Gelöscht – Kira rechnet neu.' : 'Gespeichert – Kira rechnet den Abgleich (bis ca. 30 s).');
    setOffen(false);
    setTimeout(laden, 20000);
    setTimeout(laden, 45000);
  }

  const zeilen: string[] = [];
  if (g?.erholH != null) zeilen.push(`Erholungszeit morgens ${g.erholH.toLocaleString('de-DE', { maximumFractionDigits: 0 })} h`);
  if (g?.statusText) zeilen.push(`Trainingszustand ${g.statusText}`);
  return (
    <div className="mt-2 rounded border border-border px-2 py-1.5 text-xs" data-testid="garmin-zweitmeinung">
      <div className="flex flex-wrap items-baseline justify-between gap-2">
        <span className="label">Garmin</span>
        <button className="btn btn-ghost px-2 py-0.5 text-2xs" onClick={() => setOffen((x) => !x)}>
          {g?.vorschlag ? 'Vorschlag ändern' : 'Garmin-Vorschlag eintragen'}
        </button>
      </div>
      {zeilen.length > 0 && <p className="tnum text-ink-muted">{zeilen.join(' · ')}</p>}
      {g?.deckel?.map((d) => <p key={d} className="text-ink">· {d}</p>)}
      {g?.fokus && <p className="text-ink-muted">{g.fokus}</p>}
      {g?.vorschlag?.text && (
        <p className={g.vorschlag.urteil === 'garmin_weniger' ? 'text-amber-300' : 'text-ink'} data-testid="garmin-vorschlag">{g.vorschlag.text}</p>
      )}
      {!g && <p className="text-ink-dim">Noch keine Garmin-Signale für heute (Mac meldet alle 5–10 min).</p>}
      {offen && (
        <div className="mt-1.5 flex flex-wrap items-center gap-1.5">
          <select className="input py-0.5 text-xs" value={typ} onChange={(e) => { const t = e.target.value as VorschlagTyp; setTyp(t); setZone(''); }} aria-label="Typ">
            {VORSCHLAG_TYPEN.map((t) => <option key={t} value={t}>{VORSCHLAG_LABEL[t]}</option>)}
          </select>
          {typ !== 'ruhe' && (
            <>
              <input className="input w-16 py-0.5 text-xs tnum" inputMode="numeric" value={min} onChange={(e) => setMin(e.target.value.replace(/[^0-9]/g, ''))} aria-label="Minuten" />
              <span className="text-ink-dim">min</span>
              <select className="input py-0.5 text-xs" value={zone} onChange={(e) => setZone(e.target.value)} aria-label="Zone">
                <option value="">{VORSCHLAG_ZONE[typ] || 'Zone'} (Standard)</option>
                {['Z1', 'Z2', 'Z3', 'Z4', 'Z5'].map((z) => <option key={z} value={z}>{z}</option>)}
              </select>
            </>
          )}
          <button className="btn btn-primary px-2 py-0.5 text-xs" onClick={() => speichern(false)}>Speichern</button>
          {g?.vorschlag && <button className="btn btn-ghost px-2 py-0.5 text-xs" onClick={() => speichern(true)}>Löschen</button>}
        </div>
      )}
      {msg && <p className="mt-1 text-2xs text-ink-dim">{msg}</p>}
      <p className="mt-1 text-2xs text-ink-dim">Erholungszeit und Trainingszustand können das Votum senken. Garmins Tagesvorschlag ist eine Zweitmeinung und ändert das Votum nicht.</p>
    </div>
  );
}
