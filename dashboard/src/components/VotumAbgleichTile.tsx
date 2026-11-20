import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Panel } from './UI';

type Eintrag = { datum: string; cockpit: string; ampel: string; urteil: string; diff: number | null; score: number | null; planStatus: string | null; bind: string | null };
type Antwort = {
  ok: boolean; error?: string; datum?: string;
  cockpit?: { vote: string | null; gruende: string[]; bind: string | null } | null;
  supervisor?: { ampel: string | null; score: number | null; planStatus: string | null; lauf: string | null; aktuell: boolean };
  vergleich?: { urteil: string; diff: number | null };
  historie?: Eintrag[];
  quote?: { tage: number; gleich: number; cockpitStrenger: number; supervisorStrenger: number };
};

const FARBE: Record<string, string> = { gleich: 'var(--ck-green, #16a34a)', cockpit_strenger: 'var(--ck-orange, #ea580c)', supervisor_strenger: '#7c3aed', offen: '#9ca3af' };
const TEXT: Record<string, string> = { gleich: 'gleiche Stufe', cockpit_strenger: 'Cockpit strenger', supervisor_strenger: 'Supervisor strenger', offen: 'noch offen' };
const VOTE_DE: Record<string, string> = { QUALITY: 'Quality (blau)', TRAIN: 'Train (grün)', EASY: 'Easy (gelb)', REST: 'Rest (rot)' };

/** Etappe 2 Vorbereitung (nur v2): Cockpit-Votum gegen Supervisor-Ampel, nur lesend. */
export function VotumAbgleichTile() {
  const [v2, setV2] = useState(false);
  const [d, setD] = useState<Antwort | null>(null);
  const [laedt, setLaedt] = useState(false);
  useEffect(() => {
    fetch(`${SERVER_BASE}/api/instanz`).then((r) => r.json()).then((j) => setV2(j?.instanz === 'v2')).catch(() => setV2(false));
  }, []);
  const laden = useCallback(async (neu = false) => {
    setLaedt(true);
    try { setD(await (await fetch(`${SERVER_BASE}/api/rechenkern/votum${neu ? '?neu=1' : ''}`, { cache: 'no-store' })).json()); }
    catch (e) { setD({ ok: false, error: String(e) }); }
    finally { setLaedt(false); }
  }, []);
  useEffect(() => { if (v2) laden(); }, [v2, laden]);
  if (!v2) return null;
  const u = d?.vergleich?.urteil || 'offen';
  const q = d?.quote;
  return (
    <Panel title="Votum-Abgleich · Cockpit und Supervisor (Etappe 2, Vorbereitung)">
      <div className="flex items-center justify-between gap-3 mb-3">
        <p className="text-sm opacity-80">
          {laedt ? 'Lese beide Voten …' : !d ? '' : !d.ok ? `Kein Abgleich: ${d.error}` : 'Zwei unabhängige Regelwerke bewerten denselben Tag. Ziel von Etappe 2: eine Regel, ein Votum.'}
        </p>
        <button className="text-xs px-2 py-1 rounded border border-current/20 whitespace-nowrap shrink-0" onClick={() => laden(true)} disabled={laedt}>Neu lesen</button>
      </div>
      {d?.ok && (
        <>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs mb-3">
            <div className="rounded border border-current/10 p-2">
              <div className="opacity-60">Cockpit (Regeln im Frontend)</div>
              <div className="text-base font-semibold">{d.cockpit?.vote ? VOTE_DE[d.cockpit.vote] : 'Morgenwerte fehlen'}</div>
              {d.cockpit?.gruende?.map((g) => <div key={g} className="opacity-70">{g}</div>)}
            </div>
            <div className="rounded border border-current/10 p-2">
              <div className="opacity-60">Supervisor (Apps Script)</div>
              <div className="text-base font-semibold">{d.supervisor?.ampel ?? '–'}{d.supervisor?.score != null ? ` · ${d.supervisor.score}` : ''}</div>
              <div className="opacity-70">Plan-Status {d.supervisor?.planStatus ?? '–'}</div>
              <div className="opacity-70">{d.supervisor?.aktuell ? `Lauf ${new Date(d.supervisor.lauf || '').toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })} Uhr` : 'heute noch kein Supervisor-Lauf'}</div>
            </div>
            <div className="rounded border border-current/10 p-2">
              <div className="opacity-60">Heute</div>
              <div className="text-base font-semibold" style={{ color: FARBE[u] }}>{TEXT[u]}</div>
              {d.vergleich?.diff != null && d.vergleich.diff !== 0 && <div className="opacity-70">{Math.abs(d.vergleich.diff)} Stufe(n) Unterschied</div>}
            </div>
          </div>
          <div className="rounded border border-current/10 p-2 text-xs">
            <div className="flex items-baseline gap-2">
              <span className="text-lg font-semibold tabular-nums">{q?.gleich ?? 0}/{q?.tage ?? 0}</span>
              <span className="opacity-80">Tage mit gleicher Stufe (täglich ab 12 Uhr mitgeschrieben)</span>
            </div>
            <div className="flex flex-wrap gap-1 mt-2" aria-label="Tägliche Votum-Vergleiche">
              {(d.historie || []).map((h) => (
                <span key={h.datum} title={`${h.datum}: Cockpit ${h.cockpit}, Supervisor ${h.ampel} → ${TEXT[h.urteil]}`} className="inline-block w-3 h-3 rounded-sm" style={{ background: FARBE[h.urteil] }} />
              ))}
              {!d.historie?.length && <span className="opacity-60">Erster Eintrag folgt heute ab 12 Uhr.</span>}
            </div>
            <div className="opacity-60 mt-1">Der Supervisor überschreibt seine Ampel täglich und führt keine Historie, deshalb beginnt die Reihe erst jetzt.</div>
          </div>
        </>
      )}
    </Panel>
  );
}
