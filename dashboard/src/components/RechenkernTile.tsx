import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Panel } from './UI';

type Spalte = { spalte: string; geprueft: number; treffer: number; sheetFehler: number; alteVersion: number; festwert: number; maxDiff: number; beispiele: string[] };
type Drift = { spalte: string; verglichen: number; abweichend: number; beispiele: string[] };
type Lauf = { datum: string; alleOk: boolean; abweichend: { spalte: string; n: number; bsp: string }[] };
type Antwort = { ok: boolean; error?: string; heute?: string; stand?: number; spalten?: Spalte[]; drift?: Drift[]; historie?: Lauf[]; serie?: number; ziel?: number };

/** Etappe 1 (nur v2): Rechenkern im Schattenbetrieb gegen das Sheet abgleichen. */
export function RechenkernTile() {
  const [v2, setV2] = useState(false);
  const [d, setD] = useState<Antwort | null>(null);
  const [laedt, setLaedt] = useState(false);

  useEffect(() => {
    fetch(`${SERVER_BASE}/api/instanz`).then((r) => r.json()).then((j) => setV2(j?.instanz === 'v2')).catch(() => setV2(false));
  }, []);

  const laden = useCallback(async (neu = false) => {
    setLaedt(true);
    try {
      const r = await fetch(`${SERVER_BASE}/api/rechenkern/abgleich${neu ? '?neu=1' : ''}`, { cache: 'no-store' });
      setD(await r.json());
    } catch (e) {
      setD({ ok: false, error: String(e) });
    } finally {
      setLaedt(false);
    }
  }, []);
  useEffect(() => { if (v2) laden(); }, [v2, laden]);

  if (!v2) return null;
  const alleOk = d?.ok && d.spalten?.every((s) => s.treffer === s.geprueft);

  return (
    <Panel title="Rechenkern · Abgleich mit dem Sheet (Etappe 1)">
      <div className="flex items-center justify-between gap-3 mb-3">
        <p className="text-sm opacity-80">
          {laedt ? 'Rechne nach …' : !d ? '' : !d.ok ? `Noch kein Abgleich: ${d.error}` : alleOk
            ? 'Alle Formeln der aktuellen Version stimmen mit dem Formelblatt timeline überein.'
            : 'Es gibt Abweichungen – Details unten.'}
        </p>
        <button className="text-xs px-2 py-1 rounded border border-current/20 whitespace-nowrap shrink-0" onClick={() => laden(true)} disabled={laedt}>Neu abgleichen</button>
      </div>
      {d?.ok && (
        <>
          <div className="mb-3 rounded border border-current/10 p-2 text-xs">
            <div className="flex items-baseline gap-2">
              <span className="text-lg font-semibold tabular-nums">{d.serie ?? 0}/{d.ziel ?? 14}</span>
              <span className="opacity-80">Tage in Folge ohne Abweichung (täglicher Abgleich ab 10 Uhr)</span>
            </div>
            <div className="flex flex-wrap gap-1 mt-2" aria-label="Tägliche Läufe">
              {(d.historie || []).map((h) => (
                <span key={h.datum} title={h.alleOk ? `${h.datum}: alles gleich` : `${h.datum}: ${h.abweichend.map((x) => `${x.spalte} (${x.n})`).join(', ')}`}
                  className="inline-block w-3 h-3 rounded-sm" style={{ background: h.alleOk ? 'var(--ck-green, #16a34a)' : 'var(--ck-orange, #ea580c)' }} />
              ))}
              {!d.historie?.length && <span className="opacity-60">Erster täglicher Lauf folgt heute ab 10 Uhr.</span>}
            </div>
            <div className="opacity-60 mt-1">Verlaufsdaten in v2 kommen bereits aus timeline (Etappe 1b).</div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs tabular-nums">
              <thead className="opacity-60 text-left">
                <tr><th className="py-1 pr-2">Spalte</th><th className="pr-2">Treffer</th><th className="pr-2">Sheet-Fehler</th><th className="pr-2">ältere Formel</th><th className="pr-2">fester Wert</th><th>erste Abweichung</th></tr>
              </thead>
              <tbody>
                {d.spalten!.map((s) => (
                  <tr key={s.spalte} className="border-t border-current/10">
                    <td className="py-1 pr-2 font-medium">{s.spalte}</td>
                    <td className="pr-2" style={{ color: s.treffer === s.geprueft ? 'var(--ck-green, #16a34a)' : 'var(--ck-orange, #ea580c)' }}>{s.treffer}/{s.geprueft}</td>
                    <td className="pr-2">{s.sheetFehler || '–'}</td>
                    <td className="pr-2">{s.alteVersion || '–'}</td>
                    <td className="pr-2">{s.festwert || '–'}</td>
                    <td className="opacity-70">{s.beispiele[0] || '–'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs opacity-60 mt-2">
            Geprüft werden nur Zellen mit der heutigen Formelversion. Ältere Formeln und feste Werte gelten als Wahrheit der Vergangenheit und bleiben unverändert.
          </p>
          {!!d.drift?.length && (
            <div className="mt-4">
              <div className="text-sm font-medium mb-1">Wertekopie KK_TIMELINE gegenüber timeline</div>
              <div className="flex flex-wrap gap-x-4 gap-y-1 text-xs tabular-nums">
                {d.drift.map((x) => (
                  <span key={x.spalte} title={x.beispiele.join('\n')} style={{ color: x.abweichend ? 'var(--ck-orange, #ea580c)' : undefined }}>
                    {x.spalte}: {x.abweichend}/{x.verglichen}
                  </span>
                ))}
              </div>
              <p className="text-xs opacity-60 mt-1">Das Cockpit liest heute aus KK_TIMELINE. Abweichende Tage sind dort älter als im Formelblatt.</p>
            </div>
          )}
          <p className="text-xs opacity-50 mt-2">Stand {d.stand ? new Date(d.stand).toLocaleString('de-DE') : '–'} · nur lesend</p>
        </>
      )}
    </Panel>
  );
}
