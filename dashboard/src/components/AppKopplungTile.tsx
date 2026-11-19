import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Panel } from './UI';

type Geraet = { id: string; name: string; at: number; zuletzt: number | null };
const zeit = (ms: number | null) =>
  ms ? new Date(ms).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : 'noch nie';

async function j(r: Response) {
  const t = await r.text();
  try {
    return JSON.parse(t);
  } catch {
    return { ok: false, error: `HTTP ${r.status}` };
  }
}

/** Technik → iPhone-App koppeln: Kopplungscode erzeugen, gekoppelte Geräte sehen und widerrufen. */
export function AppKopplungTile() {
  const [code, setCode] = useState<{ code: string; bis: number } | null>(null);
  const [geraete, setGeraete] = useState<Geraet[]>([]);
  const [err, setErr] = useState<string | null>(null);
  const [abgleich, setAbgleich] = useState<any>(null);
  const laden = useCallback(async () => {
    const r = await j(await fetch(`${SERVER_BASE}/api/widget/geraete?cb=${Date.now()}`, { cache: 'no-store' }));
    if (r.ok) setGeraete(r.geraete || []);
    else setErr(r.error);
  }, []);
  useEffect(() => {
    laden();
    const id = window.setInterval(laden, 15000);
    return () => window.clearInterval(id);
  }, [laden]);
  const erzeugen = async () => {
    setErr(null);
    const r = await j(await fetch(`${SERVER_BASE}/api/widget/koppeln`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}' }));
    if (r.ok) setCode({ code: r.code, bis: r.gueltigBis });
    else setErr(r.error);
  };
  const widerrufen = async (id: string) => {
    const r = await j(await fetch(`${SERVER_BASE}/api/widget/widerrufen`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ id }) }));
    if (r.ok) setGeraete(r.geraete || []);
  };
  const pruefen = async () => {
    setAbgleich({ laden: true });
    setAbgleich(await j(await fetch(`${SERVER_BASE}/api/widget/abgleich?cb=${Date.now()}`, { cache: 'no-store' })));
  };
  return (
    <Panel title="iPhone-App & Widgets" right={<button type="button" className="btn btn-ghost text-xs px-2 py-1" onClick={erzeugen} data-testid="app-koppeln">Kopplungscode erzeugen</button>}>
      <div className="space-y-3 text-xs" data-testid="app-kopplung">
        <div className="text-ink-muted">Die Coach-Kira-App und ihre Widgets holen den Tagesstand mit einem eigenen Schlüssel. Zum Koppeln hier einen Code erzeugen und in der App unter „Mit Cockpit verbinden“ eingeben.</div>
        {code && (
          <div className="flex flex-wrap items-baseline gap-3">
            <span className="font-mono text-2xl tracking-[0.3em] text-ink" data-testid="app-code">{code.code}</span>
            <span className="text-2xs text-ink-dim">gültig bis {new Date(code.bis).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })}, einmal verwendbar</span>
          </div>
        )}
        {err && <div className="text-orange-300">{err}</div>}
        <div>
          <div className="text-ink-dim mb-1">Gekoppelte Geräte</div>
          {geraete.length === 0 ? (
            <div className="text-ink-dim">Noch keins.</div>
          ) : (
            <ul className="divide-y divide-border/60">
              {geraete.map((g) => (
                <li key={g.id} className="py-1.5 flex flex-wrap items-center justify-between gap-2">
                  <span>
                    <span className="text-ink">{g.name}</span>
                    <span className="text-ink-dim"> · gekoppelt {zeit(g.at)} · zuletzt abgefragt {zeit(g.zuletzt)}</span>
                  </span>
                  <button type="button" className="btn btn-ghost text-2xs px-2 py-0.5" onClick={() => widerrufen(g.id)}>Widerrufen</button>
                </li>
              ))}
            </ul>
          )}
        </div>
        <div className="flex flex-wrap items-center gap-2 text-2xs text-ink-dim">
          <button type="button" className="btn btn-ghost text-2xs px-2 py-0.5" onClick={pruefen} data-testid="widget-abgleich">Widget-Werte prüfen</button>
          {abgleich?.laden && <span>rechnet …</span>}
          {abgleich && !abgleich.laden && (abgleich.ok ? (
            <span data-testid="widget-abgleich-werte">Widget zeigt: Votum {abgleich.vote ?? '–'} · Readiness {abgleich.readiness ?? '–'} · Soll heute {abgleich.sollHeute ?? '–'} ESS · ACWR {abgleich.acwr ?? '–'} · Tourform {abgleich.tourform ?? '–'} %</span>
          ) : <span className="text-orange-300">{abgleich.error}</span>)}
        </div>
      </div>
    </Panel>
  );
}
