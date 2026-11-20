import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Panel } from './UI';

type Spalte = { spalte: string; geprueft: number; treffer: number; sheetFehler: number; alteVersion: number; festwert: number; maxDiff: number; beispiele: string[] };
type Drift = { spalte: string; verglichen: number; abweichend: number; beispiele: string[] };
type Lauf = { datum: string; alleOk: boolean; abweichend: { spalte: string; n: number; bsp: string }[] };
type Kette = { start: string; tage: number; exakt: number; maxAcwr: number };
type SimTag = { date: string; ess: number; atl: number; ctl: number; acwr: number };
type Antwort = { kette?: Kette | null; ok: boolean; error?: string; heute?: string; stand?: number; spalten?: Spalte[]; drift?: Drift[]; historie?: Lauf[]; serie?: number; ziel?: number };

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
            {d.kette && (
              <div className="opacity-80 mt-1">
                Plan-Kette ab {d.kette.start}: {d.kette.exakt}/{d.kette.tage} Tage exakt wie die Sheet-Prognose (Etappe 1d).
              </div>
            )}
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
      {d?.ok && <WasWaereWenn />}
    </Panel>
  );
}

/** Etappe 1d: geplanten ESS eines Tages ändern und ATL/CTL/ACWR mit dem Rechenkern durchrechnen (nur Anzeige). */
function WasWaereWenn() {
  const [basis, setBasis] = useState<SimTag[]>([]);
  const [plan, setPlan] = useState<SimTag[]>([]);
  const [tag, setTag] = useState('');
  const [wert, setWert] = useState('');
  const [fehler, setFehler] = useState<string | null>(null);
  const rechnen = useCallback(async (ess?: Record<string, number>) => {
    setFehler(null);
    const q = ess ? `?ess=${encodeURIComponent(JSON.stringify(ess))}` : '';
    try {
      const j = await (await fetch(`${SERVER_BASE}/api/rechenkern/simulation${q}`, { cache: 'no-store' })).json();
      if (!j.ok) return setFehler(j.error || 'Fehler');
      setBasis(j.basis); setPlan(j.plan);
      if (!tag && j.basis[0]) { setTag(j.basis[0].date); setWert(String(j.basis[0].ess)); }
    } catch (e) { setFehler(String(e)); }
  }, [tag]);
  useEffect(() => { rechnen(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  if (!basis.length) return fehler ? <p className="text-xs mt-3 opacity-70">Plan-Rechner: {fehler}</p> : null;
  const fmt = (x: number | null | undefined, n = 2) => (x == null || !Number.isFinite(x) ? '–' : x.toLocaleString('de-DE', { maximumFractionDigits: n, minimumFractionDigits: n }));
  const farbe = (a: number) => (a > 1.5 ? 'var(--ck-red, #dc2626)' : a > 1.3 ? 'var(--ck-orange, #ea580c)' : a < 0.8 ? 'var(--ck-yellow, #ca8a04)' : undefined);
  return (
    <div className="mt-4">
      <div className="text-sm font-medium mb-1">Plan-Rechner · Was wäre, wenn …</div>
      <div className="flex flex-wrap items-center gap-2 text-xs mb-2">
        <label className="flex items-center gap-1">Tag
          <select className="bg-transparent border border-current/20 rounded px-1 py-0.5" value={tag}
            onChange={(e) => { setTag(e.target.value); const b = basis.find((x) => x.date === e.target.value); if (b) setWert(String(b.ess)); }}>
            {basis.map((b) => <option key={b.date} value={b.date}>{new Date(b.date + 'T12:00:00').toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit' })}</option>)}
          </select>
        </label>
        <label className="flex items-center gap-1">ESS
          <input className="w-20 bg-transparent border border-current/20 rounded px-1 py-0.5 tabular-nums" inputMode="numeric" value={wert} onChange={(e) => setWert(e.target.value)} />
        </label>
        <button className="px-2 py-0.5 rounded border border-current/20" onClick={() => { const v = Number(wert.replace(',', '.')); if (tag && Number.isFinite(v)) rechnen({ [tag]: v }); }}>Durchrechnen</button>
        <button className="px-2 py-0.5 rounded border border-current/20 opacity-70" onClick={() => rechnen()}>Plan wie im Sheet</button>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-xs tabular-nums">
          <thead className="opacity-60 text-left"><tr><th className="py-1 pr-2">Tag</th><th className="pr-2">ESS</th><th className="pr-2">ATL</th><th className="pr-2">CTL</th><th className="pr-2 whitespace-nowrap">ACWR Plan</th><th className="whitespace-nowrap">ACWR neu</th></tr></thead>
          <tbody>
            {plan.map((p, i) => (
              <tr key={p.date} className="border-t border-current/10">
                <td className="py-1 pr-2 whitespace-nowrap">{p.date.slice(8, 10)}.{p.date.slice(5, 7)}.</td>
                <td className="pr-2" style={{ fontWeight: p.ess !== basis[i]?.ess ? 600 : undefined }}>{fmt(p.ess, 0)}</td>
                <td className="pr-2">{fmt(p.atl, 0)}</td>
                <td className="pr-2">{fmt(p.ctl, 0)}</td>
                <td className="pr-2">{fmt(basis[i]?.acwr)}</td>
                <td style={{ color: farbe(p.acwr), fontWeight: Math.abs(p.acwr - (basis[i]?.acwr ?? p.acwr)) > 0.005 ? 600 : undefined }}>{fmt(p.acwr)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="text-xs opacity-60 mt-1">Nur Anzeige: Der Rechner speichert nichts. Grundlage sind das Formelblatt timeline und die Sheet-Prognoseformeln.</p>
    </div>
  );
}
