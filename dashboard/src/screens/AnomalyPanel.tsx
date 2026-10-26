import { useEffect, useMemo, useState } from 'react';
import { Panel } from '../components/UI';
import { fetchChartData } from '../lib/api';
import { buildDays, RECOVERY_FETCH_METRICS, type RecoveryDay } from './RecoveryPanel';
import { detectAnomalies } from '../lib/anomalies';
import { localChat } from '../lib/localLLM';
import { useEngine } from '../lib/engine';
import { fmtDateShort } from '../lib/format';

export function AnomalyPanel() {
  const [engine] = useEngine();
  const [days, setDays] = useState<RecoveryDay[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [open, setOpen] = useState(false);
  const [kira, setKira] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [kErr, setKErr] = useState<string | null>(null);

  useEffect(() => {
    fetchChartData('120d', RECOVERY_FETCH_METRICS)
      .then((r) => setDays(buildDays(r.data)))
      .catch((e) => setErr(String(e instanceof Error ? e.message : e)));
  }, []);

  const res = useMemo(() => (days ? detectAnomalies(days) : null), [days]);

  return (
    <Panel title="Auffälligkeiten · Frühwarnung" right={
      <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => setOpen((o) => !o)}>
        {open ? '▴ Einklappen' : '▾ Ausklappen'}
      </button>
    }>
      {err ? (
        <p className="text-xs text-red-300">{err}</p>
      ) : !days ? (
        <p className="text-xs text-ink-dim">Lade…</p>
      ) : (
        <>
          {res!.items.length === 0 ? (
            <p className="text-xs text-green-300">
              Keine Auffälligkeiten am {fmtDateShort(res!.date ?? undefined)} – alle Werte im persönlichen Normalbereich.
            </p>
          ) : (
            <p className="text-xs">
              <span className={res!.items.some((i) => i.level === 'bad') ? 'text-red-300' : 'text-orange-300'}>
                {res!.items.length} Auffälligkeit(en)
              </span>{' '}
              am {fmtDateShort(res!.date ?? undefined)}: {res!.items.map((i) => i.label).join(', ')}
            </p>
          )}
          {open && (
            <div>
              <ul className="mt-2 space-y-2">
                {res!.items.map((i) => (
                  <li key={i.key} className="rounded border border-border px-3 py-2">
                    <div className="flex items-center gap-2 text-sm">
                      <span className={`inline-block w-2 h-2 rounded-full ${i.level === 'bad' ? 'bg-red-400' : 'bg-orange-400'}`} />
                      <span className="font-semibold">{i.label}</span>
                      <span className="text-ink-muted">{i.value}</span>
                      <span className="ml-auto text-2xs text-ink-dim">
                        {i.baseline}
                        {i.z != null ? ' · z = ' + i.z.toFixed(1).replace('.', ',') : ''}
                      </span>
                    </div>
                    <p className="text-2xs text-ink-dim mt-1">{i.hint}</p>
                  </li>
                ))}
              </ul>
              <p className="mt-2 text-2xs text-ink-dim">
                Fest berechnet: letzter Tag mit Morgenwerten gegen deinen 28-Tage-Schnitt (ab 2 Standardabweichungen), dazu ACWR, Monotonie, Befinden und 3-Tage-Muster. Keine Diagnose.
              </p>
              {res!.items.length > 0 && <div className="mt-4">
                <div className="label mb-1">Kira erklärt</div>
                {res!.items.length > 0 ? (
                  <>
                    <button
                      className="btn btn-primary text-xs px-2 py-1"
                      onClick={async () => {
                        setBusy(true);
                        setKErr(null);
                        setKira('');
                        try {
                          const out = await localChat(
                            [
                              { role: 'system', content: 'Du bist Coach Kira, Sportwissenschaftlerin. Du sprichst Marc per Du an. Antworte auf Deutsch in 3 bis 5 Sätzen, ohne Überschrift. Keine Diagnose; bei Krankheitsverdacht zu Pause und gegebenenfalls Arzt raten. Nutze nur die gegebenen Werte.' },
                              { role: 'user', content: 'Auffälligkeiten am ' + res!.date + ':\n' + res!.items.map((i) => '- ' + i.label + ': ' + i.value + ' (' + i.baseline + (i.z != null ? ', z ' + i.z.toFixed(1) : '') + ')').join('\n') + '\n\nWas bedeutet das zusammen, und was soll ich heute konkret tun?' },
                            ],
                            { numCtx: 8192, onToken: setKira },
                          );
                          setKira(out);
                        } catch (e) {
                          setKErr(e instanceof Error ? e.message : String(e));
                        } finally {
                          setBusy(false);
                        }
                      }}
                    >
                      {busy ? `Kira (${engine === 'local' ? 'lokal' : 'Cloud'}) schreibt…` : `Kira erklärt (${engine === 'local' ? 'lokal' : 'Cloud'})`}
                    </button>
                    {kErr && <p className="mt-2 text-xs text-red-300">{kErr}</p>}
                    {kira && <p className="mt-2 text-sm text-ink-muted whitespace-pre-wrap leading-relaxed">{kira}</p>}
                  </>
                ) : res!.items.length > 0 ? (
                  <p className="mt-2 text-2xs text-ink-dim">Kira-Erklärung nur lokal – Schalter auf „Lokal (Mac)“ stellen.</p>
                ) : null}
              </div>}
            </div>
          )}
        </>
      )}
    </Panel>
  );
}
