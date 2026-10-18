import { localStatusText, useEngine, useLocalStatus } from '../lib/engine';

/** Gemeinsamer Wahlschalter Lokal (Mac) / Cloud mit Statuspunkt. Gilt für Chat und alle Briefings. */
export function EngineSwitch({ compact = false }: { compact?: boolean }) {
  const [engine, setEngine] = useEngine();
  const [status, recheck] = useLocalStatus(engine === 'local');
  const st = localStatusText(status);
  const btn = (active: boolean) =>
    `px-2.5 py-1 text-xs rounded ${active ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;
  return (
    <div className="flex flex-wrap items-center gap-2 min-w-0">
      <div className="inline-flex bg-bg rounded border border-border p-0.5" role="tablist" aria-label="KI-Modell">
        <button type="button" className={btn(engine === 'local')} onClick={() => setEngine('local')} aria-pressed={engine === 'local'}>
          Lokal (Mac)
        </button>
        <button type="button" className={btn(engine === 'pplx')} onClick={() => setEngine('pplx')} aria-pressed={engine === 'pplx'}>
          Cloud
        </button>
      </div>
      {engine === 'local' && (
        <span className="inline-flex items-center gap-1.5 text-2xs text-ink-dim">
          <span
            className={`inline-block w-2 h-2 rounded-full ${st.tone === 'ok' ? 'bg-green-400' : st.tone === 'warn' ? 'bg-orange-400' : 'bg-ink-dim'}`}
            aria-hidden
          />
          {compact && st.tone === 'ok' ? 'lokal bereit' : st.text}
          {st.tone !== 'ok' && (
            <button type="button" className="underline hover:text-ink" onClick={recheck}>
              Erneut prüfen
            </button>
          )}
        </span>
      )}
      {engine === 'pplx' && <span className="text-2xs text-ink-dim">Perplexity über Apps Script</span>}
    </div>
  );
}

export function SourceBadge({ source }: { source: { engine: 'local' | 'cloud'; model: string | null } | null | undefined }) {
  if (!source) return null;
  const local = source.engine === 'local';
  return (
    <span className={`chip border text-2xs ${local ? 'border-green-400/40 text-green-300' : 'border-border-strong text-ink-muted'}`}>
      {local ? `lokal · ${source.model || 'coach-kira'}` : `Cloud · ${source.model || 'Perplexity'}`}
    </span>
  );
}
