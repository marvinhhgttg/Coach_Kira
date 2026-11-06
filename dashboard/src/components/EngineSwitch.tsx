import { cloudStatus } from '../lib/cloudLLM';
import { useEffect, useState } from 'react';
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
        <button type="button" className={`${btn(engine === 'pplx')} ${engine === 'pplx' ? 'engine-cloud-on' : ''}`} onClick={() => setEngine('pplx')} aria-pressed={engine === 'pplx'}>
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
      {engine === 'pplx' && <CloudStatusText compact={compact} />}
    </div>
  );
}

function CloudStatusText({ compact }: { compact: boolean }) {
  const [cloud, setCloud] = useState<{ ok: boolean; configured: boolean; used: number; limit: number } | null>(null);
  useEffect(() => {
    let off = false;
    cloudStatus().then((c) => !off && setCloud(c));
    return () => {
      off = true;
    };
  }, []);
  const ok = !!cloud?.ok && cloud.configured && cloud.used < cloud.limit;
  return (
    <span className="inline-flex items-center gap-1.5 text-2xs text-ink-dim" data-testid="engine-cloud-status">
      <span className={`inline-block w-2 h-2 rounded-full ${cloud == null ? 'bg-ink-dim' : ok ? 'bg-blue-500' : 'bg-orange-400'}`} aria-hidden />
      {cloud == null ? 'Cloud wird geprüft …' : ok ? (compact ? 'Cloud bereit' : 'Cloud bereit · Perplexity') : cloud.configured ? 'Cloud-Kontingent für heute aufgebraucht' : 'Cloud nicht eingerichtet'}
    </span>
  );
}

export function SourceBadge({ source }: { source: { engine: 'local' | 'cloud'; model: string | null } | null | undefined }) {
  if (!source) return null;
  const local = source.engine === 'local';
  return (
    <span className={`chip border text-2xs ${local ? 'border-green-400/40 text-green-300' : 'border-blue-400/40 text-blue-300'}`}>
      {local ? `lokal · ${source.model || 'coach-kira'}` : `Cloud · ${source.model || 'Perplexity'}`}
    </span>
  );
}

/** Kompakter globaler Schalter für die Kopfzeile: Lokal/Cloud mit Statuspunkt, prüft alle 60 s. */
export function HeaderEngineSwitch() {
  const [engine, setEngine] = useEngine();
  const [status, recheck] = useLocalStatus(engine === 'local');
  const st = localStatusText(status);
  useEffect(() => {
    if (engine !== 'local') return;
    const id = window.setInterval(recheck, 60000);
    return () => window.clearInterval(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [engine]);
  const [cloud, setCloud] = useState<{ ok: boolean; configured: boolean; used: number; limit: number } | null>(null);
  const recheckCloud = () => cloudStatus().then(setCloud);
  useEffect(() => {
    if (engine !== 'pplx') return;
    recheckCloud();
    const id = window.setInterval(recheckCloud, 60000);
    return () => window.clearInterval(id);
  }, [engine]);
  const cloudOk = !!cloud?.ok && cloud.configured && cloud.used < cloud.limit;
  const dot =
    engine !== 'local'
      ? cloud == null ? 'bg-ink-dim' : cloudOk ? 'bg-blue-500' : 'bg-orange-400'
      : st.tone === 'ok' ? 'bg-green-400' : st.tone === 'warn' ? 'bg-orange-400' : 'bg-ink-dim';
  const title =
    engine === 'local'
      ? `KI-Modell: lokal auf dem Mac · ${st.text}${st.tone !== 'ok' ? ' (Klick auf den Punkt prüft erneut)' : ''}`
      : `KI-Modell: Cloud (Perplexity über den Cockpit-Server)${cloud ? (cloud.configured ? ` · heute ${cloud.used}/${cloud.limit} Anfragen` : ' · Schlüssel fehlt auf dem Server') : ''}`;
  const btn = (active: boolean) =>
    `px-2.5 py-1 text-xs rounded transition-colors ${active ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;
  return (
    <div
      className="inline-flex items-center gap-1.5 rounded-md border border-border bg-bg-panel p-0.5 pl-2 shadow-sm"
      role="group"
      aria-label="KI-Modell wählen (gilt für Chat und Briefings)"
      title={title}
      data-testid="engine-switch-header"
    >
      <button
        type="button"
        onClick={engine === 'local' ? recheck : recheckCloud}
        className={`inline-block w-2 h-2 rounded-full ${dot}`}
        aria-label={title}
      />
      <span className="hidden lg:inline text-2xs text-ink-dim">KI</span>
      <button type="button" className={btn(engine === 'local')} onClick={() => setEngine('local')} aria-pressed={engine === 'local'}>
        Lokal
      </button>
      <button type="button" className={`${btn(engine === 'pplx')} ${engine === 'pplx' ? 'engine-cloud-on' : ''}`} onClick={() => setEngine('pplx')} aria-pressed={engine === 'pplx'}>
        Cloud
      </button>
    </div>
  );
}
