import { useEffect, useRef, useState } from 'react';
import { DATEN_AB, WINDOWS, isCustom, winGleich, winLabel, type WinKey, type WinPreset } from '../lib/windows';

/**
 * Einheitliche Zeitraumwahl: feste Zeiträume plus „Eigener“ (von–bis).
 * `presets` schränkt die festen Zeiträume ein (Standard: alle).
 */
export function ZeitraumWahl({
  value, onChange, presets = WINDOWS, testIdPrefix,
}: { value: WinKey; onChange: (w: WinKey) => void; presets?: WinPreset[]; testIdPrefix?: string }) {
  const heute = new Date().toLocaleDateString('sv-SE');
  const [offen, setOffen] = useState(false);
  const start = (() => { const d = new Date(); d.setDate(d.getDate() - 29); return d.toLocaleDateString('sv-SE'); })();
  const [von, setVon] = useState(isCustom(value) ? value.von : start);
  const [bis, setBis] = useState(isCustom(value) ? value.bis : heute);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!offen) return;
    const zu = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOffen(false); };
    document.addEventListener('mousedown', zu);
    return () => document.removeEventListener('mousedown', zu);
  }, [offen]);
  const btn = (aktiv: boolean) => `px-1.5 sm:px-2 py-1 text-xs rounded tnum whitespace-nowrap ${aktiv ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;
  const gueltig = von && bis && von <= bis;
  return (
    <div ref={ref} className="relative inline-flex flex-wrap bg-bg rounded border border-border p-0.5" role="group" aria-label="Zeitraum">
      {presets.map((w) => (
        <button key={w} type="button" onClick={() => onChange(w)} className={btn(winGleich(value, w))} data-testid={testIdPrefix ? `${testIdPrefix}-${w}` : undefined}>
          {winLabel(w)}
        </button>
      ))}
      <button type="button" onClick={() => setOffen((o) => !o)} className={btn(isCustom(value))} aria-expanded={offen} data-testid="zeitraum-eigener">
        {isCustom(value) ? winLabel(value) : 'Eigener'}
      </button>
      {offen && (
        <div className="absolute right-0 top-full mt-1 z-40 w-64 rounded-md border border-border bg-bg-panel p-3 shadow-lg text-xs" data-testid="zeitraum-popup">
          <div className="text-2xs text-ink-dim mb-2">Eigener Zeitraum</div>
          <div className="grid grid-cols-2 gap-2">
            <label className="flex flex-col gap-1 text-ink-muted">von
              <input type="date" className="input" min={DATEN_AB} max={heute} value={von} onChange={(e) => setVon(e.target.value)} />
            </label>
            <label className="flex flex-col gap-1 text-ink-muted">bis
              <input type="date" className="input" min={DATEN_AB} max={heute} value={bis} onChange={(e) => setBis(e.target.value)} />
            </label>
          </div>
          <div className="mt-3 flex justify-end gap-2">
            <button type="button" className="btn btn-ghost text-xs px-2 py-1" onClick={() => setOffen(false)}>Abbrechen</button>
            <button type="button" className="btn btn-primary text-xs px-2 py-1" disabled={!gueltig}
              onClick={() => { onChange({ von, bis }); setOffen(false); }}>Übernehmen</button>
          </div>
        </div>
      )}
    </div>
  );
}
