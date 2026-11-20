import { useEffect, useState } from 'react';
import { ladeP4, rezeptbuch, type P4Tag, type Block } from '../lib/paket4';

// ── Hilfsfunktionen ────────────────────────────────────────────────
function fmtDateShort(iso: string): string {
  const p = iso.split('-');
  if (p.length < 3) return iso;
  const tag = parseInt(p[2]);
  const monat = parseInt(p[1]);
  return `${tag}.${String(monat).padStart(2, '0')}.${p[0]}`;
}

function fmtInt(n: number): string {
  return n.toLocaleString('de-DE');
}

function fmtOne(n: number): string {
  return n.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

// Sport-Mix: die zwei größten Anteile als Text
function sportMixText(mix: Record<string, number>): string {
  const entries = Object.entries(mix).sort((a, b) => b[1] - a[1]);
  if (entries.length === 0) return '–';
  const top = entries.slice(0, 2);
  return top.map(([name, anteil]) => `${name} ${Math.round(anteil * 100)} %`).join(', ');
}

// Zonen als Prozent ganzzahlig
function zonenText(zonen: { locker: number; grau: number; hart: number }): string {
  const l = Math.round(zonen.locker * 100);
  const g = Math.round(zonen.grau * 100);
  const h = Math.round(zonen.hart * 100);
  return `locker ${l} % · Grauzone ${g} % · hart ${h} %`;
}

// Block-Zeile für die Tabelle
function blockTabZeile(block: Block): string[] {
  const k = block.k;
  return [
    `${fmtDateShort(block.start)}–${fmtDateShort(block.ende)}`,
    block.esDelta != null ? `${block.esDelta > 0 ? '+' : block.esDelta < 0 ? '−' : ''}${Math.abs(block.esDelta).toLocaleString('de-DE')}` : '–',
    fmtInt(Math.round(k.wochenLast)),
    String(k.ruhetage),
    block.k.laengste ? fmtInt(Math.round(block.k.laengste)) : '–',
    zonenText(k.zonen),
    sportMixText(k.sportMix),
  ];
}

// ── Komponente ─────────────────────────────────────────────────────
export function RezeptbuchTile() {
  const [tage, setTage] = useState<P4Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeP4();
        if (!cancelled) setTage(t);
        setLoading(false);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Rezeptbuch lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  const rb = rezeptbuch(tage);

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Rezeptbuch</div>
        <p className="mt-1 text-2xs text-ink-dim">
          Deine 4-Wochen-Blöcke mit dem größten Zuwachs beim Endurance Score – und was sie gemeinsam hatten.
        </p>
      </div>

      {/* Muster */}
      <div className="mt-3">
        <div className="text-xs font-semibold text-ink mb-1">So trainierst du erfolgreich</div>
        {rb.muster.length > 0 ? (
          <ul className="text-2xs text-ink-muted space-y-1 ml-4 list-disc">
            {rb.muster.map((m, i) => <li key={i}>{m}</li>)}
          </ul>
        ) : (
          <p className="text-2xs text-ink-muted">Keine deutlichen Unterschiede zu deinem Durchschnitt.</p>
        )}
      </div>

      {/* Beste Blöcke */}
      {rb.top.length > 0 && (
        <div className="mt-3">
          <div className="text-xs font-semibold text-ink mb-1">Deine besten Blöcke</div>
          <div className="overflow-x-auto">
            <table className="w-full text-2xs border-collapse">
              <thead>
                <tr className="text-ink-dim text-left border-b border-border">
                  <th className="p-1 font-normal whitespace-nowrap">Zeitraum</th>
                  <th className="p-1 font-normal whitespace-nowrap">ES-Änd.</th>
                  <th className="p-1 font-normal whitespace-nowrap">Wochenlast</th>
                  <th className="p-1 font-normal whitespace-nowrap">Ruhetage</th>
                  <th className="p-1 font-normal whitespace-nowrap">Längste</th>
                  <th className="p-1 font-normal whitespace-nowrap">Zonen</th>
                  <th className="p-1 font-normal whitespace-nowrap">Sport-Mix</th>
                </tr>
              </thead>
              <tbody>
                {rb.top.map((block, i) => {
                  const zeile = blockTabZeile(block);
                  return (
                    <tr key={i} className="border-b border-border">
                      <td className="p-1 whitespace-nowrap">{zeile[0]}</td>
                      <td className="p-1 whitespace-nowrap text-green-600">{zeile[1]}</td>
                      <td className="p-1 whitespace-nowrap tnum">{zeile[2]}</td>
                      <td className="p-1 whitespace-nowrap tnum">{zeile[3]}</td>
                      <td className="p-1 whitespace-nowrap tnum">{zeile[4]}</td>
                      <td className="p-1">{zeile[5]}</td>
                      <td className="p-1 whitespace-nowrap">{zeile[6]}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Schwächste Blöcke */}
      {rb.schwach.length > 0 && (
        <div className="mt-3">
          <div className="text-xs font-semibold text-ink mb-1">Schwächste Blöcke</div>
          <div className="overflow-x-auto">
            <table className="w-full text-2xs border-collapse">
              <thead>
                <tr className="text-ink-dim text-left border-b border-border">
                  <th className="p-1 font-normal whitespace-nowrap">Zeitraum</th>
                  <th className="p-1 font-normal whitespace-nowrap">ES-Änd.</th>
                  <th className="p-1 font-normal whitespace-nowrap">Wochenlast</th>
                  <th className="p-1 font-normal whitespace-nowrap">Ruhetage</th>
                  <th className="p-1 font-normal whitespace-nowrap">Längste</th>
                  <th className="p-1 font-normal whitespace-nowrap">Zonen</th>
                  <th className="p-1 font-normal whitespace-nowrap">Sport-Mix</th>
                </tr>
              </thead>
              <tbody>
                {rb.schwach.map((block, i) => {
                  const zeile = blockTabZeile(block);
                  return (
                    <tr key={i} className="border-b border-border">
                      <td className="p-1 whitespace-nowrap">{zeile[0]}</td>
                      <td className="p-1 whitespace-nowrap text-orange-500">{zeile[1]}</td>
                      <td className="p-1 whitespace-nowrap tnum">{zeile[2]}</td>
                      <td className="p-1 whitespace-nowrap tnum">{zeile[3]}</td>
                      <td className="p-1 whitespace-nowrap tnum">{zeile[4]}</td>
                      <td className="p-1">{zeile[5]}</td>
                      <td className="p-1 whitespace-nowrap">{zeile[6]}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Fußzeile */}
      <p className="mt-3 text-2xs text-ink-dim">{rb.alle.toLocaleString('de-DE')} Blöcke ausgewertet. Zusammenhang, keine Ursache – Krankheit oder Urlaub können Blöcke prägen.</p>
    </div>
  );
}
