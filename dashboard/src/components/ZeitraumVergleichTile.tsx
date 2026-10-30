import { useEffect, useState, useMemo } from 'react';
import { ladeP4, zeitraumVergleich, type P4Tag } from '../lib/paket4';

// ── Hilfsfunktionen ────────────────────────────────────────────────
function fmtValue(val: number | null, einheit: string): string {
  if (val === null) return '–';
  if (einheit === 'h') {
    return val.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' ' + einheit;
  }
  return Math.round(val).toLocaleString('de-DE') + (einheit ? ' ' + einheit : '');
}

function diffFarbe(diff: number | null, besserWennHoeher: boolean | null): string {
  if (diff === null || besserWennHoeher === null) return 'text-ink-muted';
  if (besserWennHoeher) {
    return diff > 0 ? 'text-green-600' : diff < 0 ? 'text-orange-500' : 'text-ink-muted';
  } else {
    return diff < 0 ? 'text-green-600' : diff > 0 ? 'text-orange-500' : 'text-ink-muted';
  }
}

// ── Komponente ─────────────────────────────────────────────────────
export function ZeitraumVergleichTile() {
  const [tage, setTage] = useState<P4Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Voreinstellungen: A und B als { von, bis }
  const today = new Date();
  const toISODate = (d: Date): string => d.toISOString().split('T')[0];

  const [praeKey, setPraeKey] = useState<string>('last4');
  const [aVon, setAVon] = useState('');
  const [aBis, setABis] = useState('');
  const [bVon, setBVon] = useState('');
  const [bBis, setBBis] = useState('');

  // Setze Standard-Zeiträume beim Laden und bei Präferenzwechsel
  useEffect(() => {
    if (tage.length > 0) {
      const td = toISODate(today); // heute als String
      const yesterday = new Date(today);
      yesterday.setDate(yesterday.getDate() - 1);
      const yesterdayStr = toISODate(yesterday);

      const setPrae = (pv: string, pbis: string, bbv: string, bbis: string) => {
        setAVon(pv); setABis(pbis); setBVon(bbv); setBBis(bbis);
      };

      if (praeKey === 'last4') {
        // A = letzte 28 Tage bis gestern, B = 28 Tage davor
        const aBisDate = new Date(yesterday);
        aBisDate.setDate(aBisDate.getDate() - 27);
        const bBisDate = new Date(aBisDate);
        bBisDate.setDate(bBisDate.getDate() - 1);
        const bVonDate = new Date(bBisDate);
        bVonDate.setDate(bVonDate.getDate() - 27);
        setPrae(toISODate(aBisDate), yesterdayStr, toISODate(bVonDate), toISODate(bBisDate));
      } else if (praeKey === 'thisYear') {
        // A = 1. des Monats bis gestern, B = gleicher Zeitraum letztes Jahr
        const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
        const lastYearMonthStart = new Date(today.getFullYear() - 1, today.getMonth(), 1);
        const lastYearYesterday = new Date(yesterday);
        lastYearYesterday.setFullYear(lastYearYesterday.getFullYear() - 1);
        setPrae(toISODate(monthStart), yesterdayStr, toISODate(lastYearMonthStart), toISODate(lastYearYesterday));
      } else if (praeKey === 'last90') {
        // A = letzte 90 Tage bis gestern, B = gleiche 90 Tage letztes Jahr
        const aBisDate = new Date(yesterday);
        aBisDate.setDate(aBisDate.getDate() - 89);
        const bBisDate = new Date(aBisDate);
        bBisDate.setDate(bBisDate.getDate() - 1);
        const bVonDate = new Date(bBisDate);
        bVonDate.setDate(bVonDate.getDate() - 89);
        setPrae(toISODate(aBisDate), yesterdayStr, toISODate(bVonDate), toISODate(bBisDate));
      }
    }
  }, [praeKey, tage]);

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

  // Berechme Vergleichsdaten, wenn A und B gesetzt sind
  const ergebnis = useMemo(() => {
    if (!aVon || !aBis || !bVon || !bBis) return [];
    return zeitraumVergleich(tage, { von: aVon, bis: aBis }, { von: bVon, bis: bBis });
  }, [tage, aVon, aBis, bVon, bBis]);

  const praeKnöpfe = [
    { key: 'last4', label: '4 Wochen gegen die 4 davor' },
    { key: 'thisYear', label: 'Dieser Monat gegen Vorjahr' },
    { key: 'last90', label: 'Letzte 90 gegen Vorjahr' },
  ];

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Zeitraumvergleich lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Zeitraumvergleich</div>
      </div>

      {/* Präferenz-Knöpfe */}
      <div className="mt-3 flex flex-wrap gap-2">
        {praeKnöpfe.map(k => (
          <button
            key={k.key}
            className={`chip rounded-lg border border-border p-3 cursor-pointer text-xs ${praeKey === k.key ? 'bg-blue-500/20 border-blue-400' : ''}`}
            onClick={() => setPraeKey(k.key)}
          >
            {k.label}
          </button>
        ))}
      </div>

      {/* Datumseingaben */}
      <div className="mt-3 grid gap-3 sm:grid-cols-2">
        <div>
          <label className="block text-2xs text-ink-dim mb-1">A – von</label>
          <input type="date" value={aVon} onChange={e => setAVon(e.target.value)} className="w-full border border-border rounded px-2 py-1 text-xs" />
          <label className="block text-2xs text-ink-dim mb-1 mt-2">A – bis</label>
          <input type="date" value={aBis} onChange={e => setABis(e.target.value)} className="w-full border border-border rounded px-2 py-1 text-xs" />
        </div>
        <div>
          <label className="block text-2xs text-ink-dim mb-1">B – von</label>
          <input type="date" value={bVon} onChange={e => setBVon(e.target.value)} className="w-full border border-border rounded px-2 py-1 text-xs" />
          <label className="block text-2xs text-ink-dim mb-1 mt-2">B – bis</label>
          <input type="date" value={bBis} onChange={e => setBBis(e.target.value)} className="w-full border border-border rounded px-2 py-1 text-xs" />
        </div>
      </div>

      {/* Tabelle */}
      {ergebnis.length > 0 && (
        <div className="mt-3 overflow-x-auto">
          <table className="w-full text-xs border-collapse">
            <thead>
              <tr className="text-ink-dim text-left border-b border-border">
                <th className="p-1 font-normal">Kennzahl</th>
                <th className="p-1 font-normal">A</th>
                <th className="p-1 font-normal">B</th>
                <th className="p-1 font-normal">Differenz</th>
              </tr>
            </thead>
            <tbody>
              {ergebnis.map((row) => {
                const diff = row.a != null && row.b != null ? row.a - row.b : null;
                return (
                  <tr key={row.key} className="border-b border-border">
                    <td className="p-1 font-medium">{row.label}</td>
                    <td className="p-1 tnum">{fmtValue(row.a, row.einheit)}</td>
                    <td className="p-1 tnum">{fmtValue(row.b, row.einheit)}</td>
                    <td className={`p-1 tnum ${diffFarbe(diff, row.besserWennHoeher)}`}>
                      {diff !== null ? `${diff > 0 ? '+' : diff < 0 ? '−' : ''}${Math.abs(diff).toLocaleString('de-DE', { maximumFractionDigits: row.einheit === 'h' ? 1 : 0 })}${row.einheit ? ' ' + row.einheit : ''}` : '–'}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
