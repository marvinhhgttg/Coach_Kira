import { useEffect, useState } from 'react';
import { ladeP3, wochenrhythmus, type P3Tag, type Fenster } from '../lib/paket3';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
function fmtDe(v: number): string {
  return v.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

// ── Komponente ───────────────────────────────────────────────────────────
export function WochenrhythmusTile() {
  const [tage, setTage] = useState<P3Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [fenster, setFenster] = useState<Fenster>('alle');

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeP3();
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
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Wochenrhythmus lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  const result = wochenrhythmus(tage, fenster);
  const kuerzel = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];

  // Zeilenmittel und Hintergrund-Farben berechnen
  interface ZellInfo {
    wert: number | null;
    mittel: number | null;
    bgClass?: string;
  }

  const zeilenMittel: (number | null)[] = [];
  type ZellArray = Array<Array<ZellInfo>>;
  const zellen: ZellArray = [];

  for (const zeile of result.zeilen) {
    const werte = zeile.werte;
    const nichtNull = werte.filter((v): v is number => v !== null);
    const mittel = nichtNull.length > 0 ? nichtNull.reduce((a, b) => a + b, 0) / nichtNull.length : null;
    zeilenMittel.push(mittel);

    const rowZellen: ZellInfo[] = [];
    for (const v of werte) {
      let bgClass: string | undefined;
      if (v !== null && mittel != null && mittel > 0) {
        const abw = v / mittel - 1;
        if (zeile.key === 'load') {
          // Last-Zeile: neutral bis auf über Mittel -> bg-sky-500/20
          if (abw > 0) {
            bgClass = 'bg-sky-500/20';
          }
        } else {
          // Schlaf, Readiness, HRV: > +3% grün, < -3% orange
          if (abw > 0.03) bgClass = 'bg-green-500/20';
          else if (abw < -0.03) bgClass = 'bg-orange-500/20';
        }
      }
      rowZellen.push({ wert: v, mittel, bgClass });
    }
    zellen.push(rowZellen);
  }

  // Fenster-Labels für Knöpfe
  const fensterLabel: Record<Fenster, string> = {
    alle: 'Alle',
    '90': '90 Tage',
    sommer: 'Sommer',
    winter: 'Winter',
  };

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Wochenrhythmus</div>
        <p className="mt-1 text-2xs text-ink-muted">
          Ø je Wochentag. Farbe zeigt, ob ein Tag über (grün) oder unter (orange) deinem eigenen Wochenschnitt liegt – bei Last umgekehrt nicht gewertet.
        </p>
      </div>

      {/* Umschalter */}
      <div className="mt-3 flex gap-1">
        {(['alle', '90' as Fenster, 'sommer', 'winter'] as Fenster[]).map((f) => (
          <button
            key={f}
            className={`px-3 py-1 rounded-full text-xs border transition ${fenster === f ? 'bg-ink/20 border-ink font-semibold' : 'border-ink-dim/30 text-ink-muted hover:border-ink/40'}`}
            onClick={() => setFenster(f)}
          >
            {fensterLabel[f]}
          </button>
        ))}
      </div>

      {/* Heatmap-Tabelle */}
      <table className="mt-3 w-full text-xs border-collapse">
        <thead>
          <tr>
            <th className="text-left text-ink-muted font-normal pb-1 pr-2"></th>
            {kuerzel.map((tag) => (
              <th key={tag} className="text-center text-ink-muted font-normal px-1 pb-1">{tag}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {result.zeilen.map((zeile, ri) => (
            <tr key={zeile.key}>
              <td className="font-medium text-ink-muted pr-2 align-top pt-1">{zeile.label}</td>
              {zellen[ri].map((z, ci) => (
                <td
                  key={ci}
                  className={`text-center px-1 py-2 rounded ${z.bgClass ?? ''}`}
                  title={`${kuerzel[ci]}: ${z.wert !== null ? (zeile.key === 'sleep' ? fmtDe(z.wert) + ' h' : String(Math.round(z.wert))) : '–'} (n=${result.zeilen[ri].n[ci]})`}
                >
                  {z.wert !== null ? (zeile.key === 'sleep' ? fmtDe(z.wert) : String(Math.round(z.wert))) : '–'}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      {/* Hinweiszeile */}
      <div className="mt-3 text-xs text-ink-muted">
        Deine stärksten Tage: {result.besteTage.join(', ')}. Schwächster Tag: {result.schwaechsterTag}. Qualität eher auf die starken Tage legen, den schwächsten Tag locker halten. Der Wochen-Autopilot nutzt das als weiche Vorliebe.
      </div>
    </div>
  );
}
