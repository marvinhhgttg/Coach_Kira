import { useEffect, useState } from 'react';
import { ladeP3, intensitaetProWoche, type P3Tag } from '../lib/paket3';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
function fmtDateShort(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  return `${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

// Typ-Chip Farben und Text
function typChipInfo(typ: string): { text: string; cls: string } {
  switch (typ) {
    case 'polarisiert': return { text: typ, cls: 'bg-green-500/20 text-green-600' };
    case 'pyramidal': return { text: typ, cls: 'text-ink' };
    case 'schwellenlastig': return { text: typ, cls: 'bg-yellow-400/20 text-yellow-700' };
    case 'Grundlage': return { text: typ, cls: 'text-ink-muted' };
    case 'Ruhewoche': return { text: typ, cls: 'text-ink-dim' };
    default: return { text: typ, cls: 'text-ink-muted' };
  }
}

// ── Komponente ───────────────────────────────────────────────────────────
export function IntensitaetTile() {
  const [tage, setTage] = useState<P3Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

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
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Intensitätsverteilung lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  const ws = intensitaetProWoche(tage, 12);

  // Fazit: Typen zählen
  const typZahlen = new Map<string, number>();
  for (const w of ws) {
    typZahlen.set(w.typ, (typZahlen.get(w.typ) ?? 0) + 1);
  }

  const typTexte: string[] = [];
  for (const [typ, count] of typZahlen.entries()) {
    if (count > 0) {
      typTexte.push(`${count}× ${typ}`);
    }
  }

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Intensitätsverteilung</div>
        <p className="mt-1 text-2xs text-ink-muted">
          Lastanteile je Woche nach Zone (locker Z1–Z2, Grauzone Z3, hart Z4–Z5) und Wochentyp.
        </p>
      </div>

      {/* Wochen-Tabelle */}
      <div className="mt-3 space-y-2">
        {ws.map((w) => {
          const chip = typChipInfo(w.typ);
          return (
            <div key={w.woche} className="text-xs">
              {/* Zeile: Datum | Balken | Last | Typ */}
              <div className="flex items-center gap-2">
                {/* Datum */}
                <span className="text-ink-muted whitespace-nowrap w-16">{fmtDateShort(w.woche)}</span>

                {/* Gestapelter Balken */}
                <div className="flex-1 flex rounded-full overflow-hidden h-5">
                  {/* locker */}
                  <div
                    className="bg-green-500/50"
                    style={{ width: `${(w.locker * 100).toFixed(1)}%` }}
                    title={`locker ${(w.locker * 100).toFixed(1)}%`}
                  />
                  {/* grau */}
                  <div
                    className="bg-yellow-400/70"
                    style={{ width: `${(w.grau * 100).toFixed(1)}%` }}
                    title={`grau ${(w.grau * 100).toFixed(1)}%`}
                  />
                  {/* hart */}
                  <div
                    className="bg-red-400/70"
                    style={{ width: `${(w.hart * 100).toFixed(1)}%` }}
                    title={`hart ${(w.hart * 100).toFixed(1)}%`}
                  />
                </div>

                {/* Wochenlast */}
                <span className="tnum w-12 text-right">{w.last} ESS</span>

                {/* Typ-Chip */}
                <span className={`px-2 py-0.5 rounded-full text-xs ${chip.cls}`}>{chip.text}</span>
              </div>
            </div>
          );
        })}
      </div>

      {/* Fazit */}
      <div className="mt-3 text-xs text-ink-muted">
        <p>Letzte 12 Wochen: {typTexte.length > 0 ? typTexte.join(', ') : 'keine Daten'}. Für lange Bergtouren ist viel locker plus wenig gezielt hart meist günstiger als viel Grauzone.</p>
      </div>
    </div>
  );
}
