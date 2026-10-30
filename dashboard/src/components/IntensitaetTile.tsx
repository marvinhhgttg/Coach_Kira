import { useEffect, useState } from 'react';
import { ladeP3, intensitaetProWoche, type P3Tag } from '../lib/paket3';
import { useGarminTage, type GarminTage } from '../lib/garmin';

type Basis = 'last' | 'zeit';
/** Zeitanteile je Woche aus den Garmin-Zonensekunden (Z1–Z2 locker, Z3 grau, Z4–Z5 hart). */
function zeitAnteile(g: GarminTage | null, wocheMontag: string) {
  if (!g) return null;
  const z = [0, 0, 0, 0, 0];
  const start = new Date(wocheMontag + 'T12:00:00');
  for (let i = 0; i < 7; i++) {
    const d = new Date(start); d.setDate(start.getDate() + i);
    const zs = g.get(d.toLocaleDateString('sv-SE'))?.summe?.zonen_s;
    if (zs) for (let k = 0; k < 5; k++) z[k] += zs[k] || 0;
  }
  const sum = z.reduce((a, b) => a + b, 0);
  if (sum <= 0) return null;
  return { locker: (z[0] + z[1]) / sum, grau: z[2] / sum, hart: (z[3] + z[4]) / sum, std: sum / 3600 };
}
function typAus(locker: number, grau: number, hart: number, ruhe: boolean): string {
  void locker;
  if (ruhe) return 'Ruhewoche';
  if (hart >= 0.1 && grau < 0.15) return 'polarisiert';
  if (grau >= 0.25) return 'schwellenlastig';
  if (grau >= 0.1 && grau > hart) return 'pyramidal';
  return 'Grundlage';
}

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
  const { tage: garmin } = useGarminTage();
  const [basis, setBasis] = useState<Basis>('zeit');

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

  const wsLast = intensitaetProWoche(tage, 12);
  const hatGarmin = !!garmin && wsLast.some((w) => zeitAnteile(garmin, w.woche));
  const nachZeit = basis === 'zeit' && hatGarmin;
  const ws = wsLast.map((w) => {
    if (!nachZeit) return { ...w, std: null as number | null };
    const z = zeitAnteile(garmin, w.woche);
    if (!z) return { ...w, locker: 0, grau: 0, hart: 0, std: null as number | null };
    return { ...w, locker: z.locker, grau: z.grau, hart: z.hart, std: z.std, typ: typAus(z.locker, z.grau, z.hart, w.typ === 'Ruhewoche') as typeof w.typ };
  });

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
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="label">Coach Kira · Intensitätsverteilung</div>
          {hatGarmin && (
            <div className="inline-flex bg-bg rounded border border-border p-0.5 text-xs" data-testid="intensitaet-basis">
              <button className={`px-2 py-0.5 rounded ${basis === 'zeit' ? 'bg-bg-subtle text-ink' : 'text-ink-muted'}`} onClick={() => setBasis('zeit')}>Zeit in HF-Zonen</button>
              <button className={`px-2 py-0.5 rounded ${basis === 'last' ? 'bg-bg-subtle text-ink' : 'text-ink-muted'}`} onClick={() => setBasis('last')}>Last nach Zone</button>
            </div>
          )}
        </div>
        <p className="mt-1 text-2xs text-ink-muted">
          {nachZeit
            ? 'Zeitanteile je Woche aus den Garmin-Zonensekunden (locker Z1–Z2, Grauzone Z3, hart Z4–Z5) und Wochentyp.'
            : 'Lastanteile je Woche nach eingetragener Zone (locker Z1–Z2, Grauzone Z3, hart Z4–Z5) und Wochentyp.'}
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
                <span className="tnum w-16 text-right whitespace-nowrap">{w.last} ESS</span>
                {nachZeit && <span className="tnum w-12 text-right text-ink-muted">{w.std != null ? `${w.std.toFixed(1).replace('.', ',')} h` : '—'}</span>}

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
