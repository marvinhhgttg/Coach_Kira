import { useEffect, useMemo, useState } from 'react';
import { ladeP4, tourform as berechneTourform, type P4Tag } from '../lib/paket4';
import { fetchExtras } from '../lib/extras';
import type { Ziel } from '../lib/planning';
import { useGarminTage } from '../lib/garmin';
import { lerneModelle } from '../lib/essRechner';
import { eigeneSkitouren, tourLastPersoenlich } from '../lib/tourEss';

// ── Hilfsfunktionen ────────────────────────────────────────────────
function fmtDateShort(iso: string): string {
  const p = iso.split('-');
  if (p.length < 3) return iso;
  const tag = parseInt(p[2]);
  const monat = parseInt(p[1]);
  return `${tag}.${String(monat).padStart(2, '0')}.${p[0]}`;
}

function fmtPct(n: number): string {
  return `${n.toLocaleString('de-DE')} %`;
}

function fmtNum(n: number, dec = 0): string {
  if (dec === 0) return n.toLocaleString('de-DE');
  return n.toLocaleString('de-DE', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

// Balken-Farbe nach Prozent (0..1)
function barColor(p: number): string {
  const pct = Math.round(p * 100);
  if (pct >= 80) return 'bg-green-500';
  if (pct >= 60) return 'bg-yellow-500';
  return 'bg-orange-500';
}

function barColorText(p: number): string {
  const pct = Math.round(p * 100);
  if (pct >= 80) return 'text-green-500';
  if (pct >= 60) return 'text-yellow-500';
  return 'text-orange-500';
}

// Heutiges Datum (Mitternacht)
function heute(): Date {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

// Wochen zwischen heute und Zieltour berechnen
function wochenBisZiel(z: Ziel): number {
  const he = heute();
  const zielDatum = new Date(z.datum + 'T12:00:00');
  const diffMs = zielDatum.getTime() - he.getTime();
  return Math.max(0, Math.round(diffMs / (7 * 86400000)));
}

// ── Komponente ─────────────────────────────────────────────────────
export function TourformTile() {
  const [tage, setTage] = useState<P4Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeP4();
        if (!cancelled) setTage(t);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Tourform lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  return <TourformBody p4={tage} />;
}

// ── Interne Komponente mit Daten ───────────────────────────────────
function TourformBody({ p4 }: { p4: P4Tag[] }) {
  const [ziele, setZiele] = useState<Ziel[]>([]);
  const [fetchError, setFetchError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const items = await fetchExtras<{ ziel?: Ziel }>('ziel', 60);
        if (!cancelled) {
          const h = new Date().toLocaleDateString('sv-SE');
          const sorted = items
            .map((it) => it.data?.ziel)
            .filter((z): z is Ziel => !!z && !!z.datum && z.datum >= h)
            .sort((a, b) => a.datum.localeCompare(b.datum));
          setZiele(sorted);
        }
      } catch (e) {
        if (!cancelled) setFetchError(e instanceof Error ? e.message : 'Unbekannter Fehler');
      }
    })();
    return () => { cancelled = true; };
  }, []);
  const { tage: garmin } = useGarminTage();
  const modelle = useMemo(() => (garmin ? lerneModelle(garmin) : []), [garmin]);

  if (ziele.length === 0) {
    return (
      <div className="panel-raised p-4">
        <div>
          <div className="label">Coach Kira · Tourform</div>
          <p className="mt-1 text-2xs text-ink-dim">
            Wie gut du für deine Zieltouren vorbereitet bist – nach dem Vorbild der Marathonform: Wochenumfang, lange Tage, Höhenmeter.
          </p>
        </div>
        <div className="mt-3 text-2xs text-ink-muted">Noch keine Zieltour eingetragen – oben im Zielplaner anlegen.</div>
      </div>
    );
  }

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Tourform</div>
        <p className="mt-1 text-2xs text-ink-dim">
          Wie gut du für deine Zieltouren vorbereitet bist – nach dem Vorbild der Marathonform: Wochenumfang, lange Tage, Höhenmeter.
        </p>
      </div>

      {/* Ziel-Karten */}
      <div className="mt-3 space-y-2">
        {ziele.map((z, i) => {
          const pers = modelle.length ? tourLastPersoenlich(z, modelle, eigeneSkitouren(garmin)) : null;
          const wf = berechneTourform(p4, z, pers?.ess);
          const wo = wochenBisZiel(z);

          // Prozent-Farbe: >= 80 grün, >= 60 gelb, sonst orange
          const pctClass = wf.prozent >= 80 ? 'text-green-500' : wf.prozent >= 60 ? 'text-yellow-500' : 'text-orange-500';
          const hoeheStr = wf.hoehe == null ? 'nicht erfasst' : `${Math.round(wf.hoehe * 100)} %`;

          return (
            <div key={z.datum + z.name} className="rounded-lg border border-border p-3">
              {/* Name, Datum, Wochen */}
              <div className="flex items-baseline gap-2">
                <span className="text-sm font-semibold">{z.name}</span>
                <span className="text-2xs text-ink-dim">{fmtDateShort(z.datum)}</span>
                <span className="text-2xs text-ink-dim">{wo} Wochen</span>
              </div>

              {/* Tourform Prozent groß */}
              <div className={`mt-1 text-2xl font-bold ${pctClass}`}>
                Tourform {fmtPct(wf.prozent)}
              </div>

              {/* Drei Balken */}
              <div className="mt-2 space-y-1">
                {/* Wochenumfang */}
                <div className="flex items-center gap-2">
                  <span className="text-2xs w-16 text-right shrink-0">Wo.</span>
                  <div className="flex-1 h-3 bg-gray-200 rounded-full overflow-hidden">
                    <div className={`h-full ${barColor(wf.volumen)} rounded-full`} style={{ width: `${Math.round(wf.volumen * 100)}%` }} />
                  </div>
                  <span className={`text-2xs w-10 ${barColorText(wf.volumen)}`}>{Math.round(wf.volumen * 100)} %</span>
                </div>
                {/* Lange Tage */}
                <div className="flex items-center gap-2">
                  <span className="text-2xs w-16 text-right shrink-0">Lange</span>
                  <div className="flex-1 h-3 bg-gray-200 rounded-full overflow-hidden">
                    <div className={`h-full ${barColor(wf.lange)} rounded-full`} style={{ width: `${Math.round(wf.lange * 100)}%` }} />
                  </div>
                  <span className={`text-2xs w-10 ${barColorText(wf.lange)}`}>{Math.round(wf.lange * 100)} %</span>
                </div>
                {/* Höhenmeter */}
                <div className="flex items-center gap-2">
                  <span className="text-2xs w-16 text-right shrink-0">Höhe</span>
                  <div className="flex-1 h-3 bg-gray-200 rounded-full overflow-hidden">
                    {wf.hoehe == null ? (
                      <div className="h-full bg-orange-300 rounded-full" style={{ width: '0%' }} />
                    ) : (
                      <div className={`h-full ${barColor(wf.hoehe)} rounded-full`} style={{ width: `${Math.round(wf.hoehe * 100)}%` }} />
                    )}
                  </div>
                  <span className={`text-2xs w-10 ${wf.hoehe == null ? 'text-orange-400' : barColorText(wf.hoehe)}`}>
                    {hoeheStr}
                  </span>
                </div>
              </div>

              {/* Zahlen darunter */}
              <p className="mt-2 text-2xs text-ink-muted">
                Wochenlast {fmtNum(Math.round(wf.wochenLastIst))} von {fmtNum(Math.round(wf.wochenLastZiel))} ESS ·
                lange Tage {wf.langeTage}/{wf.langeZiel} ·
                Tourlast ≈ {fmtNum(wf.tourLast)} ESS{pers ? ' (ESS-Rechner)' : ' (Faustregel)'}
              </p>

              {/* Hinweise */}
              {wf.hinweise.length > 0 && (
                <ul className="mt-1 text-2xs text-ink-muted list-disc ml-4 space-y-0.5">
                  {wf.hinweise.map((h, j) => <li key={j}>{h}</li>)}
                </ul>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
