import { useEffect, useState } from 'react';
import { fetchExtras, type ExtraItem } from '../lib/extras';

// ── Typdefinitionen (parallel zur Beispiel-JSON) ─────────────────────
interface AcwrMax {
  wert: number;
  datum: string;
}

interface Qualitaetstag {
  datum: string;
  tag: string;
  score?: number;
  bereitschaft?: number;
  grund: string;
}

interface TagPrognose {
  datum: string;
  tag: string;
  last?: number;
  atl?: number;
  ctl?: number;
  acwr?: number;
  mono?: number;
  sport: string;
  zone: string;
  fest?: boolean;
  bereitschaft?: number;
}

interface Modell {
  n: number;
  mae: number;
  naiv: number;
}

interface Tourfenster {
  datum: string;
  tag: string;
  score: number;
  wetter: string;
  kalender: string;
  bereitschaft: number;
  plan?: string;
}

interface Signal {
  name: string;
  schwere: 'gelb' | 'orange';
  text: string;
  art: string;
}

interface Fruehwarnung {
  stufe: 'gruen' | 'gelb' | 'orange' | 'rot';
  signale: Signal[];
  prognose: string;
}

export interface PrognoseDaten {
  erstellt: string;
  text: string;
  vorausblick: {
    ctl_heute: number;
    ctl_7: number;
    ctl_14: number;
    ctl_28: number;
    trend: string;
    basis_28: string;
    acwr_max: AcwrMax | null;
    mono_hinweise: string[];
    qualitaetstag: Qualitaetstag | null;
    tage: TagPrognose[];
    modell: Modell | null;
  };
  tourfenster: Tourfenster[];
  fruehwarnung: Fruehwarnung;
}

// ── Hilfsfunktionen ──────────────────────────────────────────────────
const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

function fmtDe(v: number): string {
  return v.toLocaleString('de-DE');
}

function dateShortFromISO(iso: string): string {
  const dp = iso.slice(0, 10);
  const p = dp.split('-');
  if (p.length < 3) return dp;
  const dow = WD[new Date(dp + 'T12:00:00').getDay()];
  return `${dow} ${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

function dateFullFromISO(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  const dow = WD[new Date(iso.slice(0, 10) + 'T12:00:00').getDay()];
  return `${dow}, ${parseInt(p[2])}.${p[1]}.${p[0]}`;
}

function statusColor(b: number | undefined): string {
  if (b === undefined) return 'text-ink-muted';
  if (b >= 70) return 'text-green-300';
  if (b >= 50) return 'text-yellow-300';
  return 'text-orange-300';
}

function stufeColor(s: string): { text: string; border: string } {
  switch (s) {
    case 'gruen':     return { text: 'text-green-300', border: 'border-green-400/40' };
    case 'gelb':      return { text: 'text-yellow-300', border: 'border-yellow-400/40' };
    case 'orange':    return { text: 'text-orange-300', border: 'border-orange-400/50' };
    case 'rot':       return { text: 'text-red-300', border: 'border-red-400/60' };
    default:          return { text: 'text-ink-muted', border: 'border-border' };
  }
}

// ── Komponente ───────────────────────────────────────────────────────
export function VorausblickTile({ compact = false }: { compact?: boolean }) {
  const [item, setItem] = useState<ExtraItem<PrognoseDaten> | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const items = await fetchExtras<PrognoseDaten>('prognose', 1);
        if (!cancelled) {
          setItem(items.length > 0 ? items[0] : null);
          setLoading(false);
        }
      } catch (e) {
        if (!cancelled) {
          setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
          setLoading(false);
        }
      }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) {
    return (
      <div className="panel-raised p-4">
        <p className="text-sm text-ink-dim">Vorausblick lädt…</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="panel-raised p-4">
        <p className="text-sm text-orange-300">{error}</p>
      </div>
    );
  }

  if (!item || !item.data) {
    return (
      <div className="panel-raised p-4">
        <p className="text-sm text-ink-dim">
          Noch kein Vorausblick – der Mac rechnet ihn morgens, sobald deine Morgenwerte da sind.
        </p>
      </div>
    );
  }

  const d = item.data;
  const vk = d.vorausblick;
  const erstelltDp = d.erstellt || item.datum;

  // Zeit / Datum aus dem ISO-Datum
  let hhmm = '', ttmm = '';
  if (erstelltDp) {
    const m = erstelltDp.match(/(\d{2}):(\d{2})/);
    hhmm = m ? `${m[1]}:${m[2]}` : '';
    ttmm = dateShortFromISO(erstelltDp);
  }

  const sw = stufeColor(d.fruehwarnung?.stufe || '');
  const stageLabel: Record<string, string> = {
    gruen:  'Frühwarnung: grün',
    gelb:   'Frühwarnung: gelb',
    orange: 'Frühwarnung: orange',
    rot:    'Frühwarnung: rot',
  };

  // ── compact: drei kurze Kacheln ────────────────────────────────
  if (compact) {
    const qt = vk.qualitaetstag;
    return (
      <div className="panel-raised p-4">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div>
            <div className="label">Coach Kira · Vorausblick</div>
            {ttmm && (
              <p className="text-2xs text-ink-dim mt-1">
                Vom Mac berechnet{hhmm ? ` um ${hhmm}` : ''} ({ttmm}) · lokales Modell
              </p>
            )}
          </div>
          <span className={`chip border text-2xs ${sw.text} ${sw.border}`}>{stageLabel[d.fruehwarnung?.stufe] || 'Frühwarnung'}</span>
        </div>

        <p className="mt-2 text-sm text-ink">{d.text}</p>

        <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
          {/* Qualitätstag */}
          <div className="rounded-lg border border-border p-3">
            {qt ? (
              <>
                <span className="text-2xs text-ink-dim">Qualitätstag</span>
                <p className="tnum mt-1">{qt.tag}</p>
                {qt.bereitschaft !== undefined && (
                  <span className={`tnum text-sm ${statusColor(qt.bereitschaft)}`}>
                    Bereitschaft {fmtDe(qt.bereitschaft)}
                  </span>
                )}
              </>
            ) : (
              <p className="text-2xs text-ink-dim">Kein Qualitätstag</p>
            )}
          </div>

          {/* Bestes Tourfenster */}
          <div className="rounded-lg border border-border p-3">
            {d.tourfenster && d.tourfenster.length > 0 ? (
              <>
                <span className="text-2xs text-ink-dim">Bestes Tourfenster</span>
                <p className="tnum mt-1">{d.tourfenster[0].tag}</p>
                <span className="tnum text-sm text-ink-muted">Eignung {d.tourfenster[0].score}/100</span>
              </>
            ) : (
              <p className="text-2xs text-ink-dim">Kein gutes Tourfenster</p>
            )}
          </div>

          {/* Fitness in 28 Tagen */}
          <div className="rounded-lg border border-border p-3">
            {vk.ctl_28 !== undefined ? (
              <>
                <span className="text-2xs text-ink-dim">Fitness in 4 Wochen</span>
                <p className="tnum mt-1">{fmtDe(vk.ctl_28)}</p>
                <span className="text-2xs text-ink-dim">{vk.trend}</span>
              </>
            ) : (
              <p className="text-2xs text-ink-dim">Keine Daten</p>
            )}
          </div>
        </div>

      </div>
    );
  }

  // ── expanded (compact = false) ─────────────────────────────────
  const signale = d.fruehwarnung?.signale || [];
  const acwr = vk.acwr_max;
  const monoCount = vk.mono_hinweise ? vk.mono_hinweise.length : 0;
  const tageAb1 = vk.tage ? vk.tage.slice(1, 8) : [];

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div className="flex flex-wrap items-start justify-between gap-2">
        <div>
          <div className="label">Coach Kira · Vorausblick</div>
          {ttmm && (
            <p className="text-2xs text-ink-dim mt-1">
              Vom Mac berechnet{hhmm ? ` um ${hhmm}` : ''} ({ttmm}) · lokales Modell
            </p>
          )}
        </div>
        <span className={`chip border text-2xs ${sw.text} ${sw.border}`}>{stageLabel[d.fruehwarnung?.stufe] || 'Frühwarnung'}</span>
      </div>

      {/* Text */}
      <p className="mt-2 text-sm text-ink">{d.text}</p>

      {/* Raster */}
      <div className="grid gap-3 md:grid-cols-3 mt-3">

        {/* ── Karte 1: Frühwarnung ──────────────────────────────────── */}
        <div className="rounded-lg border border-border p-3">
          <span className="label">Frühwarnung</span>

          {signale.length > 0 ? (
            <ul className="mt-2 space-y-1 text-xs">
              {signale.map((s, i) => (
                <li key={i}>
                  <span className={s.schwere === 'gelb' ? 'text-yellow-300' : 'text-orange-300'}>
                    {s.text}
                  </span>
                </li>
              ))}
            </ul>
          ) : (
            <p className="mt-2 text-xs text-ink-muted">Keine Frühwarnzeichen.</p>
          )}

          {d.fruehwarnung?.prognose && (
            <p className="mt-2 text-xs text-ink-muted">{d.fruehwarnung.prognose}</p>
          )}
        </div>

        {/* ── Karte 2: Chancen ─────────────────────────────────────── */}
        <div className="rounded-lg border border-border p-3">
          <span className="label">Chancen</span>

          {/* Qualitätstag */}
          {vk.qualitaetstag && (
            <div className="mt-2">
              <p className="text-2xs text-ink-dim">Bester Qualitätstag</p>
              <p className="text-sm">{vk.qualitaetstag.tag}</p>
              {vk.qualitaetstag.grund && (
                <p className="text-2xs text-ink-dim">{vk.qualitaetstag.grund}</p>
              )}
            </div>
          )}

          {/* Tourfenster-Liste */}
          {d.tourfenster && d.tourfenster.length > 0 ? (
            <div className="mt-3">
              <p className="text-2xs text-ink-dim">Tourfenster</p>
              {d.tourfenster.map((tf, i) => (
                <div key={i} className="text-2xs mt-1">
                  <span>{tf.tag}</span>
                  <span className="ml-1 tnum">Eignung {tf.score}/100</span>
                  <span className="ml-1 text-ink-dim">{tf.wetter}</span>
                  <span className="ml-1 text-ink-dim">{tf.kalender}</span>
                </div>
              ))}
            </div>
          ) : (
            <p className="mt-2 text-xs text-ink-muted">
              Kein gutes Tourfenster in den nächsten 9 Tagen.
            </p>
          )}
        </div>

        {/* ── Karte 3: Fitness-Trend ───────────────────────────────── */}
        <div className="rounded-lg border border-border p-3">
          <span className="label">Fitness-Trend</span>

          {/* CTL-Reihe */}
          <div className="mt-2 flex items-baseline gap-1">
            {vk.ctl_heute !== undefined && (
              <>
                <span className="tnum">{fmtDe(vk.ctl_heute)}</span>
                <span className="text-2xs text-ink-dim">→</span>
              </>
            )}
            {vk.ctl_7 !== undefined && (
              <>
                <span className="tnum">{fmtDe(vk.ctl_7)}</span>
                <span className="text-2xs text-ink-dim">→</span>
              </>
            )}
            {vk.ctl_14 !== undefined && (
              <>
                <span className="tnum">{fmtDe(vk.ctl_14)}</span>
                <span className="text-2xs text-ink-dim">→</span>
              </>
            )}
            {vk.ctl_28 !== undefined && (
              <span className="tnum">{fmtDe(vk.ctl_28)}</span>
            )}
          </div>

          {/* Trend-Wort */}
          {vk.trend && (
            <p className="mt-1 text-sm">{vk.trend}</p>
          )}

          {/* basis_28 Hinweis */}
          {vk.basis_28 && (
            <p className="text-2xs text-ink-dim">ab Tag 15: {vk.basis_28}</p>
          )}

          {/* ACWR max */}
          {acwr && (
            <p className="text-2xs text-ink-dim">
              ACWR max: {fmtDe(acwr.wert)} am {dateShortFromISO(acwr.datum)}
            </p>
          )}

          {/* Monotonie-Hinweise */}
          {monoCount > 0 && (
            <p className="text-2xs text-ink-dim">
              {monoCount} Hinweis(e) zur Monotonie im Plan.
            </p>
          )}

          {/* Fehler-Hinweis falls modell */}
          {vk.modell && (
            <p className="mt-1 text-2xs text-ink-dim">
              Fehler ±{fmtDe(vk.modell.mae)} (naiv ±{fmtDe(vk.modell.naiv)})
            </p>
          )}

          {/* Bereitschafts-Prognosen */}
          {tageAb1.length > 0 && (
            <div className="mt-2 flex gap-2 overflow-x-auto">
              {tageAb1.map((t, i) => (
                <div key={i} className="flex flex-col items-center min-w-[3rem]">
                  <span className="text-2xs text-ink-dim">{t.tag?.split(' ')[0] || ''}</span>
                  <span className={`tnum ${statusColor(t.bereitschaft)}`}>
                    {t.bereitschaft !== undefined ? fmtDe(t.bereitschaft) : '–'}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
