import { useEffect, useState } from 'react';
import {
  ladeTage,
  erholungsDauer,
  erholungJeBand,
  lastBand,
  type BandErholung,
  zonenKlasse,
  type ETag,
} from '../lib/recoveryPrice';
import { fetchExtras, type ExtraItem } from '../lib/extras';

// ── Typdefinitionen (aus VorausblickTile-Vorlage) ───────────────────────
interface TagPrognose {
  datum: string;
  tag: string;
  last?: number;
  sport: string;
  zone: string;
  bereitschaft?: number;
}

interface PrognoseDaten {
  vorausblick: {
    tage: TagPrognose[];
  };
}

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

function fmtDe(v: number): string {
  return v.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function dateShortFromISO(iso: string): string {
  const dp = iso.slice(0, 10);
  const p = dp.split('-');
  if (p.length < 3) return dp;
  const dow = WD[new Date(dp + 'T12:00:00').getDay()];
  return `${dow} ${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

// ── Komponente ───────────────────────────────────────────────────────────
export function ErholungsuhrTile() {
  const [tage, setTage] = useState<ETag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeTage();
        if (!cancelled) setTage(t);
        setLoading(false);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const erholungData = erholungsDauer(tage);
  const p90 = erholungData.p90;
  const median = erholungData.median;

  // Basis = Mittel Readiness der letzten 28 Tage vor heute (ohne krank)
  const today = new Date().toLocaleDateString('sv-SE');
  const basisReadiness: number[] = [];
  for (let i = tage.length - 2; i >= 0 && basisReadiness.length < 28; i--) {
    const t = tage[i];
    if (t.krank) continue;
    if (t.readiness == null) continue;
    basisReadiness.push(t.readiness);
  }
  const basis = basisReadiness.length > 0 ? basisReadiness.reduce((a, b) => a + b, 0) / basisReadiness.length : NaN;

  // Prognose laden
  const [prognoseData, setPrognoseData] = useState<ExtraItem<PrognoseDaten>[]>();
  const [prognoseLoading, setPrognoseLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const items = await fetchExtras<PrognoseDaten>('prognose', 1);
        if (!cancelled) {
          setPrognoseData(items);
          setPrognoseLoading(false);
        }
      } catch {
        if (!cancelled) setPrognoseLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // grosserTag: jüngster Tag (heute oder gestern) mit load >= p90
  const heute = tage.length > 0 ? tage[tage.length - 1] : null;
  const gestern = tage.length > 1 ? tage[tage.length - 2] : null;
  const grosserTag = (() => {
    if (heute && heute.load != null && heute.load >= p90) return heute;
    if (gestern && gestern.load != null && gestern.load >= p90) return gestern;
    return null;
  })();

  // Erwarteter Erholungstag aus Prognose finden
  let erwartet: TagPrognose | null = null;
  if (grosserTag && prognoseData) {
    const firstPrognose = prognoseData.length > 0 ? prognoseData[0] : null;
    const tageList = firstPrognose?.data?.vorausblick.tage || [];
    for (const tp of tageList) {
      if (tp.datum > grosserTag.date && tp.bereitschaft != null) {
        if (tp.bereitschaft >= basis - 3) {
          erwartet = tp;
          break;
        }
      }
    }
  }

  // Planprüfung: Tage nach grosserTag und vor erwartet (oder alle wenn kein erwartet)
  const warnTage: string[] = [];
  if (grosserTag) {
    const firstPrognose = prognoseData ? prognoseData[0] : null;
    const alleTage = firstPrognose?.data?.vorausblick.tage || [];
    for (const tp of alleTage) {
      if (tp.datum <= grosserTag.date) continue;
      if (erwartet && tp.datum >= erwartet.datum) continue;
      const klass = zonenKlasse(tp.zone);
      if (klass === 'grau' || klass === 'hart' || (tp.last != null && tp.last >= 0.9 * p90)) {
        warnTage.push(tp.datum);
      }
    }
  }

  if (loading || prognoseLoading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Erholungsuhr lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  // Ohne großen Tag: Erfahrungswert für die Last von heute bzw. gestern (alle Lastbereiche)
  if (!grosserTag) {
    const bz = erholungJeBand(tage);
    const tagMitLast = heute && heute.load != null && heute.load > 0 ? heute : gestern && gestern.load != null && gestern.load > 0 ? gestern : null;
    const band = tagMitLast ? bz.baender.find((b) => b.band === lastBand(tagMitLast.load!, bz.p50, bz.p90)) : undefined;
    const tageTxt = (m: number | null) => (m == null ? 'unbekannt' : m <= 1 ? 'am nächsten Morgen' : `nach ${m} Tagen`);
    const grenzTxt = (b: BandErholung) => (b.bis == null ? `ab ${Math.round(b.von)}` : b.band === 'locker' ? `unter ${Math.round(b.bis)}` : `${Math.round(b.von)}–${Math.round(b.bis) - 1}`);
    return (
      <div className="panel-raised p-4" data-testid="erholungsuhr">
        <div className="label">Coach Kira · Erholungsuhr</div>
        {tagMitLast && band ? (
          <p className="mt-2 text-xs text-ink">
            {tagMitLast === heute ? 'Heute' : 'Gestern'} {Math.round(tagMitLast.load!)} ESS ({band.label}):{' '}
            {band.median != null && band.median <= 1 && band.sofortAnteil != null
              ? `Nach solchen Tagen bist du in ${Math.round(band.sofortAnteil * 100)} % der Fälle schon am nächsten Morgen wieder auf deiner Basis (~${fmtDe(basis)}).`
              : `Nach solchen Tagen bist du meist ${tageTxt(band.median)} wieder auf deiner Basis (~${fmtDe(basis)})${band.sofortAnteil != null && band.faelle >= 5 ? `, in ${Math.round(band.sofortAnteil * 100)} % der Fälle schon am nächsten Morgen` : ''}.`}
          </p>
        ) : (
          <p className="mt-2 text-xs text-ink">Gestern und heute ohne Training – Erholung läuft.</p>
        )}
        <div className="mt-2 grid grid-cols-3 gap-2 text-2xs tnum">
          {bz.baender.map((b) => (
            <div key={b.band} className={`rounded border p-1.5 ${band?.band === b.band ? 'border-accent' : 'border-border'}`}>
              <div className="text-ink-muted">{b.label} · {grenzTxt(b)} ESS</div>
              <div className="text-ink text-xs">{b.median == null ? '—' : b.median <= 1 ? '1 Tag' : `${b.median === 9 ? '9+' : b.median} Tage`}</div>
              <div className="text-ink-dim">{b.faelle} Fälle</div>
            </div>
          ))}
        </div>
        <p className="mt-1 text-2xs text-ink-dim">Tage bis die Bereitschaft wieder höchstens 3 Punkte unter deinem 28-Tage-Schnitt liegt (Median, ohne Krankheitstage).</p>
      </div>
    );
  }

  const grosserTagLabel = dateShortFromISO(grosserTag.date);
  const loadText = fmtDe(grosserTag.load!);

  if (erwartet) {
    const erwartetLabel = dateShortFromISO(erwartet.datum);
    return (
      <div className="panel-raised p-4">
        {/* Kopf */}
        <div>
          <div className="label">Coach Kira · Erholungsuhr</div>
        </div>

        <p className="mt-2 text-xs text-ink">
          Nach deinem großen Tag am {grosserTagLabel} ({loadText} ESS) bist du voraussichtlich am {erwartetLabel} wieder auf deiner Basis (~{fmtDe(basis)}).
        </p>

        <p className="mt-1 text-2xs text-ink-muted">
          Bisher nach Tagen ab {Math.round(p90)} ESS meist nach {median === null ? 'unbekannt' : String(Math.round(median))} Tagen.
        </p>

        {/* Warn-Tage */}
        {warnTage.length > 0 && (
          <div className="mt-2">
            {warnTage.map((date, i) => (
              <p key={i} className="text-2xs text-orange-300">
                Qualität am {dateShortFromISO(date)} liegt vor der erwarteten Erholung – lieber verschieben oder locker halten.
              </p>
            ))}
          </div>
        )}

        {/* Prognosedaten-Tabelle */}
        {prognoseData && prognoseData.length > 0 && (
          <div className="mt-3">
            <span className="label">Vorausblick</span>
            <div className="overflow-x-auto mt-2">
              <table className="w-full text-xs tnum">
                <thead>
                  <tr className="text-2xs text-ink-dim">
                    <th className="text-left pl-2">Tag</th>
                    <th className="text-right pr-2">Bereitschaft</th>
                    <th className="text-right">Last</th>
                    <th className="text-left">Zone</th>
                  </tr>
                </thead>
                <tbody>
                  {(prognoseData[0]?.data?.vorausblick.tage || []).map((tp, i) => (
                    <tr key={i} className="border-t border-border">
                      <td className="pl-2 text-ink">{dateShortFromISO(tp.datum)}</td>
                      <td className="text-right">
                        <span className={tp.bereitschaft != null && tp.bereitschaft >= basis ? 'text-green-300' : tp.bereitschaft != null && tp.bereitschaft >= basis - 3 ? 'text-yellow-300' : 'text-orange-300'}>
                          {tp.bereitschaft != null ? fmtDe(tp.bereitschaft) : '–'}
                        </span>
                      </td>
                      <td className="text-right">
                        {tp.last != null ? fmtDe(tp.last) : '–'}
                      </td>
                      <td className="text-left">
                        {zonenKlasse(tp.zone)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    );
  }

  // Erwartung ohne konkreten Tag (in den nächsten 7 Tagen)
  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Erholungsuhr</div>
      </div>

      <p className="mt-2 text-xs text-ink">
        Nach deinem großen Tag am {grosserTagLabel} ({loadText} ESS) bist du laut Prognose in den nächsten 7 Tagen voraussichtlich noch nicht ganz zurück auf deiner Basis (~{fmtDe(basis)}).
      </p>

      <p className="mt-1 text-2xs text-ink-muted">
        Bisher nach Tagen ab {Math.round(p90)} ESS meist nach {median === null ? 'unbekannt' : String(Math.round(median))} Tagen.
      </p>

      {/* Warn-Tage */}
      {warnTage.length > 0 && (
        <div className="mt-2">
          {warnTage.map((date, i) => (
            <p key={i} className="text-2xs text-orange-300">
              Qualität am {dateShortFromISO(date)} liegt vor der erwarteten Erholung – lieber verschieben oder locker halten.
            </p>
          ))}
        </div>
      )}
    </div>
  );
}
