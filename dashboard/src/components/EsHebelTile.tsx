import { useEffect, useRef, useState } from 'react';
import { ladeP3, esBacktest, esBreakEven, esZiel, type P3Tag } from '../lib/paket3';
import { useGarminTage, fmtDauer } from '../lib/garmin';
import { laengsteProTag, langEffekt, langStatus, langEmpfehlung, sportDe } from '../lib/esLang';
import { heuteBerlin } from '../lib/kernVotumKi';

function fmtDateShort(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  return `${parseInt(p[2])}.${p[1].padStart(2, '0')}.${p[0]}`;
}

// ── Komponente ───────────────────────────────────────────────────────────
export function EsHebelTile() {
  const [tage, setTage] = useState<P3Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const { tage: garmin } = useGarminTage();

  // Ziel-Rechner state
  const [zielScore, setZielScore] = useState<number>(() => {
    try {
      const n = Number(localStorage.getItem('ck_es_ziel'));
      if (Number.isFinite(n) && n >= 1000 && n <= 20000) return n;
    } catch {}
    return 7000;
  });
  const [zielDatum, setZielDatum] = useState<string>(() => {
    try {
      const saved = localStorage.getItem('ck_es_datum');
      if (saved) return saved;
    } catch {}
    const d = new Date();
    d.setDate(d.getDate() + 140);
    return d.toISOString().slice(0, 10);
  });

  const zielRef = useRef<HTMLDivElement>(null);
  const [zielErgebnis, setZielErgebnis] = useState<{ tage: number; ladeProTag: number | null; machbar: boolean } | null>(null);

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

  // Ziel-Rechner berechnen wenn Inputs sich ändern
  useEffect(() => {
    if (tage.length === 0) return;
    // Ziel-Datum speichern
    try { localStorage.setItem('ck_es_datum', zielDatum); } catch {}
    // Ziel-Score speichern
    try { localStorage.setItem('ck_es_ziel', String(zielScore)); } catch {}
    const result = esZiel(tage, zielScore, zielDatum);
    setZielErgebnis(result);
  }, [tage.length, zielScore, zielDatum]);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Endurance-Score-Hebel lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  // letzterES: letzter Tag mit es
  let letzterES: number | null = null;
  for (let i = tage.length - 1; i >= 0; i--) {
    if (tage[i].es !== null) {
      letzterES = tage[i].es;
      break;
    }
  }

  // last28: Ø load der letzten 28 Tage (null = 0)
  const last28Count = Math.min(tage.length, 28);
  let last28: number | null = null;
  if (last28Count > 0) {
    // 28 abgeschlossene Tage vor heute, fehlende Last = 0 (wie im Modell)
    const fenster = tage.slice(Math.max(0, tage.length - 29), tage.length - 1);
    last28 = fenster.reduce((acc, x) => acc + (x.load ?? 0), 0) / Math.max(1, fenster.length);
  }

  // Gleichgewicht: esBreakEven(letzterES)
  let gleichgewicht: number | null = null;
  if (letzterES !== null) {
    gleichgewicht = esBreakEven(letzterES);
  }

  // Belastbarkeit: b = esBacktest(tage)
  const backtest = tage.length > 0 ? esBacktest(tage) : { n: 0, maeModell: 0, maeNaiv: 0, horizont: 28 };
  const notBetter = backtest.maeModell >= backtest.maeNaiv;

  // Ziel-Rechner: z = esZiel(tage, ziel, datum)
  const ziel = zielScore;
  const datum = zielDatum;
  const z = zielErgebnis;

  // Lange Einheiten (Garmin-Tagesdetails): Status, eigene Wirkung, Empfehlung
  const lang = garmin ? laengsteProTag(garmin) : null;
  const heute = heuteBerlin();
  const eff = lang ? langEffekt(tage, lang) : null;
  const st = lang ? langStatus(lang, heute) : null;
  const emp = st ? langEmpfehlung(st) : null;
  const extraProTag = eff && emp ? (eff.bonus4 * emp.sehrLangProMonat) / 30 : 0;
  const zLang = eff && eff.bonus4 > 0 && z && z.tage > 0 ? esZiel(tage, zielScore, zielDatum, extraProTag) : null;
  const fmtH = (h: number) => fmtDauer(h * 3600);
  const vz = (x: number | null) => (x == null ? '—' : `${x >= 0 ? '+' : '−'}${Math.abs(Math.round(x))}`);

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Endurance-Score-Hebel</div>
        <p className="mt-1 text-2xs text-ink-muted">
          Wie deine mittlere Tageslast und deine langen Einheiten deinen Garmin Endurance Score bewegen – mit deinem eigenen Fortschreibungsmodell.
        </p>
      </div>

      {/* Kennzahlen Raster 3 Spalten */}
      <div className="mt-3 grid grid-cols-1 sm:grid-cols-3 gap-2">
        {/* Endurance Score */}
        <div className="rounded-lg border border-border p-3">
          <p className="text-xs text-ink-muted">Endurance Score</p>
          <p className="mt-1 text-lg font-mono">{letzterES !== null ? Math.round(letzterES) : '—'}</p>
        </div>

        {/* Ø Last 28 Tage */}
        <div className="rounded-lg border border-border p-3">
          <p className="text-xs text-ink-muted">Ø Last 28 Tage</p>
          <p className="mt-1 text-lg font-mono">{last28 !== null ? Math.round(last28) : '—'} ESS/Tag</p>
        </div>

        {/* Gleichgewicht */}
        <div className="rounded-lg border border-border p-3">
          <p className="text-xs text-ink-muted">Gleichgewicht</p>
          <p className="mt-1 text-lg font-mono">
            {gleichgewicht !== null ? Math.round(gleichgewicht) : '—'} ESS/Tag
          </p>
          <p className="mt-1 text-2xs text-ink-muted">
            darüber steigt der Score, darunter sinkt er (bei deinem jetzigen Niveau)
          </p>
        </div>
      </div>

      {/* Belastbarkeit */}
      <div className="mt-3">
        <p className="text-xs text-ink-muted">
          Belastbarkeit: Geprüft an {backtest.n} Zeitpunkten: 28 Tage voraus im Schnitt ±{Math.round(backtest.maeModell)} Punkte daneben, „bleibt gleich" ±{Math.round(backtest.maeNaiv)}.
        </p>
        {notBetter && (
          <p className="mt-1 text-xs text-orange-300">
            Das Modell ist derzeit nicht besser als „bleibt gleich" – Werte nur als Richtung lesen.
          </p>
        )}
      </div>

      {/* Lange Einheiten */}
      <div className="mt-4">
        <p className="text-xs text-ink-muted mb-2">Lange Einheiten:</p>
        {!lang ? (
          <p className="text-xs text-ink-dim">Garmin-Tagesdetails laden…</p>
        ) : st && emp && eff ? (
          <>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              <div className={`rounded-lg border p-3 ${emp.fenster14Leer ? 'border-orange-400/60' : 'border-border'}`}>
                <p className="text-xs text-ink-muted">Längste der letzten 14 Tage</p>
                <p className="mt-1 text-lg font-mono">{st.laengste14 ? fmtH(st.laengste14.h) : '—'}</p>
                <p className="mt-1 text-2xs text-ink-muted">
                  {st.laengste14 ? `${sportDe(st.laengste14.sport)} am ${fmtDateShort(st.laengste14.date)}` : 'keine Aktivität'}
                  {emp.fenster14Leer ? ' – Garmins 2-Wochen-Fenster hat keine lange Einheit' : ''}
                </p>
              </div>
              <div className="rounded-lg border border-border p-3">
                <p className="text-xs text-ink-muted">Ab 2 h</p>
                <p className="mt-1 text-lg font-mono">{st.n2_28} / {st.n2_84}</p>
                <p className="mt-1 text-2xs text-ink-muted">in 28 / 84 Tagen (Ziel etwa 4 / 12)</p>
              </div>
              <div className="rounded-lg border border-border p-3">
                <p className="text-xs text-ink-muted">Ab 4 h</p>
                <p className="mt-1 text-lg font-mono">{st.n4_84}</p>
                <p className="mt-1 text-2xs text-ink-muted">
                  in 84 Tagen (Ziel etwa 6){st.letzte4 ? `, zuletzt ${fmtDateShort(st.letzte4.date)}` : ''}
                </p>
              </div>
            </div>

            <div className="mt-3">
              <p className="text-xs font-medium">Empfehlung</p>
              <ul className="mt-1 text-xs space-y-1 list-disc pl-4">
                {emp.punkte.map((p) => <li key={p}>{p}</li>)}
              </ul>
              <p className="mt-2 text-xs">
                Nächste lange Einheit ({String(emp.wocheDauerH).replace('.', ',')} h):{' '}
                <span className={emp.wocheUeberfaellig ? 'text-orange-300' : ''}>
                  {emp.wocheUeberfaellig ? `überfällig seit ${fmtDateShort(emp.wocheFaellig)}` : `bis ${fmtDateShort(emp.wocheFaellig)}`}
                </span>
                {' · '}Nächste sehr lange (ab 4 h):{' '}
                {emp.sehrLangAufbau ? (
                  <span className="text-ink-muted">nach dem Aufbau auf etwa 3 h</span>
                ) : (
                  <span className={emp.sehrLangUeberfaellig ? 'text-orange-300' : ''}>
                    {emp.sehrLangUeberfaellig ? `überfällig seit ${fmtDateShort(emp.sehrLangFaellig)}` : `bis ${fmtDateShort(emp.sehrLangFaellig)}`}
                  </span>
                )}
              </p>
            </div>

            <div className="mt-3">
              <p className="text-xs font-medium">Was lange Einheiten bei dir zusätzlich bringen</p>
              <p className="mt-1 text-2xs text-ink-muted">
                Score-Punkte in den Tagen 0–3 danach, über das hinaus, was ihre ESS-Last schon erklärt ({eff.basisTage} Tage geprüft):
              </p>
              <div className="mt-1 grid grid-cols-3 gap-2 text-xs">
                {eff.klassen.map((k) => (
                  <div key={k.lab} className="rounded border border-border px-2 py-1">
                    <span className="text-ink-muted">{k.lab}: </span>
                    <span className="font-mono">{vz(k.extra)}</span>
                    <span className="text-2xs text-ink-muted"> (n={k.n})</span>
                  </div>
                ))}
              </div>
              {eff.n4 < 8 && (
                <p className="mt-1 text-2xs text-orange-300">
                  Erst {eff.n4} sehr lange Einheiten mit Score-Verlauf – Größenordnung, keine feste Zahl. Gerechnet wird vorsichtig mit {vz(eff.bonus4)} je Einheit.
                </p>
              )}
            </div>
          </>
        ) : null}
      </div>

      {/* Ziel-Rechner */}
      <div ref={zielRef} className="mt-4">
        <p className="text-xs text-ink-muted mb-2">Ziel-Rechner:</p>

        <div className="flex gap-3 mb-2 flex-wrap">
          {/* Ziel-Score */}
          <div>
            <label htmlFor="es-ziel-score" className="text-2xs text-ink-muted">Ziel-Score</label>
            <input
              id="es-ziel-score"
              type="number"
              value={zielScore}
              onChange={(e) => setZielScore(Number(e.target.value) || 7000)}
              className="ml-1 px-2 py-1 border border-border rounded text-xs w-32"
            />
          </div>

          {/* Zieldatum */}
          <div>
            <label htmlFor="es-ziel-datum" className="text-2xs text-ink-muted">Zieldatum</label>
            <input
              id="es-ziel-datum"
              type="date"
              value={zielDatum}
              onChange={(e) => setZielDatum(e.target.value)}
              className="ml-1 px-2 py-1 border border-border rounded text-xs w-36"
            />
          </div>
        </div>

        {/* Ergebnis */}
        {z !== null && z.tage > 0 && (
          <div className="mt-2">
            {z.machbar && z.ladeProTag !== null ? (
              <p className="text-xs">
                Für {ziel} bis {fmtDateShort(datum)} brauchst du im Schnitt etwa {Math.round(z.ladeProTag)} ESS pro Tag ({Math.round(z.ladeProTag * 7).toLocaleString("de-DE")} pro Woche) über {z.tage} Tage.
              </p>
            ) : (
              <p className="text-xs text-orange-300">
                Mit bis zu 250 ESS pro Tag nicht erreichbar – Ziel oder Datum anpassen.
              </p>
            )}
            {zLang && zLang.machbar && zLang.ladeProTag !== null && z.ladeProTag !== null && (
              <p className="mt-1 text-xs">
                Mit 2 sehr langen Einheiten (ab 4 h) pro Monat reichen etwa {Math.round(zLang.ladeProTag)} ESS pro Tag
                {z.ladeProTag - zLang.ladeProTag >= 1 ? ` (${Math.round(z.ladeProTag - zLang.ladeProTag)} weniger)` : ''}.
              </p>
            )}
            {last28 !== null && (
              <p className="mt-1 text-xs text-ink-muted">Aktuell: {Math.round(last28)} ESS/Tag.</p>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
