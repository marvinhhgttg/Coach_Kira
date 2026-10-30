import { useEffect, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { fetchExtras } from '../lib/extras';
import { ALL_RANGE } from '../lib/windows';
import { prognoseCheck, type CheckErgebnis, type Fehler, type IstTag, type PrognoseItem } from '../lib/prognoseCheck';
import { fmtDateShort } from '../lib/format';

const METR = ['load_fb_day', 'fbCTL_obs', 'garminEnduranceScore', 'Garmin_Training_Readiness'];
const zahl = (v: unknown) => (v === '' || v == null || !Number.isFinite(Number(v)) ? null : Number(v));
const f1 = (v: number) => v.toLocaleString('de-DE', { maximumFractionDigits: 1, minimumFractionDigits: 0 });
const vz = (v: number) => `${v > 0 ? '+' : v < 0 ? '−' : '±'}${f1(Math.abs(v))}`;

let cache: Promise<CheckErgebnis | null> | null = null;
export function ladePrognoseCheck(): Promise<CheckErgebnis | null> {
  if (!cache)
    cache = Promise.all([fetchExtras<any>('prognose', 1000), fetchChartData(ALL_RANGE, METR)])
      .then(([items, res]: [any[], any]) => {
        const ist: IstTag[] = ((res && res.data) || []).map((r: any) => ({
          date: String(r.date),
          load: zahl(r.load_fb_day),
          ctl: zahl(r.fbCTL_obs),
          es: zahl(r.garminEnduranceScore),
          ready: zahl(r.Garmin_Training_Readiness),
        }));
        ist.sort((a, b) => a.date.localeCompare(b.date));
        return prognoseCheck(items as PrognoseItem[], ist, new Date().toLocaleDateString('sv-SE'));
      })
      .catch(() => {
        cache = null;
        return null;
      });
  return cache;
}

/** Urteil: besser als „bleibt wie heute“? */
function urteil(f: Fehler): { t: string; c: string } {
  if (!f.n) return { t: 'noch keine Fälle', c: 'text-ink-dim' };
  if (f.n < 5) return { t: 'zu wenige Fälle', c: 'text-ink-dim' };
  if (f.mae <= f.maeNaiv * 0.9) return { t: 'trifft besser als „bleibt gleich“', c: 'text-green-600 dark:text-green-400' };
  if (f.mae <= f.maeNaiv * 1.1) return { t: 'nicht besser als „bleibt gleich“', c: 'text-amber-600 dark:text-amber-300' };
  return { t: 'schlechter als „bleibt gleich“', c: 'text-orange-600 dark:text-orange-300' };
}

function Zeile({ was, h, f, extra }: { was: string; h: string; f: Fehler; extra?: string }) {
  const u = urteil(f);
  return (
    <tr className="border-b border-border">
      <td className="pt-1">{was}</td>
      <td className="pt-1">{h}</td>
      <td className="pt-1 tnum">{f.n}</td>
      <td className="pt-1 tnum">{f.n ? `±${f1(f.mae)}` : '–'}</td>
      <td className="pt-1 tnum">{f.n ? vz(f.bias) : '–'}</td>
      <td className="pt-1 tnum">{f.n ? `±${f1(f.maeNaiv)}` : '–'}</td>
      <td className={`pt-1 ${u.c}`}>{u.t}{extra ? <span className="text-ink-dim"> · {extra}</span> : null}</td>
    </tr>
  );
}

/** Prognose-Check: Sind Fitness-, Endurance-Score-, Bereitschafts- und Qualitätstag-Prognosen eingetroffen? */
export function PrognoseCheckTile() {
  const [c, setC] = useState<CheckErgebnis | null | undefined>(undefined);
  useEffect(() => {
    let off = false;
    ladePrognoseCheck().then((x) => !off && setC(x));
    return () => {
      off = true;
    };
  }, []);
  const H = ['7', '14', '28'] as const;
  return (
    <div className="mt-6" data-testid="prognose-check">
      <div className="label">Prognose-Check – ist es so gekommen?</div>
      <p className="text-2xs text-ink-dim mt-1">
        Zwei Blickwinkel: Gespeicherte Prognosen (so, wie der Mac sie jeden Tag erstellt hat, mit dem damaligen Plan) gegen das, was wirklich kam – und das reine Modell, rückgerechnet mit deinen echten Lasten über die ganze Historie.
        Der Unterschied zeigt, wie viel Fehler vom Modell und wie viel von Planänderungen kommt. Vergleich jeweils mit „bleibt wie heute“. Fest berechnet, ohne KI.
      </p>
      {c === undefined && <div className="text-2xs text-ink-dim mt-2">lädt …</div>}
      {c === null && <div className="text-2xs text-ink-dim mt-2">Daten nicht erreichbar.</div>}
      {c && (
        <>
          <table className="w-full text-2xs mt-3" data-testid="prognose-check-tabelle">
            <thead className="text-ink-dim text-left">
              <tr className="border-b border-border">
                <th className="pb-1">Prognose</th>
                <th>Horizont</th>
                <th>Fälle</th>
                <th>Fehler Ø</th>
                <th>Richtung</th>
                <th>„bleibt gleich“</th>
                <th>Urteil</th>
              </tr>
            </thead>
            <tbody>
              {H.map((h) => {
                const f = c.echt.fitness[h];
                const extra = f.n && (f.lastPlan || f.lastIst) ? `Plan ${Math.round(f.lastPlan)} / Ist ${Math.round(f.lastIst)} ESS` : undefined;
                return <Zeile key={'fe' + h} was="Fitness (CTL) · gespeichert" h={`${h} Tage`} f={f} extra={extra} />;
              })}
              {H.map((h) => <Zeile key={'fm' + h} was="Fitness (CTL) · Modell mit Ist-Last" h={`${h} Tage`} f={c.modell.fitness[h]} />)}
              {c.echt.es && H.map((h) => <Zeile key={'ee' + h} was="Endurance Score · gespeichert" h={`${h} Tage`} f={c.echt.es![h]} />)}
              {H.map((h) => <Zeile key={'em' + h} was="Endurance Score · Modell mit Ist-Last" h={`${h} Tage`} f={c.modell.es[h]} />)}
              {(['1', '2', '3'] as const).map((h) => <Zeile key={'b' + h} was="Bereitschaft · gespeichert" h={`${h} Tag${h === '1' ? '' : 'e'}`} f={c.echt.bereitschaft[h]} />)}
            </tbody>
          </table>
          <div className="mt-3 text-2xs" data-testid="prognose-check-qualitaet">
            <span className="font-semibold">Qualitätstag / Tourfenster:</span>{' '}
            {c.echt.qualitaetstag.n
              ? `${c.echt.qualitaetstag.eingetroffen} von ${c.echt.qualitaetstag.n} eingetroffen (Bereitschaft am Tag ≥ 55 und höchstens 10 Punkte unter der Prognose).`
              : 'noch keine abgelaufenen Vorschläge.'}
            {c.echt.qualitaetstag.faelle.length > 0 && (
              <span className="text-ink-dim">
                {' '}
                {c.echt.qualitaetstag.faelle
                  .slice(-6)
                  .map((q) => `${fmtDateShort(q.fuer)}: ${q.prog ?? '–'} → ${q.ist ?? '–'} ${q.ok ? 'ja' : 'nein'}`)
                  .join(' · ')}
              </span>
            )}
          </div>
          <p className="mt-2 text-2xs text-ink-dim">
            Fehler Ø = mittlere Abweichung in Punkten der jeweiligen Skala (Fitness wie im Sheet, Endurance Score wie bei Garmin). Richtung: + = Prognose zu hoch.
            Gespeicherte Prognosen gibt es seit {c.seit ? fmtDateShort(c.seit) : '–'}; die Zahl der Fälle wächst mit jedem Tag.
          </p>
        </>
      )}
    </div>
  );
}
