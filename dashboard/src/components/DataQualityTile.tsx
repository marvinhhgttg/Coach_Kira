import { useState, useEffect } from 'react';
import { fetchChartData } from '../lib/api';
import { pruefeQualitaet, PLATZHALTER } from '../lib/dataQuality';

function fmtPct(n: number): string {
  return new Intl.NumberFormat('de-DE', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(Math.round(n * 100)) + ' %';
}

function fmtNum(n: number, decimals = 0): string {
  return new Intl.NumberFormat('de-DE', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(n);
}

const FELDER = [
  { key: 'load_fb_day', label: 'Last (ESS)' },
  { key: 'sleep_hours', label: 'Schlaf h' },
  { key: 'sleep_score_0_100', label: 'Schlafscore' },
  { key: 'rhr_bpm', label: 'Ruhepuls' },
  { key: 'hrv_status', label: 'HRV' },
  { key: 'hrv_threshholds', label: 'HRV-Band' },
  { key: 'Garmin_Training_Readiness', label: 'Readiness' },
  { key: 'garminEnduranceScore', label: 'Endurance Score' },
  { key: 'kcal_in', label: 'kcal gegessen' },
  { key: 'protein_g', label: 'Protein g' },
  { key: 'elev_m_day', label: 'Höhenmeter' },
  { key: 'befinden_1_5', label: 'Befinden' },
  { key: 'Sport_x', label: 'Sport' },
  { key: 'Zone', label: 'Zone' },
] as const;

export function DataQualityTile() {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [qualitaet, setQualitaet] = useState<{ felder: ReturnType<typeof pruefeQualitaet>['felder']; doppelteDaten: string[]; fehlendeTage: string[] } | null>(null);

  useEffect(() => {
    let cancelled = false;
    const heute = new Date().toLocaleDateString('sv-SE');

    fetchChartData('9999d', FELDER.map(f => f.key))
      .then(r => {
        if (cancelled) return;
        const q = pruefeQualitaet(r.data, FELDER, heute);
        setQualitaet({ felder: q.felder, doppelteDaten: q.doppelteDaten, fehlendeTage: q.fehlendeTage });
        setLoading(false);
      })
      .catch(e => {
        if (cancelled) return;
        setError(e.message || 'Daten konnten nicht geladen werden');
        setLoading(false);
      });

    return () => { cancelled = true; };
  }, []);

  if (loading) return <div className="panel-raised p-4">Lade Datenqualität …</div>;
  if (error) return <div className="panel-raised p-4 text-orange-300">{error}</div>;
  if (!qualitaet) return <div className="panel-raised p-4">Keine Daten verfügbar.</div>;

  const { felder, doppelteDaten, fehlendeTage } = qualitaet;

  const urteilColor: Record<string, string> = {
    belastbar: 'text-green-300',
    eingeschränkt: 'text-yellow-300',
    'zu wenig Daten': 'text-orange-300',
  };

  const placeholderInfo = Object.entries(PLATZHALTER).map(([feld, werte]) =>
    `${feld} = ${werte.map(v => fmtNum(v)).join(', ')}`
  );

  return (
    <div className="panel-raised p-4">
      <div>
        <div className="label">Coach Kira · Datenqualität</div>
        <p className="text-2xs text-ink-dim mt-1">
          Füllgrad je Feld, Platzhalter und Lücken – Analysen blenden Platzhalter-Tage aus.
        </p>
      </div>

      <div className="mt-3 overflow-x-auto">
        <table className="w-full text-xs" cellSpacing={0}>
          <thead>
            <tr className="text-ink-dim">
              <th className="text-left font-normal p-1 pb-2">Feld</th>
              <th className="text-center font-normal p-1 pb-2">30 Tage</th>
              <th className="text-center font-normal p-1 pb-2">90 Tage</th>
              <th className="text-center font-normal p-1 pb-2">gesamt</th>
              <th className="text-center font-normal p-1 pb-2">Platzhalter</th>
              <th className="text-center font-normal p-1 pb-2">längste Serie</th>
              <th className="text-center font-normal p-1 pb-2">Urteil</th>
            </tr>
          </thead>
          <tbody>
            {felder.map((f, i) => (
              <tr key={i} className="border-t border-border">
                <td className="p-1 text-ink">{f.label}</td>
                <td className="p-1 text-center tnum">{fmtPct(f.quote30)}</td>
                <td className="p-1 text-center tnum">{fmtPct(f.quote90)}</td>
                <td className="p-1 text-center tnum">{fmtPct(f.quoteAlle)}</td>
                <td className="p-1 text-center tnum">{f.platzhalter}</td>
                <td className="p-1 text-center tnum">{f.laengsteSerie}</td>
                <td className={`p-1 text-center ${urteilColor[f.urteil] || ''}`}>{f.urteil}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="mt-3 text-xs space-y-1">
        <p>
          <span className="text-ink-dim">Doppelte Daten:</span>{' '}
          {doppelteDaten.length > 0 ? doppelteDaten.join(', ') : 'keine'}
        </p>
        <p>
          <span className="text-ink-dim">Fehlende Tage: {fehlendeTage.length}</span>
          {fehlendeTage.length > 0 && fehlendeTage.length <= 5 && (
            <><span className="text-ink-dim"> ({fehlendeTage.map(d => {
              const p = d.split('-');
              return `${p[2]}.${p[1]}`;
            }).join(', ')})</span></>
          )}
        </p>
        {placeholderInfo.length > 0 && (
          <p className="text-ink-dim">
            Platzhalter: {placeholderInfo.join(' wird als nicht erfasst gewertet. ')} wird als nicht erfasst gewertet.
          </p>
        )}
      </div>
    </div>
  );
}
