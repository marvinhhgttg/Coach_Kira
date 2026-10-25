import { useMemo, useState } from 'react';
import { Panel } from './UI';
import { fmtNum } from '../lib/format';
import { useGarminTage, fmtDauer, type GarminTage } from '../lib/garmin';
import { WINDOWS, winLabel, type WinKey } from '../lib/windows';

export type SportProfil = {
  sport: string; n: number; sek: number; km: number; hm: number; last: number; lastSek: number;
  bb: number; bbSek: number; hf: number; hfN: number; te: number; teN: number;
};

/** Kennzahlen je Sport aus den Garmin-Aktivitäten (Fenster in Tagen, 0 = alle). */
export function sportProfile(tage: GarminTage | null, win: number): SportProfil[] {
  if (!tage) return [];
  const von = win === 0 ? '0000' : new Date(Date.now() - win * 86400000).toLocaleDateString('sv-SE');
  const m = new Map<string, SportProfil>();
  for (const [d, g] of tage) {
    if (d < von) continue;
    for (const a of g.akt || []) {
      const sport = a.sport || 'Sonstige';
      const p = m.get(sport) || { sport, n: 0, sek: 0, km: 0, hm: 0, last: 0, lastSek: 0, bb: 0, bbSek: 0, hf: 0, hfN: 0, te: 0, teN: 0 };
      const sek = a.bewegt_s || a.dauer_s || 0;
      p.n++; p.sek += sek; p.km += a.km || 0; p.hm += a.hm || 0;
      if (a.last != null && sek > 0) { p.last += a.last; p.lastSek += sek; }
      if (a.bb != null && sek > 0) { p.bb += a.bb; p.bbSek += sek; }
      if (a.hf_avg != null) { p.hf += a.hf_avg; p.hfN++; }
      if (a.te_ae != null) { p.te += a.te_ae; p.teN++; }
      m.set(sport, p);
    }
  }
  return [...m.values()].sort((a, b) => b.sek - a.sek);
}
/** Garmin-Belastung je Stunde für einen Sport (oder null). */
export function lastJeStunde(p?: SportProfil): number | null {
  return p && p.lastSek >= 1800 ? p.last / (p.lastSek / 3600) : null;
}

export function GarminSportProfilTile() {
  const { tage } = useGarminTage();
  const [win, setWin] = useState<WinKey>(120);
  const rows = useMemo(() => sportProfile(tage, win), [tage, win]);
  return (
    <Panel
      title="Sportprofil · Garmin"
      right={
        <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5">
          {WINDOWS.map((w) => (
            <button key={w} onClick={() => setWin(w)} className={`px-1.5 sm:px-2 py-1 text-xs rounded tnum ${win === w ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}>{winLabel(w)}</button>
          ))}
        </div>
      }
    >
      <p className="text-2xs text-ink-muted mb-2">
        Je Sport aus allen Garmin-Aktivitäten: Umfang, Dichte (Garmin-Belastung je Stunde) und Kosten (Body Battery je Stunde). Damit lässt sich eine geplante Last in eine Dauer übersetzen.
      </p>
      {tage === null ? (
        <div className="text-sm text-ink-muted">Lade Garmin-Daten …</div>
      ) : !rows.length ? (
        <div className="text-sm text-ink-muted">Keine Garmin-Aktivitäten im Zeitraum.</div>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full min-w-[760px] text-xs tnum" data-testid="sportprofil">
            <thead>
              <tr className="text-ink-dim text-left">
                {['Sport', 'Einh.', 'Σ Zeit', 'Σ km', 'Σ hm', 'Ø Dauer', 'Ø km', 'hm/km', 'Last/h', 'BB/h', 'Ø HF', 'Ø TE'].map((h) => (
                  <th key={h} className="label py-1 pr-2 font-normal">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((p) => {
                const lph = lastJeStunde(p);
                const bbh = p.bbSek >= 1800 ? p.bb / (p.bbSek / 3600) : null;
                return (
                  <tr key={p.sport} className="border-t border-border">
                    <td className="py-1.5 pr-2 font-semibold">{p.sport}</td>
                    <td className="py-1.5 pr-2">{p.n}</td>
                    <td className="py-1.5 pr-2">{fmtNum(p.sek / 3600, 1)} h</td>
                    <td className="py-1.5 pr-2">{fmtNum(p.km)}</td>
                    <td className="py-1.5 pr-2">{fmtNum(p.hm)}</td>
                    <td className="py-1.5 pr-2 whitespace-nowrap">{fmtDauer(p.sek / p.n)}</td>
                    <td className="py-1.5 pr-2">{fmtNum(p.km / p.n, 1)}</td>
                    <td className="py-1.5 pr-2">{p.km >= 1 ? fmtNum(p.hm / p.km, 1) : '—'}</td>
                    <td className="py-1.5 pr-2">{lph != null ? fmtNum(lph) : '—'}</td>
                    <td className={`py-1.5 pr-2 ${bbh != null && bbh <= -15 ? 'text-orange-300' : ''}`}>{bbh != null ? fmtNum(bbh) : '—'}</td>
                    <td className="py-1.5 pr-2">{p.hfN ? fmtNum(p.hf / p.hfN) : '—'}</td>
                    <td className="py-1.5 pr-2">{p.teN ? fmtNum(p.te / p.teN, 1) : '—'}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </Panel>
  );
}
