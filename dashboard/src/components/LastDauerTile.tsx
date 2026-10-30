import { useMemo, useState } from 'react';
import { fmtNum } from '../lib/format';
import { useGarminTage, fmtDauer } from '../lib/garmin';
import { sportProfile, lastJeStunde } from './GarminSportProfilTile';

/** Übersetzt eine geplante Last in eine voraussichtliche Dauer je Sport (Garmin-Belastung je Stunde, letzte 120 Tage). */
export function LastDauerTile({ vorschlag }: { vorschlag?: number }) {
  const { tage } = useGarminTage();
  const [last, setLast] = useState<string>(vorschlag ? String(Math.round(vorschlag)) : '150');
  const profile = useMemo(() => sportProfile(tage, 120).filter((p) => lastJeStunde(p) != null && p.n >= 2), [tage]);
  const l = Number(last.replace(',', '.'));
  return (
    <div className="panel-raised p-4" data-testid="last-dauer">
      <div className="label">Coach Kira · Last in Dauer</div>
      <p className="mt-1 text-2xs text-ink-muted">
        Wie lange eine geplante Last je Sport ungefähr dauert – aus deiner Garmin-Belastung je Stunde der letzten 120 Tage (Durchschnitt deiner üblichen Intensität).
      </p>
      <div className="mt-2 flex items-center gap-2 text-xs">
        <span className="text-ink-muted">Geplante Last</span>
        <input className="input w-24 tnum" inputMode="numeric" value={last} onChange={(e) => setLast(e.target.value)} />
      </div>
      {tage === null ? (
        <div className="mt-2 text-sm text-ink-muted">Lade Garmin-Daten …</div>
      ) : !profile.length || !Number.isFinite(l) || l <= 0 ? (
        <div className="mt-2 text-sm text-ink-muted">Keine Daten oder keine gültige Last.</div>
      ) : (
        <div className="mt-2 grid gap-1 sm:grid-cols-2 text-xs tnum">
          {profile.map((p) => {
            const lph = lastJeStunde(p)!;
            const sek = (l / lph) * 3600;
            const kmh = p.sek ? p.km / (p.sek / 3600) : 0;
            return (
              <div key={p.sport} className="flex items-baseline gap-2">
                <span className="font-semibold w-14">{p.sport}</span>
                <span>{fmtDauer(sek)}</span>
                {kmh > 0 && <span className="text-ink-muted">≈ {fmtNum((sek / 3600) * kmh)} km</span>}
                <span className="text-ink-dim">({fmtNum(lph)}/h)</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
