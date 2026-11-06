import { ZONEN_FARBEN, zonenProzent } from '../lib/garmin';

/** Gestapelter HF-Zonenbalken (Sekunden je Zone Z1..Z5). */
export function GarminZonenBalken({ zonen, hoehe = 8, beschriftung = false }: { zonen?: number[] | null; hoehe?: number; beschriftung?: boolean }) {
  const pct = zonenProzent(zonen);
  if (!pct || !zonen) return null;
  return (
    <div>
      <div className="flex w-full overflow-hidden rounded-full bg-bg-subtle" style={{ height: hoehe }}>
        {pct.map((p, i) =>
          p > 0 ? (
            <div
              key={i}
              style={{ width: `${p}%`, background: ZONEN_FARBEN[i] }}
              title={`Z${i + 1}: ${Math.round((zonen[i] || 0) / 60)} min (${p} %)`}
            />
          ) : null,
        )}
      </div>
      {beschriftung && (
        <div className="mt-0.5 text-2xs text-ink-dim tnum">
          {pct.map((p, i) => (p > 0 ? `Z${i + 1} ${p} %` : null)).filter(Boolean).join(' · ')}
        </div>
      )}
    </div>
  );
}
