import { useState, useEffect } from 'react';
import { fetchChartData } from '../lib/api';

export function BefindenHinweis({ erfasst }: { erfasst: boolean }) {
  const [anzahl, setAnzahl] = useState<number | null>(null);

  useEffect(() => {
    fetchChartData('9999d', ['befinden_1_5'])
      .then(r => {
        const gez = r.data.filter(row => {
          const v = row.befinden_1_5;
          return v !== null && v !== undefined && String(v).trim() !== '';
        }).length;
        setAnzahl(gez);
      })
      .catch(() => {
        // Fehler still schlucken – harmless notice
      });
  }, []);

  if (!erfasst) {
    return (
      <div className="rounded-lg border border-yellow-400/40 p-2 text-xs text-yellow-200">
        <p>Wie fühlst du dich heute? Ein Tipp auf 1–5 reicht. Befinden ist das einzige Signal, das Garmin nicht kennt.</p>
        {anzahl !== null && (
          <p className="mt-1 text-2xs text-ink-dim">Befinden: {anzahl} Werte bisher – ab 90 wird es auswertbar.</p>
        )}
      </div>
    );
  }

  // erfasst === true
  if (anzahl !== null) {
    return (
      <p className="text-2xs text-ink-dim">
        Befinden: {anzahl} Werte bisher{anzahl >= 90 ? ' – auswertbar' : ' – ab 90 wird es auswertbar'}
      </p>
    );
  }

  return null;
}
