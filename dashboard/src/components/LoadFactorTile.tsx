import { fmtNum } from '../lib/format';
import type { FaktorVorschlag } from '../lib/planning';

export function LoadFactorTile({
  vorschlaege,
  busy,
  error,
  onRefresh,
}: {
  vorschlaege: FaktorVorschlag[] | null;
  busy: boolean;
  error: string | null;
  onRefresh: () => void;
}) {
  return (
    <div className="panel-raised p-4">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Lastfaktoren</div>
          <p className="text-2xs text-ink-dim mt-1">
            Vergleich geplante vs. tatsächliche Last je Sport und Zone (Plan-Schnappschuss am Morgen gegen Garmin-Ist). Nur Vorschläge – KK_LOAD_FACTORS änderst du selbst.
          </p>
        </div>
        <button
          className="btn btn-primary text-xs px-2 py-1"
          onClick={onRefresh}
          disabled={busy}
        >
          {busy ? 'Rechne…' : 'Neu berechnen'}
        </button>
      </div>

      {error && <p className="mt-2 text-xs text-orange-300">{error}</p>}

      {vorschlaege && vorschlaege.length > 0 && (
        <div className="mt-3 overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="border-b border-border text-ink-dim">
                <th className="py-1 pr-3 font-normal">Sport</th>
                <th className="py-1 pr-3 font-normal">Zone</th>
                <th className="py-1 pr-3 font-normal text-right">Tage (n)</th>
                <th className="py-1 pr-3 font-normal text-right">Ist/Plan</th>
                <th className="py-1 pr-3 font-normal text-right">Faktor (ESS/h)</th>
                <th className="py-1 font-normal text-right">Vorschlag</th>
              </tr>
            </thead>
            <tbody>
              {vorschlaege.map((f, i) => {
                const pct = Math.round(f.verhaeltnis * 100);
                // Vorschlag nur wenn n >= 4 und |verhaeltnis - 1| > 0.08
                const showVorschlag = f.n >= 4 && Math.abs(f.verhaeltnis - 1) > 0.08;
                const dev = Math.abs(f.verhaeltnis - 1);
                return (
                  <tr key={i} className="border-b border-border/60">
                    <td className="py-1 pr-3">{f.typ}</td>
                    <td className="py-1 pr-3">{f.zone}</td>
                    <td className="py-1 pr-3 text-right tnum">{f.n}</td>
                    <td className="py-1 pr-3 text-right">
                      <span
                        className={
                          dev > 0.15
                            ? 'text-orange-300 tnum'
                            : dev > 0.08
                              ? 'tnum'
                              : 'tnum'
                        }
                      >
                        {pct} %
                      </span>
                    </td>
                    <td className="py-1 pr-3 text-right tnum">{f.faktor}</td>
                    <td className="py-1 text-right">
                      {showVorschlag ? (
                        <span className="tnum">{f.vorschlag}</span>
                      ) : (
                        <span title="zu wenig Daten oder Abweichung klein">–</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {vorschlaege && vorschlaege.length === 0 && !busy && (
        <p className="mt-3 text-xs text-ink-muted">Keine Vorschläge vorhanden.</p>
      )}

      <p className="mt-2 text-2xs text-ink-dim">
        Lesart: 115 % heißt, diese Einheiten erzeugen bei dir im Mittel 15 % mehr ESS als geplant.
        Ursachen können auch längere Dauer oder andere Zonen sein – deshalb erst ab 4 Tagen ein Vorschlag.
      </p>
    </div>
  );
}
