import { type ReadinessForecast, type Model, explainModel } from '../lib/readinessModel';
import { fmtNum } from '../lib/format';

const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

function dateShort(d: string): string {
  const dp = d.length >= 10 ? d.slice(0, 10) : d;
  const parts = dp.split('-');
  if (parts.length < 3) return d;
  const dayOfWeek = WD[new Date(d + 'T12:00:00').getDay()];
  return `${dayOfWeek} ${parseInt(parts[2])}.${parts[1].padStart(2, '0')}`;
}

function readinessColor(r: number): string {
  if (r >= 70) return 'text-green-300';
  if (r >= 50) return 'text-yellow-300';
  return 'text-orange-300';
}

export function ReadinessStrip({
  forecast,
  model,
  error,
}: {
  forecast: ReadinessForecast[];
  model: Model | null;
  error: string | null;
}) {
  return (
    <div className="panel-raised p-4">
      <div>
        <div className="label">Bereitschafts-Prognose</div>
        {model ? (
          <p className="text-2xs text-ink-dim mt-1">
            Gelernt aus {model.n} Tagen deiner Daten · mittlerer Fehler ±{fmtNum(model.mae, 0)} Punkte (naiv „wie heute“: ±{fmtNum(model.baselineMae, 0)})
          </p>
        ) : (
          <p className="text-2xs text-ink-dim mt-1">Zu wenig Daten für ein Modell.</p>
        )}
      </div>

      {error && <p className="mt-2 text-xs text-orange-300">{error}</p>}

      {forecast && forecast.length > 0 && (
        <div className="mt-3 flex gap-1 overflow-x-auto">
          {forecast.map((f, i) => (
            <div key={i} className="flex flex-col items-center min-w-[4rem]">
              <span className="text-2xs text-ink-dim">{dateShort(f.date)}</span>
              <span className={`tnum text-base ${readinessColor(f.readiness)}`}>
                {f.readiness}
              </span>
              <span className="text-2xs text-ink-dim">
                Band {f.band[0]}–{f.band[1]}
              </span>
            </div>
          ))}
        </div>
      )}

      {forecast && forecast.length > 0 && (
        <p className="mt-2 text-2xs text-ink-dim">Morgenwert nach dem jeweiligen Plantag.</p>
      )}

      {model && (
        <div className="mt-2">
          <p className="text-2xs text-ink-dim">
            Wichtigste Einflüsse: {explainModel(model).join(' · ')}
          </p>
        </div>
      )}
    </div>
  );
}
