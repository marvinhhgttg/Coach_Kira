import type { PlanDay } from '../lib/api';
import { fmtNum, fmtDateShort } from '../lib/format';

export function SupervisorPlanTile({
  days,
  lockedDates,
  status,
  busy,
  onApply,
  onRefresh,
}: {
  days: PlanDay[];
  lockedDates: Set<string>;
  status: string | null;
  busy: boolean;
  onApply: () => void;
  onRefresh: () => void;
}) {
  const hasDays = days.length > 0;

  return (
    <div className="panel-raised p-4 border-accent-strong/20">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Wochenplan</div>
          <p className="text-2xs text-ink-dim mt-1">
            Aus dem KI-Report (Supervisor) ·{' '}
            {hasDays
              ? days.length + ' Tage ab ' + fmtDateShort(days[0].date)
              : 'noch kein Plan'}
            {status ? ' · ' + status : ''}
          </p>
        </div>
        {hasDays && (
          <div className="flex flex-wrap gap-2">
            <button
              className="btn btn-ghost text-xs px-2 py-1"
              onClick={onApply}
            >
              ESS übernehmen
            </button>
            <button
              className="btn btn-primary text-xs px-2 py-1"
              onClick={onRefresh}
              disabled={busy}
            >
              {busy ? 'KI-Report läuft…' : 'Neu berechnen'}
            </button>
          </div>
        )}
      </div>

      {!hasDays && (
        <p className="mt-3 text-sm text-ink-muted">
          Noch kein Wochenplan vorhanden. „Neu berechnen“ startet den KI-Report;
          lokal dauert das einige Minuten.
        </p>
      )}

      {hasDays && (
        <div className="mt-4 grid gap-2 md:grid-cols-2 xl:grid-cols-4">
          {days.map((d) => {
            const locked = lockedDates.has(d.date);
            const diff =
              typeof d.recommended_load_ess === 'number' &&
              typeof d.original_load_ess === 'number' &&
                  d.recommended_load_ess !== d.original_load_ess
                ? d.recommended_load_ess - d.original_load_ess
                : null;
            return (
              <div
                key={d.date}
                className="rounded border border-border bg-bg-subtle/50 p-3 flex flex-col gap-1"
              >
                <div className="flex items-center justify-between">
                  <span className="tnum text-xs text-ink">
                    {d.day} {fmtDateShort(d.date)}
                  </span>
                  {locked && (
                    <span className="chip border border-border-strong text-2xs text-ink-muted">
                      fest
                    </span>
                  )}
                </div>

                <div className="tnum">
                  <span className="text-lg font-semibold text-accent">
                    {fmtNum(d.recommended_load_ess ?? 0)} ESS
                  </span>{' '}
                  <span className="text-2xs text-ink-dim">
                    Plan {fmtNum(d.original_load_ess ?? 0)}
                  </span>
                  {diff !== null && (
                    <span
                      className={`text-2xs ${
                        diff < 0 ? 'text-green-300' : 'text-orange-300'
                      }`}
                    >
                      ({diff < 0 ? '' : '+'}
                      {fmtNum(diff)})
                    </span>
                  )}
                </div>

                <div className="text-2xs text-ink-muted">
                  {d.recommended_zone || 'Zone offen'}
                </div>

                {d.training_1 && (
                  <div className="text-xs text-ink">{d.training_1}</div>
                )}
                {d.training_2 && (
                  <div className="text-2xs text-ink-dim">
                    Alternative: {d.training_2}
                  </div>
                )}
                {d.projected_effect && (
                  <p className="text-2xs text-ink-muted leading-snug">
                    {d.projected_effect}
                  </p>
                )}
                {d.weather_recommendation && (
                  <p className="text-2xs text-ink-dim leading-snug">
                    Wetter: {d.weather_recommendation}
                  </p>
                )}
              </div>
            );
          })}
        </div>
      )}

      {hasDays && (
        <p className="mt-3 text-2xs text-ink-dim">
          „ESS übernehmen“ setzt Last und Zone der freien Tage in die Simulation;
          gesperrte Tage bleiben. Gespeichert wird erst mit „In timeline
          speichern“.
        </p>
      )}
    </div>
  );
}
