import { useGarminTage, fmtDauer, STRESS_Q_DE, teLabelDe } from '../lib/garmin';
import { Panel } from './UI';
import { GarminZonenBalken } from './GarminZonenBalken';
import { fmtNum, fmtDateShort } from '../lib/format';
import type { ReactNode } from 'react';

/** Datum als sv-SE String (yyyy-mm-dd), gestern. */
function yesterdayStr(): string {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return d.toLocaleDateString('sv-SE');
}

/* ------------------------------------------------------------------ */
/* Hilfsformatierer                                                    */
/* ------------------------------------------------------------------ */

/** "+61" / "−60" Text für Body Battery Laden/Verbrauchen. */
function fmtSignedNum(v: number | null | undefined): string {
  if (v == null || !Number.isFinite(v)) return '—';
  const sign = v > 0 ? '+' : '';
  return `${sign}${fmtNum(v)}`;
}

/* ------------------------------------------------------------------ */
/* Wrapper mit useGarminTage                                           */
/* ------------------------------------------------------------------ */

import type { GarminTag, GarminTage } from '../lib/garmin';

function YesterdayContentWrapper({
  datum,
  children,
}: {
  datum: string;
  children: (tage: GarminTage) => ReactNode;
}) {
  const { tage } = useGarminTage();
  if (tage === null) return <div className="text-sm text-ink-dim">Lade Garmin-Daten …</div>;
  return <>{children(tage)}</>;
}

function NoDataText({ datum }: { datum: string }) {
  return <div className="text-sm text-ink-dim">Noch keine Garmin-Daten für {fmtDateShort(datum)}.</div>;
}

/* ------------------------------------------------------------------ */
/* Hauptkomponente – Gestern-Kachel                                    */
/* ------------------------------------------------------------------ */

export function GarminGesternTile() {
  return <YesterdayPanelInner datum={yesterdayStr()} />;
}

function YesterdayPanelInner({ datum }: { datum: string }) {
  return (
    <YesterdayContentWrapper datum={datum}>
      {(tage) => {
        const g = tage.get(datum);
        if (!g) return <NoDataText datum={datum} />;
        return <YesterdayContent g={g} datum={datum} />;
      }}
    </YesterdayContentWrapper>
  );
}

/* ------------------------------------------------------------------ */
/* Inhalt: Aktivitätszeilen + Summen                                   */
/* ------------------------------------------------------------------ */

function ActivityRow({ akt }: { akt: import('../lib/garmin').GarminAkt }) {
  const start = (akt.start && akt.start.includes(':')) ? akt.start.slice(0, 5) : '—';
  const name = akt.name || '—';
  const kmStr = akt.km != null ? `${fmtNum(akt.km, 1)} km` : '';
  const hmStr = akt.hm != null ? `${fmtNum(akt.hm)} hm` : '';
  const dauerStr = fmtDauer(akt.dauer_s);
  const lastStr = akt.last != null ? `Last ${fmtNum(akt.last)}` : '';
  const teStr = teLabelDe(akt.te_label);

  return (
    <div className="flex flex-col gap-0.5">
      <div className="flex items-baseline gap-2 text-sm flex-wrap">
        <span className="text-2xs tnum text-ink-dim">{start}</span>
        <span className="font-semibold truncate">{name}</span>
        {[kmStr, hmStr, dauerStr].filter(Boolean).join(' · ')}
        {lastStr && <span className="text-2xs tnum text-ink-dim">{lastStr}</span>}
        {teStr && teStr !== '—' && <span className="text-2xs px-1.5 py-0.5 rounded bg-accent-strong/20 text-accent">{teStr}</span>}
      </div>
      <GarminZonenBalken zonen={akt.zonen_s} hoehe={6} />
    </div>
  );
}

function SumLine({ summe }: { summe: import('../lib/garmin').GarminSumme }) {
  const n = summe.n ?? 0;
  const kmStr = summe.km != null ? `${fmtNum(summe.km, 1)} km` : '';
  const hmStr = summe.hm != null ? `${fmtNum(summe.hm)} hm` : '';
  const dauerStr = fmtDauer(summe.dauer_s);
  const bbStr = summe.bb != null ? `Body Battery ${fmtSignedNum(summe.bb)}` : '';

  return (
    <div className="flex flex-col gap-0.5">
      <div className="flex items-baseline gap-2 text-sm flex-wrap font-semibold">
        <span>Summe: {n} {n === 1 ? "Einheit" : "Einheiten"}</span>
        {[kmStr, hmStr, dauerStr].filter(Boolean).join(' · ')}
        {bbStr && <span className="text-2xs tnum text-ink-dim">{bbStr}</span>}
      </div>
      <GarminZonenBalken zonen={summe.zonen_s} hoehe={6} beschriftung />
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Kleine Kennzahl-Zelle (wiederverwendbar)                            */
/* ------------------------------------------------------------------ */

function KPCell({ label, value, hint }: { label: string; value?: ReactNode; hint?: ReactNode }) {
  return (
    <div className="flex flex-col">
      <span className="text-2xs text-ink-dim">{label}</span>
      {value != null && <span className="text-sm tnum font-semibold">{value}</span>}
      {hint != null && <span className="text-2xs text-ink-dim tnum">{hint}</span>}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Stress-Balken (gestapelt, ruhe/niedrig/mittel/hoch)                 */
/* ------------------------------------------------------------------ */

const STRESS_KINDS = ['ruhe', 'niedrig', 'mittel', 'hoch'] as const;
const STRESS_COLORS: Record<string, string> = {
  ruhe: '#22c55e',
  niedrig: '#38bdf8',
  mittel: '#f59e0b',
  hoch: '#ef4444',
};

function StressBar({ stressMin }: { stressMin?: import('../lib/garmin').GarminTagWerte['stress_min'] }) {
  if (!stressMin) return null;
  const entries = STRESS_KINDS.map((k) => ({ key: k, s: stressMin[k] ?? 0 }));
  const total = entries.reduce((a, e) => a + (e.s > 0 ? e.s : 0), 0);
  if (total <= 0) return null;

  return (
    <div>
      <div className="flex w-full overflow-hidden rounded-full bg-bg-subtle" style={{ height: 6 }}>
        {entries.filter((e) => e.s > 0).map((e) => (
          <div
            key={e.key}
            style={{ width: `${(e.s / total) * 100}%`, background: STRESS_COLORS[e.key] }}
            title={`${e.key}: ${fmtNum(Math.round(e.s / 60))} min`}
            className="transition-all"
          />
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Hauptinhalt: Gestern-Kachel                                         */
/* ------------------------------------------------------------------ */

function YesterdayContent({ g, datum }: { g: GarminTag; datum: string }) {
  const tag = g.tag || {};
  const aktListe = g.akt;

  return (
    <Panel title="Gestern · Garmin">
      {/* 1. Aktivitäten + Summen */}
      {aktListe && aktListe.length > 0 ? (
        <div className="flex flex-col gap-3">
          {aktListe.map((a) => (
            <ActivityRow key={a.id ?? a.start} akt={a} />
          ))}
          {g.summe && <SumLine summe={g.summe} />}
        </div>
      ) : (
        <div className="text-sm text-ink-dim">Ruhetag</div>
      )}

      {/* 2. Tageswerte-Raster */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-4">
        {/* Body Battery */}
        <KPCell
          label="Body Battery"
          value={
            tag.bb_min != null && tag.bb_max != null
              ? `${fmtNum(tag.bb_min)}–${fmtNum(tag.bb_max)}`
              : undefined
          }
          hint={
            (tag.bb_geladen != null || tag.bb_verbraucht != null)
              ? `${fmtSignedNum(tag.bb_geladen)} geladen / ${fmtSignedNum(tag.bb_verbraucht)} verbraucht`
              : undefined
          }
        />

        {/* Stress */}
        <KPCell
          label="Stress"
          value={tag.stress_avg != null ? fmtNum(tag.stress_avg) : undefined}
          hint={tag.stress_q != null ? STRESS_Q_DE[tag.stress_q] ?? tag.stress_q : undefined}
        />

        {/* Stress-Minuten Balken */}
        <div className="flex flex-col">
          <span className="text-2xs text-ink-dim">Stressminuten</span>
          {tag.stress_min && <StressBar stressMin={tag.stress_min} />}
        </div>

        {/* Schritte */}
        <KPCell
          label="Schritte"
          value={tag.schritte != null ? fmtNum(tag.schritte) : undefined}
          hint={tag.schritte_ziel != null ? `Ziel ${fmtNum(tag.schritte_ziel)}` : undefined}
        />

        {/* Intensitätsminuten */}
        <KPCell
          label="Intensitätsmin."
          value={
            (tag.int_mod != null || tag.int_vig != null)
              ? `${fmtNum((tag.int_mod ?? 0) + 2 * (tag.int_vig ?? 0))}`
              : undefined
          }
          hint={
            (tag.int_mod != null || tag.int_vig != null)
              ? `mod ${fmtNum(tag.int_mod ?? 0)} · intensiv ${fmtNum(tag.int_vig ?? 0)}`
              : undefined
          }
        />

        {/* Kalorien */}
        <KPCell
          label="Kalorien"
          value={tag.kcal_gesamt != null ? fmtNum(tag.kcal_gesamt) : undefined}
          hint={tag.kcal_aktiv != null ? `aktiv ${fmtNum(tag.kcal_aktiv)}` : undefined}
        />

        {/* Atmung */}
        <KPCell label="Atmung" value={tag.atem_avg != null ? `${fmtNum(tag.atem_avg)} /min` : undefined} />

        {/* SpO2 */}
        <KPCell label="SpO2" value={tag.spo2_avg != null ? `${fmtNum(tag.spo2_avg)} %` : undefined} />
      </div>

      {/* 3. Fußzeile */}
      <div className="text-2xs text-ink-dim mt-4">Quelle: Garmin Personal API</div>
    </Panel>
  );
}
