import { fmtDauer, fmtStunden, HRV_STATUS_DE } from '../lib/garmin';
import { Panel } from './UI';
import { GarminZonenBalken } from './GarminZonenBalken';
import { fmtNum, fmtDateShort } from '../lib/format';
import type { ReactNode } from 'react';

/** Datum als sv-SE String (yyyy-mm-dd). */
function todayStr(): string {
  return new Date().toLocaleDateString('sv-SE');
}

/* ------------------------------------------------------------------ */
/* Hilfsformatierer                                                    */
/* ------------------------------------------------------------------ */
function fmtBedarf(min?: number | null): string {
  if (min == null || !Number.isFinite(min)) return '—';
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${h}:${String(m).padStart(2, '0')} h`;
}

/** "+1:23 h" / "−1:23 h" aus Sekundendifferenz. */
function fmtDiffDauer(diffSek?: number | null): string {
  if (diffSek == null || !Number.isFinite(diffSek)) return '';
  const abs = Math.abs(Math.round(diffSek));
  const h = Math.floor(abs / 3600);
  const m = Math.round((abs % 3600) / 60);
  const sign = diffSek >= 0 ? '+' : '−';
  return `${sign}${h}:${String(m).padStart(2, '0')} h`;
}

/** "HH:MM – HH:MM · N× wach" aus Schlaf.start, Schlaf.ende, Schlaf.wachphasen. */
function fmtSchlafZeit(start?: string | null, wachPhasen?: number | null, ende?: string | null): string {
  if (!start) return '';
  const s = start.includes(':') ? start : start.slice(0, 5);
  const e = (ende && ende.includes(':')) ? ende : '—';
  return `${s} – ${e} · ${wachPhasen ?? 0}× wach`;
}

/** Kleine Kennzahl-Zelle. */
function KPCell({ label, value, hint }: { label: string; value: ReactNode; hint?: ReactNode }) {
  return (
    <div className="flex flex-col">
      <span className="text-2xs text-ink-dim">{label}</span>
      {value != null && <span className="text-sm tnum font-semibold">{value}</span>}
      {hint != null && <span className="text-2xs text-ink-dim tnum">{hint}</span>}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Hauptkomponente – Nacht-Kachel                                     */
/* ------------------------------------------------------------------ */

export function GarminNachtTile() {
  return <NightPanelInner datum={todayStr()} />;
}

function NightPanelInner({ datum }: { datum: string }) {
  return (
    <NightContentWrapper datum={datum}>
      {(tage) => {
        const g = tage.get(datum);
        if (!g || !g.schlaf) {
          return <NoDataText datum={datum} />;
        }
        return <NightContent g={g} datum={datum} />;
      }}
    </NightContentWrapper>
  );
}

/* ------------------------------------------------------------------ */
/* Wrapper mit useGarminTage                                           */
/* ------------------------------------------------------------------ */

import { useGarminTage } from '../lib/garmin';

function NightContentWrapper({
  datum,
  children,
}: {
  datum: string;
  children: (tage: import('../lib/garmin').GarminTage) => ReactNode;
}) {
  const { tage } = useGarminTage();
  if (tage === null) return <div className="text-sm text-ink-dim">Lade Garmin-Daten …</div>;
  return <>{children(tage)}</>;
}

function NoDataText({ datum }: { datum: string }) {
  return <div className="text-sm text-ink-dim">Noch keine Garmin-Daten für {fmtDateShort(datum)}.</div>;
}

/* ------------------------------------------------------------------ */
/* Inhalt einer Nacht-Kachel                                           */
/* ------------------------------------------------------------------ */

function NightContent({ g, datum }: { g: import('../lib/garmin').GarminTag; datum: string }) {
  const sch = g.schlaf || {};
  const hrv = g.hrv || {};
  const tag = g.tag || {};

  const dauerS = sch.dauer_s;
  const tiefS = sch.tief_s ?? 0;
  const leichtS = sch.leicht_s ?? 0;
  const remS = sch.rem_s ?? 0;
  const wachS = sch.wach_s ?? 0;
  const gesamt = tiefS + leichtS + remS + wachS;

  // Bedarf und Differenz
  const bedarfMin = sch.bedarf_min;
  const diffSek = dauerS != null && bedarfMin != null ? dauerS - bedarfMin * 60 : null;  // Fix Cloud-Review

  // Schlafzeit-Text
  const schlafZeit = fmtSchlafZeit(sch.start, sch.wachphasen, sch.ende);

  // Score

  return (
    <Panel title="Letzte Nacht · Garmin">
      {/* 1. Kopfzeile */}
      <div className="flex items-baseline gap-3 mb-3">
        <span className="text-2xl font-semibold tnum">{fmtStunden(dauerS)}</span>
        <div className="flex flex-col">
          <span className="text-sm tnum">Score {sch.score ?? '—'}</span>
          <div className="flex items-baseline gap-2 mt-0.5">
            <span className="text-2xs text-ink-dim">Bedarf {fmtBedarf(bedarfMin)}</span>
            <span className={`text-2xs tnum ${diffSek != null && diffSek < -30 * 60 ? 'text-orange-300' : ''}`}>
              {fmtDiffDauer(diffSek)}
            </span>
          </div>
        </div>
      </div>

      {schlafZeit && <div className="text-2xs text-ink-dim mb-3">{schlafZeit}</div>}

      {/* 2. Schlafphasen-Balken */}
      {gesamt > 0 && (
        <div>
          {/* Balken: Höhe 10, rounded-full */}
          <div className="flex w-full overflow-hidden rounded-full bg-bg-subtle" style={{ height: 10 }}>
            {[
              { s: tiefS, color: '#6366f1', label: 'Tief' },
              { s: leichtS, color: '#38bdf8', label: 'Leicht' },
              { s: remS, color: '#a855f7', label: 'REM' },
              { s: wachS, color: '#f97316', label: 'Wach' },
            ]
              .filter((seg) => seg.s > 0)
              .map((seg) => {
                const pct = Math.round((seg.s / gesamt!) * 100);
                return (
                  <div
                    key={seg.label}
                    style={{ width: `${pct}%`, background: seg.color }}
                    title={`${seg.label}: ${fmtNum(Math.round(seg.s / 60))} min (${pct} %)`}
                    className="transition-all"
                  />
                );
              })}
          </div>
          {/* Legende */}
          <div className="flex flex-wrap gap-x-4 gap-y-1 mt-1 text-2xs text-ink-dim tnum">
            {[
              { label: 'Tief', s: tiefS, color: '#6366f1' },
              { label: 'Leicht', s: leichtS, color: '#38bdf8' },
              { label: 'REM', s: remS, color: '#a855f7' },
              { label: 'Wach', s: wachS, color: '#f97316' },
            ]
              .filter((seg) => seg.s > 0 && gesamt > 0)
              .map((seg) => {
                const pct = Math.round((seg.s / gesamt) * 100);
                return (
                  <span key={seg.label} style={{ color: seg.color }}>
                    {seg.label}: {fmtNum(Math.round(seg.s / 60))} min ({pct} %)
                  </span>
                );
              })}
          </div>
        </div>
      )}

      {/* 3. Kennzahlen-Raster */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-4">
        {/* HRV Nacht */}
        <KPCell
          label="HRV Nacht"
          value={hrv.nacht != null ? fmtNum(hrv.nacht) : undefined}
          hint={hrv.woche != null ? `Woche ${hrv.woche} · Band ${fmtNum(hrv.lo)}–${fmtNum(hrv.hi)}` : undefined}
        />

        {/* Body Battery beim Aufwachen */}
        <KPCell
          label="BB beim Aufwachen"
          value={g.bb_aufwachen != null ? fmtNum(g.bb_aufwachen) : undefined}
          hint={g.bb_aufwachen == null && tag.bb_max != null ? `Max ${fmtNum(tag.bb_max)}` : undefined}
        />

        {/* Schlafstress */}
        <KPCell label="Schlafstress" value={sch.stress != null ? fmtNum(sch.stress) : undefined} />

        {/* HF im Schlaf */}
        <KPCell label="HF im Schlaf" value={sch.hf != null ? `${fmtNum(sch.hf)} bpm` : undefined} />

        {/* Atmung */}
        <KPCell label="Atmung" value={sch.atem != null ? `${fmtNum(sch.atem)} /min` : undefined} />

        {/* SpO2 */}
        <KPCell label="SpO2" value={sch.spo2 != null ? `${fmtNum(sch.spo2)} %` : undefined} />

        {/* Ruhepuls */}
        <KPCell label="Ruhepuls" value={tag.rhr != null ? `${fmtNum(tag.rhr)} bpm` : undefined} />
      </div>

      {/* 4. Fußzeile */}
      <div className="text-2xs text-ink-dim mt-4">Quelle: Garmin Personal API</div>
    </Panel>
  );
}
