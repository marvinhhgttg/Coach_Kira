import { fmtDauer, fmtStunden, HRV_STATUS_DE } from '../lib/garmin';
import { SCHLAF_SCORE_SKALA, farbeVon, nameVon } from '../lib/garminSkala';
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

/* ------------------------------------------------------------------ */
/* 28.11.2026: Vergleich mit deinem Schnitt + letzte 7 Nächte          */
/* ------------------------------------------------------------------ */
const addTage = (d: string, n: number) => { const x = new Date(d + 'T12:00:00'); x.setDate(x.getDate() + n); return x.toLocaleDateString('sv-SE'); };
const minVonZeit = (z?: string | null): number | null => {
  const m = String(z || '').match(/(\d{1,2}):(\d{2})/);
  if (!m) return null;
  let v = Number(m[1]) * 60 + Number(m[2]);
  if (v < 12 * 60) v += 24 * 60; // nach Mitternacht → hinter den Abend legen
  return v;
};
const zeitVonMin = (v: number) => { const x = Math.round(v) % (24 * 60); return `${String(Math.floor(x / 60)).padStart(2, '0')}:${String(x % 60).padStart(2, '0')}`; };
const aufwachMin = (z?: string | null): number | null => { const m = String(z || '').match(/(\d{1,2}):(\d{2})/); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };

type Zeile = { label: string; jetzt: number | null; schnitt: number | null; fmt: (v: number) => string; besser: 'hoch' | 'tief' | null; schwelle: number };
function NachtVergleich({ datum }: { datum: string }) {
  const { tage } = useGarminTage();
  if (!tage) return null;
  const g = tage.get(datum);
  if (!g?.schlaf) return null;
  const vor = Array.from({ length: 14 }, (_, i) => tage.get(addTage(datum, -(i + 1)))).filter((x): x is import('../lib/garmin').GarminTag => !!x?.schlaf);
  if (vor.length < 4) return null;
  const mw = (f: (t: import('../lib/garmin').GarminTag) => number | null | undefined) => {
    const xs = vor.map(f).filter((v): v is number => v != null && Number.isFinite(v) && v > 0);
    return xs.length >= 3 ? xs.reduce((a, b) => a + b, 0) / xs.length : null;
  };
  const s = g.schlaf;
  const hm = (sek: number) => fmtStunden(sek);
  const min = (sek: number) => `${Math.round(sek / 60)} min`;
  const zeilen: Zeile[] = [
    { label: 'Dauer', jetzt: s.dauer_s ?? null, schnitt: mw((t) => t.schlaf?.dauer_s), fmt: hm, besser: 'hoch', schwelle: 20 * 60 },
    { label: 'Score', jetzt: s.score ?? null, schnitt: mw((t) => t.schlaf?.score), fmt: (v: number) => fmtNum(Math.round(v)), besser: 'hoch', schwelle: 3 },
    { label: 'Tiefschlaf', jetzt: s.tief_s ?? null, schnitt: mw((t) => t.schlaf?.tief_s), fmt: min, besser: 'hoch', schwelle: 10 * 60 },
    { label: 'REM', jetzt: s.rem_s ?? null, schnitt: mw((t) => t.schlaf?.rem_s), fmt: min, besser: 'hoch', schwelle: 10 * 60 },
    { label: 'Wach', jetzt: s.wach_s ?? null, schnitt: mw((t) => t.schlaf?.wach_s), fmt: min, besser: 'tief', schwelle: 5 * 60 },
    { label: 'HRV Nacht', jetzt: g.hrv?.nacht ?? null, schnitt: mw((t) => t.hrv?.nacht), fmt: (v: number) => `${fmtNum(Math.round(v))} ms`, besser: 'hoch', schwelle: 2 },
    { label: 'HF im Schlaf', jetzt: s.hf ?? null, schnitt: mw((t) => t.schlaf?.hf), fmt: (v: number) => `${fmtNum(Math.round(v))} bpm`, besser: 'tief', schwelle: 1.5 },
    { label: 'Ruhepuls', jetzt: g.tag?.rhr ?? null, schnitt: mw((t) => t.tag?.rhr), fmt: (v: number) => `${fmtNum(Math.round(v))} bpm`, besser: 'tief', schwelle: 1.5 },
    { label: 'Schlafstress', jetzt: s.stress ?? null, schnitt: mw((t) => t.schlaf?.stress), fmt: (v: number) => fmtNum(Math.round(v)), besser: 'tief', schwelle: 3 },
    { label: 'BB beim Aufwachen', jetzt: g.bb_aufwachen ?? null, schnitt: mw((t) => t.bb_aufwachen), fmt: (v: number) => fmtNum(Math.round(v)), besser: 'hoch', schwelle: 5 },
  ].filter((z) => z.jetzt != null && z.schnitt != null) as Zeile[];
  // Rhythmus: Einschlafen und Aufwachen gegenüber dem Schnitt
  const ein = minVonZeit(s.start), einS = mw((t) => minVonZeit(t.schlaf?.start));
  const auf = aufwachMin(s.ende), aufS = mw((t) => aufwachMin(t.schlaf?.ende));
  const dz = (a: number, b: number) => { const d = Math.round(a - b); return `${d >= 0 ? '+' : '−'}${Math.abs(d)} min`; };
  return (
    <div className="mt-4 border-t border-border pt-3" data-testid="nacht-vergleich">
      <div className="label mb-1.5">Gegenüber deinem Schnitt</div>
      <div className="grid grid-cols-1 gap-x-6 gap-y-1 sm:grid-cols-2 text-xs tnum">
        {zeilen.map((z) => {
          const d = z.jetzt! - z.schnitt!;
          const gut = z.besser == null || Math.abs(d) < z.schwelle ? null : (z.besser === 'hoch') === d > 0;
          const farbe = gut == null ? 'text-ink-dim' : gut ? 'text-green-300' : 'text-orange-300';
          const pfeil = Math.abs(d) < z.schwelle ? '≈' : d > 0 ? '▲' : '▼';
          return (
            <div key={z.label} className="flex items-baseline justify-between gap-2">
              <span className="text-ink-muted">{z.label}</span>
              <span>
                <span className="text-ink font-semibold">{z.fmt(z.jetzt!)}</span>
                <span className="text-ink-dim"> · Ø {z.fmt(z.schnitt!)} </span>
                <span className={farbe}>{pfeil}</span>
              </span>
            </div>
          );
        })}
      </div>
      {(ein != null && einS != null) || (auf != null && aufS != null) ? (
        <p className="mt-1.5 text-2xs text-ink-dim tnum">
          Rhythmus: eingeschlafen {s.start?.slice(0, 5)}{ein != null && einS != null ? ` (Ø ${zeitVonMin(einS)}, ${dz(ein, einS)})` : ''}
          {' · '}aufgewacht {s.ende?.slice(0, 5)}{auf != null && aufS != null ? ` (Ø ${zeitVonMin(aufS)}, ${dz(auf, aufS)})` : ''}
        </p>
      ) : null}
    </div>
  );
}

function SiebenNaechte({ datum }: { datum: string }) {
  const { tage } = useGarminTage();
  if (!tage) return null;
  const xs = Array.from({ length: 7 }, (_, i) => { const d = addTage(datum, i - 6); return { d, s: tage.get(d)?.schlaf }; });
  if (xs.filter((x) => x.s?.dauer_s).length < 3) return null;
  const maxS = Math.max(9.5 * 3600, ...xs.map((x) => (x.s?.dauer_s || 0) + 1800));
  const minS = 4 * 3600; // Achse ab 4 h, damit Unterschiede sichtbar werden
  const hoehe = (sek: number) => Math.max(0, ((sek - minS) / (maxS - minS)) * H);
  const bedarf = xs.map((x) => x.s?.bedarf_min).filter((v): v is number => v != null && v > 0);
  const bedarfS = bedarf.length ? (bedarf.reduce((a, b) => a + b, 0) / bedarf.length) * 60 : null;
  const H = 72;
  const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  return (
    <div className="mt-4 border-t border-border pt-3" data-testid="sieben-naechte">
      <div className="label mb-1.5">Letzte 7 Nächte</div>
      {/* Scores in eigener Zeile, damit nichts mit der Bedarfslinie kollidiert */}
      <div className="flex gap-1.5">
        {xs.map(({ d, s }) => (
          <span key={d} className="flex-1 text-center text-2xs tnum" style={{ color: s?.score != null ? farbeVon(SCHLAF_SCORE_SKALA, s.score) : undefined }}>{s?.score ?? ''}</span>
        ))}
      </div>
      <div className="relative mt-0.5 flex items-end gap-1.5" style={{ height: H }}>
        {bedarfS != null && (
          <div className="pointer-events-none absolute left-0 right-0 z-10 border-t border-dashed border-white/40" style={{ bottom: hoehe(bedarfS) }} title={`Ø Bedarf ${fmtStunden(bedarfS)}`} />
        )}
        {xs.map(({ d, s }) => {
          const ges = s ? (s.tief_s || 0) + (s.leicht_s || 0) + (s.rem_s || 0) + (s.wach_s || 0) : 0;
          const h = s?.dauer_s ? hoehe(s.dauer_s) : 0;
          const teil = (v?: number) => (ges > 0 && v ? (v / ges) * h : 0);
          const heute = d === datum;
          return (
            <div key={d} className="flex flex-1 justify-center" title={s?.dauer_s ? `${fmtDateShort(d)}: ${fmtStunden(s.dauer_s)} · Score ${s.score ?? '—'}` : `${fmtDateShort(d)}: keine Daten`}>
              <div className="flex w-full max-w-[28px] flex-col-reverse overflow-hidden rounded-sm" style={{ height: Math.max(h, 2), opacity: heute ? 1 : 0.7, outline: heute ? '1px solid rgba(255,255,255,0.6)' : undefined }}>
                <div style={{ height: teil(s?.tief_s), background: '#6366f1' }} />
                <div style={{ height: teil(s?.leicht_s), background: '#38bdf8' }} />
                <div style={{ height: teil(s?.rem_s), background: '#a855f7' }} />
                <div style={{ height: teil(s?.wach_s), background: '#f97316' }} />
                {ges === 0 && h > 0 && <div style={{ height: h, background: '#64748b' }} />}
              </div>
            </div>
          );
        })}
      </div>
      <div className="mt-0.5 flex gap-1.5">
        {xs.map(({ d, s }) => (
          <div key={d} className="flex flex-1 flex-col items-center">
            <span className="text-2xs tnum text-ink-muted">{s?.dauer_s ? fmtStunden(s.dauer_s).replace(' h', '') : '–'}</span>
            <span className={`text-2xs ${d === datum ? 'text-ink font-semibold' : 'text-ink-dim'}`}>{WT[new Date(d + 'T12:00:00').getDay()]}</span>
          </div>
        ))}
      </div>
      {bedarfS != null && <p className="mt-1 text-2xs text-ink-dim">Zahl oben: Schlafscore · gestrichelt: dein Ø Schlafbedarf laut Garmin ({fmtStunden(bedarfS)}) · Balken ab 4 h: Tief, Leicht, REM, Wach.</p>}
    </div>
  );
}

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
          <span className="text-sm tnum font-semibold" style={{ color: sch.score != null ? farbeVon(SCHLAF_SCORE_SKALA, sch.score) : undefined }} title={nameVon(SCHLAF_SCORE_SKALA, sch.score)}>Score {sch.score ?? '—'}{sch.score != null ? ` · ${nameVon(SCHLAF_SCORE_SKALA, sch.score)}` : ''}</span>
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

      {/* 4. 28.11.2026: Vergleich mit deinem Schnitt und die letzten 7 Nächte */}
      <NachtVergleich datum={datum} />
      <SiebenNaechte datum={datum} />

      {/* 5. Fußzeile */}
      <div className="text-2xs text-ink-dim mt-4">Quelle: Garmin Personal API · Schnitt = deine letzten 14 Nächte vor heute</div>
    </Panel>
  );
}
