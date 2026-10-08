import { useEffect, useMemo, useRef, useState } from 'react';
import { ApiError, fetchChartData } from '../lib/api';
import { fmtDateShort, fmtNum } from '../lib/format';
import { ErrorBox, Panel, Skeleton } from '../components/UI';
import {
  buildDays,
  demandOf,
  fmtHm,
  LEVEL_CELL,
  LEVEL_TEXT,
  localIsoDate,
  RECOVERY_FETCH_METRICS,
  SLEEP_GOAL_H,
  VOTE_LABEL,
  VOTE_LEVEL,
  VOTE_ORDER,
  type RecoveryDay,
} from './RecoveryPanel';

// ---------------------------------------------------------------------------
// Monatsbericht (Command Center): kompakte Bilanz je Kalendermonat + Vergleich Vormonat.
// Rein regelbasiert aus der Timeline.
// ---------------------------------------------------------------------------

const MONTHS_DE = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];

type MonthStats = {
  key: string;
  label: string;
  days: RecoveryDay[];
  nDays: number;
  partial: boolean;
  ess: number;
  sessions: number;
  restDays: number;
  essPerWeek: number;
  bySport: { sport: string; ess: number; n: number }[];
  readiness: number | null;
  sleepH: number | null;
  nightsGoal: number;
  nightsWithSleep: number;
  hrv: number | null;
  rhr: number | null;
  befinden: number | null;
  votes: { v: (typeof VOTE_ORDER)[number]; n: number }[];
  checked: number;
  ok: number;
  ctlStart: number | null;
  ctlEnd: number | null;
  monoAvg: number | null;
  peakLoad: RecoveryDay | null;
  bestReadiness: RecoveryDay | null;
  bestSleep: RecoveryDay | null;
  lowestRhr: RecoveryDay | null;
};

const avg = (xs: (number | null)[]) => {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};
const argmax = (ds: RecoveryDay[], f: (d: RecoveryDay) => number | null, min = false) => {
  let best: RecoveryDay | null = null;
  let bv = min ? Infinity : -Infinity;
  for (const d of ds) {
    const v = f(d);
    if (v == null || !Number.isFinite(v)) continue;
    if (min ? v < bv : v > bv) {
      bv = v;
      best = d;
    }
  }
  return best;
};
const normSport = (s: string) => {
  const t = (s || '').trim().toLowerCase();
  if (!t || t === 'off' || t === 'krank') return '';
  if (t.startsWith('run') || t === 'laufen') return 'Run';
  if (t.includes('bike') && t.includes('hike')) return 'Bike/Hike';
  if (t.includes('bike') || t.includes('rad')) return 'Bike';
  if (t.includes('hike') || t === 'walk') return 'Hike/Walk';
  if (t.startsWith('ski')) return 'Ski';
  return s.trim();
};

function stats(all: RecoveryDay[], key: string): MonthStats {
  const [y, m] = key.split('-').map(Number);
  const today = localIsoDate();
  const days = all.filter((d) => d.date.startsWith(key) && d.date <= today);
  const lastDay = new Date(y, m, 0).getDate();
  const partial = today.startsWith(key) && Number(today.slice(8, 10)) < lastDay;
  const ess = days.reduce((a, d) => a + (d.essDay ?? 0), 0);
  const sessions = days.filter((d) => (d.essDay ?? 0) > 0).length;
  const sp = new Map<string, { ess: number; n: number }>();
  for (const d of days) {
    const e = d.essDay ?? 0;
    const s = normSport(d.actual.sport);
    if (e <= 0 || !s) continue;
    const o = sp.get(s) || { ess: 0, n: 0 };
    o.ess += e;
    o.n++;
    sp.set(s, o);
  }
  const sleepDays = days.filter((d) => d.sleepH != null && d.sleepH > 0);
  const checked = days.filter((d) => d.vote && d.essDay != null && (d.activityDone || d.essDay === 0));
  const misses = checked.filter(
    (d) => VOTE_ORDER.indexOf(demandOf(d.essDay, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn)) > VOTE_ORDER.indexOf(d.vote!),
  );
  const ctlDays = days.filter((d) => d.ctlPost != null);
  // CTL zu Monatsbeginn = Stand Ende Vormonat (falls vorhanden)
  const prevEnd = all.filter((d) => d.date < `${key}-01` && d.ctlPost != null).slice(-1)[0];
  return {
    key,
    label: `${MONTHS_DE[m - 1]} ${y}`,
    days,
    nDays: days.length,
    partial,
    ess,
    sessions,
    restDays: days.length - sessions,
    essPerWeek: days.length ? (ess / days.length) * 7 : 0,
    bySport: [...sp.entries()].map(([sport, o]) => ({ sport, ...o })).sort((a, b) => b.ess - a.ess),
    readiness: avg(days.map((d) => d.readiness)),
    sleepH: avg(sleepDays.map((d) => d.sleepH)),
    nightsGoal: sleepDays.filter((d) => d.sleepH! >= SLEEP_GOAL_H - 1e-9).length,
    nightsWithSleep: sleepDays.length,
    hrv: avg(days.map((d) => d.hrv)),
    rhr: avg(days.map((d) => d.rhr)),
    befinden: avg(days.map((d) => d.befinden)),
    votes: VOTE_ORDER.slice()
      .reverse()
      .map((v) => ({ v, n: days.filter((d) => d.vote === v).length }))
      .filter((x) => x.n > 0),
    checked: checked.length,
    ok: checked.length - misses.length,
    ctlStart: prevEnd?.ctlPost ?? ctlDays[0]?.ctlPost ?? null,
    ctlEnd: ctlDays[ctlDays.length - 1]?.ctlPost ?? null,
    monoAvg: avg(days.map((d) => d.monotonyPost)),
    peakLoad: argmax(days, (d) => d.essDay),
    bestReadiness: argmax(days, (d) => d.readiness),
    bestSleep: argmax(days, (d) => d.sleepH),
    lowestRhr: argmax(days, (d) => d.rhr, true),
  };
}

function pctDiff(a: number | null, b: number | null) {
  if (a == null || b == null || b === 0) return null;
  return ((a - b) / Math.abs(b)) * 100;
}

function Delta({ cur, prev, digits = 0, unit = '', invert = false, pct = false }: { cur: number | null; prev: number | null; digits?: number; unit?: string; invert?: boolean; pct?: boolean }) {
  if (cur == null || prev == null) return null;
  const d = pct ? pctDiff(cur, prev) : cur - prev;
  if (d == null) return null;
  const good = invert ? d < 0 : d > 0;
  const neutral = Math.abs(d) < (pct ? 3 : Math.pow(10, -digits) * 5);
  return (
    <span className={`ml-1 text-2xs ${neutral ? 'text-ink-dim' : good ? 'text-green-300' : 'text-orange-300'}`}>
      {d > 0 ? '+' : d < 0 ? '−' : '±'}
      {fmtNum(Math.abs(d), pct ? 0 : digits)}
      {pct ? ' %' : unit}
    </span>
  );
}

function conclusions(cur: MonthStats, prev: MonthStats | null): string[] {
  const out: string[] = [];
  const perWeekPrev = prev?.nDays ? prev.essPerWeek : null;
  const dl = pctDiff(cur.essPerWeek, perWeekPrev);
  const dr = cur.readiness != null && prev?.readiness != null ? cur.readiness - prev.readiness : null;
  if (dl != null) {
    if (dl > 10 && dr != null && dr >= -2) out.push(`Mehr Last pro Woche (${dl > 0 ? '+' : ''}${fmtNum(dl)} %) bei stabiler Readiness – Belastung wird gut vertragen.`);
    else if (dl > 10 && dr != null && dr < -2) out.push(`Mehr Last pro Woche (+${fmtNum(dl)} %) und Readiness ${fmtNum(dr)} Punkte – Erholung hinkt hinterher.`);
    else if (dl < -10) out.push(`Weniger Last pro Woche (${fmtNum(dl)} %) als im Vormonat${dr != null && dr > 2 ? ', Readiness dafür höher' : ''}.`);
    else out.push(`Last pro Woche auf Vormonatsniveau (${dl > 0 ? '+' : ''}${fmtNum(dl)} %).`);
  }
  if (cur.sleepH != null) {
    const gap = cur.sleepH - SLEEP_GOAL_H;
    if (gap < -0.25) out.push(`Schlaf im Schnitt ${fmtHm(-gap)} h unter Ziel; nur ${cur.nightsGoal} von ${cur.nightsWithSleep} Nächten ≥ 7:30 h.`);
    else if (gap < 0) out.push(`Schlaf knapp unter Ziel (Ø ${fmtHm(cur.sleepH)} h); ${cur.nightsGoal} von ${cur.nightsWithSleep} Nächten ≥ 7:30 h.`);
    else out.push(`Schlafziel im Schnitt erreicht (Ø ${fmtHm(cur.sleepH)} h; ${cur.nightsGoal} von ${cur.nightsWithSleep} Nächten ≥ 7:30 h).`);
  }
  if (cur.monoAvg != null && cur.monoAvg > 1.8) out.push(`Monotonie im Schnitt ${fmtNum(cur.monoAvg, 2)} – mehr Wechsel zwischen leichten und schweren Tagen einbauen.`);
  if (cur.checked >= 5 && cur.ok / cur.checked < 0.6) out.push(`Nur ${cur.ok} von ${cur.checked} Tagen passten zum Tagesvotum – häufig härter trainiert als empfohlen.`);
  if (cur.ctlStart != null && cur.ctlEnd != null) {
    const d = pctDiff(cur.ctlEnd, cur.ctlStart)!;
    out.push(`Fitness (CTL) ${d >= 0 ? '+' : ''}${fmtNum(d, 1)} % im Monat (${fmtNum(cur.ctlStart)} → ${fmtNum(cur.ctlEnd)}).`);
  }
  return out;
}

export function MonthReport({ days: given }: { days?: RecoveryDay[] } = {}) {
  const [raw, setRaw] = useState<any[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [sel, setSel] = useState<string>(localIsoDate().slice(0, 7));
  const boxRef = useRef<HTMLDivElement>(null);
  const [visible, setVisible] = useState(false);

  // Erst laden, wenn der Bericht in Sichtweite kommt (entlastet den Seitenstart)
  useEffect(() => {
    const el = boxRef.current;
    if (!el || visible) return;
    if (typeof IntersectionObserver === 'undefined') return setVisible(true);
    const io = new IntersectionObserver((es) => es.some((e) => e.isIntersecting) && setVisible(true), { rootMargin: '600px' });
    io.observe(el);
    return () => io.disconnect();
  }, [visible]);

  useEffect(() => {
    if (!visible || given) return;
    fetchChartData('360d', RECOVERY_FETCH_METRICS)
      .then((r) => setRaw(r.data as any[]))
      .catch((e) => setErr(e instanceof ApiError ? e.message : String(e)));
  }, [visible]);

  const all = useMemo(() => given ?? (raw ? buildDays(raw) : []), [raw, given]);
  const months = useMemo(() => {
    const ks = [...new Set(all.map((d) => d.date.slice(0, 7)))].sort();
    return ks.slice(-6);
  }, [all]);
  const cur = useMemo(() => (all.length ? stats(all, sel) : null), [all, sel]);
  const prevKey = useMemo(() => {
    const [y, m] = sel.split('-').map(Number);
    const d = new Date(y, m - 2, 1);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }, [sel]);
  const prev = useMemo(() => (all.some((d) => d.date.startsWith(prevKey)) ? stats(all, prevKey) : null), [all, prevKey]);

  return (
    <div ref={boxRef}>
    <Panel
      title="Monatsbericht"
      right={
        <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5">
          {months.map((k) => {
            const [y, m] = k.split('-').map(Number);
            return (
              <button
                key={k}
                onClick={() => setSel(k)}
                className={`px-2 py-1 text-xs rounded tnum ${sel === k ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
                data-testid={`month-${k}`}
              >
                {MONTHS_DE[m - 1].slice(0, 3)} {String(y).slice(2)}
              </button>
            );
          })}
        </div>
      }
    >
      {err ? (
        <ErrorBox message={err} />
      ) : !cur ? (
        <Skeleton className="h-48" />
      ) : cur.nDays === 0 ? (
        <p className="text-sm text-ink-muted">Keine Daten für diesen Monat.</p>
      ) : (
        <div className="space-y-4">
          <div className="flex flex-wrap items-baseline justify-between gap-2">
            <div className="text-base font-semibold">
              {cur.label}
              {cur.partial && <span className="ml-2 text-2xs text-ink-dim font-normal">laufend · {cur.nDays} Tage</span>}
            </div>
            <span className="text-2xs text-ink-dim">{prev ? `Vergleich: ${prev.label}` : 'kein Vormonat in den Daten'}</span>
          </div>

          <div className="grid gap-3 grid-cols-2 md:grid-cols-4 xl:grid-cols-8">
            <Kpi label="Last gesamt" value={`${fmtNum(cur.ess)} ESS`} delta={cur.partial ? null : <Delta cur={cur.ess} prev={prev?.ess ?? null} pct />} />
            <Kpi label="Ø Last / Woche" value={`${fmtNum(cur.essPerWeek)} ESS`} delta={<Delta cur={cur.essPerWeek} prev={prev?.essPerWeek ?? null} pct />} />
            <Kpi label="Einheiten · Ruhetage" value={`${cur.sessions} · ${cur.restDays}`} />
            <Kpi label="Ø Readiness" value={fmtNum(cur.readiness)} delta={<Delta cur={cur.readiness} prev={prev?.readiness ?? null} />} />
            <Kpi
              label="Ø Schlaf"
              value={`${fmtHm(cur.sleepH)} h`}
              sub={`${cur.nightsGoal}/${cur.nightsWithSleep} Nächte ≥ 7:30`}
              delta={cur.sleepH != null && prev?.sleepH != null ? <Delta cur={cur.sleepH * 60} prev={prev.sleepH * 60} unit=" min" /> : null}
            />
            <Kpi label="Ø HRV" value={`${fmtNum(cur.hrv)} ms`} delta={<Delta cur={cur.hrv} prev={prev?.hrv ?? null} digits={1} unit=" ms" />} />
            <Kpi label="Ø RHR" value={`${fmtNum(cur.rhr, 1)} bpm`} delta={<Delta cur={cur.rhr} prev={prev?.rhr ?? null} digits={1} unit=" bpm" invert />} />
            <Kpi
              label="Fitness (CTL)"
              value={fmtNum(cur.ctlEnd)}
              sub={cur.ctlStart != null ? `Start ${fmtNum(cur.ctlStart)}` : undefined}
              delta={<Delta cur={cur.ctlEnd} prev={cur.ctlStart} pct />}
            />
          </div>

          <div className="grid gap-4 grid-cols-[minmax(0,1fr)] lg:grid-cols-3">
            <div className="panel-raised p-3 min-w-0">
              <div className="label mb-2">Sportarten</div>
              <div className="space-y-1.5">
                {cur.bySport.map((s) => (
                  <div key={s.sport} className="text-xs">
                    <div className="flex justify-between tnum">
                      <span>
                        {s.sport} <span className="text-ink-dim">· {s.n}×</span>
                      </span>
                      <span>
                        {fmtNum(s.ess)} ESS <span className="text-ink-dim">({fmtNum((s.ess / Math.max(cur.ess, 1)) * 100)} %)</span>
                      </span>
                    </div>
                    <div className="h-1.5 rounded bg-bg-subtle mt-0.5 overflow-hidden">
                      <div className="h-full bg-accent/60 wk-bar" style={{ width: `${(s.ess / Math.max(cur.ess, 1)) * 100}%` }} />
                    </div>
                  </div>
                ))}
                {!cur.bySport.length && <span className="text-xs text-ink-dim">Keine Einheiten.</span>}
              </div>
            </div>

            <div className="panel-raised p-3 min-w-0">
              <div className="label mb-2">Voten & Einhaltung</div>
              <div className="flex flex-wrap gap-1">
                {cur.votes.map(({ v, n }) => (
                  <span key={v} className={`chip border text-2xs ${LEVEL_CELL[VOTE_LEVEL[v]]} ${LEVEL_TEXT[VOTE_LEVEL[v]]} border-border`}>
                    {n}× {VOTE_LABEL[v]}
                  </span>
                ))}
              </div>
              <div className="mt-2 text-xs flex justify-between">
                <span className="text-ink-muted">Einheiten passend zum Votum</span>
                <span className={`tnum font-semibold ${cur.checked && cur.ok / cur.checked >= 0.7 ? 'text-green-300' : 'text-orange-300'}`}>
                  {cur.checked ? `${cur.ok} / ${cur.checked}` : '—'}
                </span>
              </div>
              <div className="mt-1 text-xs flex justify-between">
                <span className="text-ink-muted">Ø Monotonie</span>
                <span className="tnum">{fmtNum(cur.monoAvg, 2)}</span>
              </div>
              {cur.befinden != null && (
                <div className="mt-1 text-xs flex justify-between">
                  <span className="text-ink-muted">Ø Befinden</span>
                  <span className="tnum">{fmtNum(cur.befinden, 1)} / 5</span>
                </div>
              )}
            </div>

            <div className="panel-raised p-3 min-w-0">
              <div className="label mb-2">Spitzenwerte</div>
              <dl className="text-xs space-y-1 tnum">
                <Peak label="Höchste Tageslast" d={cur.peakLoad} v={cur.peakLoad ? `${fmtNum(cur.peakLoad.essDay)} ESS ${cur.peakLoad.actual.sport}` : ''} />
                <Peak label="Beste Readiness" d={cur.bestReadiness} v={cur.bestReadiness ? fmtNum(cur.bestReadiness.readiness) : ''} />
                <Peak label="Längster Schlaf" d={cur.bestSleep} v={cur.bestSleep ? `${fmtHm(cur.bestSleep.sleepH)} h` : ''} />
                <Peak label="Niedrigster RHR" d={cur.lowestRhr} v={cur.lowestRhr ? `${fmtNum(cur.lowestRhr.rhr)} bpm` : ''} />
              </dl>
            </div>
          </div>

          <div className="rounded border border-border px-3 py-2">
            <div className="label mb-1">Fazit</div>
            <ul className="text-sm space-y-0.5">
              {conclusions(cur, prev).map((t) => (
                <li key={t}>· {t}</li>
              ))}
            </ul>
            <p className="mt-1 text-2xs text-ink-dim">Regelbasiert aus der Timeline, keine KI-Bewertung.{cur.partial ? ' Laufender Monat: Summen sind noch unvollständig, Wochen-Schnitte vergleichbar.' : ''}</p>
          </div>
        </div>
      )}
    </Panel>
    </div>
  );
}

function Kpi({ label, value, sub, delta }: { label: string; value: string; sub?: string; delta?: React.ReactNode }) {
  return (
    <div className="panel-raised p-2.5 min-w-0">
      <div className="text-2xs text-ink-dim">{label}</div>
      <div className="tnum text-sm font-semibold whitespace-nowrap">
        {value}
        {delta}
      </div>
      {sub && <div className="text-2xs text-ink-dim tnum">{sub}</div>}
    </div>
  );
}

function Peak({ label, d, v }: { label: string; d: RecoveryDay | null; v: string }) {
  return (
    <div className="flex justify-between gap-2">
      <dt className="text-ink-muted">{label}</dt>
      <dd>{d ? `${v} · ${fmtDateShort(d.date).slice(0, 6)}` : '—'}</dd>
    </div>
  );
}
