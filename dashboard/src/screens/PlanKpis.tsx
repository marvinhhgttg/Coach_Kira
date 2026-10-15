import { useEffect, useMemo, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { fmtNum } from '../lib/format';
import { Panel } from '../components/UI';
import { localIsoDate } from './RecoveryPanel';

const KPI_METRICS = [
  'load_fb_day',
  'fbATL_obs',
  'fbCTL_obs',
  'fbACWR_obs',
  'Garmin_Training_Readiness',
  'hrv_status',
  'rhr_bpm',
  'sleep_score_0_100',
  'Trainingszustand',
  'garminEnduranceScore',
] as const;

type Row = Record<string, any>;
type HistoryRow = {
  date: string;
  load: number;
  atl: number | null;
  ctl: number | null;
  acwr: number | null;
  readiness: number | null;
  hrv: number | null;
  rhr: number | null;
  sleep: number | null;
  zustand: string;
  es: number | null;
};
type RiskRow = { cls: string; label: string; n: number; avg: number | null; below40Pct: number | null };
type Bew = 'gut' | 'normal' | 'belastet';

function num(v: unknown): number | null {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function fmtDateShort(d: string): string {
  return d.slice(8, 10) + '.' + d.slice(5, 7) + '.';
}

const clamp = (x: number) => Math.max(0, Math.min(100, x));

export function form(atl: number, ctl: number): number {
  if (ctl === 0) return 0;
  return ((ctl - atl) / ctl) * 100;
}

export function formClass(f: number): string {
  if (f < -15) return 'form_low';
  if (f < -5) return 'form_mid';
  if (f < 5) return 'form_high';
  return 'form_veryhigh';
}

export function acwrClass(a: number): string {
  if (a < 0.9) return 'acwr_low';
  if (a < 1.1) return 'acwr_mid';
  if (a < 1.3) return 'acwr_high';
  return 'acwr_veryhigh';
}

const FORM_LABELS: [string, string][] = [
  ['form_low', '< −15 %'],
  ['form_mid', '−15 bis −5 %'],
  ['form_high', '−5 bis +5 %'],
  ['form_veryhigh', '≥ +5 %'],
];
const ACWR_LABELS: [string, string][] = [
  ['acwr_low', '< 0,9'],
  ['acwr_mid', '0,9 bis 1,1'],
  ['acwr_high', '1,1 bis 1,3'],
  ['acwr_veryhigh', '≥ 1,3'],
];

/** Klasse am Tag i → Readiness am Tag i+1 */
export function riskTable(hist: HistoryRow[], classOf: (h: HistoryRow) => string | null, labels: [string, string][]): RiskRow[] {
  const g = new Map<string, { n: number; sum: number; below40: number }>();
  for (let i = 0; i < hist.length - 1; i++) {
    const r = hist[i + 1].readiness;
    const c = classOf(hist[i]);
    if (r == null || c == null) continue;
    const e = g.get(c) ?? { n: 0, sum: 0, below40: 0 };
    e.n += 1;
    e.sum += r;
    if (r < 40) e.below40 += 1;
    g.set(c, e);
  }
  return labels.map(([cls, label]) => {
    const e = g.get(cls);
    return { cls, label, n: e?.n ?? 0, avg: e && e.n ? e.sum / e.n : null, below40Pct: e && e.n ? (e.below40 / e.n) * 100 : null };
  });
}

export function isRampe(ctlToday: number, ctl7dAgo: number): number {
  if (ctl7dAgo === 0) return 0;
  return (ctlToday / ctl7dAgo - 1) * 100;
}

export function planRampeW1(sim: Array<{ ctl: number }>, len: number): number | null {
  if (len < 2 || !sim[0].ctl) return null;
  return (sim[Math.min(7, len - 1)].ctl / sim[0].ctl - 1) * 100;
}

export function planRampeW2(sim: Array<{ ctl: number }>, len: number): number | null {
  if (len <= 8) return null;
  const a = sim[7].ctl;
  return a ? (sim[len - 1].ctl / a - 1) * 100 : null;
}

const meanOf = (xs: (number | null)[]) => {
  const v = xs.filter((x): x is number => x != null);
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};

export function recoveryIndex(hrvDev: number | null, rhrD: number | null, sleep: number | null, readiness: number | null): number | null {
  const parts: number[] = [];
  if (hrvDev != null) parts.push(clamp(50 + 5 * hrvDev));
  if (rhrD != null) parts.push(clamp(50 - 10 * rhrD));
  if (sleep != null) parts.push(clamp(sleep));
  if (readiness != null) parts.push(clamp(readiness));
  if (parts.length < 2) return null;
  return Math.round(parts.reduce((a, b) => a + b, 0) / parts.length);
}

export function recoveryBewertung(idx: number): Bew {
  if (idx >= 70) return 'gut';
  if (idx >= 50) return 'normal';
  return 'belastet';
}

/** Erholungs-Index für Tag i (Basis = 28 Vortage) */
export function recoveryAt(hist: HistoryRow[], i: number) {
  const h = hist[i];
  const prior = hist.slice(Math.max(0, i - 28), i);
  const hb = meanOf(prior.map((x) => x.hrv));
  const rb = meanOf(prior.map((x) => x.rhr));
  const hrvDev = h.hrv != null && hb ? (h.hrv / hb - 1) * 100 : null;
  const rhrD = h.rhr != null && rb != null ? h.rhr - rb : null;
  const idx = recoveryIndex(hrvDev, rhrD, h.sleep, h.readiness);
  return { date: h.date, hrvDev, rhrD, sleep: h.sleep, readiness: h.readiness, idx, bew: idx != null ? recoveryBewertung(idx) : null };
}

export function fmtSigned(x: number | null, digits = 1): string {
  if (x === null) return '—';
  const s = fmtNum(x, digits);
  return x >= 0 ? '+' + s : s;
}

export function PlanKpis({ sim }: { sim: Array<{ date: string; sim_load: number; atl: number; ctl: number; acwr: number }> }) {
  const [hist, setHist] = useState<HistoryRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    fetchChartData('360d', KPI_METRICS)
      .then((res) => {
        if (cancelled) return;
        const today = localIsoDate();
        const rows: HistoryRow[] = ((res.data ?? []) as Row[])
          .map((r) => ({
            date: String(r.date ?? ''),
            load: num(r.load_fb_day) ?? 0,
            atl: num(r.fbATL_obs),
            ctl: num(r.fbCTL_obs),
            acwr: num(r.fbACWR_obs),
            readiness: num(r.Garmin_Training_Readiness),
            hrv: num(r.hrv_status),
            rhr: num(r.rhr_bpm),
            sleep: num(r.sleep_score_0_100),
            zustand: typeof r.Trainingszustand === 'string' ? r.Trainingszustand.trim() : '',
            es: num(r.garminEnduranceScore),
          }))
          .filter((r) => r.date.length === 10 && r.date <= today && r.ctl != null)
          .sort((a, b) => (a.date < b.date ? -1 : 1));
        setHist(rows);
        setLoading(false);
      })
      .catch((e: unknown) => {
        if (cancelled) return;
        setErr(e instanceof Error ? e.message : 'Fehler beim Laden');
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const simLen = sim.length;
  const last = hist[hist.length - 1];
  const formToday = last && last.atl != null && last.ctl ? form(last.atl, last.ctl) : null;
  const acwrToday = last?.acwr ?? null;
  const ctl7 = hist.length >= 8 ? hist[hist.length - 8].ctl : null;
  const isRampeToday = last?.ctl && ctl7 ? isRampe(last.ctl, ctl7) : null;
  const planW1 = useMemo(() => planRampeW1(sim, simLen), [sim, simLen]);
  const planW2 = useMemo(() => planRampeW2(sim, simLen), [sim, simLen]);

  const formMinSim = useMemo(() => {
    let best: { f: number; date: string } | null = null;
    for (const d of sim.slice(1)) {
      const f = form(d.atl, d.ctl);
      if (!best || f < best.f) best = { f, date: d.date };
    }
    return best;
  }, [sim]);

  const acwrMaxSim = useMemo(() => {
    let best: { a: number; date: string } | null = null;
    for (const d of sim.slice(1)) if (!best || d.acwr > best.a) best = { a: d.acwr, date: d.date };
    return best;
  }, [sim]);

  const formRiskTable = useMemo(
    () => riskTable(hist, (h) => (h.atl != null && h.ctl ? formClass(form(h.atl, h.ctl)) : null), FORM_LABELS),
    [hist],
  );
  const acwrRiskTable = useMemo(() => riskTable(hist, (h) => (h.acwr != null ? acwrClass(h.acwr) : null), ACWR_LABELS), [hist]);

  const rampeRefTable = useMemo(() => {
    const g = new Map<string, number[]>();
    for (let i = 7; i < hist.length; i++) {
      const z = hist[i].zustand;
      const c = hist[i].ctl;
      const c7 = hist[i - 7].ctl;
      if (!z || c == null || !c7) continue;
      g.set(z, [...(g.get(z) ?? []), isRampe(c, c7)]);
    }
    return ['Formaufbau', 'Formerhalt', 'Unproduktiv', 'Ermüdet'].map((z) => {
      const v = g.get(z) ?? [];
      return { z, n: v.length, avg: v.length >= 5 ? v.reduce((a, b) => a + b, 0) / v.length : null };
    });
  }, [hist]);

  const konstanz = useMemo(() => {
    const loads = hist.map((h) => h.load);
    const now = loads.slice(-28);
    const plan = loads.concat(sim.slice(1).map((d) => d.sim_load)).slice(-28);
    return {
      ruhTageHeute: now.filter((l) => l < 1).length,
      langeTageHeute: now.filter((l) => l >= 200).length,
      ruhTagePlan: plan.filter((l) => l < 1).length,
      langeTagePlan: plan.filter((l) => l >= 200).length,
    };
  }, [hist, sim]);

  const ausdauerRefTable = useMemo(() => {
    const g = new Map<string, number[]>();
    for (let i = 28; i < hist.length; i++) {
      const es = hist[i].es;
      const es28 = hist[i - 28].es;
      if (es == null || es28 == null) continue;
      const rest = hist.slice(i - 27, i + 1).filter((h) => h.load < 1).length;
      const cls = rest <= 4 ? '0–4' : rest <= 7 ? '5–7' : rest <= 10 ? '8–10' : '11+';
      g.set(cls, [...(g.get(cls) ?? []), es - es28]);
    }
    return ['0–4', '5–7', '8–10', '11+'].map((cls) => {
      const v = g.get(cls) ?? [];
      return { cls, n: v.length, avg: v.length ? v.reduce((a, b) => a + b, 0) / v.length : null };
    });
  }, [hist]);

  const recoveryHistory = useMemo(() => {
    const out: ReturnType<typeof recoveryAt>[] = [];
    for (let i = Math.max(0, hist.length - 14); i < hist.length; i++) out.push(recoveryAt(hist, i));
    return out;
  }, [hist]);
  // heute: letzter Tag mit Index (Morgenwerte evtl. noch nicht da)
  const recToday = [...recoveryHistory].reverse().find((r) => r.idx != null) ?? null;
  const recoveryIdxToday = recToday?.idx ?? null;
  const recoveryBew = recToday?.bew ?? null;
  const hrvDevPctToday = recToday?.hrvDev ?? null;
  const rhrDevToday = recToday?.rhrD ?? null;
  const sleepToday = recToday?.sleep ?? null;
  const readyToday = recToday?.readiness ?? null;

  const formRiskToday = formToday == null ? null : formRiskTable.find((r) => r.cls === formClass(formToday))?.below40Pct ?? null;
  const acwrRiskToday = acwrToday == null ? null : acwrRiskTable.find((r) => r.cls === acwrClass(acwrToday))?.below40Pct ?? null;

  const formColor = formToday === null ? 'text-ink' : formToday >= 5 ? 'text-green-300' : formToday >= -5 ? 'text-yellow-300' : formToday >= -15 ? 'text-orange-300' : 'text-red-300';
  const acwrColor = acwrToday === null ? 'text-ink' : acwrClass(acwrToday) === 'acwr_mid' ? 'text-green-300' : acwrClass(acwrToday) === 'acwr_high' ? 'text-yellow-300' : 'text-orange-300';
  const rampeColor = isRampeToday === null ? 'text-ink' : isRampeToday > 8 ? 'text-orange-300' : isRampeToday >= 3 ? 'text-green-300' : isRampeToday >= 0 ? 'text-yellow-300' : 'text-orange-300';
  const recoveryColor = recoveryIdxToday === null ? 'text-ink' : recoveryBew === 'gut' ? 'text-green-300' : recoveryBew === 'normal' ? 'text-yellow-300' : 'text-orange-300';

  const loadingNode = (
    <div className="panel-raised p-3 min-w-0">
      <div className="label">Kennzahlen werden geladen…</div>
    </div>
  );
  const errNode = (
    <div className="panel-raised p-3 min-w-0">
      <div className="label text-red-300">Fehler: {err}</div>
    </div>
  );

  return (
    <Panel title="Planungs-Kennzahlen" right={<span className="text-2xs text-ink-dim">aus deiner Historie (360 T)</span>}>
      <div className="grid gap-3 grid-cols-1 md:grid-cols-2 xl:grid-cols-5">
        {loading ? loadingNode : err ? errNode : (
          <>
            {/* Karte 1: Form % · Belastbarkeit morgen */}
            <div className="panel-raised p-3 min-w-0">
              <div className="label">Form % · Belastbarkeit morgen</div>
              <div className={`mt-1 text-xl font-semibold tnum ${formColor}`}>
                {formToday === null ? '—' : fmtSigned(formToday, 1) + ' %'}
              </div>
              <div className="text-2xs text-ink-dim">
                Risiko Readiness &lt; 40 morgen: {formRiskToday === null ? '—' : fmtNum(formRiskToday, 0) + ' %'}
              </div>
              <div className="mt-2 text-2xs text-ink-dim">
                Plan: tiefste Form {formMinSim ? fmtSigned(formMinSim.f, 1) + ' % (' + fmtDateShort(formMinSim.date) + ')' : '—'}
                {' · Risiko ' + (formMinSim ? fmtNum(formRiskTable.find((r) => r.cls === formClass(formMinSim.f))?.below40Pct ?? 0, 0) + ' %' : '—')}
              </div>
              <div className="mt-2 text-2xs text-ink-dim">
                <div className="font-semibold">Klassen (Form heute → Readiness morgen)</div>
                {formRiskTable.map((r) => (
                  <div key={r.cls} className="flex justify-between gap-2">
                    <span>{r.label}</span>
                    <span>Ø {r.avg === null ? '—' : fmtNum(r.avg, 0)} · &lt; 40: {r.below40Pct === null ? '—' : fmtNum(r.below40Pct, 0) + ' %'}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Karte 2: ACWR-Ampel */}
            <div className="panel-raised p-3 min-w-0">
              <div className="label">ACWR-Ampel</div>
              <div className={`mt-1 text-xl font-semibold tnum ${acwrColor}`}>
                {acwrToday === null ? '—' : fmtNum(acwrToday, 2)}
              </div>
              <div className="text-2xs text-ink-dim">
                Risiko Readiness &lt; 40 morgen: {acwrRiskToday === null ? '—' : fmtNum(acwrRiskToday, 0) + ' %'}
              </div>
              <div className="mt-2 text-2xs text-ink-dim">
                Plan: max. ACWR {acwrMaxSim ? fmtNum(acwrMaxSim.a, 2) + ' (' + fmtDateShort(acwrMaxSim.date) + ')' : '—'}
                {' · Risiko ' + (acwrMaxSim ? fmtNum(acwrRiskTable.find((r) => r.cls === acwrClass(acwrMaxSim.a))?.below40Pct ?? 0, 0) + ' %' : '—')}
              </div>
              <div className="mt-2 text-2xs text-ink-dim">
                <div className="font-semibold">Klassen (ACWR heute → Readiness morgen)</div>
                {acwrRiskTable.map((r) => (
                  <div key={r.cls} className="flex justify-between gap-2">
                    <span>{r.label}</span>
                    <span>Ø {r.avg === null ? '—' : fmtNum(r.avg, 0)} · &lt; 40: {r.below40Pct === null ? '—' : fmtNum(r.below40Pct, 0) + ' %'}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Karte 3: CTL-Rampe / Woche */}
            <div className="panel-raised p-3 min-w-0">
              <div className="label">CTL-Rampe / Woche</div>
              <div className={`mt-1 text-xl font-semibold tnum ${rampeColor}`}>
                {isRampeToday === null ? '—' : fmtSigned(isRampeToday, 1) + ' %'}
              </div>
              <div className="text-2xs text-ink-dim">
                Plan W1 {planW1 === null ? '—' : fmtSigned(planW1, 1) + ' %'}
                {' · W2 ' + (planW2 === null ? '—' : fmtSigned(planW2, 1) + ' %')}
              </div>
              <div className="mt-2 text-2xs text-ink-dim">
                Ziel 3–5 %/Woche · &gt; 8 % zu steil
              </div>
              <div className="mt-2 text-2xs text-ink-dim">
                <div className="font-semibold">Garmin-Trainingszustand · Ø Rampe</div>
                {rampeRefTable.map((r) => (
                  <div key={r.z} className="flex justify-between gap-2">
                    <span>{r.z}</span>
                    <span>{r.avg === null ? (r.n === 0 ? '—' : 'n=' + r.n) : fmtSigned(r.avg, 1) + ' % (n=' + r.n + ')'}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Karte 4: Konstanz (28 T) */}
            <div className="panel-raised p-3 min-w-0">
              <div className="label">Konstanz (28 T)</div>
              <div className="mt-1 text-xl font-semibold tnum">
                {konstanz.ruhTageHeute} → {konstanz.ruhTagePlan} <span className="text-xs font-normal text-ink-dim">Ruhetage</span>
              </div>
              <div className="text-2xs text-ink-dim">heute → Planende · lange Tage (≥ 200 ESS) {konstanz.langeTageHeute} → {konstanz.langeTagePlan}</div>
              <div className="mt-2 text-2xs text-ink-dim">
                <div className="font-semibold">Ruhetage / 28 T → Ø Δ Ausdauerwert</div>
                {ausdauerRefTable.map((r) => (
                  <div key={r.cls} className="flex justify-between gap-2">
                    <span>{r.cls}</span>
                    <span>{r.avg === null ? (r.n === 0 ? '—' : 'n=' + r.n) : fmtSigned(r.avg, 0) + ' (n=' + r.n + ')'}</span>
                  </div>
                ))}
              </div>
              <div className="mt-1 text-2xs text-ink-dim">
                mehr Ruhetage gehen bei dir mit sinkendem Ausdauerwert einher
              </div>
            </div>

            {/* Karte 5: Erholungs-Index */}
            <div className="panel-raised p-3 min-w-0">
              <div className="label">Erholungs-Index</div>
              <div className={`mt-1 text-xl font-semibold tnum ${recoveryColor}`}>
                {recoveryIdxToday === null ? '—' : recoveryIdxToday + ' · ' + (recoveryBew === 'gut' ? 'gut erholt' : recoveryBew === 'normal' ? 'normal' : 'belastet')}
              </div>
              <div className="text-2xs text-ink-dim">
                HRV {hrvDevPctToday === null ? '—' : fmtSigned(hrvDevPctToday, 1) + ' % zum Schnitt'}
                {' · Ruhepuls ' + (rhrDevToday === null ? '—' : fmtSigned(rhrDevToday, 0) + ' bpm')}
                {' · Schlaf ' + (sleepToday === null ? '—' : fmtNum(sleepToday, 0))}
                {' · Readiness ' + (readyToday === null ? '—' : fmtNum(readyToday, 0))}
              </div>
              <div className="mt-2 flex items-end gap-0.5 h-10">
                {recoveryHistory.map((h) => {
                  const maxIdx = 100;
                  const hVal = h.idx ?? 0;
                  const hPct = Math.max(0, Math.min(100, hVal));
                  const bg = h.bew === 'gut' ? 'bg-green-400/60' : h.bew === 'normal' ? 'bg-yellow-400/60' : 'bg-orange-400/60';
                  return (
                    <div
                      key={h.date}
                      className={`flex-1 ${bg}`}
                      style={{ height: hPct + '%' }}
                      title={fmtDateShort(h.date) + ': ' + (h.idx === null ? '—' : h.idx)}
                    />
                  );
                })}
              </div>
            </div>
          </>
        )}
      </div>
    </Panel>
  );
}
