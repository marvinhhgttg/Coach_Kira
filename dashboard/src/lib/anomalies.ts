/**
 * Auffälligkeiten: fest im Code berechnet (keine KI). Vergleich des letzten Tages
 * mit dem persönlichen 28-Tage-Schnitt davor (Mittelwert, Standardabweichung).
 */
import type { RecoveryDay } from '../screens/RecoveryPanel';

export type Anomaly = {
  key: string;
  label: string;
  value: string;
  baseline: string;
  z: number | null;
  level: 'warn' | 'bad';
  hint: string;
};

const WIN = 28;
const MIN_N = 14;

function stats(xs: (number | null | undefined)[]) {
  const v = xs.filter((x): x is number => x != null && Number.isFinite(x));
  if (v.length < MIN_N) return null;
  const m = v.reduce((a, b) => a + b, 0) / v.length;
  const sd = Math.sqrt(v.reduce((a, b) => a + (b - m) ** 2, 0) / (v.length - 1));
  return { m, sd: sd > 0 ? sd : null, n: v.length };
}
const f = (v: number, d = 0) => v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });

/** days: chronologisch, letzter Eintrag = heute (oder letzter Tag mit Morgenwerten). */
export function detectAnomalies(days: RecoveryDay[]): { date: string | null; items: Anomaly[] } {
  const idx = (() => {
    for (let i = days.length - 1; i >= 0; i--) if (days[i].hasMorningData) return i;
    return -1;
  })();
  if (idx < 0) return { date: null, items: [] };
  const d = days[idx];
  const prior = days.slice(Math.max(0, idx - WIN), idx);
  const out: Anomaly[] = [];

  // higherIsBad: true bei RHR; false bei Readiness/HRV/Schlaf
  const metric = (
    key: string,
    label: string,
    val: number | null,
    series: (number | null)[],
    higherIsBad: boolean,
    unit: string,
    dec: number,
    hint: string,
  ) => {
    if (val == null) return;
    const s = stats(series);
    if (!s || !s.sd) return;
    const z = (val - s.m) / s.sd;
    const bad = higherIsBad ? z : -z;
    if (bad >= 2) out.push({ key, label, value: `${f(val, dec)}${unit}`, baseline: `Ø ${f(s.m, dec)}${unit} ± ${f(s.sd, dec)}`, z, level: bad >= 3 ? 'bad' : 'warn', hint });
  };
  metric('rhr', 'Ruhepuls', d.rhr, prior.map((p) => p.rhr), true, ' bpm', 0, 'Erhöhter Ruhepuls kann auf unvollständige Erholung, Stress oder einen beginnenden Infekt hindeuten.');
  metric('hrv', 'HRV', d.hrv, prior.map((p) => p.hrv), false, ' ms', 0, 'Niedrige HRV spricht für eine hohe Belastung des Nervensystems.');
  metric('readiness', 'Readiness', d.readiness, prior.map((p) => p.readiness), false, '', 0, 'Deutlich niedrigere Bereitschaft als sonst.');
  metric('sleep', 'Schlafdauer', d.sleepH, prior.map((p) => p.sleepH), false, ' h', 1, 'Deutlich weniger Schlaf als üblich.');
  metric('sleepScore', 'Schlafscore', d.sleepScore, prior.map((p) => p.sleepScore), false, '', 0, 'Schlafqualität klar unter deinem Schnitt.');

  // Feste Schwellen (Belastung Stand Vortag, wie im Tagesentscheid)
  if (d.acwr != null && d.acwr > 1.3)
    out.push({ key: 'acwr', label: 'ACWR', value: f(d.acwr, 2), baseline: 'Zielbereich 0,8–1,3', z: null, level: d.acwr > 1.5 ? 'bad' : 'warn', hint: 'Akute Belastung deutlich über der chronischen – Verletzungs- und Überlastungsrisiko steigt.' });
  if (d.acwr != null && d.acwr < 0.8)
    out.push({ key: 'acwrLow', label: 'ACWR', value: f(d.acwr, 2), baseline: 'Zielbereich 0,8–1,3', z: null, level: 'warn', hint: 'Belastung unter dem Erhaltungsbereich – Form kann abbauen.' });
  // Monotonie nur als Verlaufssignal melden (Einzelwert allein ist meist unkritisch)
  const ma = d.monoAssess;
  if (d.monotony != null && ma && ma.cap)
    out.push({
      key: 'mono',
      label: 'Monotonie',
      value: f(d.monotony, 2),
      baseline: `${ma.streakNoRelief} Tage ohne Entlastung, davon ${ma.demandingInStreak} fordernd · Zonen: < 1,5 gut, > 2,0 aufmerksam, > 2,5 ungünstig, > 3,0 Warnzone`,
      z: null,
      level: ma.cap === 'EASY' ? 'bad' : 'warn',
      hint: ma.text,
    });
  else if (ma && ma.streakNoRelief >= 3 && ma.demandingInStreak >= 3)
    out.push({ key: 'mono', label: 'Fordernde Tage', value: `${ma.demandingInStreak} in Folge`, baseline: '3–4 ähnliche, fordernde Tage: beobachten', z: null, level: 'warn', hint: 'Achte auf Schlaf, HRV, Ruhepuls, Beine und Motivation – verschlechtern sie sich, Belastung reduzieren.' });
  if (d.befinden != null && d.befinden <= 2)
    out.push({ key: 'befinden', label: 'Befinden', value: `${d.befinden} / 5`, baseline: 'deine Selbsteinschätzung', z: null, level: d.befinden <= 1 ? 'bad' : 'warn', hint: 'Dein Gefühl ist ein wichtiges Signal, auch wenn die Messwerte gut aussehen.' });

  // Mehrtägige Muster
  const last3 = days.slice(Math.max(0, idx - 2), idx + 1);
  if (last3.length === 3 && last3.every((x) => x.rhrDelta != null && x.rhrDelta >= 3))
    out.push({ key: 'rhr3', label: 'Ruhepuls 3 Tage erhöht', value: last3.map((x) => `+${f(x.rhrDelta!, 1)}`).join(' / ') + ' bpm', baseline: 'gegenüber 28-T-Schnitt', z: null, level: 'warn', hint: 'Mehrere Tage in Folge erhöht – Erholung priorisieren.' });
  if (last3.length === 3 && last3.every((x) => x.hrv != null && x.hrvLow != null && x.hrv < x.hrvLow))
    out.push({ key: 'hrv3', label: 'HRV 3 Tage unter Band', value: last3.map((x) => f(x.hrv!)).join(' / ') + ' ms', baseline: 'unter dem Garmin-Normalband', z: null, level: 'bad', hint: 'Anhaltend unter dem Normalband – Belastung reduzieren.' });

  out.sort((a, b) => (a.level === b.level ? 0 : a.level === 'bad' ? -1 : 1));
  return { date: d.date, items: out };
}
