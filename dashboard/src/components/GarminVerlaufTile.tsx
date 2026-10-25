import { useMemo, useState } from 'react';
import { Bar, CartesianGrid, ComposedChart, Line, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts';
import { Panel } from './UI';
import { fmtDateShort, fmtNum } from '../lib/format';
import { useGarminTage, type GarminTag } from '../lib/garmin';
import { WINDOWS, winLabel, type WinKey } from '../lib/windows';

type Serie = { key: string; label: string; farbe: string; art: 'linie' | 'balken' | 'flaeche'; stapel?: string; achse?: 'rechts' };
type Ansicht = {
  id: string; titel: string; hinweis: string; woche?: boolean;
  serien: Serie[]; wert: (g: GarminTag) => Record<string, number | null>;
};

const n = (v: unknown): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : null);
const h = (s?: number) => (s != null && s > 0 ? Math.round((s / 3600) * 100) / 100 : null);

const ANSICHTEN: Ansicht[] = [
  {
    id: 'umfang', titel: 'Umfang je Woche', hinweis: 'km und Stunden links, Höhenmeter rechts · Summe Montag–Sonntag', woche: true,
    serien: [
      { key: 'km', label: 'km', farbe: '#38bdf8', art: 'balken' },
      { key: 'std', label: 'Stunden', farbe: '#22c55e', art: 'balken' },
      { key: 'hm', label: 'Höhenmeter', farbe: '#f97316', art: 'linie', achse: 'rechts' },
    ],
    wert: (g) => ({ km: n(g.summe?.km) ?? 0, hm: n(g.summe?.hm) ?? 0, std: h(g.summe?.bewegt_s || g.summe?.dauer_s) ?? 0 }),
  },
  {
    id: 'zonen', titel: 'Zeit in HF-Zonen je Woche', hinweis: 'Stunden je Zone aus den Garmin-Zonensekunden', woche: true,
    serien: [1, 2, 3, 4, 5].map((z, i) => ({ key: `z${z}`, label: `Z${z}`, farbe: ['#94a3b8', '#38bdf8', '#22c55e', '#f97316', '#ef4444'][i], art: 'balken' as const, stapel: 'z' })),
    wert: (g) => Object.fromEntries([0, 1, 2, 3, 4].map((i) => [`z${i + 1}`, h(g.summe?.zonen_s?.[i]) ?? 0])),
  },
  {
    id: 'schlaf', titel: 'Schlafphasen', hinweis: 'Stunden je Phase pro Nacht',
    serien: [
      { key: 'tief', label: 'Tief', farbe: '#6366f1', art: 'balken', stapel: 's' },
      { key: 'rem', label: 'REM', farbe: '#a855f7', art: 'balken', stapel: 's' },
      { key: 'leicht', label: 'Leicht', farbe: '#38bdf8', art: 'balken', stapel: 's' },
      { key: 'wach', label: 'Wach', farbe: '#f97316', art: 'balken', stapel: 's' },
    ],
    wert: (g) => ({ tief: h(g.schlaf?.tief_s), rem: h(g.schlaf?.rem_s), leicht: h(g.schlaf?.leicht_s), wach: h(g.schlaf?.wach_s) }),
  },
  {
    id: 'bb', titel: 'Body Battery', hinweis: 'Tagesspanne (min–max) und Wert beim Aufwachen',
    serien: [
      { key: 'max', label: 'Maximum', farbe: '#22c55e', art: 'linie' },
      { key: 'min', label: 'Minimum', farbe: '#ef4444', art: 'linie' },
      { key: 'auf', label: 'beim Aufwachen', farbe: '#38bdf8', art: 'linie' },
    ],
    wert: (g) => ({ max: n(g.tag?.bb_max), min: n(g.tag?.bb_min), auf: n(g.bb_aufwachen) }),
  },
  {
    id: 'stress', titel: 'Stress', hinweis: 'Tagesdurchschnitt links · Stunden mit hohem/mittlerem Stress rechts',
    serien: [
      { key: 'avg', label: 'Ø Stress', farbe: '#f59e0b', art: 'linie' },
      { key: 'hoch', label: 'hoch (h)', farbe: '#ef4444', art: 'balken', achse: 'rechts', stapel: 'st' },
      { key: 'mittel', label: 'mittel (h)', farbe: '#f97316', art: 'balken', achse: 'rechts', stapel: 'st' },
    ],
    wert: (g) => ({ avg: n(g.tag?.stress_avg), hoch: g.tag?.stress_min?.hoch != null ? g.tag.stress_min.hoch / 60 : null, mittel: g.tag?.stress_min?.mittel != null ? g.tag.stress_min.mittel / 60 : null }),
  },
  {
    id: 'hrv', titel: 'HRV Nacht und Normalband', hinweis: 'ms · gestrichelt = Garmin-Normalband',
    serien: [
      { key: 'hi', label: 'Band oben', farbe: '#64748b', art: 'linie' },
      { key: 'lo', label: 'Band unten', farbe: '#64748b', art: 'linie' },
      { key: 'nacht', label: 'HRV Nacht', farbe: '#a855f7', art: 'linie' },
      { key: 'woche', label: 'HRV Woche', farbe: '#38bdf8', art: 'linie' },
    ],
    wert: (g) => ({ nacht: n(g.hrv?.nacht), woche: n(g.hrv?.woche), lo: n(g.hrv?.lo), hi: n(g.hrv?.hi) }),
  },
  {
    id: 'atem', titel: 'Atmung und SpO2 im Schlaf', hinweis: 'Atemzüge/min links · SpO2 % rechts',
    serien: [
      { key: 'atem', label: 'Atmung', farbe: '#22c55e', art: 'linie' },
      { key: 'spo2', label: 'SpO2', farbe: '#38bdf8', art: 'linie', achse: 'rechts' },
    ],
    wert: (g) => ({ atem: n(g.schlaf?.atem), spo2: n(g.schlaf?.spo2) }),
  },
  {
    id: 'schritte', titel: 'Schritte und Intensitätsminuten', hinweis: 'Schritte links · Intensitätsminuten (mod + 2 × intensiv) rechts',
    serien: [
      { key: 'schritte', label: 'Schritte', farbe: '#38bdf8', art: 'balken' },
      { key: 'im', label: 'Intensitätsmin.', farbe: '#f97316', art: 'linie', achse: 'rechts' },
    ],
    wert: (g) => ({ schritte: n(g.tag?.schritte), im: g.tag?.int_mod != null || g.tag?.int_vig != null ? (g.tag?.int_mod || 0) + 2 * (g.tag?.int_vig || 0) : null }),
  },
  {
    id: 'koerper', titel: 'Gewicht, VO2max, Fitnessalter', hinweis: 'Gewicht kg links · VO2max (aus Aktivitäten) und Fitnessalter rechts',
    serien: [
      { key: 'kg', label: 'Gewicht kg', farbe: '#e5e7eb', art: 'linie' },
      { key: 'vo2', label: 'VO2max', farbe: '#22c55e', art: 'linie', achse: 'rechts' },
      { key: 'fa', label: 'Fitnessalter', farbe: '#a855f7', art: 'linie', achse: 'rechts' },
    ],
    wert: (g) => {
      const vo2 = (g.akt || []).map((a) => a.vo2max).filter((v): v is number => typeof v === 'number');
      return { kg: n(g.gewicht_kg), vo2: vo2.length ? Math.max(...vo2) : null, fa: n(g.fitnessalter) };
    },
  },
];

function montag(d: string): string {
  const dt = new Date(d + 'T12:00:00');
  const wd = (dt.getDay() + 6) % 7;
  dt.setDate(dt.getDate() - wd);
  return dt.toLocaleDateString('sv-SE');
}

export function GarminVerlaufTile() {
  const { tage } = useGarminTage();
  const [ansichtId, setAnsichtId] = useState('umfang');
  const [win, setWin] = useState<WinKey>(120);
  const ansicht = ANSICHTEN.find((a) => a.id === ansichtId)!;

  const rows = useMemo(() => {
    if (!tage) return [];
    const daten = [...tage.keys()].sort();
    const bis = new Date().toLocaleDateString('sv-SE');
    const von = win === 0 ? '0000' : new Date(Date.now() - win * 86400000).toLocaleDateString('sv-SE');
    const sel = daten.filter((d) => d >= von && d <= bis);
    if (!ansicht.woche) return sel.map((d) => ({ date: d, ...ansicht.wert(tage.get(d)!) }));
    const wochen = new Map<string, Record<string, number>>();
    for (const d of sel) {
      const w = montag(d);
      const acc = wochen.get(w) || {};
      const v = ansicht.wert(tage.get(d)!);
      for (const [k, x] of Object.entries(v)) acc[k] = (acc[k] || 0) + (x || 0);
      wochen.set(w, acc);
    }
    return [...wochen.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([w, v]) => ({
      date: w, ...Object.fromEntries(Object.entries(v).map(([k, x]) => [k, Math.round(x * 10) / 10])),
    }));
  }, [tage, ansicht, win]);

  const rechts = ansicht.serien.some((s) => s.achse === 'rechts');
  const leer = !rows.length || ansicht.serien.every((s) => rows.every((r: any) => r[s.key] == null));
  const isHrv = ansicht.id === 'hrv';

  return (
    <Panel
      title="Garmin-Verläufe"
      right={
        <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5">
          {WINDOWS.map((w) => (
            <button key={w} onClick={() => setWin(w)} className={`px-1.5 sm:px-2 py-1 text-xs rounded tnum ${win === w ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}>
              {winLabel(w)}
            </button>
          ))}
        </div>
      }
    >
      <div className="flex flex-wrap gap-1 mb-2" data-testid="garmin-verlauf-ansichten">
        {ANSICHTEN.map((a) => (
          <button key={a.id} onClick={() => setAnsichtId(a.id)} className={`chip text-xs border ${a.id === ansichtId ? 'border-accent text-ink' : 'border-border text-ink-muted hover:text-ink'}`}>
            {a.titel}
          </button>
        ))}
      </div>
      <div className="text-2xs text-ink-dim mb-1">{ansicht.hinweis}</div>
      <div className="h-64" data-testid="garmin-verlauf-chart">
        {tage === null ? (
          <div className="h-full flex items-center justify-center text-sm text-ink-muted">Lade Garmin-Daten …</div>
        ) : leer ? (
          <div className="h-full flex items-center justify-center text-sm text-ink-muted">Keine Garmin-Daten im Zeitraum.</div>
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={rows} margin={{ top: 6, right: rechts ? 4 : 12, left: -12, bottom: 0 }}>
              <CartesianGrid stroke="#1f2731" vertical={false} />
              <XAxis dataKey="date" tickFormatter={fmtDateShort} tickLine={false} axisLine={{ stroke: '#1f2731' }} tick={{ fill: '#8b97a6', fontSize: 11 }} minTickGap={24} />
              <YAxis yAxisId="left" tickLine={false} axisLine={{ stroke: '#1f2731' }} tick={{ fill: '#8b97a6', fontSize: 11 }} domain={['auto', 'auto']} />
              {rechts && <YAxis yAxisId="right" orientation="right" tickLine={false} axisLine={{ stroke: '#1f2731' }} tick={{ fill: '#8b97a6', fontSize: 11 }} domain={['auto', 'auto']} />}
              <Tooltip
                contentStyle={{ background: '#0f141a', border: '1px solid #1f2731', fontSize: 12 }}
                labelFormatter={(v) => (ansicht.woche ? `Woche ab ${fmtDateShort(String(v))}` : fmtDateShort(String(v)))}
                formatter={(v: any, name: string) => [v == null ? '—' : fmtNum(Number(v), Number(v) % 1 ? 1 : 0), ansicht.serien.find((s) => s.key === name)?.label || name]}
              />
              {ansicht.serien.map((s) =>
                s.art === 'balken' ? (
                  <Bar key={s.key} yAxisId={s.achse === 'rechts' ? 'right' : 'left'} dataKey={s.key} fill={s.farbe} fillOpacity={0.85} stackId={s.stapel} maxBarSize={18} isAnimationActive={false} />
                ) : (
                  <Line key={s.key} yAxisId={s.achse === 'rechts' ? 'right' : 'left'} dataKey={s.key} stroke={s.farbe} strokeWidth={isHrv && (s.key === 'hi' || s.key === 'lo') ? 1 : 2} strokeDasharray={isHrv && (s.key === 'hi' || s.key === 'lo') ? '3 3' : undefined} dot={false} connectNulls isAnimationActive={false} />
                ),
              )}
            </ComposedChart>
          </ResponsiveContainer>
        )}
      </div>
      <div className="mt-1 flex flex-wrap gap-3 text-2xs text-ink-muted">
        {ansicht.serien.filter((s) => !(isHrv && s.key === 'lo')).map((s) => (
          <span key={s.key} className="inline-flex items-center gap-1"><span className="inline-block w-2.5 h-2.5 rounded-sm" style={{ background: s.farbe }} />{isHrv && s.key === 'hi' ? 'Normalband' : s.label}</span>
        ))}
      </div>
    </Panel>
  );
}
