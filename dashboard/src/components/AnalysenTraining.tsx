import { useEffect, useMemo, useState, type ReactNode } from 'react';
import { Bar, BarChart, CartesianGrid, ComposedChart, Legend, Line, LineChart, ResponsiveContainer, Scatter, Tooltip, XAxis, YAxis } from 'recharts';
import { Panel } from './UI';
import { KiraFazit } from './KiraFazit';
import { useGarminTage } from '../lib/garmin';
import {
  entkopplung, fetchWetterStunden, hausrunden, laufoekonomie, leistungskurve, tageszeit, temperatur, trainingswirkung, TE_LABELS, type Sport,
} from '../lib/analysenTraining';
import { de, dauerTxt, kurz, monatName, tempoTxt, vz } from '../lib/stat';

const AX = { fontSize: 10 };
const TT = { contentStyle: { background: 'var(--color-bg-raised, #111827)', border: '1px solid #334155', fontSize: 11 }, labelStyle: { color: '#e5e7eb' } };

export function SportSchalter({ sport, setSport }: { sport: Sport; setSport: (s: Sport) => void }) {
  const btn = (a: boolean) => `px-2 py-1 text-xs rounded ${a ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;
  return (
    <div className="inline-flex bg-bg rounded border border-border p-0.5">
      <button className={btn(sport === 'lauf')} onClick={() => setSport('lauf')}>Laufen</button>
      <button className={btn(sport === 'rad')} onClick={() => setSport('rad')}>Rad</button>
    </div>
  );
}
export const Hinweis = ({ children }: { children: ReactNode }) => <p className="text-2xs text-ink-dim mb-2">{children}</p>;
export const Laedt = () => <p className="text-2xs text-ink-dim">lädt …</p>;
export const Kz = ({ lab, wert, sub, ton }: { lab: string; wert: ReactNode; sub?: ReactNode; ton?: 'gut' | 'warn' }) => (
  <div className="rounded border border-border p-2">
    <div className="text-2xs text-ink-dim">{lab}</div>
    <div className={`text-base tnum ${ton === 'gut' ? 'text-emerald-400' : ton === 'warn' ? 'text-amber-400' : 'text-ink'}`}>{wert}</div>
    {sub && <div className="text-2xs text-ink-muted">{sub}</div>}
  </div>
);

// ─── 1 Aerobe Entkopplung ───────────────────────────────────────────────────
export function EntkopplungTile() {
  const { tage } = useGarminTage();
  const [sport, setSport] = useState<Sport>('lauf');
  const r = useMemo(() => entkopplung(tage, sport), [tage, sport]);
  const facts = [
    `${sport === 'lauf' ? 'Laufen' : 'Rad'}, Einheiten ab 60 min, flach: Entkopplung = wie stark Tempo (bzw. Leistung) je Herzschlag in der 2. Hälfte gegenüber der 1. Hälfte sinkt. Unter 5 % gilt als gute Grundlagenausdauer.`,
    `${r.punkte.length} Einheiten mit Rundendaten, Median ${de(r.median, 1)} %, letzte 6 Einheiten ${de(r.aktuell, 1)} %, Trend ${vz(r.trendJe90, 1)} Prozentpunkte je 90 Tage, ${r.unter5} Einheiten unter 5 %.`,
  ];
  return (
    <Panel title="Aerobe Entkopplung" right={<SportSchalter sport={sport} setSport={setSport} />}>
      <Hinweis>Wie stark steigt dein Puls in der zweiten Hälfte langer, flacher Einheiten bei gleichem Tempo bzw. gleicher Leistung? Unter 5 % = stabile Grundlage. Berechnet vom Mac aus den Garmin-Runden.</Hinweis>
      {tage === null ? <Laedt /> : !r.punkte.length ? (
        <p className="text-2xs text-amber-400">
          Noch keine Rundendaten: {r.kandidaten} passende lange Einheiten gefunden. Der Mac lädt die Runden nach, sobald das Mac-Update eingespielt ist
          (Terminal: <code>zsh /tmp/ck-analysen/install-analysen.sh</code>). Die Runden kommen über deinen Garmin-Login von „An Garmin senden“.
        </p>
      ) : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="Median" wert={`${de(r.median, 1)} %`} ton={(r.median ?? 9) < 5 ? 'gut' : 'warn'} sub={`${r.punkte.length} Einheiten`} />
            <Kz lab="letzte 6" wert={`${de(r.aktuell, 1)} %`} ton={(r.aktuell ?? 9) < 5 ? 'gut' : 'warn'} />
            <Kz lab="Trend je 90 Tage" wert={`${vz(r.trendJe90, 1)} Pkt`} ton={(r.trendJe90 ?? 0) < 0 ? 'gut' : 'warn'} sub="negativ = besser" />
          </div>
          <div className="h-44">
            <ResponsiveContainer>
              <ComposedChart data={r.punkte.map((p) => ({ ts: new Date(p.datum + 'T12:00:00').getTime(), y: p.y, name: p.a.name, km: p.a.km }))} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="ts" type="number" domain={['dataMin', 'dataMax']} tickFormatter={(t) => new Date(t).toLocaleDateString('de-DE', { month: 'short', year: '2-digit' })} tick={AX} />
                <YAxis tick={AX} unit=" %" width={48} />
                <Tooltip {...TT} labelFormatter={(t) => kurz(new Date(t as number).toLocaleDateString('sv-SE'))} formatter={(v: any) => [`${de(v, 1)} %`, 'Entkopplung']} />
                <Scatter dataKey="y" fill="#22c55e" />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title={`Entkopplung ${sport}`} facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 2 Hitze und Kälte ──────────────────────────────────────────────────────
export function TemperaturTile() {
  const { tage } = useGarminTage();
  const [sport, setSport] = useState<Sport>('rad');
  const [wetter, setWetter] = useState<Map<string, number> | null>(null);
  const [fehler, setFehler] = useState<string | null>(null);
  useEffect(() => {
    if (!tage || !tage.size) return;
    const ds = [...tage.keys()].sort();
    const bis = new Date(Date.now() - 2 * 86400000).toLocaleDateString('sv-SE');
    fetchWetterStunden(ds[0], bis).then(setWetter).catch((e) => setFehler(String(e.message || e)));
  }, [tage]);
  const r = useMemo(() => temperatur(tage, sport, wetter), [tage, sport, wetter]);
  const facts = [
    `${sport === 'lauf' ? 'Laufen' : 'Rad'} flach, Lufttemperatur aus Garmin-Wetter bzw. Open-Meteo. Werte = Herzschläge je km relativ zum eigenen Formstand (± 30 Tage), negativ = effizienter.`,
    ...r.bins.filter((b) => b.n).map((b) => `${b.lab}: ${b.n} Einheiten, ${vz(b.relPct, 1)} % (Median ${de(b.schlaege)} Schläge/km, HF ${de(b.hf)})`),
    `Steigung: ${vz(r.proGrad == null ? null : r.proGrad * 10, 1)} % je +10 °C.`,
  ];
  return (
    <Panel title="Hitze und Kälte" right={<SportSchalter sport={sport} setSport={setSport} />}>
      <Hinweis>Wie viel Puls kostet dich dein Tempo bei Kälte und Hitze? Lufttemperatur aus dem Garmin-Wetter der Einheit, sonst Open-Meteo für den Raum Pfaffenhofen/Ulm (die Uhr selbst misst durch die Körperwärme zu hoch). Formstand ist herausgerechnet.</Hinweis>
      {fehler && <p className="text-2xs text-amber-400">Wetterdaten nicht ladbar: {fehler}</p>}
      {tage === null || (!wetter && !fehler) ? <Laedt /> : r.n < 5 ? <p className="text-2xs text-ink-dim">Zu wenige Einheiten mit Temperatur ({r.n}).</p> : (
        <>
          <div className="h-44">
            <ResponsiveContainer>
              <BarChart data={r.bins.filter((b) => b.n)} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} />
                <YAxis tick={AX} unit=" %" width={44} />
                <Tooltip {...TT} formatter={(v: any, _n, p: any) => [`${vz(v, 1)} % · ${p.payload.n} Einheiten · ${de(p.payload.schlaege)} Schläge/km`, 'Puls je km']} />
                <Bar dataKey="relPct" fill="#f97316" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <p className="text-2xs text-ink-muted mt-1">
            {r.n} Einheiten · je +10 °C {vz(r.proGrad == null ? null : r.proGrad * 10, 1)} % Puls je km
            {r.optimum ? ` · am effizientesten bei ${r.optimum.lab}` : ''}. Balken unter null = weniger Puls für dasselbe Tempo.
          </p>
          <KiraFazit title={`Temperatur ${sport}`} facts={facts} evidenz={r.n < 20 ? 'schwach' : 'mittel'} regeln={[`Dieser Bericht betrifft nur ${sport === 'lauf' ? 'Laufen' : 'Rad'}.`]} />
        </>
      )}
    </Panel>
  );
}

// ─── 3 Tageszeit ────────────────────────────────────────────────────────────
export function TageszeitTile() {
  const { tage } = useGarminTage();
  const [sport, setSport] = useState<Sport>('lauf');
  const r = useMemo(() => tageszeit(tage, sport), [tage, sport]);
  const facts = [
    ...r.effizienz.filter((z) => z.n).map((z) => `${z.lab}: ${z.n} Einheiten${z.n < 5 ? ' (zu wenige für eine Aussage)' : ''}, Puls je km ${vz(z.relPct, 1)} % zum eigenen Formstand`),
    'Unterschiede unter 1 % beim Puls je km sind Rauschen.',
    ...r.schlaf.map((s) => `Schlaf ${s.lab}: ${s.n} Nächte, Score ${de(s.score)}, HRV ${de(s.hrv)} ms, Tiefschlaf ${de(s.tief)} min`),
  ];
  return (
    <Panel title="Tageszeit" right={<SportSchalter sport={sport} setSport={setSport} />}>
      <Hinweis>Wann bist du am effizientesten (flache Einheiten, Formstand herausgerechnet), und wie schläfst du nach Abendeinheiten?</Hinweis>
      {tage === null ? <Laedt /> : (
        <div className="grid md:grid-cols-2 gap-3">
          <table className="w-full text-2xs tnum">
            <thead><tr className="text-ink-dim text-left"><th>Startzeit</th><th>Einheiten</th><th>Puls je km</th></tr></thead>
            <tbody>
              {r.effizienz.map((z) => (
                <tr key={z.id} className="border-t border-border/50">
                  <td className="py-1">{z.lab}</td><td>{z.n}</td>
                  <td className={z.n < 4 ? 'text-ink-dim' : (z.relPct ?? 0) < -1 ? 'text-emerald-400' : (z.relPct ?? 0) > 1 ? 'text-amber-400' : ''}>{z.n ? `${vz(z.relPct, 1)} %` : '–'}{z.n > 0 && z.n < 4 ? ' (wenig Daten)' : ''}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <table className="w-full text-2xs tnum">
            <thead><tr className="text-ink-dim text-left"><th>Nacht</th><th>n</th><th>Score</th><th>HRV</th><th>Tiefschlaf</th></tr></thead>
            <tbody>
              {r.schlaf.map((s) => (
                <tr key={s.id} className="border-t border-border/50">
                  <td className="py-1">{s.lab}</td><td>{s.n}</td><td>{de(s.score)}</td><td>{de(s.hrv)} ms</td><td>{de(s.tief)} min</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {tage && <KiraFazit title={`Tageszeit ${sport}`} facts={facts} evidenz={r.effizienz.filter((z) => z.n >= 5).length < 2 || Math.max(...r.effizienz.filter((z) => z.n >= 5).map((z) => Math.abs(z.relPct ?? 0))) < 1 ? 'schwach' : 'mittel'} regeln={['Keine pauschale Regel zu Abendtraining ableiten, wenn eine Gruppe weniger als 5 Einheiten hat.']} />}
    </Panel>
  );
}

// ─── 4 Rad-Leistungskurve ───────────────────────────────────────────────────
export function LeistungskurveTile() {
  const { tage } = useGarminTage();
  const r = useMemo(() => leistungskurve(tage), [tage]);
  const facts = [
    `Rad (draußen und Rolle): beste 20 Minuten aller Zeiten ${r.best ? `${r.best.w} W am ${kurz(r.best.datum)}` : '–'}. FTP-Schätzung (95 % der besten 20 min der letzten 90 Tage) aktuell ${de(r.ftpJetzt)} W, zu Beginn ${de(r.ftpStart)} W.`,
    ...r.reihe.slice(-6).map((m) => `${monatName(m.monat)}: beste 20 min ${de(m.best20)} W, NP-Median ${de(m.npMedian)} W, TSS ${m.tss}, IF-Median ${de(m.ifMedian, 2)}`),
  ];
  return (
    <Panel title="Rad-Leistungskurve">
      <Hinweis>Beste 20 Minuten je Monat, Normalized Power und geschätzte FTP (95 % der besten 20 min der letzten 90 Tage). Darunter die Wochen-TSS.</Hinweis>
      {tage === null ? <Laedt /> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="FTP-Schätzung jetzt" wert={`${de(r.ftpJetzt)} W`} sub={`Start ${de(r.ftpStart)} W`} ton={(r.ftpJetzt ?? 0) >= (r.ftpStart ?? 0) ? 'gut' : 'warn'} />
            <Kz lab="Beste 20 Minuten" wert={`${de(r.best?.w)} W`} sub={r.best ? kurz(r.best.datum) : ''} />
            <Kz lab="TSS letzte Woche" wert={de(r.wochen[r.wochen.length - 1]?.tss)} sub={`Schnitt 16 Wo. ${de(r.wochen.reduce((a, w) => a + w.tss, 0) / Math.max(1, r.wochen.length))}`} />
          </div>
          <div className="h-48">
            <ResponsiveContainer>
              <LineChart data={r.reihe.map((m) => ({ ...m, lab: monatName(m.monat) }))} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} />
                <YAxis tick={AX} width={40} domain={['auto', 'auto']} />
                <Tooltip {...TT} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Line dataKey="best20" name="beste 20 min" stroke="#f97316" dot={{ r: 2 }} connectNulls />
                <Line dataKey="npMedian" name="NP-Median" stroke="#38bdf8" dot={{ r: 2 }} connectNulls />
                <Line dataKey="ftp" name="FTP-Schätzung" stroke="#a855f7" strokeDasharray="4 3" dot={false} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
          <div className="h-24 mt-1">
            <ResponsiveContainer>
              <BarChart data={r.wochen.map((w) => ({ ...w, lab: kurz(w.woche).slice(0, 5) }))} margin={{ left: -10, right: 8 }}>
                <XAxis dataKey="lab" tick={AX} interval={1} />
                <YAxis tick={AX} width={40} />
                <Tooltip {...TT} formatter={(v: any) => [v, 'TSS']} />
                <Bar dataKey="tss" fill="#64748b" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title="Rad-Leistungskurve" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 5 Trainingswirkung ─────────────────────────────────────────────────────
const LAB_FARBE: Record<string, string> = Object.fromEntries(Object.values(TE_LABELS).map((x) => [x.lab, x.farbe]));
export function TrainingswirkungTile() {
  const { tage } = useGarminTage();
  const r = useMemo(() => trainingswirkung(tage), [tage]);
  const labs = [...new Set(r.wochen.flatMap((w) => Object.keys(w.labels)))];
  const facts = [
    `Letzte 4 Wochen: ${de(r.qualiJeWoche, 1)} Qualitätseinheiten je Woche (Training Effect aerob ≥ 3,5, anaerob ≥ 2 oder Last ≥ 150), anaerober Reiz ${de(r.anJeWoche, 1)} je Woche – ${r.urteil}.`,
    `Hauptnutzen laut Garmin (4 Wochen): ${Object.entries(r.labelsGesamt).map(([k, v]) => `${k} ${v}`).join(', ')}.`,
  ];
  return (
    <Panel title="Trainingswirkung">
      <Hinweis>Welchen Reiz setzen deine Einheiten laut Garmin Training Effect? Gestapelt je Woche, dazu die Zahl der Qualitätseinheiten.</Hinweis>
      {tage === null ? <Laedt /> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="Qualität je Woche (4 Wo.)" wert={de(r.qualiJeWoche, 1)} sub={r.urteil} ton={r.urteil === 'ausgewogen' ? 'gut' : 'warn'} />
            <Kz lab="anaerober Reiz je Woche" wert={de(r.anJeWoche, 1)} sub="Summe TE anaerob" />
            <Kz lab="Einheiten (4 Wo.)" wert={de(r.wochen.slice(-4).reduce((a, w) => a + w.n, 0))} />
          </div>
          <div className="h-48">
            <ResponsiveContainer>
              <BarChart data={r.wochen.map((w) => ({ lab: kurz(w.woche).slice(0, 5), ...w.labels, Qualität: w.qualitaet }))} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} interval={1} />
                <YAxis tick={AX} width={30} allowDecimals={false} />
                <Tooltip {...TT} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                {labs.map((l) => <Bar key={l} dataKey={l} stackId="a" fill={LAB_FARBE[l] || '#475569'} />)}
              </BarChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title="Trainingswirkung" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 6 Laufökonomie ─────────────────────────────────────────────────────────
export function LaufoekonomieTile() {
  const { tage } = useGarminTage();
  const r = useMemo(() => laufoekonomie(tage), [tage]);
  const facts = [
    `${r.n} flache Läufe. Kadenz steigt um ${de(r.kadenzProTempo == null ? null : -r.kadenzProTempo, 1)} Schritte/min je 1 min/km schneller.`,
    `Kadenz bei gleichem Tempo: Trend ${vz(r.kadenzTrend90, 1)} Schritte/min je 90 Tage. Schrittlänge: ${vz(r.schrittTrend90 == null ? null : r.schrittTrend90 * 100, 1)} cm je 90 Tage.`,
    `Läufe mit höherer Kadenz als fürs Tempo üblich: Median ${de(r.effHoch)} Schläge/km (${r.nHoch}), mit niedrigerer: ${de(r.effTief)} Schläge/km (${r.nTief}). Weniger Schläge/km = effizienter, also war hier ${r.effHoch == null || r.effTief == null ? 'kein Vergleich möglich' : Math.abs(r.effHoch - r.effTief) / r.effTief < 0.01 ? 'kein klarer Unterschied' : r.effHoch < r.effTief ? 'die HÖHERE Kadenz effizienter' : 'die NIEDRIGERE Kadenz effizienter'}.`,
  ];
  return (
    <Panel title="Laufökonomie">
      <Hinweis>Schrittfrequenz und Schrittlänge flacher Läufe je Monat. „Kadenz bereinigt“ = Abweichung von der Kadenz, die du bei diesem Tempo üblicherweise läufst.</Hinweis>
      {tage === null ? <Laedt /> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="Kadenz bei gleichem Tempo" wert={`${vz(r.kadenzTrend90, 1)}/min`} sub="je 90 Tage" ton={(r.kadenzTrend90 ?? 0) >= 0 ? 'gut' : 'warn'} />
            <Kz lab="Schrittlänge" wert={`${vz(r.schrittTrend90 == null ? null : r.schrittTrend90 * 100, 1)} cm`} sub="je 90 Tage" />
            <Kz lab="Höhere Kadenz → Puls je km" wert={`${de(r.effHoch)} vs ${de(r.effTief)}`} sub="niedriger = effizienter" ton={(r.effHoch ?? 0) < (r.effTief ?? 0) ? 'gut' : undefined} />
          </div>
          <div className="h-44">
            <ResponsiveContainer>
              <LineChart data={r.reihe.map((m) => ({ ...m, lab: monatName(m.monat), schrittCm: m.schritt == null ? null : m.schritt * 100 }))} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} />
                <YAxis yAxisId="k" tick={AX} width={36} domain={['auto', 'auto']} />
                <YAxis yAxisId="s" orientation="right" tick={AX} width={36} domain={['auto', 'auto']} />
                <Tooltip {...TT} formatter={(v: any, n: any) => [de(v, n === 'Schrittlänge cm' ? 0 : 0), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Line yAxisId="k" dataKey="kadenz" name="Kadenz" stroke="#22c55e" dot={{ r: 2 }} />
                <Line yAxisId="s" dataKey="schrittCm" name="Schrittlänge cm" stroke="#38bdf8" dot={{ r: 2 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title="Laufökonomie" facts={facts} regeln={['Eine höhere Kadenz nur empfehlen, wenn die Fakten sagen, dass die höhere Kadenz effizienter war.']} />
        </>
      )}
    </Panel>
  );
}

// ─── 15 Hausrunden ──────────────────────────────────────────────────────────
export function HausrundenTile() {
  const { tage } = useGarminTage();
  const runden = useMemo(() => hausrunden(tage), [tage]);
  const [sel, setSel] = useState<string | null>(null);
  const r = runden.find((x) => x.id === sel) || runden[0];
  const istLauf = r?.sport === 'Laufen' || r?.sport === 'Wandern';
  const vTxt = (v: number) => (istLauf ? `${tempoTxt(60 / v)} /km` : `${de(v, 1)} km/h`);
  const facts = runden.slice(0, 8).map((x) => `${x.sport} ${x.ort} ${de(x.km, 1)} km/${de(x.hm)} hm, ${x.n}×: Tempo-Trend ${vz(x.tempoTrendPct, 1)} %, Puls je km ${vz(x.effTrendPct, 1)} % (negativ = effizienter)`);
  return (
    <Panel title="Hausrunden">
      <Hinweis>Strecken, die du immer wieder fährst oder läufst – gleicher Ort, ±6 % Distanz, ähnliche Höhenmeter. Hier ist der Vergleich am fairsten, weil Strecke und Gelände gleich bleiben.</Hinweis>
      {tage === null ? <Laedt /> : !runden.length ? <p className="text-2xs text-ink-dim">Noch keine Strecke mindestens 4-mal.</p> : (
        <>
          <div className="overflow-x-auto">
            <table className="w-full text-2xs tnum">
              <thead><tr className="text-ink-dim text-left"><th>Strecke</th><th>×</th><th>Tempo-Trend</th><th>Puls je km</th><th>beste</th><th>zuletzt</th></tr></thead>
              <tbody>
                {runden.map((x) => {
                  const l = x.sport === 'Laufen' || x.sport === 'Wandern';
                  const f = (v: number) => (l ? `${tempoTxt(60 / v)}` : `${de(v, 1)}`);
                  return (
                    <tr key={x.id} className={`border-t border-border/50 cursor-pointer ${x.id === r?.id ? 'bg-bg-subtle' : ''}`} onClick={() => setSel(x.id)}>
                      <td className="py-1">{x.sport} · {x.ort} · {de(x.km, 1)} km · {de(x.hm)} hm</td><td>{x.n}</td>
                      <td className={(x.tempoTrendPct ?? 0) > 1 ? 'text-emerald-400' : (x.tempoTrendPct ?? 0) < -1 ? 'text-amber-400' : ''}>{vz(x.tempoTrendPct, 1)} %</td>
                      <td className={(x.effTrendPct ?? 0) < -1 ? 'text-emerald-400' : (x.effTrendPct ?? 0) > 1 ? 'text-amber-400' : ''}>{vz(x.effTrendPct, 1)} %</td>
                      <td>{f(x.bestV)}</td><td>{f(x.letzteV)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {r && (
            <div className="h-44 mt-2">
              <ResponsiveContainer>
                <ComposedChart data={r.laeufe.map((l) => ({ ts: new Date(l.datum + 'T12:00:00').getTime(), v: istLauf ? 60 / l.v : l.v, hf: l.a.hf_avg, datum: l.datum, dauer: l.a.bewegt_s }))} margin={{ left: -10, right: 8 }}>
                  <CartesianGrid strokeOpacity={0.12} />
                  <XAxis dataKey="ts" type="number" domain={['dataMin', 'dataMax']} tickFormatter={(t) => new Date(t).toLocaleDateString('de-DE', { month: 'short', year: '2-digit' })} tick={AX} />
                  <YAxis yAxisId="v" tick={AX} width={40} reversed={istLauf} domain={['auto', 'auto']} tickFormatter={(v) => (istLauf ? tempoTxt(v) : de(v, 0))} />
                  <YAxis yAxisId="h" orientation="right" tick={AX} width={32} domain={['auto', 'auto']} />
                  <Tooltip {...TT} labelFormatter={(t) => kurz(new Date(t as number).toLocaleDateString('sv-SE'))} formatter={(v: any, n: any, p: any) => (n === 'Ø HF' ? [v, n] : [vTxt(istLauf ? 60 / v : v) + ` · ${dauerTxt(p.payload.dauer)}`, 'Tempo'])} />
                  <Line yAxisId="v" dataKey="v" name="Tempo" stroke="#22c55e" dot={{ r: 3 }} />
                  <Line yAxisId="h" dataKey="hf" name="Ø HF" stroke="#ef4444" strokeOpacity={0.6} dot={{ r: 2 }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          )}
          <KiraFazit title="Hausrunden" facts={facts} />
        </>
      )}
    </Panel>
  );
}
