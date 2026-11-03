import { useEffect, useMemo, useState } from 'react';
import { Bar, BarChart, CartesianGrid, ComposedChart, Legend, Line, ReferenceLine, ResponsiveContainer, Scatter, Tooltip, XAxis, YAxis } from 'recharts';
import { Panel } from './UI';
import { KiraFazit } from './KiraFazit';
import { Hinweis, Kz, Laedt, SportSchalter } from './AnalysenTraining';
import { useGarminTage } from '../lib/garmin';
import { useTimeline } from '../lib/analysenErholung';
import { ALLTAG_KALENDER } from '../lib/extras';
import type { Sport } from '../lib/analysenTraining';
import {
  DAUER_BINS, alltagErholung, entFaktoren, entSaison, evidenz, fetchKalenderHistorie, formBeiSpitzen, hitzeRegel, windRad, type KalTag,
} from '../lib/analysenPlus';
import { de, kurz, vz } from '../lib/stat';

const AX = { fontSize: 10 };
const TT = { contentStyle: { background: 'var(--color-bg-raised, #111827)', border: '1px solid #334155', fontSize: 11 }, labelStyle: { color: '#e5e7eb' } };
const Ev = ({ n, d }: { n: number; d?: number | null }) => {
  const e = evidenz(n, d);
  return <span className={`text-2xs px-1.5 py-0.5 rounded border ${e === 'belastbar' ? 'border-emerald-700 text-emerald-400' : 'border-amber-700 text-amber-400'}`}>{e}</span>;
};
const staerke = (d: number | null) => { const x = Math.abs(d ?? 0); return x >= 0.5 ? 'deutlich' : x >= 0.3 ? 'mittel' : x >= 0.15 ? 'leicht' : 'kaum'; };

// ─── 1 Hitze-Faustregel ─────────────────────────────────────────────────────
export function HitzeRegelTile() {
  const { tage } = useGarminTage();
  const [sport, setSport] = useState<Sport>('lauf');
  const r = useMemo(() => hitzeRegel(tage, sport), [tage, sport]);
  const xEinheit = sport === 'lauf' ? 'km/h' : r.mitNp ? 'W' : 'km/h';
  const facts = [
    `${sport === 'lauf' ? 'Laufen' : 'Rad'} flach, ${r.n} Einheiten mit Garmin-Wetter (davon ${r.nWarm} ab 20 °C). Modell: Ø HF = Intensität (${xEinheit}) + Wärme über 15 °C + Kälte unter 15 °C + Formtrend, R² ${de(r.r2, 2)}.`,
    `Je °C über 15 °C: ${vz(r.warm, 1)} Schläge; je °C unter 15 °C: ${vz(r.kalt == null ? null : -r.kalt, 1)} Schläge je °C kälter.`,
    ...r.regel.map((g) => `${g.temp} °C: ${vz(g.mehrSchlaege, 0)} Schläge gegenüber 15 °C – für gleiche HF ${de(g.wenigerX, 1)} ${xEinheit} weniger.`),
    r.schwuele != null ? `Schwüle (ab 20 °C): je +20 % Luftfeuchte ${vz(r.schwuele, 1)} Schläge.` : 'Für Schwüle zu wenige warme Einheiten.',
  ];
  return (
    <Panel title="Hitze-Faustregel" right={<SportSchalter sport={sport} setSport={setSport} />}>
      <Hinweis>Wie viele Schläge kostet dich Wärme bei gleichem {sport === 'lauf' ? 'Tempo' : 'Druck'}? Lufttemperatur aus dem Garmin-Wetter der Einheit; Intensität und Formtrend sind herausgerechnet. Daraus: wie viel du an warmen Tagen rausnehmen solltest, um in deiner Zone zu bleiben.</Hinweis>
      {tage === null ? <Laedt /> : r.n < 12 || r.warm == null ? <p className="text-2xs text-ink-dim">Zu wenige Einheiten mit Garmin-Wetter ({r.n}).</p> : (
        <>
          <div className="flex items-center gap-2 mb-2"><Ev n={r.nWarm * 4} /><span className="text-2xs text-ink-dim">{r.nWarm} Einheiten ab 20 °C – die Faustregel wird mit dem Sommer genauer.</span></div>
          <div className="grid grid-cols-3 gap-2 mb-2">
            {r.regel.map((g) => (
              <Kz key={g.temp} lab={`bei ${g.temp} °C`} wert={`${vz(g.mehrSchlaege, 0)} Schläge`} ton={(g.mehrSchlaege ?? 0) > 8 ? 'warn' : undefined}
                sub={`gleiche HF: ${de(g.wenigerX, sport === 'lauf' ? 1 : 0)} ${xEinheit} weniger`} />
            ))}
          </div>
          <div className="h-40">
            <ResponsiveContainer>
              <ComposedChart data={r.pts.map((p) => ({ temp: p.temp, hf: p.hf, x: p.x, name: p.name }))} margin={{ left: -10, right: 20 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="temp" type="number" tick={AX} domain={[(m: number) => Math.floor(m / 5) * 5, (m: number) => Math.ceil(m / 5) * 5]} tickFormatter={(t) => `${de(t)} °C`} allowDecimals={false} />
                <YAxis dataKey="hf" tick={AX} width={40} domain={['dataMin - 5', 'dataMax + 5']} />
                <Tooltip {...TT} formatter={(v: any, n: any) => [de(v, n === 'hf' ? 0 : 1), n === 'hf' ? 'Ø HF' : n]} />
                <Scatter dataKey="hf" fill="#f97316" />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <p className="text-2xs text-ink-muted mt-1">Punkte: Ø HF je Einheit gegen Lufttemperatur (ohne Herausrechnen). Je °C über 15 °C {vz(r.warm, 1)} Schläge, unter 15 °C {vz(r.kalt == null ? null : -r.kalt, 1)} je °C kälter.</p>
          <KiraFazit title={`Hitze-Faustregel ${sport}`} facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 2 Wind beim Rad ────────────────────────────────────────────────────────
export function WindRadTile() {
  const { tage } = useGarminTage();
  const r = useMemo(() => windRad(tage), [tage]);
  const facts = [
    `Rad draußen flach, ${r.n} Einheiten mit Normalized Power und Garmin-Wind (Wetterstation, ohne Windrichtung). Wert = Watt je km/h relativ zum Formstand ± 30 Tage.`,
    `Je +1 km/h Wind: ${vz(r.proWind, 2)} % Watt je km/h. Windstärken: ${r.bins.filter((b) => b.n).map((b) => `${b.lab} ${b.n}× ${vz(b.relPct, 1)} %`).join(', ')}.`,
    `Effizienztrend je 90 Tage: roh ${vz(r.trendRohJe90, 1)} %, windbereinigt ${vz(r.trendKorrJe90, 1)} % (positiv = mehr Watt für dasselbe Tempo). Streuung um den Trend: roh ${de(r.streuRoh, 1)} %, bereinigt ${de(r.streuKorr, 1)} %.`,
  ];
  const daten = r.punkte.map((p) => ({ ts: new Date(p.datum + 'T12:00:00').getTime(), roh: p.y, korr: p.yk, wind: p.wind }));
  const klein = Math.abs(r.proWind ?? 0) < 0.3;
  return (
    <Panel title="Wind beim Rad">
      <Hinweis>Bremst Wind deine Rad-Effizienz (Normalized Power je km/h)? Die Rundkurse um Pfaffenhofen gleichen Gegen- und Rückenwind teilweise aus; Garmin liefert nur die Windstärke der Wetterstation, keine Richtung.</Hinweis>
      {tage === null ? <Laedt /> : r.n < 10 ? <p className="text-2xs text-ink-dim">Zu wenige Radeinheiten mit Wind ({r.n}).</p> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="je +10 km/h Wind" wert={`${vz(r.proWind == null ? null : r.proWind * 10, 1)} %`} sub="Watt je km/h" ton={klein ? undefined : 'warn'} />
            <Kz lab="Trend roh je 90 Tage" wert={`${vz(r.trendRohJe90, 1)} %`} />
            <Kz lab="Trend windbereinigt" wert={`${vz(r.trendKorrJe90, 1)} %`} sub={`${r.n} Einheiten · Median-Wind ${de(r.windMed)} km/h`} />
          </div>
          <div className="h-40">
            <ResponsiveContainer>
              <ComposedChart data={daten} margin={{ left: -10, right: 20 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="ts" type="number" domain={['dataMin', 'dataMax']} tickFormatter={(t) => new Date(t).toLocaleDateString('de-DE', { month: 'short', year: '2-digit' })} tick={AX} />
                <YAxis tick={AX} width={36} domain={['auto', 'auto']} tickFormatter={(t) => de(t, 1)} />
                <Tooltip {...TT} labelFormatter={(t) => kurz(new Date(t as number).toLocaleDateString('sv-SE'))} formatter={(v: any, n: any) => [de(v, n === 'wind' ? 0 : 2), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Scatter dataKey="roh" name="W je km/h roh" fill="#64748b" />
                <Scatter dataKey="korr" name="windbereinigt" fill="#38bdf8" />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <p className="text-2xs text-ink-muted mt-1">
            {klein ? 'Ergebnis: Wind verändert deine Rad-Effizienz kaum messbar – der Rad-Effizienztrend ist also nicht vom Wind verzerrt.' : 'Wind wirkt messbar – nutze den bereinigten Trend.'} Die meisten Einheiten lagen unter 10 km/h Wind.
          </p>
          <KiraFazit title="Wind Rad" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 3 Entkopplung über die Saison ──────────────────────────────────────────
export function EntSaisonTile() {
  const { tage } = useGarminTage();
  const [sport, setSport] = useState<Sport>('rad');
  const r = useMemo(() => entSaison(tage, sport), [tage, sport]);
  const facts = [
    `${sport === 'lauf' ? 'Laufen' : 'Rad'}, flach, ab 60 min, ${r.n} Einheiten. Modell: Entkopplung = Dauer + Zeit. Je Stunde länger ${vz(r.proStunde, 1)} Punkte, Trend bei gleicher Dauer ${vz(r.trendJe90, 1)} Punkte je 90 Tage (negativ = Grundlage wächst).`,
    ...r.tabelle.map((t) => `${t.lab}: ${t.n} Einheiten, Median ${de(t.alle, 1)} %${DAUER_BINS.map((b, i) => (t.b[i] != null ? `, ${b.lab} ${de(t.b[i], 1)} %` : '')).join('')}`),
  ];
  return (
    <Panel title="Entkopplung über die Saison" right={<SportSchalter sport={sport} setSport={setSport} />}>
      <Hinweis>Wächst deine Grundlage? Entkopplung langer Einheiten je Quartal und nach Dauer getrennt – längere Einheiten entkoppeln naturgemäß stärker, deshalb ist die Dauer herausgerechnet.</Hinweis>
      {tage === null ? <Laedt /> : r.n < 6 ? <p className="text-2xs text-ink-dim">Zu wenige Einheiten mit Entkopplung ({r.n}).</p> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="Trend bei gleicher Dauer" wert={`${vz(r.trendJe90, 1)} Pkt/90 T`} ton={(r.trendJe90 ?? 0) < 0 ? 'gut' : 'warn'} sub="negativ = besser" />
            <Kz lab="je Stunde länger" wert={`${vz(r.proStunde, 1)} Pkt`} />
            <Kz lab="Einheiten" wert={r.n} sub={<Ev n={r.n} />} />
          </div>
          <div className="h-40">
            <ResponsiveContainer>
              <ComposedChart data={r.linie} margin={{ left: -10, right: 20 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="ts" type="number" domain={['dataMin', 'dataMax']} tickFormatter={(t) => new Date(t).toLocaleDateString('de-DE', { month: 'short', year: '2-digit' })} tick={AX} />
                <YAxis tick={AX} unit=" %" width={44} />
                <ReferenceLine y={5} stroke="#22c55e" strokeDasharray="3 3" />
                <Tooltip {...TT} labelFormatter={(t) => kurz(new Date(t as number).toLocaleDateString('sv-SE'))} formatter={(v: any, n: any) => [`${de(v, 1)} %`, n]} />
                <Scatter dataKey="y" name="Einheit" fill="#64748b" />
                <Line dataKey="m" name="Median der letzten 8" stroke="#22c55e" dot={false} strokeWidth={2} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div className="overflow-x-auto mt-2">
            <table className="w-full text-2xs tnum">
              <thead className="text-ink-dim"><tr><th className="text-left font-normal">Quartal</th><th className="text-right font-normal">n</th><th className="text-right font-normal">alle</th>{DAUER_BINS.map((b) => <th key={b.lab} className="text-right font-normal">{b.lab}</th>)}</tr></thead>
              <tbody>{r.tabelle.map((t) => (
                <tr key={t.q} className="border-t border-border"><td>{t.lab}</td><td className="text-right">{t.n}</td><td className="text-right">{de(t.alle, 1)} %</td>{DAUER_BINS.map((_, i) => <td key={i} className="text-right">{t.b[i] == null ? '–' : `${de(t.b[i], 1)} %`}</td>)}</tr>
              ))}</tbody>
            </table>
          </div>
          <KiraFazit title={`Entkopplung Saison ${sport}`} facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 4 Was die Entkopplung beeinflusst ──────────────────────────────────────
export function EntFaktorenTile() {
  const { tage } = useGarminTage();
  const tl = useTimeline();
  const r = useMemo(() => entFaktoren(tage, tl, 'alle'), [tage, tl]);
  const fmtS = (f: { schwelle: number; einheit: string; id: string }) => (f.id === 'start' ? '14 Uhr' : `${de(f.schwelle, f.einheit === 'h' ? 1 : 0)}${f.einheit ? ' ' + f.einheit : ''}`);
  const facts = [
    `${r.n} lange flache Einheiten (Laufen und Rad). Entkopplung um Dauer und Sportart bereinigt; jeder Faktor am Median (Temperatur 18 °C, Startzeit 14 Uhr) geteilt. Positiv = stärker entkoppelt. Zusammenhang, keine Ursache.`,
    ...r.faktoren.map((f) => `${f.lab} (Grenze ${fmtS(f)}): ${f.niedrig} ${vz(f.m1, 1)} Pkt (${f.n1}×), ${f.hoch} ${vz(f.m2, 1)} Pkt (${f.n2}×), Effekt ${staerke(f.d)}.`),
  ];
  return (
    <Panel title="Was die Entkopplung beeinflusst">
      <Hinweis>Woran liegt es, wenn eine lange Einheit „wegbricht“? Entkopplung nach Schlaf, HRV, Essen am Vortag, Wetter, Startzeit und Vorbelastung – Dauer und Sportart sind herausgerechnet. Werte in Prozentpunkten gegenüber deinem Erwartungswert.</Hinweis>
      {tage === null || tl === null ? <Laedt /> : !r.faktoren.length ? <p className="text-2xs text-ink-dim">Zu wenige Einheiten ({r.n}).</p> : (
        <>
          <div className="flex items-center gap-2 mb-2"><Ev n={Math.round(r.n / 2)} /><span className="text-2xs text-ink-dim">{r.n} Einheiten – Unterschiede unter 1 Punkt sind eher Rauschen.</span></div>
          <div className="space-y-1">
            {r.faktoren.map((f) => {
              const besser = Math.abs((f.m1 ?? 0) - (f.m2 ?? 0)) < 0.3 ? 'kein Unterschied' : (f.m1 ?? 0) < (f.m2 ?? 0) ? f.niedrig : f.hoch;
              return (
                <div key={f.id} className="grid grid-cols-[1fr_auto_auto] gap-2 items-center text-2xs border-t border-border pt-1">
                  <div><div className="text-ink">{f.lab}</div><div className="text-ink-dim">Grenze {fmtS(f)} · {besser === 'kein Unterschied' ? besser : `besser: ${besser}`}</div></div>
                  <div className="tnum text-right text-ink-muted">{f.niedrig}: {vz(f.m1, 1)} ({f.n1})<br />{f.hoch}: {vz(f.m2, 1)} ({f.n2})</div>
                  <div className={`w-16 text-right ${staerke(f.d) === 'deutlich' || staerke(f.d) === 'mittel' ? 'text-amber-400' : 'text-ink-dim'}`}>{staerke(f.d)}</div>
                </div>
              );
            })}
          </div>
          <KiraFazit title="Einflüsse Entkopplung" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 5 Alltag und Erholung ──────────────────────────────────────────────────
export function AlltagErholungTile() {
  const { tage } = useGarminTage();
  const [kal, setKal] = useState<KalTag[] | null>(null);
  const [fehler, setFehler] = useState<string | null>(null);
  useEffect(() => {
    if (!tage || !tage.size) return;
    const ds = [...tage.keys()].sort();
    const bis = new Date(Date.now() - 86400000).toLocaleDateString('sv-SE');
    fetchKalenderHistorie(ds[0], bis, ALLTAG_KALENDER).then(setKal).catch((e) => setFehler(String(e.message || e)));
  }, [tage]);
  const r = useMemo(() => alltagErholung(tage, kal), [tage, kal]);
  const facts = r ? [
    `${r.n} Tage (Kalender: ${ALLTAG_KALENDER.join(', ')}, ohne ganztägige Termine), davon ${r.mitTerminen} mit Terminen. Belegte Stunden (6–22 Uhr) gegen HRV, Schlafscore und Body Battery der folgenden Nacht.`,
    ...r.bins.map((b) => `${b.lab}: ${b.n} Tage, HRV ${vz(b.hrv, 1)} % zur Basis, Schlafscore ${de(b.schlaf)}, Body Battery ${de(b.bb)}, Tagesstress ${de(b.stress)}.`),
    `Bei gleicher Trainingslast und Wochentag: je Termin-Stunde HRV ${vz(r.hrv?.proStunde, 2)} %, Schlafscore ${vz(r.schlaf?.proStunde, 2)}, Body Battery ${vz(r.bb?.proStunde, 2)}.`,
    `Termine bis in den Abend (${r.nachAbend.n1} Tage): HRV ${vz(r.nachAbend.hrv1, 1)} %, Schlafscore ${de(r.nachAbend.schlaf1)}; andere Termintage (${r.nachAbend.n2}): HRV ${vz(r.nachAbend.hrv2, 1)} %, Schlafscore ${de(r.nachAbend.schlaf2)}.`,
  ] : [];
  return (
    <Panel title="Alltag und Erholung">
      <Hinweis>Kosten dich volle Kalendertage Erholung? Belegte Stunden aus deinen drei Kalendern (Marc, Marc (geteilt), Natalia & Marc; ganztägige Termine zählen nicht) gegen HRV und Schlaf der Nacht danach. Trainingslast und Wochenende sind herausgerechnet.</Hinweis>
      {fehler ? (
        <p className="text-2xs text-amber-400">{fehler}. Den Patch findest du im Repo unter docs/patches/Apps-Script-Patch-Kalender-Historie.md.</p>
      ) : !r ? <Laedt /> : r.n < 20 ? <p className="text-2xs text-ink-dim">Zu wenige Tage ({r.n}).</p> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="je Termin-Stunde: HRV" wert={`${vz(r.hrv?.proStunde, 1)} %`} ton={(r.hrv?.proStunde ?? 0) < -0.5 ? 'warn' : undefined} />
            <Kz lab="je Termin-Stunde: Schlaf" wert={`${vz(r.schlaf?.proStunde, 1)}`} sub="Schlafscore" />
            <Kz lab="Abendtermine" wert={`${vz(r.nachAbend.hrv1, 1)} %`} sub={`HRV · andere Tage ${vz(r.nachAbend.hrv2, 1)} %`} />
          </div>
          <div className="h-40">
            <ResponsiveContainer>
              <BarChart data={r.bins} margin={{ left: -10, right: 20 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} />
                <YAxis yAxisId="a" tick={AX} unit=" %" width={40} />
                <YAxis yAxisId="b" orientation="right" tick={AX} width={30} domain={['dataMin - 5', 'dataMax + 3']} />
                <ReferenceLine yAxisId="a" y={0} stroke="#94a3b8" />
                <Tooltip {...TT} formatter={(v: any, n: any, p: any) => [`${de(v, 1)} · ${p.payload.n} Tage`, n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Bar yAxisId="a" dataKey="hrv" name="HRV zur Basis" fill="#38bdf8" />
                <Bar yAxisId="b" dataKey="schlaf" name="Schlafscore" fill="#a78bfa" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="flex items-center gap-2 mt-1"><Ev n={r.mitTerminen} /><span className="text-2xs text-ink-dim">Belegte Stunden des Tages → Nacht danach.</span></div>
          <KiraFazit title="Alltag und Erholung" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 6 Form bei Spitzenleistungen ───────────────────────────────────────────
export function FormSpitzenTile() {
  const { tage } = useGarminTage();
  const tl = useTimeline();
  const [sport, setSport] = useState<Sport>('rad');
  const r = useMemo(() => formBeiSpitzen(tage, tl, sport), [tage, tl, sport]);
  const p = r?.profil, b = r?.basis;
  const mass = sport === 'rad' ? 'beste 20 min (W) harter Radeinheiten' : 'Tempo je Herzschlag flacher Läufe ab 5 km';
  const facts = p && b ? [
    `${sport === 'lauf' ? 'Laufen' : 'Rad'}: Maß = ${mass}, relativ zum eigenen Niveau ± 45 Tage. Spitze = obere 15 % (${p.n} von ${r!.n} Einheiten).`,
    `Vor Spitzen: Last der 7 Tage davor ${vz(p.l7rel, 0)} % zum Wochenschnitt der 4 Wochen davor (übrige Einheiten ${vz(b.l7rel, 0)} %); Tage seit letzter harter Einheit ${de(p.ohneHart)} (übrige ${de(b.ohneHart)}); Last am Vortag ${de(p.gestern)} (übrige ${de(b.gestern)}); Vortag praktisch frei ${de(p.gesternFrei)} % (übrige ${de(b.gesternFrei)} %); Garmin-Verhältnis akut/chronisch ${de(p.acwr, 2)} (übrige ${de(b.acwr, 2)}).`,
    sport === 'lauf' ? 'Achtung: Nach harten Tagen ist die HF oft gedämpft, das lässt Tempo je Herzschlag besser aussehen.' : 'Nur harte Radeinheiten zählen, damit lockere Fahrten und Pendeln nicht mitspielen.',
  ] : [];
  return (
    <Panel title="Form bei Spitzenleistungen" right={<SportSchalter sport={sport} setSport={setSport} />}>
      <Hinweis>Wie warst du vorbelastet, wenn dir richtig gute Einheiten gelungen sind? Daraus lässt sich ableiten, wie viel Entlastung du vor einem Highlight brauchst. Maß: {mass}, relativ zu deinem jeweiligen Niveau.</Hinweis>
      {tage === null || tl === null ? <Laedt /> : !p || !b || !p.n ? <p className="text-2xs text-ink-dim">Zu wenige Einheiten ({r?.n ?? 0}).</p> : (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2">
            <Kz lab="7-Tage-Last davor" wert={`${vz(p.l7rel, 0)} %`} sub={`übrige ${vz(b.l7rel, 0)} %`} ton={(p.l7rel ?? 0) < (b.l7rel ?? 0) - 5 ? 'gut' : undefined} />
            <Kz lab="Last am Vortag" wert={de(p.gestern)} sub={`übrige ${de(b.gestern)}`} />
            <Kz lab="Vortag frei" wert={`${de(p.gesternFrei)} %`} sub={`übrige ${de(b.gesternFrei)} %`} />
            <Kz lab="akut / chronisch" wert={de(p.acwr, 2)} sub={`übrige ${de(b.acwr, 2)}`} />
          </div>
          <div className="flex items-center gap-2 mb-1"><Ev n={p.n} /><span className="text-2xs text-ink-dim">{p.n} Spitzen von {r!.n} Einheiten.</span></div>
          <div className="space-y-0.5 max-h-40 overflow-y-auto">
            {r!.spitzen.map((s) => (
              <div key={s.datum + s.name} className="grid grid-cols-[auto_1fr_auto] gap-2 text-2xs border-t border-border pt-0.5">
                <span className="tnum text-ink-muted">{kurz(s.datum)}</span>
                <span className="truncate">{s.name.split('\n')[0]}</span>
                <span className="tnum text-ink-dim">{vz(s.rel, 0)} % · 7 T {vz(s.l7rel, 0)} % · Vortag {de(s.gestern)}</span>
              </div>
            ))}
          </div>
          <KiraFazit title={`Form Spitzen ${sport}`} facts={facts} />
        </>
      )}
    </Panel>
  );
}
