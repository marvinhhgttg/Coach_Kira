import { useMemo } from 'react';
import { Bar, BarChart, CartesianGrid, ComposedChart, Legend, Line, ResponsiveContainer, Scatter, Tooltip, XAxis, YAxis, ReferenceLine } from 'recharts';
import { Panel } from './UI';
import { KiraFazit } from './KiraFazit';
import { Hinweis, Kz, Laedt } from './AnalysenTraining';
import { useGarminTage } from '../lib/garmin';
import { eiweiss, energiebilanz, erholungsFaktoren, fruehwarnung, schlafrhythmus, stressBilanz, useTimeline } from '../lib/analysenErholung';
import { de, hhmm, kurz, monatName, vz } from '../lib/stat';

const AX = { fontSize: 10 };
const TT = { contentStyle: { background: 'var(--color-bg-raised, #111827)', border: '1px solid #334155', fontSize: 11 }, labelStyle: { color: '#e5e7eb' } };

// ─── 7 Was dich am besten erholt ────────────────────────────────────────────
export function ErholungsFaktorenTile() {
  const { tage } = useGarminTage();
  const tl = useTimeline();
  const r = useMemo(() => erholungsFaktoren(tage, tl), [tage, tl]);
  const staerke = (d: number | null) => (d == null ? '–' : d >= 0.5 ? 'deutlich' : d >= 0.3 ? 'mittel' : d >= 0.15 ? 'leicht' : 'kaum');
  const facts = [
    `${r.n} Morgen. Ziel: nächtliche HRV relativ zur eigenen 28-Tage-Basis. Jeder Faktor am Vortag am Median geteilt; Stärke = Effektgröße (Cohen d). Zusammenhang, keine bewiesene Ursache.`,
    ...r.faktoren.map((f) => `${f.lab}: nach „${f.gut}“ HRV ${vz(f.hrv1, 1)} % (${f.n1} Nächte), nach „${f.schlecht}“ ${vz(f.hrv2, 1)} % (${f.n2}) – ${staerke(f.d)}`),
  ];
  return (
    <Panel title="Was dich am besten erholt">
      <Hinweis>Welche Abende führen zu einer guten Nacht? Verglichen wird deine nächtliche HRV mit deiner eigenen 28-Tage-Basis – geteilt nach dem, was am Vortag war. Das zeigt Zusammenhänge, keine bewiesenen Ursachen.</Hinweis>
      {tage === null || tl === null ? <Laedt /> : (
        <>
          <table className="w-full text-2xs tnum">
            <thead><tr className="text-ink-dim text-left"><th>Faktor</th><th>besser nach</th><th>HRV</th><th>schlechter nach</th><th>HRV</th><th>Stärke</th></tr></thead>
            <tbody>
              {r.faktoren.map((f) => (
                <tr key={f.id} className="border-t border-border/50">
                  <td className="py-1">{f.lab}</td>
                  <td>{f.gut} <span className="text-ink-dim">({f.n1})</span></td><td className="text-emerald-400">{vz(f.hrv1, 1)} %</td>
                  <td>{f.schlecht} <span className="text-ink-dim">({f.n2})</span></td><td className="text-amber-400">{vz(f.hrv2, 1)} %</td>
                  <td className={(f.d ?? 0) >= 0.3 ? 'text-ink' : 'text-ink-dim'}>{staerke(f.d)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <KiraFazit title="Erholungsfaktoren" facts={facts} regeln={['Zusammenhänge sind beobachtet, nicht verursacht. Ernährungsfaktoren: nur erfasste Tage.']} />
        </>
      )}
    </Panel>
  );
}

// ─── 8 Schlafrhythmus ───────────────────────────────────────────────────────
export function SchlafrhythmusTile() {
  const { tage } = useGarminTage();
  const tl = useTimeline();
  const r: any = useMemo(() => schlafrhythmus(tage, tl), [tage, tl]);
  const facts = r && r.bettMedian != null ? [
    `Letzte 28 Nächte: ins Bett im Median ${hhmm(r.bettMedian)} (Streuung ± ${de(r.bettSd)} min), aufstehen ${hhmm(r.aufMedian)} (± ${de(r.aufSd)} min). Wochenende verschiebt die Schlafmitte um ${vz(r.jetlag)} min.`,
    `Regelmäßige Nächte (Schlafmitte ±30 min zum 14-Tage-Schnitt): ${r.regel.n}, Score ${de(r.regel.score)}, Bereitschaft ${de(r.regel.ready)}. Unregelmäßige (> 60 min): ${r.unregel.n}, Score ${de(r.unregel.score)}, Bereitschaft ${de(r.unregel.ready)}.`,
  ] : [];
  const daten = r?.naechte?.slice(-90).map((n: any) => ({ ts: new Date(n.d + 'T12:00:00').getTime(), bett: n.bett, auf: n.auf, score: n.score })) || [];
  return (
    <Panel title="Schlafrhythmus">
      <Hinweis>Wie regelmäßig gehst du ins Bett und stehst auf – und was macht das mit Schlafscore und Bereitschaft? Letzte 90 Nächte.</Hinweis>
      {!r ? <Laedt /> : r.bettMedian == null ? <p className="text-2xs text-ink-dim">Zu wenige Nächte.</p> : (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2">
            <Kz lab="Ins Bett (28 N.)" wert={hhmm(r.bettMedian)} sub={`± ${de(r.bettSd)} min`} ton={(r.bettSd ?? 99) <= 45 ? 'gut' : 'warn'} />
            <Kz lab="Aufstehen" wert={hhmm(r.aufMedian)} sub={`± ${de(r.aufSd)} min`} ton={(r.aufSd ?? 99) <= 45 ? 'gut' : 'warn'} />
            <Kz lab="Wochenend-Verschiebung" wert={`${vz(r.jetlag)} min`} sub="Schlafmitte" ton={Math.abs(r.jetlag ?? 0) <= 45 ? 'gut' : 'warn'} />
            <Kz lab="Regelmäßig vs. unregelmäßig" wert={`${de(r.regel.score)} / ${de(r.unregel.score)}`} sub={`Score · Bereitschaft ${de(r.regel.ready)} / ${de(r.unregel.ready)}`} />
          </div>
          <div className="h-44">
            <ResponsiveContainer>
              <ComposedChart data={daten} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="ts" type="number" domain={['dataMin', 'dataMax']} tickFormatter={(t) => new Date(t).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit' })} tick={AX} />
                <YAxis tick={AX} width={44} domain={[-180, 600]} ticks={[-120, 0, 120, 240, 360, 480]} tickFormatter={(v) => hhmm(v)} />
                <Tooltip {...TT} labelFormatter={(t) => kurz(new Date(t as number).toLocaleDateString('sv-SE'))} formatter={(v: any, n: any) => [n === 'Score' ? v : hhmm(v), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Scatter dataKey="bett" name="ins Bett" fill="#a855f7" />
                <Scatter dataKey="auf" name="aufgestanden" fill="#f59e0b" />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title="Schlafrhythmus" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 9 Atem/SpO2-Frühwarnung ────────────────────────────────────────────────
const FARBE = { gruen: 'bg-emerald-500/70', gelb: 'bg-amber-400/80', rot: 'bg-red-500/80' };
export function FruehwarnungAtemTile() {
  const { tage } = useGarminTage();
  const r = useMemo(() => fruehwarnung(tage), [tage]);
  const letzte = r.tage.slice(-21);
  const heute = r.tage[r.tage.length - 1];
  const facts = [
    `Signale je Morgen gegen die eigene 28-Tage-Basis: Atmung nachts ≥ +1/min, SpO2 ≤ −2 %, Ruhepuls ≥ +4, HRV ≤ −15 %. 2 Signale = gelb, 3+ = rot (rot nur, wenn der Nachbartag auch auffällig ist).`,
    heute ? `Heute (${kurz(heute.d)}): ${heute.stufe === 'gruen' && heute.signale.length === 1 ? 'grün, aber mit einem Einzelsignal' : heute.stufe}${heute.signale.length ? ' – ' + heute.signale.join(', ') : ''}. Ein einzelnes Signal ist laut Regel noch keine Warnung, soll aber genannt und beobachtet werden (SpO2 der Uhr schwankt nachts, z. B. durch Liegeposition).` : '',
    `Letzte Auffälligkeiten: ${r.episoden.slice(0, 5).map((e) => `${kurz(e.von)}${e.bis !== e.von ? '–' + kurz(e.bis) : ''} ${e.stufe} (${e.signale.join(', ')})`).join('; ') || 'keine'}.`,
  ].filter(Boolean);
  return (
    <Panel title="Atmung und SpO2 – Frühwarnung">
      <Hinweis>Erhöhte Atemfrequenz nachts und niedrigere Sauerstoffsättigung kündigen oft einen Infekt an, zusammen mit Ruhepuls und HRV. Ergänzt die bestehende Frühwarnung um diese beiden Signale.</Hinweis>
      {tage === null ? <Laedt /> : !heute ? <p className="text-2xs text-ink-dim">Zu wenige Nächte.</p> : (
        <>
          <div className="flex items-center gap-3 mb-2">
            <span className={`inline-block w-3 h-3 rounded-full ${heute.stufe === 'gruen' && heute.signale.length === 1 ? 'bg-emerald-500/70 ring-2 ring-amber-400' : FARBE[heute.stufe]}`} />
            <span className="text-sm text-ink">Heute {heute.stufe === 'gruen' ? (heute.signale.length === 1 ? 'ein Einzelsignal – allein noch keine Warnung (gelb ab zwei Signalen)' : 'unauffällig') : heute.stufe === 'gelb' ? 'zwei Signale' : 'mehrere Signale'}</span>
            <span className="text-2xs text-ink-muted tnum">Atmung {de(heute.atem)} · SpO2 {de(heute.spo2)} % · Ruhepuls {de(heute.rhr)} · HRV {de(heute.hrv)}</span>
          </div>
          {heute.signale.length > 0 && <p className="text-2xs text-amber-400 mb-2">{heute.signale.join(' · ')}</p>}
          <div className="flex gap-0.5 mb-1" title="letzte 21 Tage">
            {letzte.map((t) => <div key={t.d} className={`h-5 flex-1 rounded-sm ${FARBE[t.stufe]}`} title={`${kurz(t.d)}: ${t.signale.join(', ') || 'unauffällig'}`} />)}
          </div>
          <div className="text-2xs text-ink-dim mb-2">letzte 21 Tage · Tippen/Hovern zeigt die Signale</div>
          <div className="text-2xs text-ink-muted">
            {r.episoden.slice(0, 6).map((e) => (
              <div key={e.von}>{kurz(e.von)}{e.bis !== e.von ? `–${kurz(e.bis)}` : ''}: <span className={e.stufe === 'rot' ? 'text-red-400' : 'text-amber-400'}>{e.stufe}</span> ({e.signale.join(', ')})</div>
            ))}
          </div>
          <KiraFazit title="Atem-Frühwarnung" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 10 Stress-Bilanz ───────────────────────────────────────────────────────
export function StressBilanzTile() {
  const { tage } = useGarminTage();
  const r = useMemo(() => stressBilanz(tage), [tage]);
  const facts = r ? [
    `Garmin-Stress im Tagesschnitt nach Wochentag: ${r.proWt.map((w) => `${w.lab} ${de(w.avg)}`).join(', ')}.`,
    `Werktag ohne Training: Ø ${de(r.werktagRuhe.avg)} (${de(r.werktagRuhe.hoch)} min mittel/hoch, ${r.werktagRuhe.n} Tage); Werktag mit Training: Ø ${de(r.werktagTraining.avg)} (${de(r.werktagTraining.hoch)} min); Wochenende: Ø ${de(r.wochenende.avg)} (${de(r.wochenende.hoch)} min); Tage mit harter Einheit: Ø ${de(r.hartTage.avg)}.`,
    `Monatsverlauf: ${r.reihe.slice(-6).map((m) => `${monatName(m.monat)} ${de(m.avg)}`).join(', ')}.`,
  ] : [];
  return (
    <Panel title="Stress-Bilanz">
      <Hinweis>Garmin-Stress (aus der HRV über den Tag) nach Wochentagen und Tagesart. Kosten dich Bürotage mehr als harte Einheiten?</Hinweis>
      {!r ? <Laedt /> : (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2">
            <Kz lab="Werktag ohne Training" wert={de(r.werktagRuhe.avg)} sub={`${de(r.werktagRuhe.hoch)} min mittel/hoch`} />
            <Kz lab="Werktag mit Training" wert={de(r.werktagTraining.avg)} sub={`${de(r.werktagTraining.hoch)} min`} ton={(r.werktagTraining.avg ?? 0) < (r.werktagRuhe.avg ?? 0) ? 'gut' : undefined} />
            <Kz lab="Wochenende" wert={de(r.wochenende.avg)} sub={`${de(r.wochenende.hoch)} min`} />
            <Kz lab="Tage mit harter Einheit" wert={de(r.hartTage.avg)} sub={`${r.hartTage.n} Tage`} />
          </div>
          <div className="h-40">
            <ResponsiveContainer>
              <BarChart data={r.proWt} margin={{ left: -10, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} />
                <YAxis yAxisId="a" tick={AX} width={30} domain={[0, 'auto']} />
                <YAxis yAxisId="b" orientation="right" tick={AX} width={36} />
                <Tooltip {...TT} formatter={(v: any, n: any) => [de(v), n]} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
                <Bar yAxisId="a" dataKey="avg" name="Ø Stress" fill="#f97316" />
                <Bar yAxisId="b" dataKey="hoch" name="min mittel/hoch" fill="#64748b" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title="Stress-Bilanz" facts={facts} />
        </>
      )}
    </Panel>
  );
}

// ─── 11 Energiebilanz ───────────────────────────────────────────────────────
export function EnergiebilanzTile() {
  const { tage } = useGarminTage();
  const tl = useTimeline();
  const r = useMemo(() => energiebilanz(tage, tl), [tage, tl]);
  const k = r?.khVorHart;
  const facts = r && k ? [
    `Gegessen (Abendformular) gegen Garmin-Gesamtverbrauch: letzte 4 Wochen im Schnitt ${vz(r.saldo4)} kcal/Tag (positiv = Überschuss).`,
    `Wochen: ${r.wochen.slice(-6).map((w) => `ab ${kurz(w.woche)} ${vz(w.saldoTag)} kcal/Tag`).join(', ')}.`,
    `Harte Einheiten nach wenig Kohlenhydraten am Vortag (< ${de(k.schwelle)} g, ${k.wenig.n}×): Puls je km ${vz(k.wenig.effRel, 1)} % zum Formstand, Bereitschaft danach ${de(k.wenig.ready)}. Nach vielen (${k.viel.n}×): ${vz(k.viel.effRel, 1)} %, Bereitschaft ${de(k.viel.ready)}.`,
    'Garmin schätzt den Verbrauch, Abweichungen von ±10 % sind normal.',
  ] : [];
  return (
    <Panel title="Energiebilanz">
      <Hinweis>Gegessene Kalorien aus deinem Abendformular gegen den Garmin-Gesamtverbrauch, als Saldo je Woche. Dazu: wie laufen harte Einheiten nach wenig oder vielen Kohlenhydraten am Vortag?</Hinweis>
      {!r ? <Laedt /> : (
        <>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <Kz lab="Saldo 4 Wochen" wert={`${vz(r.saldo4)} kcal/Tag`} sub="+ Überschuss · − Defizit" />
            <Kz lab="Wenig KH am Vortag" wert={`${vz(k?.wenig.effRel, 1)} %`} sub={`Puls je km · ${k?.wenig.n}× · < ${de(k?.schwelle)} g`} />
            <Kz lab="Viel KH am Vortag" wert={`${vz(k?.viel.effRel, 1)} %`} sub={`Puls je km · ${k?.viel.n}×`} />
          </div>
          <div className="h-44">
            <ResponsiveContainer>
              <BarChart data={r.wochen.slice(-20).map((w) => ({ ...w, lab: kurz(w.woche).slice(0, 5) }))} margin={{ left: -4, right: 8 }}>
                <CartesianGrid strokeOpacity={0.12} />
                <XAxis dataKey="lab" tick={AX} interval={1} />
                <YAxis tick={AX} width={44} />
                <ReferenceLine y={0} stroke="#94a3b8" />
                <Tooltip {...TT} formatter={(v: any) => [`${vz(v)} kcal/Tag`, 'Saldo']} />
                <Bar dataKey="saldoTag" name="Saldo je Tag" fill="#38bdf8" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <KiraFazit title="Energiebilanz" facts={facts} regeln={['Nur erfasste Ernährungstage; ähnliche Trainingslast und Essenszeit sind nicht kontrolliert. Keine Aussage „weniger essen hilft“.']} />
        </>
      )}
    </Panel>
  );
}

// ─── 12 Eiweiß und Erholung ─────────────────────────────────────────────────
export function EiweissTile() {
  const { tage } = useGarminTage();
  const tl = useTimeline();
  const r: any = useMemo(() => eiweiss(tage, tl), [tage, tl]);
  const facts = r?.terzile?.length ? [
    `${r.n} Tage, Eiweiß im Median ${de(r.proteinMedian)} g. Folgemorgen nach Drittel: ${r.terzile.map((t: any) => `${t.lab}: Bereitschaft ${de(t.ready)}, HRV ${de(t.hrv)} ms (${t.n})`).join('; ')}.`,
    `Nur Tage mit viel Last (ab ${de(r.mLast)}): ${r.hochLast.map((t: any) => `${t.lab}: Bereitschaft ${de(t.ready)}, HRV ${de(t.hrv)} (${t.n})`).join('; ')}.`,
  ] : [];
  return (
    <Panel title="Eiweiß und Erholung">
      <Hinweis>Hängen Eiweißmenge und Bereitschaft bzw. HRV am nächsten Morgen zusammen? Aufgeteilt in drei gleich große Gruppen, zusätzlich nur für Tage mit viel Trainingslast.</Hinweis>
      {!r ? <Laedt /> : !r.terzile?.length ? <p className="text-2xs text-ink-dim">Zu wenige Tage mit Ernährungsdaten.</p> : (
        <>
          <div className="grid md:grid-cols-2 gap-3">
            {[['alle Tage', r.terzile], [`Tage mit viel Last (ab ${de(r.mLast)})`, r.hochLast]].map(([lab, xs]: any) => (
              <div key={lab}>
                <div className="label mb-1">{lab}</div>
                <div className="h-36">
                  <ResponsiveContainer>
                    <BarChart data={xs} margin={{ left: -10, right: 8 }}>
                      <CartesianGrid strokeOpacity={0.12} />
                      <XAxis dataKey="kurz" tick={AX} interval={0} />
                      <YAxis yAxisId="r" tick={AX} width={34} domain={[0, 100]} />
                      <YAxis yAxisId="h" orientation="right" tick={AX} width={30} domain={[0, 'auto']} />
                      <Tooltip {...TT} labelFormatter={(_l, p: any) => p?.[0]?.payload ? `${p[0].payload.lab} · ${p[0].payload.n} Tage` : ''} formatter={(v: any, n: any) => [de(v, 1), n]} />
                      <Legend wrapperStyle={{ fontSize: 10 }} />
                      <Bar yAxisId="r" dataKey="ready" name="Bereitschaft" fill="#22c55e" />
                      <Bar yAxisId="h" dataKey="hrv" name="HRV ms" fill="#a855f7" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            ))}
          </div>
          <KiraFazit title="Eiweiß und Erholung" facts={facts} regeln={['Nur erfasste Ernährungstage; Makronährstoffe ändern sich gemeinsam. Keine feste Ernährungsregel ableiten.']} />
        </>
      )}
    </Panel>
  );
}
