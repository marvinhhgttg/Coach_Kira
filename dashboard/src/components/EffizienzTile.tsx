import { useMemo, useState } from 'react';
import { CartesianGrid, ComposedChart, Line, ResponsiveContainer, Scatter, Tooltip, XAxis, YAxis } from 'recharts';
import { Panel } from './UI';
import { KiraFazit } from './KiraFazit';
import { useGarminTage } from '../lib/garmin';
import {
  effPunkte, gleitend, GRENZEN_STANDARD, nachGelaende, trend, wertVon,
  type EffPunkt, type EffTrend, type Gelaende, type Kennzahl, type Sportart, type Untergrund,
} from '../lib/effizienz';

const de = (v: number | null | undefined, d = 0) => (v == null || !Number.isFinite(v) ? '–' : v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d }));
const tempoTxt = (t: number | null) => (t == null ? '–' : `${Math.floor(t)}:${String(Math.round((t % 1) * 60)).padStart(2, '0')}`);
const datumKurz = (ts: number) => new Date(ts).toLocaleDateString('de-DE', { month: 'short', year: '2-digit' });
const datumTag = (d: string) => new Date(d + 'T12:00:00').toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: '2-digit' });

const KENNZAHL: Record<Sportart, Record<Kennzahl, { titel: string; einheit: string; erklaer: string; d: number }>> = {
  lauf: {
    hf: { titel: 'Ø HF zu Ø Tempo', einheit: 'Schläge/km', erklaer: 'Ø HF × Pace (min/km) – Herzschläge für einen Kilometer', d: 0 },
    leistung: { titel: 'Ø Laufleistung zu Ø Tempo', einheit: 'W je km/h', erklaer: 'Ø Laufleistung ÷ Ø km/h – Watt für 1 km/h Tempo', d: 1 },
  },
  rad: {
    hf: { titel: 'Ø HF zu Ø Geschwindigkeit', einheit: 'Schläge/km', erklaer: 'Ø HF × min/km – Herzschläge für einen Kilometer', d: 0 },
    leistung: { titel: 'Normalized Power zu Ø Geschwindigkeit', einheit: 'W je km/h', erklaer: 'NP ÷ Ø km/h – Watt für 1 km/h Geschwindigkeit', d: 2 },
  },
};
const FARBE: Record<string, string> = { running: '#22c55e', trail_running: '#a3e635', road_biking: '#38bdf8', cycling: '#0ea5e9', gravel_cycling: '#f59e0b', mountain_biking: '#ef4444' };
const TYP_NAME: Record<string, string> = { running: 'Laufen', trail_running: 'Trail', road_biking: 'Rennrad', cycling: 'Rad', gravel_cycling: 'Gravel', mountain_biking: 'MTB' };

/** Ein Tick je 1./2./3. Monat, je nach Zeitraum – nie mehr als ~7 Beschriftungen. */
function monatsTicks(a?: number, b?: number): number[] {
  if (a == null || b == null) return [];
  const monate = Math.max(1, Math.round((b - a) / (30.4 * 86400000)));
  const schritt = monate > 18 ? 3 : monate > 8 ? 2 : 1;
  const d = new Date(a); d.setDate(1); d.setHours(12, 0, 0, 0); d.setMonth(d.getMonth() + 1);
  const out: number[] = [];
  while (d.getTime() <= b) { if ((d.getMonth() % schritt) === 0) out.push(d.getTime()); d.setMonth(d.getMonth() + 1); }
  return out;
}

function trendTxt(t: EffTrend) {
  if (t.prozent == null) return { txt: t.n < 2 ? 'zu wenig Einheiten' : 'Zeitraum zu kurz', cls: 'text-ink-dim' };
  const p = t.prozent;
  const cls = Math.abs(p) < 2 ? 'text-ink-muted' : p < 0 ? 'text-emerald-400' : 'text-amber-400';
  const wort = Math.abs(p) < 2 ? 'etwa gleich' : p < 0 ? 'effizienter' : 'weniger effizient';
  return { txt: `${p > 0 ? '+' : ''}${de(p, 1)} % (${wort})`, cls };
}

function Diagramm({ ps, k, sport }: { ps: EffPunkt[]; k: Kennzahl; sport: Sportart }) {
  const meta = KENNZAHL[sport][k];
  const t = useMemo(() => trend(ps, k), [ps, k]);
  const rows = useMemo(() => {
    const g = gleitend(ps, k, 5);
    const xs = ps.map((p, i) => ({ ...p, y: wertVon(p, k), med: g[i] })).filter((r) => r.y != null);
    if (xs.length >= 2 && t.linieStart != null && t.linieEnde != null) {
      xs[0] = { ...xs[0], lin: t.linieStart } as any;
      xs[xs.length - 1] = { ...xs[xs.length - 1], lin: t.linieEnde } as any;
    }
    return xs;
  }, [ps, k, t]);
  const tt = trendTxt(t);
  const ticks = useMemo(() => monatsTicks(rows[0]?.ts, rows[rows.length - 1]?.ts), [rows]);
  return (
    <div className="rounded border border-border p-2" data-testid={`eff-${sport}-${k}`}>
      <div className="flex items-baseline justify-between gap-2">
        <div className="text-xs text-ink">{meta.titel}</div>
        <div className={`text-2xs tnum ${tt.cls}`}>{tt.txt}</div>
      </div>
      <div className="text-2xs text-ink-dim">{meta.erklaer} · niedriger = effizienter</div>
      {sport === 'rad' && k === 'leistung' && <div className="text-2xs text-ink-dim">Achtung: steigt bei höherem Tempo von selbst (Luftwiderstand) – schneller unterwegs heißt hier nicht automatisch schlechter.</div>}
      {rows.length < 2 ? (
        <p className="text-2xs text-ink-dim mt-3 mb-2">{k === 'leistung' ? 'Zu wenige Einheiten mit Leistungsdaten.' : 'Zu wenige Einheiten in diesem Gelände.'}</p>
      ) : (
        <div className="h-44 mt-1">
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={rows} margin={{ top: 4, right: 8, left: -4, bottom: 0 }}>
              <CartesianGrid strokeOpacity={0.12} />
              <XAxis dataKey="ts" type="number" domain={['dataMin', 'dataMax']} ticks={ticks} tickFormatter={datumKurz} tick={{ fontSize: 10 }} />
              <YAxis domain={['auto', 'auto']} tick={{ fontSize: 10 }} width={52} tickFormatter={(v) => de(v, meta.d > 1 ? 1 : 0)} />
              <Tooltip
                content={({ active, payload }) => {
                  if (!active || !payload?.length) return null;
                  const p = payload[0].payload as EffPunkt & { y: number; med: number | null };
                  return (
                    <div className="panel-raised p-2 text-2xs space-y-0.5">
                      <div className="text-ink">{datumTag(p.datum)} · {TYP_NAME[p.typ] || p.typ}</div>
                      {p.name && <div className="text-ink-dim">{p.name}</div>}
                      <div className="tnum">{de(p.y, meta.d)} {meta.einheit}{p.med != null ? ` · Verlauf ${de(p.med, meta.d)}` : ''}</div>
                      <div className="tnum text-ink-muted">{de(p.km, 1)} km · {de(p.hm)} hm ({de(p.hmProKm, 1)} hm/km) · {de(p.minuten)} min</div>
                      <div className="tnum text-ink-muted">Ø HF {de(p.hf)} · {sport === 'lauf' ? `${tempoTxt(p.tempo)} min/km` : `${de(p.v, 1)} km/h`}{p.watt != null ? ` · ${sport === 'lauf' ? 'Ø' : 'NP'} ${de(p.watt)} W` : ''}</div>
                    </div>
                  );
                }}
              />
              <Scatter dataKey="y" isAnimationActive={false} shape={(s: any) => <circle cx={s.cx} cy={s.cy} r={2.6} fill={FARBE[s.payload.typ] || '#94a3b8'} fillOpacity={0.75} />} />
              <Line dataKey="med" stroke="#e5e7eb" strokeWidth={1.6} dot={false} connectNulls isAnimationActive={false} />
              <Line dataKey="lin" stroke="#a855f7" strokeDasharray="4 3" strokeWidth={1.2} dot={false} connectNulls isAnimationActive={false} />
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      )}
      {t.n >= 2 && (
        <div className="grid grid-cols-2 gap-2 mt-1 text-2xs tnum text-ink-muted">
          {([['erste 90 Tage', t.frueh], ['letzte 90 Tage', t.spaet]] as const).map(([lab, b]) => (
            <div key={lab}>
              <span className="text-ink-dim">{lab} ({b.n}):</span> {de(b.wert, meta.d)} · HF {de(b.hf)} · {sport === 'lauf' ? `${tempoTxt(b.tempo)}/km` : `${de(b.v, 1)} km/h`} · {de(b.hmProKm, 1)} hm/km
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function EffizienzTile() {
  const { tage } = useGarminTage();
  const [sport, setSport] = useState<Sportart>('lauf');
  const [untergrund, setUntergrund] = useState<Untergrund>('alle');
  const [grenzen, setGrenzen] = useState(GRENZEN_STANDARD);
  const gr = grenzen[sport];

  const alle = useMemo(() => effPunkte(tage, sport, sport === 'rad' ? untergrund : 'alle'), [tage, sport, untergrund]);
  const gelaende = useMemo(() => ({ flach: nachGelaende(alle, 'flach', gr), berg: nachGelaende(alle, 'berg', gr) }), [alle, gr]);

  const facts = useMemo(() => {
    const f: string[] = [];
    const sp = sport === 'lauf' ? 'Laufen' : `Radfahren (${untergrund === 'alle' ? 'Straße und Gravel' : untergrund === 'strasse' ? 'nur Straße' : 'nur Gravel'})`;
    f.push(`${sp}, Garmin-Aktivitäten ${alle[0] ? datumTag(alle[0].datum) : '–'} bis ${alle.length ? datumTag(alle[alle.length - 1].datum) : '–'}. Flach = höchstens ${gr.flach} hm/km, viele Höhenmeter = mindestens ${gr.berg} hm/km. Bei allen Kennzahlen gilt: niedriger = effizienter.`);
    for (const g of ['flach', 'berg'] as Gelaende[]) {
      for (const k of ['hf', 'leistung'] as Kennzahl[]) {
        const t = trend(gelaende[g], k), m = KENNZAHL[sport][k];
        f.push(`${g === 'flach' ? 'Flach' : 'Viele Höhenmeter'} – ${m.titel} (${m.einheit}): ${t.n} Einheiten, Trendlinie ${de(t.linieStart, m.d)} → ${de(t.linieEnde, m.d)} (${t.prozent == null ? 'kein Trend berechenbar' : `${t.prozent > 0 ? '+' : ''}${de(t.prozent, 1)} %`}); erste 90 Tage Median ${de(t.frueh.wert, m.d)} bei ${de(t.frueh.hmProKm, 1)} hm/km, HF ${de(t.frueh.hf)} und ${sport === 'lauf' ? tempoTxt(t.frueh.tempo) + ' min/km' : de(t.frueh.v, 1) + ' km/h'}; letzte 90 Tage Median ${de(t.spaet.wert, m.d)} bei ${de(t.spaet.hmProKm, 1)} hm/km, HF ${de(t.spaet.hf)} und ${sport === 'lauf' ? tempoTxt(t.spaet.tempo) + ' min/km' : de(t.spaet.v, 1) + ' km/h'}.`);
      }
    }
    f.push('Hinweis: Wetter, Wind, Untergrund und Gruppenfahrten beeinflussen die Werte; Leistungsdaten gibt es nicht für alle Einheiten. Bei vielen Höhenmetern hängen die Werte stark von der Steilheit ab – die hm/km der ersten und letzten 90 Tage mitlesen. Laufleistung kann zwischen Uhrenmodellen unterschiedlich skaliert sein. Beim Rad steigt NP je km/h mit höherem Tempo von selbst (Luftwiderstand wächst mit dem Quadrat der Geschwindigkeit), ein Anstieg dort ist bei gleichzeitig höherem Tempo kein Formverlust.');
    return f;
  }, [alle, gelaende, gr, sport, untergrund]);

  const btn = (aktiv: boolean) => `px-2 py-1 text-xs rounded ${aktiv ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;
  const setGrenze = (art: 'flach' | 'berg', v: number) => setGrenzen((g) => ({ ...g, [sport]: { ...g[sport], [art]: Math.max(0, Math.min(40, v || 0)) } }));

  return (
    <Panel
      title="Effizienz flach und bergig"
      right={
        <div className="inline-flex bg-bg rounded border border-border p-0.5">
          <button className={btn(sport === 'lauf')} onClick={() => setSport('lauf')}>Laufen</button>
          <button className={btn(sport === 'rad')} onClick={() => setSport('rad')}>Rad</button>
        </div>
      }
    >
      <div data-testid="effizienz">
        <p className="text-2xs text-ink-dim mb-2">
          Wie viel Puls und Leistung kostet dich dein Tempo – und wie entwickelt sich das? Aus den Garmin-Aktivitäten
          (draußen, {sport === 'lauf' ? 'ab 20 min, 3 km und Ø HF 115' : 'ab 30 min, 10 km und Ø HF 105; flach erst ab 15 km/h'}, Tempo aus der Bewegungszeit). Punkte = Einheiten, weiße Linie = Verlauf (Median der letzten 5),
          lila gestrichelt = robuster Trend (Ausreißer kippen ihn nicht).
        </p>
        <div className="flex flex-wrap items-center gap-x-4 gap-y-1 mb-3 text-2xs text-ink-muted">
          <label className="inline-flex items-center gap-1">flach bis
            <input type="number" min={0} max={40} step={0.5} value={gr.flach} onChange={(e) => setGrenze('flach', Number(e.target.value))} className="w-12 bg-bg border border-border rounded px-1 py-0.5 tnum text-ink" /> hm/km
          </label>
          <label className="inline-flex items-center gap-1">viele Höhenmeter ab
            <input type="number" min={0} max={40} step={0.5} value={gr.berg} onChange={(e) => setGrenze('berg', Number(e.target.value))} className="w-12 bg-bg border border-border rounded px-1 py-0.5 tnum text-ink" /> hm/km
          </label>
          {sport === 'rad' && (
            <div className="inline-flex bg-bg rounded border border-border p-0.5">
              {(['alle', 'strasse', 'gravel'] as Untergrund[]).map((u) => (
                <button key={u} className={btn(untergrund === u)} onClick={() => setUntergrund(u)}>{u === 'alle' ? 'alle' : u === 'strasse' ? 'Straße' : 'Gravel'}</button>
              ))}
            </div>
          )}
          <span>
            {Object.entries(FARBE).filter(([t]) => alle.some((p) => p.typ === t)).map(([t, c]) => (
              <span key={t} className="mr-2"><span style={{ color: c }}>●</span> {TYP_NAME[t]}</span>
            ))}
          </span>
        </div>
        {tage === null ? (
          <p className="text-2xs text-ink-dim">lädt …</p>
        ) : !alle.length ? (
          <p className="text-2xs text-ink-dim">Keine passenden Garmin-Aktivitäten gefunden.</p>
        ) : (
          (['flach', 'berg'] as Gelaende[]).map((g) => (
            <div key={g} className="mb-3">
              <div className="label mb-1">
                {g === 'flach' ? `Vorwiegend flach (≤ ${de(gr.flach, 1)} hm/km)` : `Viele Höhenmeter (≥ ${de(gr.berg, 1)} hm/km)`} · {gelaende[g].length} Einheiten
              </div>
              <div className="grid gap-2 md:grid-cols-2">
                <Diagramm ps={gelaende[g]} k="hf" sport={sport} />
                <Diagramm ps={gelaende[g]} k="leistung" sport={sport} />
              </div>
            </div>
          ))
        )}
        {alle.length > 0 && <KiraFazit title={`Effizienz ${sport === 'lauf' ? 'Laufen' : 'Rad'} flach und bergig`} facts={facts} />}
      </div>
    </Panel>
  );
}
