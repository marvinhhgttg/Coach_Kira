import { useMemo, useState } from 'react';
import { Panel } from './UI';
import { fmtNum, fmtDateShort } from '../lib/format';
import { useGarminTage, fmtDauer, type GarminTage, type GarminAkt } from '../lib/garmin';

type Periode = 'woche' | 'monat' | 'jahr' | 'alle';
const PERIODEN: { k: Periode; l: string }[] = [
  { k: 'woche', l: 'Woche' }, { k: 'monat', l: 'Monat' }, { k: 'jahr', l: 'Jahr' }, { k: 'alle', l: 'Alle' },
];

const iso = (d: Date) => d.toLocaleDateString('sv-SE');
function grenzen(p: Periode, versatz: number): { von: string; bis: string; titel: string } {
  const h = new Date(); h.setHours(12, 0, 0, 0);
  if (p === 'woche') {
    const mo = new Date(h); mo.setDate(h.getDate() - ((h.getDay() + 6) % 7) - 7 * versatz);
    const so = new Date(mo); so.setDate(mo.getDate() + 6);
    return { von: iso(mo), bis: iso(so), titel: `${fmtDateShort(iso(mo))}–${fmtDateShort(iso(so))}` };
  }
  if (p === 'monat') {
    const a = new Date(h.getFullYear(), h.getMonth() - versatz, 1, 12);
    const e = new Date(a.getFullYear(), a.getMonth() + 1, 0, 12);
    return { von: iso(a), bis: iso(e), titel: a.toLocaleDateString('de-DE', { month: 'long', year: 'numeric' }) };
  }
  if (p === 'jahr') {
    const y = h.getFullYear() - versatz;
    return { von: `${y}-01-01`, bis: `${y}-12-31`, titel: String(y) };
  }
  return { von: '0000', bis: '9999', titel: 'alle Daten' };
}

type Werte = { n: number; km: number; hm: number; sek: number; kcal: number; tage: number; sport: Map<string, { n: number; km: number; hm: number; sek: number }> };
type Best = { label: string; wert: string; datum: string; name?: string };

function auswerten(t: GarminTage, von: string, bis: string): { w: Werte; best: Best[] } {
  const w: Werte = { n: 0, km: 0, hm: 0, sek: 0, kcal: 0, tage: 0, sport: new Map() };
  const alle: { d: string; a: GarminAkt }[] = [];
  for (const [d, g] of t) {
    if (d < von || d > bis || !g.akt?.length) continue;
    w.tage++;
    for (const a of g.akt) {
      alle.push({ d, a });
      const sek = a.bewegt_s || a.dauer_s || 0;
      w.n++; w.km += a.km || 0; w.hm += a.hm || 0; w.sek += sek; w.kcal += a.kcal || 0;
      const s = a.sport || 'Sonstige';
      const x = w.sport.get(s) || { n: 0, km: 0, hm: 0, sek: 0 };
      x.n++; x.km += a.km || 0; x.hm += a.hm || 0; x.sek += sek;
      w.sport.set(s, x);
    }
  }
  const top = (f: (a: GarminAkt) => number | undefined | null) => {
    let b: { d: string; a: GarminAkt; v: number } | null = null;
    for (const x of alle) { const v = f(x.a); if (v != null && Number.isFinite(v) && (!b || v > b.v)) b = { ...x, v }; }
    return b;
  };
  const best: Best[] = [];
  const add = (label: string, b: ReturnType<typeof top>, fmt: (v: number) => string) => { if (b) best.push({ label, wert: fmt(b.v), datum: b.d, name: b.a.name }); };
  add('Längste Distanz', top((a) => a.km), (v) => `${fmtNum(v, 1)} km`);
  add('Meiste Höhenmeter', top((a) => a.hm), (v) => `${fmtNum(v)} hm`);
  add('Längste Dauer', top((a) => a.bewegt_s || a.dauer_s), (v) => fmtDauer(v));
  add('Höchste Belastung', top((a) => a.last), (v) => fmtNum(v));
  add('Bester TE aerob', top((a) => a.te_ae), (v) => fmtNum(v, 1));
  add('Höchste 20-min-Leistung Rad', top((a) => (a.sport === 'Bike' ? a.watt_20 : null)), (v) => `${fmtNum(v)} W`);
  add('Höchste NP Rad', top((a) => (a.sport === 'Bike' ? a.watt_np : null)), (v) => `${fmtNum(v)} W`);
  add('Höchster VO2max', top((a) => a.vo2max), (v) => fmtNum(v));
  add('Höchster Punkt', top((a) => a.hoehe_max), (v) => `${fmtNum(v)} m`);
  return { w, best };
}

function Delta({ jetzt, vorher, einheit = '' }: { jetzt: number; vorher: number; einheit?: string }) {
  if (!vorher) return null;
  const d = jetzt - vorher;
  const pct = Math.round((d / vorher) * 100);
  return <span className={`text-2xs ${d >= 0 ? 'text-green-300' : 'text-orange-300'}`}>{d >= 0 ? '+' : ''}{fmtNum(d)}{einheit} ({pct >= 0 ? '+' : ''}{pct} %)</span>;
}

export function GarminBerichtTile() {
  const { tage } = useGarminTage();
  const [p, setP] = useState<Periode>('monat');
  const [versatz, setVersatz] = useState(0);
  const g = grenzen(p, versatz);
  const gv = grenzen(p, versatz + 1);
  const res = useMemo(() => (tage ? auswerten(tage, g.von, g.bis) : null), [tage, g.von, g.bis]);
  const vor = useMemo(() => (tage && p !== 'alle' ? auswerten(tage, gv.von, gv.bis) : null), [tage, gv.von, gv.bis, p]);

  return (
    <Panel
      title="Umfang und Bestwerte · Garmin"
      right={
        <div className="flex items-center gap-2">
          <div className="inline-flex bg-bg rounded border border-border p-0.5">
            {PERIODEN.map((x) => (
              <button key={x.k} onClick={() => { setP(x.k); setVersatz(0); }} className={`px-2 py-1 text-xs rounded ${p === x.k ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}>{x.l}</button>
            ))}
          </div>
          {p !== 'alle' && (
            <div className="inline-flex items-center gap-1 text-xs">
              <button className="btn btn-ghost px-2 py-1" onClick={() => setVersatz((v) => v + 1)} aria-label="früher">‹</button>
              <button className="btn btn-ghost px-2 py-1" disabled={versatz === 0} onClick={() => setVersatz((v) => Math.max(0, v - 1))} aria-label="später">›</button>
            </div>
          )}
        </div>
      }
    >
      {tage === null || !res ? (
        <div className="text-sm text-ink-muted">Lade Garmin-Daten …</div>
      ) : (
        <div className="space-y-3" data-testid="garmin-bericht">
          <div className="text-xs text-ink-muted">{g.titel}{vor && p !== 'alle' ? ` · Vergleich mit ${gv.titel}` : ''}</div>
          <div className="grid gap-2 grid-cols-2 md:grid-cols-5">
            {[
              ['Einheiten', res.w.n, vor?.w.n, ''],
              ['Distanz', res.w.km, vor?.w.km, ' km'],
              ['Höhenmeter', res.w.hm, vor?.w.hm, ' hm'],
              ['Bewegungszeit', res.w.sek / 3600, vor ? vor.w.sek / 3600 : undefined, ' h'],
              ['Kalorien (Training)', res.w.kcal, vor?.w.kcal, ' kcal'],
            ].map(([l, v, pv, e]) => (
              <div key={l as string} className="rounded border border-border p-2">
                <div className="label">{l as string}</div>
                <div className="text-lg font-semibold tnum">{fmtNum(v as number)}{e as string}</div>
                {pv != null && <Delta jetzt={v as number} vorher={pv as number} einheit={e as string} />}
              </div>
            ))}
          </div>
          {res.w.sport.size > 0 && (
            <div className="overflow-x-auto">
              <table className="w-full min-w-[520px] text-xs tnum">
                <thead><tr className="text-left">{['Sport', 'Einh.', 'km', 'hm', 'Zeit', 'Anteil Zeit'].map((h) => <th key={h} className="label py-1 pr-2 font-normal">{h}</th>)}</tr></thead>
                <tbody>
                  {[...res.w.sport.entries()].sort((a, b) => b[1].sek - a[1].sek).map(([s, x]) => (
                    <tr key={s} className="border-t border-border">
                      <td className="py-1 pr-2 font-semibold">{s}</td>
                      <td className="py-1 pr-2">{x.n}</td>
                      <td className="py-1 pr-2">{fmtNum(x.km)}</td>
                      <td className="py-1 pr-2">{fmtNum(x.hm)}</td>
                      <td className="py-1 pr-2">{fmtNum(x.sek / 3600, 1)} h</td>
                      <td className="py-1 pr-2">{res.w.sek ? Math.round((x.sek / res.w.sek) * 100) : 0} %</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {res.best.length > 0 && (
            <div>
              <div className="label mb-1">Bestwerte im Zeitraum</div>
              <div className="grid gap-1 sm:grid-cols-2 lg:grid-cols-3 text-xs">
                {res.best.map((b) => (
                  <div key={b.label} className="flex items-baseline gap-2 min-w-0">
                    <span className="text-ink-muted shrink-0">{b.label}</span>
                    <span className="font-semibold tnum shrink-0">{b.wert}</span>
                    <span className="text-ink-dim truncate">{fmtDateShort(b.datum)}{b.name ? ` · ${b.name}` : ''}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
          {res.w.n === 0 && <div className="text-sm text-ink-muted">Keine Garmin-Aktivitäten in diesem Zeitraum.</div>}
        </div>
      )}
    </Panel>
  );
}
