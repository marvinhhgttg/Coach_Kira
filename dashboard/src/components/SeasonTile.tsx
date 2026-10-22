import { useEffect, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { fetchExtras, saveExtra, type ExtraItem } from '../lib/extras';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { fmtDateShort, fmtNum } from '../lib/format';
import { buildSeason, mondayOf, type Saisonwoche, type Ziel } from '../lib/planning';

interface SaisonDaten { wochen: Saisonwoche[]; parameter: { wochen: number; rampePct: number; startWeekly: number; maxWeekly: number } }
const heute = () => new Date().toLocaleDateString('sv-SE');
const PHASE_CLS: Record<string, string> = {
  Aufbau: 'text-accent',
  Entlastung: 'text-green-300',
  Tapering: 'text-yellow-300',
  Ziel: 'text-orange-300 font-semibold',
  Erholung: 'text-green-300',
};

export function SeasonTile() {
  const [item, setItem] = useState<ExtraItem<SaisonDaten> | null>(null);
  const [ziele, setZiele] = useState<Ziel[]>([]);
  const [startWeekly, setStartWeekly] = useState<number | null>(null);
  const [swLoading, setSwLoading] = useState(true);
  const [wochen, setWochen] = useState(16);
  const [rampe, setRampe] = useState(5);
  const [maxW, setMaxW] = useState<number | ''>('');
  const [busy, setBusy] = useState<string | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [vorschau, setVorschau] = useState<Saisonwoche[] | null>(null);

  useEffect(() => {
    fetchExtras<SaisonDaten>(['saison'], 1)
      .then((xs) => setItem(xs[0] || null))
      .catch((e) => setErr(/JSON|Typ|Failed to fetch|Load failed/.test(String(e)) ? 'Patch „Zukunftsplanung“ noch nicht aktiv.' : String(e)));
    fetchExtras<{ ziel: Ziel }>(['ziel'], 60)
      .then((xs) => setZiele(xs.map((x) => x.data?.ziel).filter((z): z is Ziel => !!z && z.datum >= heute())))
      .catch(() => setZiele([]));
    fetchChartData('28d' as any, ['load_fb_day'])
      .then((r: any) => {
        const xs = ((r && r.data) || []).map((x: any) => Number(x.load_fb_day) || 0);
        setStartWeekly(xs.length ? Math.round((xs.reduce((a: number, b: number) => a + b, 0) / xs.length) * 7) : null);
      })
      .catch(() => setStartWeekly(null))
      .finally(() => setSwLoading(false));
  }, []);

  const params = () => ({ startWeekly: startWeekly || 800, wochen, rampePct: rampe, maxWeekly: maxW ? Number(maxW) : Math.round((startWeekly || 800) * 1.6), ziele });

  async function erstellen() {
    setErr(null);
    const p = params();
    const ws = buildSeason(p);
    setVorschau(ws);
    setBusy('Kira formuliert den Fokus…');
    try {
      const tab = ws.map((w) => `${w.start} KW${w.kw} ${w.phase} ${w.ziel_ess} ESS | ${w.sportmix}${w.ziel ? ` | ZIEL: ${w.ziel}` : ''}`).join('\n');
      const zt = ziele.map((z) => `${z.datum} ${z.name} (${z.prioritaet}, ${z.hm} Hm, ${z.dauer_h} h)`).join('; ') || 'keine';
      const t = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, keine Emojis. Antworte nur mit JSON.' },
          {
            role: 'user',
            content:
              `Saisonrahmen (Wochen, Phase, Ziel-Wochenlast, Sportmix):\n${tab}\n\nZiele: ${zt}\nMarcs Ziele allgemein: Ausdauer für lange Bergtouren und Bike/Hike/Bike-Touren, Kondition und Tempohärte.\n\n` +
              'Gib für jede Woche einen kurzen Trainingsfokus (max. 12 Wörter, z. B. „lange Z2-Tage, ein Höhenmeter-Block am Wochenende“). Lastzahlen nicht ändern. ' +
              'Antworte nur mit JSON: {"wochen":[{"start":"2026-10-26","fokus":"…"}],"fazit":"2 Sätze zur Saison"}',
          },
        ],
        { temperature: 0.3, numCtx: 8192 },
      );
      const j: any = parseJsonLoose<any>(t) || {};
      const fk = new Map<string, string>(((Array.isArray(j.wochen) ? j.wochen : []) as any[]).map((x) => [String(x.start ?? '').slice(0, 10), String(x.fokus ?? '')]));
      const ws2 = ws.map((w) => ({ ...w, fokus: fk.get(w.start) || '' }));
      const txt = `Saisonrahmen ${ws2.length} Wochen ab ${ws2[0].start}: ${ws2.filter((w) => w.phase === 'Aufbau').length} Aufbau-, ${ws2.filter((w) => w.phase === 'Entlastung').length} Entlastungswochen, Wochenlast ${ws2[0].ziel_ess} → ${Math.max(...ws2.map((w) => w.ziel_ess))} ESS. ${j.fazit ? String(j.fazit) : ''}`;
      await saveExtra('saison', heute(), txt, { wochen: ws2, parameter: { wochen, rampePct: rampe, startWeekly: p.startWeekly, maxWeekly: p.maxWeekly } });
      const xs = await fetchExtras<SaisonDaten>(['saison'], 1);
      setItem(xs[0] || null);
      setVorschau(null);
    } catch (e) {
      setErr(String(e instanceof Error ? e.message : e));
    } finally {
      setBusy(null);
    }
  }

  const ws = vorschau || item?.data?.wochen || [];
  const thisMon = mondayOf(new Date()).toLocaleDateString('sv-SE');

  return (
    <div className="panel-raised p-4">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Saisonrahmen</div>
          <p className="text-2xs text-ink-dim mt-1">
            12–24 Wochen im Voraus: Aufbau im 3:1-Rhythmus, Tapering vor A-Zielen, Erholung danach, Sportmix nach Jahreszeit. Der KI-Report und der Wochen-Autopilot richten sich danach.
          </p>
        </div>
      </div>
      <div className="mt-3 flex flex-wrap items-end gap-3 text-xs">
        <label className="flex flex-col gap-1">Wochen
          <select className="input" value={wochen} onChange={(e) => setWochen(Number(e.target.value))}>{[12, 16, 20, 24].map((x) => <option key={x} value={x}>{x}</option>)}</select>
        </label>
        <label className="flex flex-col gap-1">Steigerung je Aufbauwoche
          <select className="input" value={rampe} onChange={(e) => setRampe(Number(e.target.value))}>{[3, 4, 5, 6, 8].map((x) => <option key={x} value={x}>{x} %</option>)}</select>
        </label>
        <label className="flex flex-col gap-1">Höchste Wochenlast
          <input className="input w-28 tnum" type="number" placeholder={startWeekly ? String(Math.round(startWeekly * 1.6)) : ''} value={maxW} onChange={(e) => setMaxW(e.target.value ? Number(e.target.value) : '')} />
        </label>
        <div className="text-ink-muted pb-1">Start: Ø {startWeekly != null ? fmtNum(startWeekly) : swLoading ? 'lädt…' : '–'} ESS/Woche (letzte 4 Wochen) · {ziele.length} Ziel(e)</div>
        <button className="btn btn-primary text-xs px-2 py-1" onClick={erstellen} disabled={!!busy || startWeekly == null}>{busy || (item ? 'Neu erstellen' : 'Saisonrahmen erstellen')}</button>
      </div>
      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}
      {item && !vorschau && <p className="mt-2 text-xs text-ink-muted">{item.text}</p>}
      {ws.length > 0 && (
        <div className="mt-3 overflow-x-auto">
          <table className="w-full text-xs text-left">
            <thead>
              <tr className="border-b border-border text-ink-dim">
                <th className="py-1 pr-3 font-normal">KW</th>
                <th className="py-1 pr-3 font-normal">ab</th>
                <th className="py-1 pr-3 font-normal">Phase</th>
                <th className="py-1 pr-3 font-normal text-right">Ziel ESS</th>
                <th className="py-1 pr-3 font-normal">Fokus</th>
                <th className="py-1 font-normal">Sportmix</th>
              </tr>
            </thead>
            <tbody>
              {ws.map((w) => (
                <tr key={w.start} className={`border-b border-border/60 ${w.start === thisMon ? 'bg-accent/10' : ''}`}>
                  <td className="py-1 pr-3 tnum">{w.kw}</td>
                  <td className="py-1 pr-3 tnum whitespace-nowrap">{fmtDateShort(w.start)}</td>
                  <td className={`py-1 pr-3 whitespace-nowrap ${PHASE_CLS[w.phase] || ''}`}>{w.phase}{w.ziel ? ` · ${w.ziel}` : ''}</td>
                  <td className="py-1 pr-3 tnum text-right">{w.ziel_ess}</td>
                  <td className="py-1 pr-3 text-ink">{w.fokus || '–'}</td>
                  <td className="py-1 text-ink-dim">{w.sportmix}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
