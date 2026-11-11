import { useEffect, useState } from 'react';
import { fetchExtras, saveExtra, type ExtraItem } from '../lib/extras';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { fmtDateShort } from '../lib/format';
import { monotonyPlanViolations } from '../lib/monotony';
import type { PlanRow } from '../screens/StrategyPanel';

interface Befund { art: string; schwere: 'hoch' | 'mittel' | 'niedrig'; text: string }
const GRUENDE = [
  { v: 'infekt', l: 'Ich fühle mich krank / Infekt' },
  { v: 'muede', l: 'Sehr müde, schlecht geschlafen' },
  { v: 'ausfall', l: 'Training ausgefallen' },
  { v: 'reise', l: 'Reise / volle Tage' },
];
const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const wd = (d: string) => WT[new Date(d + 'T12:00:00').getDay()];
const heute = () => new Date().toLocaleDateString('sv-SE');

export function DisruptionTile({
  sim,
  loads,
  sports,
  zones,
  locks,
  simulateWith,
  onApply,
}: {
  sim: PlanRow[];
  loads: number[];
  sports: string[];
  zones: string[];
  locks: boolean[];
  simulateWith: (loads: number[]) => PlanRow[];
  onApply: (loads: number[], sports: string[], zones: string[]) => void;
}) {
  const [item, setItem] = useState<ExtraItem<{ befunde: Befund[] }> | null>(null);
  const [grund, setGrund] = useState('');
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);
  const [vor, setVor] = useState<{ loads: number[]; sports: string[]; zones: string[]; aend: { i: number; grund: string }[]; fazit: string; v0: number; v1: number } | null>(null);

  useEffect(() => {
    fetchExtras<{ befunde: Befund[] }>(['umplanung'], 3)
      .then((xs) => setItem(xs.find((x) => x.datum >= new Date(Date.now() - 86400000).toLocaleDateString('sv-SE')) || null))
      .catch(() => setItem(null));
  }, []);

  const befunde = item?.data?.befunde || [];
  const wichtig = befunde.filter((b) => b.schwere !== 'niedrig');

  async function umplanen() {
    setBusy(true);
    setErr(null);
    setVor(null);
    try {
      const anlass = [
        ...befunde.map((b) => b.text),
        ...(grund ? [GRUENDE.find((g) => g.v === grund)?.l || grund] : []),
      ].join(' | ');
      if (!anlass) throw new Error('Kein Anlass erkannt – wähle oben einen Grund.');
      const rest = sim.map((r, i) => ({ i, r })).filter(({ i }) => i <= 7);
      const tab = rest.map(({ i, r }) => `${r.date} ${wd(r.date)} | ${Math.round(loads[i])} ESS ${sports[i]} ${zones[i]}${locks[i] ? ' FEST' : ''}${i === 0 ? ' (heute)' : ''}`).join('\n');
      const t = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, keine Emojis. Antworte nur mit JSON.' },
          {
            role: 'user',
            content:
              `Anlass: ${anlass}\n\nPlan heute und die nächsten 7 Tage:\n${tab}\n\n` +
              'Passe den Rest-Plan an. Regeln: FEST-Tage nicht ändern. Bei Infekt-Verdacht: heute und morgen Ruhe oder höchstens 20–30 min sehr locker, danach 2–3 Tage sanfter Wiedereinstieg (≤ 60 % der normalen Last, nur Z1/Z2), keine Intensität bis die Werte normal sind. ' +
              'Bei Müdigkeit: heute locker, Qualität um 1–2 Tage schieben. Bei ausgefallenem Training: nicht nachholen, höchstens eine Einheit leicht verlängern. Bei Reise: kurze Einheiten oder Ruhetag an den Reisetagen. ' +
              'Gesamtlast lieber senken als verdichten; Monotonie-Regel beachten. ' +
              'Antworte nur mit JSON mit ALLEN Tagen: {"tage":[{"datum":"2026-10-22","last":0,"sport":"Off","zone":"Off","grund":"…"}],"fazit":"1–2 Sätze"}',
          },
        ],
        { temperature: 0.2, numCtx: 8192 },
      );
      const j: any = parseJsonLoose<any>(t) || {};
      const by = new Map<string, any>(((Array.isArray(j.tage) ? j.tage : []) as any[]).map((x) => [String(x.datum ?? '').slice(0, 10), x]));
      const nl = loads.slice(), ns = sports.slice(), nz = zones.slice();
      const aend: { i: number; grund: string }[] = [];
      rest.forEach(({ i, r }) => {
        const x = by.get(r.date);
        if (!x || locks[i]) return;
        const l = Math.max(0, Math.min(400, Math.round(Number(x.last ?? nl[i]))));
        const s = String(x.sport || ns[i]);
        const z = String(x.zone || nz[i]);
        if (l !== Math.round(nl[i]) || s !== ns[i] || z !== nz[i]) aend.push({ i, grund: String(x.grund || '') });
        nl[i] = l; ns[i] = s; nz[i] = z;
      });
      if (!aend.length) throw new Error('Kira hat keine Änderung vorgeschlagen.');
      const mv = (ls: number[], rows: PlanRow[]) => monotonyPlanViolations(rows.map((r, i) => ({ label: r.date, load: ls[i], mono: i === 0 ? null : r.monotony, locked: locks[i] }))).length;
      setVor({ loads: nl, sports: ns, zones: nz, aend, fazit: String(j.fazit || ''), v0: mv(loads, sim), v1: mv(nl, simulateWith(nl)) });
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className={`panel-raised p-4 ${wichtig.length ? 'border border-orange-400/40' : ''}`}>
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Umplanen bei Störungen</div>
          <p className="text-2xs text-ink-dim mt-1">Der Mac prüft jeden Morgen auf Infekt-Verdacht, starke Müdigkeit, ausgefallenes Training und Reisetage. Du kannst auch selbst einen Grund wählen.</p>
        </div>
      </div>
      {befunde.length > 0 ? (
        <ul className="mt-2 text-xs flex flex-col gap-0.5">
          {befunde.map((b, i) => (
            <li key={i} className={b.schwere === 'hoch' ? 'text-orange-300' : b.schwere === 'mittel' ? 'text-yellow-300' : 'text-ink-muted'}>{b.text}</li>
          ))}
          <li className="text-2xs text-ink-dim">Erkannt am {item && fmtDateShort(item.datum)}</li>
        </ul>
      ) : (
        <p className="mt-2 text-xs text-ink-muted">Heute keine Störung erkannt.</p>
      )}
      <div className="mt-3 flex flex-wrap items-center gap-2 text-xs">
        <select className="input" value={grund} onChange={(e) => setGrund(e.target.value)}>
          <option value="">{befunde.length ? 'zusätzlicher Grund (optional)' : 'Grund wählen …'}</option>
          {GRUENDE.map((g) => <option key={g.v} value={g.v}>{g.l}</option>)}
        </select>
        <button className="btn btn-primary text-xs px-2 py-1" onClick={umplanen} disabled={busy || (!befunde.length && !grund)}>{busy ? 'Kira plant um…' : 'Rest-Plan anpassen'}</button>
      </div>
      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}
      {vor && (
        <div className="mt-3 flex flex-col gap-2 text-xs">
          {vor.fazit && <p className="text-sm text-ink">{vor.fazit}</p>}
          <table className="w-full max-w-2xl">
            <tbody>
              {vor.aend.map(({ i, grund: g }) => (
                <tr key={i} className="border-b border-border/60">
                  <td className="py-1 pr-3 tnum whitespace-nowrap">{wd(sim[i].date)} {fmtDateShort(sim[i].date)}</td>
                  <td className="py-1 pr-3 tnum whitespace-nowrap"><span className="line-through text-ink-dim">{Math.round(loads[i])} {sports[i]} {zones[i]}</span> → {vor.loads[i]} {vor.sports[i]} {vor.zones[i]}</td>
                  <td className="py-1 text-ink-muted">{g}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="text-2xs text-ink-dim">Verstöße gegen die Monotonie-Regel: {vor.v0} → {vor.v1}</p>
          <div className="flex gap-2">
            <button
              className="btn btn-primary text-xs px-2 py-1"
              onClick={() => {
                onApply(vor.loads, vor.sports, vor.zones);
                // Grund als Ereignis-Marker merken (krank/Reise), damit die Lernmodelle ihn kennen
                const art = grund === 'infekt' ? 'krank' : grund === 'reise' ? 'reise' : null;
                if (art) saveExtra('ereignis', heute(), art === 'krank' ? 'krank – aus Umplanen' : 'Reise – aus Umplanen', { arten: [art], notiz: 'aus Umplanen bei Störungen' }).catch(() => undefined);
              }}
            >
              In Slider übernehmen
            </button>
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => setVor(null)}>Verwerfen</button>
          </div>
        </div>
      )}
    </div>
  );
}
