import { useEffect, useMemo, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { ALL_RANGE } from '../lib/windows';
import { buildDays, RECOVERY_FETCH_METRICS, VOTE_LABEL, type Vote } from '../screens/RecoveryPanel';
import { votumBilanz, votumTreffer, type VotumTag } from '../lib/votumTreffer';
import { fmtDateShort } from '../lib/format';

const f1 = (v: number | null) => (v == null ? '—' : `${v > 0 ? '+' : v < 0 ? '−' : ''}${Math.abs(v).toLocaleString('de-DE', { maximumFractionDigits: 2 })}`);
const VOTEN: Vote[] = ['REST', 'EASY', 'TRAIN', 'QUALITY'];
const KIRA_AB = '2026-05-05';

/** „Votum befolgt“ – rückwirkend über die ganze Historie, getrennt von der Treffsicherheit der Kira-Empfehlungen. */
export function VotumTrefferTile() {
  const [tage, setTage] = useState<VotumTag[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [fenster, setFenster] = useState<'alles' | 'vor' | 'ab'>('alles');

  useEffect(() => {
    let off = false;
    fetchChartData(ALL_RANGE, [...RECOVERY_FETCH_METRICS])
      .then((res: any) => {
        if (off) return;
        setTage(votumTreffer(buildDays(res.data || [])));
      })
      .catch((e) => !off && setErr(String(e?.message || e)));
    return () => {
      off = true;
    };
  }, []);

  const sicht = useMemo(() => {
    if (!tage) return [];
    if (fenster === 'vor') return tage.filter((t) => t.datum < KIRA_AB);
    if (fenster === 'ab') return tage.filter((t) => t.datum >= KIRA_AB);
    return tage;
  }, [tage, fenster]);
  const bil = useMemo(() => votumBilanz(sicht), [sicht]);
  const jeVotum = useMemo(
    () => VOTEN.map((v) => ({ v, ...votumBilanz(sicht.filter((t) => t.votum === v)) })).filter((x) => x.n > 0),
    [sicht],
  );
  const zeitraum = sicht.length ? `${fmtDateShort(sicht[sicht.length - 1].datum)}–${fmtDateShort(sicht[0].datum)}` : '';

  return (
    <div className="mt-6" data-testid="votum-treffer">
      <div className="label">Votum befolgt · ganze Historie</div>
      <p className="text-2xs text-ink-dim mt-1">
        Rückwirkend berechnet: das Tagesvotum nach den heutigen festen Regeln des Recovery Briefs (Stand Vortag, ohne KI) gegen den tatsächlichen
        Anspruch deiner Einheit. „Befolgt“ heißt: nicht härter als das Votum. Getrennt von Kiras Empfehlungen oben – das hier misst die Regel, nicht Kira.
      </p>
      <div className="mt-2 flex flex-wrap items-center gap-2 text-2xs" data-testid="votum-fenster">
        {([['alles', 'Alles'], ['vor', 'vor Mai 2026'], ['ab', 'ab Mai 2026']] as const).map(([k, l]) => (
          <button key={k} onClick={() => setFenster(k)} className={`btn btn-ghost px-2 py-0.5 text-2xs ${fenster === k ? 'border-accent text-ink' : 'text-ink-dim'}`}>
            {l}
          </button>
        ))}
        {zeitraum && <span className="text-ink-dim">{zeitraum} · {bil.n} Tage</span>}
      </div>
      {err && <div className="text-2xs text-orange-600 mt-2">Daten nicht lesbar: {err}</div>}
      {!tage && !err && <div className="text-2xs text-ink-dim mt-2">lädt …</div>}
      {tage && (
        <>
          <div className="mt-3 flex flex-wrap gap-6 items-baseline" data-testid="votum-bilanz">
            <div><span className="text-lg font-semibold">{bil.befolgtPct ?? '—'} %</span> <span className="text-2xs text-ink-dim">befolgt (n={bil.n})</span></div>
            <div><span className="text-lg font-semibold">{f1(bil.scoreBefolgt)}</span> <span className="text-2xs text-ink-dim">Reaktion befolgt ({bil.nBefolgt})</span></div>
            <div><span className="text-lg font-semibold">{f1(bil.scoreHaerter)}</span> <span className="text-2xs text-ink-dim">Reaktion härter als Votum ({bil.nHaerter})</span></div>
          </div>
          <p className="text-2xs text-ink-dim mt-1">Reaktion = gleiche Formel wie oben (Bereitschaft, HRV, Ruhepuls, Schlafscore am Morgen danach gegen das 7-Tage-Mittel; positiv = besser erholt).</p>
          <table className="w-full text-2xs mt-3">
            <thead className="text-ink-dim text-left">
              <tr className="border-b border-border"><th className="pb-1">Votum</th><th>Tage</th><th>befolgt</th><th>Reaktion befolgt</th><th>Reaktion härter</th></tr>
            </thead>
            <tbody>
              {jeVotum.map((x) => (
                <tr key={x.v} className="border-b border-border">
                  <td className="pt-1">{VOTE_LABEL[x.v]}</td>
                  <td className="pt-1">{x.n}</td>
                  <td className="pt-1">{x.befolgtPct} %</td>
                  <td className="pt-1">{f1(x.scoreBefolgt)}</td>
                  <td className="pt-1">{x.nHaerter ? f1(x.scoreHaerter) : '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="text-2xs text-ink-dim mt-2">
            Grenzen: Das Votum wird mit den heutigen Regeln nachgerechnet, nicht mit denen, die damals galten. Tage ohne Morgenwerte (Readiness) am Tag oder am Folgemorgen
            fehlen. Bei Votum Quality ist jede Einheit „befolgt“.
          </p>
        </>
      )}
    </div>
  );
}
