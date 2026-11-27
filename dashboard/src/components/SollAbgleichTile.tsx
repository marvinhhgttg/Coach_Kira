import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Panel } from './UI';

type Rahmen = { min: number; max: number; halte: number; acwrGrenze: number | null; gedeckelt: boolean };
type Hist = { datum: string; vote: string | null; min: number; max: number; plan: number | null; planLage: string | null; ist: number | null; istLage: string | null; gesperrt?: boolean };
type Antwort = {
  ok: boolean; error?: string; datum?: string; vote?: string | null;
  soll?: { rahmen: Rahmen | null; plan: { load: number; sport: string; zone: string; gesperrt: boolean; quelle: string } | null; planLage: string | null; ist: number | null; istLage: string | null } | null;
  rueckblick?: { tage: number; imRahmen: number; darueber: number; darunter: number } | null;
  historie?: Hist[];
  quote?: { tage: number; planImRahmen: number; planDarueber: number; planDarunter: number };
};
const FARBE: Record<string, string> = { 'im Rahmen': 'var(--ck-green, #16a34a)', 'darüber': 'var(--ck-orange, #ea580c)', 'darunter': '#2563eb' };
const VOTE_DE: Record<string, string> = { QUALITY: 'Quality', TRAIN: 'Train', EASY: 'Easy', REST: 'Rest' };

/** Etappe 2b (nur v2): Soll-Rahmen aus dem Kern gegen den Plan im Sheet, nur lesend. */
export function SollAbgleichTile() {
  const [v2, setV2] = useState(false);
  const [d, setD] = useState<Antwort | null>(null);
  const [laedt, setLaedt] = useState(false);
  useEffect(() => {
    fetch(`${SERVER_BASE}/api/instanz`).then((r) => r.json()).then((j) => setV2(j?.instanz === 'v2')).catch(() => setV2(false));
  }, []);
  const laden = useCallback(async (neu = false) => {
    setLaedt(true);
    try { setD(await (await fetch(`${SERVER_BASE}/api/rechenkern/soll${neu ? '?neu=1' : ''}`, { cache: 'no-store' })).json()); }
    catch (e) { setD({ ok: false, error: String(e) }); }
    finally { setLaedt(false); }
  }, []);
  useEffect(() => { if (v2) laden(); }, [v2, laden]);
  if (!v2) return null;
  const s = d?.soll;
  const r = s?.rahmen;
  const q = d?.quote;
  const rb = d?.rueckblick;
  return (
    <Panel title="Soll-Abgleich · Kern-Rahmen gegen Plan im Sheet (Etappe 2b)">
      <div className="flex items-center justify-between gap-3 mb-3">
        <p className="text-sm opacity-80">Der Kern rechnet den Soll-Rahmen für heute (Votum × Fitness halten, gedeckelt bei ACWR 1,3). Der Plan im Sheet (coachE_ESS_day) bleibt dein Planwert und wird nur dagegen geprüft.</p>
        <button type="button" className="btn btn-ghost text-xs px-2 py-1 shrink-0" onClick={() => laden(true)} disabled={laedt}>{laedt ? 'lädt …' : 'Neu prüfen'}</button>
      </div>
      {d && !d.ok && <div className="text-sm text-orange-300">{d.error || 'Soll-Abgleich nicht verfügbar.'}</div>}
      {d?.ok && (
        <>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs mb-3">
            <div className="rounded border border-current/10 p-2">
              <div className="opacity-60">Rahmen aus dem Kern</div>
              <div className="text-base font-semibold tabular-nums">{r ? `${r.min}–${r.max} ESS` : '–'}</div>
              <div className="opacity-70">Votum {d.vote ? VOTE_DE[d.vote] || d.vote : '–'}{r ? ` · halten ≈ ${r.halte}` : ''}{r?.gedeckelt ? ' · durch ACWR gedeckelt' : ''}</div>
            </div>
            <div className="rounded border border-current/10 p-2">
              <div className="opacity-60">Plan im Sheet</div>
              <div className="text-base font-semibold tabular-nums">{s?.plan ? `${s.plan.load} ESS` : '–'}</div>
              <div className="opacity-70">{s?.plan ? `${s.plan.sport || '–'}${s.plan.zone ? ' · ' + s.plan.zone : ''}${s.plan.gesperrt ? ' · gesperrt' : ''}` : 'kein Plan für heute'}</div>
            </div>
            <div className="rounded border border-current/10 p-2">
              <div className="opacity-60">Heute</div>
              <div className="text-base font-semibold" style={{ color: s?.planLage ? FARBE[s.planLage] : undefined }}>{s?.planLage ? `Plan ${s.planLage}` : 'noch offen'}</div>
              <div className="opacity-70">{s?.ist != null ? `Ist ${s.ist} ESS · ${s.istLage ?? '–'}` : 'Einheit noch nicht erfasst'}</div>
            </div>
          </div>
          <div className="rounded border border-current/10 p-2 text-xs">
            <div className="flex items-baseline gap-2">
              <span className="text-lg font-semibold tabular-nums">{q?.planImRahmen ?? 0}/{q?.tage ?? 0}</span>
              <span className="opacity-80">Tage, an denen der Plan im Kern-Rahmen lag (täglich ab 12 Uhr mitgeschrieben)</span>
            </div>
            <div className="flex flex-wrap gap-1 mt-2" aria-label="Täglicher Soll-Abgleich">
              {(d.historie || []).map((h) => (
                <span key={h.datum} title={`${h.datum}: Rahmen ${h.min}–${h.max}, Plan ${h.plan ?? '–'} → ${h.planLage ?? 'offen'}${h.ist != null ? `, Ist ${h.ist}` : ''}`} className="inline-block w-3 h-3 rounded-sm" style={{ background: h.planLage ? FARBE[h.planLage] : '#9ca3af' }} />
              ))}
              {!d.historie?.length && <span className="opacity-60">Erster Eintrag folgt heute nach 12 Uhr bzw. beim nächsten Prüflauf.</span>}
            </div>
            {rb && <div className="opacity-60 mt-1">Rückblick 30 Tage, Ist-Last gegen Rahmen: {rb.imRahmen} im Rahmen, {rb.darueber} darüber, {rb.darunter} darunter (von {rb.tage}).</div>}
          </div>
        </>
      )}
    </Panel>
  );
}
