import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { IST_V2 } from '../lib/v2guard';
import { Panel } from './UI';

/**
 * 24.11.2026: Automatik – Heute-Banner (vorläufiges Votum, Plan-Vorschlag) und Technik-Kachel
 * (Schalter Formular/Hinweise/Plancheck/Workout, Protokoll der automatischen Einträge, Mac-Puls).
 */
type Aenderung = { d: string; sport: string; zone: string; alt: number; neu: number };
type Vorschlag = { id: string; grenze: number; maxAlt: number; tagAlt: string; maxNeu: number; erreicht: boolean; aenderungen: Aenderung[] };
type LogEintrag = { at: number; art: string; typ?: string; datum?: string; ok: boolean; uebersprungen?: string; fehler?: string; werte?: Record<string, string>; name?: string; schluessel?: string[]; verworfen?: boolean };
type Status = {
  ok: boolean; error?: string;
  einstellung: { formular: boolean; hinweise: boolean; plancheck: boolean; workout: boolean; erfolg: boolean };
  log: LogEintrag[]; puls: { at: number } | null;
  vorschlag: Vorschlag | null; vorschlagStatus: { id?: string; status?: string };
  hinweise: { id: string; titel: string; text: string }[];
  votum: { vote: string | null; label: string; text: string; gruende: string[]; vorlaeufig?: boolean } | null;
  schlaf: { lichtAus: string; bedarfH: number; aufstehen: string } | null;
};

const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
const tag = (d: string) => `${WT[new Date(d + 'T12:00:00Z').getUTCDay()]} ${d.slice(8, 10)}.${d.slice(5, 7)}.`;
const komma = (v: number, n = 2) => v.toFixed(n).replace('.', ',');
const zeit = (ms: number) => new Date(ms).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });

async function j(r: Response) {
  const t = await r.text();
  try { return JSON.parse(t); } catch { return { ok: false, error: `HTTP ${r.status}` }; }
}

function useAutoStatus(ms = 60000) {
  const [s, setS] = useState<Status | null>(null);
  const laden = useCallback(async () => {
    if (IST_V2) return;
    try {
      const r = await j(await fetch(`${SERVER_BASE}/api/auto/status?cb=${Date.now()}`, { cache: 'no-store' }));
      if (r.ok) setS(r);
    } catch {}
  }, []);
  useEffect(() => {
    laden();
    const id = window.setInterval(laden, ms);
    return () => window.clearInterval(id);
  }, [laden, ms]);
  return [s, laden] as const;
}

/** Heute: vorläufiges Votum (Garmin live, Sheet folgt) und Plan-Vorschlag bei ACWR-Prognose über dem Korridor. */
export function AutoBanner() {
  const [s, laden] = useAutoStatus();
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<string | null>(null);
  if (!s) return null;
  const pv = s.vorschlag && s.vorschlagStatus?.id !== s.vorschlag.id ? s.vorschlag : null;
  const vorl = s.votum?.vorlaeufig ? s.votum : null;
  if (!pv && !vorl && !msg) return null;
  const aktion = async (a: 'annehmen' | 'verwerfen') => {
    if (!pv) return;
    setBusy(true);
    setMsg(null);
    const r = await j(await fetch(`${SERVER_BASE}/api/auto/vorschlag`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ aktion: a, id: pv.id }) }));
    setBusy(false);
    setMsg(r.ok ? r.text : r.error || 'Fehler');
    laden();
  };
  return (
    <div className="mb-4 space-y-3">
      {vorl && (
        <div className="panel-raised border-l-4 border-sky-500 p-4">
          <p className="text-base font-semibold text-ink">Vorläufiges Votum: {vorl.label}</p>
          <p className="mt-1 text-sm text-ink-muted">Aus den Garmin-Livewerten von heute Morgen. Die Morgenwerte werden automatisch ins Sheet eingetragen, dann gilt das endgültige Votum. {vorl.gruende?.[0] || ''}</p>
        </div>
      )}
      {pv && (
        <div className="panel-raised border-l-4 border-orange-500 p-4">
          <p className="text-base font-semibold text-ink">Plan-Vorschlag: ACWR-Prognose {komma(pv.maxAlt)} am {tag(pv.tagAlt)}</p>
          <p className="mt-1 text-sm text-ink-muted">Der Korridor endet bei {komma(pv.grenze)}. Kira schlägt kleinere Lasten vor{pv.erreicht ? '' : ' (Grenze nicht ganz erreicht, gesperrte Tage bleiben)'}:</p>
          <ul className="mt-2 space-y-1 text-sm">
            {pv.aenderungen.map((x) => (
              <li key={x.d} className="tabular-nums text-ink">{tag(x.d)} · {x.sport} {x.zone} · {x.alt} → <b>{x.neu}</b> ESS</li>
            ))}
          </ul>
          <p className="mt-2 text-sm text-ink-muted">Danach höchstens {komma(pv.maxNeu)}. Bestätigen geht auch im Uhr-Menü.</p>
          <div className="mt-3 flex gap-2">
            <button className="btn btn-primary" disabled={busy} onClick={() => aktion('annehmen')}>{busy ? 'Speichere…' : 'Übernehmen'}</button>
            <button className="btn" disabled={busy} onClick={() => aktion('verwerfen')}>Verwerfen</button>
          </div>
        </div>
      )}
      {msg && <p className="text-sm text-ink-muted">{msg}</p>}
    </div>
  );
}

const SCHALTER: { k: keyof Status['einstellung']; label: string; unter: string }[] = [
  { k: 'formular', label: 'Formular automatisch', unter: 'Morgens, nach jeder Einheit und abends um 23:30 vom Mac eintragen – manuelle Einträge haben Vorrang' },
  { k: 'hinweise', label: 'Hinweise an Uhr und iPhone', unter: 'Morgenvotum, Befinden, Licht aus, ACWR/Ruhepuls, Plan-Vorschlag' },
  { k: 'plancheck', label: 'Plancheck', unter: 'ACWR-Prognose über 1,50 → Vorschlag mit kleineren Lasten' },
  { k: 'workout', label: 'Workout für morgen', unter: 'Abends ab 20 Uhr in den Garmin-Kalender, wenn noch keins da ist' },
  { k: 'erfolg', label: 'Erfolgsmeldungen', unter: 'Uhr und iPhone melden eingetragene Werte, Workout im Kalender, Plan angepasst' },
];

function zeile(e: LogEintrag) {
  if (e.art === 'eintrag') {
    if (e.uebersprungen) return `${e.typ} ${e.datum}: übersprungen – ${e.uebersprungen}`;
    return `${e.typ} ${e.datum}: ${e.ok ? 'eingetragen' : 'Fehler'}${e.werte ? ' (' + Object.entries(e.werte).map(([k, v]) => `${k} ${v}`).join(', ') + ')' : ''}${e.fehler ? ' – ' + e.fehler : ''}`;
  }
  if (e.art === 'workout') return `Workout ${e.datum || ''}: ${e.ok ? e.name : 'Fehler – ' + e.fehler}`;
  if (e.art === 'plan') return e.verworfen ? 'Plan-Vorschlag verworfen' : `Plan-Vorschlag ${e.ok ? 'übernommen' : 'Fehler – ' + e.fehler}`;
  if (e.art === 'wiederherstellung') return `Nach Neustart wiederhergestellt: ${(e.schluessel || []).join(', ')}`;
  return `${e.art}: ${e.ok ? 'ok' : e.fehler}`;
}

/** Technik: Schalter, Mac-Puls und Protokoll. */
export function AutomatikTile() {
  const [s, laden] = useAutoStatus(30000);
  const [err, setErr] = useState<string | null>(null);
  if (IST_V2) return null;
  const setzen = async (k: string, v: boolean) => {
    setErr(null);
    const r = await j(await fetch(`${SERVER_BASE}/api/auto/einstellung`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ [k]: v }) }));
    if (!r.ok) setErr(r.error || 'Fehler');
    laden();
  };
  const pulsAlt = s?.puls ? Math.round((Date.now() - s.puls.at) / 60000) : null;
  return (
    <Panel title="Automatik" right={<span className={`text-sm ${pulsAlt != null && pulsAlt <= 20 ? 'text-green-400' : 'text-orange-300'}`}>{pulsAlt == null ? 'Mac noch nicht gemeldet' : `Mac vor ${pulsAlt} min`}</span>}>
      {!s ? <p className="text-sm text-ink-muted">Lädt…</p> : (
        <div className="space-y-4">
          <div className="grid gap-2 sm:grid-cols-2">
            {SCHALTER.map((x) => (
              <label key={x.k} className="flex cursor-pointer items-start gap-3 rounded-lg border border-border p-3">
                <input type="checkbox" className="mt-1" checked={!!s.einstellung[x.k]} onChange={(e) => setzen(x.k, e.target.checked)} />
                <span><span className="block text-sm font-semibold text-ink">{x.label}</span><span className="block text-sm text-ink-muted">{x.unter}</span></span>
              </label>
            ))}
          </div>
          {err && <p className="text-sm text-orange-300">{err}</p>}
          {s.schlaf && <p className="text-sm text-ink-muted">Licht aus heute: <b className="text-ink">{s.schlaf.lichtAus}</b> (Bedarf {komma(s.schlaf.bedarfH, 1)} h bis {s.schlaf.aufstehen})</p>}
          <div>
            <p className="mb-1 text-sm font-semibold text-ink">Fällige Hinweise</p>
            {s.hinweise.length ? s.hinweise.map((h) => <p key={h.id} className="text-sm text-ink-muted"><b className="text-ink">{h.titel}</b> – {h.text}</p>) : <p className="text-sm text-ink-muted">keine</p>}
          </div>
          <div>
            <p className="mb-1 text-sm font-semibold text-ink">Protokoll</p>
            {s.log.length ? (
              <ul className="max-h-64 space-y-1 overflow-auto text-sm">
                {s.log.map((e, i) => <li key={i} className={e.ok ? 'text-ink-muted' : 'text-orange-300'}><span className="tabular-nums">{zeit(e.at)}</span> · {zeile(e)}</li>)}
              </ul>
            ) : <p className="text-sm text-ink-muted">noch keine automatischen Schritte</p>}
          </div>
        </div>
      )}
    </Panel>
  );
}
