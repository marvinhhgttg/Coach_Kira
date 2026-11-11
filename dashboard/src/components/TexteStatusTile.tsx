import { useCallback, useEffect, useState } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Panel } from './UI';
import { useEngine } from '../lib/engine';

type Item = {
  key: string;
  gruppe: 'morgens' | 'aktivitaet' | 'abends';
  art: string;
  label: string;
  datum: string;
  erkannt: number;
  versuche: number;
  phase: 'wartet' | 'lokal' | 'cloud' | 'fertig';
  engine?: 'local' | 'cloud' | null;
  model?: string | null;
  fertigAt?: number;
  fehler?: string | null;
  grund?: string;
};
type Status = {
  ok: boolean;
  now: number;
  takt: { letzter: number | null; naechster: number; fehler: string | null; versuche: number; minuten: number };
  einstellung: { cloudErsatz: boolean };
  mac: { at: number | null; erreichbar: boolean; ollama: boolean | null; arbeit: string | null };
  items: Item[];
  auftraege?: Auftrag[];
};
type Auftrag = { id: string; key: string; art: string; datum: string; label: string; engine: 'local' | 'cloud'; phase: 'angefordert' | 'abgeholt' | 'cloud' | 'fertig' | 'fehler'; at: number; endAt?: number; ok?: boolean; result?: string; model?: string | null };

const uhr = (ms?: number | null) => (ms ? new Date(ms).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '–');
const GRUPPE: Record<Item['gruppe'], string> = { morgens: 'Morgens', aktivitaet: 'Nach Aktivität', abends: 'Abends' };

async function laden(pfad: string, body?: unknown): Promise<Status> {
  const r = await fetch(`${SERVER_BASE}${pfad}${body ? '' : `?cb=${Date.now()}`}`, body
    ? { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) }
    : { cache: 'no-store' });
  const t = await r.text();
  try {
    return JSON.parse(t);
  } catch {
    throw new Error(`Cockpit-Server antwortet nicht (HTTP ${r.status}).`);
  }
}

/** Nach dem Absenden sofort nachsehen lassen (zählt nicht als Versuch). */
export function texteNachAbsenden() {
  window.setTimeout(() => {
    laden('/api/texte/pruefen', {}).then(() => window.dispatchEvent(new Event('ck-texte'))).catch(() => {});
  }, 4000);
}

function Phase({ it, st }: { it: Item; st: Status }) {
  const chip = (cls: string, txt: string) => <span className={`chip border text-2xs whitespace-nowrap ${cls}`}>{txt}</span>;
  if (it.phase === 'fertig')
    return it.engine === 'cloud'
      ? chip('border-blue-400/40 text-blue-300', `fertig · Cloud${it.model ? ` · ${it.model}` : ''}`)
      : chip('border-green-400/40 text-green-300', `fertig · lokal${it.model ? ` · ${it.model}` : ''}`);
  if (it.phase === 'lokal') return chip('border-green-400/40 text-green-300', 'lokal in Arbeit');
  if (it.phase === 'cloud') return chip('border-blue-400/40 text-blue-300', 'Cloud-Ersatz läuft');
  const voll = it.versuche >= st.takt.versuche;
  if (voll && !st.einstellung.cloudErsatz) return chip('border-orange-400/40 text-orange-300', 'lokal gescheitert · Cloud-Ersatz aus');
  if (voll) return chip('border-orange-400/40 text-orange-300', it.fehler ? 'Cloud-Ersatz fehlgeschlagen · neuer Versuch' : 'Cloud-Ersatz startet');
  return chip('border-border text-ink-muted', `wartet auf Mac · Prüfung ${it.versuche}/${st.takt.versuche}`);
}

const heuteIso = () => new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Berlin' });
const plusTag = (iso: string, n: number) => { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const kurz = (iso: string) => `${iso.slice(8, 10)}.${iso.slice(5, 7)}.`;

function AuftragZeile({ a }: { a: Auftrag }) {
  const cloud = a.engine === 'cloud';
  const ton = a.phase === 'fehler' ? 'text-orange-300' : cloud ? 'text-blue-300' : 'text-green-300';
  const txt =
    a.phase === 'angefordert' ? 'wartet, bis der Mac den Auftrag abholt (max. 1 min)'
    : a.phase === 'abgeholt' ? 'Mac rechnet …'
    : a.phase === 'cloud' ? 'Cloud rechnet …'
    : a.phase === 'fertig' ? `fertig ${uhr(a.endAt)}${a.model ? ` · ${a.model}` : ''}`
    : `Fehler: ${a.result || 'unbekannt'}`;
  return (
    <li className="flex flex-wrap items-center justify-between gap-2 py-1" data-testid="texte-auftrag">
      <span className="text-ink-dim">{uhr(a.at)} · <span className="text-ink">{a.label}{a.art === 'aktivitaet' || a.art === 'ernaehrung' ? ` ${kurz(a.datum)}` : ''}</span> · {cloud ? 'Cloud' : 'lokal'}</span>
      <span className={ton}>{txt}</span>
    </li>
  );
}

function ManuellStarten({ st, onUpdate }: { st: Status; onUpdate: () => void }) {
  const [engine] = useEngine();
  const cloud = engine === 'pplx';
  const [msg, setMsg] = useState<string | null>(null);
  const tag = heuteIso();
  const aktTage = Array.from(new Set(st.items.filter((i) => i.art === 'aktivitaet').map((i) => i.datum))).sort().reverse();
  const laeuft = (key: string) => (st.auftraege || []).some((a) => a.key === key && ['angefordert', 'abgeholt', 'cloud'].includes(a.phase));
  const laeufe: { art: string; datum: string; key: string; label: string }[] = [
    { art: 'status', datum: tag, key: `morgen:${tag}:status`, label: 'KI-Analyse' },
    { art: 'plan', datum: tag, key: `morgen:${tag}:plan`, label: 'Plan-Briefing' },
    { art: 'consolidated', datum: tag, key: `morgen:${tag}:consolidated`, label: 'Konsolidierte Analyse' },
    ...aktTage.map((d) => ({ art: 'aktivitaet', datum: d, key: `aktivitaet:${d}`, label: `Aktivitätskommentar ${kurz(d)}` })),
    { art: 'ernaehrung', datum: plusTag(tag, 1), key: `ernaehrung:${plusTag(tag, 1)}`, label: `Ernährungsplan ${kurz(plusTag(tag, 1))}` },
  ];
  const starten = async (l: (typeof laeufe)[number]) => {
    setMsg(null);
    try {
      const r = await fetch(`${SERVER_BASE}/api/texte/starten`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ art: l.art, datum: l.datum, engine: cloud ? 'cloud' : 'local' }),
      });
      const j = await r.json().catch(() => ({ ok: false, error: `HTTP ${r.status}` }));
      if (!j.ok) setMsg(j.error || 'Start fehlgeschlagen');
      else setMsg(cloud ? `${l.label}: Cloud rechnet.` : `${l.label}: Auftrag an den Mac – er holt ihn innerhalb einer Minute ab.`);
    } catch (e: any) {
      setMsg(String(e.message || e));
    }
    onUpdate();
  };
  const auftr = (st.auftraege || []).slice().reverse().slice(0, 6);
  return (
    <div className="pt-3 border-t border-border/60 space-y-2" data-testid="texte-manuell">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <span className="text-ink font-medium">Manuell starten</span>
        <span className="text-2xs text-ink-dim">
          läuft mit der Auswahl oben:{' '}
          <span className={cloud ? 'text-blue-300' : 'text-green-300'}>{cloud ? 'Cloud' : 'Lokal (Mac)'}</span>
          {!cloud && !st.mac.erreichbar ? ' · Mac gerade nicht erreichbar – der Auftrag wartet, bis er sich meldet' : ''}
        </span>
      </div>
      <div className="flex flex-wrap gap-2">
        {laeufe.map((l) => (
          <button
            key={l.key}
            type="button"
            className={`btn text-xs px-2.5 py-1 border ${cloud ? 'border-blue-400/40 hover:bg-blue-500/10' : 'border-green-400/40 hover:bg-green-500/10'}`}
            onClick={() => starten(l)}
            disabled={laeuft(l.key)}
            title={`${l.label} jetzt ${cloud ? 'in der Cloud' : 'lokal auf dem Mac'} erzeugen – ein vorhandener Text wird ersetzt`}
            data-testid="texte-start"
          >
            {laeuft(l.key) ? `${l.label} läuft …` : l.label}
          </button>
        ))}
      </div>
      <div className="text-2xs text-ink-dim">Ein vorhandener Text wird durch den neuen ersetzt.</div>
      {msg && <div className="text-2xs text-ink-muted">{msg}</div>}
      {auftr.length > 0 && <ul className="divide-y divide-border/40 text-2xs">{auftr.map((a) => <AuftragZeile key={a.id} a={a} />)}</ul>}
    </div>
  );
}

export function TexteStatusTile({ compact = false }: { compact?: boolean }) {
  const [st, setSt] = useState<Status | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const neu = useCallback(() => {
    laden('/api/texte/status').then((s) => { setSt(s); setErr(null); }).catch((e) => setErr(String(e.message || e)));
  }, []);
  useEffect(() => {
    neu();
    const id = window.setInterval(neu, 10000);
    window.addEventListener('ck-texte', neu);
    return () => { window.clearInterval(id); window.removeEventListener('ck-texte', neu); };
  }, [neu]);

  const umschalten = async () => {
    if (!st) return;
    setBusy(true);
    try { setSt(await laden('/api/texte/einstellung', { cloudErsatz: !st.einstellung.cloudErsatz })); } catch (e: any) { setErr(String(e.message || e)); }
    setBusy(false);
  };
  const jetzt = async () => {
    setBusy(true);
    try { setSt(await laden('/api/texte/pruefen', {})); } catch (e: any) { setErr(String(e.message || e)); }
    setBusy(false);
  };

  const items = st?.items || [];
  const offen = items.filter((i) => i.phase !== 'fertig').length;
  const macTon = !st ? 'bg-ink-dim' : st.mac.erreichbar && st.mac.ollama !== false ? 'bg-green-400' : 'bg-orange-400';
  const macText = !st
    ? 'wird geprüft …'
    : st.mac.erreichbar
      ? st.mac.ollama === false ? `Mac erreichbar, Ollama nicht bereit · ${uhr(st.mac.at)}` : st.mac.arbeit ? `Mac rechnet gerade · ${uhr(st.mac.at)}` : `Mac bereit · zuletzt ${uhr(st.mac.at)}`
      : st.mac.at ? `Mac nicht erreichbar · zuletzt ${uhr(st.mac.at)}` : 'Mac nicht erreichbar · noch kein Signal';

  return (
    <Panel
      title="Kira-Texte nach dem Absenden"
      right={
        <div className="flex flex-wrap items-center gap-2 text-2xs text-ink-dim">
          <span className="inline-flex items-center gap-1.5" data-testid="texte-mac"><span className={`inline-block w-2 h-2 rounded-full ${macTon}`} aria-hidden />{macText}</span>
          <button type="button" className="btn btn-ghost text-xs px-2 py-1" onClick={jetzt} disabled={busy}>Jetzt prüfen</button>
        </div>
      }
    >
      <div className="space-y-3 text-xs" data-testid="texte-status">
        {err && <div className="text-orange-300">{err}</div>}
        {st && items.length === 0 && <div className="text-ink-muted">Heute noch nichts offen. Sobald Morgenwerte, eine Aktivität oder Abendwerte im Sheet stehen, erscheinen hier die zugehörigen Texte.</div>}
        {items.length > 0 && (
          <ul className="divide-y divide-border/60">
            {items.map((it) => (
              <li key={it.key} className="py-1.5 flex flex-wrap items-center justify-between gap-2" data-testid="texte-item">
                <span className="min-w-0">
                  <span className="text-ink-dim">{GRUPPE[it.gruppe]} · </span>
                  <span className="text-ink">{it.label}</span>
                  {!compact && (
                    <span className="text-ink-dim"> · erkannt {uhr(it.erkannt)}{it.fertigAt && it.phase === 'fertig' ? ` · fertig ${uhr(it.fertigAt)}` : ''}</span>
                  )}
                  {it.phase === 'wartet' && it.grund && !compact && <span className="text-ink-dim"> · {it.grund}</span>}
                  {it.fehler && it.phase !== 'fertig' && <span className="text-orange-300"> · {it.fehler}</span>}
                </span>
                {st && <Phase it={it} st={st} />}
              </li>
            ))}
          </ul>
        )}
        {st && !compact && <ManuellStarten st={st} onUpdate={neu} />}
        {st && (
          <div className="flex flex-wrap items-center justify-between gap-2 text-2xs text-ink-dim">
            <span>
              {offen ? `${offen} offen · ` : ''}Prüfung alle {st.takt.minuten} min · nächste ca. {uhr(st.takt.naechster)}
              {st.takt.fehler ? ` · Prüfung gestört: ${st.takt.fehler}` : ''}
            </span>
            <label className="inline-flex items-center gap-1.5 cursor-pointer">
              <input type="checkbox" checked={st.einstellung.cloudErsatz} onChange={umschalten} disabled={busy} data-testid="texte-cloud-ersatz" />
              Cloud-Ersatz nach {st.takt.versuche} × {st.takt.minuten} min ohne lokalen Text
            </label>
          </div>
        )}
      </div>
    </Panel>
  );
}
