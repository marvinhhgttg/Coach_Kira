import { useEffect, useMemo, useState } from 'react';
import type { ExtraItem } from '../lib/extras';
import { fetchExtras } from '../lib/extras';
import { fmtDateShort, fmtNum } from '../lib/format';
import { VotumTrefferTile } from './VotumTrefferTile';

// --- Section A: Treffsicherheit types ---
interface TrefferTag {
  datum: string;
  empfohlen: number | null;
  ist: number | null;
  abweichungPct: number | null;
  gefolgt: boolean | null;
  score: number | null;
  quelle: string;
}

function parseTrefferTag(item: ExtraItem): TrefferTag {
  const d = item.data || {};
  return {
    datum: String(d.datum ?? item.datum),
    empfohlen: typeof d.empfohlen === 'number' ? d.empfohlen : null,
    ist: typeof d.ist === 'number' ? d.ist : null,
    abweichungPct: typeof d.abweichung_pct === 'number' ? d.abweichung_pct : null,
    gefolgt: d.gefolgt !== undefined ? !!d.gefolgt : null,
    score: typeof d.score === 'number' ? d.score : null,
    quelle: String(d.quelle ?? item.quelle),
  };
}

// --- Section B: Selbsttest types ---
interface SelbsttestEintrag {
  datum: string;
  gespeichert: string;
  quelle: string;
  tests: { name: string; ok: boolean; detail: string; sekunden: number }[];
  text?: string;
}

function parseSelbsttest(item: ExtraItem): SelbsttestEintrag {
  const d = item.data || {};
  return {
    datum: String(d.datum ?? item.datum),
    gespeichert: item.gespeichert,
    quelle: String(d.quelle ?? item.quelle),
    tests: (d.tests || []).map((t: unknown) => {
      const o = t as Record<string, unknown>;
      return {
        name: String(o.name || ''),
        ok: !!o.ok,
        detail: String(o.detail || ''),
        sekunden: typeof o.sekunden === 'number' ? o.sekunden : 0,
      };
    }),
    text: item.text || undefined,
  };
}

// --- Section A: weekly summary type ---
interface TrefferWoche {
  datum: string;
  quelle: string;
  text: string;
}

function parseTrefferWoche(item: ExtraItem): TrefferWoche {
  const d = item.data || {};
  return {
    datum: String(d.datum ?? item.datum),
    quelle: String(d.quelle ?? item.quelle),
    text: item.text,
  };
}

export function KiraExtrasPanel() {
  const [trefferTagItems, setTrefferTagItems] = useState<ExtraItem[]>([]);
  const [trefferWocheItems, setTrefferWocheItems] = useState<ExtraItem[]>([]);
  const [selbsttestItems, setSelbsttestItems] = useState<ExtraItem[]>([]);
  const [errTag, setErrTag] = useState<string | null>(null);
  const [fenster, setFenster] = useState<'90' | '180' | 'alles'>('alles');
  const [errWoche, setErrWoche] = useState<string | null>(null);
  const [errTest, setErrTest] = useState<string | null>(null);

  // Load all extra types in useEffect
  useEffect(() => {
    let cancelled = false;

    async function loadAll() {
      // Section A: treffer_tag (14 entries)
      try {
        const items = await fetchExtras(['treffer_tag'], 1000);
        if (!cancelled) setTrefferTagItems(items);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (msg.includes('JSON')) setErrTag('Patch „KI-Extras" noch nicht aktiv');
        else if (!cancelled) setErrTag('Treffer-Tage konnten nicht geladen werden');
      }

      // Section A: treffer_woche (latest)
      try {
        const items = await fetchExtras(['treffer_woche'], 4);
        if (!cancelled) setTrefferWocheItems(items);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (msg.includes('JSON')) setErrWoche('Patch „KI-Extras" noch nicht aktiv');
        else if (!cancelled) setErrWoche('Wochenbilanz konnte nicht geladen werden');
      }

      // Section B: selbsttest (14 entries)
      try {
        const items = await fetchExtras(['selbsttest'], 14);
        if (!cancelled) setSelbsttestItems(items);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (msg.includes('JSON')) setErrTest('Patch „KI-Extras" noch nicht aktiv');
        else if (!cancelled) setErrTest('Selbsttest konnte nicht geladen werden');
      }
    }

    loadAll();
    return () => { cancelled = true; };
  }, []);

  // Section A: parse data
  const alleTrefferTage = useMemo(() => {
    // je Datum nur den neuesten Eintrag, absteigend sortiert
    const m = new Map<string, ReturnType<typeof parseTrefferTag>>();
    for (const t of trefferTagItems.map(parseTrefferTag)) if (t.datum && !m.has(t.datum)) m.set(t.datum, t);
    return [...m.values()].sort((a, b) => (a.datum < b.datum ? 1 : -1));
  }, [trefferTagItems]);
  const trefferTage = useMemo(() => {
    if (fenster === 'alles') return alleTrefferTage;
    const grenze = new Date(Date.now() - Number(fenster) * 86400000).toISOString().slice(0, 10);
    return alleTrefferTage.filter((t) => t.datum >= grenze);
  }, [alleTrefferTage, fenster]);
  const zeitraum = trefferTage.length ? `${fmtDateShort(trefferTage[trefferTage.length - 1].datum)}–${fmtDateShort(trefferTage[0].datum)}` : '';
  const nachgetragen = trefferTage.filter((t) => t.quelle === 'plan-briefing').length;
  const wocheEintraege = useMemo(() => trefferWocheItems.map(parseTrefferWoche), [trefferWocheItems]);
  const selbsttestEintraege = useMemo(() => selbsttestItems.map(parseSelbsttest), [selbsttestItems]);

  // Section A: metrics
  const metrics = useMemo(() => {
    if (!trefferTage.length) return null;
    const gefolgtList = trefferTage.filter((t) => t.gefolgt === true);
    const nichtGefolgtList = trefferTage.filter((t) => t.gefolgt === false);

    const gefolgtAnzahl = gefolgtList.length;
    const folgeQuote = trefferTage.length > 0 ? Math.round((gefolgtAnzahl / trefferTage.length) * 100) : 0;

    const scoreGefolgt = gefolgtList.length > 0
      ? gefolgtList.reduce((s: number, t) => s + (t.score ?? 0), 0) / gefolgtList.length
      : null;
    const scoreNichtGefolgt = nichtGefolgtList.length > 0
      ? nichtGefolgtList.reduce((s: number, t) => s + (t.score ?? 0), 0) / nichtGefolgtList.length
      : null;

    return { folgeQuote, scoreGefolgt, scoreNichtGefolgt, gesamt: trefferTage.length };
  }, [trefferTage]);

  // Section A: latest weekly summary
  const neuesteWoche = wocheEintraege.length > 0 ? wocheEintraege[0] : null;

  // Section B: latest self-test entry
  const neuesterTest = selbsttestEintraege.length > 0 ? selbsttestEintraege[0] : null;

  // Section B: last 14 days history
  const testVerlauf = useMemo(() => {
    return selbsttestEintraege.slice(0, 14).map((e) => ({
      datum: e.datum,
      okCount: e.tests.filter((t) => t.ok).length,
    }));
  }, [selbsttestEintraege]);

  return (
    <div className="panel-raised p-4">
      {/* Section A: Kira-Treffsicherheit */}
      <div className="label">Coach Kira · Treffsicherheit</div>
      <p className="text-2xs text-ink-dim mt-1">Wie hat dein Körper reagiert, wenn du Kiras Empfehlung gefolgt bist – und wenn nicht?</p>
      <div className="mt-2 flex flex-wrap items-center gap-2 text-2xs" data-testid="treffer-fenster">
        {([['90', '90 Tage'], ['180', '180 Tage'], ['alles', 'Alles']] as const).map(([k, l]) => (
          <button key={k} onClick={() => setFenster(k)} className={`btn btn-ghost px-2 py-0.5 text-2xs ${fenster === k ? 'border-accent text-ink' : 'text-ink-dim'}`}>{l}</button>
        ))}
        {zeitraum && <span className="text-ink-dim">{zeitraum} · {trefferTage.length} Tage{nachgetragen ? ` · davon ${nachgetragen} aus Kiras Plan-Briefings nachgetragen` : ''}</span>}
      </div>
      <p className="text-2xs text-ink-dim mt-1">Ausgewertet werden alle Tage mit gespeicherter Kira-Empfehlung (seit Mai 2026). Für die Zeit davor gibt es keine gespeicherten Empfehlungen – in der Timeline ist die Soll-Last vergangener Tage durch die Ist-Last ersetzt.</p>

      {(errTag || errWoche) && (
        <p className="mt-2 text-xs text-orange-300">
          {[errTag, errWoche].filter(Boolean).join(' · ')}
        </p>
      )}

      {metrics && (
        <div className="mt-3 rounded border border-border p-3">
          <div className="flex flex-wrap gap-4 text-sm">
            <div>
              <span className="text-lg tnum">{metrics.folgeQuote} %</span>
              <span className="ml-1 text-2xs text-ink-dim">gefolgt (n={metrics.gesamt})</span>
            </div>
            <div>
              <span className="text-lg tnum">
                {metrics.scoreGefolgt !== null ? (metrics.scoreGefolgt > 0 ? '+' : '') + fmtNum(metrics.scoreGefolgt, 1) : '—'}
              </span>
              <span className="ml-1 text-2xs text-ink-dim">Reaktion gefolgt</span>
            </div>
            <div>
              <span className={`text-lg tnum ${metrics.scoreNichtGefolgt !== null && metrics.scoreNichtGefolgt < 0 ? 'text-orange-300' : ''}`}>
                {metrics.scoreNichtGefolgt !== null ? (metrics.scoreNichtGefolgt > 0 ? '+' : '') + fmtNum(metrics.scoreNichtGefolgt, 1) : '—'}
              </span>
              <span className="ml-1 text-2xs text-ink-dim">Reaktion nicht gefolgt</span>
            </div>
          </div>
          <p className="mt-2 text-2xs text-ink-dim">
            Reaktion = Veränderung von Bereitschaft, HRV, Ruhepuls und Schlaf am Morgen danach gegenüber dem 7-Tage-Mittel (positiv = besser erholt).
          </p>

          {/* Weekly summary */}
          {neuesteWoche && (
            <div className="mt-3 rounded border border-border p-2">
              <div className="text-xs font-semibold text-ink mb-1">Wochenbilanz · {fmtDateShort(neuesteWoche.datum)}</div>
              <pre className="whitespace-pre-line text-xs text-ink">{neuesteWoche.text}</pre>
              <div className="text-2xs text-ink-dim mt-1">Quelle: {neuesteWoche.quelle}</div>
            </div>
          )}

          {/* Treffer-Tage Tabelle */}
          {trefferTage.length > 0 && (
            <div className="mt-3 overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead>
                  <tr className="border-b border-border">
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Datum</th>
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Empfohlen</th>
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Gemacht</th>
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Abw. %</th>
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Gefolgt</th>
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Reaktion</th>
                    <th className="pb-1 pr-2 font-normal text-ink-dim">Quelle</th>
                  </tr>
                </thead>
                <tbody>
                  {trefferTage.slice(0, 14).map((t, i) => (
                    <tr key={i} className="border-b border-border">
                      <td className="pr-2 pt-1 tnum">{fmtDateShort(t.datum)}</td>
                      <td className="pr-2 pt-1">
                        <span className="tnum">{t.empfohlen !== null ? fmtNum(t.empfohlen) : '–'}</span>
                      </td>
                      <td className="pr-2 pt-1">
                        <span className="tnum">{t.ist !== null ? fmtNum(t.ist) : '–'}</span>
                      </td>
                      <td className="pr-2 pt-1 tnum text-xs">
                        {t.abweichungPct !== null ? (t.abweichungPct > 0 ? '+' : '') + fmtNum(t.abweichungPct) + ' %' : 'Ruhetag'}
                      </td>
                      <td className="pr-2 pt-1">
                        {t.gefolgt !== null && (
                          <span className={t.gefolgt ? 'text-green-300' : 'text-orange-300'}>
                            {t.gefolgt ? 'ja' : 'nein'}
                          </span>
                        )}
                      </td>
                      <td className="pr-2 pt-1 tnum">
                        {t.score !== null ? (
                          <span className={`tnum ${
                            t.score > 0.5 ? 'text-green-300' : t.score < -0.5 ? 'text-orange-300' : ''
                          }`}>
                            {fmtNum(t.score)}
                          </span>
                        ) : '—'}
                      </td>
                      <td className="pt-1 text-ink-dim">{t.quelle === 'plan-briefing' ? 'nachgetragen' : t.quelle}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      <VotumTrefferTile />

      {/* Section B: Selbsttest lokales Modell */}
      <div className="mt-6">
        <div className="label">Selbsttest lokales Modell</div>
        {errTest && <p className="mt-2 text-xs text-orange-300">{errTest}</p>}

        {neuesterTest && (
          <div className="mt-3 rounded border border-border p-3">
            {neuesterTest.text && (
              <pre className="whitespace-pre-line text-xs text-ink">{neuesterTest.text}</pre>
            )}

            {neuesterTest.tests.length > 0 && (
              <div className="mt-3 overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="border-b border-border">
                      <th className="pb-1 pr-2 font-normal text-ink-dim">Name</th>
                      <th className="pb-1 pr-2 font-normal text-ink-dim">Ergebnis</th>
                      <th className="pb-1 pr-2 font-normal text-ink-dim">Detail</th>
                      <th className="pb-1 pr-2 font-normal text-ink-dim">Zeit</th>
                    </tr>
                  </thead>
                  <tbody>
                    {neuesterTest.tests.map((t, i) => (
                      <tr key={i} className="border-b border-border">
                        <td className="pr-2 pt-1 text-xs">{t.name}</td>
                        <td className="pr-2 pt-1">
                          {t.ok ? (
                            <span className="text-green-300 text-xs">bestanden</span>
                          ) : (
                            <span className="text-orange-300 text-xs">nicht bestanden</span>
                          )}
                        </td>
                        <td className="pr-2 pt-1 text-xs text-ink-dim">{t.detail || '–'}</td>
                        <td className="pt-1 tnum text-xs">{t.sekunden}s</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {testVerlauf.length > 0 && (
              <div className="mt-3">
                <div className="text-2xs text-ink-dim mb-1">Verlauf der letzten 14 Tage:</div>
                <div className="flex flex-wrap gap-2 text-2xs">
                  {testVerlauf.map((v, i) => (
                    <span key={i} className="text-ink-dim">
                      {fmtDateShort(v.datum)}: <span className="tnum">{v.okCount}/6</span>
                    </span>
                  ))}
                </div>
                <p className="mt-1 text-2xs text-ink-dim">Läuft nachts auf dem Mac.</p>
              </div>
            )}

            <div className="mt-2 text-2xs text-ink-dim">
              Erstellt {neuesterTest.gespeichert} · {neuesterTest.quelle}
            </div>
          </div>
        )}

        {!neuesterTest && !errTest && (
          <p className="mt-2 text-xs text-ink-muted">Keine Selbsttest-Daten vorhanden.</p>
        )}
      </div>
    </div>
  );
}
