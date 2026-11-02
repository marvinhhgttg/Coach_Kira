import { useMemo, useState } from 'react';
import { Panel } from './UI';
import { KiraFazit } from './KiraFazit';
import { useGarminTage } from '../lib/garmin';
import { GRUPPEN, rekorde, zeitraum, type Rekord, type RekordGruppe } from '../lib/rekorde';

const kurz = (d: string) => new Date(d + 'T12:00:00').toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: '2-digit' });

function Karte({ r }: { r: Rekord }) {
  return (
    <div className={`rounded border p-2 ${r.neu ? 'border-emerald-500/60' : 'border-border'}`} data-testid="rekord">
      <div className="flex items-start justify-between gap-2">
        <div className="text-2xs text-ink-dim leading-tight">{r.titel}</div>
        {r.neu && <span className="text-2xs text-emerald-400 whitespace-nowrap">neu</span>}
      </div>
      <div className="text-lg text-ink tnum leading-tight mt-0.5">{r.text}</div>
      <div className="text-2xs text-ink-muted tnum">
        {r.bis ? `${kurz(r.datum)}–${kurz(r.bis)}` : kurz(r.datum)}
        {r.aktId ? (
          <> · <a className="underline decoration-dotted hover:text-ink" href={`https://connect.garmin.com/modern/activity/${r.aktId}`} target="_blank" rel="noreferrer">Garmin</a></>
        ) : null}
      </div>
      {r.detail && <div className="text-2xs text-ink-dim mt-0.5 line-clamp-2">{r.detail}</div>}
      {r.zweiter && <div className="text-2xs text-ink-dim mt-0.5">{r.zweiter}</div>}
    </div>
  );
}

export function RekordeTile() {
  const { tage } = useGarminTage();
  const alle = useMemo(() => rekorde(tage), [tage]);
  const zr = useMemo(() => zeitraum(tage), [tage]);
  const gruppen = GRUPPEN.filter((g) => alle.some((r) => r.gruppe === g.id));
  const [gruppe, setGruppe] = useState<RekordGruppe | 'neu'>('lauf');
  const neu = alle.filter((r) => r.neu);
  const sicht = gruppe === 'neu' ? neu : alle.filter((r) => r.gruppe === gruppe);

  const facts = useMemo(() => {
    const f: string[] = [`Rekorde aus den Garmin-Aktivitäten ${zr ? `${kurz(zr.von)} bis ${kurz(zr.bis)}` : ''}. „neu“ = in den letzten 30 Tagen aufgestellt.`];
    for (const g of GRUPPEN) {
      const rs = alle.filter((r) => r.gruppe === g.id);
      if (!rs.length) continue;
      f.push(`${g.label}: ` + rs.map((r) => `${r.titel} ${r.text} am ${r.bis ? `${kurz(r.datum)}–${kurz(r.bis)}` : kurz(r.datum)}${r.neu ? ' (NEU)' : ''}${r.zweiter ? ` [${r.zweiter}]` : ''}`).join('; '));
    }
    f.push('Hinweis: Die Garmin-Laufleistung ist seit August 2026 höher skaliert (gleiches Tempo, mehr Watt) – Laufleistungs-Rekorde danach sind nur bedingt mit früher vergleichbar.');
    return f;
  }, [alle, zr]);

  const btn = (aktiv: boolean) => `px-2 py-1 text-xs rounded ${aktiv ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;

  return (
    <Panel title="Deine Rekorde">
      <div data-testid="rekorde">
        <p className="text-2xs text-ink-dim mb-2">
          Bestwerte aus deinen Garmin-Aktivitäten{zr ? ` (${kurz(zr.von)} bis ${kurz(zr.bis)})` : ''}. Tempo aus der Bewegungszeit, mit Mindestumfängen und
          GPS-Plausibilitätsprüfung. Grün umrandet = in den letzten 30 Tagen aufgestellt.
        </p>
        {tage === null ? (
          <p className="text-2xs text-ink-dim">lädt …</p>
        ) : !alle.length ? (
          <p className="text-2xs text-ink-dim">Keine Garmin-Aktivitäten gefunden.</p>
        ) : (
          <>
            <div className="flex flex-wrap gap-1 mb-3">
              <div className="inline-flex flex-wrap bg-bg rounded border border-border p-0.5">
                {gruppen.map((g) => (
                  <button key={g.id} className={btn(gruppe === g.id)} onClick={() => setGruppe(g.id)}>{g.label}</button>
                ))}
                <button className={btn(gruppe === 'neu')} onClick={() => setGruppe('neu')}>Neu ({neu.length})</button>
              </div>
            </div>
            {sicht.length ? (
              <div className="grid gap-2 grid-cols-2 md:grid-cols-3 xl:grid-cols-4">
                {sicht.map((r) => <Karte key={r.gruppe + r.titel} r={r} />)}
              </div>
            ) : (
              <p className="text-2xs text-ink-dim">In den letzten 30 Tagen kein neuer Rekord.</p>
            )}
            {gruppe === 'lauf' && (
              <p className="text-2xs text-ink-dim mt-2">Laufleistung: Garmin zeigt seit August 2026 bei gleichem Tempo mehr Watt – Rekorde danach sind nur bedingt mit früher vergleichbar.</p>
            )}
            <div className="mt-3">
              <KiraFazit title="Deine Rekorde" facts={facts} />
            </div>
          </>
        )}
      </div>
    </Panel>
  );
}
