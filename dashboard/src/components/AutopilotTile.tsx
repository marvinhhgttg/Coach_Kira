import { useEffect, useState } from 'react';
import { fetchExtras, type ExtraItem } from '../lib/extras';
import { fmtDateShort } from '../lib/format';

export interface AutopilotTag {
  datum: string;
  tag: string;
  last: number;
  sport: string;
  zone: string;
  fest?: boolean;
  grund?: string;
}

export function AutopilotTile({ planDates, onApply }: { planDates: string[]; onApply: (tage: AutopilotTag[]) => void }) {
  const [data, setData] = useState<ExtraItem<any> | null>(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    setBusy(true);
    setErr(null);
    fetchExtras(['autopilot'], 3)
      .then((list) => {
        if (list.length === 0) {
          setData(null);
          return;
        }
        // Suche den neuesten Eintrag, dessen data.tage-Elemente mindestens teilweise in planDates liegen
        let best: ExtraItem<any> | null = null;
        for (let i = 0; i < list.length; i++) {
          const d = list[i].data || {};
          const items: AutopilotTag[] = Array.isArray(d.tage) ? d.tage : [];
          const overlap = items.filter((t: AutopilotTag) => planDates.includes(t.datum));
          if (overlap.length > 0) {
            best = list[i];
            break;
          }
        }
        // Wenn kein Passender, nimm den neuesten (erste in Liste)
        setData(best || list[0]);
      })
      .catch((e) => setErr(e instanceof Error ? e.message : String(e)))
      .finally(() => setBusy(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [planDates.join(",")]);

  if (busy) {
    return (
      <div className="panel-raised p-4">
        <div className="label">Coach Kira · Wochen-Autopilot</div>
        <p className="text-2xs text-ink-dim mt-1">Lade Autopilot-Vorschlag…</p>
      </div>
    );
  }

  if (!data) {
    return (
      <div className="panel-raised p-4">
        <div className="label">Coach Kira · Wochen-Autopilot</div>
        <p className="text-2xs text-ink-dim mt-1">Sonntagabend prüft der Mac die neue Woche (Monotonie, Kalender, Wetter, Saisonrahmen) und schlägt Änderungen vor.</p>
        <p className="mt-3 text-sm text-ink-muted">Noch kein Vorschlag – der nächste kommt am Sonntagabend.</p>
      </div>
    );
  }

  const d = data.data || {};
  const tage: AutopilotTag[] = Array.isArray(d.tage) ? d.tage : [];
  const aenderungen: any[] = Array.isArray(d.aenderungen) ? d.aenderungen : [];
  const verstoesse: any[] = Array.isArray(d.verstoesse) ? d.verstoesse : [];

  return (
    <div className="panel-raised p-4">
      <div>
        <div className="label">Coach Kira · Wochen-Autopilot</div>
        <p className="text-2xs text-ink-dim mt-1">Sonntagabend prüft der Mac die neue Woche (Monotonie, Kalender, Wetter, Saisonrahmen) und schlägt Änderungen vor.</p>
      </div>

      {/* Wochenzeile */}
      <div className="mt-3 flex flex-wrap items-baseline gap-x-2 text-sm">
        <span>Woche ab {fmtDateShort(data.datum)}</span>
        <span className="text-ink-muted">
          Wochenlast {d.wochenlast_alt} → {d.wochenlast_neu}
        </span>
        <span className="text-ink-muted">
          ESS{d.saison_ziel ? ` (Saisonziel ${d.saison_ziel})` : ''}
        </span>
      </div>

      {/* Fazit */}
      {d.fazit && <p className="mt-2 text-sm text-ink">{d.fazit}</p>}

      {/* Änderungen-Tabelle */}
      {aenderungen.length > 0 && (
        <div className="mt-3 overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="border-b border-border text-ink-dim">
                <th className="py-1 pr-3 font-normal">Tag</th>
                <th className="py-1 pr-3 font-normal">Datum</th>
                <th className="py-1 pr-3 font-normal">Alt last/sport/zone</th>
                <th className="py-1 pr-3 font-normal">Neu last/sport/zone</th>
                <th className="py-1 font-normal">Grund</th>
              </tr>
            </thead>
            <tbody>
              {aenderungen.map((a: any, i: number) => (
                <tr key={i} className="border-b border-border/60">
                  <td className="py-1 pr-3 whitespace-nowrap">{a.tag}</td>
                  <td className="py-1 pr-3 whitespace-nowrap">{fmtDateShort(a.datum)}</td>
                  <td className="py-1 pr-3 whitespace-nowrap">
                    <span className="line-through text-ink-dim tnum">{a.alt?.last ?? ''}</span>
                    <span className="text-ink-dim"> {a.alt?.sport ?? ''} {a.alt?.zone ?? ''}</span>
                  </td>
                  <td className="py-1 pr-3 whitespace-nowrap">
                    <span className="tnum">{a.neu?.last ?? ''}</span>
                    <span> {a.neu?.sport ?? ''} {a.neu?.zone ?? ''}</span>
                  </td>
                  <td className="py-1 text-ink-muted">{a.grund ?? ''}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Keine Änderung */}
      {aenderungen.length === 0 && tage.length > 0 && (
        <p className="mt-3 text-xs text-ink-muted">Plan passt – keine Änderung nötig.</p>
      )}

      {/* Regelverstöße */}
      {d.verstoesse_vorher && d.verstoesse ? (
        <div className="mt-3">
          <p className="text-xs text-ink">
            Regelverstöße: vorher {d.verstoesse_vorher.length} → nachher {d.verstoesse.length}
          </p>
          {d.verstoesse.length > 0 && (
            <ul className="mt-1 text-xs text-orange-300 list-disc pl-5">
              {d.verstoesse.map((v: any, i: number) => (
                <li key={i}>{v}</li>
              ))}
            </ul>
          )}
        </div>
      ) : null}

      {/* Knopf */}
      {aenderungen.length > 0 && (
        <>
          <button className="mt-3 btn btn-primary text-xs px-2 py-1" onClick={() => onApply(tage)}>
            In Slider übernehmen
          </button>
          <p className="mt-1 text-2xs text-ink-dim">Gespeichert wird erst mit „In timeline speichern“.</p>
        </>
      )}

      {/* Quelle */}
      <p className="mt-3 text-2xs text-ink-dim">Erstellt {data.gespeichert} · {data.quelle}</p>
    </div>
  );
}
