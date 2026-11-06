import { useEffect, useState } from 'react';
import {
  ladeTage,
  reaktionen,
  preisGruppen,
  erwartetDR,
  sportGruppe,
  zonenKlasse,
  type ETag,
} from '../lib/recoveryPrice';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

function fmtDe(v: number): string {
  return v.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function fmtSigned(v: number): string {
  const r = fmtDe(Math.abs(v));
  return v >= 0 ? '+' + r : '−' + r;
}

function dateShortFromISO(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  const dow = WD[new Date(iso.slice(0, 10) + 'T12:00:00').getDay()];
  return `${dow} ${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

function lastDayColor(dR: number): string {
  if (dR >= 0) return 'text-green-300';
  if (dR > -8) return 'text-yellow-300';
  return 'text-orange-300';
}

function urteilColor(u: string): string {
  if (u === 'belastbar') return 'text-green-300';
  if (u === 'eingeschränkt') return 'text-yellow-300';
  return 'text-ink-dim';
}

// ── Komponente ───────────────────────────────────────────────────────────
export function ErholungspreisTile({ plan }: { plan?: Array<{ date: string; load: number; sport: string; zone: string }> }) {
  const [tage, setTage] = useState<ETag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeTage();
        if (!cancelled) setTage(t);
        setLoading(false);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const reaktionenList = reaktionen(tage);
  const preisData = preisGruppen(reaktionenList);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Erholungspreis lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  // Sport-Tabelle: n >= 5 filtern
  const sportRows = preisData.sport.filter((g) => g.n >= 5);
  // Zone-Tabelle: locker, grau, hart
  const zoneRows = preisData.zone;

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Erholungspreis</div>
        <p className="mt-1 text-2xs text-ink-dim">
          Wie stark deine Trainingsbereitschaft am Folgemorgen auf die Last reagiert – gelernt aus deinen Daten,
          getrennt nach Sport und Zone. Zusammenhang, keine Ursache; die Readiness enthält Schlaf und HRV schon selbst.
        </p>
      </div>

      {/* Sport-Tabelle */}
      {sportRows.length > 0 && (
        <div className="mt-3">
          <span className="label">Nach Sport</span>
          <div className="overflow-x-auto mt-2">
            <table className="w-full text-xs tnum">
              <thead>
                <tr className="text-2xs text-ink-dim">
                  <th className="text-left pl-2">Sport</th>
                  <th className="text-right pr-2">n</th>
                  <th className="text-left">Urteil</th>
                  <th className="text-right">Ø Last</th>
                  <th className="text-right">100 ESS</th>
                  <th className="text-right">150 ESS</th>
                  <th className="text-right">200 ESS</th>
                </tr>
              </thead>
              <tbody>
                {sportRows.map((g) => (
                  <tr key={g.name} className="border-t border-border">
                    <td className="pl-2 text-ink">{g.name}</td>
                    <td className="text-right pr-2">{g.n}</td>
                    <td className={`text-left ${urteilColor(g.urteil)}`}>{g.urteil}</td>
                    <td className="text-right">{fmtDe(g.avgLoad)}</td>
                    {([100, 150, 200] as const).map((l) => {
                      const wert = g.a + g.b * l;
                      return (
                        <td key={l} className={`text-right ${lastDayColor(wert)}`}>
                          {fmtSigned(wert)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Zone-Tabelle */}
      {zoneRows.length > 0 && (
        <div className="mt-3">
          <span className="label">Nach Zone</span>
          <div className="overflow-x-auto mt-2">
            <table className="w-full text-xs tnum">
              <thead>
                <tr className="text-2xs text-ink-dim">
                  <th className="text-left pl-2">{'Zone' as any}</th>
                  <th className="text-right pr-2">n</th>
                  <th className="text-right">Ø Last</th>
                  <th className="text-right">Ø dR</th>
                </tr>
              </thead>
              <tbody>
                {zoneRows.map((g) => (
                  <tr key={g.name} className="border-t border-border">
                    <td className={`pl-2 ${g.name === 'grau' ? 'bg-gray-800/30 rounded-l text-ink-muted' : 'text-ink'}`}>
                      {g.name === 'grau' ? 'Grauzone' : g.name.charAt(0).toUpperCase() + g.name.slice(1)}
                    </td>
                    <td className="text-right pr-2">{g.n}</td>
                    <td className="text-right">{fmtDe(g.avgLoad)}</td>
                    <td className={`text-right ${lastDayColor(g.avgDR)}`}>{fmtSigned(g.avgDR)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Plan-Abschnitt */}
      {plan && plan.length > 0 && (
        <div className="mt-3">
          <span className="label">Dein Plan: erwartete Bereitschaft am Folgemorgen</span>
          <div className="overflow-x-auto mt-2">
            <table className="w-full text-xs tnum">
              <thead>
                <tr className="text-2xs text-ink-dim">
                  <th className="text-left pl-2">Tag</th>
                  <th className="text-right pr-2">Last / Sport + Zone</th>
                  <th className="text-right">Erwartet</th>
                </tr>
              </thead>
              <tbody>
                {plan.filter((p) => p.load > 30).slice(0, 7).map((p, i) => {
                  const erwartet = erwartetDR(preisData, p.load, p.sport, p.zone);
                  return (
                    <tr key={i} className="border-t border-border">
                      <td className="pl-2 text-ink">{dateShortFromISO(p.date)}</td>
                      <td className="text-right pr-2">
                        <span className="tnum">{p.load}</span> · {sportGruppe(p.sport)} ({zonenKlasse(p.zone)})
                      </td>
                      <td className="text-right" title={erwartet.quelle}>
                        <span className={lastDayColor(erwartet.wert)}>
                          {fmtSigned(Math.round(erwartet.wert * 10) / 10)}
                        </span>
                        ± <span className="text-ink-muted">{fmtDe(erwartet.sd)}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Reaktions-Verlauf */}
      {reaktionenList.length > 0 && (
        <div className="mt-3">
          <span className="label">Letzte Trainingstage: erwartet und eingetreten</span>
          <div className="overflow-x-auto mt-2">
            <table className="w-full text-xs tnum">
              <thead>
                <tr className="text-2xs text-ink-dim">
                  <th className="text-left pl-2">Datum</th>
                  <th className="text-right pr-2">Last / Sport</th>
                  <th className="text-right">Erwartet</th>
                  <th className="text-right">Eingetreten</th>
                  <th className="text-left"></th>
                </tr>
              </thead>
              <tbody>
                {reaktionenList.slice(-10).map((r, i) => {
                  const zoneVon = r.klasse === 'hart' ? 'Z4' : r.klasse === 'grau' ? 'Z3' : 'Z2';
                  const e = erwartetDR(preisData, r.load, r.gruppe, zoneVon);
                  const erwartetW = Math.round(e.wert * 10) / 10;
                  const erwarteteSd = e.sd;

                  let hinweis = '';
                  let hinweisColor = '';
                  if (r.dR < erwartetW - 1.5 * erwarteteSd) {
                    hinweis = 'deutlich schlechter als erwartet';
                    hinweisColor = 'text-orange-300';
                  } else if (r.dR > erwartetW + 1.5 * erwarteteSd) {
                    hinweis = 'deutlich besser als erwartet';
                    hinweisColor = 'text-green-300';
                  }

                  return (
                    <tr key={i} className="border-t border-border">
                      <td className="pl-2 text-ink">{dateShortFromISO(r.date)}</td>
                      <td className="text-right pr-2">
                        <span className="tnum">{r.load}</span> · {r.gruppe}
                      </td>
                      <td className="text-right">
                        <span className={lastDayColor(erwartetW)}>
                          {fmtSigned(erwartetW)}
                        </span>
                      </td>
                      <td className="text-right">{fmtSigned(r.dR)}</td>
                      <td className={`text-left text-2xs ${hinweisColor}`}>{hinweis}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
