import { useEffect, useRef, useState } from 'react';
import { wochenplanAbgleich, EVT_WOCHENPLAN } from '../lib/geraeteSync';
import { ladeP3, esBacktest, esBreakEven, esZiel, type P3Tag } from '../lib/paket3';
import { useGarminTage, fmtDauer } from '../lib/garmin';
import { laengsteProTag, langEffekt, langStatus, langEmpfehlung, langPlan, sportDe, parseLangEinstellung, essProStunde, esGesamt, zielBedarf, WT_KURZ, type LangEinstellung, type ZielKurz } from '../lib/esLang';
import { fetchExtras } from '../lib/extras';
import { fetchPlanBriefing } from '../lib/api';
import { heuteBerlin } from '../lib/kernVotumKi';

function fmtDateShort(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  return `${parseInt(p[2])}.${p[1].padStart(2, '0')}.${p[0]}`;
}

// ── Komponente ───────────────────────────────────────────────────────────
export function EsHebelTile() {
  const [tage, setTage] = useState<P3Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const { tage: garmin } = useGarminTage();
  // Lange Einheiten: Wochentage und sehr lange pro Monat (Geräte-Einstellungen, gelten auf allen Geräten)
  const [langEin, setLangEin] = useState<LangEinstellung>(() => {
    try { return parseLangEinstellung(localStorage.getItem('ck_lang_tage'), localStorage.getItem('ck_lang_sehr')); } catch { return parseLangEinstellung(null, null); }
  });
  const langSpeichern = (e: LangEinstellung) => {
    setLangEin(e);
    try {
      localStorage.setItem('ck_lang_tage', e.tage.join(','));
      localStorage.setItem('ck_lang_sehr', String(e.sehrProMonat));
    } catch { /* egal */ }
  };
  const tagUmschalten = (wt: number) =>
    langSpeichern({ ...langEin, tage: langEin.tage.includes(wt) ? langEin.tage.filter((x) => x !== wt) : [...langEin.tage, wt].sort((a, b) => a - b) });
  // Zieltouren (Zielplaner/Tourform) steuern Länge, Sport und Tapering der langen Einheiten
  const [ziele, setZiele] = useState<ZielKurz[]>([]);
  useEffect(() => {
    let alive = true;
    fetchExtras<{ ziel?: ZielKurz }>('ziel', 60)
      .then((it) => { if (alive) setZiele(it.map((x) => x.data?.ziel).filter((z): z is ZielKurz => !!z && !!z.datum)); })
      .catch(() => {});
    return () => { alive = false; };
  }, []);
  // Kiras Vorschlag aus dem letzten Plan-Briefing (Feld lange_einheit je Tag)
  const [kiVorschlag, setKiVorschlag] = useState<{ stand: string | null; quelle?: string; tage: Record<string, { dauer_h?: number; sport?: string; es_effekt?: number; hinweis?: string; ess?: number }> } | null>(null);
  const [wpStand, setWpStand] = useState(0);
  useEffect(() => {
    const neu = () => setWpStand((n) => n + 1);
    window.addEventListener(EVT_WOCHENPLAN, neu);
    wochenplanAbgleich().catch(() => {});
    return () => window.removeEventListener(EVT_WOCHENPLAN, neu);
  }, []);
  useEffect(() => {
    let alive = true;
    // zuerst Kiras Wochenplan (Plan › Diese Woche, auf allen Geräten gleich), sonst das Plan-Briefing vom Morgenlauf
    try {
      const wp = JSON.parse(localStorage.getItem('ck_kira_wochenplan_v1') || 'null');
      const tage: Record<string, any> = {};
      for (const x of (wp && wp.tage) || []) if (x && x.date && x.langH != null) tage[x.date] = { dauer_h: x.langH, sport: sportDe(x.sport), es_effekt: x.esEffekt ?? undefined, hinweis: x.why || undefined, ess: x.load };
      if (Object.keys(tage).length) { setKiVorschlag({ stand: wp.erstellt || null, quelle: 'Wochenplan', tage }); return () => { alive = false; }; }
    } catch { /* weiter mit dem Plan-Briefing */ }
    fetchPlanBriefing()
      .then((r: any) => {
        if (!alive || !r || !r.ok) return;
        const tage: Record<string, any> = {};
        for (const x of r.recommendations || []) if (x && x.date && x.lange_einheit && typeof x.lange_einheit === 'object') tage[String(x.date).slice(0, 10)] = { ...x.lange_einheit, ess: x.recommended_load_ess };
        setKiVorschlag({ stand: r.briefingTimestamp || null, quelle: 'Plan-Briefing', tage });
      })
      .catch(() => {});
    return () => { alive = false; };
  }, [wpStand]);

  // Ziel-Rechner state
  const [zielScore, setZielScore] = useState<number>(() => {
    try {
      const n = Number(localStorage.getItem('ck_es_ziel'));
      if (Number.isFinite(n) && n >= 1000 && n <= 20000) return n;
    } catch {}
    return 7000;
  });
  const [zielDatum, setZielDatum] = useState<string>(() => {
    try {
      const saved = localStorage.getItem('ck_es_datum');
      if (saved) return saved;
    } catch {}
    const d = new Date();
    d.setDate(d.getDate() + 140);
    return d.toISOString().slice(0, 10);
  });

  const zielRef = useRef<HTMLDivElement>(null);
  const [zielErgebnis, setZielErgebnis] = useState<{ tage: number; ladeProTag: number | null; machbar: boolean } | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeP3();
        if (!cancelled) setTage(t);
        setLoading(false);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // Ziel-Rechner berechnen wenn Inputs sich ändern
  useEffect(() => {
    if (tage.length === 0) return;
    // Ziel-Datum speichern
    try { localStorage.setItem('ck_es_datum', zielDatum); } catch {}
    // Ziel-Score speichern
    try { localStorage.setItem('ck_es_ziel', String(zielScore)); } catch {}
    const result = esZiel(tage, zielScore, zielDatum);
    setZielErgebnis(result);
  }, [tage.length, zielScore, zielDatum]);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Endurance-Score-Hebel lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  // letzterES: letzter Tag mit es
  let letzterES: number | null = null;
  for (let i = tage.length - 1; i >= 0; i--) {
    if (tage[i].es !== null) {
      letzterES = tage[i].es;
      break;
    }
  }

  // last28: Ø load der letzten 28 Tage (null = 0)
  const last28Count = Math.min(tage.length, 28);
  let last28: number | null = null;
  if (last28Count > 0) {
    // 28 abgeschlossene Tage vor heute, fehlende Last = 0 (wie im Modell)
    const fenster = tage.slice(Math.max(0, tage.length - 29), tage.length - 1);
    last28 = fenster.reduce((acc, x) => acc + (x.load ?? 0), 0) / Math.max(1, fenster.length);
  }

  // Gleichgewicht: esBreakEven(letzterES)
  let gleichgewicht: number | null = null;
  if (letzterES !== null) {
    gleichgewicht = esBreakEven(letzterES);
  }

  // Belastbarkeit: b = esBacktest(tage)
  const backtest = tage.length > 0 ? esBacktest(tage) : { n: 0, maeModell: 0, maeNaiv: 0, horizont: 28 };
  const notBetter = backtest.maeModell >= backtest.maeNaiv;

  // Ziel-Rechner: z = esZiel(tage, ziel, datum)
  const ziel = zielScore;
  const datum = zielDatum;
  const z = zielErgebnis;

  // Lange Einheiten (Garmin-Tagesdetails): Status, eigene Wirkung, Empfehlung
  const lang = garmin ? laengsteProTag(garmin) : null;
  const heute = heuteBerlin();
  const eff = lang ? langEffekt(tage, lang) : null;
  const st = lang ? langStatus(lang, heute) : null;
  const emp = st ? langEmpfehlung(st, langEin) : null;
  const ep = lang ? essProStunde(tage, lang) : null;
  const zb = st ? zielBedarf(ziele, st) : [];
  const termine = st && eff ? langPlan(langEin, st, eff, 14, ep || undefined, zb) : [];
  const extraProTag = eff && emp ? (eff.bonus4 * emp.sehrLangProMonat) / 30 : 0;
  const zLang = eff && eff.bonus4 > 0 && emp && emp.sehrLangProMonat > 0 && z && z.tage > 0 ? esZiel(tage, zielScore, zielDatum, extraProTag) : null;
  const fmtH = (h: number) => fmtDauer(h * 3600);
  const vz = (x: number | null) => (x == null ? '—' : `${x >= 0 ? '+' : '−'}${Math.abs(Math.round(x))}`);

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Endurance-Score-Hebel</div>
        <p className="mt-1 text-2xs text-ink-muted">
          Wie deine mittlere Tageslast und deine langen Einheiten deinen Garmin Endurance Score bewegen – mit deinem eigenen Fortschreibungsmodell.
        </p>
      </div>

      {/* Kennzahlen Raster 3 Spalten */}
      <div className="mt-3 grid grid-cols-1 sm:grid-cols-3 gap-2">
        {/* Endurance Score */}
        <div className="rounded-lg border border-border p-3">
          <p className="text-xs text-ink-muted">Endurance Score</p>
          <p className="mt-1 text-lg font-mono">{letzterES !== null ? Math.round(letzterES) : '—'}</p>
        </div>

        {/* Ø Last 28 Tage */}
        <div className="rounded-lg border border-border p-3">
          <p className="text-xs text-ink-muted">Ø Last 28 Tage</p>
          <p className="mt-1 text-lg font-mono">{last28 !== null ? Math.round(last28) : '—'} ESS/Tag</p>
        </div>

        {/* Gleichgewicht */}
        <div className="rounded-lg border border-border p-3">
          <p className="text-xs text-ink-muted">Gleichgewicht</p>
          <p className="mt-1 text-lg font-mono">
            {gleichgewicht !== null ? Math.round(gleichgewicht) : '—'} ESS/Tag
          </p>
          <p className="mt-1 text-2xs text-ink-muted">
            darüber steigt der Score, darunter sinkt er (bei deinem jetzigen Niveau)
          </p>
        </div>
      </div>

      {/* Belastbarkeit */}
      <div className="mt-3">
        <p className="text-xs text-ink-muted">
          Belastbarkeit: Geprüft an {backtest.n} Zeitpunkten: 28 Tage voraus im Schnitt ±{Math.round(backtest.maeModell)} Punkte daneben, „bleibt gleich" ±{Math.round(backtest.maeNaiv)}.
        </p>
        {notBetter && (
          <p className="mt-1 text-xs text-orange-300">
            Das Modell ist derzeit nicht besser als „bleibt gleich" – Werte nur als Richtung lesen.
          </p>
        )}
      </div>

      {/* Lange Einheiten */}
      <div className="mt-4">
        <p className="text-xs text-ink-muted mb-2">Lange Einheiten:</p>
        {!lang ? (
          <p className="text-xs text-ink-dim">Garmin-Tagesdetails laden…</p>
        ) : st && emp && eff ? (
          <>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              <div className={`rounded-lg border p-3 ${emp.fenster14Leer ? 'border-orange-400/60' : 'border-border'}`}>
                <p className="text-xs text-ink-muted">Längste der letzten 14 Tage</p>
                <p className="mt-1 text-lg font-mono">{st.laengste14 ? fmtH(st.laengste14.h) : '—'}</p>
                <p className="mt-1 text-2xs text-ink-muted">
                  {st.laengste14 ? `${sportDe(st.laengste14.sport)} am ${fmtDateShort(st.laengste14.date)}` : 'keine Aktivität'}
                  {emp.fenster14Leer ? ' – Garmins 2-Wochen-Fenster hat keine lange Einheit' : ''}
                </p>
              </div>
              <div className="rounded-lg border border-border p-3">
                <p className="text-xs text-ink-muted">Ab 2 h</p>
                <p className="mt-1 text-lg font-mono">{st.n2_28} / {st.n2_84}</p>
                <p className="mt-1 text-2xs text-ink-muted">in 28 / 84 Tagen (Ziel etwa {langEin.tage.length * 4} / {langEin.tage.length * 12})</p>
              </div>
              <div className="rounded-lg border border-border p-3">
                <p className="text-xs text-ink-muted">Ab 4 h</p>
                <p className="mt-1 text-lg font-mono">{st.n4_84}</p>
                <p className="mt-1 text-2xs text-ink-muted">
                  in 84 Tagen (Ziel etwa {langEin.sehrProMonat * 3}){st.letzte4 ? `, zuletzt ${fmtDateShort(st.letzte4.date)}` : ''}
                </p>
              </div>
            </div>

            <div className="mt-3">
              <p className="text-xs font-medium">Deine Vorgabe</p>
              <div className="mt-1 flex flex-wrap items-center gap-1">
                <span className="text-2xs text-ink-muted mr-1">Tage für lange Einheiten:</span>
                {[1, 2, 3, 4, 5, 6, 7].map((wt) => {
                  const an = langEin.tage.includes(wt);
                  return (
                    <button
                      key={wt}
                      type="button"
                      aria-pressed={an}
                      onClick={() => tagUmschalten(wt)}
                      className={`px-2 py-0.5 rounded border text-xs ${an ? 'bg-accent-strong/20 border-accent-strong/60 text-accent font-medium' : 'border-border text-ink-muted'}`}
                    >
                      {WT_KURZ[wt]}
                    </button>
                  );
                })}
                <span className="text-2xs text-ink-muted ml-1">= {langEin.tage.length} pro Woche</span>
              </div>
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <label htmlFor="es-lang-sehr" className="text-2xs text-ink-muted">Sehr lange (ab 4 h) pro Monat:</label>
                <select
                  id="es-lang-sehr"
                  value={langEin.sehrProMonat}
                  onChange={(e) => langSpeichern({ ...langEin, sehrProMonat: Number(e.target.value) })}
                  className="px-2 py-0.5 border border-border rounded text-xs"
                >
                  {[0, 1, 2, 3, 4].map((n) => <option key={n} value={n}>{n}</option>)}
                </select>
                <span className="text-2xs text-ink-muted">gilt auf allen Geräten; Kira berücksichtigt es beim nächsten Wochenplan („Neu planen“ unter Plan › Diese Woche) und im Plan-Briefing</span>
              </div>
            </div>

            {zb.length > 0 && (
              <div className="mt-3">
                <p className="text-xs font-medium">Zieltouren (aus der Tourform)</p>
                <ul className="mt-1 text-xs space-y-1">
                  {zb.slice(0, 3).map((b) => (
                    <li key={b.ziel.datum + b.ziel.name}>
                      <span className="font-medium">{b.ziel.name}</span> am {fmtDateShort(b.ziel.datum)} ({b.wochen} Wochen, {b.ziel.dauer_h} h): längste Einheit bis{' '}
                      {fmtDateShort(b.aufbauBis)} auf etwa {fmtH(b.zielH)} ({b.sport}{b.hmZiel ? `, ${b.hmZiel.toLocaleString('de-DE')} Hm` : ''}) – nötig{' '}
                      {fmtH(b.noetigProWoche)} pro Woche, erlaubt {fmtH(b.schrittH)}:{' '}
                      {b.machbar ? <span className="text-green-600">machbar</span> : <span className="text-orange-300">knapp, nicht erzwingen</span>}. Tapering ab {fmtDateShort(b.taperAb)}.
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {termine.length > 0 && (
              <div className="mt-3">
                <p className="text-xs font-medium">Termine in den nächsten 14 Tagen</p>
                <div className="mt-1 overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="text-2xs text-ink-muted text-left">
                        <th className="py-1 pr-2 font-normal">Tag</th>
                        <th className="py-1 pr-2 font-normal">Art</th>
                        <th className="py-1 pr-2 font-normal">Richtwert</th>
                        <th className="py-1 pr-2 font-normal">ESS</th>
                        <th className="py-1 pr-2 font-normal">Endurance</th>
                        <th className="py-1 font-normal">Kira schlägt vor</th>
                      </tr>
                    </thead>
                    <tbody>
                      {termine.map((t) => {
                        const k = kiVorschlag?.tage[t.date];
                        return (
                          <tr key={t.date} className="border-t border-border align-top">
                            <td className="py-1 pr-2 whitespace-nowrap">{t.wt} {fmtDateShort(t.date).slice(0, -5)}</td>
                            <td className="py-1 pr-2">
                              {{ sehr_lang: 'sehr lang', lang: 'lang', ziel: 'Zieltour', taper: 'Tapering' }[t.art]}
                              {t.sport ? <span className="text-2xs text-ink-muted"> · {t.sport}</span> : null}
                            </td>
                            <td className="py-1 pr-2 font-mono whitespace-nowrap">{fmtH(t.dauerH)}</td>
                            <td className="py-1 pr-2 font-mono whitespace-nowrap">{t.essRad ? `${t.essRad.von}–${t.essRad.bis}` : '—'}</td>
                            <td className="py-1 pr-2 font-mono whitespace-nowrap" title={`Dauer-Zuschlag ${vz(t.esEffekt)}, Rest aus der Last`}>{t.esGes != null ? vz(t.esGes) : vz(t.esEffekt)}</td>
                            <td className="py-1">
                              {k ? (
                                <span>
                                  <span className="font-mono">{k.dauer_h != null ? fmtH(Number(k.dauer_h)) : '—'}</span>
                                  {k.sport ? ` ${k.sport}` : ''}
                                  {k.dauer_h != null && eff && k.ess != null ? (
                                    <span className="font-mono">, {Math.round(Number(k.ess))} ESS, {vz(esGesamt(Number(k.ess), Number(k.dauer_h), eff).gesamt)}</span>
                                  ) : k.es_effekt != null ? <span className="font-mono">, {vz(Number(k.es_effekt))}</span> : null}
                                  {k.dauer_h != null && eff && k.ess != null && k.es_effekt != null && Math.abs(Number(k.es_effekt) - esGesamt(Number(k.ess), Number(k.dauer_h), eff).gesamt) > 10 ? (
                                    <span className="text-2xs text-ink-muted"> (Kira schätzte {vz(Number(k.es_effekt))})</span>
                                  ) : null}
                                  {k.hinweis ? <span className="text-2xs text-ink-muted"> – {k.hinweis}</span> : null}
                                </span>
                              ) : (
                                <span className="text-2xs text-ink-muted">noch kein Vorschlag</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <p className="mt-1 text-2xs text-ink-muted">
                  Richtwert, ESS-Spanne (ruhige lange Einheiten der Sportart, sonst Rad) und Endurance (Dauer-Zuschlag + 0,17 Punkte je ESS gegenüber einem Ruhetag) aus deiner Historie. Kiras Vorschlag kommt aus {kiVorschlag?.quelle === 'Wochenplan' ? 'dem Wochenplan (Plan › Diese Woche)' : 'dem Plan-Briefing des Morgenlaufs'}
                  {kiVorschlag?.stand ? ` vom ${new Date(kiVorschlag.stand).toLocaleString('de-DE', { day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' })}` : ''} – neue Vorgaben wirken nach „Neu planen“.
                </p>
              </div>
            )}

            <div className="mt-3">
              <p className="text-xs font-medium">Empfehlung</p>
              <ul className="mt-1 text-xs space-y-1 list-disc pl-4">
                {emp.punkte.map((p) => <li key={p}>{p}</li>)}
              </ul>
              <p className="mt-2 text-xs">
                Nächste lange Einheit ({String(emp.wocheDauerH).replace('.', ',')} h):{' '}
                <span className={emp.wocheUeberfaellig ? 'text-orange-300' : ''}>
                  {emp.wocheUeberfaellig ? `überfällig seit ${fmtDateShort(emp.wocheFaellig)}` : `bis ${fmtDateShort(emp.wocheFaellig)}`}
                </span>
                {' · '}Nächste sehr lange (ab 4 h):{' '}
                {emp.sehrLangAufbau ? (
                  <span className="text-ink-muted">nach dem Aufbau auf etwa 3 h</span>
                ) : (
                  <span className={emp.sehrLangUeberfaellig ? 'text-orange-300' : ''}>
                    {emp.sehrLangUeberfaellig ? `überfällig seit ${fmtDateShort(emp.sehrLangFaellig)}` : `bis ${fmtDateShort(emp.sehrLangFaellig)}`}
                  </span>
                )}
              </p>
            </div>

            <div className="mt-3">
              <p className="text-xs font-medium">Was lange Einheiten bei dir zusätzlich bringen</p>
              <p className="mt-1 text-2xs text-ink-muted">
                Score-Punkte in den Tagen 0–3 danach, über das hinaus, was ihre ESS-Last schon erklärt ({eff.basisTage} Tage geprüft{eff.quelle ? `, ${eff.quelle}` : ''}):
              </p>
              <div className="mt-1 grid grid-cols-2 gap-2 text-xs sm:grid-cols-4">
                {eff.klassen.map((k) => (
                  <div key={k.lab} className="rounded border border-border px-2 py-1">
                    <span className="text-ink-muted">{k.lab}: </span>
                    <span className="font-mono">{vz(k.extra)}</span>
                    <span className="text-2xs text-ink-muted"> (n={k.n})</span>
                  </div>
                ))}
              </div>
              {eff.n4 < 12 && (
                <p className="mt-1 text-2xs text-orange-300">
                  Erst {eff.n4} sehr lange Einheiten mit Score-Verlauf, die Werte streuen stark – Größenordnung, keine feste Zahl. Gerechnet wird vorsichtig mit {vz(eff.bonus4)} je Einheit.
                </p>
              )}
            </div>
          </>
        ) : null}
      </div>

      {/* Ziel-Rechner */}
      <div ref={zielRef} className="mt-4">
        <p className="text-xs text-ink-muted mb-2">Ziel-Rechner:</p>

        <div className="flex gap-3 mb-2 flex-wrap">
          {/* Ziel-Score */}
          <div>
            <label htmlFor="es-ziel-score" className="text-2xs text-ink-muted">Ziel-Score</label>
            <input
              id="es-ziel-score"
              type="number"
              value={zielScore}
              onChange={(e) => setZielScore(Number(e.target.value) || 7000)}
              className="ml-1 px-2 py-1 border border-border rounded text-xs w-32"
            />
          </div>

          {/* Zieldatum */}
          <div>
            <label htmlFor="es-ziel-datum" className="text-2xs text-ink-muted">Zieldatum</label>
            <input
              id="es-ziel-datum"
              type="date"
              value={zielDatum}
              onChange={(e) => setZielDatum(e.target.value)}
              className="ml-1 px-2 py-1 border border-border rounded text-xs w-36"
            />
          </div>
        </div>

        {/* Ergebnis */}
        {z !== null && z.tage > 0 && (
          <div className="mt-2">
            {z.machbar && z.ladeProTag !== null ? (
              <p className="text-xs">
                Für {ziel} bis {fmtDateShort(datum)} brauchst du im Schnitt etwa {Math.round(z.ladeProTag)} ESS pro Tag ({Math.round(z.ladeProTag * 7).toLocaleString("de-DE")} pro Woche) über {z.tage} Tage.
              </p>
            ) : (
              <p className="text-xs text-orange-300">
                Mit bis zu 250 ESS pro Tag nicht erreichbar – Ziel oder Datum anpassen.
              </p>
            )}
            {zLang && zLang.machbar && zLang.ladeProTag !== null && z.ladeProTag !== null && (
              <p className="mt-1 text-xs">
                Mit {emp?.sehrLangProMonat} sehr langen Einheit{emp?.sehrLangProMonat === 1 ? '' : 'en'} (ab 4 h) pro Monat reichen etwa {Math.round(zLang.ladeProTag)} ESS pro Tag
                {z.ladeProTag - zLang.ladeProTag >= 1 ? ` (${Math.round(z.ladeProTag - zLang.ladeProTag)} weniger)` : ''}.
              </p>
            )}
            {last28 !== null && (
              <p className="mt-1 text-xs text-ink-muted">Aktuell: {Math.round(last28)} ESS/Tag.</p>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
