import { useEffect, useMemo, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { ALL_RANGE } from '../lib/windows';
import { fmtDateShort } from '../lib/format';
import { FENSTER_LABEL, type FensterArt, type Zeitfenster } from '../lib/planning';
import {
  kalenderVorschlaege, ladeFenster, loescheFenster, rueckblick, saisonAnpassen, speichereFenster, vorschlagErfasst,
  type FensterRueckblick, type IstZeile, type KalVorschlag,
} from '../lib/zeitfenster';

const heute = () => new Date().toLocaleDateString('sv-SE');
const zahl = (v: unknown) => (v === '' || v == null || !Number.isFinite(Number(v)) ? null : Number(v));
const ARTEN: FensterArt[] = ['entlastung', 'pause', 'aktiv'];
const ART_HINT: Record<FensterArt, string> = {
  entlastung: 'Dienstreise, Familienwoche – wird als Entlastungswoche eingeplant (ca. 65 %)',
  pause: 'Messe, Krankenhaus, Umzug – Pause, Fitnessverlust wird vorab eingerechnet',
  aktiv: 'Bergurlaub, Skiwoche – Belastungswoche, die Woche davor wird zur Entlastung',
};
const ART_CLS: Record<FensterArt, string> = { entlastung: 'text-green-600 dark:text-green-300', pause: 'text-ink-muted', aktiv: 'text-orange-600 dark:text-orange-300' };
const vorschlagArt = (v: KalVorschlag): FensterArt => (v.typ === 'reise' ? 'entlastung' : 'aktiv');
const vorschlagTitel = (v: KalVorschlag) => {
  const s = v.stichworte[0] || (v.typ === 'reise' ? 'reise' : 'urlaub');
  return s.charAt(0).toUpperCase() + s.slice(1);
};
const f2 = (x: number | null, d = 0) => (x == null ? '–' : x.toLocaleString('de-DE', { maximumFractionDigits: d }));

/** Zeitfenster: wann eine Entlastung passt, kein Training geht oder ein Aktivurlaub ansteht – mit Kalendervorschlägen und Rückblick. */
export function ZeitfensterTile() {
  const [fenster, setFenster] = useState<Zeitfenster[] | null>(null);
  const [kal, setKal] = useState<KalVorschlag[] | null>(null);
  const [kalAlt, setKalAlt] = useState<KalVorschlag[]>([]);
  const [kalErr, setKalErr] = useState<string | null>(null);
  const [ist, setIst] = useState<IstZeile[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [msg, setMsg] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [von, setVon] = useState('');
  const [bis, setBis] = useState('');
  const [art, setArt] = useState<FensterArt>('entlastung');
  const [titel, setTitel] = useState('');
  const [artWahl, setArtWahl] = useState<Record<string, FensterArt>>({});

  const neuLaden = async () => {
    try {
      setFenster(await ladeFenster());
      setErr(null);
    } catch (e) {
      setFenster([]);
      setErr(/Typ|JSON|zeitfenster/i.test(String(e)) ? 'Apps-Script-Patch „Zeitfenster“ fehlt – Zeitfenster lassen sich erst danach speichern.' : String(e));
    }
  };
  useEffect(() => {
    void neuLaden();
    kalenderVorschlaege(heute(), 200).then(setKal).catch((e) => setKalErr(String(e instanceof Error ? e.message : e)));
    kalenderVorschlaege('2025-08-26', 430).then((xs) => setKalAlt(xs.filter((x) => x.bis < heute()))).catch(() => setKalAlt([]));
    fetchChartData(ALL_RANGE, ['load_fb_day', 'Garmin_Training_Readiness', 'rhr_bpm', 'sleep_score_0_100'])
      .then((r: any) =>
        setIst(((r && r.data) || []).map((x: any) => ({ date: String(x.date), load: zahl(x.load_fb_day), ready: zahl(x.Garmin_Training_Readiness), rhr: zahl(x.rhr_bpm), schlaf: zahl(x.sleep_score_0_100) }))),
      )
      .catch(() => setIst([]));
  }, []);

  async function nachAenderung(text: string) {
    const fs = await ladeFenster();
    setFenster(fs);
    try {
      const r = await saisonAnpassen(fs);
      setMsg(r ? `${text} Saisonrahmen angepasst (${r.geaendert} von ${r.wochen} Wochen neu gelegt) – Autopilot und Vorausblick am Mac lesen denselben Stand.` : `${text} Noch kein Saisonrahmen gespeichert – beim Erstellen werden die Zeitfenster berücksichtigt.`);
    } catch (e) {
      setMsg(`${text} Saisonrahmen konnte nicht angepasst werden: ${String(e instanceof Error ? e.message : e)}`);
    }
  }

  async function hinzufuegen(f: Zeitfenster) {
    setBusy(true);
    setMsg(null);
    try {
      if (f.bis < f.von) throw new Error('„bis“ liegt vor „von“.');
      if ((fenster || []).some((x) => x.von === f.von && x.id !== f.id)) throw new Error('An diesem Tag beginnt schon ein Zeitfenster – bitte anpassen oder das alte löschen.');
      await speichereFenster(f);
      await nachAenderung(`„${f.titel}“ gespeichert.`);
      setVon(''); setBis(''); setTitel('');
    } catch (e) {
      const t = String(e instanceof Error ? e.message : e);
      setMsg(/Unbekannter Typ/.test(t) ? 'Speichern nicht möglich: Apps-Script-Patch „Zeitfenster“ fehlt noch.' : t);
    } finally {
      setBusy(false);
    }
  }
  async function loeschen(f: Zeitfenster) {
    setBusy(true);
    try {
      await loescheFenster(f);
      await nachAenderung(`„${f.titel}“ gelöscht.`);
    } catch (e) {
      setMsg(String(e instanceof Error ? e.message : e));
    } finally {
      setBusy(false);
    }
  }

  const kommend = (fenster || []).filter((f) => f.bis >= heute());
  const vorbei = (fenster || []).filter((f) => f.bis < heute());
  const offen = (kal || []).filter((v) => !vorschlagErfasst(v, fenster || []));
  const rb = useMemo(() => {
    if (!ist?.length) return [] as (FensterRueckblick & { ausKalender: boolean })[];
    const alle: (Zeitfenster & { ausKalender: boolean })[] = [
      ...vorbei.map((f) => ({ ...f, ausKalender: false })),
      ...kalAlt.filter((v) => !vorschlagErfasst(v, vorbei)).map((v) => ({ id: 'k' + v.von, von: v.von, bis: v.bis, art: vorschlagArt(v), titel: vorschlagTitel(v), quelle: 'kalender' as const, ausKalender: true })),
    ].sort((a, b) => b.von.localeCompare(a.von));
    return alle.map((f) => { const r = rueckblick(f, ist); return r ? { ...r, ausKalender: f.ausKalender } : null; }).filter((x): x is FensterRueckblick & { ausKalender: boolean } => !!x);
  }, [ist, vorbei, kalAlt]);

  const zeitraum = (a: string, b: string) => (a === b ? fmtDateShort(a) : `${fmtDateShort(a)}–${fmtDateShort(b)}`);

  return (
    <div className="panel-raised p-4" data-testid="zeitfenster">
      <div className="label">Coach Kira · Zeitfenster</div>
      <p className="mt-1 text-2xs text-ink-dim">
        Sag Kira, wann eine Entlastung passt (Dienstreise, Familienwoche), wann kein Training möglich ist und wann ein Aktivurlaub ansteht. Der Saisonrahmen legt seine Entlastungswochen dorthin – statt stur alle 4 Wochen – und passt sich nach jeder Änderung selbst an. Gespeichert im Sheet, für alle Geräte.
      </p>

      {err && <p className="mt-2 text-xs text-amber-600 dark:text-amber-300">{err}</p>}

      {/* Eingabe */}
      <div className="mt-3 flex flex-wrap items-end gap-2 text-xs">
        <label className="flex flex-col gap-1">von<input type="date" className="input" value={von} onChange={(e) => { setVon(e.target.value); if (!bis || bis < e.target.value) setBis(e.target.value); }} /></label>
        <label className="flex flex-col gap-1">bis<input type="date" className="input" value={bis} onChange={(e) => setBis(e.target.value)} /></label>
        <label className="flex flex-col gap-1">Art
          <select className="input" value={art} onChange={(e) => setArt(e.target.value as FensterArt)}>{ARTEN.map((a) => <option key={a} value={a}>{FENSTER_LABEL[a]}</option>)}</select>
        </label>
        <label className="flex flex-col gap-1 flex-1 min-w-[160px]">Bezeichnung<input className="input" placeholder="z. B. Dienstreise Toulouse" value={titel} onChange={(e) => setTitel(e.target.value)} /></label>
        <button className="btn btn-primary text-xs px-2 py-1" disabled={busy || !von || !bis || !titel.trim() || !!err}
          onClick={() => void hinzufuegen({ id: String(Date.now()), von, bis, art, titel: titel.trim(), quelle: 'cockpit' })}>Speichern</button>
      </div>
      <p className="mt-1 text-2xs text-ink-dim">{ART_HINT[art]}. Eine Woche zählt ab 3 Tagen im Zeitfenster.</p>
      {msg && <p className="mt-2 text-2xs text-ink" data-testid="zeitfenster-msg">{msg}</p>}

      {/* Kommende Zeitfenster */}
      <div className="mt-3">
        <div className="text-2xs text-ink-dim mb-1">Geplant</div>
        {fenster === null ? <p className="text-2xs text-ink-dim">lädt …</p> : kommend.length === 0 ? <p className="text-2xs text-ink-dim">Noch keine Zeitfenster.</p> : (
          <ul className="space-y-1 text-xs" data-testid="zeitfenster-liste">
            {kommend.map((f) => (
              <li key={f.id} className="flex flex-wrap items-center gap-2">
                <span className="tnum">{zeitraum(f.von, f.bis)}</span>
                <span className="text-ink">{f.titel}</span>
                <span className={`text-2xs ${ART_CLS[f.art]}`}>{FENSTER_LABEL[f.art]}</span>
                {f.quelle === 'kalender' && <span className="text-2xs text-ink-dim">aus Kalender</span>}
                <button className="ml-auto text-2xs text-ink-dim hover:text-red-400" disabled={busy} onClick={() => void loeschen(f)}>löschen</button>
              </li>
            ))}
          </ul>
        )}
      </div>

      {/* Kalender-Vorschläge */}
      <div className="mt-3">
        <div className="text-2xs text-ink-dim mb-1">Vorschläge aus deinem Kalender (nächste 200 Tage, nur ganztägige oder lange Termine mit Reise-/Urlaubsbezug)</div>
        {kalErr ? <p className="text-2xs text-amber-600 dark:text-amber-300">{kalErr}</p> : kal === null ? <p className="text-2xs text-ink-dim">lädt …</p> : offen.length === 0 ? <p className="text-2xs text-ink-dim">Keine neuen Vorschläge.</p> : (
          <ul className="space-y-1 text-xs" data-testid="zeitfenster-vorschlaege">
            {offen.map((v) => {
              const k = v.von + v.bis;
              const a = artWahl[k] || vorschlagArt(v);
              return (
                <li key={k} className="flex flex-wrap items-center gap-2">
                  <span className="tnum">{zeitraum(v.von, v.bis)}</span>
                  <span className="text-ink">{vorschlagTitel(v)}</span>
                  <select className="input text-2xs py-0.5" value={a} onChange={(e) => setArtWahl({ ...artWahl, [k]: e.target.value as FensterArt })}>
                    {ARTEN.map((x) => <option key={x} value={x}>{FENSTER_LABEL[x]}</option>)}
                  </select>
                  <button className="btn btn-ghost text-2xs px-2 py-0.5" disabled={busy || !!err}
                    onClick={() => void hinzufuegen({ id: 'k' + v.von + Date.now(), von: v.von, bis: v.bis, art: a, titel: vorschlagTitel(v), quelle: 'kalender' })}>übernehmen</button>
                </li>
              );
            })}
          </ul>
        )}
      </div>

      {/* Rückblick */}
      <div className="mt-4">
        <div className="text-2xs text-ink-dim mb-1">Rückblick – war es eine Entlastung, und wie hat dein Körper reagiert? (gespeicherte Zeitfenster und Reise-/Urlaubszeiten aus dem Kalender seit 26.08.25)</div>
        {ist === null ? <p className="text-2xs text-ink-dim">lädt …</p> : rb.length === 0 ? <p className="text-2xs text-ink-dim">Noch keine vergangenen Zeitfenster.</p> : (
          <div className="overflow-x-auto">
            <table className="w-full text-2xs" data-testid="zeitfenster-rueckblick">
              <thead className="text-ink-dim text-left">
                <tr className="border-b border-border">
                  <th className="pb-1">Zeitraum</th><th>Was</th><th>Last (Ø/Tag)</th><th>Bereitschaft vorher → im → danach</th><th>Ruhepuls vorher → danach</th><th>Urteil</th>
                </tr>
              </thead>
              <tbody>
                {rb.map((r) => (
                  <tr key={r.f.id} className="border-b border-border">
                    <td className="pt-1 tnum whitespace-nowrap">{zeitraum(r.f.von, r.f.bis)}</td>
                    <td className="pt-1">{r.f.titel}{r.ausKalender ? <span className="text-ink-dim"> · Kalender</span> : null}</td>
                    <td className="pt-1 tnum">{f2(r.lastFenster)} statt {f2(r.lastVorher)}{r.anteil != null ? ` (${Math.round(r.anteil * 100)} %)` : ''}</td>
                    <td className="pt-1 tnum">{f2(r.readyVorher)} → {f2(r.readyIm)} → {f2(r.readyDanach)}</td>
                    <td className="pt-1 tnum">{f2(r.rhrVorher, 1)} → {f2(r.rhrDanach, 1)}</td>
                    <td className="pt-1">{r.urteil}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
