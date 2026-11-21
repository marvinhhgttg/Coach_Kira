/**
 * Coach Kira · Wochenplan – direkt von der KI (lokal coach-kira oder Cloud), ohne KI-Report im Sheet.
 * Kira plant 14 Tage, die Cockpit-Simulation prüft ACWR, Monotonie, Ruhetage und Fitness,
 * Verstöße gehen bis zu 3 Runden zurück an Kira. Ergebnis landet nur in der Simulation
 * (und wird im Browser gemerkt); gespeichert wird erst mit „In Timeline speichern“.
 */
import { useCallback, useEffect, useMemo, useState } from 'react';
import { useEngine } from '../lib/engine';
import { runPlanLoop, evaluate, localLlm, type LoopDay, type LoopSimRow } from '../lib/planLoop';
import { planChecks, type PlanRow } from '../screens/StrategyPanel';
import { fetchExtras } from '../lib/extras';
import { fmtDateShort, fmtNum } from '../lib/format';
import { WorkoutPanel } from './WorkoutPanel';
import type { PlanDay } from '../lib/api';
import { ladeLangKontext } from '../lib/briefingLocal';
import { langeEinheitenBlock } from '../lib/kiraPrompts';
import { esGesamt, essSpanne, type LangKontext } from '../lib/esLang';

const KEY = 'ck_kira_wochenplan_v1';

export type WochenplanTag = {
  date: string;
  day?: string;
  load: number;
  zone: string;
  sport: string;
  training: string;
  alt: string;
  why: string;
  locked: boolean;
  vorher: number;
  /** 21.11.2026: lange Einheit – Dauer (h) und erwarteter Einfluss auf den Endurance Score (Punkte) von Kira */
  langH?: number | null;
  esEffekt?: number | null;
};
export type Wochenplan = {
  erstellt: string;
  engine: 'local' | 'pplx';
  summary: string;
  violations: string[];
  runden: number;
  maxAcwr: number;
  finalCtl: number;
  tage: WochenplanTag[];
};

export function loadWochenplan(): Wochenplan | null {
  try {
    const j = JSON.parse(localStorage.getItem(KEY) || 'null');
    return j && Array.isArray(j.tage) ? (j as Wochenplan) : null;
  } catch {
    return null;
  }
}

const BASIS_SPORT = ['Bike', 'Run', 'Hike', 'Row', 'Strength', 'Off'];

export function KiraWochenplanTile({
  sim,
  loads,
  zones,
  sports,
  locks,
  simulateWith,
  recGap,
  onApply,
  onPlan,
}: {
  sim: PlanRow[];
  loads: number[];
  zones: string[];
  sports: string[];
  locks: boolean[];
  simulateWith: (l: number[], z: string[]) => PlanRow[];
  recGap: number;
  onApply: (loads: number[], zones: string[], sports: string[]) => void;
  onPlan?: (p: Wochenplan | null) => void;
}) {
  const [engine] = useEngine();
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState('');
  const [err, setErr] = useState<string | null>(null);
  const [plan, setPlan] = useState<Wochenplan | null>(() => loadWochenplan());
  const [openDate, setOpenDate] = useState<string | null>(null);
  const [extraCtx, setExtraCtx] = useState('');
  const [langK, setLangK] = useState<LangKontext | null>(null);

  // Nur Pläne zeigen, die zum aktuellen 14-Tage-Fenster passen
  const fenster = sim.map((r) => r.date);
  const gueltig = plan && plan.tage.some((t) => fenster.includes(t.date)) ? plan : null;
  useEffect(() => {
    onPlan?.(gueltig);
  }, [gueltig?.erstellt]); // eslint-disable-line react-hooks/exhaustive-deps

  // Lange Einheiten (Vorgabe aus dem Endurance-Score-Hebel, Stand und Termine) – vor jedem Planen frisch
  useEffect(() => {
    let off = false;
    ladeLangKontext().then((k) => { if (!off) setLangK(k); }).catch(() => {});
    return () => { off = true; };
  }, [fenster[0]]); // eslint-disable-line react-hooks/exhaustive-deps

  // Zusatzwissen aus den Extras (nur lesen): Vorausblick, Saisonrahmen, Ziele
  useEffect(() => {
    let off = false;
    fetchExtras(['prognose', 'saison', 'ziel'], 20)
      .then((items) => {
        if (off) return;
        const first = (t: string) => items.find((x) => x.typ === t);
        const z: string[] = [];
        const pg = first('prognose')?.data as any;
        if (pg) {
          const v = pg.vorausblick || {};
          if (pg.fruehwarnung) z.push(`FRÜHWARNUNG: ${pg.fruehwarnung.stufe} – ${pg.fruehwarnung.prognose || ''}`);
          if (v.halte_woche) z.push(`FITNESS HALTEN: auf Dauer ${v.halte_woche} ESS/Woche (${v.halte_tag} ESS/Tag).`);
          if (Array.isArray(v.wochen) && v.wochen.length)
            z.push('VORAUSBLICK WOCHEN: ' + v.wochen.map((w: any) => `${w.von}–${w.bis} ${w.quelle} ${w.last} ESS (${w.diff >= 0 ? '+' : ''}${w.diff} zu Fitness halten)`).join('; '));
          if (v.qualitaetstag?.tag) z.push(`BESTER QUALITÄTSTAG laut Vorausblick: ${v.qualitaetstag.tag}.`);
          const gl = pg.garmin_lage;
          if (gl && typeof gl === 'object')
            z.push(
              `GARMIN-LAGE: Body Battery Aufwachen ${gl.bb_aufwachen ?? '-'} (Schnitt ${gl.bb_schnitt ?? '-'}), Schlafbedarf ${gl.schlafbedarf_diff_min ?? '-'} min, letzte 7 Tage ${gl.std7 ?? '-'} h / ${gl.km7 ?? '-'} km / ${gl.hm7 ?? '-'} hm, Zonen Z1–Z5 ${(gl.zonen7_pct || []).join('/')} %.`,
            );
        }
        const sa = first('saison')?.data as any;
        if (sa?.wochen?.length) {
          const von = fenster[0] || '';
          const bis = fenster[fenster.length - 1] || '';
          const ws = sa.wochen.filter((w: any) => w.start && w.start <= bis && addDays(w.start, 6) >= von);
          if (ws.length) z.push('SAISONRAHMEN: ' + ws.map((w: any) => `KW ${w.kw} ab ${w.start}: ${w.phase}, Ziel ${w.ziel_ess} ESS, Sportmix ${w.sportmix || '-'}, Fokus ${w.fokus || '-'}`).join(' | '));
        }
        const ziele = items.filter((x) => x.typ === 'ziel').map((x) => (x.data as any)?.ziel).filter(Boolean);
        if (ziele.length) z.push('ZIELE: ' + ziele.map((g: any) => `${g.name} am ${g.datum} (Prio ${g.prioritaet})`).join('; '));
        setExtraCtx(z.join('\n'));
      })
      .catch(() => setExtraCtx(''));
    return () => {
      off = true;
    };
  }, [fenster[0]]); // eslint-disable-line react-hooks/exhaustive-deps

  const days: LoopDay[] = useMemo(
    () => sim.map((r, i) => ({ date: r.date, day: r.day, load: loads[i] ?? r.sim_load, zone: zones[i] ?? r.zone, sport: sports[i] ?? r.sport, locked: !!locks[i] })),
    [sim, loads, zones, sports, locks],
  );
  const sportsAllowed = useMemo(() => Array.from(new Set([...BASIS_SPORT, ...sports.filter(Boolean)])), [sports]);
  const simulate = useCallback((l: number[], z: string[]) => simulateWith(l, z) as unknown as LoopSimRow[], [simulateWith]);
  const check = useCallback(
    (rows: LoopSimRow[]) => planChecks(rows as unknown as PlanRow[], recGap).map((c) => ({ ...c, cap: c.cap === 'QUALITY' ? undefined : c.cap })),
    [recGap],
  );

  const run = useCallback(async () => {
    setBusy(true);
    setErr(null);
    try {
      const s = sim[0];
      const lk = (await ladeLangKontext().catch(() => null)) || langK; // Vorgabe evtl. gerade geändert
      if (lk) setLangK(lk);
      const langTxt = lk ? langeEinheitenBlock(lk, 'wochenplan').trim() + '\n' : '';
      const context =
        `STAND HEUTE (${s?.date || ''}): ATL ${fmtNum(s?.atl)}, CTL ${fmtNum(s?.ctl)}, ACWR ${fmtNum(s?.acwr, 2)}, Monotonie ${fmtNum(s?.monotony, 2)}.\n` +
        (extraCtx ? extraCtx + '\n' : '') +
        langTxt +
        'AUFGABE: Plane die 14 Tage als Wochenplan – passend zu Saisonrahmen, Zielen und Garmin-Lage. Wochenlast möglichst nah an Fitness halten bzw. Saison-Ziel, Entlastung wo der Körper es braucht. Sport je Tag aus der erlaubten Liste.';
      const r = await runPlanLoop({
        days,
        simulate,
        check,
        context,
        sportsAllowed,
        llm: localLlm,
        onProgress: setProgress,
      });
      const b = r.best;
      const neu: Wochenplan = {
        erstellt: new Date().toISOString(),
        engine,
        summary: b.summary,
        violations: b.violations,
        runden: r.rounds.length,
        maxAcwr: b.maxAcwr,
        finalCtl: b.finalCtl,
        tage: days.map((d, i) => ({
          date: d.date,
          day: d.day,
          load: b.loads[i],
          zone: b.zones[i],
          sport: b.sports?.[i] || d.sport,
          training: b.training?.[i] || '',
          alt: b.alt?.[i] || '',
          why: b.why[i] || '',
          locked: d.locked || i === 0,
          vorher: d.load,
          langH: b.langH?.[i] ?? null,
          esEffekt: b.esEffekt?.[i] ?? null,
        })),
      };
      setPlan(neu);
      try {
        localStorage.setItem(KEY, JSON.stringify(neu));
      } catch {
        /* egal */
      }
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
      setProgress('');
    }
  }, [sim, days, simulate, check, sportsAllowed, engine, extraCtx, langK]);

  const apply = () => {
    if (!gueltig) return;
    const by = new Map(gueltig.tage.map((t) => [t.date, t]));
    const pick = <T,>(cur: T[], f: (t: WochenplanTag) => T) => cur.map((v, i) => {
      const t = by.get(days[i]?.date);
      return !t || locks[i] || i === 0 ? v : f(t);
    });
    onApply(pick(loads, (t) => t.load), pick(zones, (t) => t.zone), pick(sports, (t) => t.sport));
  };

  // Prüfung gegen den aktuellen Stand (falls Tage inzwischen gesperrt/geändert wurden)
  const pruef = useMemo(() => {
    if (!gueltig) return null;
    const by = new Map(gueltig.tage.map((t) => [t.date, t]));
    const l = days.map((d, i) => (d.locked || i === 0 ? d.load : by.get(d.date)?.load ?? d.load));
    const z = days.map((d, i) => (d.locked || i === 0 ? d.zone : by.get(d.date)?.zone ?? d.zone));
    return evaluate(days, l, z, simulate, check);
  }, [gueltig, days, simulate, check]);

  const wochen = useMemo(() => {
    if (!gueltig) return [] as { von: string; summe: number }[];
    const out: { von: string; summe: number }[] = [];
    gueltig.tage.forEach((t) => {
      const mo = monday(t.date);
      const w = out.find((x) => x.von === mo);
      if (w) w.summe += t.load;
      else out.push({ von: mo, summe: t.load });
    });
    return out;
  }, [gueltig]);

  // Termine aus der Vorgabe (Endurance-Score-Hebel): Datum → Richtwert
  const langTermin = useMemo(
    () => new Map((langK?.termine || []).map((t) => [t.date, `Richtwert ${fmtH(t.dauerH)}${t.art === 'sehr_lang' ? ' (sehr lang)' : ''}`])),
    [langK],
  );
  // Eigene Rechnung statt Kiras Schätzung: Endurance = Dauer-Zuschlag + Lastwirkung; ESS gegen die übliche Spanne prüfen
  const langCheck = (t: WochenplanTag): { es: number | null; hinweis: string | null } | null => {
    if (t.langH == null || !langK) return null;
    const es = esGesamt(t.load, t.langH, langK.effekt).gesamt;
    const sp = essSpanne(t.langH, t.sport, langK.essProStunde);
    const teile: string[] = [];
    if (sp && (t.load < sp.von * 0.85 || t.load > sp.bis * 1.15)) teile.push(`ESS für ${fmtH(t.langH)} ${t.sport} sonst ${sp.von}–${sp.bis}`);
    if (t.esEffekt != null && Math.abs(t.esEffekt - es) > 10) teile.push(`Kira schätzte Endurance ${vzS(t.esEffekt)}, gerechnet ${vzS(es)}`);
    return { es, hinweis: teile.length ? teile.join(' · ') : null };
  };
  const openTag = gueltig?.tage.find((t) => t.date === openDate) || null;
  const openDay: PlanDay | null = openTag
    ? { date: openTag.date, day: openTag.day, recommended_load_ess: openTag.load, original_load_ess: openTag.vorher, recommended_zone: openTag.zone, training_1: `${openTag.sport}${openTag.training ? ' – ' + openTag.training : ''}`, training_2: openTag.alt, projected_effect: openTag.why }
    : null;
  const engLabel = engine === 'pplx' ? 'Cloud' : 'lokal (coach-kira)';

  return (
    <div className="panel-raised p-4 border-accent-strong/20">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Wochenplan</div>
          <p className="text-2xs text-ink-dim mt-1">
            Direkt von Kira · {gueltig ? `${gueltig.engine === 'pplx' ? 'Cloud' : 'lokal'} · ${new Date(gueltig.erstellt).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })} Uhr · ${gueltig.runden} Runde${gueltig.runden === 1 ? '' : 'n'}` : 'noch kein Plan'}
            {busy && progress ? ' · ' + progress : ''}
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {gueltig && (
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={apply} disabled={busy}>
              In Simulation übernehmen
            </button>
          )}
          <button className="btn btn-primary text-xs px-2 py-1" onClick={run} disabled={busy || !sim.length}>
            {busy ? 'Kira plant…' : `${gueltig ? 'Neu planen' : 'Plan erstellen'} · ${engLabel}`}
          </button>
        </div>
      </div>

      {err && <p className="mt-3 rounded border border-ampel-rot/40 bg-ampel-rot/10 px-3 py-2 text-xs text-red-400">{err}</p>}

      {!gueltig && !busy && !err && (
        <p className="mt-3 text-sm text-ink-muted">
          Kira plant die nächsten 14 Tage direkt – mit Fitness, Garmin-Lage, Saisonrahmen, Zielen, langen Einheiten (Vorgabe im Endurance-Score-Hebel) und gesperrten Tagen. Die Simulation prüft jeden Vorschlag
          und schickt Regelverstöße bis zu dreimal zurück. Lokal dauert das etwa 1–3 Minuten.
        </p>
      )}

      {gueltig && (
        <>
          {gueltig.summary && <p className="mt-3 text-sm leading-relaxed">{gueltig.summary}</p>}
          <div className="mt-2 flex flex-wrap gap-2 text-2xs">
            {pruef && (
              <span className={`chip border ${pruef.violations.length ? 'border-orange-400/50 text-orange-500' : 'border-green-500/50 text-green-600'}`}>
                {pruef.violations.length ? `${pruef.violations.length} Regelhinweis${pruef.violations.length === 1 ? '' : 'e'}` : 'Regeln erfüllt'}
              </span>
            )}
            {pruef && <span className="chip border border-border text-ink-muted">ACWR max {fmtNum(pruef.maxAcwr, 2)}</span>}
            {pruef && <span className="chip border border-border text-ink-muted">CTL Tag 14: {fmtNum(pruef.finalCtl)}</span>}
            {pruef && <span className="chip border border-border text-ink-muted">{pruef.restDays} Ruhetage</span>}
            {wochen.map((w) => (
              <span key={w.von} className="chip border border-border text-ink-muted">
                Woche ab {fmtDateShort(w.von)}: {fmtNum(w.summe)} ESS
              </span>
            ))}
          </div>
          {pruef && pruef.violations.length > 0 && (
            <ul className="mt-2 list-disc pl-5 text-2xs text-orange-500">
              {pruef.violations.slice(0, 5).map((v, i) => (
                <li key={i}>{v}</li>
              ))}
            </ul>
          )}
          <div className="mt-4 grid gap-2 md:grid-cols-2 xl:grid-cols-4">
            {gueltig.tage.map((t) => {
              const diff = t.load - t.vorher;
              return (
                <div key={t.date} className="rounded border border-border bg-bg-subtle/50 p-3 flex flex-col gap-1">
                  <div className="flex items-center justify-between">
                    <span className="tnum text-xs text-ink">
                      {t.day} {fmtDateShort(t.date)}
                    </span>
                    {t.locked && <span className="chip border border-border-strong text-2xs text-ink-muted">fest</span>}
                  </div>
                  <div className="tnum">
                    <span className="text-lg font-semibold text-accent">{fmtNum(t.load)} ESS</span>{' '}
                    <span className="text-2xs text-ink-dim">Plan {fmtNum(t.vorher)}</span>
                    {diff !== 0 && <span className={`text-2xs ${diff < 0 ? 'text-green-600' : 'text-orange-500'}`}> ({diff > 0 ? '+' : ''}{fmtNum(diff)})</span>}
                  </div>
                  <div className="text-2xs text-ink-muted">
                    {t.sport} · {t.zone || 'Zone offen'}
                  </div>
                  {(t.langH != null || langTermin.has(t.date)) && (
                    <div className={`text-2xs rounded px-1.5 py-0.5 self-start border ${t.langH != null ? 'border-accent-strong/50 text-accent' : 'border-orange-400/60 text-orange-500'}`}>
                      {t.langH != null
                        ? `Lange Einheit ${fmtH(t.langH)}${langCheck(t)?.es != null ? ` · Endurance ${vzS(langCheck(t)!.es as number)}` : t.esEffekt != null ? ` · Endurance ${vzS(t.esEffekt)}` : ''}`
                        : `Langer Tag laut Vorgabe – ${langTermin.get(t.date)}, von Kira nicht eingeplant`}
                    </div>
                  )}
                  {t.langH != null && langCheck(t)?.hinweis && <div className="text-2xs text-orange-500 leading-snug">{langCheck(t)!.hinweis}</div>}
                  {t.training && <div className="text-xs text-ink">{t.training}</div>}
                  {t.alt && <div className="text-2xs text-ink-dim">Alternative: {t.alt}</div>}
                  {t.why && <p className="text-2xs text-ink-muted leading-snug">{t.why}</p>}
                  {t.load > 0 && (
                    <button className="mt-1 text-2xs text-accent hover:underline self-start" onClick={() => setOpenDate(openDate === t.date ? null : t.date)}>
                      {openDate === t.date ? 'Workout schließen' : 'Workout'}
                    </button>
                  )}
                </div>
              );
            })}
          </div>
          {openDay && <WorkoutPanel key={openDay.date} day={openDay} onClose={() => setOpenDate(null)} />}
          <p className="mt-3 text-2xs text-ink-dim">
            „In Simulation übernehmen“ setzt Last, Zone und Sport der freien Tage; heute und gesperrte Tage bleiben. Der Plan liegt nur im Browser, gespeichert wird
            erst mit „In Timeline speichern“. Cloud nutzt pro Runde eine Chat-Anfrage.
          </p>
        </>
      )}
    </div>
  );
}

function vzS(x: number) {
  return `${x >= 0 ? '+' : '−'}${Math.abs(Math.round(x))}`;
}
function fmtH(h: number) {
  const m = Math.round(h * 60);
  return `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')} h`;
}
function addDays(iso: string, n: number) {
  const d = new Date(iso + 'T12:00:00');
  d.setDate(d.getDate() + n);
  return d.toLocaleDateString('sv-SE');
}
function monday(iso: string) {
  const d = new Date(iso + 'T12:00:00');
  const wd = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - wd);
  return d.toLocaleDateString('sv-SE');
}
