import { useEffect, useMemo, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { deleteExtra, fetchExtras, fetchLoadFactors, saveExtra, type ExtraItem, type LoadFactors } from '../lib/extras';
import { localChat, parseJsonLoose } from '../lib/localLLM';
import { fmtDateShort } from '../lib/format';
import { zielBereitschaft, type Ziel } from '../lib/planning';

interface Block { titel: string; von: string; bis: string; wochen_ess: number; schluesseleinheiten: string[] }
interface ZielDaten { ziel: Ziel; plan?: { bloecke: Block[]; hinweise: string[]; tapering?: string; erstellt?: string } }

const ARTEN: { v: Ziel['art']; l: string }[] = [
  { v: 'bergtour', l: 'Bergtour' },
  { v: 'bike_hike', l: 'Bike & Hike' },
  { v: 'rad', l: 'Rad/Gravel' },
  { v: 'lauf', l: 'Lauf' },
  { v: 'skitour', l: 'Skitour' },
  { v: 'sonstiges', l: 'Sonstiges' },
];
const artLabel = (a: string) => ARTEN.find((x) => x.v === a)?.l || a;
const leer: Ziel = { datum: '', name: '', art: 'bergtour', hm: 1500, dauer_h: 6, distanz_km: 15, prioritaet: 'A', notiz: '' };
const heute = () => new Date().toLocaleDateString('sv-SE');
const n = (v: unknown) => (v === '' || v == null ? null : Number(String(v).replace(',', '.')));

export function GoalPlannerTile() {
  const [items, setItems] = useState<ExtraItem<ZielDaten>[]>([]);
  const [err, setErr] = useState<string | null>(null);
  const [form, setForm] = useState<Ziel | null>(null);
  const [sel, setSel] = useState<string | null>(null);
  const [hist, setHist] = useState<{ date: string; load: number | null; elev: number | null; ctl: number | null }[]>([]);
  const [fakt, setFakt] = useState<LoadFactors | undefined>(undefined);
  const [busy, setBusy] = useState<string | null>(null);

  async function reload() {
    try {
      const xs = await fetchExtras<ZielDaten>(['ziel'], 60);
      xs.sort((a, b) => (a.datum < b.datum ? -1 : 1));
      setItems(xs);
      setErr(null);
      if (!sel) setSel(xs.find((x) => x.datum >= heute())?.datum || null);
    } catch (e) {
      setErr(/JSON|Typ|Failed to fetch|Load failed/.test(String(e)) ? 'Patch „Zukunftsplanung“ noch nicht aktiv.' : String(e instanceof Error ? e.message : e));
    }
  }
  useEffect(() => {
    reload();
    fetchChartData('365d' as any, ['load_fb_day', 'elev_m_day'])
      .then((r: any) => setHist(((r && r.data) || []).map((x: any) => ({ date: x.date, load: n(x.load_fb_day), elev: n(x.elev_m_day), ctl: null }))))
      .catch(() => setHist([]));
    fetchLoadFactors().then(setFakt).catch(() => setFakt(undefined));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const aktiv = items.filter((x) => x.datum >= heute());
  const vergangen = items.filter((x) => x.datum < heute());
  const cur = items.find((x) => x.datum === sel) || null;
  const bereit = useMemo(() => (cur?.data?.ziel && hist.length ? zielBereitschaft(cur.data.ziel, hist, fakt) : null), [cur, hist, fakt]);

  async function onSave() {
    if (!form || !form.name.trim() || !/^\d{4}-\d{2}-\d{2}$/.test(form.datum)) {
      setErr('Bitte Name und Datum angeben.');
      return;
    }
    setBusy('Speichere…');
    try {
      const txt = `${form.name} (${artLabel(form.art)}, Priorität ${form.prioritaet}): ${form.hm} Hm, ${form.dauer_h} h, ${form.distanz_km} km${form.notiz ? ` – ${form.notiz}` : ''}`;
      const alt = items.find((x) => x.datum === form.datum)?.data;
      await saveExtra('ziel', form.datum, txt, { ziel: form, plan: alt?.ziel?.name === form.name ? alt?.plan : undefined });
      setForm(null);
      setSel(form.datum);
      await reload();
    } catch (e) {
      setErr(String(e instanceof Error ? e.message : e));
    } finally {
      setBusy(null);
    }
  }

  async function onDelete(datum: string) {
    if (!window.confirm('Ziel wirklich löschen?')) return;
    try {
      await deleteExtra('ziel', datum);
      if (sel === datum) setSel(null);
      await reload();
    } catch (e) {
      setErr(String(e instanceof Error ? e.message : e));
    }
  }

  async function onPlan() {
    if (!cur?.data?.ziel || !bereit) return;
    const z = cur.data.ziel;
    setBusy('Kira plant die Vorbereitung…');
    setErr(null);
    try {
      const wochen = Math.max(1, bereit.wochen);
      const user =
        `Ziel: ${z.name} am ${z.datum} (${artLabel(z.art)}), ${z.hm} Hm, ${z.dauer_h} h, ${z.distanz_km} km, Priorität ${z.prioritaet}. ${z.notiz || ''}\n` +
        `Geschätzte Tourlast: ${bereit.tourLast} ESS. Noch ${wochen} Wochen.\n` +
        `Stand: ${bereit.punkte.join(' | ')}\n` +
        `Aktuelle Ø-Wochenlast: ${bereit.ctl ?? '?'} ESS. Marc trainiert Rad (Pendeln, Gravel), Lauf, Wandern, im Winter Skitour/Langlauf/Skike.\n\n` +
        'Erstelle einen Vorbereitungsplan in 2–5 Blöcken bis zum Ziel (Datum von/bis je Block, YYYY-MM-DD), mit Ziel-Wochenlast (ESS) und 2–4 Schlüsseleinheiten je Block ' +
        '(z. B. lange Tage mit Höhenmetern, Back-to-Back-Tage, Tempo bergauf). Steigerung höchstens 5–8 % pro Woche, jede 4. Woche Entlastung, letzte 7–10 Tage Tapering. ' +
        'Antworte nur mit JSON: {"bloecke":[{"titel":"Grundlage","von":"2026-11-02","bis":"2026-11-29","wochen_ess":950,"schluesseleinheiten":["…"]}],"tapering":"1 Satz","hinweise":["…"]}';
      const t = await localChat(
        [
          { role: 'system', content: 'Du bist Coach Kira, Marcs Trainingscoach. Deutsch, du-Form, keine Emojis. Antworte nur mit JSON.' },
          { role: 'user', content: user },
        ],
        { temperature: 0.3, numCtx: 8192 },
      );
      const j: any = parseJsonLoose<any>(t) || {};
      const bl: any[] = Array.isArray(j.bloecke) ? j.bloecke : Array.isArray(j.blocks) ? j.blocks : [];
      const bloecke: Block[] = bl.map((b) => ({
        titel: String(b.titel ?? b.name ?? 'Block'),
        von: String(b.von ?? b.start ?? '').slice(0, 10),
        bis: String(b.bis ?? b.ende ?? '').slice(0, 10),
        wochen_ess: Math.round(Number(b.wochen_ess ?? b.wochenlast ?? 0)) || 0,
        schluesseleinheiten: (Array.isArray(b.schluesseleinheiten) ? b.schluesseleinheiten : Array.isArray(b.einheiten) ? b.einheiten : []).map(String),
      }));
      if (!bloecke.length) throw new Error('Kira hat keinen verwertbaren Plan geliefert – bitte noch einmal versuchen.');
      const plan = { bloecke, hinweise: (Array.isArray(j.hinweise) ? j.hinweise : []).map(String), tapering: j.tapering ? String(j.tapering) : '', erstellt: new Date().toISOString() };
      await saveExtra('ziel', z.datum, cur.text, { ziel: z, plan });
      await reload();
    } catch (e) {
      setErr(String(e instanceof Error ? e.message : e));
    } finally {
      setBusy(null);
    }
  }

  const ampelCls = (a?: string) => (a === 'gruen' ? 'text-green-300' : a === 'gelb' ? 'text-yellow-300' : 'text-orange-300');
  const setF = (k: keyof Ziel, v: any) => setForm((f) => (f ? { ...f, [k]: v } : f));

  return (
    <div className="panel-raised p-4">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · Ziele &amp; Touren</div>
          <p className="text-2xs text-ink-dim mt-1">Trag deine Zieltouren ein – Kira prüft, wie weit du bist, und plant die Vorbereitung. Ziele fließen in Saisonrahmen und KI-Report ein.</p>
        </div>
        <button className="btn btn-primary text-xs px-2 py-1" onClick={() => setForm({ ...leer })}>Neues Ziel</button>
      </div>
      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}

      {form && (
        <div className="mt-3 grid gap-2 sm:grid-cols-2 lg:grid-cols-4 text-xs rounded border border-border p-3">
          <label className="flex flex-col gap-1 lg:col-span-2">Name<input className="input" value={form.name} onChange={(e) => setF('name', e.target.value)} placeholder="z. B. Bike & Hike Hochvogel" /></label>
          <label className="flex flex-col gap-1">Datum<input className="input" type="date" value={form.datum} onChange={(e) => setF('datum', e.target.value)} /></label>
          <label className="flex flex-col gap-1">Art
            <select className="input" value={form.art} onChange={(e) => setF('art', e.target.value)}>{ARTEN.map((a) => <option key={a.v} value={a.v}>{a.l}</option>)}</select>
          </label>
          <label className="flex flex-col gap-1">Höhenmeter<input className="input" type="number" value={form.hm} onChange={(e) => setF('hm', Number(e.target.value))} /></label>
          <label className="flex flex-col gap-1">Dauer (h)<input className="input" type="number" step="0.5" value={form.dauer_h} onChange={(e) => setF('dauer_h', Number(e.target.value))} /></label>
          <label className="flex flex-col gap-1">Distanz (km)<input className="input" type="number" value={form.distanz_km} onChange={(e) => setF('distanz_km', Number(e.target.value))} /></label>
          <label className="flex flex-col gap-1">Priorität
            <select className="input" value={form.prioritaet} onChange={(e) => setF('prioritaet', e.target.value)}>
              <option value="A">A – Saisonhöhepunkt (mit Tapering)</option><option value="B">B – wichtig</option><option value="C">C – Trainingsziel</option>
            </select>
          </label>
          <label className="flex flex-col gap-1 sm:col-span-2 lg:col-span-4">Notiz<input className="input" value={form.notiz || ''} onChange={(e) => setF('notiz', e.target.value)} placeholder="Gelände, Begleitung, Besonderheiten" /></label>
          <div className="flex gap-2 sm:col-span-2 lg:col-span-4">
            <button className="btn btn-primary text-xs px-2 py-1" onClick={onSave} disabled={!!busy}>{busy || 'Speichern'}</button>
            <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => setForm(null)}>Abbrechen</button>
          </div>
        </div>
      )}

      {aktiv.length === 0 && !form && !err && <p className="mt-3 text-sm text-ink-muted">Noch kein Ziel eingetragen.</p>}
      {aktiv.length > 0 && (
        <div className="mt-3 flex flex-wrap gap-2">
          {aktiv.map((x) => (
            <button key={x.datum} onClick={() => setSel(x.datum)} className={`rounded border px-2 py-1 text-xs text-left ${sel === x.datum ? 'border-accent text-ink bg-bg-subtle' : 'border-border text-ink-muted'}`}>
              <span className="font-semibold">{x.data?.ziel?.prioritaet}</span> · {x.data?.ziel?.name} · {fmtDateShort(x.datum)}
            </button>
          ))}
        </div>
      )}

      {cur?.data?.ziel && (
        <div className="mt-3 rounded border border-border p-3 flex flex-col gap-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="text-sm text-ink">
              {cur.data.ziel.name} · {fmtDateShort(cur.datum)} · {artLabel(cur.data.ziel.art)} · {cur.data.ziel.hm} Hm · {cur.data.ziel.dauer_h} h
            </div>
            <div className="flex gap-2">
              <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => setForm({ ...cur.data!.ziel })}>Bearbeiten</button>
              <button className="btn btn-ghost text-xs px-2 py-1" onClick={() => onDelete(cur.datum)}>Löschen</button>
            </div>
          </div>
          {bereit && (
            <div className="text-xs">
              <div className={`font-semibold ${ampelCls(bereit.ampel)}`}>
                {bereit.ampel === 'gruen' ? 'Auf Kurs' : bereit.ampel === 'gelb' ? 'Machbar, Vorbereitung nötig' : 'Knapp – gezielt aufbauen'} · noch {bereit.wochen} Wochen · Tourlast ca. {bereit.tourLast} ESS
              </div>
              <ul className="mt-1 list-disc pl-5 text-ink-muted">{bereit.punkte.map((p, i) => <li key={i}>{p}</li>)}</ul>
            </div>
          )}
          <div className="flex gap-2 items-center">
            <button className="btn btn-primary text-xs px-2 py-1" onClick={onPlan} disabled={!!busy || !bereit}>
              {busy || (cur.data.plan ? 'Vorbereitung neu planen' : 'Vorbereitung planen')}
            </button>
            {cur.data.plan?.erstellt && <span className="text-2xs text-ink-dim">Plan vom {fmtDateShort(cur.data.plan.erstellt.slice(0, 10))}</span>}
          </div>
          {cur.data.plan && (
            <div className="grid gap-2 md:grid-cols-2">
              {cur.data.plan.bloecke.map((b, i) => (
                <div key={i} className="rounded border border-border bg-bg-subtle/40 p-2 text-xs">
                  <div className="text-ink font-semibold">{b.titel}</div>
                  <div className="text-2xs text-ink-dim tnum">{b.von && fmtDateShort(b.von)} – {b.bis && fmtDateShort(b.bis)} · ca. {b.wochen_ess} ESS/Woche</div>
                  <ul className="mt-1 list-disc pl-4 text-ink-muted">{b.schluesseleinheiten.map((s, k) => <li key={k}>{s}</li>)}</ul>
                </div>
              ))}
              {(cur.data.plan.tapering || cur.data.plan.hinweise.length > 0) && (
                <div className="md:col-span-2 text-xs text-ink-muted">
                  {cur.data.plan.tapering && <p>Tapering: {cur.data.plan.tapering}</p>}
                  {cur.data.plan.hinweise.map((h, i) => <p key={i}>– {h}</p>)}
                </div>
              )}
            </div>
          )}
        </div>
      )}
      {vergangen.length > 0 && <p className="mt-2 text-2xs text-ink-dim">Vergangen: {vergangen.map((x) => `${x.data?.ziel?.name} (${fmtDateShort(x.datum)})`).join(' · ')}</p>}
    </div>
  );
}
