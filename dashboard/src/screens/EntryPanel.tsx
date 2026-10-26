import { useMemo, useRef, useState, useEffect } from 'react';
import { EreignisTile } from '../components/EreignisTile';
import { plausiCheck, PLAUSI_FELDER, FELD_LABEL, type Befund } from '../lib/plausi';
import { ApiError, fetchChartData, submitCockpitEntry, type EntryType, type SubmitEntryResponse } from '../lib/api';
import { proxySubmitEntry } from '../lib/proxy';
import { fmtDateShort } from '../lib/format';
import { Panel, StorySection } from '../components/UI';
import { fetchExtras } from '../lib/extras';

type Abschnitt = 'morgens' | 'aktivitaet' | 'abends';
/** Schlüssel in jeder Sektion: string → string | number */
type GarminVorschlag = Record<Abschnitt, Record<string, string | number> | undefined>;

type ToastFn = (kind: 'ok' | 'err' | 'info', text: string) => void;

type Field = {
  key: string;
  label: string;
  kind: 'number' | 'select' | 'hrvBand';
  min?: number;
  max?: number;
  step?: number;
  required?: boolean;
  options?: string[];
  hint?: string;
  /** automatisch berechnet, wenn leer */
  auto?: (v: Record<string, string>) => string | null;
};

const SPORTS = ['Run', 'Bike', 'Hike', 'Row', 'Swim', 'Skike', 'Ski', 'Off'];
const ZONES = ['Z1', 'Z2', 'Z3', 'Z4', 'Z5', 'Z1, Z2', 'Z2, Z3', 'Z3, Z4', 'Z4, Z5', 'Off'];
const STATUS = ['Höchstform', 'Formaufbau', 'Formerhalt', 'Erholung', 'Unproduktiv', 'Ermüdet', 'Formverlust', 'Überlastung', 'Keine Daten'];

const ratio = (a: string, b: string) => {
  const x = parseFloat(a.replace(',', '.'));
  const y = parseFloat(b.replace(',', '.'));
  return Number.isFinite(x) && Number.isFinite(y) && y > 0 ? (x / y).toFixed(2) : null;
};

const FIELDS: Record<EntryType, Field[]> = {
  Morgens: [
    { key: 'sleep_hours', label: 'Schlaf (h)', kind: 'number', min: 0, max: 14, step: 0.05, required: true, hint: 'z. B. 7,5' },
    { key: 'sleep_score_0_100', label: 'Schlafscore', kind: 'number', min: 0, max: 100, step: 1, required: true },
    { key: 'rhr_bpm', label: 'Ruhepuls (bpm)', kind: 'number', min: 25, max: 120, step: 1, required: true },
    { key: 'hrv_status', label: 'HRV (ms)', kind: 'number', min: 10, max: 200, step: 1, required: true },
    { key: 'hrv_threshholds', label: 'HRV-Normalband', kind: 'hrvBand', required: true, hint: 'unten ; oben' },
    { key: 'Garmin_Training_Readiness', label: 'Trainingsbereitschaft', kind: 'number', min: 0, max: 100, step: 1, required: true },
    { key: 'garminATL', label: 'Garmin ATL', kind: 'number', min: 0, max: 3000, step: 1, required: true },
    { key: 'garminCTL', label: 'Garmin CTL', kind: 'number', min: 0, max: 3000, step: 1, required: true },
    { key: 'garminACWR', label: 'Garmin ACWR', kind: 'number', min: 0, max: 3, step: 0.01, hint: 'leer = ATL/CTL', auto: (v) => ratio(v.garminATL || '', v.garminCTL || '') },
    { key: 'Trainingszustand', label: 'Trainingszustand', kind: 'select', options: STATUS, required: true },
  ],
  'Nach Aktivität': [
    { key: 'Sport_x', label: 'Sport', kind: 'select', options: SPORTS, required: true },
    { key: 'Zone', label: 'Zone', kind: 'select', options: ZONES, required: true },
    { key: 'load_fb_day', label: 'Belastung (Tag)', kind: 'number', min: 0, max: 1500, step: 1, required: true, hint: 'bei 2 Aktivitäten Summe' },
    { key: 'Aerobic_TE', label: 'TE aerob', kind: 'number', min: 0, max: 5, step: 0.1, required: true },
    { key: 'Anaerobic_TE', label: 'TE anaerob', kind: 'number', min: 0, max: 5, step: 0.1, required: true },
    { key: 'fbATL_obs', label: 'ATL (nach Aktivität)', kind: 'number', min: 0, max: 3000, step: 1, required: true },
    { key: 'fbCTL_obs', label: 'CTL (nach Aktivität)', kind: 'number', min: 0, max: 3000, step: 1, required: true },
    { key: 'fbACWR_obs', label: 'ACWR', kind: 'number', min: 0, max: 3, step: 0.01, hint: 'leer = ATL/CTL', auto: (v) => ratio(v.fbATL_obs || '', v.fbCTL_obs || '') },
    { key: 'garminEnduranceScore', label: 'Ausdauerwert', kind: 'number', min: 1000, max: 15000, step: 1 },
    { key: 'fb_TR_obs', label: 'Trainingsbereitschaft (nachher)', kind: 'number', min: 0, max: 100, step: 1 },
  ],
  Abends: [
    { key: 'kcal_in', label: 'kcal aufgenommen', kind: 'number', min: 0, max: 10000, step: 1, required: true },
    { key: 'kcal_out', label: 'kcal verbraucht', kind: 'number', min: 0, max: 10000, step: 1, required: true },
    {
      key: 'deficit',
      label: 'Defizit',
      kind: 'number',
      min: -10000,
      max: 10000,
      step: 1,
      hint: 'leer = verbraucht − aufgenommen',
      auto: (v) => {
        const a = parseFloat((v.kcal_in || '').replace(',', '.'));
        const b = parseFloat((v.kcal_out || '').replace(',', '.'));
        return Number.isFinite(a) && Number.isFinite(b) ? String(Math.round(b - a)) : null;
      },
    },
    { key: 'carb_g', label: 'Kohlenhydrate (g)', kind: 'number', min: 0, max: 1500, step: 1 },
    { key: 'protein_g', label: 'Eiweiß (g)', kind: 'number', min: 0, max: 1000, step: 1 },
    { key: 'fat_g', label: 'Fett (g)', kind: 'number', min: 0, max: 1000, step: 1 },
    { key: 'fiber_g', label: 'Ballaststoffe (g)', kind: 'number', min: 0, max: 300, step: 1 },
  ],
};

const TYPE_HINT: Record<EntryType, string> = {
  Morgens: 'Setzt den Tageswechsel auf heute und startet den KI-Report. Nur für heute.',
  'Nach Aktivität': 'Setzt activity_done, kopiert IST → SOLL, startet Activity Review und KI-Report. Dauert 10–30 s.',
  Abends: 'Ernährungswerte; startet den KI-Report.',
};

function localIso(d = new Date()) {
  return d.toLocaleDateString('sv-SE');
}

function validate(fields: Field[], v: Record<string, string>): string[] {
  const errs: string[] = [];
  for (const f of fields) {
    const raw = (v[f.key] || '').trim();
    if (f.kind === 'hrvBand') {
      const lo = (v[`${f.key}__lo`] || '').trim();
      const hi = (v[`${f.key}__hi`] || '').trim();
      if (!lo && !hi) {
        if (f.required) errs.push(`${f.label} fehlt`);
        continue;
      }
      const a = Number(lo.replace(',', '.'));
      const b = Number(hi.replace(',', '.'));
      if (!Number.isFinite(a) || !Number.isFinite(b) || a <= 0 || b <= a) errs.push(`${f.label}: unten < oben angeben`);
      continue;
    }
    if (!raw) {
      if (f.required && !(f.auto && f.auto(v))) errs.push(`${f.label} fehlt`);
      continue;
    }
    if (f.kind === 'number') {
      const n = Number(raw.replace(',', '.'));
      if (!Number.isFinite(n)) errs.push(`${f.label}: keine Zahl`);
      else if ((f.min != null && n < f.min) || (f.max != null && n > f.max)) errs.push(`${f.label}: ${f.min}–${f.max}`);
    }
  }
  return errs;
}

function buildPayload(fields: Field[], v: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const f of fields) {
    if (f.kind === 'hrvBand') {
      const lo = (v[`${f.key}__lo`] || '').trim();
      const hi = (v[`${f.key}__hi`] || '').trim();
      if (lo && hi) out[f.key] = `${lo};${hi}`;
      continue;
    }
    let raw = (v[f.key] || '').trim();
    if (!raw && f.auto) raw = f.auto(v) || '';
    if (raw) out[f.key] = f.kind === 'number' ? raw.replace(',', '.') : raw;
  }
  return out;
}

export function EntryPanel({
  token,
  toast,
  proxyAuthenticated,
}: {
  token: string;
  toast: ToastFn;
  proxyAuthenticated: boolean;
}) {
  const [type, setType] = useState<EntryType>('Morgens');
  const [dayOffset, setDayOffset] = useState<0 | 1>(0);
  const [values, setValues] = useState<Record<EntryType, Record<string, string>>>({
    Morgens: {},
    'Nach Aktivität': {},
    Abends: {},
  });
  const [busy, setBusy] = useState(false);
  const [confirming, setConfirming] = useState(false);
  const [result, setResult] = useState<SubmitEntryResponse | null>(null);
  const [befunde, setBefunde] = useState<Befund[]>([]);
  const histRef = useRef<Record<string, unknown>[] | null>(null);
  const [garmin, setGarmin] = useState<Record<string, GarminVorschlag>>({});
  const [garminLoaded, setGarminLoaded] = useState(false);
  const [garminTimestamp, setGarminTimestamp] = useState<Record<string, string>>({});
  const [filled, setFilled] = useState<Record<EntryType, string[]>>({ Morgens: [], 'Nach Aktivität': [], Abends: [] });

  // Garmin-Daten beim Mount, per Knopf und beim Zurückkehren in den Tab laden
  const [garminReload, setGarminReload] = useState(0);
  useEffect(() => {
    const onVis = () => { if (document.visibilityState === 'visible') setGarminReload((n) => n + 1); };
    document.addEventListener('visibilitychange', onVis);
    return () => document.removeEventListener('visibilitychange', onVis);
  }, []);
  useEffect(() => {
    let cancelled = false;
    fetchExtras<GarminVorschlag>('garmin_api', 5)
      .then((items) => {
        if (cancelled) return;
        const map: Record<string, GarminVorschlag> = {};
        const tsMap: Record<string, string> = {};
        for (const item of items) {
          map[item.datum] = item.data || {} as GarminVorschlag;
          if (item.gespeichert) {
            const m = item.gespeichert.slice(11, 16);
            if (m) tsMap[item.datum] = m;
          }
        }
        setGarmin(map);
        setGarminTimestamp(tsMap);
      })
      .catch(() => {/* still ignorieren – Patch evtl. noch nicht aktiv */})
      .finally(() => { if (!cancelled) setGarminLoaded(true); });
    return () => { cancelled = true; };
  }, [garminReload]);

  /** Garmin-Werte in leere Felder übernehmen */
  function handleGarminFill() {
    const abschnittMap: Record<EntryType, Abschnitt> = {
      Morgens: 'morgens',
      'Nach Aktivität': 'aktivitaet',
      Abends: 'abends',
    };
    const abschnitt = abschnittMap[type];
    const vorschlag = garmin[date]?.[abschnitt];
    if (!vorschlag) return;
    const fieldKeys = new Set(fields.map((f) => f.key));
    const keys: string[] = [];
    for (const [fk, fv] of Object.entries(vorschlag)) {
      if (!fieldKeys.has(fk)) continue;
      const f = fields.find((ff) => ff.key === fk);
      if (f && f.kind === 'hrvBand') {
        // Fix Cloud-Review: Band "unten;oben" auf die zwei Eingabefelder verteilen
        const [lo, hi] = String(fv).split(';').map((x) => x.trim());
        if (!lo || !hi || (v[`${fk}__lo`] || '').trim() || (v[`${fk}__hi`] || '').trim()) continue;
        set(`${fk}__lo`, lo);
        set(`${fk}__hi`, hi);
        keys.push(fk);
        continue;
      }
      const existing = (v[fk] || '').trim();
      if (existing) continue;
      if (f && f.kind === 'select' && f.options) {
        const sval = String(fv).replace(',', '.');
        if (!f.options.includes(sval)) continue;
      }
      set(fk, String(fv));
      keys.push(fk);
    }
    // Fix Cloud-Review: übernommene Felder merken (nach den set-Aufrufen)
    setFilled((prev) => ({ ...prev, [type]: keys }));
  }

  /** filled[type] bei Typ-/Datumswechsel leeren */
  function clearFilledForType(t: EntryType) {
    setFilled((prev) => ({ ...prev, [t]: [] }));
  }

  async function review() {
    // Plausibilitäts-Check gegen die letzten 30 Tage, dann Bestätigung zeigen
    let hist = histRef.current;
    if (!hist) {
      try {
        const r: any = await fetchChartData('30d', PLAUSI_FELDER);
        hist = ((r && (r.data || r.rows)) || []) as Record<string, unknown>[];
      } catch {
        hist = [];
      }
      histRef.current = hist;
    }
    const before = hist.filter((x) => String(x.date || '') < date);
    const row: Record<string, unknown> = {};
    Object.entries(payload).forEach(([k, val]) => {
      if (val !== '' && val !== undefined && val !== null) row[k] = val;
    });
    try {
      setBefunde(plausiCheck(row, before));
    } catch {
      setBefunde([]);
    }
    setConfirming(true);
  }

  const fields = FIELDS[type];
  const v = values[type];
  const date = useMemo(() => {
    const d = new Date();
    d.setDate(d.getDate() - (type === 'Morgens' ? 0 : dayOffset));
    return localIso(d);
  }, [type, dayOffset]);
  const errors = validate(fields, v);
  const payload = buildPayload(fields, v);

  // Garmin-Hilfsberechnungen
  const abschnittMap: Record<EntryType, Abschnitt> = {
    Morgens: 'morgens',
    'Nach Aktivität': 'aktivitaet',
    Abends: 'abends',
  };
  const abschnitt = abschnittMap[type];
  const garminVorschlag = garminLoaded ? garmin[date]?.[abschnitt] : undefined;
  const timeStr = garminTimestamp[date] || '';

  /** Prüft ob vorschlag mind. einen Schlüssel hat, der in FIELDS[type] vorkommt */
  const garminButtonVisible = useMemo(() => {
    if (!garminVorschlag) return false;
    const fieldKeys = new Set(fields.map((f) => f.key));
    return Object.keys(garminVorschlag).some((k) => fieldKeys.has(k));
  }, [garminVorschlag, fields]);

  /** Anzahl verfügbarer Werte im vorschlag */
  const garminCount = useMemo(() => {
    if (!garminVorschlag) return 0;
    const fieldKeys = new Set(fields.map((f) => f.key));
    return Object.keys(garminVorschlag).filter((k) => fieldKeys.has(k)).length;
  }, [garminVorschlag, fields]);

  /** Info für 'Nach Aktivität': Anzahl Aktivitäten, km, hm aus aktivitaet-Sektion */
  const aktivitaetInfo = useMemo(() => {
    if (type !== 'Nach Aktivität' || !garminVorschlag) return null;
    const ak = garminVorschlag; // Fix Cloud-Review: garminVorschlag ist bereits der Abschnitt „aktivitaet“
    const anzahl = (ak.anzahl as number | undefined);
    const kmNum = (ak.km as number | undefined);
    const hmNum = (ak.hm as number | undefined);
    return { anzahl: Number.isFinite(anzahl) ? anzahl : null, kmNum, hmNum };
  }, [garminVorschlag, type]);

  /** Anzahl übernommene Werte */
  const filledCount = filled[type]?.length || 0;

  function set(key: string, val: string) {
    setConfirming(false);
    setBefunde([]);
    setValues((s) => ({ ...s, [type]: { ...s[type], [key]: val } }));
    setFilled((prev) => ({
      ...prev,
      [type]: (prev[type] || []).filter((k) => k !== key),
    }));
  }

  function applyOffDay() {
    setConfirming(false);
    setValues((s) => ({
      ...s,
      'Nach Aktivität': {
        ...s['Nach Aktivität'],
        Sport_x: 'Off',
        Zone: 'Off',
        load_fb_day: '0',
        Aerobic_TE: '0',
        Anaerobic_TE: '0',
      },
    }));
  }

  async function submit() {
    setBusy(true);
    setResult(null);
    try {
      const res = proxyAuthenticated
        ? await proxySubmitEntry(type, date, payload)
        : await submitCockpitEntry(token.trim(), type, date, payload);
      setResult(res);
      if (!res.ok) throw new ApiError(res.error || 'Senden fehlgeschlagen');
      toast('ok', `${type} für ${fmtDateShort(date)} übertragen (Zeile ${res.row}).`);
      setValues((s) => ({ ...s, [type]: {} }));
      setFilled((prev) => ({ ...prev, [type]: [] }));
      setBefunde([]);
      histRef.current = null;
    } catch (e) {
      toast('err', `${type} nicht übertragen: ${e instanceof ApiError ? e.message : String(e)}`);
    } finally {
      setBusy(false);
      setConfirming(false);
    }
  }

  return (
    <div>
    <StorySection id="eingabe" titel="Werte eintragen" unter="Morgens, nach der Aktivität, abends – Garmin-Werte übernehmen, Rest ergänzen, prüfen und senden">
    <Panel
      title="Eingabe · statt Formular"
      right={<span className="text-2xs text-ink-dim">läuft über dieselbe onFormSubmit-Logik wie das Google-Formular</span>}
    >
      <div className="space-y-4">
        <div className="flex flex-wrap items-center gap-3">
          <div className="inline-flex bg-bg rounded border border-border p-0.5" role="tablist">
            {(Object.keys(FIELDS) as EntryType[]).map((t) => (
              <button
                key={t}
                role="tab"
                aria-selected={type === t}
                onClick={() => {
                  setType(t);
                  setConfirming(false);
                  setResult(null);
                  clearFilledForType(t);
                }}
                className={`px-3 py-1.5 text-sm rounded ${type === t ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
                data-testid={`entry-tab-${t}`}
              >
                {t}
              </button>
            ))}
          </div>
          {type !== 'Morgens' ? (
            <div className="inline-flex bg-bg rounded border border-border p-0.5">
              {([0, 1] as const).map((o) => (
                <button
                  key={o}
                  onClick={() => { setDayOffset(o); clearFilledForType(type); }}
                  className={`px-2.5 py-1 text-xs rounded ${dayOffset === o ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
                >
                  {o === 0 ? 'Heute' : 'Gestern'}
                </button>
              ))}
            </div>
          ) : null}
          <span className="text-xs text-ink-muted tnum">Datum: {fmtDateShort(date)}</span>
          {garminButtonVisible && (
            <button onClick={handleGarminFill} className="btn btn-ghost text-xs px-2 py-1" data-testid="garmin-fill">
              Garmin-Werte übernehmen
            </button>
          )}
          <button onClick={() => setGarminReload((n) => n + 1)} className="btn btn-ghost text-xs px-2 py-1" title="Garmin-Vorschlag vom Mac neu laden">
            Garmin neu laden
          </button>
          {type === 'Nach Aktivität' && (
            <button onClick={applyOffDay} className="btn btn-ghost text-xs px-2 py-1">
              Off-Tag vorbelegen
            </button>
          )}
        </div>
        <p className="text-2xs text-ink-dim">{TYPE_HINT[type]}{type !== 'Morgens' && dayOffset === 1 ? ' Für gestern werden keine Analysen gestartet.' : ''}</p>
        {garminLoaded && !garminVorschlag && type !== 'Morgens' && (
          <div className="rounded border border-orange-400/40 bg-orange-500/10 px-3 py-2 text-2xs text-orange-600 dark:text-orange-300" data-testid="garmin-missing">
            {type === 'Nach Aktivität' ? 'Noch keine Aktivität' : 'Noch keine Abendwerte'} im Garmin-Vorschlag für {fmtDateShort(date)}
            {timeStr ? ` (letzter Abruf vom Mac: ${timeStr} Uhr)` : ' (heute noch kein Abruf vom Mac)'}. Der Mac holt die Werte alle 30 Minuten – aber nur, wenn er wach ist
            (zugeklappt im Akkubetrieb schläft er). Mac aufwecken oder am Netzteil lassen, dann „Garmin neu laden“.
          </div>
        )}
        {garminVorschlag && (
          <div className="text-2xs text-ink-dim" data-testid="garmin-hint">
            Garmin-API (Stand {timeStr}): {garminCount} Werte verfügbar
            {filledCount > 0 && ` · ${filledCount} übernommen, Rest bitte ergänzen`}
            {type === 'Nach Aktivität' && aktivitaetInfo && (
              <> · {aktivitaetInfo.anzahl} Aktivität{aktivitaetInfo.anzahl !== 1 ? 'en' : ''}{aktivitaetInfo.kmNum != null && `, ${Number(aktivitaetInfo.kmNum).toFixed(1).replace('.', ',')} km`}{aktivitaetInfo.hmNum != null && `, ${Math.round(Number(aktivitaetInfo.hmNum))} hm`}</>
            )}
          </div>
        )}

        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          {fields.map((f) => {
            const autoVal = f.auto && !(v[f.key] || '').trim() ? f.auto(v) : null;
            return (
              <label key={f.key} className="block">
                <span className="label">
                  {f.label}
                  {filled[type]?.includes(f.key) && <span className="text-ink-dim"> · Garmin</span>}
                  {f.required ? ' *' : ''}
                </span>
                {f.kind === 'select' ? (
                  <select className="input w-full mt-1" value={v[f.key] || ''} onChange={(e) => set(f.key, e.target.value)}>
                    <option value="">—</option>
                    {f.options!.map((o) => (
                      <option key={o} value={o}>
                        {o}
                      </option>
                    ))}
                  </select>
                ) : f.kind === 'hrvBand' ? (
                  <div className="mt-1 flex items-center gap-1.5">
                    <input className="input w-full tnum" inputMode="numeric" placeholder="unten" value={v[`${f.key}__lo`] || ''} onChange={(e) => set(`${f.key}__lo`, e.target.value)} />
                    <span className="text-ink-dim">;</span>
                    <input className="input w-full tnum" inputMode="numeric" placeholder="oben" value={v[`${f.key}__hi`] || ''} onChange={(e) => set(`${f.key}__hi`, e.target.value)} />
                  </div>
                ) : (
                  <input
                    className="input w-full mt-1 tnum"
                    inputMode="decimal"
                    placeholder={autoVal ? `auto: ${autoVal.replace('.', ',')}` : f.hint || ''}
                    value={v[f.key] || ''}
                    onChange={(e) => set(f.key, e.target.value)}
                    data-testid={`entry-${f.key}`}
                  />
                )}
              </label>
            );
          })}
        </div>

        {errors.length > 0 && (
          <div className="text-xs text-yellow-300">Noch offen: {errors.join(' · ')}</div>
        )}

        <div className="flex flex-wrap items-center gap-3 border-t border-border pt-3">
          {!confirming ? (
            <button
              className="btn btn-primary px-4 py-1.5 text-sm"
              disabled={busy || errors.length > 0}
              onClick={review}
              data-testid="entry-review"
            >
              Prüfen &amp; senden
            </button>
          ) : (
            <>
              {befunde.length > 0 && (
                <div className="w-full rounded border border-orange-400/40 bg-orange-400/10 p-3 text-xs" data-testid="plausi-box">
                  <div className="font-semibold text-ink mb-1">Bitte kurz prüfen</div>
                  <ul className="flex flex-col gap-0.5">
                    {befunde.map((b, i) => (
                      <li key={i} className={b.schwere === 'warnung' ? 'text-orange-300' : 'text-ink-muted'}>
                        {FELD_LABEL[b.feld] || b.feld}: {b.wert} – {b.grund}
                        {b.erwartet ? ` (erwartet ${b.erwartet})` : ''}
                      </li>
                    ))}
                  </ul>
                  <div className="mt-1 text-2xs text-ink-dim">Stimmt alles, kannst du trotzdem übertragen; sonst Wert korrigieren.</div>
                </div>
              )}
              <span className="text-xs text-ink-muted">
                {type} · {fmtDateShort(date)} ·{' '}
                <span className="tnum">
                  {Object.entries(payload)
                    .map(([k, val]) => `${k}=${val}`)
                    .join(', ')}
                </span>
              </span>
              <button className="btn btn-primary px-4 py-1.5 text-sm" disabled={busy} onClick={submit} data-testid="entry-submit">
                {busy ? 'Wird übertragen…' : befunde.length ? 'Trotzdem übertragen' : 'Jetzt übertragen'}
              </button>
              <button className="btn btn-ghost px-3 py-1.5 text-sm" disabled={busy} onClick={() => { setConfirming(false); setBefunde([]); }}>
                {befunde.length ? 'Korrigieren' : 'Abbrechen'}
              </button>
            </>
          )}
        </div>

        {result && (
          <div
            className={`rounded border px-3 py-2 text-xs ${result.ok ? 'border-ampel-gruen/40 bg-ampel-gruen/10' : 'border-ampel-rot/40 bg-ampel-rot/10 text-red-300'}`}
          >
            {result.ok ? (
              <>
                <div className="font-semibold text-green-300">
                  Übertragen: {result.type} · {fmtDateShort(result.date)} · timeline Zeile {result.row}
                  {result.durationMs != null ? ` · ${(result.durationMs / 1000).toFixed(1)} s` : ''}
                </div>
                <div className="mt-1 tnum text-ink-muted">
                  {Object.entries(result.readback || {})
                    .map(([k, val]) => `${k}: ${String(val)}`)
                    .join(' · ')}
                </div>
                <div className="mt-1 text-ink-dim">
                  {result.triggersStarted ? 'Analysen (KI-Report etc.) wurden angestoßen.' : 'Vergangener Tag: keine Analysen gestartet.'}
                </div>
              </>
            ) : (
              <>Fehler: {result.error}</>
            )}
          </div>
        )}
      </div>
    </Panel>
    </StorySection>
    <StorySection id="ereignisse" titel="Besondere Tage" unter="Krankheit, Reise, Höhe, Hitze, Stress – damit Kira solche Tage richtig einordnet">
      <EreignisTile />
    </StorySection>
    </div>
  );
}
