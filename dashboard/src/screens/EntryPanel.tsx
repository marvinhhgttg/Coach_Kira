import { useMemo, useState } from 'react';
import { ApiError, submitCockpitEntry, type EntryType, type SubmitEntryResponse } from '../lib/api';
import { proxySubmitEntry } from '../lib/proxy';
import { fmtDateShort } from '../lib/format';
import { Panel } from '../components/UI';

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

  const fields = FIELDS[type];
  const v = values[type];
  const date = useMemo(() => {
    const d = new Date();
    d.setDate(d.getDate() - (type === 'Morgens' ? 0 : dayOffset));
    return localIso(d);
  }, [type, dayOffset]);
  const errors = validate(fields, v);
  const payload = buildPayload(fields, v);
  const canAuth = proxyAuthenticated || !!token.trim();

  function set(key: string, val: string) {
    setConfirming(false);
    setValues((s) => ({ ...s, [type]: { ...s[type], [key]: val } }));
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
    if (!canAuth) {
      toast('err', 'Zum Senden bitte mit PIN anmelden oder Token eingeben (Command Center).');
      return;
    }
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
    } catch (e) {
      toast('err', `${type} nicht übertragen: ${e instanceof ApiError ? e.message : String(e)}`);
    } finally {
      setBusy(false);
      setConfirming(false);
    }
  }

  return (
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
                  onClick={() => setDayOffset(o)}
                  className={`px-2.5 py-1 text-xs rounded ${dayOffset === o ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`}
                >
                  {o === 0 ? 'Heute' : 'Gestern'}
                </button>
              ))}
            </div>
          ) : null}
          <span className="text-xs text-ink-muted tnum">Datum: {fmtDateShort(date)}</span>
          {type === 'Nach Aktivität' && (
            <button onClick={applyOffDay} className="btn btn-ghost text-xs px-2 py-1">
              Off-Tag vorbelegen
            </button>
          )}
        </div>
        <p className="text-2xs text-ink-dim">{TYPE_HINT[type]}{type !== 'Morgens' && dayOffset === 1 ? ' Für gestern werden keine Analysen gestartet.' : ''}</p>

        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          {fields.map((f) => {
            const autoVal = f.auto && !(v[f.key] || '').trim() ? f.auto(v) : null;
            return (
              <label key={f.key} className="block">
                <span className="label">
                  {f.label}
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
              onClick={() => setConfirming(true)}
              data-testid="entry-review"
            >
              Prüfen &amp; senden
            </button>
          ) : (
            <>
              <span className="text-xs text-ink-muted">
                {type} · {fmtDateShort(date)} ·{' '}
                <span className="tnum">
                  {Object.entries(payload)
                    .map(([k, val]) => `${k}=${val}`)
                    .join(', ')}
                </span>
              </span>
              <button className="btn btn-primary px-4 py-1.5 text-sm" disabled={busy} onClick={submit} data-testid="entry-submit">
                {busy ? 'Wird übertragen…' : 'Jetzt übertragen'}
              </button>
              <button className="btn btn-ghost px-3 py-1.5 text-sm" disabled={busy} onClick={() => setConfirming(false)}>
                Abbrechen
              </button>
            </>
          )}
          {!canAuth && <span className="text-2xs text-ink-dim">Zum Senden mit PIN anmelden oder Token eingeben.</span>}
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
  );
}
