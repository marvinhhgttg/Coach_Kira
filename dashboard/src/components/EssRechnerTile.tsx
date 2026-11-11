/**
 * ESS-Rechner in der Simulation: Einheit → ESS und Ziel-ESS → Einheit (mit Variationen),
 * gelernt aus Marcs Garmin-Aktivitäten (garmin_tag). Übernahme der Last in einen Plantag.
 */
import { useEffect, useMemo, useState } from 'react';
import { useGarminTage } from '../lib/garmin';
import {
  PROFILE,
  lerneModelle,
  vorwaerts,
  ausDauer,
  ausDistanz,
  variationen,
  aehnliche,
  fmtTempo,
  parseTempo,
  fmtMin,
  parseMin,
  type ProfilId,
  type Loesung,
} from '../lib/essRechner';
import { fmtDateShort } from '../lib/format';

type Tag = { date: string; day?: string; load: number; sport: string; locked: boolean };

const fmt1 = (n?: number) => (n == null || !isFinite(n) ? '–' : n.toLocaleString('de-DE', { maximumFractionDigits: 1 }));
const fmt0 = (n?: number) => (n == null || !isFinite(n) ? '–' : Math.round(n).toLocaleString('de-DE'));
const num = (s: string) => {
  const v = Number(s.trim().replace(',', '.'));
  return s.trim() && isFinite(v) && v >= 0 ? v : undefined;
};

export function EssRechnerTile({ tage, onApply }: { tage: Tag[]; onApply: (i: number, ess: number, sport: string) => void }) {
  const { tage: garmin } = useGarminTage();
  const modelle = useMemo(() => (garmin ? lerneModelle(garmin) : []), [garmin]);
  const [pid, setPid] = useState<ProfilId>('run');
  const [richtung, setRichtung] = useState<'vor' | 'rueck'>('rueck');
  // vorwärts
  const [vDauer, setVDauer] = useState('');
  const [vKm, setVKm] = useState('');
  const [vHm, setVHm] = useState('');
  const [vHf, setVHf] = useState('');
  const [vTempo, setVTempo] = useState('');
  const [vNp, setVNp] = useState('');
  // rückwärts
  const [ziel, setZiel] = useState('150');
  const [vorgabe, setVorgabe] = useState<'dauer' | 'distanz'>('dauer');
  const [rDauer, setRDauer] = useState('60');
  const [rKm, setRKm] = useState('10');
  const [rHm, setRHm] = useState('');
  const [tagIdx, setTagIdx] = useState(1);

  const m = modelle.find((x) => x.id === pid) || null;
  const prof = PROFILE.find((p) => p.id === pid)!;
  useEffect(() => {
    if (modelle.length && !modelle.some((x) => x.id === pid)) setPid(modelle[0].id);
  }, [modelle, pid]);

  // Ziel aus Plantag übernehmen, wenn der Tag gewechselt wird
  useEffect(() => {
    const t = tage[tagIdx];
    if (t && t.load > 0) setZiel(String(Math.round(t.load)));
  }, [tagIdx]); // eslint-disable-line react-hooks/exhaustive-deps

  const vor = useMemo(() => {
    if (!m || richtung !== 'vor') return null;
    return vorwaerts(m, { min: parseMin(vDauer), km: num(vKm), hm: num(vHm), hf: num(vHf), v: parseTempo(pid, vTempo), np: num(vNp) });
  }, [m, richtung, vDauer, vKm, vHm, vHf, vTempo, vNp, pid]);

  const zielEss = num(ziel) || 0;
  const rueck = useMemo(() => {
    if (!m || richtung !== 'rueck' || zielEss <= 0) return null;
    const km = num(rKm);
    const hmEingabe = num(rHm);
    if (vorgabe === 'dauer') {
      const min = parseMin(rDauer);
      if (!min) return null;
      const hpk = hmEingabe != null && km ? hmEingabe / km : m.hpk.p50;
      return ausDauer(m, zielEss, min, hpk);
    }
    if (!km) return null;
    return ausDistanz(m, zielEss, km, hmEingabe ?? Math.round(km * m.hpk.p50));
  }, [m, richtung, zielEss, vorgabe, rDauer, rKm, rHm]);

  const vars = useMemo(() => {
    if (!m || richtung !== 'rueck' || zielEss <= 0) return [];
    const km = num(rKm);
    const hmE = num(rHm);
    const hpk = hmE != null && km ? hmE / km : m.hpk.p50;
    return variationen(m, zielEss, hpk);
  }, [m, richtung, zielEss, rKm, rHm]);

  const aehnlich = useMemo(() => {
    if (!m) return [];
    if (richtung === 'vor' && vor) return aehnliche(m, vor.werte);
    if (richtung === 'rueck' && rueck) return aehnliche(m, { min: rueck.min, hf: rueck.hf, km: rueck.km, hm: rueck.hm });
    return [];
  }, [m, richtung, vor, rueck]);

  const freieTage = tage.map((t, i) => ({ ...t, i })).filter((t) => t.i > 0 && !t.locked);
  const essAktuell = richtung === 'vor' ? vor?.ess : rueck?.ess;

  if (!garmin) return <div className="panel-raised p-4 text-sm text-ink-muted">ESS-Rechner: Garmin-Aktivitäten werden geladen…</div>;
  if (!modelle.length) return <div className="panel-raised p-4 text-sm text-ink-muted">ESS-Rechner: keine Garmin-Aktivitäten gefunden.</div>;

  const tempoLabel = prof.tempo === 'pace' ? 'Pace (min/km)' : prof.tempo === 'row' ? 'Split (/500 m)' : 'Tempo (km/h)';

  const LoesungZeile = ({ l, titel }: { l: Loesung; titel?: string }) => (
    <tr className="border-t border-border">
      {titel !== undefined && <td className="py-1.5 pr-2 font-medium">{titel}</td>}
      <td className="tnum py-1.5 pr-2">{fmtMin(l.min)}</td>
      {prof.tempo !== 'none' && <td className="tnum py-1.5 pr-2">{l.km ? `${fmt1(l.km)} km` : '–'}</td>}
      {prof.tempo !== 'none' && <td className="tnum py-1.5 pr-2">{fmtTempo(pid, l.v)}</td>}
      <td className="tnum py-1.5 pr-2">{fmt0(l.hf)} bpm</td>
      {m?.npLink && <td className="tnum py-1.5 pr-2">{fmt0(l.np)} W</td>}
      {pid !== 'bike_indoor' && pid !== 'row' && <td className="tnum py-1.5 pr-2">{fmt0(l.hm)} hm</td>}
      <td className="py-1.5 pr-2">
        <span className={`chip border text-2xs ${l.charakter === 'locker' ? 'border-green-500/50 text-green-600' : l.charakter === 'mittel' ? 'border-sky-500/50 text-sky-600' : 'border-orange-400/60 text-orange-500'}`}>{l.charakter}</span>
        {l.hinweise.length > 0 && <span className="ml-1 text-2xs text-orange-500" title={l.hinweise.join(' ')}>unsicher</span>}
      </td>
    </tr>
  );

  return (
    <div className="panel-raised p-4">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <div className="label">Coach Kira · ESS-Rechner</div>
          <p className="text-2xs text-ink-dim mt-1">
            Gelernt aus {m?.n ?? 0} {prof.label}-Einheiten (Garmin) · neuere zählen mehr · typischer Fehler ±{Math.round(((m?.fehlerExt && richtung === 'vor' ? m.fehlerExt : m?.fehler) || 0) * 100)} %
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <div className="inline-flex rounded border border-border p-0.5">
            <button className={`px-2.5 py-1 text-xs rounded ${richtung === 'rueck' ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`} onClick={() => setRichtung('rueck')}>ESS → Einheit</button>
            <button className={`px-2.5 py-1 text-xs rounded ${richtung === 'vor' ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`} onClick={() => setRichtung('vor')}>Einheit → ESS</button>
          </div>
        </div>
      </div>

      <div className="mt-3 flex flex-wrap gap-1">
        {modelle.map((x) => {
          const p = PROFILE.find((q) => q.id === x.id)!;
          return (
            <button key={x.id} className={`chip border text-xs ${x.id === pid ? 'border-accent text-accent' : 'border-border text-ink-muted'}`} onClick={() => setPid(x.id)}>
              {p.label} <span className="text-ink-dim">· {x.n}</span>
            </button>
          );
        })}
      </div>

      {richtung === 'vor' && m && (
        <div className="mt-4 grid gap-4 lg:grid-cols-[1fr_1fr]">
          <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
            <Feld label="Dauer (min oder h:mm)" value={vDauer} set={setVDauer} ph={pid === 'hike' ? 'leer = aus km/hm' : 'z. B. 60 oder 1:30'} />
            {prof.tempo !== 'none' && <Feld label="Distanz (km)" value={vKm} set={setVKm} ph="z. B. 10" />}
            {pid !== 'bike_indoor' && pid !== 'row' && <Feld label="Höhenmeter" value={vHm} set={setVHm} ph={`leer = typisch ${Math.round(m.hpk.p50)} hm/km`} />}
            <Feld label="Ø HF (bpm)" value={vHf} set={setVHf} ph={`typisch ${Math.round(m.hf.p50)}`} />
            {prof.tempo !== 'none' && pid !== 'hike' && <Feld label={tempoLabel} value={vTempo} set={setVTempo} ph={prof.tempo === 'kmh' ? 'z. B. 26' : 'z. B. 5:30'} />}
            {m.npLink && <Feld label="NP / Leistung (W)" value={vNp} set={setVNp} ph="z. B. 180" />}
          </div>
          <div className="rounded-lg border border-border p-3">
            {vor ? (
              <>
                <div className="text-2xs text-ink-dim">Geschätzte Last</div>
                <div className="flex items-baseline gap-2">
                  <span className="tnum text-3xl font-semibold text-accent">{fmt0(vor.ess)} ESS</span>
                  <span className="tnum text-xs text-ink-dim">
                    {fmt0(vor.lo)}–{fmt0(vor.hi)}
                  </span>
                </div>
                <div className="mt-2 grid grid-cols-2 gap-x-4 gap-y-1 text-xs sm:grid-cols-3">
                  <Wert k="Dauer" v={fmtMin(vor.werte.min)} g={vor.geschaetzt.includes('min')} />
                  {prof.tempo !== 'none' && <Wert k="Distanz" v={vor.werte.km ? `${fmt1(vor.werte.km)} km` : '–'} g={vor.geschaetzt.includes('km')} />}
                  {prof.tempo !== 'none' && <Wert k="Tempo" v={fmtTempo(pid, vor.werte.v)} g={vor.geschaetzt.includes('v')} />}
                  <Wert k="Ø HF" v={`${fmt0(vor.werte.hf)} bpm`} g={vor.geschaetzt.includes('hf')} />
                  {m.npLink && <Wert k="NP" v={`${fmt0(vor.werte.np)} W`} g={vor.geschaetzt.includes('np')} />}
                  {pid !== 'bike_indoor' && pid !== 'row' && <Wert k="Höhenmeter" v={`${fmt0(vor.werte.hm)} hm`} g={vor.geschaetzt.includes('hm')} />}
                </div>
                <p className="mt-1 text-2xs text-ink-dim">grau = vom Rechner ergänzt</p>
                {vor.hinweise.map((h, i) => (
                  <p key={i} className="mt-1 text-2xs text-orange-500">{h}</p>
                ))}
              </>
            ) : (
              <p className="text-sm text-ink-muted">Dauer eingeben – oder Distanz{pid === 'hike' ? ' und Höhenmeter' : ' und Tempo'}.</p>
            )}
          </div>
        </div>
      )}

      {richtung === 'rueck' && m && (
        <div className="mt-4 space-y-4">
          <div className="flex flex-wrap items-end gap-3">
            <Feld label="Ziel-ESS" value={ziel} set={setZiel} />
            <div className="flex flex-col gap-1 text-2xs text-ink-dim">
              <span>Vorgabe</span>
              <div className="inline-flex rounded border border-border p-0.5">
                <button className={`px-2.5 py-1 text-xs rounded ${vorgabe === 'dauer' ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`} onClick={() => setVorgabe('dauer')}>Dauer</button>
                {prof.tempo !== 'none' && (
                  <button className={`px-2.5 py-1 text-xs rounded ${vorgabe === 'distanz' ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`} onClick={() => setVorgabe('distanz')}>Distanz</button>
                )}
              </div>
            </div>
            {vorgabe === 'dauer' ? <Feld label="Dauer (min oder h:mm)" value={rDauer} set={setRDauer} /> : <Feld label="Distanz (km)" value={rKm} set={setRKm} />}
            {pid !== 'bike_indoor' && pid !== 'row' && (
              <>
                {vorgabe === 'dauer' && <Feld label="Distanz für hm/km (optional)" value={rKm} set={setRKm} />}
                <Feld label="Höhenmeter (optional)" value={rHm} set={setRHm} ph={`typisch ${Math.round(m.hpk.p50)} hm/km`} />
              </>
            )}
          </div>

          {rueck ? (
            <div className="rounded-lg border border-border p-3">
              <div className="text-2xs text-ink-dim">So erreichst du {fmt0(zielEss)} ESS</div>
              <div className="mt-1 flex flex-wrap items-baseline gap-x-5 gap-y-1">
                <Gross k="Dauer" v={fmtMin(rueck.min)} />
                {prof.tempo !== 'none' && <Gross k="Distanz" v={rueck.km ? `${fmt1(rueck.km)} km` : '–'} />}
                {prof.tempo !== 'none' && <Gross k={prof.tempo === 'pace' ? 'Pace' : prof.tempo === 'row' ? 'Split' : 'Tempo'} v={fmtTempo(pid, rueck.v)} />}
                <Gross k="Ø HF" v={`${fmt0(rueck.hf)} bpm`} />
                {m.npLink && <Gross k="NP" v={`${fmt0(rueck.np)} W`} />}
                {pid !== 'bike_indoor' && pid !== 'row' && <Gross k="Höhenmeter" v={`${fmt0(rueck.hm)} hm`} />}
              </div>
              {rueck.hinweise.map((h, i) => (
                <p key={i} className="mt-1 text-2xs text-orange-500">{h}</p>
              ))}
            </div>
          ) : (
            <p className="text-sm text-ink-muted">
              {vorgabe === 'distanz' ? 'Für diese Distanz passt das Ziel-ESS nicht in deinen Tempo-Bereich – Distanz oder ESS anpassen.' : 'Ziel-ESS und Dauer eingeben.'}
            </p>
          )}

          {vars.length > 0 && (
            <div>
              <div className="label mb-1">Variationen für {fmt0(zielEss)} ESS</div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="text-left text-2xs text-ink-dim">
                      <th className="py-1 pr-2 font-normal">Variante</th>
                      <th className="py-1 pr-2 font-normal">Dauer</th>
                      {prof.tempo !== 'none' && <th className="py-1 pr-2 font-normal">Distanz</th>}
                      {prof.tempo !== 'none' && <th className="py-1 pr-2 font-normal">{prof.tempo === 'pace' ? 'Pace' : prof.tempo === 'row' ? 'Split' : 'Tempo'}</th>}
                      <th className="py-1 pr-2 font-normal">Ø HF</th>
                      {m.npLink && <th className="py-1 pr-2 font-normal">NP</th>}
                      {pid !== 'bike_indoor' && pid !== 'row' && <th className="py-1 pr-2 font-normal">hm</th>}
                      <th className="py-1 pr-2 font-normal">Charakter</th>
                    </tr>
                  </thead>
                  <tbody>
                    {vars.map((v) => (
                      <LoesungZeile key={v.titel} l={v} titel={v.titel} />
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="mt-1 text-2xs text-ink-dim">
                Charakter nach deinem HF-Bereich beim {prof.label}: locker unter {Math.round(m.hf.p25)}, zügig ab {Math.round(m.hf.p75)} bpm. „unsicher“ = außerhalb deiner bisherigen Einheiten.
              </p>
            </div>
          )}
        </div>
      )}

      {aehnlich.length > 0 && (
        <div className="mt-4">
          <div className="label mb-1">Ähnliche echte Einheiten</div>
          <div className="grid gap-2 sm:grid-cols-3">
            {aehnlich.map((a) => (
              <div key={a.datum + a.min} className="rounded border border-border px-3 py-2 text-xs">
                <div className="text-2xs text-ink-dim">{fmtDateShort(a.datum)}{a.name ? ` · ${a.name}` : ''}</div>
                <div className="tnum">
                  <span className="font-semibold">{fmt0(a.ess)} ESS</span> · {fmtMin(a.min)}
                  {a.km > 0.5 ? ` · ${fmt1(a.km)} km` : ''} · {fmt0(a.hf)} bpm{a.hm ? ` · ${fmt0(a.hm)} hm` : ''}
                  {a.np ? ` · ${fmt0(a.np)} W` : ''}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {essAktuell != null && freieTage.length > 0 && (
        <div className="mt-4 flex flex-wrap items-center gap-2 border-t border-border pt-3 text-xs">
          <span className="text-ink-dim">In Plantag übernehmen:</span>
          <select className="input text-xs" value={tagIdx} onChange={(e) => setTagIdx(Number(e.target.value))}>
            {freieTage.map((t) => (
              <option key={t.date} value={t.i}>
                {t.day} {fmtDateShort(t.date)} · aktuell {Math.round(t.load)} ESS {t.sport}
              </option>
            ))}
          </select>
          <button className="btn btn-primary text-xs px-2 py-1" onClick={() => onApply(tagIdx, Math.round(essAktuell), prof.sport)}>
            {fmt0(essAktuell)} ESS {prof.sport} setzen
          </button>
        </div>
      )}
    </div>
  );
}

function Feld({ label, value, set, ph, hint }: { label: string; value: string; set: (s: string) => void; ph?: string; hint?: string }) {
return (
  <label className="flex flex-col gap-1 text-2xs text-ink-dim">
    <span>{label}</span>
    <input className="input tnum text-sm" value={value} placeholder={ph} onChange={(e) => set(e.target.value)} inputMode="decimal" />
    {hint && <span className="text-ink-dim">{hint}</span>}
  </label>
);
}

function Wert({ k, v, g }: { k: string; v: string; g: boolean }) {
  return (
    <div>
      <div className="text-2xs text-ink-dim">{k}</div>
      <div className={`tnum ${g ? 'text-ink-dim' : 'text-ink'}`}>{v}</div>
    </div>
  );
}
function Gross({ k, v }: { k: string; v: string }) {
  return (
    <div>
      <div className="text-2xs text-ink-dim">{k}</div>
      <div className="tnum text-lg font-semibold">{v}</div>
    </div>
  );
}
