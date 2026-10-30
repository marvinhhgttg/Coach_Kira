import { useEffect, useState } from 'react';
import type { ExtraItem } from '../lib/extras';
import { fetchExtras } from '../lib/extras';
import { fmtDateShort, fmtNum } from '../lib/format';

// Nutrition data interface
interface ErnahrungData {
  datum: string;
  gespeichert?: string;
  quelle?: string;
  ziele?: {
    kcal?: number;
    protein?: number;
    kohlenhydrate?: number;
    fett?: number;
  };
  last?: string;
  heute?: { kcal_in?: number | null; kcal_ziel?: number | null; protein?: number | null; protein_ziel?: number | null } | null;
  sport?: string;
  kurz?: string;
  mahlzeiten?: { wann: string; gericht: string; warum?: string }[];
  tipp?: string;
}

function parseNutrition(item: ExtraItem): ErnahrungData | null {
  const d = item.data || {};
  if (!d) return null;

  // Tolerant key reading: kohlenhydrate or carbs
  const rawKohlenhydrate = typeof d.kohlenhydrate === 'number' ? d.kohlenhydrate : null;
  const rawCarbs = typeof d.carbs === 'number' ? d.carbs : null;
  const kohlenhydrate = rawKohlenhydrate !== null ? rawKohlenhydrate : rawCarbs;

  const ziele: ErnahrungData['ziele'] = {};
  const rawKcal = typeof d.kcal === 'number' ? d.kcal : null;
  if (rawKcal !== null) ziele.kcal = rawKcal;
  else if (typeof d.ziele?.kcal === 'number') ziele.kcal = d.ziele.kcal;

  const rawProtein = typeof d.protein === 'number' ? d.protein : null;
  if (rawProtein !== null) ziele.protein = rawProtein;
  else if (typeof d.ziele?.protein === 'number') ziele.protein = d.ziele.protein;

  if (kohlenhydrate !== null) ziele.kohlenhydrate = kohlenhydrate;
  else if (typeof d.ziele?.kohlenhydrate === 'number') ziele.kohlenhydrate = d.ziele.kohlenhydrate;
  else if (typeof d.ziele?.carbs === 'number') ziele.kohlenhydrate = d.ziele.carbs;

  const rawFett = typeof d.fett === 'number' ? d.fett : null;
  if (rawFett !== null) ziele.fett = rawFett;
  else if (typeof d.ziele?.fett === 'number') ziele.fett = d.ziele.fett;

  const mahlzeiten: ErnahrungData['mahlzeiten'] = [];
  if (Array.isArray(d.mahlzeiten)) {
    for (const m of d.mahlzeiten as unknown[]) {
      const o = m as Record<string, unknown>;
      mahlzeiten.push({
        wann: String(o.wann || ''),
        gericht: String(o.gericht || ''),
        warum: typeof o.warum === 'string' ? o.warum : '',
      });
    }
  }

  return {
    datum: String(d.datum ?? item.datum),
    gespeichert: d.gespeichert || undefined,
    quelle: String(d.quelle ?? item.quelle),
    ziele: Object.keys(ziele).length > 0 ? ziele : undefined,
    last: d.last !== undefined && d.last !== null ? String(Math.round(Number(d.last))) : '',
    heute: d.heute || null,
    sport: String(d.sport || ''),
    kurz: typeof d.kurz === 'string' ? d.kurz : undefined,
    mahlzeiten: mahlzeiten.length > 0 ? mahlzeiten : undefined,
    tipp: typeof d.tipp === 'string' ? d.tipp : undefined,
  };
}

export function NutritionTile() {
  const [items, setItems] = useState<ExtraItem[]>([]);
  const [err, setErr] = useState<string | null>(null);

  // Load nutrition extra in useEffect
  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const data = await fetchExtras(['ernaehrung'], 3);
        if (!cancelled) setItems(data);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (msg.includes('JSON')) setErr('Patch „KI-Extras" noch nicht aktiv');
        else if (!cancelled) setErr('Ernährung konnte nicht geladen werden');
      }
    }

    load();
    return () => { cancelled = true; };
  }, []);

  const latestItem = items.length > 0 ? items[0] : null;
  const ernData: ErnahrungData | null = latestItem ? parseNutrition(latestItem) : null;

  return (
    <div className="panel-raised p-4">
      <div className="label">Coach Kira · Ernährung{ernData ? ` für ${fmtDateShort(ernData.datum)}` : ""}</div>

      {err && <p className="mt-2 text-xs text-orange-300">{err}</p>}

      {ernData ? (
        <>
          {/* Ziele */}
          {ernData.ziele && (
            <div className="mt-3 rounded border border-border p-2">
              <div className="text-xs font-semibold text-ink mb-1">Ziele</div>
              <div className="flex flex-wrap gap-x-4 text-xs tnum">
                {ernData.ziele.kcal !== undefined && (
                  <span>
                    <span className="text-lg">{fmtNum(ernData.ziele.kcal)}</span>
                    <span className="ml-1 text-2xs text-ink-dim">kcal</span>
                  </span>
                )}
                {ernData.ziele.protein !== undefined && (
                  <span>
                    <span className="text-lg">{fmtNum(ernData.ziele.protein)}</span>
                    <span className="ml-1 text-2xs text-ink-dim">g Protein</span>
                  </span>
                )}
                {ernData.ziele.kohlenhydrate !== undefined && (
                  <span>
                    <span className="text-lg">{fmtNum(ernData.ziele.kohlenhydrate)}</span>
                    <span className="ml-1 text-2xs text-ink-dim">g KH</span>
                  </span>
                )}
                {ernData.ziele.fett !== undefined && (
                  <span>
                    <span className="text-lg">{fmtNum(ernData.ziele.fett)}</span>
                    <span className="ml-1 text-2xs text-ink-dim">g Fett</span>
                  </span>
                )}
              </div>
            </div>
          )}

          {/* Letztes Training + Sport */}
          {(ernData.last || ernData.sport) && (
            <div className="mt-2 text-xs text-ink-dim">
              {ernData.last !== '' && <span>Training morgen: {ernData.last} ESS</span>}
              {ernData.sport && <span className="ml-2">· {ernData.sport}</span>}
            </div>
          )}

          {ernData.heute && ernData.heute.kcal_in != null && (
            <div className="mt-2 text-xs text-ink-muted tnum">
              Heute gegessen: {fmtNum(ernData.heute.kcal_in)} kcal
              {ernData.heute.kcal_ziel ? ` von ca. ${fmtNum(ernData.heute.kcal_ziel)}` : ''}
              {ernData.heute.protein != null ? ` · Protein ${fmtNum(ernData.heute.protein)} g` : ''}
              {ernData.heute.protein_ziel ? ` von ${fmtNum(ernData.heute.protein_ziel)} g` : ''}
            </div>
          )}

          {/* Kurz */}
          {ernData.kurz && (
            <div className="mt-2 rounded border border-border p-2">
              <pre className="whitespace-pre-line text-xs text-ink">{ernData.kurz}</pre>
            </div>
          )}

          {/* Mahlzeiten */}
          {ernData.mahlzeiten && ernData.mahlzeiten.length > 0 && (
            <div className="mt-3">
              <div className="text-xs font-semibold text-ink mb-1">Mahlzeiten</div>
              <div className="flex flex-col gap-2">
                {ernData.mahlzeiten.map((m, i) => (
                  <div key={i} className="rounded border border-border p-2">
                    <div className="text-xs text-ink">
                      <span className="font-semibold">{m.wann}:</span>{' '}
                      {m.gericht}
                    </div>
                    {m.warum && (
                      <p className="mt-1 text-2xs text-ink-dim">{m.warum}</p>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Tipp */}
          {ernData.tipp && (
            <div className="mt-3 rounded border border-border p-2">
              <div className="text-xs font-semibold text-ink mb-1">Tipp</div>
              <p className="text-xs text-ink">{ernData.tipp}</p>
            </div>
          )}

          <div className="mt-2 text-2xs text-ink-dim">
            Erstellt {latestItem?.gespeichert || ''} · {ernData.quelle}
          </div>
        </>
      ) : latestItem ? (
        /* Data exists but parse failed — show text */
        <div>
          <pre className="whitespace-pre-line text-xs text-ink">{latestItem.text}</pre>
          <div className="mt-2 text-2xs text-ink-dim">
            Erstellt {latestItem.gespeichert} · {latestItem.quelle}
          </div>
        </div>
      ) : (
        <p className="mt-2 text-xs text-ink-muted">
          Noch keine Empfehlung – der Mac erstellt sie abends ab 17 Uhr.
        </p>
      )}
    </div>
  );
}
