import { useEffect, useState } from 'react';
import {
  ladeTage,
  reaktionen,
  zonenKlasse,
  type ETag,
} from '../lib/recoveryPrice';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

function fmtDe(v: number): string {
  return v.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function fmtDeInt(v: number): string {
  return v.toLocaleString('de-DE');
}

function fmtSigned(v: number): string {
  const r = fmtDe(Math.abs(v));
  return v >= 0 ? '+' + r : '−' + r;
}

function dateShortFromISO(iso: string): string {
  const dp = iso.slice(0, 10);
  const p = dp.split('-');
  if (p.length < 3) return dp;
  const dow = WD[new Date(dp + 'T12:00:00').getDay()];
  return `${dow} ${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

function fmtDateShort(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  return `${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

// ── Komponente ───────────────────────────────────────────────────────────
export function GrauzoneTile() {
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

  const g = reaktionen.length > 0 ? { faelle: [], lockerGeplant: 0, proWoche: [] } : null;
  const reaktionenList = reaktionen(tage);

  // Grauzone-Struktur selbst berechnen (direkt aus tage)
  const faelle: { date: string; planZone: string; zone: string; load: number; dR: number | null }[] = [];
  let lockerGeplantCount = 0;

  // Map für dR-Berechnung
  const tagsByDate = new Map<string, number>();
  for (let i = 0; i < tage.length; i++) {
    tagsByDate.set(tage[i].date, i);
  }

  function calcDR(date: string): number | null {
    const idx = tagsByDate.get(date);
    if (idx == null || idx >= tage.length - 1) return null;
    const t = tage[idx];
    if (t.load == null || t.load <= 30) return null;
    const tNext = tage[idx + 1];
    if (!tNext || tNext.krank) return null;
    const rNext = tNext.readiness;
    if (rNext == null) return null;
    const basisReadiness: number[] = [];
    for (let j = idx - 27; j <= idx; j++) {
      if (j < 0 || j >= tage.length) continue;
      const tRef = tage[j];
      if (tRef.krank) continue;
      if (tRef.readiness == null) continue;
      basisReadiness.push(tRef.readiness);
    }
    if (basisReadiness.length < 14) return null;
    const basis = basisReadiness.reduce((a, b) => a + b, 0) / basisReadiness.length;
    return rNext - basis;
  }

  // Wochen-Mapping für letzte 8 Wochen
  interface GrauWoche {
    woche: string; // ISO-Woche als 'YYYY-MM-DD' (Montag)
    geplantLocker: number;
    grau: number;
  }

  const proWocheMap = new Map<string, GrauWoche>();

  for (let w = 7; w >= 0; w--) {
    const montag = new Date();
    montag.setDate(montag.getDate() - montag.getDay() + 1 - w * 7);
    const montagStr = montag.toLocaleDateString('sv-SE');
    proWocheMap.set(montagStr, { woche: montagStr, geplantLocker: 0, grau: 0 });
  }

  function dateInWeek(dateStr: string): string | null {
    const parts = dateStr.split('-');
    if (parts.length !== 3) return null;
    const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    if (isNaN(d.getTime())) return null;
    const dayOfWeek = d.getDay();
    const montagOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    const montag = new Date(d);
    montag.setDate(montag.getDate() + montagOffset);
    return montag.toLocaleDateString('sv-SE');
  }

  for (const t of tage) {
    const klasse = zonenKlasse(t.planZone);
    if (klasse === 'locker' && t.load != null && t.load > 30) {
      lockerGeplantCount++;
      const zKl = zonenKlasse(t.zone);
      const istGrau = zKl === 'grau' || zKl === 'hart';
      if (istGrau) {
        faelle.push({ date: t.date, planZone: t.planZone, zone: t.zone, load: t.load, dR: calcDR(t.date) });
      }
      const woek = dateInWeek(t.date);
      const w = woek ? proWocheMap.get(woek) : undefined;
      if (w) {
        w.geplantLocker++;
        if (istGrau) w.grau++;
      }
    }
  }

  const proWoche = Array.from(proWocheMap.values()).sort((a, b) => a.woche.localeCompare(b.woche));
  const letzteWoche = proWoche.length > 0 ? proWoche[proWoche.length - 1] : { geplantLocker: 0, grau: 0 };

  // Prozent berechnen
  const prozent = lockerGeplantCount > 0 ? Math.round(faelle.length / lockerGeplantCount * 100) : 0;

  // Ø dR der Fälle mit dR != null
  const drFaelle = faelle.filter(f => f.dR !== null).map(f => f.dR! as number);
  const avgDRFaelle = drFaelle.length > 0 ? drFaelle.reduce((a, b) => a + b, 0) / drFaelle.length : NaN;

  // Ø dR aller reaktionen mit klasse 'locker'
  const drLocker = reaktionenList.filter(r => r.klasse === 'locker').map(r => r.dR);
  const avgDRLocker = drLocker.length > 0 ? drLocker.reduce((a, b) => a + b, 0) / drLocker.length : NaN;

  // maxGeplant für Balken-Charts
  const maxGeplant = proWoche.length > 0 ? Math.max(...proWoche.map(w => w.geplantLocker)) : 1;

  // Letzte 5 Fälle (neueste zuerst)
  const faelleList = [...faelle].reverse().slice(0, 5);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Grauzonen-Detektor lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  const geplantLocker = letzteWoche.geplantLocker;
  const grau = letzteWoche.grau;

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Grauzonen-Detektor</div>
        <p className="mt-1 text-2xs text-ink-muted">
          Einheiten, die locker (Z1/Z2) geplant waren und mit Z3-Anteil endeten.
        </p>
      </div>

      {/* Zusammenfassung */}
      <div className="mt-2 text-xs">
        <p>
          {faelle.length} von {lockerGeplantCount} locker geplanten Einheiten ({fmtDeInt(prozent)} %) sind in die Grauzone gerutscht.
        </p>

        {drFaelle.length > 0 && drLocker.length > 0 ? (
          <p className="mt-1">
            Am Folgemorgen im Schnitt {fmtSigned(avgDRFaelle)} statt {fmtSigned(avgDRLocker)} Punkte.
          </p>
        ) : drFaelle.length > 0 ? (
          <p className="mt-1">
            Am Folgemorgen im Schnitt {fmtSigned(avgDRFaelle)} Punkte.
          </p>
        ) : null}

        <p className={`mt-1 ${grau >= 2 ? 'text-yellow-400' : ''}`}>
          Diese Woche: {grau} von {geplantLocker} lockeren Einheiten in der Grauzone.
        </p>
      </div>

      {/* 8-Wochen-Balken */}
      <div className="mt-3">
        <span className="label">Letzte 8 Wochen</span>
        <div className="flex gap-1 mt-2">
          {proWoche.map((w, i) => {
            const heightPercent = maxGeplant > 0 ? (w.geplantLocker / maxGeplant) * 100 : 0;
            const grauHeightPercent = w.geplantLocker > 0 ? (w.grau / maxGeplant) * 100 : 0;
            return (
              <div key={i} className="flex-1 flex flex-col items-center">
                <div className="relative w-full" style={{ height: '48px' }}>
                  {/* Geplanteil */}
                  <div
                    className="absolute bottom-0 w-full bg-ink-dim/30 rounded-t"
                    style={{ height: `${Math.max(heightPercent, 1)}%` }}
                  />
                  {/* Grauanteil */}
                  {w.grau > 0 && (
                    <div
                      className="absolute bottom-0 w-full bg-yellow-400/70 rounded-t"
                      style={{ height: `${grauHeightPercent}%` }}
                    />
                  )}
                </div>
                <span className="text-2xs text-ink-dim mt-1">
                  {fmtDateShort(w.woche)}
                </span>
              </div>
            );
          })}
        </div>
      </div>

      {/* Liste der letzten 5 Fälle */}
      {faelleList.length > 0 && (
        <div className="mt-3">
          <span className="label">Letzte Fälle</span>
          <div className="mt-2 space-y-1">
            {faelleList.map((f, i) => (
              <div key={i} className="text-xs flex items-baseline gap-2">
                <span className="text-ink">{dateShortFromISO(f.date)}</span>
                <span className="text-ink-muted">geplant {f.planZone} → {f.zone}</span>
                <span className="tnum">{f.load} ESS</span>
                <span className={f.dR != null && f.dR >= 0 ? 'text-green-300' : f.dR != null && f.dR >= -8 ? 'text-yellow-300' : 'text-orange-300'}>
                  {f.dR != null ? fmtSigned(f.dR) : '–'}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Keine Fälle */}
      {faelle.length === 0 && (
        <div className="mt-2 text-xs text-ink-muted">
          Keine locker geplanten Einheiten in die Grauzone gerutscht.
        </div>
      )}
    </div>
  );
}
