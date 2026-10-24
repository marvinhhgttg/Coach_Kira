import { useEffect, useState } from 'react';
import { ladeP3, rampe, formZone, type P3Tag } from '../lib/paket3';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
function fmtDateShort(iso: string): string {
  const p = iso.slice(0, 10).split('-');
  if (p.length < 3) return iso;
  return `${parseInt(p[2])}.${p[1].padStart(2, '0')}`;
}

function weekdayShort(iso: string): string {
  const d = new Date(iso + 'T12:00:00Z');
  const tag = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  return tag[d.getUTCDay()];
}

function formPctClass(formPct: number): string {
  if (formPct > 25) return 'text-sky-300';
  if (formPct > 5) return 'text-green-300';
  if (formPct > -10) return 'text-ink-muted';
  if (formPct > -30) return 'text-green-400';
  return 'text-orange-300';
}

function formBgClass(formPct: number): string {
  if (formPct > 25) return 'bg-sky-500/20';
  if (formPct > 5) return 'bg-green-500/15';
  if (formPct > -10) return 'bg-ink-dim/10';
  if (formPct > -30) return 'bg-green-500/30';
  return 'bg-orange-500/25';
}

function formZonenname(formPct: number): string {
  if (formPct > 25) return 'Transition';
  if (formPct > 5) return 'Frisch';
  if (formPct > -10) return 'Grau';
  if (formPct > -30) return 'Optimal';
  return 'Hohes Risiko';
}

function rampeUrteilClass(urteil: string | null): string {
  switch (urteil) {
    case 'abbauend': return 'text-ink-muted';
    case 'aufbauend': return 'text-green-400';
    case 'ambitioniert': return 'text-yellow-400';
    case 'hoch': return 'text-orange-300';
    default: return 'text-ink-muted';
  }
}

// ── Komponente ───────────────────────────────────────────────────────────
export function FormRampeTile({ plan }: { plan?: Array<{ date: string; atl: number; ctl: number }> }) {
  const [tage, setTage] = useState<P3Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

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

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Form und Rampe lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  // Form berechnen: letzter Tag mit atl und ctl
  let formPct: number | null = null;
  let zone: string = '';
  for (let i = tage.length - 1; i >= 0; i--) {
    const t = tage[i];
    if (t.ctl !== null && t.atl !== null) {
      const f = formZone(t.ctl, t.atl);
      formPct = f.formPct;
      zone = f.zone;
      break;
    }
  }

  // Rampe berechnen: r = rampe(tage)
  const rampData = rampe(tage);

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Form und Rampe</div>
        <p className="mt-1 text-2xs text-ink-muted">
          Form = (CTL − ATL) / CTL. Ergänzt den ACWR: ein Signal unter mehreren, keine harte Grenze.
        </p>
      </div>

      {/* Form heute */}
      {formPct !== null && (
        <div className="mt-3">
          <p className={`text-sm ${formPctClass(formPct)}`}>
            Form heute: {formPct >= 0 ? '+' : ''}{Math.round(formPct)} % · {zone}
          </p>
        </div>
      )}

      {/* Legende */}
      <p className="mt-2 text-2xs text-ink-muted">
        &gt; +25 % Transition (zu frisch) · +5…+25 Frisch · −10…+5 Grau · −30…−10 Optimal (Aufbau) · &lt; −30 Hohes Risiko
      </p>

      {/* Rampe */}
      {rampData.proWoche !== null && rampData.pct !== null && (
        <div className="mt-3">
          <p className={`text-xs ${rampeUrteilClass(rampData.urteil)}`}>
            Rampenrate: {rampData.proWoche >= 0 ? '+' : ''}{Math.round(rampData.proWoche)} CTL/Woche ({rampData.pct >= 0 ? '+' : ''}{Math.round(rampData.pct)} %) – {rampData.urteil}
          </p>
          <p className="mt-1 text-2xs text-ink-muted">
            Über 8 % pro Woche steigt das Risiko, 5–8 % ist ambitioniert.
          </p>
        </div>
      )}

      {/* Plan-Stripe */}
      {plan && plan.length > 0 && (
        <div className="mt-4">
          <p className="text-xs text-ink-muted mb-1">Dein Plan:</p>
          <div className="flex gap-1 flex-wrap overflow-x-auto">
            {plan.slice(0, 14).map((p, idx) => {
              const f = formZone(p.ctl, p.atl);
              return (
                <div
                  key={idx}
                  className={`rounded px-1.5 py-0.5 text-center min-w-[2.2rem] ${formBgClass(f.formPct)}`}
                  title={formZonenname(f.formPct)}
                >
                  <span className="text-2xs text-ink-muted">{weekdayShort(p.date)} {fmtDateShort(p.date)}</span>
                  <br />
                  <span className="text-2xs font-mono">{f.formPct >= 0 ? '+' : ''}{Math.round(f.formPct)}%</span>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
