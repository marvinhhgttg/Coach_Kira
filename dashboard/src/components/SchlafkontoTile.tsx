import { useEffect, useState } from 'react';
import { ladeP3, schlafKonto, type P3Tag } from '../lib/paket3';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────
function fmtDe(v: number): string {
  return v.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function minutenZuHHMM(min: number): string {
  const h = Math.floor(Math.abs(min) / 60);
  const m = Math.round(Math.abs(min) % 60);
  return `${h} h ${m.toString().padStart(2, '0')} min`;
}

function formatTime(value: string): string {
  if (!value) return '';
  const parts = value.split(':');
  if (parts.length === 2) return `${parts[0]}:${parts[1]}`;
  return value;
}

// ── Komponente ───────────────────────────────────────────────────────────
export function SchlafkontoTile() {
  const [tage, setTage] = useState<P3Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aufstehen, setAufstehen] = useState('');

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeP3();
        if (!cancelled) setTage(t);

        // Aufsteh-Zeit aus localStorage laden
        const saved = localStorage.getItem('ck_aufstehen');
        let defaultTime = '06:15';
        if (saved) {
          if (!cancelled) setAufstehen(saved);
        } else {
          const morgen = new Date();
          morgen.setDate(morgen.getDate() + 1);
          const mt = morgen.getDay(); // 0=So, 6=Sa
          defaultTime = (mt === 0 || mt === 6) ? '07:30' : '06:15';
          if (!cancelled) setAufstehen(defaultTime);
        }

        setLoading(false);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Schlafkonto lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  const lastLoad = tage.length > 0 ? tage[tage.length - 1].load ?? null : null;
  const k = schlafKonto(tage, lastLoad);

  // Quartil-Labels
  const bereichLabel = ['bei wenig Schuld', 'niedrige Schuld', 'mittlere Schuld', 'hohe Schuld'];

  // Readiness-Werte pro Quartil (gerundet)
  const readinessText = [0, 1, 2, 3].map(i =>
    k.quartile[i] != null && Math.round(k.quartile[i].readiness) > 0
      ? String(Math.round(k.quartile[i].readiness))
      : 'n/a'
  );

  // Farbe nach jetztQuartil: 0=grün, 1=grün, 2=gelb, 3=orange
  const farbeMap = ['text-green-500', 'text-green-400', 'text-yellow-400', 'text-orange-500'];
  const farbe = k.jetztQuartil !== null ? farbeMap[k.jetztQuartil] : 'text-ink-dim';

  // Schuld oder Guthaben
  let schuldText: string;
  if (k.schuld == null) {
    schuldText = 'keine Schuld';
  } else if (k.schuld < 0) {
    schuldText = `Guthaben ${fmtDe(Math.abs(k.schuld))} h`;
  } else {
    schuldText = `Schlafschuld ${fmtDe(k.schuld)} h`;
  }

  // Zuschläge in Minuten (dezimal)
  const zuschlagLastMin = k.zuschlagLast * 60;
  const zuschlagSchuldMin = k.zuschlagSchuld * 60;

  // Aufstehzeit-Default berechnen
  const morgen = new Date();
  morgen.setDate(morgen.getDate() + 1);
  const mtWochentag = morgen.getDay();

  // Licht aus berechnen: Aufstehzeit − Bedarf − 15 min Einschlafen
  const aufstMin = (aufstehen || '').split(':');
  const asH = parseInt(aufstMin[0] ?? '6') || 0;
  const asM = parseInt(aufstMin[1] ?? '0') || 0;
  const totalAs = asH * 60 + asM;
  const bedarfMin = k.bedarfH * 60;
  // über Mitternacht zurückrechnen (z. B. 06:15 − 7 h 50 − 15 min = 22:10)
  const lichtAusTotal = ((Math.round(totalAs - bedarfMin - 15) % 1440) + 1440) % 1440;
  const laH = Math.floor(lichtAusTotal / 60);
  const laM = lichtAusTotal % 60;

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Schlafkonto</div>
      </div>

      {/* Zeile 1: Schlafschuld / Guthaben */}
      <div className={`mt-2 text-lg ${farbe}`}>
        {schuldText}
      </div>

      {/* Zeile 2: Readiness-Bereich */}
      <p className="mt-1 text-xs text-ink-muted">
        In diesem Bereich lag deine Readiness im Schnitt bei{' '}
        {k.jetztQuartil !== null ? readinessText[k.jetztQuartil] : 'n/a'}{' '}
        {k.jetztQuartil !== null ? '(' + bereichLabel[k.jetztQuartil] + ')' : ''}
        {' '} (bei wenig Schuld: {readinessText[0]}).
      </p>

      {/* Block: Heute Nacht */}
      <div className="mt-3 p-3 bg-ink-dim/10 rounded-lg">
        <div className="font-semibold text-sm">Heute Nacht</div>
        <div className="mt-1 text-lg font-medium">{minutenZuHHMM(bedarfMin)}</div>
        <p className="mt-1 text-2xs text-ink-muted">
          7 h 30 Basis + {Math.round(zuschlagLastMin)} min für die Last heute + {Math.round(zuschlagSchuldMin)} min Schuldabbau
        </p>
      </div>

      {/* Aufstehzeit-Eingabe */}
      <div className="mt-3">
        <label className="block text-xs text-ink-muted mb-1">Aufstehzeit</label>
        <input
          type="time"
          value={aufstehen || (mtWochentag === 0 || mtWochentag === 6 ? '07:30' : '06:15')}
          onChange={(e) => {
            const v = e.target.value;
            setAufstehen(v);
            localStorage.setItem('ck_aufstehen', v);
          }}
          className="w-full px-3 py-2 rounded-md border bg-transparent text-sm"
        />
      </div>

      {/* Licht aus gross */}
      <div className="mt-2 text-center">
        <p className="text-sm text-ink-muted">Licht aus</p>
        <p className="text-3xl font-medium">{laH}:{laM.toString().padStart(2, '0')}</p>
      </div>
    </div>
  );
}
