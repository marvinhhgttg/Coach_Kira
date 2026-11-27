import { useEffect, useState, useMemo } from 'react';
import { ladeP4, jahresRueckblick, type P4Tag } from '../lib/paket4';

// ── Hilfsfunktionen ────────────────────────────────────────────────
function fmtDateShort(iso: string): string {
  const p = iso.split('-');
  if (p.length < 3) return iso;
  const tag = parseInt(p[2]);
  const monat = parseInt(p[1]);
  return `${tag}.${String(monat).padStart(2, '0')}.${p[0]}`;
}

function fmtOne(n: number): string {
  return n.toLocaleString('de-DE', { maximumFractionDigits: 1 });
}

function fmtInt(n: number): string {
  return n.toLocaleString('de-DE');
}

// Monatsname auf Deutsch
function monatName(monatStr: string): string {
  const monate = [
    'Januar', 'Februar', 'März', 'April', 'Mai', 'Juni',
    'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember',
  ];
  const p = monatStr.split('-');
  if (p.length < 2) return monatStr;
  const idx = parseInt(p[1]) - 1;
  if (idx < 0 || idx > 11) return monatStr;
  const jahr = p[0];
  return `${monate[idx]} ${jahr}`;
}

// Kachel-Darstellung
interface TileKachelProps {
  label: string;
  wert: string;
}

function TileKachel({ label, wert }: TileKachelProps) {
  return (
    <div className="rounded-lg border border-border p-3">
      <div className="text-2xs text-ink-muted">{label}</div>
      <div className="mt-1 text-sm font-medium">{wert}</div>
    </div>
  );
}

// ── Komponente ─────────────────────────────────────────────────────
export function JahresrueckblickTile() {
  const [tage, setTage] = useState<P4Tag[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Bestimme verfügbare Jahre aus den Daten
  const verfuegbareJahre = useMemo(() => {
    if (tage.length === 0) return [] as number[];
    const jahre = new Set<number>();
    for (const tag of tage) {
      const p = tag.date.split('-');
      if (p.length >= 1) {
        const jahr = parseInt(p[0]);
        if (jahr > 0) jahre.add(jahr);
      }
    }
    return Array.from(jahre).sort((a, b) => a - b);
  }, [tage]);

  // Aktive Auswahl: startet mit 'letzte365'
  const [auswahl, setAuswahl] = useState<number | 'letzte365'>(verfuegbareJahre.length > 0 ? verfuegbareJahre[verfuegbareJahre.length - 1] : 'letzte365');

  // Warte bis tage geladen sind, dann wähle das neueste Jahr
  useEffect(() => {
    if (tage.length > 0 && verfuegbareJahre.length > 0) {
      setAuswahl(verfuegbareJahre[verfuegbareJahre.length - 1]);
    }
  }, [tage, verfuegbareJahre]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const t = await ladeP4();
        if (!cancelled) setTage(t);
        setLoading(false);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // Rückblick berechnen
  const r = useMemo(() => {
    return jahresRueckblick(tage, auswahl);
  }, [tage, auswahl]);

  const jahreKnöpfe: (number | 'letzte365')[] = ['letzte365', ...verfuegbareJahre];

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Jahresrückblick lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  // ── Kachel-Werte berechnen ───────────────────────────────────────

  const gesamtLast = fmtInt(Math.round(r.summeLast));
  const trainingTage = fmtInt(r.einheiten);
  const kilometer = r.km > 0 ? fmtInt(Math.round(r.km)) : 'nicht erfasst';
  const hoehenmeter = r.hm > 0 ? fmtInt(Math.round(r.hm)) : 'nicht erfasst';

  const liebSport = r.topSport.name
    ? `${r.topSport.name} (${Math.round(r.topSport.anteil * 100)} %)`
    : '–';

  const besterTagStr = r.besterTag.date && r.besterTag.load
    ? `${fmtDateShort(r.besterTag.date)} mit ${fmtInt(Math.round(r.besterTag.load))} ESS`
    : '–';

  const langsteSerieStr = r.laengsteSerie.tage > 0
    ? `${r.laengsteSerie.tage} Tage bis ${fmtDateShort(r.laengsteSerie.bis)}`
    : '–';

  const esStr = (r.esStart != null && r.esEnde != null)
    ? `${fmtInt(Math.round(r.esStart))} → ${fmtInt(Math.round(r.esEnde))}`
    : '–';

  const esMaxStr = r.esMax != null && r.esMax.date
    ? `Höchstwert ${fmtInt(Math.round(r.esMax.es))} am ${fmtDateShort(r.esMax.date)}`
    : '–';


  const ctlStr = (r.ctlStart != null && r.ctlEnde != null)
    ? `${fmtOne(Math.round(r.ctlStart * 10) / 10)} → ${fmtOne(Math.round(r.ctlEnde * 10) / 10)}`
    : '–';

  const schlafKönigStr = r.besterSchlafMonat != null
    ? `${monatName(r.besterSchlafMonat.monat)} mit ${fmtOne(Math.round(r.besterSchlafMonat.schlaf * 10) / 10)} h`
    : '–';

  const readinessStr = r.besterReadinessMonat != null
    ? `${monatName(r.besterReadinessMonat.monat)} mit Ø ${fmtOne(Math.round(r.besterReadinessMonat.readiness * 10) / 10)}`
    : '–';

  const ruhigStr = r.ruhigsterMonat != null
    ? `${monatName(r.ruhigsterMonat.monat)} mit ${fmtInt(Math.round(r.ruhigsterMonat.last))} ESS`
    : '–';

  const bestBlockStr = r.besterBlock != null && r.besterBlock.esDelta != null
    ? `${fmtDateShort(r.besterBlock.start)}–${fmtDateShort(r.besterBlock.ende)} mit +${r.besterBlock.esDelta.toLocaleString('de-DE')} ES`
    : '–';

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Jahresrückblick</div>
      </div>

      {/* Auswahl-Knöpfe */}
      <div className="mt-3 flex flex-wrap gap-2">
        {jahreKnöpfe.map(k => (
          <button
            key={k === 'letzte365' ? 'letzte365' : String(k)}
            className={`chip rounded-lg border border-border p-3 cursor-pointer text-xs ${auswahl === k ? 'bg-blue-500/20 border-blue-400' : ''}`}
            onClick={() => setAuswahl(k)}
          >
            {k === 'letzte365' ? 'Letzte 365 Tage' : String(k)}
          </button>
        ))}
      </div>

      {/* Titel */}
      <div className="mt-3 text-lg font-bold">{r.titel}</div>

      {/* Kacheln-Raster */}
      <div className="mt-3 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
        <TileKachel label="Gesamtlast" wert={gesamtLast} />
        <TileKachel label="Trainingstage" wert={trainingTage} />
        <TileKachel label="Kilometer" wert={kilometer} />
        <TileKachel label="Höhenmeter" wert={hoehenmeter} />
        <TileKachel label="Lieblingssport" wert={liebSport} />
        <TileKachel label="Bester Tag" wert={besterTagStr} />
        <TileKachel label="Längste Serie" wert={langsteSerieStr} />
        <TileKachel label="Endurance Score" wert={esMaxStr !== '–' ? `${esStr} · ${esMaxStr}` : esStr} />
        <TileKachel label="Fitness (CTL)" wert={ctlStr} />
        <TileKachel label="Schlafkönig-Monat" wert={schlafKönigStr} />
        <TileKachel label="Bester Readiness-Monat" wert={readinessStr} />
        <TileKachel label="Ruhigster Monat" wert={ruhigStr} />
        <TileKachel label="Bester Block" wert={bestBlockStr} />
      </div>
    </div>
  );
}
