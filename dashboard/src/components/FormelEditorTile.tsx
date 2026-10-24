import { useEffect, useState } from 'react';
import { fetchChartData } from '../lib/api';
import { parseFormel, rechne, VARIABLEN, type Knoten } from '../lib/formel';

// ── Typen und Hilfsfunktionen ───────────────────────────────────────
interface FormelEintrag {
  name: string;
  formel: string;
}

const LS_KEY = 'ck_formeln';
const STANDARD_FORMELN: FormelEintrag[] = [
  { name: 'Ruhepuls gegen 28-Tage-Schnitt', formel: 'rhr - avg(rhr, 28)' },
  { name: 'Rampenrate CTL/Woche', formel: 'ctl - lag(ctl, 7)' },
  { name: 'Form %', formel: '(ctl - atl) / ctl * 100' },
  { name: 'Schlafschuld 7 Tage', formel: 'sum(7.5 - sleep, 7)' },
];

function fmtDateShort(iso: string): string {
  const p = iso.split('-');
  if (p.length < 3) return iso;
  const tag = parseInt(p[2]);
  const monat = parseInt(p[1]);
  return `${tag}.${String(monat).padStart(2, '0')}.${p[0]}`;
}

function fmtNum(n: number): string {
  return n.toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtNumInt(n: number): string {
  return n.toLocaleString('de-DE');
}

// Mapping API-Spalte → VARIABLEN-Name
const METRIC_MAP: Record<string, string> = {
  load_fb_day: 'load',
  fbATL_obs: 'atl',
  fbCTL_obs: 'ctl',
  fbACWR_obs: 'acwr',
  rhr_bpm: 'rhr',
  hrv_status: 'hrv',
  sleep_hours: 'sleep',
  sleep_score_0_100: 'sleepScore',
  Garmin_Training_Readiness: 'readiness',
  garminEnduranceScore: 'es',
  Monotony7: 'mono',
};

const VARIABLEN_ORDER = Object.keys(VARIABLEN);

// ── Komponente ─────────────────────────────────────────────────────
export function FormelEditorTile() {
  const [rawData, setRawData] = useState<Record<string, any>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Formeln
  const [formeln, setFormeln] = useState<FormelEintrag[]>([]);
  const [selectedIdx, setSelectedIdx] = useState<number | null>(0);

  // Neue Formel
  const [newName, setNewName] = useState('');
  const [newFormel, setNewFormel] = useState('');

  // Zeitraum
  const [tage, setTage] = useState(30);

  // Berechnetes Ergebnis
  const [resultData, setResultData] = useState<{ label: string; xVal: string; yVal: number | null }[]>([]);
  const [chartError, setChartError] = useState<string | null>(null);
  const [ast, setAst] = useState<Knoten | null>(null);

  // Statistiken
  const [statsLast, setStatsLast] = useState<number | null>(null);
  const [statsMin, setStatsMin] = useState<number | null>(null);
  const [statsMax, setStatsMax] = useState<number | null>(null);

  useEffect(() => {
    // Daten laden
    const loadChart = async () => {
      try {
        const metrics = [
          'load_fb_day', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs',
          'rhr_bpm', 'hrv_status', 'sleep_hours',
          'sleep_score_0_100', 'Garmin_Training_Readiness',
          'garminEnduranceScore', 'Monotony7',
        ];
        const res = await fetchChartData('9999d', metrics);
        if (!res || !res.ok) throw new Error('Chart-Daten nicht lesbar');
        // Nur Daten bis heute filtern
        const he = new Date();
        const datumStr = `${he.getFullYear()}-${String(he.getMonth() + 1).padStart(2, '0')}-${String(he.getDate()).padStart(2, '0')}`;
        const filtered = (res.data || []).filter((r: Record<string, any>) => r.date <= datumStr);
        // Nach Datum sortieren
        filtered.sort((a: Record<string, any>, b: Record<string, any>) => a.date.localeCompare(b.date));
        setRawData(filtered);
      } catch (e) {
        setError(e instanceof Error ? e.message : 'Unbekannter Fehler');
      } finally {
        setLoading(false);
      }
    };

    loadChart();
  }, []);

  // Formeln aus localStorage laden
  useEffect(() => {
    try {
      const saved = localStorage.getItem(LS_KEY);
      if (saved) {
        setFormeln(JSON.parse(saved));
      } else {
        const init = STANDARD_FORMELN;
        setFormeln(init);
        localStorage.setItem(LS_KEY, JSON.stringify(init));
      }
    } catch {
      setFormeln(STANDARD_FORMELN);
    }
  }, []);

  // Bei Formel-Änderung oder Zeitraum-Wechsel neu berechnen
  useEffect(() => {
    if (selectedIdx == null || selectedIdx < 0 || selectedIdx >= formeln.length) {
      setResultData([]);
      setChartError(null);
      setAst(null);
      setStatsLast(null);
      setStatsMin(null);
      setStatsMax(null);
      return;
    }

    const entry = formeln[selectedIdx];
    try {
      // Parser-AST erstellen und zwischenspeichern
      const parsed = parseFormel(entry.formel);
      setAst(parsed);

      // Reihen aufbauen: VARIABLEN-Namen → Array von Werten
      const reihen: Record<string, Array<number | null>> = {};
      for (const varName of VARIABLEN_ORDER) {
        const metric = METRIC_MAP[varName];
        if (!metric || !rawData.length) {
          reihen[varName] = new Array(rawData.length).fill(null);
          continue;
        }
        reihen[varName] = rawData.map((r: Record<string, any>) => {
          const v = r[metric];
          return v != null ? Number(v) : null;
        });
      }

      // Nur die letzten N Tage verwenden (abzüglich der Datenlänge)
      const dataLen = rawData.length;
      const startIdx = Math.max(0, dataLen - tage);
      const sliceReihen: Record<string, Array<number | null>> = {};
      for (const k of Object.keys(reihen)) {
        sliceReihen[k] = reihen[k].slice(startIdx);
      }

      // Berechne Ergebnis-Array elementweise
      const erg = rechne(parsed, sliceReihen);

      // Mit Datum verbinden (nur die Datenpunkte ab startIdx)
      const points: { label: string; xVal: string; yVal: number | null }[] = [];
      const validValues: number[] = [];

      for (let i = 0; i < erg.length; i++) {
        const datumPkt = rawData[startIdx + i];
        if (!datumPkt) continue;
        const d = new Date(datumPkt.date);
        points.push({
          label: fmtDateShort(datumPkt.date),
          xVal: datumPkt.date,
          yVal: erg[i] as number | null,
        });
        const v = erg[i];
        if (v != null && Number.isFinite(v)) {
          validValues.push(v);
        }
      }

      setResultData(points);
      setChartError(null);

      // Statistiken berechnen
      if (validValues.length > 0) {
        setStatsLast(erg[erg.length - 1] as number);
        setStatsMin(Math.min(...validValues));
        setStatsMax(Math.max(...validValues));
      } else {
        setStatsLast(null);
        setStatsMin(null);
        setStatsMax(null);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Unbekannter Fehler';
      setChartError(msg);
      setResultData([]);
    }
  }, [selectedIdx, formeln, rawData, tage]);

  // Formel speichern
  function speichereFormel() {
    if (!newName.trim() || !newFormel.trim()) return;
    const neu: FormelEintrag = { name: newName.trim(), formel: newFormel.trim() };
    const updated = [...formeln, neu];
    setFormeln(updated);
    localStorage.setItem(LS_KEY, JSON.stringify(updated));
    setSelectedIdx(updated.length - 1);
    setNewName('');
    setNewFormel('');
  }

  // Formel löschen
  function loescheFormel(idx: number) {
    const updated = formeln.filter((_, i) => i !== idx);
    setFormeln(updated);
    localStorage.setItem(LS_KEY, JSON.stringify(updated));
    if (selectedIdx === idx) {
      setSelectedIdx(updated.length > 0 ? 0 : null);
    } else if (selectedIdx != null && selectedIdx > idx) {
      setSelectedIdx(selectedIdx - 1);
    }
  }

  if (loading) {
    return <div className="panel-raised p-4"><p className="text-sm text-ink-dim">Eigene Kennzahlen lädt…</p></div>;
  }

  if (error) {
    return <div className="panel-raised p-4"><p className="text-sm text-orange-300">{error}</p></div>;
  }

  return (
    <div className="panel-raised p-4">
      {/* Kopf */}
      <div>
        <div className="label">Coach Kira · Eigene Kennzahlen</div>
        <p className="mt-1 text-2xs text-ink-dim">
          Formel eingeben, z. B. rhr - avg(rhr, 28) oder ctl - lag(ctl, 7). Rechnet lokal im Browser.
        </p>
      </div>

      {/* Formel-Eingabe */}
      <div className="mt-3 flex gap-2 items-end">
        <div>
          <label className="text-2xs text-ink-dim block mb-0.5">Name</label>
          <input
            type="text"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            placeholder="Formel-Name"
            className="border border-border rounded px-2 py-1 text-sm bg-transparent w-36"
          />
        </div>
        <div className="flex-1">
          <label className="text-2xs text-ink-dim block mb-0.5">Formel</label>
          <input
            type="text"
            value={newFormel}
            onChange={(e) => setNewFormel(e.target.value)}
            placeholder="z. B. ctl - lag(ctl, 7)"
            className="border border-border rounded px-2 py-1 text-sm bg-transparent w-full"
          />
        </div>
        <button onClick={speichereFormel} className="px-3 py-1 border border-border rounded text-sm hover:bg-gray-50 shrink-0">
          Speichern
        </button>
      </div>

      {/* Formel-Liste */}
      <div className="mt-2 flex flex-wrap gap-1">
        {formeln.map((f, i) => (
          <div key={i} className="flex items-center">
            <button
              onClick={() => setSelectedIdx(i === selectedIdx ? null : i)}
              className={`px-2 py-1 border rounded text-xs ${i === selectedIdx ? 'bg-accent/10 border-accent' : 'border-border hover:bg-gray-50'}`}
            >
              {f.name}
            </button>
            <button onClick={() => loescheFormel(i)} className="ml-0.5 text-xs text-ink-dim hover:text-orange-400 px-1" title="Löschen">×</button>
          </div>
        ))}
      </div>

      {/* Zeitraum-Umschalter */}
      <div className="mt-3 flex gap-1">
        {[30, 90, 180, 365].map(d => (
          <button
            key={d}
            onClick={() => setTage(d)}
            className={`px-2 py-1 border rounded text-xs ${tage === d ? 'bg-accent/10 border-accent' : 'border-border hover:bg-gray-50'}`}
          >
            {d} Tage
          </button>
        ))}
      </div>

      {/* Diagramm */}
      <div className="mt-3">
        {selectedIdx != null && selectedIdx >= 0 && selectedIdx < formeln.length ? (
          <>
            {/* Parser-Fehler */}
            {chartError && (
              <p className="text-orange-300 text-2xs mb-1">{chartError}</p>
            )}

            {/* recharts-Diagramm: ResponsiveContainer, LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid */}
            {/* Da recharts importiert werden muss: direkte Implementierung mit den Named-Imports */}
            <div style={{ height: 220 }}>
              {resultData.length > 0 ? (
                <div className="recharts-wrapper" style={{ height: 220 }}>
                  {/* Wir bauen das Diagramm manuell mit SVG für maximale Kompatibilität */}
                  <DiagramChartWerkzeuge daten={resultData} error={chartError} />
                </div>
              ) : (
                <p className="text-2xs text-ink-dim flex items-center justify-center" style={{ height: 220 }}>
                  Keine Daten für diese Formel verfügbar.
                </p>
              )}
            </div>

            {/* Statistiken unter dem Diagramm */}
            {statsLast != null && statsMin != null && statsMax != null && (
              <div className="mt-2 text-2xs text-ink-muted grid grid-cols-3 gap-1">
                <div>Letzter: {fmtNum(statsLast)}</div>
                <div>Minimum: {fmtNum(statsMin)}</div>
                <div>Maximum: {fmtNum(statsMax)}</div>
              </div>
            )}
          </>
        ) : (
          <p className="mt-3 text-2xs text-ink-dim">Wähle eine Formel aus, um das Diagramm zu sehen.</p>
        )}

        {/* Hilfe */}
        <div className="mt-3 text-2xs text-ink-dim">
          <details>
            <summary className="cursor-pointer hover:text-ink">Hilfe: VARIABLEN und Funktionen</summary>
            <div className="mt-2 space-y-1">
              {/* VARIABLEN */}
              <div>
                <span className="font-semibold">VARIABLEN:</span>
                {VARIABLEN_ORDER.map(v => (
                  <div key={v}><code className="text-accent">{v}</code> — {VARIABLEN[v]}</div>
                ))}
              </div>
              {/* Funktionen */}
              <div className="mt-1">
                <span className="font-semibold">FUNKTIONEN:</span>
                <ul className="ml-4 mt-0.5 space-y-0.5">
                  <li><code>avg(x, n)</code> — Durchschnitt der letzten n Werte von Reihe x</li>
                  <li><code>sum(x, n)</code> — Summe der letzten n Werte von Reihe x</li>
                  <li><code>min(x, n)</code> — Minimum der letzten n Werte (null ignoriert)</li>
                  <li><code>max(x, n)</code> — Maximum der letzten n Werte (null ignoriert)</li>
                  <li><code>lag(x, n)</code> — Wert vor n Tagen</li>
                  <li><code>abs(x)</code> — Absolutwert</li>
                  <li><code>sqrt(x)</code> — Quadratwurzel</li>
                  <li><code>log(x)</code> — Natürlicher Logarithmus</li>
                </ul>
              </div>
            </div>
          </details>
        </div>
      </div>
    </div>
  );
}

// ── Diagramm-Komponente (SVG-basiert, mit recharts-Komponenten-Props) ─
interface DiagramChartWerkzeugeProps {
  daten: { label: string; xVal: string; yVal: number | null }[];
  error?: string | null;
}

function DiagramChartWerkzeuge({ daten, error }: DiagramChartWerkzeugeProps) {
  // Filtere gültige Werte für Y-Achsen-Bereich
  const validVals = daten.map(d => d.yVal).filter(v => v != null && Number.isFinite(v as number)) as number[];
  if (validVals.length === 0 && !error) return null;

  const width = daten.length > 0 ? Math.max(daten.length * 3, 200) : 200;
  const height = 196; // 220 - padding
  const padL = 50, padR = 10, padT = 10, padB = 30;
  const chartW = width - padL - padR;
  const chartH = height - padT - padB;

  let yMin: number, yMax: number;
  if (validVals.length > 0) {
    const margin = (Math.max(...validVals) - Math.min(...validVals)) * 0.1;
    yMin = Math.min(...validVals) - margin;
    yMax = Math.max(...validVals) + margin;
  } else {
    yMin = -10; yMax = 10;
  }

  function xPos(i: number): number { return padL + (i / Math.max(daten.length - 1, 1)) * chartW; }
  function yPos(v: number): number { return padT + (1 - (v - yMin) / (yMax - yMin)) * chartH; }

  // Linien-Pfad bauen
  const pts = daten.map((d, i) => ({ x: xPos(i), y: yPos(d.yVal ?? 0) }));
  const pathParts: string[] = [];
  let started = false;
  for (let i = 0; i < pts.length; i++) {
    if (daten[i].yVal == null) { started = false; continue; }
    if (!started) { pathParts.push(`M${pts[i].x},${pts[i].y}`); started = true; }
    else { pathParts.push(`L${pts[i].x},${pts[i].y}`); }
  }
  const pathD = pathParts.join(' ');

  // Gitter-Linien (5 horizontal)
  const gridLines: string[] = [];
  for (let j = 0; j <= 4; j++) {
    const gv = yMin + (yMax - yMin) * (1 - j / 4);
    const gy = yPos(gv);
    gridLines.push(`M${padL},${gy} L${width - padR},${gy}`);
  }

  // Y-Achsen-Beschriftungen
  const yLabels: { label: string; yPos: number }[] = [];
  for (let j = 0; j <= 4; j++) {
    const gv = yMin + (yMax - yMin) * (1 - j / 4);
    const gy = yPos(gv);
    yLabels.push({ label: Math.round(gv).toString(), yPos: gy });
  }

  // X-Achsen-Beschriftungen (alle N Tage)
  const xLabels: { label: string; xPos: number }[] = [];
  const stepX = Math.max(1, Math.floor(daten.length / 8));
  for (let i = 0; i < daten.length; i += stepX) {
    xLabels.push({ label: daten[i].label, xPos: xPos(i) });
  }

  return (
    <svg width={width} height={height} className="text-accent">
      {/* Gitter */}
      <g stroke="#ddd" strokeWidth={0.5}>
        {gridLines.map((d, i) => <path key={i} d={d} />)}
      </g>

      {/* Linie */}
      {pathD && (
        <path d={pathD} fill="none" stroke="currentColor" strokeWidth={2} className="text-accent" />
      )}

      {/* Y-Achsen-Beschriftungen */}
      {yLabels.map((l, i) => (
        <text key={i} x={padL - 4} y={l.yPos + 3} textAnchor="end" fontSize={10} fill="#666">
          {l.label}
        </text>
      ))}

      {/* X-Achsen-Beschriftungen */}
      {xLabels.map((l, i) => (
        <text key={i} x={l.xPos} y={height - 2} textAnchor="middle" fontSize={9} fill="#666">
          {l.label}
        </text>
      ))}

      {/* Achsen */}
      <line x1={padL} y1={padT} x2={padL} y2={height - padB} stroke="#ccc" strokeWidth={1} />
      <line x1={padL} y1={height - padB} x2={width - padR} y2={height - padB} stroke="#ccc" strokeWidth={1} />
    </svg>
  );
}
