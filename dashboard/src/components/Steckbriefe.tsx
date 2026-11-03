import { useEffect, useMemo, useState } from 'react';
import { Panel } from './UI';
import { KiraFazit, fazitAusCache } from './KiraFazit';
import { Hinweis, Kz, Laedt } from './AnalysenTraining';
import { useGarminTage, type GarminAkt, type GarminTage } from '../lib/garmin';
import { fetchChartData } from '../lib/api';
import { rekorde, GRUPPEN } from '../lib/rekorde';
const GLAB: Record<string, string> = Object.fromEntries(GRUPPEN.map((g) => [g.id, g.label]));
import { effPunkte, nachGelaende, trend, GRENZEN_STANDARD } from '../lib/effizienz';
import { BERG, LAUF, RAD, RAD_ALLE, de, dauerTxt, eintraege, kurz, median, mittel, monatName, plusTage, vz, type E } from '../lib/stat';

// ─── gemeinsamer Druck ─────────────────────────────────────────────────────
const esc = (s: unknown) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));
const CSS = `
*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;color:#111;margin:0;padding:28px 34px;font-size:11.5px;line-height:1.45}
h1{font-size:22px;margin:0 0 2px}h2{font-size:13px;margin:18px 0 6px;border-bottom:1px solid #ddd;padding-bottom:3px;text-transform:uppercase;letter-spacing:.04em;color:#333}
.sub{color:#666;margin-bottom:10px}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}.k{border:1px solid #e2e2e2;border-radius:6px;padding:7px 9px}
.k b{display:block;font-size:16px;font-variant-numeric:tabular-nums}.k span{color:#666;font-size:10px}
table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}td,th{padding:3px 5px;border-bottom:1px solid #eee;text-align:left;vertical-align:top}th{color:#666;font-weight:600;font-size:10px}
.faz{background:#f6f7f9;border-radius:6px;padding:10px 12px;white-space:pre-wrap}.neu{color:#15803d;font-weight:600}.foot{margin-top:18px;color:#888;font-size:9.5px}
svg text{font-size:9px;fill:#555}@media print{body{padding:10mm 12mm}h2{break-after:avoid}.k,tr{break-inside:avoid}}@page{size:A4;margin:0}
`;
export function drucken(titel: string, html: string) {
  const w = window.open('', '_blank');
  if (!w) { alert('Bitte Pop-ups für coachkira.pplx.app erlauben.'); return; }
  w.document.write(`<!doctype html><html lang="de"><head><meta charset="utf-8"><title>${esc(titel)}</title><style>${CSS}</style></head><body>${html}<div class="foot">Coach Kira · erstellt ${new Date().toLocaleString('de-DE')} · Daten: Garmin Connect und Coach-Kira-Sheet</div><script>setTimeout(()=>print(),400)<\/script></body></html>`);
  w.document.close();
}
/** Einfache Balken als SVG (für den Druck). */
function svgBalken(xs: { lab: string; v: number }[], einheit: string, farbe = '#2563eb') {
  const W = 680, H = 120, max = Math.max(1, ...xs.map((x) => x.v)), bw = W / Math.max(1, xs.length);
  return `<svg width="100%" viewBox="0 0 ${W} ${H + 28}">${xs.map((x, i) => {
    const h = (x.v / max) * H;
    return `<rect x="${i * bw + 3}" y="${H - h + 10}" width="${bw - 6}" height="${h}" fill="${farbe}" rx="2"/><text x="${i * bw + bw / 2}" y="${H + 24}" text-anchor="middle">${esc(x.lab)}</text><text x="${i * bw + bw / 2}" y="${H - h + 7}" text-anchor="middle">${de(x.v)}</text>`;
  }).join('')}<text x="0" y="8">${esc(einheit)}</text></svg>`;
}
function svgLinie(xs: { lab: string; v: number | null }[], einheit: string, farbe = '#7c3aed') {
  const W = 680, H = 110, vs = xs.map((x) => x.v).filter((v): v is number => v != null);
  if (vs.length < 2) return '';
  const lo = Math.min(...vs), hi = Math.max(...vs), sx = W / Math.max(1, xs.length - 1);
  const y = (v: number) => 10 + H - ((v - lo) / Math.max(1, hi - lo)) * H;
  const pts = xs.map((x, i) => (x.v == null ? null : `${i * sx},${y(x.v)}`)).filter(Boolean).join(' ');
  return `<svg width="100%" viewBox="-30 0 ${W + 40} ${H + 30}"><polyline points="${pts}" fill="none" stroke="${farbe}" stroke-width="2"/>${xs.map((x, i) => (i % Math.ceil(xs.length / 12) === 0 ? `<text x="${i * sx}" y="${H + 26}" text-anchor="middle">${esc(x.lab)}</text>` : '')).join('')}<text x="-28" y="${y(hi)}">${de(hi)}</text><text x="-28" y="${y(lo)}">${de(lo)}</text><text x="${W - 60}" y="8">${esc(einheit)}</text></svg>`;
}

// ─── 13 Saison-Steckbrief ──────────────────────────────────────────────────
type Zeitraum = 'jahr' | '12m' | 'alles';
const GR = [
  { lab: 'Laufen', typen: LAUF }, { lab: 'Rad draußen', typen: RAD }, { lab: 'Rad indoor', typen: new Set(['indoor_cycling', 'virtual_ride']) },
  { lab: 'Wandern & Berg', typen: new Set([...BERG, 'walking']) }, { lab: 'Ski & Skike', typen: new Set(['skate_skiing_ws', 'backcountry_skiing', 'cross_country_skiing_ws']) },
  { lab: 'Rudern', typen: new Set(['indoor_rowing', 'rowing_v2']) },
];
function teilTage(tage: GarminTage, von: string, bis: string): GarminTage {
  const m: GarminTage = new Map();
  for (const [d, g] of tage) if (d >= von && d <= bis) m.set(d, g);
  return m;
}
export function SaisonSteckbriefTile() {
  const { tage } = useGarminTage();
  const [zr, setZr] = useState<Zeitraum>('12m');
  const [ctl, setCtl] = useState<{ d: string; v: number }[] | null>(null);
  useEffect(() => {
    fetchChartData('9999d', ['fbCTL_obs']).then((r: any) => setCtl(((r && r.data) || []).map((x: any) => ({ d: String(x.date).slice(0, 10), v: Number(x.fbCTL_obs) })).filter((x: any) => Number.isFinite(x.v) && x.v > 0))).catch(() => setCtl([]));
  }, []);
  const heute = new Date().toLocaleDateString('sv-SE');
  const von = zr === 'jahr' ? heute.slice(0, 4) + '-01-01' : zr === '12m' ? plusTage(heute, -365) : '2000-01-01';
  const s = useMemo(() => {
    if (!tage) return null;
    const t = teilTage(tage, von, heute);
    const es = eintraege(t);
    const ds = [...t.keys()].sort();
    const gr = GR.map((g) => {
      const xs = es.filter((e) => g.typen.has(e.a.typ || ''));
      return { lab: g.lab, n: xs.length, km: xs.reduce((a, e) => a + (e.a.km || 0), 0), hm: xs.reduce((a, e) => a + (e.a.hm || 0), 0), s: xs.reduce((a, e) => a + (e.a.dauer_s || 0), 0) };
    }).filter((g) => g.n);
    const ges = { n: es.length, km: es.reduce((a, e) => a + (e.a.km || 0), 0), hm: es.reduce((a, e) => a + (e.a.hm || 0), 0), s: es.reduce((a, e) => a + (e.a.dauer_s || 0), 0) };
    const mon = new Map<string, { km: number; hm: number; s: number }>();
    for (const e of es) { const k = e.datum.slice(0, 7); const m = mon.get(k) || { km: 0, hm: 0, s: 0 }; m.km += e.a.km || 0; m.hm += e.a.hm || 0; m.s += e.a.dauer_s || 0; mon.set(k, m); }
    const monate = [...mon.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([m, x]) => ({ monat: m, ...x }));
    const rek = rekorde(t, heute);
    const effL = trend(nachGelaende(effPunkte(t, 'lauf'), 'flach', GRENZEN_STANDARD.lauf), 'hf');
    const effR = trend(nachGelaende(effPunkte(t, 'rad'), 'flach', GRENZEN_STANDARD.rad), 'hf');
    const rhr = ds.map((d) => ({ d, v: t.get(d)?.tag?.rhr })).filter((x) => x.v);
    const hrv = ds.map((d) => ({ d, v: t.get(d)?.hrv?.nacht })).filter((x) => x.v);
    const ersterM = (xs: { d: string; v?: number }[]) => median(xs.slice(0, 30).map((x) => x.v!)), letzterM = (xs: { d: string; v?: number }[]) => median(xs.slice(-30).map((x) => x.v!));
    const groesste = es.filter((e) => (e.a.hm || 0) > 0).sort((a, b) => (b.a.hm || 0) - (a.a.hm || 0)).slice(0, 5);
    const ctlR = (ctl || []).filter((x) => x.d >= von);
    return {
      von: ds[0], bis: ds[ds.length - 1], gr, ges, monate, rek, effL, effR, groesste,
      rhr: { a: ersterM(rhr), b: letzterM(rhr) }, hrv: { a: ersterM(hrv), b: letzterM(hrv) },
      ctl: ctlR.length ? { a: ctlR[0].v, b: ctlR[ctlR.length - 1].v, max: Math.max(...ctlR.map((x) => x.v)), maxD: ctlR.reduce((p, x) => (x.v > p.v ? x : p)).d, reihe: ctlR } : null,
    };
  }, [tage, von, heute, ctl]);

  const facts = s ? [
    `Zeitraum ${kurz(s.von)}–${kurz(s.bis)}: ${s.ges.n} Einheiten, ${de(s.ges.km)} km, ${de(s.ges.hm)} hm, ${dauerTxt(s.ges.s)}.`,
    ...s.gr.map((g) => `${g.lab}: ${g.n}×, ${de(g.km)} km, ${de(g.hm)} hm, ${dauerTxt(g.s)}`),
    s.ctl ? `Fitness (CTL) ${de(s.ctl.a)} → ${de(s.ctl.b)}, Höchstwert ${de(s.ctl.max)} am ${kurz(s.ctl.maxD)}.` : '',
    `Effizienz flach (Herzschläge je km): Laufen ${vz(s.effL.prozent, 1)} %, Rad ${vz(s.effR.prozent, 1)} % (negativ = effizienter).`,
    `Ruhepuls erster vs. letzter Monat ${de(s.rhr.a)} → ${de(s.rhr.b)}, HRV ${de(s.hrv.a)} → ${de(s.hrv.b)} ms.`,
    `Rekorde: ${s.rek.filter((r) => ['lauf', 'rad', 'berg', 'gesamt'].includes(r.gruppe)).slice(0, 12).map((r) => `${r.titel} ${r.text}`).join('; ')}.`,
  ].filter(Boolean) : [];
  const titel = `Saison-Steckbrief ${zr === 'jahr' ? heute.slice(0, 4) : zr === '12m' ? 'letzte 12 Monate' : 'gesamt'}`;
  const [fazit, setFazit] = useState<string | null>(null);
  useEffect(() => { setFazit(fazitAusCache(titel, facts)); }, [titel, facts.join('|')]); // eslint-disable-line

  function pdf() {
    if (!s) return;
    const k = (lab: string, v: string, sub = '') => `<div class="k"><span>${esc(lab)}</span><b>${esc(v)}</b><span>${esc(sub)}</span></div>`;
    const html = `
<h1>${esc(titel)}</h1><div class="sub">Marc Reichpietsch · ${kurz(s.von)} bis ${kurz(s.bis)}</div>
<div class="grid">${k('Einheiten', de(s.ges.n))}${k('Distanz', de(s.ges.km) + ' km')}${k('Höhenmeter', de(s.ges.hm) + ' hm')}${k('Trainingszeit', dauerTxt(s.ges.s))}
${s.ctl ? k('Fitness (CTL)', `${de(s.ctl.a)} → ${de(s.ctl.b)}`, `Höchstwert ${de(s.ctl.max)} am ${kurz(s.ctl.maxD)}`) : ''}
${k('Effizienz Laufen flach', vz(s.effL.prozent, 1) + ' %', 'Herzschläge je km, − = besser')}${k('Effizienz Rad flach', vz(s.effR.prozent, 1) + ' %', 'Herzschläge je km, − = besser')}
${k('Ruhepuls / HRV', `${de(s.rhr.a)} → ${de(s.rhr.b)} · ${de(s.hrv.a)} → ${de(s.hrv.b)} ms`, 'erster vs. letzter Monat')}</div>
${fazit ? `<h2>Kiras Fazit</h2><div class="faz">${esc(fazit)}</div>` : ''}
<h2>Nach Sportart</h2><table><tr><th>Sport</th><th>Einheiten</th><th>km</th><th>hm</th><th>Zeit</th></tr>${s.gr.map((g) => `<tr><td>${esc(g.lab)}</td><td>${g.n}</td><td>${de(g.km)}</td><td>${de(g.hm)}</td><td>${dauerTxt(g.s)}</td></tr>`).join('')}</table>
<h2>Trainingszeit je Monat (Stunden)</h2>${svgBalken(s.monate.map((m) => ({ lab: monatName(m.monat), v: Math.round(m.s / 360) / 10 })), 'h')}
${s.ctl ? `<h2>Fitnessverlauf (CTL)</h2>${svgLinie(s.ctl.reihe.filter((_, i) => i % 3 === 0).map((x) => ({ lab: kurz(x.d).slice(3), v: x.v })), 'CTL')}` : ''}
<h2>Rekorde</h2><table><tr><th>Bereich</th><th>Rekord</th><th>Wert</th><th>Datum</th><th>Einheit</th></tr>${s.rek.filter((r) => r.gruppe !== 'koerper').map((r) => `<tr><td>${esc(GLAB[r.gruppe] || r.gruppe)}</td><td>${esc(r.titel)}</td><td>${esc(r.text)}${r.neu ? ' <span class="neu">neu</span>' : ''}</td><td>${r.bis ? kurz(r.datum) + '–' + kurz(r.bis) : kurz(r.datum)}</td><td>${esc(r.detail || '')}</td></tr>`).join('')}</table>
<h2>Größte Touren nach Höhenmetern</h2><table><tr><th>Datum</th><th>Tour</th><th>km</th><th>hm</th><th>Dauer</th></tr>${s.groesste.map((e) => `<tr><td>${kurz(e.datum)}</td><td>${esc(e.a.name || '')}</td><td>${de(e.a.km, 1)}</td><td>${de(e.a.hm)}</td><td>${dauerTxt(e.a.dauer_s)}</td></tr>`).join('')}</table>`;
    drucken(titel, html);
  }
  const btn = (a: boolean) => `px-2 py-1 text-xs rounded ${a ? 'bg-bg-subtle text-ink' : 'text-ink-muted hover:text-ink'}`;
  return (
    <Panel title="Saison-Steckbrief" right={
      <div className="inline-flex bg-bg rounded border border-border p-0.5">
        <button className={btn(zr === 'jahr')} onClick={() => setZr('jahr')}>{heute.slice(0, 4)}</button>
        <button className={btn(zr === '12m')} onClick={() => setZr('12m')}>12 Monate</button>
        <button className={btn(zr === 'alles')} onClick={() => setZr('alles')}>gesamt</button>
      </div>
    }>
      <Hinweis>Umfang, Fitness, Effizienz, Rekorde und größte Touren auf einen Blick – mit Kiras Fazit als PDF. Erst das Fazit erstellen (lokal), dann „Als PDF“: Im Druckdialog „Als PDF sichern“ wählen.</Hinweis>
      {!s ? <Laedt /> : (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2">
            <Kz lab="Einheiten" wert={de(s.ges.n)} sub={`${kurz(s.von)}–${kurz(s.bis)}`} />
            <Kz lab="Distanz · Höhe" wert={`${de(s.ges.km)} km`} sub={`${de(s.ges.hm)} hm`} />
            <Kz lab="Trainingszeit" wert={dauerTxt(s.ges.s)} />
            <Kz lab="Fitness (CTL)" wert={s.ctl ? `${de(s.ctl.a)} → ${de(s.ctl.b)}` : '–'} sub={s.ctl ? `max ${de(s.ctl.max)}` : ''} />
          </div>
          <div className="flex flex-wrap gap-2 items-center">
            <button className="btn btn-primary text-xs px-2 py-1" onClick={pdf} data-testid="saison-pdf">Als PDF</button>
            <span className="text-2xs text-ink-dim">{fazit ? 'mit Kiras Fazit' : 'noch ohne Fazit – unten erstellen'}</span>
          </div>
          <KiraFazit title={titel} facts={facts} onText={setFazit} />
        </>
      )}
    </Panel>
  );
}

// ─── 14 Touren-Steckbrief ──────────────────────────────────────────────────
const istTour = (a: GarminAkt) => (BERG.has(a.typ || '') && (a.hm || 0) >= 300) || ((a.hm || 0) >= 800 && !RAD_ALLE.has(a.typ || ''));
export function TourenSteckbriefTile() {
  const { tage } = useGarminTage();
  const touren = useMemo(() => eintraege(tage).filter((e) => istTour(e.a)).reverse(), [tage]);
  const [idx, setIdx] = useState(0);
  const t = touren[idx];
  const info = useMemo(() => {
    if (!t || !tage) return null;
    const a = t.a;
    const aufstieg = a.bewegt_s && a.hm ? a.hm / (a.bewegt_s / 3600) : null;
    const lastJe100 = a.last && a.hm ? (a.last / a.hm) * 100 : null;
    const rang = (f: (e: E) => number | null, hoch = true) => {
      const xs = touren.map(f).filter((v): v is number => v != null).sort((p, q) => (hoch ? q - p : p - q));
      const v = f(t);
      return v == null ? null : { platz: xs.indexOf(v) + 1, von: xs.length };
    };
    const basis = (k: 'ready' | 'hrv') => {
      const vs: number[] = [];
      for (let i = 1; i <= 28; i++) { const g = tage.get(plusTage(t.datum, -i)); const v = k === 'hrv' ? g?.hrv?.nacht : null; if (v) vs.push(v); }
      return median(vs);
    };
    const hrvB = basis('hrv');
    const nach = [1, 2].map((n) => { const g = tage.get(plusTage(t.datum, n)); return { n, hrv: g?.hrv?.nacht ?? null, rhr: g?.tag?.rhr ?? null, schlaf: g?.schlaf?.score ?? null }; });
    const medLast = median(touren.map((e) => (e.a.last && e.a.hm ? (e.a.last / e.a.hm) * 100 : NaN)).filter(Number.isFinite));
    const medAuf = median(touren.map((e) => (e.a.bewegt_s && e.a.hm ? e.a.hm / (e.a.bewegt_s / 3600) : NaN)).filter(Number.isFinite));
    const z = a.zonen_s || [];
    const zs = z.reduce((p, x) => p + x, 0) || 1;
    return {
      aufstieg, lastJe100, medLast, medAuf, hrvB, nach, zonen: z.map((x) => Math.round((x / zs) * 100)),
      rangHm: rang((e) => e.a.hm || null), rangAuf: rang((e) => (e.a.bewegt_s && e.a.hm ? e.a.hm / (e.a.bewegt_s / 3600) : null)), rangDauer: rang((e) => e.a.dauer_s || null),
    };
  }, [t, tage, touren]);
  const facts = t && info ? [
    `Tour ${t.a.name || ''} am ${kurz(t.datum)} (${t.a.typ}): ${de(t.a.km, 1)} km, ${de(t.a.hm)} hm, ${dauerTxt(t.a.dauer_s)} (bewegt ${dauerTxt(t.a.bewegt_s)}), Ø HF ${de(t.a.hf_avg)}, max ${de(t.a.hf_max)}, Garmin-Last ${de(t.a.last)}, Body Battery ${vz(t.a.bb)}, Höchster Punkt ${de(t.a.hoehe_max)} m.`,
    `Aufstieg ${de(info.aufstieg)} hm/h (Median deiner ${touren.length} Touren ${de(info.medAuf)}), Last je 100 hm ${de(info.lastJe100, 1)} (Median ${de(info.medLast, 1)}).`,
    `Rang unter deinen Touren: Höhenmeter ${info.rangHm?.platz}/${info.rangHm?.von}, Aufstiegstempo ${info.rangAuf?.platz}/${info.rangAuf?.von}, Dauer ${info.rangDauer?.platz}/${info.rangDauer?.von}.`,
    `Zonen: ${info.zonen.map((p, i) => `Z${i + 1} ${p} %`).join(', ')}.`,
    `Erholung danach: HRV-Basis ${de(info.hrvB)} ms; ${info.nach.map((n) => `Tag +${n.n}: HRV ${de(n.hrv)}, Ruhepuls ${de(n.rhr)}, Schlaf ${de(n.schlaf)}`).join('; ')}.`,
  ] : [];
  const titel = t ? `Touren-Steckbrief ${t.a.name || ''} ${kurz(t.datum)}` : 'Touren-Steckbrief';
  const [fazit, setFazit] = useState<string | null>(null);
  useEffect(() => { setFazit(t ? fazitAusCache(titel, facts) : null); }, [titel, facts.join('|')]); // eslint-disable-line
  function pdf() {
    if (!t || !info) return;
    const a = t.a;
    const k = (lab: string, v: string, sub = '') => `<div class="k"><span>${esc(lab)}</span><b>${esc(v)}</b><span>${esc(sub)}</span></div>`;
    drucken(titel, `<h1>${esc(a.name || 'Tour')}</h1><div class="sub">${kurz(t.datum)} · ${esc(a.ort || '')} · ${esc(a.typ)}</div>
<div class="grid">${k('Distanz', de(a.km, 1) + ' km')}${k('Höhenmeter', de(a.hm) + ' hm', `Rang ${info.rangHm?.platz}/${info.rangHm?.von}`)}${k('Dauer', dauerTxt(a.dauer_s), 'bewegt ' + dauerTxt(a.bewegt_s))}${k('Höchster Punkt', de(a.hoehe_max) + ' m')}
${k('Aufstieg', de(info.aufstieg) + ' hm/h', `Median ${de(info.medAuf)} · Rang ${info.rangAuf?.platz}/${info.rangAuf?.von}`)}${k('Ø / max HF', `${de(a.hf_avg)} / ${de(a.hf_max)}`)}${k('Garmin-Last', de(a.last), `je 100 hm ${de(info.lastJe100, 1)} (Median ${de(info.medLast, 1)})`)}${k('Body Battery', vz(a.bb))}</div>
<h2>Zonen</h2>${svgBalken(info.zonen.map((p, i) => ({ lab: 'Z' + (i + 1), v: p })), '% der Zeit', '#ea580c')}
<h2>Erholung danach</h2><table><tr><th>Tag</th><th>HRV (Basis ${de(info.hrvB)} ms)</th><th>Ruhepuls</th><th>Schlafscore</th></tr>${info.nach.map((n) => `<tr><td>+${n.n}</td><td>${de(n.hrv)}</td><td>${de(n.rhr)}</td><td>${de(n.schlaf)}</td></tr>`).join('')}</table>
${fazit ? `<h2>Kiras Fazit</h2><div class="faz">${esc(fazit)}</div>` : ''}`);
  }
  return (
    <Panel title="Touren-Steckbrief">
      <Hinweis>Karte zu jeder Berg-, Wander- oder Skitour (ab 300 hm): Kennzahlen, Rang unter deinen Touren, Zonen und wie du dich in den zwei Tagen danach erholt hast.</Hinweis>
      {tage === null ? <Laedt /> : !t ? <p className="text-2xs text-ink-dim">Noch keine Tour mit mindestens 300 hm.</p> : info && (
        <>
          <select className="bg-bg border border-border rounded px-2 py-1 text-xs mb-2 max-w-full" value={idx} onChange={(e) => setIdx(Number(e.target.value))}>
            {touren.map((e, i) => <option key={e.datum + i} value={i}>{kurz(e.datum)} · {e.a.name} · {de(e.a.hm)} hm</option>)}
          </select>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2">
            <Kz lab="Höhenmeter" wert={`${de(t.a.hm)} hm`} sub={`${de(t.a.km, 1)} km · Rang ${info.rangHm?.platz}/${info.rangHm?.von}`} />
            <Kz lab="Dauer" wert={dauerTxt(t.a.dauer_s)} sub={`bewegt ${dauerTxt(t.a.bewegt_s)}`} />
            <Kz lab="Aufstieg" wert={`${de(info.aufstieg)} hm/h`} sub={`Median ${de(info.medAuf)} · Rang ${info.rangAuf?.platz}/${info.rangAuf?.von}`} ton={(info.aufstieg ?? 0) > (info.medAuf ?? 0) ? 'gut' : undefined} />
            <Kz lab="Last je 100 hm" wert={de(info.lastJe100, 1)} sub={`Median ${de(info.medLast, 1)} · Garmin-Last ${de(t.a.last)}`} />
          </div>
          <div className="text-2xs text-ink-muted tnum mb-1">Zonen: {info.zonen.map((p, i) => `Z${i + 1} ${p} %`).join(' · ')} · Ø HF {de(t.a.hf_avg)} · Body Battery {vz(t.a.bb)}</div>
          <div className="text-2xs text-ink-muted tnum">Danach (HRV-Basis {de(info.hrvB)} ms): {info.nach.map((n) => `+${n.n} Tag HRV ${de(n.hrv)} · Ruhepuls ${de(n.rhr)} · Schlaf ${de(n.schlaf)}`).join(' | ')}</div>
          <div className="mt-2"><button className="btn btn-primary text-xs px-2 py-1" onClick={pdf}>Als PDF</button></div>
          <KiraFazit title={titel} facts={facts} onText={setFazit} />
        </>
      )}
    </Panel>
  );
}
