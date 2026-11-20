// Abgleich: Rechenkern gegen die Sheet-Werte aus KK_TIMELINE
import fs from 'node:fs';
const { rechneZeile } = await import(process.argv[2]);
const d = JSON.parse(fs.readFileSync('/tmp/tlall.json', 'utf8'));
const S = d.spalten; const rows = d.zeilen.map(z => Object.fromEntries(S.map((s, j) => [s, z[j]])));
const keys = ['fKEI','load_star','ctl_star','ctl_gap','coachE_Smart_Gains','SG_Flags','Monotony7','Strain7','coachE_ATL_morning','coachE_CTL_morning','TE_Balance_Trend'];
const heute = new Date().toISOString().slice(0,10);
const maske = JSON.parse(fs.readFileSync('/tmp/maske.json','utf8')); const istFormel = Object.fromEntries(Object.entries(maske).map(([k,v])=>[k,new Set(v)]));
const stat = Object.fromEntries(keys.map(k => [k, { n: 0, ok: 0, max: 0, bsp: [] }]));
for (let i = 0; i < rows.length; i++) {
  if (rows[i].date > heute) continue; // nur Vergangenheit
  const r = rechneZeile(rows, i, d.konstanten);
  for (const k of keys) {
    if (!istFormel[k]?.has(i)) continue; // feste Werte der Vergangenheit sind Wahrheit, nicht Prüfgegenstand
    const soll = rows[i][k], ist = r[k]; const s = stat[k]; s.n++;
    const leer = v => v === null || v === undefined || v === '';
    let gleich;
    if (leer(soll) || leer(ist)) gleich = leer(soll) && leer(ist);
    else if (typeof soll === 'number' && typeof ist === 'number') {
      const diff = Math.abs(soll - ist); gleich = diff <= Math.max(0.05, Math.abs(soll) * 0.005);
      if (diff > s.max) s.max = diff;
    } else gleich = String(soll).trim() === String(ist).trim();
    if (gleich) s.ok++; else if (s.bsp.length < 4) s.bsp.push(`${rows[i].date}: Sheet ${JSON.stringify(soll)} / Kern ${JSON.stringify(ist)}`);
  }
}
for (const k of keys) { const s = stat[k]; console.log(`${k.padEnd(20)} ${s.ok}/${s.n}  max ${s.max.toFixed(4)}  ${s.bsp.join(' | ')}`); }
