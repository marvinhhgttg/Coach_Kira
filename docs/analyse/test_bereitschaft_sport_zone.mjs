import fs from 'fs';
import * as M from './rm.mjs';
const lines = fs.readFileSync('/tmp/pruef/kk_timeline.csv','utf8').trim().split('\n');
const parse = l => { const out=[];let cur='',q=false; for(const c of l){ if(c=='"'){q=!q;continue} if(c==','&&!q){out.push(cur);cur='';continue} cur+=c } out.push(cur); return out; };
const H = parse(lines[0]); const rows=[]; const seen=new Set();
for (const l of lines.slice(1)) { const v=parse(l); const o={}; H.forEach((h,i)=>o[h]=v[i]); if(seen.has(o.date)) continue; seen.add(o.date); rows.push(o); if(o.is_today=='1') break; }
const num=v=>{ if(v===''||v==null) return null; const n=Number(String(v).replace(',','.')); return Number.isFinite(n)?n:null };
const grp=s=>{s=(s||'').toLowerCase(); return /run|lauf/.test(s)?'L':/bike|rad|gravel|ride|cycl/.test(s)?'R':'S'};
const grau=z=>/Z3|Z4|Z5/i.test(z||'');
for (const win of [180, 424]) {
  const R = rows.slice(-win-1, -1);   // ohne heute
  const H2 = R.map(x=>({date:x.date, load:num(x.load_fb_day), readiness:num(x.Garmin_Training_Readiness), hrv:num(x.hrv_status), acwr:num(x.fbACWR_obs), mono:num(x.Monotony7), _s:grp(x.Sport_x), _g:grau(x.Zone)}));
  M.setExtra(null); const a = M.trainModel(H2,'readiness');
  // Variante B: Extras über Lookup per Last ist nicht eindeutig -> Index über Closure
  let idx=0; const byLoad=[];
  const H3 = H2.map(r=>({...r}));
  // feats wird in dataset in Reihenfolge i=7..n-2 aufgerufen; wir geben über Zähler die Extras
  let order=[]; for(let i=7;i<H3.length-1;i++){const d=H3[i],n=H3[i+1]; if(n.readiness==null||d.readiness==null||d.load==null) continue; order.push([d._s=='L'?d.load:0, d._g?d.load:0]);}
  let k=0; M.setExtra(()=>order[(k++)%order.length]);
  const b = M.trainModel(H3,'readiness');
  console.log(win,'Tage: n',a.n,'MAE alt',a.mae.toFixed(2),'naiv',a.baselineMae.toFixed(2),'| MAE mit Sport/Zone',b.mae.toFixed(2),'naiv',b.baselineMae.toFixed(2), 'Gewichte neu', b.w.slice(-2).map(x=>x.toFixed(2)).join(' / '));
}
