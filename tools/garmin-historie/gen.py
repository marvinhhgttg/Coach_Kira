import json,collections,datetime as dt,statistics as st
exec(open('ana.py').read().split("print('Tage'")[0])
SP={'road_biking':'Rad','gravel_cycling':'Rad','mountain_biking':'Rad','cycling':'Rad','virtual_ride':'Rad','indoor_cycling':'Rad',
 'hiking':'Wandern','walking':'Wandern','mountaineering':'Wandern','backcountry_skiing':'Skitour','backcountry_skiing_snowboarding_ws':'Skitour',
 'running':'Lauf','trail_running':'Lauf','treadmill_running':'Lauf','skate_skiing_ws':'Skike'}
raten=collections.defaultdict(list)
for a in ga:
  s=SP.get(a[3]); 
  if not s or a[4] is None or a[4]<=0 or a[1]<1.5*3600: continue
  h=a[1]/3600; z=[x or 0 for x in a[8:13]]; tz=sum(z); z12=(z[0]+z[1])/tz if tz else None
  if s!='Skitour' and z12 is not None and z12<0.85: continue
  v=a[4]/h
  if v<10 and s=='Rad': continue
  raten[s].append([a[0][:10],round(v,1)])
kl=[]
for lab,von,bis in [('unter 2 h',0.01,2),('2–3 h',2,3),('3–4 h',3,4),('ab 4 h',4,99)]:
  v=[x[2] for x in pro if von<=x[1]<bis]
  kl.append({'lab':lab,'von':von,'bis':bis,'n':len(v),'summe':round(sum(v),1)})
faelle=[[x[0],round(x[1],1),round(x[2]-basis)] for x in pro if x[1]>=3]
out={'stand':'2026-11-20','esAb':min(es),'aktAb':min(a[0][:10] for a in ga),'basisTage':len(pro),'basisSumme':round(sum(x[2] for x in pro),1),'klassen':kl,'faelle':faelle,'raten':raten}
print({k:len(v) for k,v in raten.items()})
ts='''/**
 * Garmin-Historie (aus Garmin Connect gelesen am 21.11.2026, Stand 20.11.2026).
 * – Endurance Score täglich seit {esAb} (vorher gibt es bei Marc keinen Score), Last (ESS) je Aktivität seit 2020.
 * – klassen: Residuum des Fortschreibungsmodells ES_MODEL in den Tagen 0–3 nach dem Tag (Summe), nach Dauer der längsten Einheit.
 * – raten: ESS pro Stunde ruhiger Einheiten ab 1,5 h (Z1+Z2 ≥ 85 %, Skitour ohne Zonenfilter), je Sport.
 * Neue Tage nach „stand“ kommen live aus Timeline und Garmin-Tagesdetails dazu. Erzeugt von tools/garmin-historie.
 */
export type HistKlasse = {{ lab: string; von: number; bis: number; n: number; summe: number }};
export const GARMIN_HISTORIE: {{ stand: string; esAb: string; aktAb: string; basisTage: number; basisSumme: number; klassen: HistKlasse[]; faelle: [string, number, number][]; raten: Record<string, [string, number][]> }} = {js};
'''.format(esAb=out['esAb'],js=json.dumps(out,ensure_ascii=False))
open('/home/user/workspace/coach-kira-v2/coach-kira-dashboard/src/lib/garminHistorie.ts','w').write(ts)
