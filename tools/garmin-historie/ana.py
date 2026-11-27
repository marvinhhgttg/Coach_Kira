import json,collections,datetime as dt,statistics as st
d=json.load(open('garmin_hist.json'));ga=d['ga'];es=d['es']
a_,b_,c_,d_=155.91432,0.96845,0.1537,0.2993
load=collections.defaultdict(float);lang={};typ={}
for a in ga:
  day=a[0][:10]
  if a[4] is not None: load[day]+=a[4]
  h=a[1]/3600
  if h>lang.get(day,0): lang[day]=h;typ[day]=a[3]
D0=dt.date(2025,6,1);D1=dt.date(2026,11,20)
days=[(D0+dt.timedelta(i)).isoformat() for i in range((D1-D0).days+1)]
n=len(days);res=[None]*n
for i in range(1,n):
  e0=es.get(days[i-1]);e1=es.get(days[i])
  if e0 is None or e1 is None: continue
  w=[load.get(days[m],0) for m in range(max(0,i-56),i)]
  res[i]=e1-(a_+b_*e0+c_*load.get(days[i],0)+d_*(sum(w)/len(w)))
pro=[]
for i in range(n):
  v=[res[j] for j in range(i,min(n,i+4)) if res[j] is not None]
  if len(v)<2: continue
  pro.append((days[i],lang.get(days[i],0),sum(v),load.get(days[i],0)))
basis=st.mean(x[2] for x in pro)
print('Tage',len(pro),'basis',round(basis,1))
for lab,von,bis in [('<2',0.01,2),('2-3',2,3),('3-4',3,4),('2-4',2,4),('>=4',4,99),('4-6',4,6),('>=6',6,99)]:
  v=[x[2]-basis for x in pro if von<=x[1]<bis]
  print(lab,len(v),round(st.mean(v),1) if v else None, round(st.median(v),1) if v else None)
print('Einzelfälle >=3h:')
for x in pro:
  if x[1]>=3: print(x[0],round(x[1],1),typ[x[0]],'ESS',round(x[3]),'extra',round(x[2]-basis))
# fehlende ES bei langen Einheiten
miss=[k for k,h in lang.items() if h>=3 and k>='2025-08-21' and k not in [p[0] for p in pro]]
print('lange ohne ES-Auswertung:',miss)
