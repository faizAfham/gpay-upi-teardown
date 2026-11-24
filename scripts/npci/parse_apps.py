"""Build data/npci_upi_apps_monthly.csv from the raw NPCI "UPI Applications" responses.

Usage: python parse_apps.py ../../data/npci_upi_apps_raw.json ../../data
Share = app volume / sum of all listed apps (NPCI lists apps above a volume threshold).
"""
import json,re,sys,csv
raw,outdir=sys.argv[1],sys.argv[2]
d=json.load(open(raw))
M=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
def norm(n):
    n=n.replace('#','').replace('*','').strip(); l=n.lower()
    rules=[('google pay','Google Pay'),('phonepe','PhonePe'),('paytm','Paytm'),('navi','Navi'),('super','super.money'),
           ('cred','CRED'),('amazon','Amazon Pay'),('whatsapp','WhatsApp'),('bhim','BHIM'),('fam','FamApp'),
           ('flipkart','Flipkart UPI'),('mobikwik','MobiKwik'),('groww','Groww'),('pop upi','POP'),('slice','slice'),('jio','Jio')]
    for k,v in rules:
        if l.startswith(k) or (k in ('paytm','navi') and k in l): return v
    return 'Other'
months=[]; table={}
for y in range(2020,2027):
    for m in M:
        k=f'{y}-{m}'; v=d.get(k)
        if not v or v['s']!=200 or len(v['t'])<5000: continue
        tot={}; 
        for r in json.loads(v['t'])['data']['results']:
            n=(r['application_name'] or r['upi_application_name'] or '').strip()
            vol=r['total_volume_mn'] or r['upi_volume_mn']
            try: vol=float(str(vol).replace(',',''))
            except: continue
            a=norm(n); tot[a]=tot.get(a,0)+vol
        key=f'{y}-{M.index(m)+1:02d}'; months.append(key); table[key]=tot
apps=['PhonePe','Google Pay','Paytm','Navi','super.money','CRED','Amazon Pay','WhatsApp','BHIM','FamApp','Flipkart UPI','MobiKwik','Groww','POP','slice','Jio','Other']
with open(f'{outdir}/npci_upi_apps_monthly.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(['month','total_mn']+[f'{a}_mn' for a in apps]+[f'{a}_pct' for a in apps])
    for k in months:
        t=table[k]; T=sum(t.values())
        w.writerow([k,round(T,2)]+[round(t.get(a,0),2) for a in apps]+[round(100*t.get(a,0)/T,2) for a in apps])
print(len(months),'months',months[0],'->',months[-1])
