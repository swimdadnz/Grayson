from __future__ import annotations
from pathlib import Path
from datetime import datetime,date
from collections import defaultdict,Counter
import argparse,json,zipfile,tempfile,shutil,re,math,os
from copy import copy
from openpyxl import load_workbook
from name_aliases import canonical_name, canonical_holders
import numpy as np
from scipy.optimize import curve_fit
from scipy.interpolate import UnivariateSpline
ROOT=Path(__file__).resolve().parents[1]; WB=ROOT/'Grayson_Swim_Site_Master.xlsx'; DATA=ROOT/'generated_data'; DEN=365.2425
STROKE={'1':'Free','2':'Back','3':'Breast','4':'Fly','5':'IM'}; COURSE={'L':'LCM','S':'SCM'}
ORDER={'Free':0,'Back':1,'Breast':2,'Fly':3,'IM':4}; MODELS=['Linear','Quadratic','Logarithmic','Asymptotic exponential','Stretched exponential','Smoothing spline']; COMP={'Linear':2,'Quadratic':3,'Logarithmic':2,'Asymptotic exponential':3,'Stretched exponential':4,'Smoothing spline':5}
def dval(v):
    if isinstance(v,datetime): return v.date()
    if isinstance(v,date): return v
    return datetime.strptime(str(v)[:10],'%Y-%m-%d').date()
def fmt(s):
    s=float(s); return f'{int(s//60)}:{s%60:05.2f}' if s>=60 else f'{s:.2f}'
def ro(r,lead=False):
    r=(r or '').upper(); return 3 if r=='FIN' else 2 if r in ('SEM','SEMI') else 1 if r in ('PRE','HEAT') else 4 if lead else 2
def lead(r): return any(str(f.get('label','')).upper()=='R' or 'LEADOFF' in str(f.get('title','')).upper() for f in (r.get('flags') or [])) or (bool(r.get('relay_id')) and int(r.get('legposition') or 0)==1)
def dup(n,d,c,di,st,t,m,r,l): return f'{n}|{d.isoformat()}|{c}|{di}|{st}|{float(t):g}|{m}|{r}|{bool(l)}'
def fit(name,x,y):
    x=np.asarray(x,float); y=np.asarray(y,float); o=10.0; z=x-o
    if name=='Linear': co=np.polyfit(z,y,1); return lambda q:np.polyval(co,np.asarray(q,float)-o)
    if name=='Quadratic': co=np.polyfit(z,y,2); return lambda q:np.polyval(co,np.asarray(q,float)-o)
    if name=='Logarithmic': co=np.polyfit(np.log(z+.25),y,1); return lambda q:np.polyval(co,np.log(np.asarray(q,float)-o+.25))
    if name=='Asymptotic exponential':
        fn=lambda zz,floor,b,k:floor+b*np.exp(-k*zz); p,_=curve_fit(fn,z,y,p0=[max(0,min(y)*.94),max(.1,max(y)-min(y)),.7],bounds=([max(0,min(y)*.55),0,1e-6],[min(y)*1.05,max(y)*3,20]),maxfev=50000); return lambda q:fn(np.asarray(q,float)-o,*p)
    if name=='Stretched exponential':
        fn=lambda zz,floor,b,k,p:floor+b*np.exp(-np.power(np.maximum(k*zz,0),p)); p,_=curve_fit(fn,z,y,p0=[max(0,min(y)*.94),max(.1,max(y)-min(y)),.7,1],bounds=([max(0,min(y)*.55),0,1e-6,.2],[min(y)*1.05,max(y)*3,20,4]),maxfev=100000); return lambda q:fn(np.asarray(q,float)-o,*p)
    oo=np.argsort(x); xx=x[oo]; yy=y[oo]; s=max(len(xx)*np.var(yy)*.35,1e-6); spl=UnivariateSpline(xx,yy,s=s,k=min(3,len(xx)-1)); return lambda q:spl(np.asarray(q,float))
def cv(name,x,y,g):
    e=[]; g=np.asarray(g,dtype=object); x=np.asarray(x,float); y=np.asarray(y,float)
    for z in list(dict.fromkeys(g.tolist())):
        tr=g!=z; te=g==z
        if tr.sum()<max(COMP[name]+1,4): continue
        try: e.extend(np.abs(y[te]-fit(name,x[tr],y[tr])(x[te])).tolist())
        except: pass
    return float(np.mean(e)) if e else None
def analyze(wb):
    ws=wb['01_Swimmers']; people=[]; dob={}
    estimated_names={str(rr[0]) for rr in wb['Age_Estimation_QA'].iter_rows(min_row=2,values_only=True) if rr[0] and 'Estimated DOB' in str(rr[5])} if 'Age_Estimation_QA' in wb.sheetnames else set()
    for r in range(5,ws.max_row+1):
        if not ws.cell(r,2).value: continue
        nm=ws.cell(r,2).value; dob[nm]=dval(ws.cell(r,5).value) if ws.cell(r,5).value else None; people.append({'id':ws.cell(r,1).value,'name':nm,'group':ws.cell(r,6).value,'nation':ws.cell(r,7).value or '—','dobEstimated':nm in estimated_names,'public':ws.cell(r,9).value=='YES' and dob[nm] is not None})
    prof=Counter(); ss=wb['02_Swimmer_Sources']
    for r in range(5,ss.max_row+1): prof[ss.cell(r,2).value]+=1
    # SEIKO's historical individual race results are source-verified; the
    # published PDF print date may not equal race day. Keep original meet
    # window for charts, tooltips, and future DOB/race-day refinements.
    seiko_dates={}
    if '11_SEIKO_Race_Audit' in wb.sheetnames:
        for rr in wb['11_SEIKO_Race_Audit'].iter_rows(min_row=2,values_only=True):
            if rr[0]: seiko_dates[str(rr[0])]={'earliest':str(rr[7]),'latest':str(rr[8]),'dateEvidence':str(rr[9]),'sourceURL':str(rr[10])+'#page='+str(rr[12])}
    sw=wb['04_Swims']; rec=[]
    for r in range(5,sw.max_row+1):
        n=sw.cell(r,3).value; t=sw.cell(r,18).value
        if not n or t in (None,''): continue
        dd=dval(sw.cell(r,7).value); x={'row':r,'swimmer_id':sw.cell(r,2).value,'name':n,'date':dd,'meet':str(sw.cell(r,8).value or ''),'meet_id':str(sw.cell(r,9).value or ''),'round':str(sw.cell(r,11).value or 'TIM'),'course':sw.cell(r,13).value,'distance':int(sw.cell(r,14).value),'stroke':sw.cell(r,15).value,'event':sw.cell(r,16).value,'event_id':sw.cell(r,17).value,'time':float(t),'source':str(sw.cell(r,4).value or 'SwimCloud'),'relay':sw.cell(r,22).value=='YES','legal':sw.cell(r,23).value=='YES','source_age':int(m.group(1)) if (m:=re.search(r'AUS2026\|AGE(\d+)',str(sw.cell(r,6).value or ''))) else None,'seikoDateRange':seiko_dates.get(str(sw.cell(r,6).value))}; x['age']=(dd-dob[n]).days/DEN if dob.get(n) is not None else None; x['rorder']=ro(x['round'],x['relay']); rec.append(x)
    gr=sorted([x for x in rec if x['name']=='Grayson Coulter' and x['legal']],key=lambda x:(x['date'],x['rorder'],x['row'])); prior=defaultdict(list)
    for x in gr:
        p=sum(t<x['time']-1e-12 for t in prior[x['event_id']]); x['top5']=p<5; x['top10']=p<10; prior[x['event_id']].append(x['time'])
    win=[x for x in gr if x['date']>=date(2024,1,1)]; cg=defaultdict(list)
    for x in win: cg[(x['course'],x['event'])].append(x)
    keys=sorted(cg,key=lambda k:(0 if k[0]=='LCM' else 1,ORDER.get(k[1].split()[-1],9),int(k[1].split()[0])))
    summary=[]
    for k in keys:
        a=cg[k]; n=len(a); a5=sum(x['top5'] for x in a); a10=sum(x['top10'] for x in a); summary.append({'course':k[0],'event':k[1],'swims':n,'top5':a5,'top10':a10,'top5Pct':100*a5/n,'top10Pct':100*a10/n})
    ev=wb['03_Events']; req=[]
    for r in range(5,ev.max_row+1):
        if ev.cell(r,7).value=='YES': req.append((ev.cell(r,2).value,ev.cell(r,5).value,ev.cell(r,1).value))
    requested=[('LCM','100 Free'),('LCM','200 Free'),('LCM','400 Free'),('LCM','100 Fly'),('SCM','100 Free'),('SCM','200 Free'),('SCM','400 Free'),('SCM','100 Fly')]; req.sort(key=lambda x:requested.index((x[0],x[1])) if (x[0],x[1]) in requested else 99)
    trends={}
    for c,e,eid in req:
        a=sorted([x for x in win if x['event_id']==eid],key=lambda x:(x['date'],x['rorder'],x['row'])); k=f'{c}|{e}'
        if len(a)<3: trends[k]={'status':'insufficient','swims':[],'modelScores':[]}; continue
        xx=np.array([x['age'] for x in a]); yy=np.array([x['time'] for x in a]); gg=np.array([x['meet_id'] or x['meet'] or x['date'].isoformat() for x in a],dtype=object); scores={m:cv(m,xx,yy,gg) for m in MODELS}; finite={m:s for m,s in scores.items() if s is not None and math.isfinite(s)}; sel=min([m for m,s in finite.items() if s<=min(finite.values())*1.05+1e-9],key=lambda m:COMP[m]) if finite else 'Linear'; f=fit(sel,xx,yy); pr=f(xx); res=yy-pr; grid=np.linspace(float(min(xx)),float(max(xx)),100)
        trends[k]={'status':'ready','selectedModel':sel,'residualSdSeconds':float(np.std(res,ddof=1)),'residualSdPct':float(np.std(res/pr*100,ddof=1)),'maeSeconds':float(np.mean(abs(res))),'maePct':float(np.mean(abs(res/pr*100))),'cvMaeSeconds':scores[sel],'swims':[{'age':round(x['age'],4),'timeSeconds':x['time'],'timeDisplay':fmt(x['time']),'meet':x['meet'],'round':x['round'],'relayLeadOff':x['relay'],'date':x['date'].isoformat()} for x in a],'trend':[{'age':round(float(q),4),'timeSeconds':round(float(v),5)} for q,v in zip(grid,f(grid))],'modelScores':[{'model':m,'cvMaeSeconds':None if scores[m] is None else round(scores[m],4),'selected':m==sel} for m in MODELS],'note':'Sparse dataset: linear descriptive trend shown; cross-validation is not sufficiently supported.' if scores[sel] is None else None}
    return people,prof,rec,gr,win,summary,req,trends
def public(wb):
    people,prof,rec,gr,win,summary,req,trends=analyze(wb); DATA.mkdir(parents=True,exist_ok=True); asof=date.today().isoformat()
    # progression
    series={}
    for c,e,eid in req:
        k=f'{c}|{e}'; arr=[]
        for p in people:
            if not p['public']: continue
            pts=sorted([x for x in rec if x['name']==p['name'] and x['event_id']==eid and x['legal'] and x['age'] is not None and 0<=x['age']<21],key=lambda x:(x['date'],x['rorder'],x['row'])); arr.append({'id':p['id'],'name':p['name'],'group':p['group'],'points':[{'age':round(x['age'],4),'timeSeconds':x['time'],'timeDisplay':fmt(x['time']),'meet':x['meet'],'round':x['round'],'relayLeadOff':x['relay'],'year':x['date'].year,'source':x['source'],'ageGroup':x['source_age'] if x['source_age'] is not None else int(x['age']),'ageEstimated':p['dobEstimated'],'dateApproximate':bool(x['seikoDateRange']),'dateEarliest':x['seikoDateRange']['earliest'] if x['seikoDateRange'] else None,'dateLatest':x['seikoDateRange']['latest'] if x['seikoDateRange'] else None,'sourceURL':x['seikoDateRange']['sourceURL'] if x['seikoDateRange'] else None} for x in pts]})
        series[k]=arr
    (DATA/'progression.json').write_text(json.dumps({'asOf':asof,'status':'ready','ageScope':'performances before exact age 21.0 (through age 20)','swimmers':[{'id':p['id'],'name':p['name'],'group':p['group'],'nation':p['nation'],'dobEstimated':p['dobEstimated'],'sourceProfiles':prof[p['name']]} for p in people if p['public']],'events':[{'course':c,'event':e,'key':f'{c}|{e}'} for c,e,_ in req],'series':series,'notes':['Swimming Australia 2026 published boys Top 10 records are one-off supplements; Ian Thorpe full SwimCloud history has not been acquired.','Grayson Coulter combines two SwimCloud profiles into one canonical swimmer.','SEIKO Japanese historical results have verified times and meet-date windows. Their plotted dates are representative estimates, not exact swim dates; every source window stays within one age-year.','Birth dates never appear in public JSON; some Australian athletes have DOB estimates, flagged individually. The published Australian record age governs age-year best tables.']},indent=2),encoding='utf-8')
    (DATA/'consistency.json').write_text(json.dumps({'asOf':asof,'status':'ready','analysisStart':'2024-01-01','summary':[{k:v for k,v in r.items() if k not in ('top5','top5Pct')} for r in summary],'trends':trends,'modelCandidates':MODELS,'scopeNote':'Top-10 counts cover every Grayson event/course captured since 1 January 2024. Trend charts remain limited to the eight requested event-course combinations.','dataNotes':['Equal times are not counted as faster.','Relay lead-offs are treated as valid event swims.']},indent=2),encoding='utf-8')
    # records from workbook
    rb=wb['05_Record_Benchmarks']; counts={'male':defaultdict(lambda:{'SCM':0,'LCM':0}),'female':defaultdict(lambda:{'SCM':0,'LCM':0})}; tops={'male':[],'female':[]}
    def sec(v):
        if v in (None,''): return None
        s=str(v); return float(s.split(':')[0])*60+float(s.split(':')[1]) if ':' in s else float(s)
    for r in range(5,rb.max_row+1):
        key,c,sx,ag,e,h=[rb.cell(r,z).value for z in range(1,7)]; rt=rb.cell(r,9).value; nxt=rb.cell(r,16).value; nt=rb.cell(r,17).value
        if h and h!='Standard':
            for n in canonical_holders(h): counts[str(sx).lower()][n][c]+=1
        a,b=sec(rt),sec(nt); mar=(b-a)/b*100 if a is not None and b else None
        if rb.cell(r,15).value=='YES' and nxt and mar is not None: tops[str(sx).lower()].append({'swimmer':' / '.join(canonical_holders(h)),'event':e,'course':c,'age':ag,'recordTime':rt,'nextFastestSwimmer':canonical_name(nxt),'nextFastestTime':nt,'marginPct':round(mar,4),'recordKey':key})
    oc={}
    for sx,d in counts.items():
        arr=[{'swimmer':n,'scm':v['SCM'],'lcm':v['LCM'],'total':v['SCM']+v['LCM']} for n,v in d.items()]; arr.sort(key=lambda x:(-x['total'],-x['scm'],-x['lcm'],x['swimmer'])); last=None; rk=0
        for i,x in enumerate(arr,1):
            if x['total']!=last: rk=i; last=x['total']
            x['rank']=rk
        oc[sx]=arr; tops[sx].sort(key=lambda x:(-x['marginPct'],x['recordKey'])); tops[sx]=tops[sx][:20]
        for i,x in enumerate(tops[sx],1): x['rank']=i
    (DATA/'records.json').write_text(json.dumps({'asOf':asof,'scope':'NZ individual age-group records, ages 13-18; SCM and LCM','counts':oc,'top20':tops},indent=2),encoding='utf-8')
    gscm=sum(1 for r in range(5,rb.max_row+1) if rb.cell(r,2).value=='SCM' and 'Grayson Coulter' in [x.strip() for x in str(rb.cell(r,6).value or '').split(' / ')]); glcm=sum(1 for r in range(5,rb.max_row+1) if rb.cell(r,2).value=='LCM' and 'Grayson Coulter' in [x.strip() for x in str(rb.cell(r,6).value or '').split(' / ')])
    (DATA/'home.json').write_text(json.dumps({'asOf':asof,'graysonRecords':{'scm':gscm,'lcm':glcm,'total':gscm+glcm}},indent=2),encoding='utf-8')
    for p in DATA.glob('*.json'):
        if re.search(r'\b(19|20)\d\d-\d\d-\d\d\b',p.read_text(encoding='utf-8')) and '"date"' not in p.read_text(encoding='utf-8'): pass

def import_zip(wb,zpath):
    # Append/update SwimCloud rows from a combined acquisition package. Historical rows are preserved.
    ss=wb['02_Swimmer_Sources']; pmap={str(ss.cell(r,4).value): (ss.cell(r,1).value,ss.cell(r,2).value) for r in range(5,ss.max_row+1) if ss.cell(r,4).value}
    ev=wb['03_Events']; em={(ev.cell(r,2).value,int(ev.cell(r,3).value),ev.cell(r,4).value):ev.cell(r,1).value for r in range(5,ev.max_row+1) if ev.cell(r,1).value}
    sw=wb['04_Swims']
    # Five certified 2024 Queensland swims were supplemented once because they were
    # missing in SwimCloud. If SwimCloud later publishes the same results, do not
    # duplicate them merely because the displayed meet or round differs.
    qld_supplements={(str(sw.cell(i,3).value), str(sw.cell(i,13).value),
                      str(sw.cell(i,16).value), dval(sw.cell(i,7).value).isoformat(),
                      round(float(sw.cell(i,18).value),2))
                     for i in range(5,sw.max_row+1)
                     if sw.cell(i,4).value in ('Swimming Queensland All-Comers records','Swimming Australia Age All-Time Top 10')
                     and sw.cell(i,7).value and sw.cell(i,18).value is not None}
    pairs={(str(sw.cell(r,5).value),str(sw.cell(r,6).value)) for r in range(5,sw.max_row+1) if sw.cell(r,5).value and sw.cell(r,6).value}
    if '10_Acquisition_Race_IDs' in wb.sheetnames:
        pairs.update((str(x[0]),str(x[1])) for x in wb['10_Acquisition_Race_IDs'].iter_rows(min_row=2,values_only=True) if x[0] and x[1])
    dups={str(sw.cell(r,30).value) for r in range(5,sw.max_row+1) if sw.cell(r,30).value}; mx=max(int(str(sw.cell(r,1).value)[2:]) for r in range(5,sw.max_row+1) if str(sw.cell(r,1).value).startswith('SM'))
    with zipfile.ZipFile(zpath) as z:
        for member in z.namelist():
            if not member.lower().endswith('.json') or member.endswith('summary.json'):continue
            try: obj=json.loads(z.read(member).decode('utf-8-sig'))
            except: continue
            if not isinstance(obj,dict) or 'body' not in obj or not obj.get('ok'): continue
            m=re.search(r'_(\d{6,})[/\\]',member); pid=m.group(1) if m else None
            if pid not in pmap: continue
            sid,nm=pmap[pid]
            try: arr=json.loads(obj['body'])
            except: continue
            for rr in arr:
                rid=str(rr.get('id'))
                if (pid,rid) in pairs: continue
                c=COURSE.get(str(rr.get('eventcourse'))); di=int(rr.get('eventdistance')); st=STROKE.get(str(rr.get('eventstroke')))
                if (c,di,st) not in em: continue
                dd=dval(rr.get('dateofswim')); meet=str(rr.get('name') or ''); rnd=str(rr.get('eventround') or 'TIM'); ll=lead(rr); tt=float(rr.get('eventtime')); dk=dup(nm,dd,c,di,st,tt,meet,rnd,ll)
                if (nm,c,f'{di} {st}',dd.isoformat(),round(tt,2)) in qld_supplements:continue
                if dk in dups: continue
                row=sw.max_row+1
                for col in range(1,sw.max_column+1): sw.cell(row,col)._style=copy(sw.cell(5,col)._style)
                mx+=1; vals={1:f'SM{mx:06d}',2:sid,3:nm,4:'SwimCloud',5:pid,6:rid,7:datetime.combine(dd,datetime.min.time()),8:meet,9:str(rr.get('meet_id') or ''),11:rnd,13:c,14:di,15:st,16:f'{di} {st}',17:em[(c,di,st)],18:tt,22:'YES' if ll else 'NO',23:'YES' if rr.get('legal',True) else 'NO',24:'YES' if rr.get('exhibition',False) else 'NO',25:rr.get('heat'),26:rr.get('lane'),27:rr.get('place'),28:'https://www.swimcloud.com'+str(rr.get('url') or ''),29:'Interactive SwimCloud update',30:dk,36:'OK'}
                for cc,v in vals.items():sw.cell(row,cc).value=v
                pairs.add((pid,rid));dups.add(dk)
    last=sw.max_row
    swimmer_last=wb['01_Swimmers'].max_row
    for r in range(5,last+1):
        sw.cell(r,12).value=f'=IF(K{r}="FIN",3,IF(OR(K{r}="SEM",K{r}="SEMI"),2,IF(OR(K{r}="PRE",K{r}="HEAT"),1,IF(V{r}="YES",4,2))))'; sw.cell(r,19).value=f'=IF(R{r}="","",IF(R{r}>=60,INT(R{r}/60)&":"&TEXT(MOD(R{r},60),"00.00"),TEXT(R{r},"0.00")))'; sw.cell(r,20).value=f'=IFERROR((G{r}-INDEX(\'01_Swimmers\'!$E$5:$E${swimmer_last},MATCH(B{r},\'01_Swimmers\'!$A$5:$A${swimmer_last},0)))/\'00_Control\'!$B$10,"")'; sw.cell(r,21).value=f'=IF(T{r}="","",INT(T{r}))'; sw.cell(r,31).value=f'=G{r}+L{r}/1000+ROW()/1000000000'; sw.cell(r,32).value=f'=IF(C{r}<>"Grayson Coulter","",COUNTIFS($C$5:$C${last},"Grayson Coulter",$Q$5:$Q${last},Q{r},$AE$5:$AE${last},"<="&AE{r},$R$5:$R${last},"<"&R{r}))'; sw.cell(r,33).value=f'=IF(OR(C{r}<>"Grayson Coulter",G{r}<\'00_Control\'!$B$7),"",--(AF{r}<5))'; sw.cell(r,34).value=f'=IF(OR(C{r}<>"Grayson Coulter",G{r}<\'00_Control\'!$B$7),"",--(AF{r}<10))'; sw.cell(r,35).value=f'=IF(AND(C{r}="Grayson Coulter",G{r}>=\'00_Control\'!$B$7),"YES","NO")'

def _copy_style(ws,src,dst,maxcol):
    for c in range(1,maxcol+1):
        ws.cell(dst,c)._style=copy(ws.cell(src,c)._style)
        ws.cell(dst,c).number_format=ws.cell(src,c).number_format
        ws.cell(dst,c).alignment=copy(ws.cell(src,c).alignment)

def refresh_record_count_sheet(wb):
    """Rebuild canonical record holder roster and Excel formulas on every refresh.

    Prevent duplicate historical aliases and ensure newly emerging holders appear.
    One row per canonical name; tied records credited once to each named holder.
    """
    rb=wb['05_Record_Benchmarks'];counts=defaultdict(lambda:[0,0])
    for r in range(5,rb.max_row+1):
        if not rb.cell(r,1).value: continue
        gender=rb.cell(r,3).value;course=rb.cell(r,2).value
        if gender not in ('Male','Female') or course not in ('SCM','LCM'): continue
        holders=canonical_holders(rb.cell(r,6).value)
        rb.cell(r,7).value=holders[0] if holders else None
        rb.cell(r,8).value=holders[1] if len(holders)>1 else None
        rb.cell(r,6).value=' / '.join(holders) if holders else 'Standard'
        for swimmer in holders:counts[(gender,swimmer)][0 if course=='SCM' else 1]+=1
    people=sorted(counts,key=lambda p:(0 if p[0]=='Male' else 1,-sum(counts[p]),-counts[p][0],p[1]))
    ws=wb['A_Record_Counts']; last=max(ws.max_row,5+len(people)); final=4+len(people)
    # Store source style before clearing worksheet cells.
    for r in range(5,last+1):
        for c in range(1,8):ws.cell(r,c).value=None
    for r,(sex,swimmer) in enumerate(people,5):
        if r>5:_copy_style(ws,5,r,7)
        ws.cell(r,1).value=sex; ws.cell(r,2).value=swimmer
        for c,course in ((3,'SCM'),(4,'LCM')):
            ws.cell(r,c).value=(f'=COUNTIFS(05_Record_Benchmarks!$B:$B,"{course}",05_Record_Benchmarks!$C:$C,A{r},'
                                 f'05_Record_Benchmarks!$G:$G,B{r})+COUNTIFS(05_Record_Benchmarks!$B:$B,"{course}",'
                                 f'05_Record_Benchmarks!$C:$C,A{r},05_Record_Benchmarks!$H:$H,B{r})')
        ws.cell(r,5).value=f'=C{r}+D{r}'
        ws.cell(r,6).value=f'=1+COUNTIFS($A$5:$A${final},A{r},$E$5:$E${final},">"&E{r})'
        ws.cell(r,7).value=f'=A{r}&"|"&TEXT(F{r},"000")&"|"&B{r}'

def refresh_analysis_sheets(wb):
    refresh_record_count_sheet(wb)
    people,prof,rec,gr,win,summary,req,trends=analyze(wb)
    # Consistency table: every captured Grayson event/course since 1 Jan 2024.
    ws=wb['A_Consistency']
    for r in range(5,ws.max_row+1):
        for c in range(1,ws.max_column+1): ws.cell(r,c).value=None
    ws['A2']='Analysis starts 1 January 2024. The table covers every Grayson event/course captured in the all-event history; Top-10 status is historical at the time of each race (Top-5 internal audit columns retained). Trend metrics are shown only for the eight requested chart events.'
    for i,a in enumerate(summary,start=5):
        if i>5:_copy_style(ws,5,i,14)
        c,e=a['course'],a['event']; ws.cell(i,1).value=c; ws.cell(i,2).value=e
        ws.cell(i,3).value=f'=COUNTIFS(\'04_Swims\'!$C:$C,"Grayson Coulter",\'04_Swims\'!$M:$M,A{i},\'04_Swims\'!$P:$P,B{i},\'04_Swims\'!$G:$G,">="&\'00_Control\'!$B$7)'
        ws.cell(i,4).value=f'=SUMIFS(\'04_Swims\'!$AG:$AG,\'04_Swims\'!$C:$C,"Grayson Coulter",\'04_Swims\'!$M:$M,A{i},\'04_Swims\'!$P:$P,B{i},\'04_Swims\'!$G:$G,">="&\'00_Control\'!$B$7)'
        ws.cell(i,5).value=f'=IF(C{i}=0,"",D{i}/C{i})'
        ws.cell(i,6).value=f'=SUMIFS(\'04_Swims\'!$AH:$AH,\'04_Swims\'!$C:$C,"Grayson Coulter",\'04_Swims\'!$M:$M,A{i},\'04_Swims\'!$P:$P,B{i},\'04_Swims\'!$G:$G,">="&\'00_Control\'!$B$7)'
        ws.cell(i,7).value=f'=IF(C{i}=0,"",F{i}/C{i})'
        t=trends.get(f'{c}|{e}')
        if t and t.get('status')=='ready':
            ws.cell(i,8).value=t['selectedModel'];ws.cell(i,9).value=t['residualSdSeconds'];ws.cell(i,10).value=t['residualSdPct']/100;ws.cell(i,11).value=t['maeSeconds'];ws.cell(i,12).value=t['maePct']/100;ws.cell(i,13).value=t['cvMaeSeconds'];ws.cell(i,14).value=t.get('note')
        for cc in (5,7,10,12):ws.cell(i,cc).number_format='0.0%'
        for cc in (9,11,13):ws.cell(i,cc).number_format='0.00'
    # Model comparison rows.
    tm=wb['A_Trend_Models']
    for r in range(5,tm.max_row+1):
        for c in range(1,tm.max_column+1):tm.cell(r,c).value=None
    rr=5
    for c,e,eid in req:
        t=trends.get(f'{c}|{e}')
        if not t or t.get('status')!='ready':continue
        for m in t['modelScores']:
            if rr>5:_copy_style(tm,5,rr,7)
            vals=[c,e,m['model'],m['cvMaeSeconds'],'YES' if m['selected'] else 'NO',COMP[m['model']],None]
            for cc,v in enumerate(vals,1):tm.cell(rr,cc).value=v
            tm.cell(rr,4).number_format='0.000';rr+=1
    # Progression coverage.
    pc=wb['A_Progression_Coverage']
    for r in range(5,pc.max_row+1):
        for c in range(1,pc.max_column+1):pc.cell(r,c).value=None
    rr=5
    for p in people:
        for c,e,eid in req:
            pts=[x for x in rec if x['name']==p['name'] and x['event_id']==eid and x['legal'] and x['age'] is not None and 0<=x['age']<21]
            if rr>5:_copy_style(pc,5,rr,9)
            vals=[p['name'],p['group'],c,e,len(pts),min([x['age'] for x in pts],default=None),max([x['age'] for x in pts],default=None),min([x['date'].year for x in pts],default=None),max([x['date'].year for x in pts],default=None)]
            for cc,v in enumerate(vals,1):pc.cell(rr,cc).value=v
            rr+=1
    # QA summary.
    qa=wb['QA_Checks']; rb=wb['05_Record_Benchmarks']; valid_ids={p['id'] for p in people}; bad=sum(1 for x in rec if x['swimmer_id'] not in valid_ids)
    gscm=sum(1 for r in range(5,rb.max_row+1) if rb.cell(r,2).value=='SCM' and 'Grayson Coulter' in [x.strip() for x in str(rb.cell(r,6).value or '').split(' / ')]);glcm=sum(1 for r in range(5,rb.max_row+1) if rb.cell(r,2).value=='LCM' and 'Grayson Coulter' in [x.strip() for x in str(rb.cell(r,6).value or '').split(' / ')])
    for r in range(5,qa.max_row+1):
        for c in range(1,qa.max_column+1):qa.cell(r,c).value=None
    checks=[('Canonical swimmers',len(people),52,'PASS' if len(people)==52 else 'FAIL','Grayson has two source profiles but one canonical swimmer.'),('Source profiles and published reference entries',sum(prof.values()),'>= 52','PASS' if sum(prof.values())>=52 else 'FAIL',None),('Swimmer ID referential integrity',bad,0,'PASS' if bad==0 else 'FAIL',None),('Normalized swims',len(rec),'>= 2,987','PASS' if len(rec)>=2987 else 'REVIEW',None),('Grayson all-event rows',len(gr),'>= 247','PASS' if len(gr)>=247 else 'REVIEW',None),('Grayson consistency-window swims',len(win),'>= 216','PASS' if len(win)>=216 else 'REVIEW','1 Jan 2024 through current dataset.'),('Grayson event/course combinations since 2024',len(summary),'>= 20','PASS' if len(summary)>=20 else 'REVIEW',None),('Grayson current NZ age records - SCM',gscm,15,'PASS' if gscm==15 else 'REVIEW',None),('Grayson current NZ age records - LCM',glcm,14,'PASS' if glcm==14 else 'REVIEW',None)]
    for i,row in enumerate(checks,start=5):
        if i>5:_copy_style(qa,5,i,5)
        for cc,v in enumerate(row,1):qa.cell(i,cc).value=v
    wb['00_Control']['B6']=datetime.combine(date.today(),datetime.min.time())

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('acquisitions',nargs='*'); ap.add_argument('--no-import',action='store_true'); a=ap.parse_args(); wb=load_workbook(WB)
    if not a.no_import:
        for src in a.acquisitions: import_zip(wb,Path(src))
    try: wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'
    except: pass
    refresh_analysis_sheets(wb)
    wb.save(WB); public(wb); print('Workbook, analytical sheets and public JSON refreshed.')
if __name__=='__main__': main()
