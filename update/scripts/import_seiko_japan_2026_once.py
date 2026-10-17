#!/usr/bin/env python3
"""One-off, idempotent SEIKO Japan 2026 historical audit import.

This is deliberately NOT part of recurring SwimCloud acquisitions. Source CSV
is first-party verified PDF race candidates. The event day is represented by a
PDF print date or official meet-window midpoint and is NOT an exact race date.
The min/max dates are recorded privately for reproducibility and passed to
Progression as an approximate-date interval. Heats and finals are separate.
"""
from __future__ import annotations
from pathlib import Path
from datetime import datetime, date
from copy import copy
import csv, re, hashlib
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

ROOT=Path(__file__).resolve().parents[1]
BOOK=ROOT/'Grayson_Swim_Site_Master.xlsx'
SOURCE=ROOT/'input'/'SEIKO_four_swimmers_verified_2026.csv'
DOB_CSV=ROOT/'input'/'Australian_DOB_corrections_approved.csv'
KEEP={'Kosei Kato':('2003-10-08','Historical champion'),
      'Koki Tanaka':('1999-07-06','Historical champion'),
      'Ryunosuke Murata':('2011-12-03','Current age group star'),
      'Koki Fujiwara':('2012-07-16','Current age group star')}
MOVE=['Elijah Winnington','Flynn Southam','Jack Alexy','Jacob Whittle',
      'Lewis Clareburt','Mack Horton','Samuel Short']
ROUNDS={'Final':'FIN','Heat':'HEAT','Semi':'SEM','Timed final':'TIM'}


def iso(s):return datetime.strptime(s,'%Y-%m-%d').date()

def age(d,b):return d.year-b.year-((d.month,d.day)<(b.month,b.day))

def sid(row):
    raw='|'.join([row['swimmer'],row['source_url'],row['pdf_page'],row['course'],row['event'],row['round'],row['time']])
    return 'SEIKO26_'+hashlib.sha256(raw.encode('utf-8')).hexdigest()[:21]

def refresh():
    with SOURCE.open(encoding='utf-8-sig',newline='') as f: data=list(csv.DictReader(f))
    with DOB_CSV.open(encoding='utf-8-sig',newline='') as f: corrections=list(csv.DictReader(f))
    assert len(data)==148, f'Expected 148 source-verified Japanese races; saw {len(data)}'
    assert set(d['swimmer'] for d in data)==set(KEEP)
    assert len({sid(d) for d in data})==148,'Ambiguous or duplicated race keys'
    assert len(corrections)==6
    checks=[]
    for d in data:
        assert d['race_validation'].startswith('Source PDF name/birth year/event/course/round/time matched')
        assert d['birth_year_validation'].startswith('Printed birth year matches')
        assert d['course'] in ('LCM','SCM') and d['round'] in ROUNDS
        assert d['stroke'] in ('Free','Fly')
        assert 0<float(d['seconds'])<2000
        assert len(d['source_url'])>30 and d['source_url'].startswith('https://swim.seiko.co.jp/')
        assert int(d['distance_m']) in (50,100,200,400,800,1500)
        lo,hi,rep=map(iso,(d['event_date_earliest'],d['event_date_latest'],d['representative_date']))
        assert lo<=rep<=hi
        birth=iso(KEEP[d['swimmer']][0]);assert birth.year==int(d['dob_reference'][:4])
        assert age(lo,birth)==age(hi,birth),f"Swim window straddles age-year: {d['swimmer']} {d['source_pdf']} page {d['pdf_page']}"
        checks.append((d,lo,hi,rep))
    wb=load_workbook(BOOK)
    persons=wb['01_Swimmers']; sources=wb['02_Swimmer_Sources']; swims=wb['04_Swims']; ev=wb['03_Events']
    byname={str(persons.cell(i,2).value):i for i in range(5,persons.max_row+1) if persons.cell(i,2).value}
    for n in MOVE:
        if n not in byname:raise ValueError('Category move missing swimmer '+n)
        persons.cell(byname[n],6).value='Other fast buggers'
    for n in ['Akaru Nikkawa','Riku Hamada','Akira Ozawa']:
        if n in byname:persons.cell(byname[n],9).value='NO'
    correction_sources={}
    ageqa=wb['Age_Estimation_QA']
    for d in corrections:
        n=d['Swimmer'];dob=iso(d['Proposed DOB'])
        if n not in byname:raise ValueError('Australian DOB target missing '+n)
        persons.cell(byname[n],5).value=datetime.combine(dob,datetime.min.time())
        for row in range(1,ageqa.max_row+1):
            if ageqa.cell(row,1).value==n:
                ageqa.cell(row,3).value=datetime.combine(dob,datetime.min.time())
                if n=='Hantang Wang':lower,upper='2011-09-13','2011-10-01'
                elif n=='Koa Stotz':lower,upper='2010-08-11','2010-09-28'
                else:lower,upper=d['Proposed DOB'],d['Proposed DOB']
                ageqa.cell(row,4).value=datetime.combine(iso(lower),datetime.min.time())
                ageqa.cell(row,5).value=datetime.combine(iso(upper),datetime.min.time())
                ageqa.cell(row,6).value=('Estimated DOB - '+d['Provenance'] if d['Precision'] in ('ESTIMATED','PROVISIONAL')
                                           else 'Verified DOB - '+d['Provenance'])
                ageqa.cell(row,8).value=d['Precision']+'; '+d['Provenance']
                break
        correction_sources[n]=(d['Proposed DOB'],d['Precision'],d['Provenance'])
    for name,(dob,group) in KEEP.items():
        if name not in byname:
            row=persons.max_row+1
            for c in range(1,persons.max_column+1):persons.cell(row,c)._style=copy(persons.cell(5,c)._style)
            new=f'SW{row-4:03d}'
            for c,v in {1:new,2:name,3:name,4:'M',5:datetime.combine(iso(dob),datetime.min.time()),6:group,7:'JPN',9:'YES'}.items():persons.cell(row,c).value=v
            byname[name]=row
        else:
            row=byname[name]
            if str(persons.cell(row,7).value)!='JPN':raise ValueError('Nationality conflict for '+name)
            persons.cell(row,5).value=datetime.combine(iso(dob),datetime.min.time())
            persons.cell(row,6).value=group
            persons.cell(row,9).value='YES'
    eventmap={(str(ev.cell(i,2).value),str(ev.cell(i,5).value)):ev.cell(i,1).value for i in range(5,ev.max_row+1) if ev.cell(i,1).value}
    prof=set((str(sources.cell(i,2).value),str(sources.cell(i,3).value)) for i in range(5,sources.max_row+1))
    for name in KEEP:
        if (name,'SEIKO Japan Results') in prof:continue
        row=sources.max_row+1
        for c in range(1,sources.max_column+1): sources.cell(row,c)._style=copy(sources.cell(5,c)._style)
        for c,v in {1:persons.cell(byname[name],1).value,2:name,3:'SEIKO Japan Results',4:'SEIKO_'+re.sub('[^A-Z]','_',name.upper()),5:'https://swim.seiko.co.jp/',6:'NO',7:'Historical archived PDF result rows imported one time (race date is estimated within verified window)'}.items():sources.cell(row,c).value=v
    audit=wb['11_SEIKO_Race_Audit'] if '11_SEIKO_Race_Audit' in wb.sheetnames else wb.create_sheet('11_SEIKO_Race_Audit')
    fields=['Source Race ID','Swimmer','Event','Course','Round','Time seconds','Representative date (not exact)','Earliest meet date','Latest meet date','Date evidence','Source PDF URL','PDF filename','Page number','Race validation','Status']
    if audit.max_row==1 and audit['A1'].value is None:
        for c,v in enumerate(fields,1):audit.cell(1,c).value=v
        audit.freeze_panes='A2';audit.sheet_view.showGridLines=False
        for cell in audit[1]:cell.fill=PatternFill('solid',fgColor='092B4B');cell.font=Font(color='FFFFFF',bold=True)
        for col,width in {'A':33,'B':22,'C':16,'D':12,'E':12,'F':17,'G':29,'H':23,'I':23,'J':60,'K':75,'L':42,'M':15,'N':65,'O':24}.items():audit.column_dimensions[col].width=width
    seen={str(audit.cell(i,1).value) for i in range(2,audit.max_row+1) if audit.cell(i,1).value}
    loaded={str(swims.cell(i,6).value) for i in range(5,swims.max_row+1) if swims.cell(i,6).value}
    signature=set()
    for i in range(5,swims.max_row+1):
        d=swims.cell(i,7).value; t=swims.cell(i,18).value
        if isinstance(d,datetime) and t is not None:
            signature.add((str(swims.cell(i,3).value),str(swims.cell(i,13).value),str(swims.cell(i,16).value),d.date(),str(swims.cell(i,11).value),round(float(t),2)))
    mx=max(int(str(swims.cell(i,1).value)[2:]) for i in range(5,swims.max_row+1) if str(swims.cell(i,1).value).startswith('SM'))
    appended=0
    for d,lo,hi,rep in checks:
        race_id=sid(d);name=d['swimmer'];event=d['event'];course=d['course'];t=round(float(d['seconds']),2)
        for col in ['distance_m','stroke']:
            assert d[col]
        if (course,event) not in eventmap:raise ValueError(f'Missing {course} {event} in event dimension')
        source_sig=(name,course,event,rep,ROUNDS[d['round']],t)
        if race_id in loaded:continue
        # Different sources may use an identical reported swim on the same date.
        if source_sig in signature:
            continue
        r=swims.max_row+1;mx+=1
        for c in range(1,swims.max_column+1):swims.cell(r,c)._style=copy(swims.cell(5,c)._style)
        stub=f'{name}|{course}|{event}|{rep}|{d["round"]}|{d["time"]}|{d["source_pdf"]}|p{d["pdf_page"]}'
        vals={1:f'SM{mx:06d}',2:persons.cell(byname[name],1).value,3:name,4:'SEIKO Japan Results',5:'SEIKO Japan',6:race_id,
              7:datetime.combine(rep,datetime.min.time()),8:d['source_pdf'].removesuffix('.pdf').replace('_',' '),
              9:d['source_pdf'],11:ROUNDS[d['round']],13:course,14:int(d['distance_m']),15:d['stroke'],16:event,
              17:eventmap[(course,event)],18:t,22:'NO',23:'YES',24:'NO',28:d['source_url']+'#page='+d['pdf_page'],
              29:'One-off verified SEIKO 2026; meet date within source range, representative date NOT exact',30:stub,36:'DATE_ESTIMATED_SEIKO'}
        for c,v in vals.items():swims.cell(r,c).value=v
        # Match formulas to existing 36-column normalized data schema (future refresh regenerates other sheets).
        last=persons.max_row
        for c,v in {12:f'=IF(K{r}="FIN",3,IF(OR(K{r}="SEM",K{r}="SEMI"),2,IF(OR(K{r}="PRE",K{r}="HEAT"),1,IF(V{r}="YES",4,2))))',
            19:f'=IF(R{r}="","",IF(R{r}>=60,INT(R{r}/60)&":"&TEXT(MOD(R{r},60),"00.00"),TEXT(R{r},"0.00")))',
            20:f'=IFERROR((G{r}-INDEX(\'01_Swimmers\'!$E$5:$E${last},MATCH(B{r},\'01_Swimmers\'!$A$5:$A${last},0)))/\'00_Control\'!$B$10,"")',
            21:f'=IF(T{r}="","",INT(T{r}))',31:f'=G{r}+L{r}/1000+ROW()/1000000000'}.items():swims.cell(r,c).value=v
        loaded.add(race_id);signature.add(source_sig);appended+=1
        if race_id not in seen:
            audit.append([race_id,name,event,course,d['round'],t,rep.isoformat(),lo.isoformat(),hi.isoformat(),d['date_precision'],d['source_url'],d['source_pdf'],int(d['pdf_page']),d['race_validation'],'VERIFIED_RESULT_DATE_APPROXIMATE'])
            seen.add(race_id)
    # Add source references for matched races that already existed in this workbook.
    # Do not duplicate the source-result record in the normalized swim history.
    log=wb['09_Import_Log']
    if not any(str(log.cell(i,1).value)=='SEIKO_JPN_2026' for i in range(5,log.max_row+1)):
        i=log.max_row+1
        for c,v in {1:'SEIKO_JPN_2026',2:'SEIKO Japan four-swimmer historical supplement',3:SOURCE.name,4:datetime(2026,10,17),5:datetime.now(),6:appended,7:'VERIFIED',8:'148 source-verified rows; PDF print dates or meet-date windows (all age-years unambiguous)',9:'https://swim.seiko.co.jp/'}.items():log.cell(i,c).value=v
    # refresh lookup rows in existing formulas for all imported/new swimmers; maintain private dates.
    last=persons.max_row
    for i in range(5,swims.max_row+1):
        cell=swims.cell(i,20)
        if cell.data_type=='f':cell.value=re.sub(r"\$E\$5:\$E\$\d+",f'$E$5:$E${last}',str(cell.value)).replace("$A$5:$A$52",f'$A$5:$A${last}')
    try:wb.calculation.fullCalcOnLoad=True;wb.calculation.forceFullCalc=True;wb.calculation.calcMode='auto'
    except:pass
    wb.save(BOOK)
    print(f'Japanese audited rows: {len(data)}; appended: {appended}; current swimmers: {persons.max_row-4}; total swims: {swims.max_row-4}')
    print('Australian DOB corrections:',', '.join(corrections[i]['Swimmer'] for i in range(len(corrections))))

if __name__=='__main__':refresh()
