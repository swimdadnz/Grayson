"""One-off, idempotent reconciliation for 6 Oct 2026 SNZ SCM book and 2024 QLD swims.

Run ONLY for the supplied 2026 historical migration, not on each future refresh.
Checks 216 SCM record times/holders against a transcription verified from the
ratified book; makes no changes to actual record time or holders. For the 2024
Queensland swims, appends only rows absent from SwimCloud/source data.
"""
from __future__ import annotations
from pathlib import Path
from openpyxl import load_workbook
from copy import copy
from collections import Counter
from datetime import datetime,date
import csv, re
from name_aliases import canonical_holders

ROOT=Path(__file__).resolve().parents[1]
BOOK=ROOT/'Grayson_Swim_Site_Master.xlsx'
RATIFIED=ROOT/'input'/'ratified_scm_age_records_2026-10-06.csv'
QLD=ROOT/'input'/'queensland_age12_historical_2024.csv'
NZ_URL='https://www.swimmingnz.org/performance-and-coaching'


def secs(v):
    parts=str(v).strip().split(':')
    return round(sum(float(p)*60**i for i,p in enumerate(reversed(parts))),2)


def run():
    wb=load_workbook(BOOK)
    rb=wb['05_Record_Benchmarks']; hist=wb['06_Record_History'];sw=wb['04_Swims'];ev=wb['03_Events']
    bykey={str(rb.cell(r,1).value):r for r in range(5,rb.max_row+1) if rb.cell(r,1).value}
    # All verifications before changing workbook, including missing input rows.
    with RATIFIED.open(encoding='utf-8-sig',newline='') as f:ratified=list(csv.DictReader(f))
    assert len(ratified)==216 and len({r['Record Key'] for r in ratified})==216,'Expected the complete ratified SCM age record book'
    for d in ratified:
        k=d['Record Key'];assert k in bykey and '|SCM|' in k,k
        row=bykey[k]
        if secs(rb.cell(row,9).value)!=secs(d['Ratified Time']):raise ValueError('Ratified time mismatch at '+k)
        canonical=set(canonical_holders(d['Holder from ratified list (canonical)']))
        old=set(canonical_holders(rb.cell(row,6).value))
        if canonical!=old:raise ValueError(f'Ratified holder mismatch at {k}: {canonical} != {old}')
        if str(rb.cell(row,11).value or '').strip()!=d['Ratified Date'].strip():raise ValueError('Ratified record date mismatch at '+k)
    with QLD.open(encoding='utf-8-sig',newline='') as f: qld=list(csv.DictReader(f))
    assert len(qld)==5 and len({(x['Event'],x['Swim Date']) for x in qld})==5
    eventmap={(ev.cell(r,2).value,ev.cell(r,5).value):ev.cell(r,1).value for r in range(5,ev.max_row+1) if ev.cell(r,1).value}
    people=wb['01_Swimmers']; swimmer_id=next(people.cell(i,1).value for i in range(5,people.max_row+1) if people.cell(i,2).value=='Grayson Coulter')
    existing=set()
    for r in range(5,sw.max_row+1):
        name=sw.cell(r,3).value; dd=sw.cell(r,7).value; t=sw.cell(r,18).value
        if name and dd and t is not None:
            day=dd.date().isoformat() if isinstance(dd,datetime) else str(dd)[:10]
            existing.add((str(name),str(sw.cell(r,13).value),str(sw.cell(r,16).value),day,round(float(t),2)))
    needed=[]
    for d in qld:
        if d['Swimmer']!='Grayson Coulter' or d['Course']!='LCM' or d['Age']!='12':raise ValueError('Unexpected historical source row')
        if (d['Course'],d['Event']) not in eventmap:raise ValueError('Unknown event '+d['Event'])
        key=(d['Swimmer'],d['Course'],d['Event'],d['Swim Date'],secs(d['Time']))
        if key not in existing:needed.append((d,key))
    # Verified against official book: attribution change only, no record changes.
    for d in ratified:
        i=bykey[d['Record Key']]
        rb.cell(i,12).value='Swimming NZ, ratified Short Course Records by Age as at 6 Oct 2026'
        rb.cell(i,13).value=NZ_URL
        note=str(rb.cell(i,24).value or '')
        if '6 Oct 2026' not in note:rb.cell(i,24).value=(note+' | ' if note else '')+'Ratified against Swimming NZ SCM record list as at 6 Oct 2026.'
    active_scm={r['Record Key'] for r in ratified}
    for i in range(5,hist.max_row+1):
        if str(hist.cell(i,2).value) in active_scm and hist.cell(i,11).value=='YES':hist.cell(i,12).value=NZ_URL
    ctrl=wb['00_Control'];ctrl['A31']='NZSC 2026 ratified record source';ctrl['B31']=NZ_URL
    ctrl['C31']='Official Swimming NZ short-course book dated 6 Oct 2026; historical non-record swims retain original acquisition provenance.'
    # Attribute any direct recorded age-14 NZ records in supplemental swim history to ratified source.
    record_swims={(r['Record Key'].split('|')[4],r['Ratified Date'],secs(r['Ratified Time'])) for r in ratified if r['Holder from ratified list (canonical)']=='Grayson Coulter'}
    ratified_swims=0
    for i in range(5,sw.max_row+1):
        if sw.cell(i,3).value!='Grayson Coulter' or sw.cell(i,13).value!='SCM':continue
        day=sw.cell(i,7).value
        if not isinstance(day,datetime):continue
        stamp=day.strftime('%d.%m.%y')
        t=sw.cell(i,18).value
        if t is not None and (str(sw.cell(i,16).value),stamp,round(float(t),2)) in record_swims:
            # Defend against future changes to age approximation: check actual ratified keys by matching date/event/time.
            sw.cell(i,4).value='Swimming NZ ratified age record';sw.cell(i,28).value=NZ_URL
            sw.cell(i,29).value='Ratified Swimming NZ SCM records as at 6 Oct 2026';ratified_swims+=1
    maxnum=max(int(str(sw.cell(i,1).value)[2:]) for i in range(5,sw.max_row+1) if str(sw.cell(i,1).value).startswith('SM'))
    for d,key in needed:
        row=sw.max_row+1;maxnum+=1
        for c in range(1,sw.max_column+1):sw.cell(row,c)._style=copy(sw.cell(5,c)._style)
        dist,stroke=d['Event'].split()
        tm=secs(d['Time']);day=datetime.fromisoformat(d['Swim Date']);days=(day.date()-date(2012,1,17)).days
        num=f'SM{maxnum:06d}'
        v={1:num,2:swimmer_id,3:'Grayson Coulter',4:'Swimming Queensland All-Comers records',5:'',6:'QLD2024|'+d['Swim Date']+'|'+d['Event'],7:day,8:d['Meet'],9:'',10:'',11:'TIM',12:2,13:'LCM',14:int(dist),15:stroke,16:d['Event'],17:eventmap[('LCM',d['Event'])],18:tm,19:d['Time'],20:days/365.2425,21:12,22:'NO',23:'YES',24:'NO',28:d['Source URL'],29:d['Notes'],30:'|'.join(map(str,key)),31:None,36:'OK'}
        for c,x in v.items():sw.cell(row,c).value=x
        # Age-year/source data are regenerated in update_master.py.
    log=wb['09_Import_Log']
    if not any(str(log.cell(i,1).value)=='RATSC20261006' for i in range(5,log.max_row+1)):
        i=log.max_row+1
        vals=['RATSC20261006','Swimming NZ SCM ratified age records','Swimming_NZ_Short_Course_Records_as_at_2026-10-06.pdf','2026-10-06',datetime.now(),216,'Verified','All 216 SCM age record times/holders match database; official status now ratified','Private reference PDF']
        for c,x in enumerate(vals,1):log.cell(i,c).value=x
    if not any(str(log.cell(i,1).value)=='QLD2024AGE12' for i in range(5,log.max_row+1)):
        i=log.max_row+1;vals=['QLD2024AGE12','Swimming Queensland Age 12 historical supplement','queensland_age12_historical_2024.csv','2024-12-20',datetime.now(),len(needed),'Verified','Five official All-Comers record performances; one-time import; deduplicated vs SwimCloud','Private master only']
        for c,x in enumerate(vals,1):log.cell(i,c).value=x
    ctrl['B5']='3.2 | ratified 6 Oct 2026 SCM; one-off 2024 Queensland age-12 supplement'
    try:wb.calculation.fullCalcOnLoad=True;wb.calculation.forceFullCalc=True
    except Exception:pass
    wb.save(BOOK)
    print(f'Ratified 216 SCM age-group record combinations; appended {len(needed)} missing Queensland swims; {ratified_swims} existing recorded swims attributed to ratified source.')
if __name__=='__main__':run()
