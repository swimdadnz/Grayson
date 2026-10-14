"""ONE-TIME Australian historical boys' age-ranking supplement, not part of periodic scraping.

Usage: python scripts/import_historical_top10_once.py
Matches by canonical swimmer, event/course, official swim date and time. Safe to rerun.
Adds Thorpe's supplied SwimCloud URL as a source reference; it does NOT claim to scrape that profile.
"""
from __future__ import annotations
import csv
from collections import Counter
from datetime import datetime
from copy import copy
from pathlib import Path
from openpyxl import load_workbook

U=Path(__file__).resolve().parents[1]
F=U/'input'/'australian_historical_top10_verified.csv'
BOOK=U/'Grayson_Swim_Site_Master.xlsx'

def seconds(t):
 v=[float(x) for x in str(t).split(':')];return sum(x*60**i for i,x in enumerate(reversed(v)))

def run():
 wb=load_workbook(BOOK); swimmers=wb['01_Swimmers']; sources=wb['02_Swimmer_Sources']; swims=wb['04_Swims']
 people={str(swimmers.cell(r,2).value):(str(swimmers.cell(r,1).value),r) for r in range(5,swimmers.max_row+1) if swimmers.cell(r,2).value}
 if 'Ian Thorpe' not in people:
  row=swimmers.max_row+1
  for c in range(1,swimmers.max_column+1):swimmers.cell(row,c)._style=copy(swimmers.cell(row-1,c)._style)
  values=['SW019','Ian Thorpe','Ian Thorpe','M',datetime(1982,10,13),'Historical champion','AUS','','YES']
  for c,x in enumerate(values,1): swimmers.cell(row,c).value=x
  people['Ian Thorpe']=('SW019',row)
 if not any(str(sources.cell(i,4).value)=='939756' for i in range(5,sources.max_row+1)):
  row=sources.max_row+1
  for c in range(1,sources.max_column+1):sources.cell(row,c)._style=copy(sources.cell(row-1,c)._style)
  vals=['SW019','Ian Thorpe','SwimCloud','939756','https://www.swimcloud.com/swimmer/939756/times/','YES','Profile linked; full SwimCloud race history not acquired. Official Australian Top 10 subset appended separately.']
  for c,x in enumerate(vals,1):sources.cell(row,c).value=x
 eventmap={ (str(wb['03_Events'].cell(r,2).value), str(wb['03_Events'].cell(r,5).value)):str(wb['03_Events'].cell(r,1).value) for r in range(5,wb['03_Events'].max_row+1) if wb['03_Events'].cell(r,1).value }
 existing=set(); existing_abbreviated={}
 for r in range(5,swims.max_row+1):
  name=swims.cell(r,3).value; day=swims.cell(r,7).value; course=swims.cell(r,13).value; event=swims.cell(r,16).value; tm=swims.cell(r,18).value
  if name and day and tm is not None:
   key=(str(name),str(course),str(event),day.date().isoformat() if isinstance(day,datetime) else str(day)[:10],round(float(tm),2))
   existing.add(key)
 maxswim=max(int(str(swims.cell(i,1).value)[2:]) for i in range(5,swims.max_row+1) if str(swims.cell(i,1).value).startswith('SM'))
 counts=Counter(); source_log=[]
 with F.open(encoding='utf-8',newline='') as f:
  for record in csv.DictReader(f):
   name,course,event=record['Swimmer'],record['Course'],record['Event']; date=record['Swim Date']; tm=seconds(record['Time']); key=(name,course,event,date,round(tm,2))
   if name not in people:raise ValueError('No swimmer '+name)
   if (course,event) not in eventmap:raise ValueError('No event '+str((course,event)))
   birth=swimmers.cell(people[name][1],5).value.date(); exact=(datetime.fromisoformat(date).date()-birth).days/365.2425
   if int(exact)!=int(record['Age Group']): raise ValueError(f'Age mismatch: {key} {exact:.3f} vs {record["Age Group"]}')
   if key in existing:
    counts['matched_existing']+=1;continue
   row=swims.max_row+1;maxswim+=1
   for c in range(1,swims.max_column+1):swims.cell(row,c)._style=copy(swims.cell(5,c)._style)
   distance,stroke=event.split(' ')
   vals={1:f'SM{maxswim:06d}',2:people[name][0],3:name,4:'Swimming Australia Age All-Time Top 10',5:'',6:f'AUSTOP10|{name}|{course}|{date}|{event}|{record["Time"]}',7:datetime.fromisoformat(date),8:record['Meet'],9:'',10:'',11:'TIM',12:2,13:course,14:int(distance),15:stroke,16:event,17:eventmap[(course,event)],18:tm,19:record['Time'],20:exact,21:int(exact),22:'NO',23:'YES',24:'NO',28:record['Source URL'],29:'Verified published Australian all-time age Top-10; one-off historical supplement',30:'|'.join(map(str,key)),31:datetime.fromisoformat(date).timestamp(),36:'OK'}
   for c,v in vals.items():swims.cell(row,c).value=v
   existing.add(key);counts['appended']+=1
 # Full-range DOB lookups for any follow-up imported races; preserve original formulas.
 for r in range(5,swims.max_row+1):
  cell=swims.cell(r,20)
  if cell.data_type=='f' and "'01_Swimmers'!" in cell.value:cell.value=cell.value.replace('$E$22','$E$23').replace('$A$22','$A$23')
 wb['00_Control']['B5']='3.1 historical Top-10 reconciliation; Thorpe + QLD/AUS/world supplementary records'
 wb.save(BOOK)
 print('Australian historical import:',dict(counts),'; swimmers',len(people),'; rows',swims.max_row-4)
 return counts
if __name__=='__main__':run()
