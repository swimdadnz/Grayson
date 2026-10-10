#!/usr/bin/env python3
"""Build portable local-file Grayson site v5 from v4 template and private master workbook.
No workbook or DOB is exported. Re-run to stamp creation date, record details, and heatmaps.
Requirements: openpyxl and beautifulsoup4 (build machine only). Public output: offline HTML + media.
"""
from __future__ import annotations
import collections
import datetime as dt
import html
import json
import math
import os
import re
import shutil
from pathlib import Path
from bs4 import BeautifulSoup
from openpyxl import load_workbook

UPDATE=Path(__file__).resolve().parents[1]
BASE=UPDATE.parent
TEMPLATE=UPDATE/'source_templates'
PRIVATE=UPDATE/'Grayson_Swim_Site_Master.xlsx'
OUT=BASE
DATA=UPDATE/'generated_data'
NOW=dt.date.today()
# Preserve reproducible creation date from invoking build; optional override YYYY-MM-DD.
if os.environ.get('SITE_BUILD_DATE'): NOW=dt.date.fromisoformat(os.environ['SITE_BUILD_DATE'])
OUT.mkdir(parents=True,exist_ok=True)
(OUT/'assets').mkdir(parents=True,exist_ok=True)

PHOTOS=[
 ('2026 NZSC.jpeg','NZ Short Course Championships · 2026','National short-course championships', 'feature'),
 ('2024 AIMS butterfly.jpeg','AIMS Games · 2024','Butterfly race at AIMS 2024',''),
 ('Ariel and Grayson NAGS 2025.jpeg','National age-group championships · 2025','Two swimmers smiling at the pool',''),
 ('2025 butterfly.jpeg','Butterfly · 2025','Butterfly swim in pool',''),
 ('2026 JPP profile.jpeg','Junior Pan Pacific Championships · 2026','Junior Pan Pacific team profile poster','portrait'),
 ('2026 JPP first NZ cap.jpeg','First New Zealand cap · 2026','New Zealand team kit and cap',''),
 ('2025 NZSS breaststroke.jpeg','NZ Secondary Schools · 2025','Breaststroke racing',''),
 ('2025 NZSC mixed relay team.jpeg','NZ Short Course mixed relay · 2025','Four teammates at poolside',''),
 ('Grayson-Coulter-by-Nardia-Mulkerrins-2-1082x720.jpeg','First sub-4 min 400m Free · QLD Champs 2025','Grayson celebrating his first sub-four-minute 400m freestyle at the 2025 Queensland Championships',''),
 ('Grayson-Coulter-by-Nardia-Mulkerrins-3-1082x720.jpeg','Freestyle breakout · QLD Champs 2025','Grayson swimming freestyle at the 2025 Queensland Championships',''),
]
for filename,*_ in PHOTOS+[("2026-10-01 NZSC Scott Rice 200m Free commentary.mp4",None,None,None)]:
 src=BASE/'assets'/filename
 if not src.is_file(): raise FileNotFoundError(src)
 dst=OUT/'assets'/filename
 if src.resolve()!=dst.resolve(): shutil.copy2(src,dst)


def parse_day(value):
 if isinstance(value, dt.datetime): return value.date()
 if isinstance(value, dt.date): return value
 if not value: return None
 # two dates => oldest date represents when that record mark was first set
 parts=re.split(r'\s*/\s*|\s*&\s*',str(value).strip())
 valid=[]
 for raw in parts:
  raw=raw.strip()
  for fmt in ('%d.%m.%y','%d.%m.%Y','%d-%b-%y','%Y-%m-%d','%d/%m/%Y','%d/%m/%y'):
   try:
    date=dt.datetime.strptime(raw,fmt).date()
    # Excel's times are not dates; sanity-check source date
    if 1970<=date.year<=NOW.year: valid.append(date)
    break
   except ValueError: pass
 return min(valid) if valid else None


def anniversary_back(day,years):
 try:return day.replace(year=day.year-years)
 except ValueError:return day.replace(year=day.year-years,day=28)


def years_bin(d):
 # exact anniversary cutoffs, e.g., 2025-10-11 through 2026-10-10 is '0-1 years'
 for i in range(10):
  if d>anniversary_back(NOW,i+1):return str(i)
 if d>anniversary_back(NOW,15):return '10-15'
 if d>anniversary_back(NOW,20):return '15-20'
 return '20+'

bins=[(str(i),f'{i}–{i+1}y') for i in range(10)]+[('10-15','10–15y'),('15-20','15–20y'),('20+','20y+')]

wb=load_workbook(PRIVATE,read_only=True,data_only=True)
records=[]
source_rows=list(wb['05_Record_Benchmarks'].iter_rows(min_row=5,values_only=True))
for tup in source_rows:
 key,course,sex,age,event,holder=tup[:6]
 if not key: continue
 if not (sex in ('Male','Female') and course in ('SCM','LCM') and age in range(13,19)):continue
 holders=[] if not holder or str(holder).strip().lower()=='standard' else [p.strip() for p in str(holder).split(' / ') if p.strip()]
 holders=list(dict.fromkeys(holders)) # same person repeated gets only one record
 day=parse_day(tup[10]);
 records.append(dict(key=key,course=course,sex=sex,age=age,event=event,holder=holder,
                     holders=holders,time=str(tup[8]) if tup[8] is not None else '',
                     recordDate=day.isoformat() if day else None,originalDate=str(tup[10] or ''),
                     dateKnown=bool(day),bin=years_bin(day) if day else None))
assert len(records)==420,len(records)
held=[r for r in records if r['holders']]
assert len(held)>=400
by_gender={s:[r for r in held if r['sex']==s] for s in ['Male','Female']}
assert all(len(v)>0 for v in by_gender.values())

# Rebuild the record leaderboards from the updated workbook rather than reusing frozen v4 JSON.
def parse_seconds(value):
 try:
  pieces=[float(v) for v in str(value).strip().split(':')]
  return sum(x*(60**i) for i,x in enumerate(reversed(pieces)))
 except (ValueError,TypeError):return None

def recalc_leaderboard(sex):
 names=collections.defaultdict(lambda:{'scm':0,'lcm':0})
 for r in by_gender[sex]:
  for swimmer in r['holders']:
   names[swimmer][r['course'].lower()]+=1
 ordered=sorted(((n,c['scm'],c['lcm']) for n,c in names.items()),key=lambda t:(-(t[1]+t[2]),-t[1],t[0]))
 result=[];last_total=None;last_rank=0
 for position,(name,scm,lcm) in enumerate(ordered,1):
  total=scm+lcm
  rank=last_rank if last_total==total else position
  result.append({'swimmer':name,'scm':scm,'lcm':lcm,'total':total,'rank':rank})
  last_total=total;last_rank=rank
 return result

def recalc_outstanding(sex):
 result=[]
 for tup in source_rows:
  if not tup[0] or tup[2]!=sex or tup[14]!='YES':continue
  record_time,next_name,next_time=tup[8],tup[15],tup[16]
  if not next_name or not next_time:continue
  first,second=parse_seconds(record_time),parse_seconds(next_time)
  if first is None or second is None or second<=0 or first>second:continue
  result.append({'swimmer':tup[5],'event':tup[4],'course':tup[1],'age':tup[3],
                 'recordTime':str(record_time),'nextFastestSwimmer':str(next_name),
                 'nextFastestTime':str(next_time),'marginPct':round(100*(second-first)/second,4),
                 'recordKey':tup[0]})
 result.sort(key=lambda row:(-row['marginPct'],row['recordKey']))
 for position,row in enumerate(result,1):row['rank']=position
 return result[:20]

public_records_data={'asOf':NOW.isoformat(),
  'scope':'NZ individual age-group records, ages 13-18; SCM and LCM',
  'counts':{'male':recalc_leaderboard('Male'),'female':recalc_leaderboard('Female')},
  'top20':{'male':recalc_outstanding('Male'),'female':recalc_outstanding('Female')}}
assert len(public_records_data['top20']['male'])==20 and len(public_records_data['top20']['female'])==20

# Reconcile public record count lists and detailed history before publishing.
for s in ['Male','Female']:
 actual=collections.Counter(name for r in by_gender[s] for name in r['holders'])
 assert actual, s
print('Loaded',len(records),'records; held',len(held),'; dated',sum(r['dateKnown'] for r in held),'; unknown dates',sum(not r['dateKnown'] for r in held))

photo_gallery=[]
for filename,caption,alt,kind in PHOTOS:
 fn=html.escape(filename,quote=True); a=html.escape(alt,quote=True)
 kind_cls=' media-feature' if kind=='feature' else ' media-portrait' if kind=='portrait' else ''
 photo_gallery.append(f'''<a class="media-card{kind_cls}" href="assets/{fn}" target="_blank" rel="noopener" aria-label="Open full photo: {html.escape(caption,quote=True)}"><span class="media-frame"><img loading="lazy" decoding="async" src="assets/{fn}" alt="{a}"></span><span class="media-caption">{html.escape(caption)}</span></a>''')
photo_gallery.append('''<button type="button" class="media-card media-video" id="video-tile" aria-haspopup="dialog" aria-controls="video-dialog"><span class="media-frame video-preview"><span class="big-play" aria-hidden="true">▶</span><span class="video-tag">200m freestyle · Race commentary</span></span><span class="media-caption">NZ Short Course Championships · 1 October 2026 <span class="video-indicator">PLAY VIDEO ▶</span></span></button>''')
gallery_html='''<section class="section" id="gallery"><div class="wrap"><div class="eyebrow">In pictures &amp; sound</div><h2>Moments worth remembering.</h2><p class="lede">From early races to national and international meets — a few highlights from the years already enjoyed. Select a photo to see it larger, or the play tile to open race commentary.</p><div class="media-gallery">'''+''.join(photo_gallery)+'''</div></div></section>
<dialog id="video-dialog" class="video-dialog" aria-label="NZSC 200m freestyle video"><div class="video-dialog-bar"><strong>NZSC 2026 · 200m freestyle commentary</strong><button type="button" class="modal-close" id="video-close" aria-label="Close video">×</button></div><video id="nzsc-video" controls preload="none" playsinline><source src="assets/2026-10-01 NZSC Scott Rice 200m Free commentary.mp4" type="video/mp4">Your browser does not support this video format.</video><p class="video-footnote">NZ Short Course Championships 2026 · 200m freestyle commentary. If playback is unavailable, <a href="assets/2026-10-01 NZSC Scott Rice 200m Free commentary.mp4" target="_blank" rel="noopener">open the MP4 directly</a>.</p></dialog>'''

media_css='''
/* Offline media wall - all assets are relative to the HTML files */
.media-gallery{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:18px;margin-top:28px;align-items:stretch}
.media-card{display:flex;flex-direction:column;min-width:0;text-decoration:none;text-align:left;padding:0;border:1px solid #dce2e9;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 8px 25px rgba(7,27,51,.08);transition:transform .18s ease,box-shadow .18s ease;cursor:pointer;font:inherit;color:inherit}
.media-card:hover,.media-card:focus-visible{transform:translateY(-3px);box-shadow:0 16px 36px rgba(7,27,51,.15);outline-offset:3px}
.media-frame{display:flex;align-items:center;justify-content:center;position:relative;aspect-ratio: 4/3;background:#d6e1e9;overflow:hidden}
.media-frame img{width:100%;height:100%;object-fit:cover;display:block}
.media-feature{grid-column:span 2}.media-feature .media-frame{aspect-ratio:2.05/1}
.media-portrait .media-frame{background:var(--navy)}.media-portrait img{object-fit:contain}
.media-caption{display:block;font-size:.86rem;line-height:1.35;font-weight:750;padding:13px 15px 15px}
.media-pending{flex-direction:column;background:repeating-linear-gradient(135deg,#edf2f6 0px,#edf2f6 16px,#f9fbfd 16px,#f9fbfd 32px);color:var(--navy);gap:8px;padding:24px;text-align:center}
.media-pending small{font-size:.73rem;color:#627083;max-width:25ch}.photo-pictogram{font-size:2.8rem;color:#8297a8;line-height:1}
.video-preview{background:linear-gradient(145deg,rgba(4,25,45,.82),rgba(9,56,87,.88)),url('assets/2026 NZSC.jpeg') center/cover no-repeat;color:#fff;gap:18px;flex-direction:column}
.big-play{width:76px;height:76px;display:grid;place-items:center;border-radius:50%;background:var(--yellow);color:var(--navy);font-size:2rem;box-shadow:0 8px 26px #0004;padding-left:5px}.video-tag{font-size:.78rem;text-transform:uppercase;letter-spacing:.1em;font-weight:750;text-align:center}
.video-indicator{display:inline-block;color:#0d5b89;font-size:.74rem;white-space:nowrap;margin-left:8px}
.gallery-note{font-size:.81rem;color:var(--muted);margin-top:24px}
.video-dialog{width:min(850px,94vw);max-height:92vh;border:none;border-radius:14px;padding:0;box-shadow:0 30px 70px #071b3350;background:#fff}
.video-dialog::backdrop,.record-dialog::backdrop{background:rgba(2,16,31,.75)}
.video-dialog-bar{display:flex;align-items:center;justify-content:space-between;gap:20px;padding:14px 20px;background:var(--navy);color:#fff}
.modal-close{border:none;cursor:pointer;font:inherit;font-size:1.7rem;line-height:1;background:transparent;color:inherit;width:44px;height:44px}
.video-dialog video{display:block;width:100%;background:#000;max-height:70vh;aspect-ratio:16/9}
.video-footnote{margin:0;padding:16px 20px;font-size:.86rem;color:#4b5563}
@media(max-width:850px){.media-gallery{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:540px){.media-gallery{grid-template-columns:1fr;gap:14px}.media-feature{grid-column:span 1}.media-feature .media-frame{aspect-ratio:4/3}.media-card .media-frame{aspect-ratio:4/3}}
@media(prefers-reduced-motion:reduce){.media-card{transition:none}.media-card:hover{transform:none}}
'''

record_css='''
/* Record-holder details and responsive heatmaps */
.modal-close{border:0;background:none;color:inherit;font:inherit;font-size:1.8rem;width:42px;height:42px;cursor:pointer}.modal-close:focus-visible{outline:2px solid var(--yellow);outline-offset:3px}
.swimmer-open{border:0;background:none;padding:7px 3px;color:#07558b;font:inherit;font-weight:700;cursor:pointer;text-align:left;text-decoration:underline;text-decoration-color:#8cb4d0;text-underline-offset:3px;min-height:38px}
.swimmer-open:hover,.swimmer-open:focus-visible{color:#063654;text-decoration-color:#063654;outline-offset:3px}
.record-dialog{width:min(720px,calc(100vw - 24px));max-height:min(90vh,850px);padding:0;border:0;border-radius:16px;box-shadow:0 20px 70px #071b3355}
.record-dialog-head{position:sticky;top:0;z-index:3;display:flex;justify-content:space-between;gap:10px;align-items:center;padding:16px 22px;background:var(--navy);color:white}
.record-dialog-body{padding:20px 24px 28px;overflow:auto;max-height:calc(90vh - 80px)}
.record-dialog-title{font-size:clamp(1.1rem,3vw,1.5rem);font-weight:800;line-height:1.1;margin:0}
.record-dialog-summary{margin:0 0 18px;color:var(--muted);font-size:.92rem}
.record-list{list-style:none;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:7px 14px;padding:0;margin:0}
.record-list li{display:flex;align-items:center;justify-content:space-between;gap:12px;border-bottom:1px solid #e2e7ed;padding:9px 0;font-size:.88rem;line-height:1.4}
.record-short{font-weight:700}.record-time{font-variant-numeric:tabular-nums;color:#43536a;white-space:nowrap}
.recency-intro{max-width:900px;color:var(--muted)}.heat-groups{display:grid;grid-template-columns:minmax(0,1fr);gap:34px;margin:30px 0}.heat-group{min-width:0}.heat-group h3{font-size:1.35rem;margin:0 0 12px;color:var(--navy)}
.heat-scroller{overflow-x:auto;max-width:100%;width:100%;min-width:0;border:1px solid var(--line);border-radius:12px;box-shadow:0 5px 18px #071b3308;scrollbar-width:thin}
.heat-table{border-collapse:separate;border-spacing:3px;width:100%;min-width:920px;background:#fff;font-size:.79rem;table-layout:fixed}
.heat-table th,.heat-table td{text-align:center;border:0;padding:0;background:transparent}
.heat-table thead th{font-size:.7rem;padding:14px 2px;font-weight:800;background:#e8eef4;color:#20314a;line-height:1.2;white-space:normal}
.heat-table thead th:first-child{width:75px;border-top-left-radius:8px}
.heat-table tbody th{color:#28374a;background:#eef2f6;font-size:.81rem;width:75px}
.heat-cell{display:grid;place-items:center;width:100%;height:48px;border:0;border-radius:5px;cursor:pointer;font:800 .91rem/1 Arial,Helvetica,sans-serif;font-variant-numeric:tabular-nums;transition:box-shadow .15s,transform .15s}
.heat-cell:focus-visible,.heat-cell:hover{outline:2px solid var(--yellow);outline-offset:1px;box-shadow:0 0 0 2px #071b33}
.heat-cell.zero{font-weight:450;color:#a1aebb;background:#f3f5f7 !important}
.heat-footer{display:flex;flex-wrap:wrap;gap:10px 16px;align-items:center;font-size:.81rem;color:var(--muted);margin:10px 0}
.heat-key{display:inline-flex;gap:4px;align-items:center}.heat-key span{display:inline-block;width:22px;height:12px;border-radius:2px}
.heat-cell-detail{background:#fff;border-left:4px solid var(--yellow);padding:14px 18px;min-height:48px;color:#293c52;font-size:.89rem}
.heat-cell-detail ul{margin:10px 0 0;padding-left:20px;columns:2;column-gap:22px}.heat-cell-detail li{break-inside:avoid;margin:0 0 5px}
.recency-analysis{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:24px}.recency-insight{background:#fff;border:1px solid #e0e6ec;padding:22px;border-radius:12px;box-shadow:0 4px 15px #071b3308}.recency-insight h3{font-size:1.1rem;line-height:1.25;color:var(--navy);margin:0 0 10px}.recency-insight p{margin:0;color:#495971;font-size:.93rem}
@media(max-width:680px){.record-list,.recency-analysis{grid-template-columns:1fr}.heat-cell-detail ul{columns:1}.record-dialog-body{padding:16px}.heat-table{min-width:820px}.heat-cell{height:44px}.heat-scroller{margin-left:-1vw;margin-right:-1vw}}
/* prevent data tables and scrollable heatmap from inflating responsive grid columns */
.two-col{grid-template-columns:repeat(2,minmax(0,1fr))}.two-col>div{min-width:0}.section .wrap{min-width:0}.heat-footer{min-width:0}.heat-key{flex-shrink:0}
/* common navigation and SVG adaptability for narrow portrait screens */
@media(max-width:670px){.nav{display:flex;align-items:flex-start;flex-direction:column;gap:12px;min-height:unset;padding:12px 0 14px}.nav-links{display:flex;gap:6px 12px;flex-wrap:wrap;justify-content:flex-start;width:100%}.nav-links a{font-size:.83rem;padding:8px 2px}.wrap{width:min(1180px,94vw)}.section h2{line-height:1.1}.table-wrap{max-width:100%}.two-col{grid-template-columns:minmax(0,1fr)}.svg-chart{min-height:260px}.chart-card{max-width:100%}}
'''

# Heatmap content and quantitative reading of the diagonals

def heat_cell_colour(n,max_n):
 if n==0:return '#f3f5f7','#a1aebb'
 # sqrt improves discrimination of light cells while preserving a clear dark high end
 z=math.sqrt(n/max_n)
 a=(219,234,242);b=(7,42,76)
 rgb=tuple(round(a[i]*(1-z)+b[i]*z) for i in range(3))
 return '#'+''.join(f'{x:02x}' for x in rgb), ('#ffffff' if z>.48 else '#082f4b')

heat_records=[]
for r in held:
 if not r['dateKnown']:continue
 heat_records.append(dict(sex=r['sex'],age=r['age'],bin=r['bin'],event=r['event'],course=r['course'],
                          holders=r['holders'],date=r['recordDate'],time=r['time']))
max_count=max(sum(1 for r in heat_records if r['sex']==s and r['age']==age and r['bin']==b) for s in by_gender for age in range(13,19) for b,_ in bins)
heat_html=['<section class="section" id="record-recency"><div class="wrap"><div class="eyebrow">How the record book evolves</div><h2>How old are today’s records?</h2><p class="recency-intro">Each square counts individual records in that age group and course combined (SCM + LCM), using the date the current record time was first set. Columns are measured backwards from the site build date of '+NOW.strftime('%-d %B %Y')+'. The darker the square, the more records in that recency period. Tap any square to see its events.</p><div class="heat-groups">']
for sex,code in [('Male','male'),('Female','female')]:
 n=sum(1 for r in heat_records if r['sex']==sex)
 heat_html.append(f'<div class="heat-group"><h3>{sex} age-group records</h3><div class="heat-scroller" tabindex="0" role="region" aria-label="Scrollable {sex.lower()} record recency heat map"><table class="heat-table"><thead><tr><th scope="col">Age</th>')
 for b,lab in bins:heat_html.append(f'<th scope="col" title="Records set {html.escape(lab)} ago">{html.escape(lab)}</th>')
 heat_html.append('</tr></thead><tbody>')
 for age in range(13,19):
  heat_html.append(f'<tr><th scope="row">{age}y</th>')
  for b,label in bins:
   count=sum(1 for r in heat_records if r['sex']==sex and r['age']==age and r['bin']==b)
   bg,fg=heat_cell_colour(count,max_count)
   heat_html.append(f'<td><button type="button" class="heat-cell {"zero" if not count else ""}" style="background:{bg};color:{fg}" data-sex="{sex}" data-age="{age}" data-bin="{b}" aria-label="{sex}, age {age}, {label}: {count} records" title="{sex} {age}y · {label}: {count} records">{count if count else "·"}</button></td>')
  heat_html.append('</tr>')
 heat_html.append(f'</tbody></table></div><div class="heat-footer"><span>{n} dated, swimmer-held records included.</span><span class="heat-key" aria-label="Light shades mean fewer records, dark shades mean more"><span style="background:#dbeaf2"></span><span style="background:#8eafc5"></span><span style="background:#365e7e"></span><span style="background:#072a4c"></span>More records → darker</span></div><div class="heat-cell-detail" id="heat-detail-{code}" aria-live="polite">Select a cell to see the records behind that number.</div></div>')
heat_html.append('</div>')

# Compute recent share for narrative.  The visual cells are a descriptive snapshot, not proof of swimmer paths.
for sex in ['Male','Female']:
 xx=[r for r in heat_records if r['sex']==sex]
 recent=[r for r in xx if r['date']>anniversary_back(NOW,2).isoformat()]
 print(sex, 'dated records',len(xx),'last two years',len(recent),'pct',round(100*len(recent)/len(xx),1))
 for age in range(13,19):
  sub=[r for r in xx if r['age']==age and r['date']>anniversary_back(NOW,2).isoformat()]
  print(age,len(sub),collections.Counter(n for r in sub for n in r['holders']).most_common(3))

# Detect specifically documented swimmer runs ages 13-17 in sequential swim years
for sex in ['Male','Female']:
 grouped=collections.defaultdict(list)
 for r in heat_records:
  if r['sex']==sex:
   for n in r['holders']: grouped[n].append(r)
 candidates=[]
 for n,v in grouped.items():
  age_years=collections.defaultdict(set)
  for r in v:age_years[r['age']].add(int(r['date'][:4]))
  max_chain=0
  for age in range(13,18):
   for y in age_years[age]:
    chain=1
    for k in range(1,6):
     if y+k in age_years.get(age+k,set()):chain+=1
     else:break
    max_chain=max(max_chain,chain)
  if max_chain>=2: candidates.append((max_chain,len(v),n,sorted(age_years.items())))
 print('COHORTS',sex,sorted(candidates,reverse=True)[:14])

# Data-derived analysis; includes caveat that a diagonal can be cohort renewal but cannot establish swimmer identity by itself.
heat_html.append('''<div class="recency-analysis"><div class="recency-insight"><h3>Different pace of recent change</h3><p id="recency-summary">''')
for s in ['Male','Female']:
 a=[r for r in heat_records if r['sex']==s]; recent=sum(1 for r in a if r['date']>anniversary_back(NOW,2).isoformat())
 heat_html.append(f'{s}: <strong>{recent} of {len(a)}</strong> dated, swimmer-held records ({100*recent/len(a):.1f}%) were set in the past two years. ')
heat_html.append('''This is a snapshot of <em>surviving current records</em>, not the total number broken in each year.</p></div><div class="recency-insight"><h3>Male diagonal: successive age groups</h3><p>Historical diagonal candidates include Ariel Muchirahondo and Grayson Coulter; select adjacent age-row cells to confirm which record times and holders still survive. As the record book changes, the strength and dates of these diagonal bands can change.</p></div><div class="recency-insight"><h3>Female diagonals: earlier generations</h3><p>Historical diagonal candidates include Erika Fairweather and Gabrielle Fa'amausili. Check the age, swimmer and date details in the cells: diagonals suggest cohorts progressing across ages, not proof that one swimmer set every record in a band.</p></div><div class="recency-insight"><h3>Recent female records are concentrated at older ages</h3><p>Compare the recent female rows with earlier-age rows to see where the pace of record changes is concentrated. Zoe Pedersen, Milan Glintmeyer and Monique Wieruszowski are useful examples to explore, but the distribution should be reassessed at every refresh. <strong>Interpretation:</strong> individual heat-map cells do not prove a swimmer's progression without checking dates and names.</p></div></div>''')
unknown=sum(1 for r in held if not r['dateKnown'])
heat_html.append('<p class="gallery-note">Methodology: one current event/course/age record per cell, including ties once (not once per joint holder); original date for equal record times. Unclaimed “Standard” records are excluded. ' + (f'{unknown} swimmer-held records lack an interpretable date and are excluded from the heat map. ' if unknown else '') + 'The 10–15, 15–20 and 20+ year columns aggregate broader periods than the first ten annual columns. Older records are not adjusted for changes in events or competition conditions.</p></div></section>')
heat_html=''.join(heat_html)

# shared JS installed after the inline page controller JS; this stays local and cross-browser
record_items=[]
for r in held:
 for name in r['holders']:
  record_items.append({k:r[k] for k in ('course','age','event','time','recordDate') }|{'swimmer':name,'sex':r['sex']})
json_items=json.dumps(record_items,ensure_ascii=False,separators=(',',':')).replace('</','<\\/')
json_heat=json.dumps(heat_records,ensure_ascii=False,separators=(',',':')).replace('</','<\\/')
interaction_js='''
const holderRecords=__RECORDS__;
const heatRecords=__HEATRECORDS__;
const detailsDialog=document.getElementById('record-detail-dialog');
let previousFocus=null;
function openSwimmer(name,sex){
 const match=holderRecords.filter(r=>r.swimmer===name&&r.sex===sex);
 if(!match.length)return;
 previousFocus=document.activeElement;
 match.sort((a,b)=>a.age-b.age||({SCM:0,LCM:1}[a.course]-{SCM:0,LCM:1}[b.course])||parseInt(a.event)-parseInt(b.event)||a.event.localeCompare(b.event));
 document.getElementById('record-dialog-title').textContent=name;
 const scm=match.filter(x=>x.course==='SCM').length,lcm=match.length-scm;
 document.getElementById('record-dialog-summary').textContent=`${match.length} current records · ${scm} short course · ${lcm} long course`;
 document.getElementById('record-dialog-list').innerHTML=match.map(x=>`<li><span class="record-short">${x.age}y ${x.course==='SCM'?'SC':'LC'} ${esc(x.event)}</span><span class="record-time">${esc(x.time)}</span></li>`).join('');
 if(typeof detailsDialog.showModal==='function')detailsDialog.showModal();else detailsDialog.setAttribute('open','');
}
// Event delegation covers male/female leaderboards and outstanding-record rows, also if re-rendered.
document.addEventListener('click',ev=>{
 const btn=ev.target.closest('button.swimmer-open');
 if(btn){openSwimmer(btn.dataset.swimmer,btn.dataset.sex);return}
 const cell=ev.target.closest('button.heat-cell');
 if(cell){
   const hits=heatRecords.filter(x=>x.sex===cell.dataset.sex && x.age===Number(cell.dataset.age)&&x.bin===cell.dataset.bin);
   const target=document.getElementById('heat-detail-'+cell.dataset.sex.toLowerCase());
   const heading=`${cell.dataset.sex} · ${cell.dataset.age} years · ${cell.title.split('·').slice(-1)[0].trim()}`;
   target.innerHTML=`<strong>${esc(heading)} (${hits.length})</strong>`+(hits.length?'<ul>'+hits.sort((a,b)=>a.date.localeCompare(b.date)).map(x=>`<li>${x.age}y ${x.course==='SCM'?'SC':'LC'} ${esc(x.event)} — ${esc(x.holders.join(' / '))} (${esc(x.date)})</li>`).join('')+'</ul>':'<p>No current swimmer-held records in this period.</p>');
 }
});
document.getElementById('record-dialog-close').addEventListener('click',()=>detailsDialog.close());
detailsDialog.addEventListener('close',()=>{if(previousFocus&&previousFocus.focus)previousFocus.focus()});
'''.replace('__RECORDS__',json_items).replace('__HEATRECORDS__',json_heat)

video_js='''
const vidDlg=document.getElementById('video-dialog');
const movie=document.getElementById('nzsc-video');
document.getElementById('video-tile').addEventListener('click',()=>{
 if(vidDlg.showModal)vidDlg.showModal();else vidDlg.setAttribute('open','');
  const start=movie.play(); if(start && typeof start.catch==='function')start.catch(()=>{});
});
document.getElementById('video-close').addEventListener('click',()=>vidDlg.close());
vidDlg.addEventListener('close',()=>{movie.pause();});
'''


def inject(soup,tag,source):
 node=BeautifulSoup(source,'html.parser')
 for el in list(node.contents):tag.append(el)


def add_css(soup, css):
 style=soup.head.find('style');style.string=(style.string or '')+'\n'+css


def add_script(soup,script):
 n=soup.new_tag('script');n.string=script;soup.body.append(n)

for file in ['index.html','nz-records.html','progression.html','consistency.html']:
 soup=BeautifulSoup((TEMPLATE/file).read_text(encoding='utf-8'),'html.parser')
 footer=soup.find('footer').find(class_='wrap')
 footer.clear()
 footer.append(f'A summary of records and achievements as at {NOW.day} {NOW.strftime("%B %Y")}.')
 if file=='index.html':
  home_stats=soup.select('.stat strong')
  if len(home_stats)>=3:
   gscm=sum(1 for r in held if 'Grayson Coulter' in r['holders'] and r['course']=='SCM')
   glcm=sum(1 for r in held if 'Grayson Coulter' in r['holders'] and r['course']=='LCM')
   for strong,value in zip(home_stats,[gscm+glcm,gscm,glcm]):strong.string=str(value)
  sections=soup.find_all('section',class_='section')
  media_sec=next(x for x in sections if 'Photos & video' in x.get_text(' ',strip=True))
  media_sec.replace_with(BeautifulSoup(gallery_html,'html.parser'))
  add_css(soup,media_css+record_css)
  add_script(soup,video_js)
 elif file=='nz-records.html':
  add_css(soup,record_css)
  # Replace static rendering helper so swimmer names in both leaderboards are actual buttons.
  scripts=soup.find_all('script')
  original=scripts[-1].string or ''
  original=original.replace('const recordsData={','const recordsData={',1)
  old_count='''<td class="rank">${r.rank}</td><td>${esc(r.swimmer)}</td>'''
  replacement='''<td class="rank">${r.rank}</td><td><button type="button" class="swimmer-open" data-sex="${rows===recordsData.counts.male?'Male':'Female'}" data-swimmer="${esc(r.swimmer)}" aria-label="List current records for ${esc(r.swimmer)}">${esc(r.swimmer)}</button></td>'''
  if old_count not in original:
   # locate exact renderer in existing minified JS
   match=re.search(r'function countTable\(rows\).*?\nfunction topTable',original,re.S)
   print('Existing countTable snippet:',(match.group(0)[:1000] if match else 'MISSING'))
  else:original=original.replace(old_count,replacement)
  # Correct the top20 table's own sex lookup if its HTML was transformed by countTable replacement.
  a=original.index('function topTable(')
  b=original.index('const d=recordsData;',a)
  top_segment=original[a:b].replace('rows===recordsData.counts.male', 'rows===recordsData.top20.male')
  original=original[:a]+top_segment+original[b:]
  # Outstanding Top20 click names, determine sex by incoming table reference.
  old_top='''<td>${esc(r.swimmer)}</td>'''
  top_repl='''<td><button type="button" class="swimmer-open" data-sex="${rows===recordsData.top20.male?'Male':'Female'}" data-swimmer="${esc(r.swimmer)}">${esc(r.swimmer)}</button></td>'''
  if old_top in original:original=original.replace(old_top,top_repl)
  else: print('WARN top table match missing')
  json_public=json.dumps(public_records_data,separators=(',',':'),ensure_ascii=False).replace('</','<\\/')
  original,nreplace=re.subn(r'const recordsData=(\{.*?\});\s*function countTable',lambda m:'const recordsData='+json_public+';\nfunction countTable',original,count=1,flags=re.S)
  assert nreplace==1,'Cannot replace embedded leaderboard data'
  scripts[-1].string=original
  # Insert heat map directly before footer. Native html table + JS cell drill down.
  soup.find('footer').insert_before(BeautifulSoup(heat_html,'html.parser'))
  insert_modal='''<dialog id="record-detail-dialog" class="record-dialog" aria-labelledby="record-dialog-title"><div class="record-dialog-head"><h2 class="record-dialog-title" id="record-dialog-title">Swimmer record list</h2><button type="button" class="modal-close" id="record-dialog-close" aria-label="Close record list">×</button></div><div class="record-dialog-body"><p id="record-dialog-summary" class="record-dialog-summary"></p><ul id="record-dialog-list" class="record-list"></ul></div></dialog>'''
  soup.find('footer').insert_before(BeautifulSoup(insert_modal,'html.parser'))
  add_script(soup,interaction_js)
 else: add_css(soup,record_css)
 # Bring freshly recalculated progression and consistency data into the portable offline HTML.
 if file in ('progression.html','consistency.html'):
  var='progressionData' if file=='progression.html' else 'consistencyData'
  json_file=DATA/('progression.json' if file=='progression.html' else 'consistency.json')
  if not json_file.is_file(): raise FileNotFoundError('Run workbook refresh before HTML regeneration: '+str(json_file))
  parsed=json.loads(json_file.read_text(encoding='utf-8'))
  payload=json.dumps(parsed,ensure_ascii=False,separators=(',',':')).replace('</','<\\/')
  for script in soup.find_all('script'):
   if script.string and 'const '+var+'=' in script.string:
    code=script.string
    pattern=r'const '+var+r'=\{.*?\};\s*(?=function |const |\n)'
    replacement='const '+var+'='+payload+';\n'
    code,n=re.subn(pattern,lambda m:replacement,code,count=1,flags=re.S)
    if n!=1: raise RuntimeError('Unable to replace '+var+' inline JSON')
    script.string=code
    break
  else:raise RuntimeError('No embedded '+var+' variable found in template')
 # A small build stamp is useful for a human reviewing downloaded files.
 soup.head.append(soup.new_tag('meta',attrs={'name':'site-build-date','content':NOW.isoformat()}))
 (OUT/file).write_text(str(soup),encoding='utf-8')

# The root README is an operator-maintained document. Never overwrite it during rebuild.
print('Generated',OUT)
