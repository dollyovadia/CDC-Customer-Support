#!/usr/bin/env python3
"""Add 'Triage status' (+ overlap with a previous replies file) to a Replies
workbook and prepend a 'Summary for support' sheet.
Usage: python3 triage_summary.py REPLIES.xlsx [PREVIOUS_REPLIES.xlsx] [--title "..."]"""
import sys, collections, openpyxl
from openpyxl.styles import Font, Alignment
args=[a for a in sys.argv[1:] if not a.startswith('--')]; title=next((sys.argv[i+1] for i,a in enumerate(sys.argv) if a=='--title'), 'CDC check - login tickets')
inp=args[0]; prev_path=args[1] if len(args)>1 else None
wb=openpyxl.load_workbook(inp); ws=wb['Replies']; rr=list(ws.iter_rows(values_only=True)); h=rr[0]; rows=[dict(zip(h,r)) for r in rr[1:]]
if 'Triage status' in h: sys.exit('already triaged')
pidx={}
if prev_path:
    prev=openpyxl.load_workbook(prev_path, read_only=True)['Replies']; pr=list(prev.iter_rows(values_only=True)); ph=pr[0]
    for r in pr[1:]:
        d=dict(zip(ph,r))
        for k in (d['Customer gave: Phone'], d['Customer gave: Email']):
            if k: pidx[str(k).lower()]=d['User ID']
def status(d):
    res=d['CDC result'] or ''
    if d['Email check'] and res.startswith('Mismatch'): return 'D - Tech team: erase/release account (reply after)'
    if d['Logged in since cutoff']=='Yes': return 'A - Already logged in since 7.10: send confirmation, low priority'
    if res in ('Phone not in CDC (likely erased)','Phone and email not in CDC (likely erased)','Email not in CDC (likely erased)'): return 'C - No account in CDC: ask to register again'
    if res=='No phone or email to check': return 'E - Ask for phone + email'
    return 'B - Not logged in yet: send instructions / call back'
st=[status(d) for d in rows]
prev_hit=[pidx.get(str(d['Customer gave: Phone'] or '').lower()) or pidx.get(str(d['Customer gave: Email'] or '').lower()) for d in rows]
ci=len(h)+1; ws.cell(1,ci,'Triage status').font=Font(bold=True); ws.cell(1,ci+1,'In previous batch (User ID)').font=Font(bold=True)
for i,(s_,p_) in enumerate(zip(st,prev_hit),2): ws.cell(i,ci,s_); ws.cell(i,ci+1,p_)
ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width=48
ntick=sum(len(str(d['Ticket numbers']).split(',')) for d in rows)
sm=wb.create_sheet('Summary for support',0)
sm.append([title]); sm['A1'].font=Font(bold=True,size=13)
sm.append(['Tickets',ntick,'Users',len(rows)]); sm.append([])
sm.append(['Triage status','Users','What to do']); [setattr(c,'font',Font(bold=True)) for c in sm[4]]
todo={'A':'Send the green reply (we saw you logged in). No call needed unless they write again.',
      'B':'Send the reply with the correct email / instructions. Call back if no answer in 2 days.',
      'C':'Send the reply: register again (or send us phone/email). No CDC account exists.',
      'D':'Hold the reply until the tech team erases/releases the account, then send it.',
      'E':'Reply asking for phone + email, then rerun the check.'}
for k,n in sorted(collections.Counter(st).items()): sm.append([k,n,todo[k[0]]])
sm.append([]); sm.append(['CDC result','Users']); [setattr(c,'font',Font(bold=True)) for c in sm[sm.max_row]]
for k,n in collections.Counter(d['CDC result'] for d in rows).most_common(): sm.append([k,n])
li=sum(1 for d in rows if d['Logged in since cutoff']=='Yes')
sm.append([]); sm.append(['Logged in since 7.10', li, f"{li/len(rows):.0%}"])
if prev_path: sm.append(['Also appeared in the previous batch', sum(1 for p in prev_hit if p)])
sm.append(['Masked emails (verify identity before revealing)', sum(1 for d in rows if d['Email check'])])
sm.column_dimensions['A'].width=58; sm.column_dimensions['C'].width=80
for r in sm.iter_rows(min_row=4):
    for c in r: c.alignment=Alignment(wrap_text=True,vertical='top')
wb.save(inp); print(collections.Counter(st)); print('logged in', li, '/', len(rows))
