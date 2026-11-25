"""Extract local PST correspondents with optional signature analysis."""
import argparse, re, csv
from collections import defaultdict, Counter
from datetime import datetime
from pathlib import Path

EMAIL = re.compile(r'[^\s@<>;,]+@[^\s@<>;,]+\.[^\s@<>;,]+')
SMTP = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
RULES = {'Commercial': r'purchase order|bon de commande|invoice|facture|\bsow\b', 'Sales': r'quotation|proposal|devis|\brfp\b', 'Partner': r'partnership|partenariat|reseller', 'Institutional': r'funding|subvention|accelerator', 'Network': r'webinar|conference|conférence|workshop', 'Research': r'research|recherche|consortium', 'Automated': r'newsletter|unsubscribe|automatic reply'}

def norm(v):
    s=str(v or '').strip().lower().removeprefix('smtp:')
    return s if EMAIL.fullmatch(s) else ''

def prop(o,k,default=None):
    try: return getattr(o,k)
    except Exception: return default

def smtp(entry):
    try:
        if entry.Type == 'EX':
            u=entry.GetExchangeUser()
            if u and norm(u.PrimarySmtpAddress): return norm(u.PrimarySmtpAddress)
    except Exception: pass
    try: return norm(entry.PropertyAccessor.GetProperty(SMTP)) or norm(entry.Address)
    except Exception: return norm(prop(entry,'Address'))

def mailbox(e):
    local=e.split('@')[0]
    if re.match(r'^(no[._-]?reply|do[._-]?not[._-]?reply|mailer[._-]?daemon|postmaster|notifications?)(?:[._+\d-]|$)',local): return 'Automated candidate'
    if re.match(r'^(info|sales|support|contact|admin|office|accounts|billing|marketing|hr|team|helpdesk)(?:[._+\d-]|$)',local): return 'Generic candidate'
    return 'Individual or unknown'

def score(s,r,first,last,now):
    n=s+r; age=(now-last).days if last else 99999; span=(last-first).days if first and last else 0
    return (30 if s and r else 0)+next((p for t,p in [(50,25),(20,20),(10,15),(5,10),(1,5)] if n>=t),0)+next((p for t,p in [(30,30),(90,25),(180,15),(365,10)] if 0<=age<=t),0)+next((p for t,p in [(365,15),(180,10),(30,5)] if span>=t),0)

def signature(body):
    import phonenumbers as pn
    import pycountry
    head=re.split(r'(?im)^\s*(?:from:|de\s*:|sent:|on .+ wrote:|[-_]{5,})',body or '',maxsplit=1)[0]
    parts=re.split(r'(?im)^\s*(?:best regards|kind regards|regards|sincerely|cordialement|bien à vous|merci)[,! .]*$',head)
    if len(parts)<2: return set(),set(),set(),set(),set()
    tail='\n'.join(parts[-1].splitlines()[:20]); countries=set(); phones=set(); regions=set(); titles=set(); companies=set()
    for line in tail.splitlines():
        line=line.strip()
        if re.search(r'confidential|unsubscribe|désabonnement',line,re.I): break
        for m in pn.PhoneNumberMatcher(line,None):
            if pn.is_valid_number(m.number):
                phones.add(pn.format_number(m.number,pn.PhoneNumberFormat.E164)); regions.add(pn.region_code_for_number(m.number))
        for c in pycountry.countries:
            if line.casefold()==c.name.casefold() or line.casefold().endswith(', '+c.name.casefold()): countries.add(c.alpha_2)
        if len(line)<100 and re.search(r'\b(manager|director|engineer|consultant|officer|directeur|responsable|founder)\b',line,re.I) and '@' not in line: titles.add(line)
        if len(line)<100 and re.search(r'\b(ltd|limited|inc|gmbh|sas|sarl|llp|corporation)\.?$',line,re.I): companies.add(line)
    return countries,regions,phones,titles,companies

def main():
    p=argparse.ArgumentParser(description=__doc__); p.add_argument('--pst',type=Path,required=True); p.add_argument('--own-email',action='append',required=True); p.add_argument('--output',type=Path,required=True); p.add_argument('--signatures',action='store_true'); a=p.parse_args()
    own={norm(x) for x in a.own_email}; assert '' not in own, 'Invalid own email'
    if not a.pst.is_file(): p.error('PST not found')
    import win32com.client
    ns=win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    def find(): return next((s for s in ns.Stores if str(prop(s,'FilePath','')).casefold()==str(a.pst.resolve()).casefold()),None)
    store=find()
    if store is None: ns.AddStore(str(a.pst.resolve())); store=find()
    if store is None: raise RuntimeError('Open PST manually in Classic Outlook')
    try: sent_id=store.GetDefaultFolder(5).EntryID
    except Exception: sent_id=None
    stack=[(store.GetRootFolder(),False)]; records=defaultdict(lambda:dict(names=Counter(),sent=0,received=0,dates=[],signals=defaultdict(set),samples=[],meetings=0)); stats=Counter(); seen=set()
    while stack:
        folder,in_sent=stack.pop(); in_sent=in_sent or (sent_id is not None and prop(folder,'EntryID')==sent_id) or str(prop(folder,'Name','')).casefold() in {'sent items','sent mail','éléments envoyés'}
        try: stack.extend((f,in_sent) for f in folder.Folders); items=folder.Items; count=items.Count
        except Exception: stats['folder errors']+=1; continue
        stats['folders']+=1
        for i in range(count,0,-1):
            try:
                item=items.Item(i); cls=int(item.Class)
                if cls not in (43,53,54,55,56,57): continue
                eid=str(item.EntryID)
                if eid in seen: continue
                seen.add(eid); stats['items']+=1
                sender=norm(prop(item,'SenderEmailAddress')) or smtp(prop(item,'Sender')); outbound=sender in own or in_sent
                targets={}
                if outbound:
                    for r in item.Recipients:
                        if int(r.Type) in (1,2):
                            e=norm(prop(r,'Address')) or smtp(prop(r,'AddressEntry')); targets[e]=str(prop(r,'Name',''))
                else: targets[sender]=str(prop(item,'SenderName',''))
                dt=prop(item,'SentOn' if outbound else 'ReceivedTime'); dt=dt.replace(tzinfo=None) if isinstance(dt,datetime) and dt.year>1900 else None
                context=str(prop(item,'Subject',''))+' '+ ' '.join(str(prop(x,'FileName','')) for x in (prop(item,'Attachments',[]) or [])); thread=str(prop(item,'ConversationID') or eid)
                for e,name in targets.items():
                    if not e or e in own: continue
                    r=records[e]; r['names'][name]+=1; r['sent' if outbound else 'received']+=1
                    if dt: r['dates'].append(dt)
                    if cls!=43: r['meetings']+=1
                    for k,pattern in RULES.items():
                        if re.search(pattern,context,re.I): r['signals'][k].add(thread)
                    if a.signatures and not outbound and cls==43 and dt:
                        r['samples'].append((dt,eid)); r['samples']=sorted(r['samples'],reverse=True)[:5]
            except Exception: stats['item errors']+=1
    fields=['Email','Name','Domain','First Name','Last Name','Emails Sent','Emails Received','First Interaction','Last Interaction','Score','Mailbox','Relationship candidate','Signals','Meetings','Company candidate','Job Title candidate','Phone candidates','Country candidate ISO','Marketing permission']
    with a.output.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=fields); w.writeheader()
        for e,r in sorted(records.items()):
            name=next((n for n,_ in r['names'].most_common() if n),''); parts=name.split(); kind=mailbox(e); first,last=(parts[0],' '.join(parts[1:])) if len(parts)>1 and '@' not in name and kind=='Individual or unknown' else ('','')
            dates=r['dates']; lo=min(dates,default=None); hi=max(dates,default=None); sets=[set() for _ in range(5)]
            for _,eid in r['samples']:
                try:
                    for dest,src in zip(sets,signature(ns.GetItemFromID(eid,store.StoreID).Body)): dest.update(src)
                except Exception: stats['signature errors']+=1
            countries,regions,phones,titles,companies=sets; country=next(iter(countries)) if len(countries)==1 and countries==regions else ''
            signals=r['signals']; candidate=next(iter(signals)) if len(signals)==1 else 'Mixed signals' if signals else 'Unknown'
            w.writerow(dict(zip(fields,[e,name,e.split('@')[1],first,last,r['sent'],r['received'],lo,hi,score(r['sent'],r['received'],lo,hi,datetime.now()),kind,candidate,'; '.join(f'{k}: {len(v)}' for k,v in sorted(signals.items())),r['meetings'],next(iter(companies)) if len(companies)==1 else '',next(iter(titles)) if len(titles)==1 else '', '; '.join(sorted(phones)),country,'Unknown'])))
    print(dict(stats), 'contacts:',len(records))

if __name__=='__main__': main()
