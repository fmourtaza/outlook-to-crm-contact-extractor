"""Reconcile extracted contacts against CRM exports; produce review and import CSVs."""
import argparse,csv,re
from pathlib import Path
from collections import Counter,defaultdict
import pycountry

def read(path):
    if path.suffix.lower()=='.xlsx':
        from openpyxl import load_workbook
        w=load_workbook(path,read_only=True,data_only=True); rows=w.active.values; headers=next(rows); result=[dict(zip(headers,r)) for r in rows]; w.close(); return result
    with path.open(encoding='utf-8-sig',newline='') as f:return list(csv.DictReader(f))
def val(r,k):return str(r.get(k) or '').strip()
def key(s):return s.strip().casefold()
def write(path,rows,fields):
    with path.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore'); w.writeheader(); w.writerows(rows)
def country(code):
    c=pycountry.countries.get(alpha_2=code.upper()); return c.name if c else ''
def dns_status(domain):
    import dns.resolver,dns.exception
    try:
        mx=dns.resolver.resolve(domain,'MX',lifetime=5); return 'Null MX' if all(str(x.exchange)=='.' for x in mx) else 'MX present'
    except dns.resolver.NXDOMAIN:return 'NXDOMAIN'
    except dns.resolver.NoAnswer:return 'No MX; address fallback possible'
    except dns.exception.DNSException:return 'Inconclusive'
def main():
    p=argparse.ArgumentParser(description=__doc__); p.add_argument('--source',type=Path,required=True); p.add_argument('--contacts',type=Path,required=True); p.add_argument('--accounts',type=Path,required=True); p.add_argument('--reviewed-domains',type=Path); p.add_argument('--output',type=Path,required=True); p.add_argument('--dns',action='store_true'); p.add_argument('--accept-country-candidates',action='store_true'); a=p.parse_args()
    source,existing,accounts=map(read,(a.source,a.contacts,a.accounts)); a.output.mkdir(parents=True,exist_ok=True)
    emails={key(val(r,k)) for r in existing for k in ('Email','Email Address 2','Email Address 3') if val(r,k)}; names={(key(val(r,'First Name')),key(val(r,'Last Name'))) for r in existing}; account_counts=Counter(key(val(r,'Account Name')) for r in accounts if val(r,'Account Name'))
    reviewed=defaultdict(list)
    if a.reviewed_domains:
        for r in read(a.reviewed_domains):
            if key(val(r,'Reviewed'))=='yes' and val(r,'Source URL').startswith(('https://','http://')) and val(r,'Account Name'):reviewed[key(val(r,'Domain'))].append(r)
    dns={}; audit=[]; imports=[]; proposals={}; seen=set(); new_names=set()
    for r in source:
        e=key(val(r,'Email')); d=e.rpartition('@')[2]; first=val(r,'First Name'); last=val(r,'Last Name'); name=(key(first),key(last)); reason='Import'; company=''; status='Not checked'
        if a.dns and d:
            if d not in dns:dns[d]=dns_status(d)
            status=dns[d]
        if e in emails:reason='Existing email'
        elif e in seen:reason='Repeated source email'
        elif not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+',e) or not last:reason='Invalid email or missing surname'
        elif name in names or name in new_names:reason='Potential name duplicate'
        elif val(r,'Mailbox') in ('Generic candidate','Automated candidate'):reason='Mailbox review'
        elif status in ('NXDOMAIN','Null MX'):reason='Domain hold'
        seen.add(e)
        candidates=reviewed.get(d,[])
        if len(candidates)==1:
            x=candidates[0]; company=val(x,'Account Name'); count=account_counts[key(company)]
            if count>1:company=''
            elif count==0 and reason=='Import':proposals[key(company)]={'Account Name':company,'Address 1: Country/Region':country(val(x,'HQ ISO'))}
        loc=country(val(r,'Country candidate ISO')) if a.accept_country_candidates else ''
        phone=val(r,'Phone candidates'); business=''; mobile=''
        if phone and ';' not in phone:
            import phonenumbers as pn
            try:
                n=pn.parse(phone,None); t=pn.number_type(n)
                if pn.is_valid_number(n):
                    if t==pn.PhoneNumberType.MOBILE:mobile=phone
                    elif t==pn.PhoneNumberType.FIXED_LINE:business=phone
            except pn.NumberParseException:pass
        audit.append({**r,'Decision':reason,'DNS status':status,'Mapped Account':company,'Mapped country':loc})
        if reason=='Import':
            new_names.add(name); imports.append({'First Name':first,'Last Name':last,'Email':e,'Company Name':company,'Business Phone':business,'Mobile Phone':mobile,'Job Title':val(r,'Job Title candidate'),'Address 1: Country/Region':loc,'Do not allow Emails':'Allow','Do not allow Bulk Emails':'Do Not Allow'})
    fields=['First Name','Last Name','Email','Company Name','Business Phone','Mobile Phone','Job Title','Address 1: Country/Region','Do not allow Emails','Do not allow Bulk Emails']
    write(a.output/'contacts_import.csv',imports,fields); write(a.output/'accounts_proposed.csv',list(proposals.values()),['Account Name','Address 1: Country/Region']); write(a.output/'reconciliation.csv',audit,list(dict.fromkeys(k for r in audit for k in r)) or ['Decision']); print(dict(Counter(r['Decision'] for r in audit)))
if __name__=='__main__':main()
