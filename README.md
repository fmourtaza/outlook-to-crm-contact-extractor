# Outlook to CRM Contact Extractor

Two local Python CLI scripts for migrating historical correspondence into a CRM. No cloud AI, no mailbox probing, no automatic CRM writes. Python 3.10+, Windows and Classic Outlook for extraction; reconciliation is cross-platform.

## Setup

```powershell
python -m pip install -r requirements.txt
python extract_outlook_contacts.py --pst "C:\\path\\archive.pst" --own-email user@example.com --output contacts.csv --signatures
```

Repeat `--own-email` for every mailbox alias. Use a backup PST; Outlook may mount it and leave it mounted. Omit `--signatures` to avoid reading message bodies. Subjects and attachment filenames are inspected in memory for aggregate English/French signals; no raw subjects, bodies or attachment contents are exported. Signatures are sampled from up to five recent incoming messages. Names, company, title, phone and country are candidates, not verified facts. Country requires consistent explicit signature country and international phone evidence; no HQ or domain suffix inference. Scoring measures interaction, not commercial value. Folder/item errors are counted; copied messages with different entry IDs may be counted twice.

## Reconcile

Export all CRM Contacts (including inactive, primary and alternate emails) and Accounts as static worksheets or CSVs with English column headers. Account Name, Email, First Name, Last Name are expected. XLSX reads the first worksheet. For optional reviewed domain mapping, create a local CSV with headers `Domain,Account Name,HQ ISO,Source URL,Reviewed`; use Yes only after reviewing the organization and its evidence. Do not assign a shared domain to an account when ambiguous.

```powershell
python prepare_crm_import.py --source contacts.csv --contacts existing_contacts.xlsx --accounts existing_accounts.xlsx --reviewed-domains reviewed_domains.csv --output output
```

Optional `--dns` queries DNS externally (domain names only); MX does not verify mailbox ownership or deliverability. Optional `--accept-country-candidates` maps evidence-supported individual country candidates to full names; default is blank. Unknown and conflicting countries must stay blank. No automatic website or city enrichment. Review output retains all source columns and decisions. Account proposals must be checked against CRM and imported first; then import contacts. Map fields explicitly in the CRM wizard (these are CSVs, not tenant-specific Excel templates). If Company Name lookup matches both Account and Contact, try Account Name for Account and Email for Contact; otherwise omit the lookup and assign accounts after import. Keep duplicate detection enabled. Retry only failed rows after checking existing records. Test a few records and verify permission fields before bulk import.

New contacts default to individual email Allow and bulk email Do Not Allow; review known objections and lawful basis before import. Existing records are never updated. No consent is inferred. Generic mailboxes, potential name duplicates and invalid addresses are held for review. CSVs contain personal data and may contain spreadsheet formula-like strings: use a trusted CSV reader or import as text, not as executable spreadsheet formulas. Keep outputs local and access-controlled; never commit PSTs, CRM exports, logs or generated files. The allowlist .gitignore helps but cannot protect already tracked or force-added files.

## Limitations

Initial release; heuristics are incomplete, language coverage limited, CRM column names may differ. Live Outlook COM and Dynamics import require testing on Windows and the target tenant. No guarantee of deduplication beyond supplied exports and simple rules. Review account proposals and all candidate fields before use. This is a consolidation of the workflow, not an exact reproduction of every historical enrichment iteration.
