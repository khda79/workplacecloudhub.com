"""Version 1.0.0: mailbox populations, distinct from licensed people.

The current prepared contract does not capture governed service-account
classification for all AD users. Unknown service scope must never become zero.
Pending register hints quarantine candidates; they do not confirm identities.
"""

TABLE = 'Exchange Mailbox Reconciliation'


def population_expression(account_classification=False):
    expression='''VAR _scope=SELECTCOLUMNS('Exchange Mailbox Reconciliation',
    "Mailbox",'Exchange Mailbox Reconciliation'[Mailbox GUID],
    "Batch",'Exchange Mailbox Reconciliation'[Loaded Batch ID],
    "Object",LOWER(TRIM('Exchange Mailbox Reconciliation'[Object GUID])),
    "Address",LOWER(TRIM('Exchange Mailbox Reconciliation'[Primary SMTP Address])),
    "Retained",'Exchange Mailbox Reconciliation'[Retained On-premises])
VAR _exceptions=SELECTCOLUMNS(FILTER('Mailbox Migration Exceptions',
    'Mailbox Migration Exceptions'[Requires Review]=1),
    "ExceptionBatch",'Mailbox Migration Exceptions'[Loaded Batch ID],
    "Identity",LOWER(TRIM('Mailbox Migration Exceptions'[Mailbox Identity])),
    "Candidate",LOWER(TRIM('Mailbox Migration Exceptions'[Candidate Identity])),
    "ExceptionAddress",LOWER(TRIM('Mailbox Migration Exceptions'[Primary SMTP Address])))
VAR _batches=DISTINCT(SELECTCOLUMNS(_scope,"Batch",[Batch]))
VAR _valid=COUNTROWS(_batches)=1 && COUNTROWS(FILTER(_scope,[Batch]="" || [Mailbox]=""))=0
VAR _keys=DISTINCT(SELECTCOLUMNS(_scope,"Mailbox",[Mailbox],"Batch",[Batch]))
RETURN GENERATE(FILTER(_keys,_valid),
    VAR _mailbox=[Mailbox] VAR _batch=[Batch]
    VAR _rows=FILTER(_scope,[Mailbox]=_mailbox && [Batch]=_batch)
    VAR _retained=COUNTROWS(FILTER(_rows,[Retained]="Matched retained decision"))>0
    VAR _pending=SUMX(_rows,VAR _object=[Object] VAR _address=[Address]
        RETURN IF(COUNTROWS(FILTER(_exceptions,[ExceptionBatch]=_batch &&
            ((_object<>"" && ([Identity]="local_mailboxes|" & _object || [Candidate]="local_mailboxes|" & _object)) ||
             (_address<>"" && [ExceptionAddress]=_address))))>0,1,0))>0
    RETURN ROW("Population",SWITCH(TRUE(),_retained,"Register retained",_pending,"Register review","Potential before service review")))'''
    if account_classification:
        from exchange_account_classification import classified_scope
        expression=expression.replace('VAR _scope=','VAR _scopeRaw=',1)
        expression=expression.replace('VAR _batches=',classified_scope()+'\nVAR _batches=',1)
        expression=expression.replace('RETURN ROW("Population",',
            'VAR _service=MAXX(_rows,[Service Retained])\n'
            +'    VAR _known=MINX(_rows,[Classification Known])\n'
            +'    RETURN ROW("Service Retained",_service,"Classification Known",_known,"Population",')
        expression=expression.replace('_retained,"Register retained",_pending',
            '_retained,"Register retained",_service=1,"Service retained",_pending')
    return expression


def population_measures(account_classification=False):
    result=[]
    def add(name,expression,fmt='#,0',description=None):
        result.append(dict(name=name,expression=expression,formatString=fmt,
            displayFolder='Migration mailbox populations',description=description or
            'Distinct same-batch native mailbox GUIDs. Mailbox counts are not licence quantities. Confirmed register retention takes precedence over duplicate pending hints.'))
    for name,predicate in [('Total','TRUE()'),('Register Retained','[Population]="Register retained"'),
            ('Register Review','[Population]="Register review"'),
            ('Potential Before Service Review','[Population]="Potential before service review"')]:
        add('Exchange Population '+name,'VAR _mailboxes=('+population_expression(account_classification)+')\n'
            +'RETURN IF(COUNTROWS(_mailboxes)>0,COALESCE(COUNTROWS(FILTER(_mailboxes,'+predicate+')),0))')
    add('Exchange Population Service Label','"Not captured"','@',
        'The prepared AD user source lacks governed AccountType; Workforce excludes non-human service accounts. Generic accounts and managed-service object classes are not substitutes. No service count or complete retained union can be inferred.')
    add('Exchange Population Potential Label','''VAR _n=[Exchange Population Potential Before Service Review]
RETURN IF(ISBLANK(_n),"Unknown", "At most " & FORMAT(_n,"#,0"))''','@',
        'Upper bound for the currently unquarantined mailbox pool, before service-account exclusion and technical checks. Excludes confirmed retained mailboxes and pending exact native-ID/SMTP hints, deduplicated. Not an EXO licence requirement.')
    add('Exchange Scenario SE Required Label','''VAR _n=[Exchange Scenario SE Need]
RETURN IF(ISBLANK(_n),"Unknown",">= " & FORMAT(_n,"#,0"))''','@',
        'Known minimum estimated SE user-licence scope: distinct confirmed retained user identities. Service-account scope, special-mailbox accessors and other accessors are incomplete. Observed equivalents are not subtracted from this total. Not a net purchase quantity.')
    add('Exchange Population Note','''"INCOMPLETE SERVICE SCOPE: classification is not captured; Generic is not Service. Potential EXO excludes confirmed retention and " & FORMAT([Exchange Population Register Review],"0") & " other mailboxes with pending register hints; upper bound before service/technical review, not licence demand. SE = known minimum user scope; service accounts and other accessors uncounted. ESTIMATE: F3 <=2 GB, no archive/hold/delegation, frontline eligibility assumed; E3 for other known user cases <=100 GB. Assigned suites deducted; free tenant stock shared, assumed allocatable, not reserved. " & FORMAT([Exchange Scenario Register Review Rows],"0") & " register rows to reconcile; " & FORMAT([Exchange Scenario Review People],"0") & " unpriced people; " & FORMAT([Exchange Scenario Plan Remediation],"0") & " plan repairs, not purchases."''','@')
    if account_classification:
        expressions={
            'Exchange Population Service Label':'VAR _mailboxes=('+population_expression(True)+''')
VAR _known=COALESCE(SUMX(_mailboxes,[Classification Known]),0)
VAR _services=COALESCE(SUMX(_mailboxes,[Service Retained]),0)
RETURN IF(_known=0,"Not captured",IF(_known<COUNTROWS(_mailboxes),">= ","") & FORMAT(_services,"#,0"))''',
            'Exchange Population Note':'''"SERVICE SCOPE: only unique same-batch native AD classifications qualify; unknown types stay explicit. Register and service retention overlap is counted once in EXO and SE sizing. Potential EXO excludes both confirmed retention reasons and pending register hints; upper bound before technical review, not licence demand. SE = known minimum distinct retained users; other accessors uncounted. ESTIMATE: F3 <=2 GB, no archive/hold/delegation, frontline eligibility assumed; E3 for other known user cases <=100 GB. Assigned suites deducted; tenant stock shared, assumed allocatable, not reserved. Register rows to reconcile: " & FORMAT([Exchange Scenario Register Review Rows],"0")'''
        }
        for measure in result:
            if measure['name'] in expressions:measure['expression']=expressions[measure['name']]
            if measure['name']=='Exchange Population Service Label':
                measure['description']='Distinct mailbox scope with a unique same-batch native AD classification. Service/register overlap is displayed in both reason counts but excluded once from EXO and SE sizing. Missing or unknown classifications are not zero.'
    return result
