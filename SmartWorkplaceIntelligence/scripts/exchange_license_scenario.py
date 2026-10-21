"""Version 1.1.0: explicit mailbox migration sizing assumptions, not an order.

Measures only: existing imported rows, keys, relationships and partitions remain
unchanged. Register review rows quarantine migration candidates, never grant rights.
"""

TABLE = 'Exchange Mailbox Reconciliation'


def people_expression(account_classification=False):
    expression='''VAR _scope=SELECTCOLUMNS('Exchange Mailbox Reconciliation',
    "Person",IF('Exchange Mailbox Reconciliation'[Cloud Account Key]<>"","cloud|" & 'Exchange Mailbox Reconciliation'[Cloud Account Key],
        IF('Exchange Mailbox Reconciliation'[AD Account Key]<>"","ad|" & 'Exchange Mailbox Reconciliation'[AD Account Key],"mailbox|" & 'Exchange Mailbox Reconciliation'[Mailbox GUID])),
    "Batch",'Exchange Mailbox Reconciliation'[Loaded Batch ID],
    "Cloud",'Exchange Mailbox Reconciliation'[Cloud Account Key],
    "Object",LOWER(TRIM('Exchange Mailbox Reconciliation'[Object GUID])),
    "Address",LOWER(TRIM('Exchange Mailbox Reconciliation'[Primary SMTP Address])),
    "Kind",'Exchange Mailbox Reconciliation'[Mailbox Type],
    "Size",'Exchange Mailbox Reconciliation'[Mailbox Size GB],
    "Archive",'Exchange Mailbox Reconciliation'[Archive State],
    "Hold",'Exchange Mailbox Reconciliation'[Hold and Retention],
    "Access",'Exchange Mailbox Reconciliation'[Full Access Count],
    "Send",'Exchange Mailbox Reconciliation'[Send As Count],
    "Behalf",'Exchange Mailbox Reconciliation'[Send On Behalf],
    "Retained",'Exchange Mailbox Reconciliation'[Retained On-premises],
    "Link",'Exchange Mailbox Reconciliation'[Account Link])
VAR _exceptions=SELECTCOLUMNS(FILTER('Mailbox Migration Exceptions',
    'Mailbox Migration Exceptions'[Requires Review]=1),
    "ExceptionBatch",'Mailbox Migration Exceptions'[Loaded Batch ID],
    "Identity",LOWER(TRIM('Mailbox Migration Exceptions'[Mailbox Identity])),
    "Candidate",LOWER(TRIM('Mailbox Migration Exceptions'[Candidate Identity])),
    "ExceptionAddress",LOWER(TRIM('Mailbox Migration Exceptions'[Primary SMTP Address])))
VAR _users=DISTINCT(SELECTCOLUMNS(FILTER(_scope,[Cloud]<>""),"User",[Cloud]))
VAR _batches=DISTINCT(SELECTCOLUMNS(_scope,"Batch",[Batch]))
VAR _plans=CALCULATETABLE(SELECTCOLUMNS(FILTER('Native User Service Plan',
    LOWER('Native User Service Plan'[IsEnabled] & "")="true" && 'Native User Service Plan'[AssignmentState]="Success" &&
    'Native User Service Plan'[ServicePlanName] IN {"EXCHANGE_S_STANDARD","EXCHANGE_S_ENTERPRISE","EXCHANGE_S_DESKLESS"}),
    "PlanUser",'Native User Service Plan'[TenantUserKey],"PlanBatch",'Native User Service Plan'[Loaded Batch ID],"Plan",'Native User Service Plan'[ServicePlanName]),
    TREATAS(_users,'Native User Service Plan'[TenantUserKey]),TREATAS(_batches,'Native User Service Plan'[Loaded Batch ID]))
VAR _allLicenses=CALCULATETABLE(SELECTCOLUMNS('Native User License',
    "LicenseUser",'Native User License'[TenantUserKey],"LicenseBatch",'Native User License'[Loaded Batch ID],
    "LicenseSku",'Native User License'[TenantSkuKey],"LicenseState",'Native User License'[AssignmentState]),
    TREATAS(_users,'Native User License'[TenantUserKey]),TREATAS(_batches,'Native User License'[Loaded Batch ID]))
VAR _licenses=FILTER(_allLicenses,[LicenseState] IN {"Active","Success"})
VAR _skus=SELECTCOLUMNS('Native License SKU',"SkuKey",'Native License SKU'[TenantSkuKey],
    "SkuBatch",'Native License SKU'[Loaded Batch ID],"SkuPart",'Native License SKU'[SkuPartNumber])
VAR _keys=DISTINCT(SELECTCOLUMNS(_scope,"Person",[Person],"Batch",[Batch]))
RETURN GENERATE(_keys,
    VAR _person=[Person] VAR _batch=[Batch]
    VAR _rows=FILTER(_scope,[Person]=_person && [Batch]=_batch)
    VAR _user=MAXX(_rows,[Cloud])
    VAR _local=COUNTROWS(FILTER(_rows,[Retained]="Matched retained decision"))>0
    VAR _remote=COUNTROWS(FILTER(_rows,[Retained]<>"Matched retained decision"))>0
    VAR _special=COUNTROWS(FILTER(_rows,[Kind]<>"UserMailbox"))>0
    VAR _pending=SUMX(_rows,VAR _object=[Object] VAR _address=[Address]
        RETURN IF(COUNTROWS(FILTER(_exceptions,[ExceptionBatch]=_batch &&
            ((_object<>"" && ([Identity]="local_mailboxes|" & _object || [Candidate]="local_mailboxes|" & _object)) ||
             (_address<>"" && [ExceptionAddress]=_address))))>0,1,0))>0
    VAR _unknown=COUNTROWS(FILTER(_rows,ISBLANK([Size]) || [Size]<0 || [Size]>100 ||
        NOT([Archive] IN {"None","Disabled","Local","Active","HostedProvisioned","HostedPending"}) ||
        NOT([Hold] IN {"None observed","Review - preserve settings"}) || ISBLANK([Access]) || ISBLANK([Send])))>0
    VAR _large=COUNTROWS(FILTER(_rows,[Size]>2 || NOT([Archive] IN {"None","Disabled"}) ||
        [Hold]="Review - preserve settings" || [Access]>0 || [Send]>0 || [Behalf]<>""))>0
    VAR _review=_batch="" || _user="" || COUNTROWS(FILTER(_rows,[Link]<>"Unique AD GUID + native SID"))>0 || (_local && _remote)
    VAR _bucket=SWITCH(TRUE(),_special,"Special",_local && NOT(_remote),"SE",_pending,"Register review",_review,"Identity review",
        _unknown,"Technical review",_large,"E3","F3")
    VAR _userPlans=FILTER(_plans,[PlanUser]=_user && [PlanBatch]=_batch && _user<>"")
    VAR _p2=COUNTROWS(FILTER(_userPlans,[Plan]="EXCHANGE_S_ENTERPRISE"))>0
    VAR _p1=COUNTROWS(FILTER(_userPlans,[Plan]="EXCHANGE_S_STANDARD"))>0
    VAR _kiosk=COUNTROWS(FILTER(_userPlans,[Plan]="EXCHANGE_S_DESKLESS"))>0
    VAR _assigned=SELECTCOLUMNS(FILTER(_licenses,[LicenseUser]=_user && [LicenseBatch]=_batch && _user<>""),"Key",[LicenseSku])
    VAR _suite=FILTER(_skus,[SkuKey] IN _assigned && [SkuBatch]=_batch)
    VAR _e3=COUNTROWS(FILTER(_suite,[SkuPart] IN {"SPE_E3","SPE_E5"}))>0
    VAR _f3=COUNTROWS(FILTER(_suite,[SkuPart]="SPE_F1"))>0
    VAR _covered=SWITCH(_bucket,"SE",IF(_e3 || _p1 || _p2,1,0),"E3",IF(_e3,1,0),
        "F3",IF(_f3 || _e3,1,0),0)
    VAR _failed=SELECTCOLUMNS(FILTER(_allLicenses,[LicenseUser]=_user && [LicenseBatch]=_batch &&
        NOT([LicenseState] IN {"Active","Success"})),"Key",[LicenseSku])
    VAR _failedTarget=COUNTROWS(FILTER(_skus,[SkuKey] IN _failed && [SkuBatch]=_batch &&
        ([SkuPart] IN {"SPE_E3","SPE_E5"} || (_bucket="F3" && [SkuPart]="SPE_F1"))))>0
    VAR _target=IF(_bucket IN {"E3","F3"} && _covered=0 && _failedTarget,"Assignment review",_bucket)
    VAR _planReview=_covered=1 && SWITCH(_bucket,"E3",NOT(_p2),"F3",NOT(_p2 || _p1 || _kiosk),FALSE())
    RETURN ROW("Product",_target,"Covered",_covered,"Plan Review",IF(_planReview,1,0)))'''
    if account_classification:
        from exchange_account_classification import classified_scope
        expression=expression.replace('VAR _scope=','VAR _scopeRaw=',1)
        expression=expression.replace('VAR _users=',classified_scope()+'\nVAR _users=',1)
        expression=expression.replace('[Retained]="Matched retained decision"))>0',
            '[Retained]="Matched retained decision" || [Service Retained]=1))>0',1)
        expression=expression.replace('[Retained]<>"Matched retained decision"))>0',
            '[Retained]<>"Matched retained decision" && [Service Retained]=0))>0',1)
        expression=expression.replace('VAR _review=_batch=',
            'VAR _classificationReview=COUNTROWS(FILTER(_rows,[Classification Known]=0))>0\n    VAR _review=_classificationReview || _batch=',1)
    return expression


def stock_expression(product):
    assert product in ('E3', 'F3')
    sku='SPE_E3' if product=='E3' else 'SPE_F1'
    return f'''VAR _batch=[Dataset Loaded Batch ID]
VAR _q=FILTER(ALL('License Report Qualification'),
    'License Report Qualification'[Loaded Batch ID]=_batch &&
    'License Report Qualification'[Scope]="M365_Licenses_ReportSkuCapacity.csv" &&
    'License Report Qualification'[State]="Qualified")
VAR _rows=FILTER(ALL('License SKU Capacity'),
    'License SKU Capacity'[Loaded Batch ID]=_batch && 'License SKU Capacity'[SKU Part Number]="{sku}" &&
    'License SKU Capacity'[Capacity Class]="Governed" && LOWER('License SKU Capacity'[Excluded] & "")="false")
VAR _snapshots=DISTINCT(SELECTCOLUMNS(_rows,"Snapshot",'License SKU Capacity'[Snapshot ID]))
VAR _ids=SELECTCOLUMNS(_rows,"Sku",'License SKU Capacity'[SKU ID])
VAR _valid=_batch<>"" && COUNTROWS(_q)=1 && COUNTROWS(_rows)>0 &&
    COUNTROWS(_snapshots)=1 && COUNTROWS(_ids)=COUNTROWS(DISTINCT(_ids)) &&
    COUNTROWS(FILTER(_rows,ISBLANK('License SKU Capacity'[Available Units]) ||
        'License SKU Capacity'[Available Units]<0 || 'License SKU Capacity'[SKU ID]="" || 'License SKU Capacity'[Snapshot ID]=""))=0
RETURN IF(_valid,SUMX(_rows,'License SKU Capacity'[Available Units]))'''


def scenario_measures(account_classification=False):
    result=[]
    def add(name,expression,fmt='#,0',description=None):
        result.append(dict(name=name,expression=expression,formatString=fmt,displayFolder='Migration licensing scenario',
            description=description or 'Estimate under approved mailbox sizing assumptions. Not a contractual order. Distinct same-batch people; retained and unresolved exception identities never migrate automatically.'))
    for product in ('E3','F3','SE'):
        for metric,condition,aggregate in [
            ('Need',f'[Product]="{product}"','COUNTROWS(_selected)'),
            ('Covered',f'[Product]="{product}"','SUMX(_selected,[Covered])')]:
            add(f'Exchange Scenario {product} {metric}',
                'VAR _people=('+people_expression(account_classification)+')\nVAR _selected=FILTER(_people,'+condition+')\n'
                +'RETURN IF([Dataset Loaded Batch ID]<>"",COALESCE('+aggregate+',0))')
        add(f'Exchange Scenario {product} Gap',f'VAR _need=[Exchange Scenario {product} Need] VAR _covered=[Exchange Scenario {product} Covered] RETURN IF(NOT ISBLANK(_need),MAX(0,_need-_covered))')
        if product!='SE':
            add(f'Exchange Scenario {product} Available',stock_expression(product))
            add(f'Exchange Scenario {product} Buy',f'VAR _gap=[Exchange Scenario {product} Gap] VAR _free=[Exchange Scenario {product} Available] RETURN IF(NOT ISBLANK(_gap) && NOT ISBLANK(_free),MAX(0,_gap-_free))')
            add(f'Exchange Scenario {product} Buy Label',f'VAR _n=[Exchange Scenario {product} Buy] RETURN IF(ISBLANK(_n),"To qualify",FORMAT(_n,"#,0"))','@')
        else:
            add('Exchange Scenario SE Buy Label','"To qualify"','@','Existing CAL stock, active SA, accessors and base/additive rights are not inventoried. Gap is exposure before those rights; never a net purchase quantity.')
    add('Exchange Scenario Review People','VAR _people=('+people_expression(account_classification)+')\nRETURN IF([Dataset Loaded Batch ID]<>"",COALESCE(COUNTROWS(FILTER(_people,[Product] IN {"Register review","Identity review","Technical review","Assignment review"})),0))')
    add('Exchange Scenario Plan Remediation','VAR _people=('+people_expression(account_classification)+')\nRETURN IF([Dataset Loaded Batch ID]<>"",COALESCE(SUMX(_people,[Plan Review]),0))')
    add('Exchange Scenario Register Review Rows','VAR _batch=[Dataset Loaded Batch ID] RETURN IF(_batch<>"",CALCULATE(COUNTROWS(\'Mailbox Migration Exceptions\'),REMOVEFILTERS(\'Mailbox Migration Exceptions\'),\'Mailbox Migration Exceptions\'[Loaded Batch ID]=_batch,\'Mailbox Migration Exceptions\'[Requires Review]=1))')
    add('Exchange Scenario Row Target','VAR _people=('+people_expression(account_classification)+')\nVAR _products=DISTINCT(SELECTCOLUMNS(_people,"Target",[Product])) RETURN IF(COUNTROWS(_products)=1,CONCATENATEX(_products,[Target],""),"Multiple targets / review")','@')
    add('Exchange Scenario Note','''"ESTIMATE: F3 <=2 GB, no archive/hold/delegation; frontline eligibility assumed. E3 for other known user cases <=100 GB. Assigned suites deducted even if a service plan needs repair; tenant stock assumed allocatable, shared across countries, not reserved. SE = confirmed retained users; CAL stock/SA and other accessors unknown. Unpriced people: " & FORMAT([Exchange Scenario Review People],"0") & " | Register rows to reconcile: " & FORMAT([Exchange Scenario Register Review Rows],"0") & " | Plan repairs (not purchases): " & FORMAT([Exchange Scenario Plan Remediation],"0")''','@')
    if account_classification:
        from exchange_mailbox_populations import population_expression
        for measure in result:
            if measure['name'].endswith('Buy Label') and ' SE ' not in measure['name']:
                product='E3' if ' E3 ' in measure['name'] else 'F3'
                measure['expression']='VAR _mailboxes=('+population_expression(True)+')\n'+f'''VAR _n=[Exchange Scenario {product} Buy]
VAR _known=COALESCE(SUMX(_mailboxes,[Classification Known]),0)
RETURN IF(ISBLANK(_n) || _known=0,"Not captured",IF(_known<COUNTROWS(_mailboxes),">= ","") & FORMAT(_n,"#,0"))'''
            if measure['name']=='Exchange Scenario Note':
                measure['expression']=measure['expression'].replace('SE = confirmed retained users','SE = register/service retained users')
    return result
