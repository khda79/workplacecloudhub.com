"""Same-batch mailbox/account/license review model. No licensing purchases inferred.

Builds definitions only; applying them and calculating a model are separate steps.
Native immutable keys, unique matches and explicit successful plans are required.
"""

import re

TABLE = "Exchange Mailbox Reconciliation"
SERVER_TABLE = "Exchange Mailbox Server Reference"


def server_expression():
    return '''let
    Readers = CreatePreparedReaders(),
    Source = Readers[ReadMailboxMigrationCsv](
        Readers[GetPreparedEvidenceFile]("Source-Exchange_OnPrem_Mailboxes_AllDomains.csv"),
        {"TenantKey", "ObjectGUID", "RecipientType"}, {"ServerName"}),
    TenantChecked = if List.AllTrue(List.Transform(Source[TenantKey], each _ = PreparedExpectedTenantKey))
        then Source else error "Mailbox server source tenant mismatch",
    Local = Table.SelectRows(TenantChecked, each not Text.StartsWith(Text.Lower([RecipientType]), "remote")),
    Keyed = Table.TransformColumns(Local, {{"ObjectGUID", each if _ = null then "" else Text.Lower(Text.Trim(_)), type text}}),
    Keys = List.Buffer(Keyed[ObjectGUID]),
    Checked = if not List.Contains(Keys, "") and List.Count(Keys) = List.Count(List.Distinct(Keys))
        then Keyed else error "Missing or duplicate native mailbox object identity",
    Projected = Table.RenameColumns(Table.SelectColumns(Checked, {"ObjectGUID", "ServerName"}, MissingField.UseNull),
        {{"ObjectGUID", "Object GUID"}, {"ServerName", "Server Name"}}),
    Display = Table.TransformColumns(Projected, {{"Server Name", each if _ = null or Text.Trim(_) = "" then "Not observed" else _, type text}})
in
    Readers[Finish](Display)'''


def table_expression():
    body = '''GENERATE(
    SELECTCOLUMNS('Mailbox Migration Inventory',
        "Mailbox GUID", 'Mailbox Migration Inventory'[Mailbox GUID],
        "Object GUID", 'Mailbox Migration Inventory'[Object GUID],
        "Mailbox", 'Mailbox Migration Inventory'[Display Name],
        "Primary SMTP Address", 'Mailbox Migration Inventory'[Primary SMTP Address],
        "Country", 'Mailbox Migration Inventory'[Country],
        "Exchange Version", 'Mailbox Migration Inventory'[Exchange Version],
        "Mailbox Type", 'Mailbox Migration Inventory'[Mailbox Type],
        "Mailbox Size GB", 'Mailbox Migration Inventory'[Mailbox Size GB],
        "Loaded Batch ID", 'Mailbox Migration Inventory'[Loaded Batch ID],
        "Archive State", 'Mailbox Migration Inventory'[Archive State],
        "Hold and Retention", 'Mailbox Migration Inventory'[Hold and Retention],
        "Full Access Count", 'Mailbox Migration Inventory'[Full Access Count],
        "Send As Count", 'Mailbox Migration Inventory'[Send As Count],
        "Send On Behalf", 'Mailbox Migration Inventory'[Send On Behalf]),
    VAR _object = LOWER(TRIM([Object GUID]))
    VAR _batch = [Loaded Batch ID]
    VAR _servers = FILTER(ALL('Exchange Mailbox Server Reference'), _object <> "" && _batch <> "" &&
        'Exchange Mailbox Server Reference'[Object GUID] = _object && 'Exchange Mailbox Server Reference'[Loaded Batch ID] = _batch)
    VAR _ad = FILTER(ALL('AD User Evidence'), _object <> "" && _batch <> "" &&
        LOWER(TRIM('AD User Evidence'[ObjectGUID])) = _object && 'AD User Evidence'[Loaded Batch ID] = _batch)
    VAR _adUnique = COUNTROWS(_ad) = 1
    VAR _adKey = IF(_adUnique, MAXX(_ad, 'AD User Evidence'[TenantADUserKey]))
    VAR _cloudKey = IF(_adUnique && MAXX(_ad, 'AD User Evidence'[CloudMatchStatus]) = "Unique native SID",
        MAXX(_ad, 'AD User Evidence'[TenantUserKey]))
    VAR _cloud = FILTER(ALL('Native User Evidence'), _cloudKey <> "" &&
        'Native User Evidence'[TenantUserKey] = _cloudKey && 'Native User Evidence'[Loaded Batch ID] = _batch)
    VAR _cloudUnique = COUNTROWS(_cloud) = 1
    VAR _user = IF(_cloudUnique, _cloudKey)
    VAR _licenses = FILTER(ALL('Native User License'), _user <> "" &&
        'Native User License'[TenantUserKey] = _user && 'Native User License'[Loaded Batch ID] = _batch)
    VAR _skuKeys = SELECTCOLUMNS(_licenses, "Key", 'Native User License'[TenantSkuKey])
    VAR _skus = FILTER(ALL('Native License SKU'), 'Native License SKU'[TenantSkuKey] IN _skuKeys &&
        'Native License SKU'[Loaded Batch ID] = _batch)
    VAR _plans = FILTER(ALL('Native User Service Plan'), _user <> "" &&
        'Native User Service Plan'[TenantUserKey] = _user && 'Native User Service Plan'[Loaded Batch ID] = _batch &&
        LEFT('Native User Service Plan'[ServicePlanName], 11) = "EXCHANGE_S_")
    VAR _success = FILTER(_plans, LOWER('Native User Service Plan'[IsEnabled]) = "true" &&
        'Native User Service Plan'[AssignmentState] = "Success")
    VAR _p1 = COUNTROWS(FILTER(_success, 'Native User Service Plan'[ServicePlanName] = "EXCHANGE_S_STANDARD")) > 0
    VAR _p2 = COUNTROWS(FILTER(_success, 'Native User Service Plan'[ServicePlanName] = "EXCHANGE_S_ENTERPRISE")) > 0
    VAR _kiosk = COUNTROWS(FILTER(_success, 'Native User Service Plan'[ServicePlanName] = "EXCHANGE_S_DESKLESS")) > 0
    VAR _foundation = COUNTROWS(FILTER(_success, 'Native User Service Plan'[ServicePlanName] = "EXCHANGE_S_FOUNDATION")) > 0
    VAR _planNames = DISTINCT(SELECTCOLUMNS(_success, "Plan", 'Native User Service Plan'[ServicePlanName]))
    VAR _skuNames = DISTINCT(SELECTCOLUMNS(_skus, "SKU", 'Native License SKU'[SkuPartNumber]))
    VAR _exception = FILTER(ALL('Mailbox Migration Exceptions'), _object <> "" &&
        'Mailbox Migration Exceptions'[Loaded Batch ID] = _batch &&
        LOWER('Mailbox Migration Exceptions'[Mailbox Identity]) = "local_mailboxes|" & _object &&
        'Mailbox Migration Exceptions'[Status] = "Matched local" && 'Mailbox Migration Exceptions'[Retained] = 1 &&
        'Mailbox Migration Exceptions'[Requires Review] = 0)
    VAR _retained = COUNTROWS(_exception) > 0
    VAR _userMailbox = [Mailbox Type] = "UserMailbox"
    VAR _size = [Mailbox Size GB]
    VAR _sizeUnknown = ISBLANK(_size) || _size < 0
    VAR _features = NOT([Archive State] IN {"None", "Disabled", "Not observed", "Unknown", ""}) ||
        [Hold and Retention] = "Review - preserve settings" || [Full Access Count] > 0 ||
        [Send As Count] > 0 || [Send On Behalf] <> ""
    RETURN ROW(
        "Server Name", IF(COUNTROWS(_servers)=1,MAXX(_servers,'Exchange Mailbox Server Reference'[Server Name]),"Not observed"),
        "AD Account Key", _adKey,
        "Cloud Account Key", _user,
        "Account Link", IF(NOT(_adUnique), IF(COUNTROWS(_ad)>1,"Ambiguous AD identity","No same-batch AD identity"),
            IF(_cloudUnique,"Unique AD GUID + native SID",IF(COUNTROWS(_cloud)>1,"Ambiguous cloud identity","AD only; cloud unresolved"))),
        "AD Account", IF(_adUnique, MAXX(_ad, 'AD User Evidence'[UserPrincipalName])),
        "Entra Account", IF(_cloudUnique, MAXX(_cloud, 'Native User Evidence'[UserPrincipalName])),
        "AD Job Title", IF(_adUnique, MAXX(_ad, 'AD User Evidence'[Title])),
        "Entra Job Title", IF(_cloudUnique, MAXX(_cloud, 'Native User Evidence'[JobTitle])),
        "Observed SKUs", IF(_user="","Unknown - unresolved account",COALESCE(CONCATENATEX(_skuNames,[SKU],"; ",[SKU],ASC),"None observed")),
        "Successful Exchange Plans", IF(_user="","Unknown - unresolved account",COALESCE(CONCATENATEX(_planNames,[Plan],"; ",[Plan],ASC),"None observed")),
        "Other Exchange Plan States", CONCATENATEX(DISTINCT(SELECTCOLUMNS(FILTER(_plans,
            LOWER('Native User Service Plan'[IsEnabled]) <> "true" || 'Native User Service Plan'[AssignmentState] <> "Success"),
            "State", 'Native User Service Plan'[ServicePlanName] & ": " & 'Native User Service Plan'[AssignmentState])),[State],"; ",[State],ASC),
        "Mailbox Plan Observation", SWITCH(TRUE(),_user="","Unknown - unresolved account",_p2,"Plan 2 observed",_p1,"Plan 1 observed",_kiosk,"Kiosk observed",_foundation,"Foundation only - no mailbox right","No mailbox plan observed"),
        "Retained On-premises", IF(_retained,"Matched retained decision","No matched retained decision; review register"),
        "F3 Review", SWITCH(TRUE(),NOT(_userMailbox),"Special mailbox - review accessors and features",
            _sizeUnknown,"Unknown mailbox size - review",_size>2,"Above 2 GB - Kiosk alone insufficient",
            _features,"Feature or delegation review required","At most 2 GB - frontline eligibility and usage unverified"),
        "E3 Review", SWITCH(TRUE(),NOT(_userMailbox),"Special mailbox rules; no automatic E3 requirement",
            NOT(_sizeUnknown) && _size>100,"Above 100 GB - quota and archive design review",
            NOT(_sizeUnknown) && _size>50,"Above 50 GB - Plan 2 or equivalent review",
            _features,"Archive, hold and delegation rights review","No E3 purchase conclusion; assess complete workload"),
        "SE CAL Review", IF(_p2,"Plan 2 observed - base/additive equivalency to verify",
            IF(_p1,"Plan 1 observed - base equivalency to verify","CAL coverage unverified; reconcile contracts and accessors")),
        "Licensing Decision", "Unqualified - contract, active coverage, accessors and frontline eligibility required")
)'''
    # Materialize narrow source projections once, outside the mailbox iterator.
    # SELECTCOLUMNS preserves duplicate rows (uniqueness checks must not hide
    # them). Snapshot columns use unqualified references inside their iterators.
    # Whole-table ALL would expand calculated
    # columns and repeatedly evaluate unrelated identity/license dependencies.
    # Exchange plans are filtered once before the per-account lookup.
    snapshots = []
    for index, table in enumerate((SERVER_TABLE, 'AD User Evidence',
            'Native User Evidence', 'Native User License', 'Native License SKU',
            'Native User Service Plan', 'Mailbox Migration Exceptions')):
        fields = sorted(set(re.findall(re.escape("'" + table + "'") + r'\[([^\]]+)\]', body)))
        pairs = ', '.join('"' + field + '", \'' + table + '\'[' + field + ']' for field in fields)
        projection = "SELECTCOLUMNS('" + table + "', " + pairs + ')'
        if table == 'Native User Service Plan':
            projection = 'FILTER(' + projection + ', LEFT([ServicePlanName],11)="EXCHANGE_S_")'
        name = '_source' + str(index)
        snapshots.append('VAR ' + name + ' = ' + projection)
        body = body.replace("ALL('" + table + "')", name)
        body = re.sub(re.escape("'" + table + "'") + r'\[([^\]]+)\]', r'[\1]', body)
    return '\n'.join(snapshots) + '\nRETURN\n' + body


def measures():
    t = "'" + TABLE + "'"
    definitions = {
        'Exchange Reconciled Mailboxes': f'COUNTROWS({t})',
        'Exchange Cloud Linked Mailboxes': f'CALCULATE(COUNTROWS({t}),KEEPFILTERS({t}[Account Link]="Unique AD GUID + native SID"))',
        'Exchange Observed Plan Mailboxes': f'CALCULATE(COUNTROWS({t}),KEEPFILTERS({t}[Mailbox Plan Observation] IN {{"Plan 1 observed","Plan 2 observed","Kiosk observed"}}))',
        'Exchange Retained Decision Mailboxes': f'CALCULATE(COUNTROWS({t}),KEEPFILTERS({t}[Retained On-premises]="Matched retained decision"))',
        'Exchange Referenced Servers': f'CALCULATE(DISTINCTCOUNT({t}[Server Name]),KEEPFILTERS({t}[Server Name]<>""),KEEPFILTERS({t}[Server Name]<>"Not observed"))',
        'Exchange Kiosk Size Reviews': f'CALCULATE(COUNTROWS({t}),KEEPFILTERS({t}[Mailbox Type]="UserMailbox"),KEEPFILTERS({t}[Mailbox Plan Observation]="Kiosk observed"),KEEPFILTERS({t}[Mailbox Size GB]>2))',
        'Exchange Licensing Review Note': '"Mailbox observations, not purchase quantities | F3 eligibility and contracts unverified | Server names are mailbox references, not a complete active-server inventory."',
    }
    return [dict(name=n, expression=e, formatString='@' if n.endswith('Note') else '#,0',
                 displayFolder='Reconciliation', description='Same-batch mailbox grain. Not distinct licensed users, accessors, required CALs or purchase quantities.')
            for n, e in definitions.items()]
