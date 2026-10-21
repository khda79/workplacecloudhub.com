"""Version 1.0.0: guarded optional AD classification reader for future batches.

Keeps all AD accounts, including non-human/disabled accounts. No name heuristics.
The callable reader can be installed without processing a new imported table.
"""

EXPRESSION='ReadPreparedADAccountClassification'
TABLE='AD Account Classification'


def reader_expression():
    return '''() as table =>
let
    Readers = CreatePreparedReaders(),
    Source = Readers[ReadMailboxMigrationCsv](Readers[GetPreparedEvidenceFile]("ADUserSource.csv"),
        {"TenantKey", "ObjectGUID", "TenantADUserKey"}, {"AccountType"}),
    TenantChecked = if List.AllTrue(List.Transform(Source[TenantKey], each _ = PreparedExpectedTenantKey))
        then Source else error "AD account classification tenant mismatch",
    Keyed = Table.TransformColumns(TenantChecked,
        {{"ObjectGUID", each if _ = null then "" else Text.Lower(Text.Trim(_)), type text},
         {"TenantADUserKey", each if _ = null then "" else Text.Trim(_), type text}}),
    Objects = List.Buffer(Keyed[ObjectGUID]),
    AccountKeys = List.Buffer(Keyed[TenantADUserKey]),
    Checked = if not List.Contains(Objects, "") and not List.Contains(AccountKeys, "") and
        List.Count(Objects) = List.Count(List.Distinct(Objects)) and
        List.Count(AccountKeys) = List.Count(List.Distinct(AccountKeys))
        then Keyed else error "Missing or duplicate native AD account classification identity",
    Projected = Table.RenameColumns(Table.SelectColumns(Checked,
        {"ObjectGUID", "TenantADUserKey", "AccountType"}, MissingField.UseNull),
        {{"ObjectGUID", "Object GUID"}, {"TenantADUserKey", "AD Account Key"}, {"AccountType", "Account Type"}}),
    Normalized = Table.TransformColumns(Projected,
        {{"Account Type", each if _ = null or Text.Trim(_) = "" then "Not captured" else Text.Trim(_), type text}}),
    Service = Table.AddColumn(Normalized, "Service Classification", each
        if [Account Type] = "Not captured" then "Not captured"
        else if List.Contains({"service", "service account"}, Text.Lower([Account Type])) then "Service"
        else if List.Contains({"named account", "ext account", "generic account", "admin account", "system account", "resource account"},
            Text.Lower([Account Type])) then "Other classified account"
        else "Unknown classification", type text)
in
    Readers[Finish](Service)'''


def classified_scope():
    """Enrich a DAX _scopeRaw without inferring identity from names or SMTP."""
    return '''VAR _classification=SELECTCOLUMNS(ALL('AD Account Classification'),
    "ClassObject",'AD Account Classification'[Object GUID],
    "ClassBatch",'AD Account Classification'[Loaded Batch ID],
    "ClassState",'AD Account Classification'[Service Classification])
VAR _scope=GENERATE(_scopeRaw,
    VAR _object=[Object] VAR _batch=[Batch]
    VAR _matches=FILTER(_classification,_object<>"" && _batch<>"" && [ClassObject]=_object && [ClassBatch]=_batch)
    VAR _unique=COUNTROWS(_matches)=1
    VAR _state=IF(_unique,MAXX(_matches,[ClassState]))
    RETURN ROW("Service Retained",IF(_unique && _state="Service",1,0),
        "Classification Known",IF(_unique && _state IN {"Service","Other classified account"},1,0)))'''


def prepare_activation_model(model):
    """Return the complete future model; no source processing or deployment.

    Keep separate from the working cached PBIP until its new imported table can
    be materialized in an authorized refresh. The current view remains usable.
    """
    import copy
    from exchange_license_scenario import scenario_measures
    from exchange_mailbox_populations import population_measures
    result=copy.deepcopy(model);body=result['model']
    assert not any(t['name']==TABLE for t in body['tables']),'Classification already activated'
    columns=['Object GUID','AD Account Key','Account Type','Service Classification','Loaded Batch ID']
    body['tables'].append(dict(name=TABLE,description='Utility facet: every governed AD account type, including disabled and non-human accounts. Native same-batch lookup only; no relationship or uniqueness semantics changed in existing tables.',
        columns=[dict(name=name,dataType='string',sourceColumn=name,summarizeBy='none',isHidden=name in {'Object GUID','AD Account Key','Loaded Batch ID'}) for name in columns],
        partitions=[dict(name=TABLE,mode='import',source=dict(type='m',expression=EXPRESSION+'()'))]))
    assert any(e['name']==EXPRESSION for e in body['expressions'])
    target=next(t for t in body['tables'] if t['name']=='Exchange Mailbox Reconciliation')
    definitions={m['name']:m for m in scenario_measures(True)+population_measures(True)}
    for i,m in enumerate(target['measures']):
        if m['name'] in definitions:target['measures'][i]={**m,**definitions[m['name']]}
    audit=next(t for t in body['tables'] if t['name']=='Loaded Batch Audit')['partitions'][0]['source']
    expr=audit['expression'];expr='\n'.join(expr) if isinstance(expr,list) else expr
    assert expr.rstrip().endswith(')')
    audit['expression']=expr.rstrip()[:-1]+',\nROW("Table Name", "'+TABLE+'", "Loaded Batch ID", MAXX(\''+TABLE+'\', \''+TABLE+'\'[Loaded Batch ID]), "Loaded Rows", COUNTROWS(\''+TABLE+'\'))\n)'
    return result
