"""Generic offline TOM definitions for the optional migration reporting model.

Applied through Modeling MCP; this helper neither refreshes nor edits PBIP files.
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONTRACT = json.loads((ROOT / 'config/sharepoint-migration-contract.json.txt').read_text())
P = "'SPO Migration Plan'"
H = "'SPO Migration History'"
W = "'SPO Migration Web'"
C = "'SPO Migration Collection'"
S = "'SPO Migration Scope'"
F = "'SPO Migration Findings'"
Q = "'SPO Migration Source Summary'"


def definitions():
    files = [t['file'] for t in CONTRACT['tables']]
    reader = '''(FileName as text, Schema as list) as table =>
let
    Expected = FILES,
    Manifest = GetPreparedBatchManifest(),
    Entries = List.Select(Manifest[Files], each List.Contains(Expected, [File])),
    Names = List.Transform(Entries, each [File]),
    GroupReady = if List.IsEmpty(Entries) then false else
        if List.Count(Names) = List.Count(Expected) and List.Count(List.Distinct(Names)) = List.Count(Expected)
        then true else error "Incomplete migration evidence group in selected batch",
    Allowed = if List.Contains(Expected, FileName) then FileName else error "Unexpected migration table",
    Empty = #table(List.Transform(Schema, each _{0}), {}),
    Parsed = if GroupReady then Table.PromoteHeaders(Csv.Document(GetPreparedEvidenceFile(Allowed),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true]) else Empty,
    HeaderChecked = if Table.ColumnNames(Parsed) = List.Transform(Schema, each _{0}) then Parsed
        else error "Migration evidence schema mismatch",
    TenantChecked = if List.AnyTrue(List.Transform(Table.Column(HeaderChecked, "TenantKey"),
        each _ <> PreparedExpectedTenantKey)) then error "Foreign migration tenant" else HeaderChecked,
    Nulls = Table.ReplaceValue(TenantChecked, "", null, Replacer.ReplaceValue, Table.ColumnNames(TenantChecked)),
    Conversions = List.Transform(Schema, (s) => {s{0}, (v) => if v = null then null else
        if s{1} = type date then Date.FromText(v, [Culture="en-US"])
        else if s{1} = type datetime then DateTimeZone.RemoveZone(DateTimeZone.ToUtc(DateTimeZone.FromText(v)))
        else if s{1} = type logical then Logical.FromText(v)
        else if s{1} = Int64.Type then Int64.From(v, "en-US")
        else if s{1} = Currency.Type then Currency.From(v, "en-US") else v, s{1}}),
    Typed = Table.TransformColumns(Nulls, Conversions),
    Compatible = if not GroupReady and Allowed = "SharePointMigrationSourceSummary.csv" then
        Table.FromRecords({[TenantKey=PreparedExpectedTenantKey, Source="Migration workbook",
            Qualification="Not available in selected batch", #"Coverage State"="Not available in selected batch"]},
            Value.Type(Typed), MissingField.UseNull) else Typed
in Compatible'''.replace('FILES', json.dumps(files).replace('[', '{').replace(']', '}'))
    expressions = [{'name': 'GetSharePointMigrationEvidence', 'kind': 'M', 'expression': reader,
                    'description': 'Optional complete nine-file group from the selected immutable batch. Old batches yield typed empty compatibility; corrupt or incomplete groups fail.'}]
    tables = []
    type_map = {'string': ('String', 'type text'), 'int64': ('Int64', 'Int64.Type'),
                'decimal': ('Decimal', 'Currency.Type'), 'boolean': ('Boolean', 'type logical')}
    for t in CONTRACT['tables']:
        cols, schema = [], []
        for c in t['columns']:
            name = c['name']
            dtype, mtype = ('DateTime', 'type date' if name.endswith('Date') else 'type datetime') if c['type'] == 'dateTime' else type_map[c['type']]
            hidden = (name.endswith(' Key') or name in ('TenantKey', 'Snapshot ID', 'Source Row', 'Native Web Count', 'Legacy Web ID',
                      'Source SHA256', 'Binding SHA256', 'Capture UTC', 'Captured UTC', 'Run ID', 'Source URL Aliases', 'Storage Bytes',
                      'Reported Size', 'Project Scope Key', 'Present In Register', 'Is Validated', 'Requires Review')) or name.startswith('Imported ')
            col = {'Name': name, 'DataType': dtype, 'SourceColumn': name, 'IsHidden': hidden,
                   'SummarizeBy': 'None', 'Description': 'Prepared, tenant-qualified migration evidence: ' + name + '.'}
            if dtype == 'DateTime':
                col['FormatString'] = 'yyyy-MM-dd' if mtype == 'type date' else 'yyyy-MM-dd HH:mm'
            if name == 'Resolved Target URL':
                col['DataCategory'] = 'WebUrl'
            cols.append(col)
            schema.append('{"' + name + '", ' + mtype + '}')
        query = 'let\n    Source = GetSharePointMigrationEvidence("' + t['file'] + '", {' + ', '.join(schema) + '}),\n'
        query += '''    Identity = Table.AddColumn(Source, "Project Identity", each Binary.ToText(Json.FromValue({[TenantKey], [Project Key]}), BinaryEncoding.Base64), type text),
    Categories = Table.TransformColumns(Identity, List.Transform(List.Intersect({{"Business Unit", "Phase"}, Table.ColumnNames(Identity)}), (c) => {c, each if _ = null then "(Unspecified)" else _, type text}))'''
        cols.append({'Name': 'Project Identity', 'DataType': 'String', 'IsHidden': True, 'SourceColumn': 'Project Identity'})
        if t['projection'] == 'History':
            query += ''',
    Snapshots = Table.Group(Categories, {"Project Identity", "Snapshot Date", "Snapshot ID"}, {{"Observed", each List.Max([Capture UTC]), type nullable datetime}}),
    Daily = Table.Group(Snapshots, {"Project Identity", "Snapshot Date"}, {{"Latest", each Table.First(Table.Sort(_, {{"Observed", Order.Descending}, {"Snapshot ID", Order.Descending}}))[Snapshot ID], type text}}),
    Joined = Table.NestedJoin(Categories, {"Project Identity", "Snapshot Date", "Snapshot ID"}, Daily, {"Project Identity", "Snapshot Date", "Latest"}, "Daily Match", JoinKind.LeftOuter),
    Latest = Table.AddColumn(Joined, "Is Latest Daily Snapshot", each not Table.IsEmpty([Daily Match]), type logical),
    Result = Table.RemoveColumns(Latest, {"Daily Match"})
in Result'''
            cols.append({'Name': 'Is Latest Daily Snapshot', 'DataType': 'Boolean', 'IsHidden': True})
        else:
            if t['projection'] == 'Plan':
                query += ''',
    Variance = Table.AddColumn(Categories, "Schedule Variance Days", each if [Disposition]="MIGRATE" and [Execution Status]="DONE" and [Planned Date]<>null and [Actual Date]<>null then Duration.Days([Actual Date]-[Planned Date]) else null, Int64.Type)
in Variance'''
                cols.append({'Name': 'Schedule Variance Days', 'DataType': 'Int64', 'IsHidden': True, 'SummarizeBy': 'None'})
            else:
                query += '\nin Categories'
        tables.append({'Name': t['table'], 'Mode': 'Import', 'MExpression': query, 'Columns': cols,
                       'IsHidden': t['projection'].startswith('Legacy') or t['projection'] == 'Scope',
                       'Description': 'SPO means SharePoint Online. Grain: ' + ', '.join(t['uniqueKey']) + '; source: one selected immutable prepared batch.'})

    def dimension(name, columns, query, hidden=False):
        tables.append({'Name': name, 'Mode': 'Import', 'MExpression': query, 'IsHidden': hidden,
                       'Description': 'Migration reporting dimension; current and historical members retained.',
                       'Columns': [{'Name': c, 'DataType': 'DateTime' if c == 'Date' else 'String', 'IsHidden': c.endswith('Identity') or c == 'Operation Key',
                                    'SourceColumn': c, 'SummarizeBy': 'None', **({'FormatString': 'yyyy-MM-dd'} if c == 'Date' else {})} for c in columns]})
    for col, suffix in [('Source Type', 'Source Type'), ('Business Unit', 'Business Unit'), ('Phase', 'Phase')]:
        dimension('SPO Migration ' + suffix, [col], 'let Source = Table.Combine({Table.SelectColumns(#"SPO Migration Plan", {"' + col + '"}), Table.SelectColumns(#"SPO Migration History", {"' + col + '"})}) in Table.Distinct(Source)')
    dimension('SPO Migration Project', ['Project Identity', 'Project Key'], '''let
        Source = Table.Combine(List.Transform({#"SPO Migration Plan", #"SPO Migration History", #"SPO Migration Source Summary"}, each Table.SelectColumns(_, {"Project Identity", "Project Key"})))
        in Table.Distinct(Source)''')
    dimension('SPO Migration Operation', ['Operation Key', 'Migration ID', 'Source Type'], '''let
        Source = Table.Combine(List.Transform({#"SPO Migration Plan", #"SPO Migration History"}, each Table.SelectColumns(_, {"Operation Key", "Migration ID", "Source Type"})))
        in Table.Distinct(Source, {"Operation Key"})''', True)
    dimension('SPO Migration Snapshot Date', ['Date'], '''let
        Native = Table.SelectColumns(Date, {"Date"}),
        Historical = Table.RenameColumns(Table.SelectColumns(#"SPO Migration History", {"Snapshot Date"}), {{"Snapshot Date", "Date"}}),
        Bounds = List.RemoveNulls(Table.Column(Table.Combine({Native, Historical}), "Date")),
        Start = List.Min(Bounds), End = List.Max(Bounds),
        Result = Table.FromColumns({List.Dates(Start, Duration.Days(End-Start)+1, #duration(1,0,0,0))}, type table [Date=date])
        in Result''')
    relationships = []

    def rel(f, fc, t, tc, active=True):
        relationships.append({'name': 'SpoMigration' + str(len(relationships) + 1), 'fromTable': f, 'fromColumn': fc,
                              'toTable': t, 'toColumn': tc, 'fromCardinality': 'Many', 'toCardinality': 'One',
                              'crossFilteringBehavior': 'OneDirection', 'isActive': active})
    for fact in ['SPO Migration Plan', 'SPO Migration History']:
        for col in ['Source Type', 'Business Unit', 'Phase']:
            rel(fact, col, 'SPO Migration ' + col, col)
        rel(fact, 'Operation Key', 'SPO Migration Operation', 'Operation Key')
    for fact in ['SPO Migration Plan', 'SPO Migration History', 'SPO Migration Source Summary', 'SPO Migration Findings', 'SPO Migration Legacy Web', 'SPO Migration Legacy Collection']:
        rel(fact, 'Project Identity', 'SPO Migration Project', 'Project Identity')
    rel('SPO Migration Scope', 'Operation Key', 'SPO Migration Plan', 'Operation Key')
    rel('SPO Migration Scope', 'Web Key', 'SPO Migration Web', 'Web Key')
    rel('SPO Migration Web', 'Collection Key', 'SPO Migration Collection', 'Collection Key')
    rel('SPO Migration Web', 'Web Application Key', 'SharePoint On-prem Web Applications', 'Web Application Key')
    rel('SPO Migration History', 'Snapshot Date', 'SPO Migration Snapshot Date', 'Date')
    for col in ['Planned Date', 'Actual Date', 'Validation Date']:
        rel('SPO Migration Plan', col, 'Date', 'Date', False)

    measures = []

    def measure(name, expression, folder='Progress', fmt='#,##0', table='SPO Migration Plan', description=None, hidden=False):
        measures.append({'name': 'SPO ' + name, 'tableName': table, 'expression': expression,
                         'formatString': fmt, 'displayFolder': 'Migration\\' + folder, 'isHidden': hidden,
                         'description': description or name + '. Operations use distinct source-type-qualified identities. Declared progress is separate from technical observation and validation.'})

    technical = f"ISCROSSFILTERED('SharePoint On-prem Web Applications') || ISFILTERED({W}[Collection Key]) || ISFILTERED({C}[Source URL]) || ISFILTERED({W}[Source URL]) || ISFILTERED({W}[Coverage Kind])"
    measure('Operations', f'''VAR _technical = {technical}
VAR _scope = VALUES({S}[Operation Key])
VAR _sp = CALCULATE(DISTINCTCOUNT({P}[Operation Key]), KEEPFILTERS({P}[Source Type]="SharePoint on-prem"), KEEPFILTERS(TREATAS(_scope, {P}[Operation Key])))
VAR _shared = CALCULATE(DISTINCTCOUNT({P}[Operation Key]), KEEPFILTERS({P}[Source Type]="Shared folders"))
RETURN IF(_technical, IF(NOT ISBLANK(_sp) || NOT ISBLANK(_shared), COALESCE(_sp,0)+COALESCE(_shared,0)), DISTINCTCOUNT({P}[Operation Key]))''')
    measure('Migrate Operations', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Disposition]="MIGRATE"))')
    measure('Migrate Population', f'CALCULATE([SPO Migrate Operations], REMOVEFILTERS({P}[Execution Status]), REMOVEFILTERS({P}[Validation Status]), REMOVEFILTERS({P}[Is Validated]))', description='Shared denominator for reported and validated progress: every MIGRATE operation in the selected source/project/business unit/phase and technical context, including cancellations. Execution and validation selections do not shrink the denominator.')
    measure('Reported Done', f'VAR _n=[SPO Migrate Population] RETURN IF(_n>0, COALESCE(CALCULATE([SPO Migrate Operations], KEEPFILTERS({P}[Execution Status]="DONE")),0))')
    measure('Validated Done', f'VAR _n=[SPO Migrate Population] RETURN IF(_n>0, COALESCE(CALCULATE([SPO Reported Done], KEEPFILTERS({P}[Is Validated]=TRUE())),0))')
    measure('Reported Completion', 'DIVIDE([SPO Reported Done], [SPO Migrate Population])', fmt='0.0%')
    measure('Validated Completion', 'DIVIDE([SPO Validated Done], [SPO Migrate Population])', fmt='0.0%')
    for label, expression in [('Reported Progress', '[SPO Reported Done]'), ('Validated Progress', '[SPO Validated Done]')]:
        measure(label, 'VAR _n=[SPO Migrate Population] RETURN IF(_n>0, FORMAT(' + expression + ',"#,##0") & " / " & FORMAT(_n,"#,##0") & " (" & FORMAT(DIVIDE(' + expression + ',_n),"0.0%") & ")", "n/a")', fmt='General')
    measure('Remaining to Execute', f'VAR _n=[SPO Migrate Operations] RETURN IF(_n>0, COALESCE(CALCULATE([SPO Migrate Operations], KEEPFILTERS({P}[Execution Status]<>"DONE"), KEEPFILTERS({P}[Execution Status]<>"CANCELLED")),0))')
    measure('Remaining to Close', 'IF([SPO Migrate Population]>0, [SPO Migrate Population]-[SPO Validated Done])')
    measure('Cancelled Migrations', f'CALCULATE([SPO Migrate Operations], KEEPFILTERS({P}[Execution Status]="CANCELLED"))')
    measure('Reference Date', f'CALCULATE(MAX({Q}[Snapshot Date]), REMOVEFILTERS({Q}[Source]))', fmt='yyyy-MM-dd')
    measure('Page Context', f'''VAR _date=[SPO Reference Date]
VAR _states=CONCATENATEX(VALUES({Q}[Qualification]), {Q}[Qualification], ", ")
RETURN "Batch " & IF(ISBLANK(_date),"not available",FORMAT(_date,"yyyy-MM-dd")) & " | " & COALESCE(_states,"No migration evidence") & " | " &
IF({technical}, "On-prem filter; shared folders not applicable", "Declared progress; validation is separate")''', fmt='General')
    for decision in ['ARCHIVE', 'DELETE', 'KEEP', 'TBD', 'NOTINPLAN']:
        measure(decision.title() + ' Operations', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Disposition]="{decision}"))', 'Decisions')
    measure('Other Actions Done', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Disposition]<>"MIGRATE"), KEEPFILTERS({P}[Execution Status]="DONE"))', 'Decisions')
    active = f'KEEPFILTERS({P}[Execution Status]<>"DONE"), KEEPFILTERS({P}[Execution Status]<>"CANCELLED")'
    for label, condition in [('Overdue', f'{P}[Planned Date]<_ref'), ('Due Today', f'{P}[Planned Date]=_ref'),
                             ('Due Next 7 Days', f'{P}[Planned Date]>_ref && {P}[Planned Date]<=_ref+7'),
                             ('Due Days 8 to 14', f'{P}[Planned Date]>_ref+7 && {P}[Planned Date]<=_ref+14'),
                             ('Due Within 14 Days', f'{P}[Planned Date]>=_ref && {P}[Planned Date]<=_ref+14')]:
        measure(label, f'VAR _ref=[SPO Reference Date] RETURN IF(NOT ISBLANK(_ref) && [SPO Migrate Operations]>0, COALESCE(CALCULATE([SPO Migrate Operations], {active}, KEEPFILTERS(NOT ISBLANK({P}[Planned Date])), KEEPFILTERS({condition})),0))', 'Schedule')
    measure('Missing Planned Date', f'CALCULATE([SPO Migrate Operations], {active}, KEEPFILTERS(ISBLANK({P}[Planned Date])))', 'Schedule')
    measure('Done Missing Actual Date', f'CALCULATE([SPO Reported Done], KEEPFILTERS(ISBLANK({P}[Actual Date])))', 'Schedule')
    measure('Planned Calendar Operations', f'CALCULATE([SPO Migrate Operations], USERELATIONSHIP({P}[Planned Date],\'Date\'[Date]))', 'Schedule')
    measure('Actual Calendar Completions', f'CALCULATE([SPO Reported Done], USERELATIONSHIP({P}[Actual Date],\'Date\'[Date]))', 'Schedule')
    measure('Validation Calendar Completions', f'CALCULATE([SPO Validated Done], USERELATIONSHIP({P}[Validation Date],\'Date\'[Date]))', 'Schedule')
    measure('Cumulative Declared Completions', '''VAR _day=MAX('Date'[Date]) VAR _reference=[SPO Reference Date]
RETURN IF(_day<=_reference, CALCULATE([SPO Actual Calendar Completions], FILTER(ALLSELECTED('Date'), 'Date'[Date]<=_day)))''', 'Schedule')
    eligible = f'FILTER({P}, {P}[Disposition]="MIGRATE" && {P}[Execution Status]="DONE" && NOT ISBLANK({P}[Planned Date]) && NOT ISBLANK({P}[Actual Date]) && [SPO Operations]>0)'
    measure('Variance Eligible Operations', f'COUNTROWS({eligible})', 'Schedule')
    measure('Mean Schedule Variance Days', f'AVERAGEX({eligible}, DATEDIFF({P}[Planned Date], {P}[Actual Date], DAY))', 'Schedule', '0.0')
    measure('Variance Excluded Operations', 'IF([SPO Reported Done]>0, [SPO Reported Done]-COALESCE([SPO Variance Eligible Operations],0))', 'Schedule')
    measure('Historical Migrate Operations', f'''VAR _day=CALCULATE(MAX({H}[Snapshot Date]), REMOVEFILTERS('SPO Migration Business Unit'), REMOVEFILTERS('SPO Migration Phase'), REMOVEFILTERS('SPO Migration Source Type'), REMOVEFILTERS('SPO Migration Operation'))
RETURN CALCULATE(DISTINCTCOUNT({H}[Operation Key]), KEEPFILTERS({H}[Snapshot Date]=_day), KEEPFILTERS({H}[Is Latest Daily Snapshot]=TRUE()), KEEPFILTERS({H}[Present In Register]=TRUE()), KEEPFILTERS({H}[Disposition]="MIGRATE"))''', 'History', table='SPO Migration History')
    measure('Historical Reported Done', f'VAR _n=[SPO Historical Migrate Operations] RETURN IF(_n>0, COALESCE(CALCULATE([SPO Historical Migrate Operations], KEEPFILTERS({H}[Execution Status]="DONE")),0))', 'History', table='SPO Migration History')
    measure('Historical Validated Done', f'VAR _n=[SPO Historical Migrate Operations] RETURN IF(_n>0, COALESCE(CALCULATE([SPO Historical Reported Done], KEEPFILTERS({H}[Is Validated]=TRUE())),0))', 'History', table='SPO Migration History')
    measure('Historical Completion', 'DIVIDE([SPO Historical Reported Done], [SPO Historical Migrate Operations])', 'History', '0.0%', table='SPO Migration History')
    measure('Historical Removed Operations', f'CALCULATE(DISTINCTCOUNT({H}[Operation Key]), KEEPFILTERS({H}[Is Latest Daily Snapshot]=TRUE()), KEEPFILTERS({H}[Present In Register]=FALSE()))', 'History', table='SPO Migration History')
    measure('Source Identity Review', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Source Type]="SharePoint on-prem"), KEEPFILTERS({P}[Source Qualification]<>"Direct native match"))', 'Reconciliation')
    measure('Target Review', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Disposition] IN {{"MIGRATE","ARCHIVE"}}), KEEPFILTERS({P}[Target Approval Evidence]="Not recorded"))', 'Reconciliation')
    measure('Required Targets Missing', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Disposition] IN {{"MIGRATE","ARCHIVE"}}), KEEPFILTERS(ISBLANK({P}[Resolved Target URL])))', 'Reconciliation')
    for label, value in [('Targets Observed', 'Site observed'), ('Targets Not Observed', 'Not observed'), ('Targets Not Verifiable', 'Not verifiable')]:
        measure(label, f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Disposition] IN {{"MIGRATE","ARCHIVE"}}), KEEPFILTERS({P}[Target Observation]="{value}"))', 'Reconciliation')
    measure('Observed Webs', f'''VAR _sharedOnly=SELECTEDVALUE('SPO Migration Source Type'[Source Type])="Shared folders"
VAR _business=ISFILTERED('SPO Migration Business Unit') || ISFILTERED('SPO Migration Phase')
VAR _projects=VALUES('SPO Migration Project'[Project Identity])
RETURN IF(NOT _sharedOnly, IF(_business, CALCULATE(DISTINCTCOUNT({W}[Web Key]), TREATAS(_projects,{W}[Project Identity]), TREATAS(VALUES({S}[Web Key]), {W}[Web Key])), CALCULATE(DISTINCTCOUNT({W}[Web Key]), TREATAS(_projects,{W}[Project Identity]))))''', 'Reconciliation', table='SPO Migration Web')
    for label, value in [('Webs Without Plan', 'NOTINPLAN'), ('Whole Webs Covered', 'Whole web coverage'), ('Partial Webs', 'Partial content only'), ('Conflicting Webs', 'Conflict')]:
        extra = ", REMOVEFILTERS('SPO Migration Business Unit'), REMOVEFILTERS('SPO Migration Phase')" if label == 'Webs Without Plan' else ''
        measure(label, f'CALCULATE([SPO Observed Webs], KEEPFILTERS({W}[Coverage Kind]="{value}"){extra})', 'Reconciliation', table='SPO Migration Web', description=label + '. Observed webs without a plan have no attributable business unit or phase, so that global counter preserves the technical context and source flow but ignores those two business filters.')
    measure('On-prem Coverage Context', 'IF(SELECTEDVALUE(\'SPO Migration Source Type\'[Source Type])="Shared folders", "Not applicable: no SMB inventory, volume or ACL evidence", "Observed webs; partial inventory cannot prove absence or deletion")', 'Reconciliation', 'General', table='SPO Migration Web')
    measure('Current Collection Storage GiB', f'''VAR _keys=VALUES({S}[Collection Key])
RETURN IF(SELECTEDVALUE('SPO Migration Source Type'[Source Type])<>"Shared folders",
DIVIDE(CALCULATE(SUM({C}[Storage Bytes]), TREATAS(_keys, {C}[Collection Key])),POWER(1024,3)))''', 'Volumes', '0.0', table='SPO Migration Collection')
    measure('Fully Qualified Migrating GiB', f'''VAR _collections=VALUES({S}[Collection Key])
VAR _eligible=CALCULATE(COUNTROWS({C}), KEEPFILTERS({C}[Is Fully Migrating]=TRUE()), TREATAS(_collections, {C}[Collection Key]))
RETURN IF(_eligible>0 && SELECTEDVALUE('SPO Migration Source Type'[Source Type])<>"Shared folders",
DIVIDE(CALCULATE(SUM({C}[Storage Bytes]), KEEPFILTERS({C}[Is Fully Migrating]=TRUE()),TREATAS(_collections,{C}[Collection Key])),POWER(1024,3)))''', 'Volumes', '0.0', table='SPO Migration Collection')
    measure('Volume Context', 'IF(ISBLANK([SPO Fully Qualified Migrating GiB]),"Not measurable: complete migration scope is not qualified", "Current source storage in fully qualified scope; not transferred bytes")', 'Volumes', 'General', table='SPO Migration Collection')
    measure('Operations With Findings', f'CALCULATE([SPO Operations], KEEPFILTERS({P}[Requires Review]=TRUE()))', 'Quality')
    measure('Business Review Findings', f'''VAR _keys=FILTER(VALUES({P}[Operation Key]), [SPO Operations]>0) RETURN CALCULATE(DISTINCTCOUNT({F}[Operation Key]), TREATAS(_keys, {F}[Operation Key]),
KEEPFILTERS({F}[Finding Rule] IN {{"Decision pending","Missing planned date","Done without actual date","Imported date needs clarification","Target approval pending","Migration validation not documented","Inconsistent validation"}}))''', 'Quality', table='SPO Migration Findings')
    measure('Finding Count', f'VAR _keys=FILTER(VALUES({P}[Operation Key]), [SPO Operations]>0) RETURN CALCULATE(COUNTROWS({F}), TREATAS(_keys,{F}[Operation Key]))', 'Quality', table='SPO Migration Findings')
    measure('Global Findings', f'CALCULATE(COUNTROWS({F}), KEEPFILTERS(ISBLANK({F}[Operation Key])))', 'Quality', table='SPO Migration Findings')
    measure('Global Quality Context', f'''VAR _available=COUNTROWS(FILTER({Q}, {Q}[Qualification]="Configured"))>0
RETURN IF(_available,"Source/global findings: " & FORMAT(COALESCE([SPO Global Findings],0),"#,##0") & ". Source-level findings are not attributable to a business unit or operation. Rule categories overlap.", [SPO Source Qualification Context])''', 'Quality', 'General', table='SPO Migration Findings')
    measure('Visible Finding', f'''VAR _key=SELECTEDVALUE({F}[Operation Key]) VAR _keys=FILTER(VALUES({P}[Operation Key]), [SPO Operations]>0)
RETURN IF(ISBLANK(_key) || CONTAINS(_keys, {P}[Operation Key], _key), 1)''', 'Quality', table='SPO Migration Findings', hidden=True)
    measure('Source Qualification Context', f'CONCATENATEX(VALUES({Q}[Qualification]), {Q}[Qualification], ", ")', 'Quality', 'General', table='SPO Migration Source Summary')
    measure('Source Age Hours', f'MAX({Q}[Age Hours])', 'Quality', '0.0', table='SPO Migration Source Summary')
    measure('Permission Evidence Context', '"Permission strategy is declared; effective ACL equivalence is not measured"', 'Quality', 'General')
    return {'expressions': expressions, 'tables': tables, 'relationships': relationships, 'measures': measures,
            'existingColumns': [{'tableName': 'Date', 'name': 'Week Start', 'expression': "'Date'[Date]-WEEKDAY('Date'[Date],2)+1", 'dataType': 'DateTime', 'formatString': 'yyyy-MM-dd', 'summarizeBy': 'None', 'displayFolder': 'Migration calendar', 'description': 'Monday of the reporting week; used for declared migration deadlines.'}]}


if __name__ == '__main__':
    print(json.dumps(definitions()))
