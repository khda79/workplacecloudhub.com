"""Offline, tenant-scoped migration planning evidence. Never writes its inputs.

Workbook rows are declarations; inventory presence and migration validation are
separate evidence. Uses only the Python standard library on preparation hosts.
"""
import argparse
import collections
import csv
import datetime as dt
import hashlib
import json
import posixpath
import re
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET

VERSION = '1.0.3'
DISABLED_PROJECT_KEY = 'Not configured'
CONTRACT = Path(__file__).resolve().parents[1] / 'config/sharepoint-migration-contract.json.txt'
UTC = dt.timezone.utc
TABLES = {'MigrationSP': 'tblMigrationSP', 'MigrationSharedFolders': 'tblMigrationSharedFolders',
          'LegacyWebs': 'tblLegacyWebs', 'LegacyCollections': 'tblLegacyCollections'}
FLOWS = {'MigrationSP': 'SharePoint on-prem', 'MigrationSharedFolders': 'Shared folders'}
FLOW_CODES = {'SharePoint on-prem': 'SHAREPOINT-ONPREM', 'Shared folders': 'SHAREDFOLDERS'}
DERIVED = {'TenantKey', 'ProjectKey', 'SourceType', 'DateIssue', 'SourceIssue', 'TargetIssue',
           'ValidationIssue', 'TenantIssue', 'LegacyLinkIssue', 'ScopeIssue', 'ReviewRequired', 'SizeConverted', 'ConvertedUnit'}
DISPOSITIONS = {'MIGRATE', 'ARCHIVE', 'DELETE', 'KEEP', 'TBD', 'NOTINPLAN'}
EXECUTIONS = {'TO_PLAN', 'PLANNED', 'IN_PROGRESS', 'DONE', 'BLOCKED', 'FAILED', 'CANCELLED', 'NOTINPLAN'}
CONTENT_OBJECT_TYPES = {'FOLDER', 'LIBRARY', 'LIST', 'FOLDER_PROBABLE', 'LIBRARY_PROBABLE'}


def text(value):
    return str(value or '').strip()


def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1048576), b''):
            h.update(chunk)
    return h.hexdigest().upper()


def token(*parts):
    return hashlib.sha256(json.dumps(parts, ensure_ascii=True, separators=(',', ':')).encode()).hexdigest()


def utc(value):
    result = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
    if result.tzinfo is None:
        raise ValueError('Capture timestamp requires an explicit UTC offset')
    return result.astimezone(UTC)


def date(value):
    if not text(value):
        return ''
    if re.fullmatch(r'\d+(?:\.0+)?', text(value)):
        serial = float(value)
        if not 61 <= serial < 2958466 or serial != int(serial):
            raise ValueError('Invalid Excel date serial')
        return (dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))).isoformat()
    return dt.date.fromisoformat(text(value)[:10]).isoformat()


def url_key(raw):
    """Business identity ignores HTTP/HTTPS, never significant resource parts."""
    parsed = urlsplit(text(raw))
    if parsed.scheme.lower() not in ('http', 'https') or not parsed.hostname or parsed.username or parsed.password or parsed.fragment:
        raise ValueError('Expected an absolute HTTP(S) source URL without credentials or fragment')
    host = parsed.hostname.casefold()
    port = parsed.port
    if port and (parsed.scheme.casefold(), port) not in [('http', 80), ('https', 443)]:
        host += ':' + str(port)
    # Preserve encoded separators, which can change the resource boundary.
    path = unquote(re.sub('%2f', '%252F', re.sub('%5c', '%255C', parsed.path, flags=re.I), flags=re.I)).casefold().rstrip('/')
    if any(segment in ('.', '..') for segment in path.split('/')) or '\\' in path:
        raise ValueError('Non-canonical source URL path')
    return host, path, parsed.query


def ancestor(parent, child):
    return parent[0] == child[0] and parent[2] == child[2] and (child[1] == parent[1] or child[1].startswith(parent[1] + '/'))


def read_workbook(path):
    """Read actual table bounds; ignore formula caches in derived columns."""
    with zipfile.ZipFile(path) as archive:
        if any(n.endswith('vbaProject.bin') or n.startswith('xl/externalLinks/') for n in archive.namelist()):
            raise ValueError('Macros and external workbook links are unsupported')
        if sum(i.file_size for i in archive.infolist()) > 256 * 1024 * 1024:
            raise ValueError('Workbook exceeds the uncompressed size limit')

        def xml(name):
            if archive.getinfo(name).file_size > 48 * 1024 * 1024:
                raise ValueError('Oversized workbook part')
            raw = archive.read(name)
            if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():
                raise ValueError('Unsafe workbook XML')
            return ET.fromstring(raw)

        def target(base, relative):
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(base), relative)) if not relative.startswith('/') else relative.lstrip('/')
            if not resolved.startswith('xl/') or '..' in resolved.split('/'):
                raise ValueError('Unsafe workbook relationship')
            return resolved

        strings = [''.join(e.itertext()) for e in xml('xl/sharedStrings.xml')] if 'xl/sharedStrings.xml' in archive.namelist() else []
        workbook = xml('xl/workbook.xml')
        props = workbook.find('.//{*}workbookPr')
        if props is not None and props.get('date1904') in ('1', 'true'):
            raise ValueError('Expected the Excel 1900 date system')
        rels = {r.get('Id'): r.get('Target') for r in xml('xl/_rels/workbook.xml.rels') if r.get('TargetMode') != 'External'}
        sheets = {}
        for sheet in workbook.findall('.//{*}sheet'):
            rid = next(v for k, v in sheet.attrib.items() if k.endswith('}id'))
            sheets[sheet.get('name')] = target('xl/workbook.xml', rels[rid])
        if not set(TABLES).union({'Parameters'}).issubset(sheets):
            raise ValueError('Required migration worksheets are missing')

        def cell_value(cell):
            kind = cell.get('t')
            if kind == 'e':
                raise ValueError('Workbook contains an error cell in an input table')
            raw = cell.findtext('{*}v', '')
            return strings[int(raw)] if kind == 's' else ''.join(cell.find('{*}is').itertext()) if kind == 'inlineStr' else raw

        parameters = {}
        for row in xml(sheets['Parameters']).findall('.//{*}sheetData/{*}row'):
            cells = {re.sub(r'\d', '', c.get('r')): cell_value(c) for c in row}
            if cells.get('A') in ('TenantKey', 'ProjectKey', 'SchemaVersion', 'ReviewDate', 'LegacySnapshotDate'):
                if any(c.find('{*}f') is not None for c in row if c.get('r', '').startswith('B')):
                    raise ValueError('Identity and review parameters require values')
                parameters[cells['A']] = cells.get('B', '')
        if parameters.get('SchemaVersion') != '6.3':
            raise ValueError('Unsupported migration workbook schema version')
        result = {}
        for name, table_name in TABLES.items():
            sheet_path = sheets[name]
            sheet = xml(sheet_path)
            links = {r.get('Id'): r.get('Target') for r in xml(posixpath.dirname(sheet_path) + '/_rels/' + posixpath.basename(sheet_path) + '.rels') if r.get('TargetMode') != 'External'}
            tables = []
            for part in sheet.findall('.//{*}tablePart'):
                rid = next(v for k, v in part.attrib.items() if k.endswith('}id'))
                table = xml(target(sheet_path, links[rid]))
                if table.get('displayName') == table_name:
                    tables.append(table)
            if len(tables) != 1:
                raise ValueError('Expected exactly one structured migration input table: ' + table_name)
            table = tables[0]
            headers = [c.get('name') for c in table.findall('.//{*}tableColumn')]
            match = re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', table.get('ref', ''))
            if not match or table.get('totalsRowCount', '0') != '0' or len(headers) != len(set(headers)):
                raise ValueError('Invalid migration table bounds or headers')
            def col(raw):
                out = 0
                for ch in raw:
                    out = out * 26 + ord(ch) - 64
                return out
            left, top, right, bottom = col(match[1]), int(match[2]), col(match[3]), int(match[4])
            if right - left + 1 != len(headers):
                raise ValueError('Table metadata width mismatch')
            required = {'MigrationID', 'Disposition', 'ExecutionStatus', 'ValidationStatus', 'ValidationDate', 'TargetURL', 'PlannedDate', 'ActualDate'} if name in FLOWS else {'LegacyWebID'} if name == 'LegacyWebs' else {'LegacyCollectionID'}
            if not required.issubset(headers):
                raise ValueError('Required migration input columns missing: ' + name)
            rows, seen_header = [], False
            for row in sheet.findall('.//{*}sheetData/{*}row'):
                number = int(row.get('r'))
                if not top <= number <= bottom:
                    continue
                values = {}
                for cell in row:
                    column = col(re.sub(r'\d', '', cell.get('r')))
                    if not left <= column <= right:
                        continue
                    header = headers[column - left]
                    formula = cell.find('{*}f') is not None
                    if formula and (number == top or header not in DERIVED):
                        raise ValueError('Unexpected formula in migration input: ' + header)
                    values[header] = '' if formula else cell_value(cell)
                if number == top:
                    seen_header = True
                    if [values.get(h, '') for h in headers] != headers:
                        raise ValueError('Worksheet headers differ from structured table metadata')
                elif any(text(v) for k, v in values.items() if k not in DERIVED):
                    rows.append({**{h: values.get(h, '') for h in headers}, '_row': number})
            if not seen_header:
                raise ValueError('Missing migration table header row')
            result[name] = rows
        return parameters, result


def read_csv(path, delimiter=','):
    with Path(path).open(encoding='utf-8-sig', newline='') as stream:
        reader = csv.DictReader(stream, delimiter=delimiter)
        if not reader.fieldnames or len(set(reader.fieldnames)) != len(reader.fieldnames):
            raise ValueError('Missing or duplicate CSV columns')
        rows = list(reader)
        if any(None in row or any(v is None for v in row.values()) for row in rows):
            raise ValueError('Malformed evidence CSV')
        return rows


def inventory(last, producer, tenant, asof, maximum=168, overrides=None):
    """Verify the producer receipt, every declared file and source timestamps."""
    path = last / ('SmartInventory_' + producer + '.current.json.txt')
    receipt = json.loads(path.read_text(encoding='utf-8-sig'))
    if receipt.get('Owner') != 'SmartInventory-SourceReceipt' or receipt.get('TenantKey') != tenant or receipt.get('Producer') != producer + '.ps1' or receipt.get('Status') not in ('Completed', 'CompletedWithWarnings'):
        raise ValueError('Invalid SharePoint producer receipt identity or status')
    publication_run = text(receipt.get('RunId'))
    if not publication_run:
        raise ValueError('Missing SharePoint receipt publication RunId')
    captured = utc(receipt['CompletedAtUtc'])
    age = (asof - captured).total_seconds() / 3600
    overrides = overrides or {}
    group_limit = max([float(overrides.get(name, maximum)) for name in receipt.get('RequiredFiles', [])] or [maximum])
    if age < -5 / 60 or age > group_limit:
        raise ValueError('SharePoint capture is stale or in the future')
    files = {}
    inventory_run = None
    for entry in receipt.get('Files', []):
        name = entry.get('File')
        if not name or Path(name).name != name or '/' in name or '\\' in name or name in files:
            raise ValueError('Invalid or duplicate receipt filename')
        if entry.get('RunId') != publication_run or entry.get('Status') != 'Success':
            raise ValueError('Invalid SharePoint receipt entry publication identity or status: ' + name)
        candidate = last / name
        limit = float(overrides.get(name, maximum))
        observed_age = (asof - utc(entry.get('CompletedAtUtc', receipt['CompletedAtUtc']))).total_seconds() / 3600
        if not 0 < limit <= 8760 or observed_age < -5 / 60 or observed_age > limit:
            raise ValueError('Receipt payload violates its source acquisition-age policy')
        if sha(candidate) != text(entry.get('SHA256')).upper() or ('Bytes' in entry and candidate.stat().st_size != entry['Bytes']):
            raise ValueError('SharePoint receipt payload mismatch: ' + name)
        rows = read_csv(candidate, ';')
        if len(rows) != entry.get('Rows'):
            raise ValueError('SharePoint receipt row count mismatch')
        # Collectors own their CSV RunId; the receipt publisher owns a separate RunId.
        # Match receipt entries to the publication, and CSV rows to one inventory run.
        for row_number, row in enumerate(rows, 2):
            if row.get('TenantKey') != tenant:
                raise ValueError(f'SharePoint CSV tenant mismatch: {name}, row {row_number}')
            row_run = text(row.get('RunId'))
            if not row_run:
                raise ValueError(f'Missing SharePoint CSV inventory RunId: {name}, row {row_number}')
            if inventory_run is None:
                inventory_run = row_run
            elif row_run != inventory_run:
                raise ValueError(f'Mixed SharePoint CSV inventory RunId: {name}, row {row_number}')
        files[name] = rows
    if not set(receipt.get('RequiredFiles', [])).issubset(files):
        raise ValueError('Incomplete SharePoint receipt group')
    full = receipt.get('ConsumerScopeQualified') is True and receipt.get('FullInventoryQualified') is True and receipt.get('IsPartialInventory') is not True
    return files, {'Source': producer, 'Source SHA256': sha(path), 'Captured UTC': captured.isoformat(),
                   'Age Hours': round(age, 2), 'Qualification': 'Complete' if full else 'Observed only',
                   'Coverage State': receipt.get('ScopeQualification', 'Unknown'), 'Run ID': receipt.get('RunId', '')}


class ApprovedAliases(collections.defaultdict):
    """Exact-path host equivalences from an explicitly approved private registry."""
    def __init__(self, hosts=None):
        super().__init__(set)
        self.hosts = collections.defaultdict(set)
        for first, second in (hosts or {}).items():
            for host in [first, second]:
                if not isinstance(host, str) or url_key('https://' + host) != (host.casefold(), '', ''):
                    raise ValueError('Invalid approved host mapping')
            self.hosts[first.casefold()].add(second.casefold())
            self.hosts[second.casefold()].add(first.casefold())

    def get(self, key, default=None):
        values = set(super().get(key, default or set()))
        host, path, query = key[1]
        pending, visited = [host], {host}
        while pending:
            for other in self.hosts.get(pending.pop(), set()) - visited:
                visited.add(other); pending.append(other)
                values.add((other, path, query))
        return values


def bindings(path, tenant, project, asof):
    if not path:
        return {}, {}
    document = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if document.get('Identity') != {'TenantKey': tenant, 'ProjectKey': project} or document.get('SchemaVersion') != 1:
        raise ValueError('Alias registry identity/version mismatch')
    pairs = ApprovedAliases(document.get('ApprovedHostMappings', {}))
    for row in document.get('Bindings', []):
        if not text(row.get('Confirmation')) or not row.get('ConfirmedAtLocal') or utc(row['ConfirmedAtLocal']) > asof:
            raise ValueError('Alias binding lacks dated approval')
        first, second = url_key(row['SourceURL']), url_key(row['AliasURL'])
        if first[1:] != second[1:]:
            raise ValueError('Alias binding changes the resource path')
        pairs[(row['MigrationID'], first)].add(second)
        pairs[(row['MigrationID'], second)].add(first)
    targets = {}
    for row in document.get('TargetBindings', []):
        operation = row.get('OperationKey')
        if not operation or operation in targets or not row.get('EvidenceSHA256') or not row.get('ConfirmedAtUtc') or utc(row['ConfirmedAtUtc']) > asof:
            raise ValueError('Invalid or repeated target approval evidence')
        if not row.get('SourceReference') or not row.get('TargetURL') or not row.get('ParentMigrationID'):
            raise ValueError('Incomplete parent target binding')
        targets[operation] = row
    return pairs, targets


def native_key(row):
    fields = ['TenantKey', 'FarmId', 'SiteCollectionId', 'WebId']
    if any(not text(row.get(f)) for f in fields):
        raise ValueError('Incomplete native web identity')
    return '|'.join(text(row[f]).casefold() for f in fields)


def resolve(parameters, tables, tenant, webs, collections_rows, source_states, asof, aliases=None, target_bindings=None, target_tenant_url='', review_days=7, source_sha='', reporting_date=''):
    if parameters.get('TenantKey') != tenant or not text(parameters.get('ProjectKey')):
        raise ValueError('Workbook tenant/project identity is missing or different')
    project = parameters['ProjectKey']
    snapshot_date = date(reporting_date) if reporting_date else asof.date().isoformat()
    snapshot = token(tenant, project, snapshot_date, source_sha, target_tenant_url, review_days, [(s['Source'], s.get('Source SHA256', ''), s.get('Binding SHA256', '')) for s in source_states])
    base = {'TenantKey': tenant, 'Project Key': project, 'Snapshot ID': snapshot, 'Snapshot Date': snapshot_date}
    aliases = aliases if aliases is not None else {}
    target_bindings = target_bindings or {}
    by_native, by_url = {}, collections.defaultdict(set)
    for web in webs:
        if web.get('TenantKey') != tenant:
            raise ValueError('Foreign native web identity')
        key = native_key(web)
        if key in by_native:
            raise ValueError('Duplicate native web key')
        by_native[key] = web
        by_url[url_key(web['Url'])].add(key)
    output = {name: [] for name in ['Plan', 'Web', 'Collection', 'LegacyWeb', 'LegacyCollection', 'Scope', 'Finding', 'History', 'SourceSummary']}
    findings, seen = [], set()

    def finding(operation, rule, severity='Review', observed='', action='Review the source workbook'):
        findings.append({**base, 'Finding Key': token(snapshot, operation.get('Operation Key', ''), rule),
                         'Operation Key': operation.get('Operation Key', ''), 'Source Type': operation.get('Source Type', ''),
                         'Migration ID': operation.get('Migration ID', ''), 'Source Name': operation.get('Source Name', ''),
                         'Severity': severity, 'Finding Rule': rule, 'Observed Value': text(observed), 'Expected Action': action,
                         'Source Reference': operation.get('Source Reference', ''), 'Capture UTC': asof.isoformat()})

    for sheet, flow in FLOWS.items():
        for row in tables[sheet]:
            mid = text(row.get('MigrationID'))
            if not mid:
                raise ValueError('A nonempty migration row requires MigrationID')
            op_key = token(tenant, project, FLOW_CODES[flow], mid)
            if op_key in seen:
                raise ValueError('Duplicate tenant/project/source-type/migration identity')
            seen.add(op_key)
            for field, expected in [('TenantKey', tenant), ('ProjectKey', project)]:
                if text(row.get(field)) and row[field] != expected:
                    raise ValueError('Conflicting explicit workbook identity')
            source = text(row.get('SourceURL') if sheet == 'MigrationSP' else row.get('SourcePath'))
            op = {**base, 'Operation Key': op_key, 'Migration ID': mid, 'Source Type': flow,
                  'Business Unit': text(row.get('BusinessUnit')) or 'Unknown', 'Phase': text(row.get('Phase')) or 'Unknown',
                  'Project Scope Key': token(tenant, project), 'Source Reference': source,
                  'Source URL Aliases': text(row.get('SourceURLAliases')), 'Source Name': text(row.get('SourceName') or row.get('FolderName')),
                  'Disposition': text(row.get('Disposition')), 'Execution Status': text(row.get('ExecutionStatus')),
                  'Validation Status': text(row.get('ValidationStatus')), 'Target URL': text(row.get('TargetURL')),
                  'Target Review Status': text(row.get('TargetReviewStatus')), 'Target Mapping Level': text(row.get('TargetObjectType')) or ('FOLDER' if sheet == 'MigrationSharedFolders' else 'Unknown'),
                  'Target Approval Evidence': 'Workbook confirmed' if row.get('TargetReviewStatus') == 'CONFIRMED' else 'Not recorded',
                  'Target Observation': 'Not verifiable', 'Target Precision Status': 'Review required',
                  'Source Object Type': text(row.get('SourceObjectType')) or ('FOLDER' if sheet == 'MigrationSharedFolders' else 'Unknown'),
                  'Source Scope': text(row.get('SourceScope')), 'Source Qualification': 'Not observed',
                  'Business Contact': text(row.get('BusinessContact')), 'Comments': text(row.get('Comments')),
                  'Permission Strategy': text(row.get('PermissionStrategy')), 'Audience': text(row.get('Audience')),
                  'Source Row': row['_row'], 'Source SHA256': source_sha, 'Capture UTC': asof.isoformat(),
                  'Native Web Count': 0, 'Is Validated': False, 'Requires Review': False, 'Present In Register': True,
                  'Imported Date Meaning': text(row.get('ImportedDateMeaning')), 'Legacy Web ID': text(row.get('LegacyWebID'))}
            for source_field, target_field in [('PlannedDate', 'Planned Date'), ('ActualDate', 'Actual Date'), ('ValidationDate', 'Validation Date'), ('ImportedMigrationDate', 'Imported Migration Date')]:
                try:
                    op[target_field] = date(row.get(source_field))
                except ValueError:
                    op[target_field] = ''
                    finding(op, 'Invalid ' + target_field, observed=row.get(source_field))
            for field in ('ImportedTargetSite', 'ImportedTargetLibrary', 'ImportedTargetFolder'):
                op[re.sub(r'(?<!^)(?=[A-Z])', ' ', field)] = text(row.get(field))
            if op['Disposition'] not in DISPOSITIONS or op['Execution Status'] not in EXECUTIONS:
                finding(op, 'Invalid business status', observed=op['Disposition'] + '/' + op['Execution Status'])
            if op['Disposition'] == 'TBD':
                finding(op, 'Decision pending')
            if op['Imported Date Meaning'] == 'TO_CLARIFY':
                finding(op, 'Imported date needs clarification')
            done = op['Execution Status'] == 'DONE'
            if done and not op['Actual Date']:
                finding(op, 'Done without actual date')
            if op['Actual Date'] and op['Actual Date'] > snapshot_date:
                finding(op, 'Actual date in future')
            valid_date = bool(op['Validation Date']) and op['Validation Date'] <= snapshot_date and (not op['Actual Date'] or op['Validation Date'] >= op['Actual Date'])
            op['Is Validated'] = done and op['Validation Status'] == 'VALIDATED' and valid_date
            if op['Validation Status'] == 'VALIDATED' and not op['Is Validated']:
                finding(op, 'Inconsistent validation')
            if op['Disposition'] == 'MIGRATE' and done and not op['Is Validated']:
                finding(op, 'Migration validation not documented', action='Record migration acceptance and its date')
            if op['Disposition'] == 'MIGRATE' and op['Execution Status'] not in ('DONE', 'CANCELLED') and not op['Planned Date']:
                finding(op, 'Missing planned date')
            target = op['Target URL']
            if target.startswith('/') and not target.startswith('//') and target_tenant_url:
                target = target_tenant_url.rstrip('/') + target
            try:
                host, _, _ = url_key(target)
                if not host.endswith('.sharepoint.com') or (target_tenant_url and host != url_key(target_tenant_url)[0]):
                    raise ValueError()
                op['Resolved Target URL'] = target
            except ValueError:
                op['Resolved Target URL'] = ''
                if op['Disposition'] in ('MIGRATE', 'ARCHIVE'):
                    finding(op, 'Missing or unresolved target', observed=op['Target URL'])
            proof = target_bindings.get(op_key)
            parent = next((r for r in tables['MigrationSP'] if proof and r['MigrationID'] == proof['ParentMigrationID']), None)
            parent_matches = False
            if parent and parent.get('TargetURL') == op['Target URL']:
                try:
                    parent_key, source_key = url_key(parent['SourceURL']), url_key(source)
                    parent_keys = {parent_key} | aliases.get((parent['MigrationID'], parent_key), set())
                    source_keys = {source_key} | aliases.get((mid, source_key), set())
                    parent_matches = any(ancestor(p, c) for p in parent_keys for c in source_keys)
                    # A separately approved exact target edge can cross source hosts.
                    # It never grants a general source-domain equivalence or scope.
                    if proof.get('ParentSourceReference') == parent['SourceURL']:
                        parent_matches = parent_matches or ancestor((source_key[0], parent_key[1], parent_key[2]), source_key)
                except ValueError:
                    pass
            if proof and proof['SourceReference'] == source and proof['TargetURL'] == op['Target URL'] and parent_matches:
                op['Target Approval Evidence'] = 'Approved parent mapping'
            approved = op['Target Approval Evidence'] != 'Not recorded'
            op['Target Precision Status'] = 'Approved at stated level' if approved else 'Review required'
            if op['Disposition'] in ('MIGRATE', 'ARCHIVE') and not approved:
                finding(op, 'Target approval pending', action='Review target mapping at the required site/library/folder level')
            if sheet == 'MigrationSharedFolders':
                op['Source Qualification'] = 'Declared path only' if re.fullmatch(r'\\\\[^\\]+\\[^\\]+(?:\\.*)?', source) else 'Invalid source path'
                if op['Source Qualification'] == 'Invalid source path':
                    finding(op, 'Invalid source path')
            else:
                candidates = set()
                try:
                    normalized = url_key(source)
                    candidates.update(by_url.get(normalized, set()))
                    for alias in aliases.get((mid, normalized), set()):
                        candidates.update(by_url.get(alias, set()))
                    stored = {v.strip().casefold() for field in ('OnPremWebKey', 'OnPremAdditionalWebKeys') for v in text(row.get(field)).splitlines() if v.strip()}
                    # A stale/native key with a different URL is not silently rebound.
                    permitted = {normalized} | aliases.get((mid, normalized), set())
                    candidates.update(k for k in stored if k in by_native and url_key(by_native[k]['Url']) in permitted)
                    if len(candidates) > 1 and not aliases.get((mid, normalized)):
                        op['Source Qualification'] = 'Ambiguous'
                        finding(op, 'Ambiguous source identity')
                        candidates.clear()
                    elif candidates:
                        op['Source Qualification'] = 'Direct native match'
                    elif op['Source Object Type'] in CONTENT_OBJECT_TYPES:
                        parents = [(len(u[1]), keys) for u, keys in by_url.items() if ancestor(u, normalized)]
                        if parents:
                            candidates = set(max(parents, key=lambda v: v[0])[1])
                            op['Source Qualification'] = 'Parent context only'
                except ValueError:
                    op['Source Qualification'] = 'Invalid source URL'
                if not candidates:
                    finding(op, 'Source identity not qualified', observed=op['Source Qualification'])
                if candidates and op['Disposition'] == 'DELETE' and done:
                    finding(op, 'Declared deletion still observed', action='Review decision and completion evidence; preserve the original date')
                op['_native'] = candidates
                op['Native Web Count'] = len(candidates)
            output['Plan'].append(op)

    # Native hierarchy is derived only within the same tenant/farm/collection.
    parents = {}
    native_groups = collections.defaultdict(list)
    paths = collections.defaultdict(set)
    for key, web in by_native.items():
        group = (web['FarmId'], web['SiteCollectionId'])
        native_groups[group].append((key, web))
        paths[(group, url_key(web['Url']))].add(key)
    for key, web in by_native.items():
        host, path, query = url_key(web['Url'])
        parents[key] = ''
        while path:
            path = path.rpartition('/')[0]
            candidates = paths.get(((web['FarmId'], web['SiteCollectionId']), (host, path, query)), set())
            if candidates:
                parents[key] = next(iter(candidates)) if len(candidates) == 1 else ''
                break
    claims = collections.defaultdict(list)
    for op in output['Plan']:
        if op['Source Type'] == 'Shared folders':
            continue
        for native in op.get('_native', set()):
            own = by_native[native]
            is_context = op['Source Qualification'] == 'Parent context only'
            partial = is_context or op['Source Object Type'] in CONTENT_OBJECT_TYPES
            cover = 'ParentContext' if is_context else 'PartialContent' if partial else 'WholeWeb'
            claims[native].append((op, cover, len(url_key(own['Url'])[1]), True))
            if op['Source Scope'] == 'INCLUDE_CHILDREN' and not partial:
                for child_key, child in native_groups[(own['FarmId'], own['SiteCollectionId'])]:
                    if child_key != native and child['FarmId'] == own['FarmId'] and child['SiteCollectionId'] == own['SiteCollectionId'] and ancestor(url_key(own['Url']), url_key(child['Url'])):
                        claims[child_key].append((op, 'WholeWeb', len(url_key(own['Url'])[1]), False))
    for key, web in by_native.items():
        unique = {}
        for claim in claims[key]:
            signature = (claim[0]['Operation Key'], claim[1])
            if signature not in unique or (claim[2], claim[3]) > (unique[signature][2], unique[signature][3]):
                unique[signature] = claim
        current = list(unique.values())
        nearest = max((r[2] for r in current if r[1] == 'WholeWeb'), default=-1)
        winners = [r for r in current if r[1] == 'WholeWeb' and r[2] == nearest]
        decisions = {r[0]['Disposition'] for r in winners}
        coverage = 'Conflict' if len(decisions) > 1 else 'Whole web coverage' if winners and decisions != {'NOTINPLAN'} else 'Partial content only' if any(r[1] == 'PartialContent' for r in current) else 'NOTINPLAN'
        for op, kind, depth, direct in current:
            output['Scope'].append({**base, 'Scope Key': token(snapshot, op['Operation Key'], key, kind), 'Operation Key': op['Operation Key'], 'Web Key': key,
                                    'Collection Key': '|'.join(key.split('|')[:3]), 'Coverage Kind': kind, 'Is Effective': kind == 'WholeWeb' and depth == nearest,
                                    'Is Direct': direct, 'Effective Disposition': op['Disposition'] if depth == nearest else '', 'Has Conflict': len(decisions) > 1})
        output['Web'].append({**base, 'Web Key': key, 'Collection Key': '|'.join(key.split('|')[:3]), 'Parent Web Key': parents[key],
                              'Source URL': web['Url'], 'Source Name': web.get('Title', ''), 'Farm Key': tenant + '|' + web['FarmId'],
                              'Web Application Key': tenant + '|' + web['FarmId'] + '|' + web['WebApplicationId'],
                              'Coverage Kind': coverage, 'Run ID': web['RunId'], 'Captured UTC': web['CollectedAtUtc']})
    collection_webs = collections.defaultdict(list)
    scopes_by_web = collections.defaultdict(list)
    for web in output['Web']:
        collection_webs[web['Collection Key']].append(web)
    for scope in output['Scope']:
        scopes_by_web[scope['Web Key']].append(scope)
    collection_seen = set()
    for row in collections_rows:
        if row.get('TenantKey') != tenant:
            raise ValueError('Foreign native collection identity')
        collection = '|'.join(text(row[f]).casefold() for f in ('TenantKey', 'FarmId', 'SiteCollectionId'))
        if collection in collection_seen:
            raise ValueError('Duplicate native collection identity')
        collection_seen.add(collection)
        native = collection_webs[collection]
        fully = bool(native) and all(w['Coverage Kind'] == 'Whole web coverage' for w in native)
        migration_only = fully and all({s['Effective Disposition'] for s in scopes_by_web[w['Web Key']] if s['Is Effective']} == {'MIGRATE'} for w in native)
        complete = source_states and source_states[0].get('Qualification') == 'Complete'
        output['Collection'].append({**base, 'Collection Key': collection, 'Source URL': row['Url'], 'Storage Bytes': row.get('StorageBytes', ''),
                                     'Is Fully Covered': fully and complete, 'Is Fully Migrating': migration_only and complete,
                                     'Captured UTC': row['CollectedAtUtc'], 'Run ID': row['RunId']})
    for sheet, short, idfield in [('LegacyWebs', 'LegacyWeb', 'LegacyWebID'), ('LegacyCollections', 'LegacyCollection', 'LegacyCollectionID')]:
        legacy_seen = set()
        for row in tables[sheet]:
            ident = text(row[idfield])
            if not ident or ident in legacy_seen:
                raise ValueError('Missing or duplicate historical identity')
            legacy_seen.add(ident)
            output[short].append({**base, 'Legacy Key': token(tenant, project, ident), 'Legacy ID': ident,
                                  'Legacy Collection Key': token(tenant, project, text(row.get('LegacyCollectionID'))) if short == 'LegacyWeb' else '',
                                  'Source URL': row.get('SourceURL') or row.get('SiteCollectionURL', ''),
                                  'Source Name': row.get('SiteTitle', ''), 'Reported Size': row.get('SizeMB_Reported', ''),
                                  'Snapshot Date': date(parameters.get('LegacySnapshotDate'))})
    review = date(parameters.get('ReviewDate'))
    state = 'Missing review date' if not review else 'Future review date' if review > snapshot_date else 'Stale review' if (dt.date.fromisoformat(snapshot_date) - dt.date.fromisoformat(review)).days > review_days else 'Current'
    if state != 'Current':
        finding({}, 'Workbook ' + state, action='Review the migration plan and its ReviewDate')
    output['SourceSummary'] = [{**base, **s, 'Review Date': review, 'Review State': state} for s in source_states]
    output['SourceSummary'].append({**base, 'Source': 'Migration workbook', 'Source SHA256': source_sha, 'Captured UTC': asof.isoformat(),
                                    'Qualification': 'Configured', 'Coverage State': 'Declared operations', 'Review Date': review, 'Review State': state})
    for op in output['Plan']:
        mine = [f for f in findings if f['Operation Key'] == op['Operation Key']]
        op['Requires Review'] = bool(mine)
        op['Next Action'] = mine[0]['Finding Rule'] if mine else 'No recorded action'
        op.pop('_native', None)
    output['Finding'] = findings
    output['History'] = [dict(op, **{'History Key': token(snapshot, op['Operation Key'])}) for op in output['Plan']]
    return output


def merge_history(current, previous, tenant, project, snapshot, snapshot_date, captured_utc=''):
    """Preserve the old attributes and explicitly record disappearance/reopening."""
    if any(r.get('TenantKey') != tenant for r in previous):
        raise ValueError('Historical tenant/project mismatch')
    keys = [r['History Key'] for r in previous]
    if len(keys) != len(set(keys)):
        raise ValueError('Duplicate prior snapshot operation key')
    result = {r['History Key']: r for r in previous if r['Snapshot ID'] != snapshot}
    present = {r['Operation Key'] for r in current}
    latest = {}
    for row in previous:
        if row.get('Project Key') != project:
            continue
        op = row['Operation Key']
        if op not in latest or (row['Snapshot Date'], row.get('Capture UTC', '')) > (latest[op]['Snapshot Date'], latest[op].get('Capture UTC', '')):
            latest[op] = row
    for key, row in latest.items():
        if key not in present:
            current.append(dict(row, **{'Snapshot ID': snapshot, 'Snapshot Date': snapshot_date, 'History Key': token(snapshot, key),
                                       'Present In Register': False, 'Capture UTC': captured_utc or row.get('Capture UTC', '')}))
    result.update({r['History Key']: r for r in current})
    return sorted(result.values(), key=lambda r: (r['Snapshot Date'], r.get('Capture UTC', ''), r['Operation Key']))


def write_outputs(output, destination):
    definitions = json.loads(CONTRACT.read_text(encoding='utf-8'))['tables']
    destination.mkdir(parents=True, exist_ok=True)
    for table in definitions:
        rows = output[table['projection']]
        with (destination / table['file']).open('w', encoding='utf-8-sig', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=[c['name'] for c in table['columns']], extrasaction='ignore')
            writer.writeheader()
            writer.writerows(rows)


def build(args):
    asof = utc(args.as_of) if args.as_of else dt.datetime.now(UTC)
    if getattr(args, 'data_root', None):
        source_root, output_root = Path(args.data_root).resolve(), Path(args.output_root).resolve()
        if output_root == source_root or output_root == source_root/'DATA-LAST' or source_root/'DATA-LAST' in output_root.parents:
            raise ValueError('Use a dedicated output directory, never the raw source root')
    if not args.enabled:
        output = {t['projection']: [] for t in json.loads(CONTRACT.read_text(encoding='utf-8'))['tables']}
        # Every uniqueKey part must be non-blank: the prepared batch validator rejects an empty key part.
        output['SourceSummary'] = [{'TenantKey': args.tenant, 'Project Key': DISABLED_PROJECT_KEY, 'Snapshot ID': token(args.tenant, 'disabled'), 'Snapshot Date': asof.date().isoformat(),
                                   'Source': 'Migration workbook', 'Qualification': 'Not configured', 'Coverage State': 'Not configured', 'Captured UTC': asof.isoformat()}]
        if args.previous_history:
            output['History'] = read_csv(args.previous_history)
            if any(r.get('TenantKey') != args.tenant for r in output['History']):
                raise ValueError('Foreign tenant in retained disabled-feature history')
    else:
        source = Path(args.data_root) / args.workbook
        parameters, tables = read_workbook(source)
        last = Path(args.data_root) / 'DATA-LAST'
        policy = json.loads(Path(args.source_policy).read_text(encoding='utf-8-sig')) if getattr(args, 'source_policy', '') else {}
        maximum, overrides = policy.get('MaxSourceAgeHours', args.max_age_hours), policy.get('AgeOverrides', {})
        content, content_state = inventory(last, 'SmartM365-SharePoint-OnPrem-Content-Inventory', args.tenant, asof, maximum, overrides)
        _, infra_state = inventory(last, 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory', args.tenant, asof, maximum, overrides)
        registry = Path(args.data_root) / args.bindings if args.bindings else None
        aliases, targets = bindings(registry, args.tenant, parameters.get('ProjectKey'), asof)
        if registry:
            content_state['Binding SHA256'] = sha(registry)
        # Collection observation only; no assertion of library/folder/content existence.
        spo = Path(args.output_root) / 'DimSharePointSite.csv'
        if not spo.exists():
            raise ValueError('Prepared SPO site projection is required for target observations')
        observed = set()
        for row in read_csv(spo):
            if row.get('TenantKey') != args.tenant:
                raise ValueError('Foreign tenant in SPO projection')
            if row.get('SiteUrl'):
                observed.add(url_key(row['SiteUrl']))
        spo_state = {'Source': 'Prepared SPO usage sites', 'Source SHA256': sha(spo), 'Captured UTC': '',
                     'Qualification': 'Observed only', 'Coverage State': 'Usage report; not an exhaustive existence census'}
        output = resolve(parameters, tables, args.tenant, content['SharePoint_OnPrem_Webs.csv'], content['SharePoint_OnPrem_SiteCollections.csv'],
                         [content_state, infra_state, spo_state], asof, aliases, targets, args.target_tenant_url, args.review_days, sha(source), args.reporting_date)
        for row in output['Plan']:
            if row['Resolved Target URL']:
                url = url_key(row['Resolved Target URL'])
                row['Target Observation'] = 'Site observed' if any(ancestor(site, url) for site in observed if site[1]) else 'Not observed' if observed else 'Not verifiable'
        if args.previous_history:
            previous = read_csv(args.previous_history)
        else:
            previous = []
        summary = output['SourceSummary'][-1]
        # Rebuild new history rows after target observation qualification.
        output['History'] = [dict(r, **{'History Key': token(r['Snapshot ID'], r['Operation Key'])}) for r in output['Plan']]
        output['History'] = merge_history(output['History'], previous, args.tenant, parameters['ProjectKey'], summary['Snapshot ID'], summary['Snapshot Date'], asof.isoformat())
    write_outputs(output, Path(args.output_root))
    audit = {'Version': VERSION, 'Enabled': args.enabled, 'TenantKey': args.tenant, 'AsOfUtc': asof.isoformat(),
             'Rows': {k: len(v) for k, v in output.items()}, 'CollectionOrRefreshPerformed': False}
    (Path(args.output_root) / 'SharePointMigrationAudit.json.txt').write_text(json.dumps(audit, indent=2), encoding='utf-8')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--validate-workbook')
    parser.add_argument('--data-root')
    parser.add_argument('--output-root')
    parser.add_argument('--tenant')
    parser.add_argument('--enabled', action='store_true')
    parser.add_argument('--workbook', default='SmartWorkplaceIntelligence-Migration-to-SPO.xlsx')
    parser.add_argument('--bindings', default='')
    parser.add_argument('--previous-history', default='')
    parser.add_argument('--target-tenant-url', default='')
    parser.add_argument('--review-days', type=int, default=7)
    parser.add_argument('--max-age-hours', type=int, default=168)
    parser.add_argument('--as-of')
    parser.add_argument('--reporting-date', default='')
    parser.add_argument('--source-policy', default='')
    parser.add_argument('--validate-sources', action='store_true')
    args = parser.parse_args()
    if args.validate_workbook:
        parameters, tables = read_workbook(args.validate_workbook)
        seen = set()
        populated = any(tables[sheet] for sheet in FLOWS)
        if populated and (not parameters.get('TenantKey') or not parameters.get('ProjectKey')):
            raise ValueError('Active workbook requires tenant and project identity')
        if args.tenant and parameters.get('TenantKey') != args.tenant:
            raise ValueError('Workbook differs from the selected tenant')
        for sheet, flow in FLOWS.items():
            for row in tables[sheet]:
                key = (flow, row.get('MigrationID'))
                if not key[1] or key in seen:
                    raise ValueError('Missing or duplicate migration operation identity')
                seen.add(key)
                if any(text(row.get(f)) and row[f] != parameters.get(f) for f in ['TenantKey','ProjectKey']):
                    raise ValueError('Conflicting explicit workbook identity')
        if args.bindings:
            bindings(args.bindings, parameters.get('TenantKey'), parameters.get('ProjectKey'), dt.datetime.now(UTC))
        if args.validate_sources:
            asof = utc(args.as_of) if args.as_of else dt.datetime.now(UTC)
            policy = json.loads(Path(args.source_policy).read_text(encoding='utf-8-sig')) if args.source_policy else {}
            for producer in ['SmartM365-SharePoint-OnPrem-Content-Inventory','SmartM365-SharePoint-OnPrem-Infrastructure-Inventory']:
                inventory(Path(args.data_root)/'DATA-LAST',producer,args.tenant,asof,policy.get('MaxSourceAgeHours',args.max_age_hours),policy.get('AgeOverrides',{}))
        print(json.dumps({'SchemaVersion': '6.3', 'Rows': {k: len(v) for k, v in tables.items()}}))
        return
    if not args.tenant or not args.data_root or not args.output_root:
        parser.error('tenant, data-root and output-root are required')
    for name in [args.workbook, args.bindings]:
        if name and (Path(name).name != name or '/' in name or '\\' in name or name in ('.', '..')):
            parser.error('Configured inputs must be safe file names below the captured root')
    build(args)


if __name__ == '__main__':
    main()
