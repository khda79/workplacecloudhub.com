"""Current-only SmartInventory CMDB preparation. Python 3.10+, standard library.

No APIs, collector execution, report edits or history. Completion evidence is
mandatory; transport timestamps are never treated as acquisition timestamps.
"""
import argparse
import collections
import csv
import datetime as dt
import hashlib
import json
import os
import shutil
import uuid
from pathlib import Path
import cmdb_freshness

VERSION = '0.3.10'
OWNER = 'SmartInventory-CMDB-Prepared'
CONTRACT = Path(__file__).with_name('cmdb-prepared-contract.json.txt')
REGISTRY = Path(__file__).resolve().parents[2] / 'Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt'
MANIFEST = 'current.json.txt'
UTC = dt.timezone.utc
# Native AD group membership JSON can exceed Python's 128 KiB CSV default.
CSV_FIELD_LIMIT = 64 * 1024 * 1024


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(4 * 1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest().upper()


def utc(value):
    try:
        date = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
        if date.tzinfo is None:
            raise ValueError()
        return date.astimezone(UTC)
    except (ValueError, TypeError, AttributeError):
        raise ValueError('Acquisition timestamp requires an explicit timezone') from None


def rows(path):
    previous_limit = csv.field_size_limit()
    csv.field_size_limit(max(previous_limit, CSV_FIELD_LIMIT))
    try:
        with Path(path).open(encoding='utf-8-sig', newline='') as stream:
            reader = csv.DictReader(stream)
            if not reader.fieldnames or len(set(reader.fieldnames)) != len(reader.fieldnames):
                raise ValueError('Missing or duplicate CSV header: ' + Path(path).name)
            for row in reader:
                if None in row or any(value is None for value in row.values()):
                    raise ValueError('Malformed logical CSV row: ' + Path(path).name)
                yield row
    finally:
        csv.field_size_limit(previous_limit)


def header(path):
    with Path(path).open(encoding='utf-8-sig', newline='') as stream:
        return next(csv.reader(stream), [])


def normalized(value):
    return str(value or '').strip().casefold()


def check_rows(path, definition, tenant, exact=False, identity=None):
    columns = header(path)
    required = definition['columns']
    if (exact and columns != required) or not set(required).issubset(columns):
        raise ValueError('CSV schema mismatch: ' + path.name)
    seen, count = set(), 0
    for row in rows(path):
        count += 1
        if 'TenantKey' in columns and row['TenantKey'] != tenant:
            raise ValueError('Foreign or blank TenantKey: ' + path.name)
        if identity and any(row[field] != value for field, value in identity.items() if field in columns):
            raise ValueError('Source reporting identity mismatch: ' + path.name)
        key = tuple(normalized(row[field]) for field in definition['key'])
        optional = {'AssignedByGroupId'} if definition.get('name') == 'license_paths' else set()
        if (definition.get('name') == 'ad_members' and row.get('MembershipKind') == 'PrimaryGroup'
                and row.get('ResolutionStatus') == 'UnresolvedPrimaryGroup'):
            optional = {'GroupObjectGUID'}  # The native group SID remains mandatory.
        if any(not normalized(row[field]) for field in definition['key'] if field not in optional):
            raise ValueError('Blank immutable key: ' + path.name)
        if key and key in seen:
            raise ValueError('Duplicate immutable key: ' + path.name)
        seen.add(key)
    if count == 0 and not definition['allowEmpty']:
        raise ValueError('Unexpected empty table: ' + path.name)
    return count


def check_publication(path, proof):
    """Collecting preserves the previous proof; replacement or failed replacement does not."""
    run_path = path.with_name(path.name.replace('.current.json.txt', '.run.json.txt'))
    protocol = proof.get('PublicationProtocol')
    if protocol is not None and (type(protocol) is not int or protocol != 1):
        raise ValueError('Unsupported source publication protocol: ' + path.name)
    if protocol is None and not run_path.exists():
        return  # Existing completed receipts remain compatible, never fabricated.
    if run_path == path or run_path.is_symlink() or not run_path.is_file():
        raise ValueError('Missing or unsafe source run state: ' + path.name)
    run = load_json(run_path)
    if (run.get('Owner') != 'SmartInventory-SourceRun' or run.get('ContractVersion') != '1.0'
            or run.get('Status') not in {'Collecting', 'Publishing', 'Completed', 'Failed'}
            or not run.get('RunId') or type(run.get('PublicationStarted')) is not bool
            or type(run.get('UnqualifiedPublication')) is not bool
            or any(run.get(field) != proof.get(field) for field in
                   ('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId', 'Producer'))):
        raise ValueError('Invalid source run state: ' + path.name)
    if run['Status'] == 'Publishing' or run['UnqualifiedPublication']:
        raise ValueError('Canonical source publication is in progress or remains unqualified: ' + path.name)


def producer_records(source, identity):
    registry_hash = sha(REGISTRY)
    registry = load_json(REGISTRY)
    records, receipts = {}, []
    producers = set()
    for definition in registry['Producers']:
        name = definition['Script']
        if name in producers:
            raise ValueError('Duplicate producer registry entry')
        producers.add(name)
        path = source / definition['Receipt']
        context = f"Producer={name!r}; Receipt={definition['Receipt']!r}"
        if Path(definition['Receipt']).name != definition['Receipt'] or path.is_symlink():
            raise ValueError('Unsafe or linked producer receipt: ' + context)
        if not path.is_file():
            raise ValueError('Missing producer completion proof: ' + context)
        digest = sha(path)
        proof = load_json(path)
        check_publication(path, proof)
        receipt_kind = (proof.get('Owner'), proof.get('ContractVersion'))
        if receipt_kind not in {('SmartInventory-CmdbSourceReceipt', '1.1'),
                                ('SmartInventory-SourceReceipt', '1.2')}:
            raise ValueError('Producer completion proof owner or version mismatch: ' + context)
        shared = receipt_kind == ('SmartInventory-SourceReceipt', '1.2')
        coverage = accepted_ad_coverage(proof) if proof.get('DomainCoverage') is not None else None
        if shared and (proof.get('ScopeQualification') != 'ConsumerScope'
                       or set(proof.get('RequiredFiles', [])) != set(definition['Files'])):
            raise ValueError('Producer consumer-scope declaration mismatch: ' + context)
        if any(proof.get(field) != value for field, value in identity.items()):
            raise ValueError('Producer completion proof tenant identity mismatch: ' + context)
        if proof.get('Producer') != name or not proof.get('ScriptVersion') or not proof.get('RunId'):
            raise ValueError('Producer completion proof lineage mismatch: ' + context)
        if (proof.get('Status') != 'Completed' or proof.get('IsPartialInventory') is not bool(coverage)
                or type(proof.get('Errors')) is not int or proof['Errors'] != 0):
            raise ValueError('Incomplete producer completion proof: ' + context
                             + f"; Status={proof.get('Status')!r}"
                             + f"; IsPartialInventory={proof.get('IsPartialInventory')!r}"
                             + f"; Errors={proof.get('Errors')!r} ({type(proof.get('Errors')).__name__})")
        if proof.get('Scope') != definition['Scope']:
            raise ValueError('Producer full scope mismatch: ' + context)
        start, end = utc(proof['StartedAtUtc']), utc(proof['CompletedAtUtc'])
        local = {}
        for record in proof['Files']:
            file = record['File']
            if Path(file).name != file or (file in records and file in definition['Files']) or file in local:
                raise ValueError('Unsafe or repeated source proof filename: ' + context)
            if any(record.get(field) != proof[field] for field in ('Producer','ScriptVersion','RunId','StartedAtUtc','Scope')):
                raise ValueError('Source record lineage or scope differs from producer: ' + context)
            if record.get('DomainCoverage') != coverage or record.get('IsPartialInventory') is not bool(coverage):
                raise ValueError('Incomplete producer source: domain coverage differs from producer: ' + context)
            completed = utc(record['CompletedAtUtc'])
            if not start <= completed <= end:
                raise ValueError('Invalid acquisition interval inside producer receipt: ' + context)
            local[file] = record
        if (not set(definition['Files']).issubset(local)
                or (not shared and set(local) != set(definition['Files']))):
            raise ValueError('Missing producer completion proof or unexpected source file: ' + context)
        if sha(path) != digest:
            raise ValueError('Producer proof changed during validation: ' + context)
        # Additional exports belong to the shared producer, not to the CMDB source contract.
        records.update({file: local[file] for file in definition['Files']})
        receipts.append({'File':path.name, 'SHA256':digest, 'Producer':name,
                         'RunId':proof['RunId'], 'ScriptVersion':proof['ScriptVersion'],
                         'Scope':proof['Scope'], 'StartedAtUtc':proof['StartedAtUtc'],
                         'CompletedAtUtc':proof['CompletedAtUtc'],
                         'Qualifications':proof.get('Qualifications', []), 'DomainCoverage':coverage})
    return records, receipts, registry_hash


def accepted_ad_coverage(proof):
    """Only explicit, bounded AD connectivity gaps qualify; never generic partial data."""
    coverage = proof['DomainCoverage']
    if (proof.get('Producer') != 'SmartM365-ActiveDirectory-Inventory.ps1'
            or proof.get('Owner') != 'SmartInventory-SourceReceipt' or proof.get('ContractVersion') != '1.2'
            or proof.get('ConsumerScopeQualified') is not True or proof.get('IsPartialInventory') is not True):
        raise ValueError('Invalid partial AD coverage declaration')
    return cmdb_freshness.partial_ad_coverage(proof['Producer'], coverage)


def validate_license_assignment_parents(source, contract):
    """Check native parent IDs without dropping paths or inventing identities."""
    definitions = {item['name']: item for item in contract['sources']}
    user_file = definitions['users']['file']
    sku_file = definitions['skus']['file']
    path_file = definitions['license_paths']['file']
    users = {normalized(row['Object Id']) for row in rows(source / user_file)}
    skus = {normalized(row['Id']) for row in rows(source / sku_file)}
    missing_users, missing_skus = set(), set()
    missing_user_rows = missing_sku_rows = 0
    for row in rows(source / path_file):
        user_id, sku_id = normalized(row['UserId']), normalized(row['SkuId'])
        if not user_id or user_id not in users:
            missing_user_rows += 1
            missing_users.add(user_id)
        if not sku_id or sku_id not in skus:
            missing_sku_rows += 1
            missing_skus.add(sku_id)
    if missing_user_rows or missing_sku_rows:
        # Counts and file names suffice for diagnosis; do not expose account IDs.
        raise ValueError(
            'License assignment parent identity missing: '
            + f'Source={path_file}; UserParent={user_file} (UserId -> Object Id); '
            + f'MissingUserRows={missing_user_rows}; MissingUserIds={len(missing_users)}; '
            + f'SkuParent={sku_file} (SkuId -> Id); '
            + f'MissingSkuRows={missing_sku_rows}; MissingSkuIds={len(missing_skus)}. '
            + 'Refresh the affected parent inventory and revalidate coherent current sources; '
            + 'no assignment paths were excluded or historical exports substituted.')


def application_coverage_status(reported, observed, unresolved):
    parts = []
    if reported != observed:
        parts.append('Relation count differs')
    if unresolved:
        parts.append('Device links unresolved')
    return '; '.join(parts) or 'Complete'


def assess_application_coverage(source, contract):
    """Weekly app snapshots may differ from current devices; retain every relation."""
    definitions = {item['name']: item for item in contract['sources']}
    applications = {}
    for row in rows(source / definitions['apps']['file']):
        if row['CollectionScope'] != 'AllPlatforms' or row['RelationCollectionScope'] != 'All':
            raise ValueError('All-platform, All-mode applications are required')
        raw = row['DeviceCount'].strip()
        if not raw.isascii() or not raw.isdecimal():
            raise ValueError('Invalid application reported device count')
        applications[normalized(row['AppId'])] = int(raw)
    managed = {normalized(row['ManagedDeviceId']) for row in rows(source / definitions['managed']['file'])}
    counts = collections.Counter()
    observed = collections.Counter()
    missing_devices, missing_apps = set(), set()
    for row in rows(source / definitions['app_relations']['file']):
        aid, mid = normalized(row['AppId']), normalized(row['DeviceId'])
        counts['RelationRows'] += 1
        observed[aid] += 1
        if aid not in applications:
            counts['UnresolvedApplicationRelationRows'] += 1
            missing_apps.add(aid)
        if mid not in managed:
            counts['UnresolvedDeviceRelationRows'] += 1
            missing_devices.add(mid)
    coverage = {field: counts[field] for field in (
        'RelationRows', 'UnresolvedApplicationRelationRows', 'UnresolvedDeviceRelationRows')}
    coverage.update(DistinctUnresolvedApplicationIds=len(missing_apps),
                    DistinctUnresolvedDeviceIds=len(missing_devices),
                    CountMismatchApplications=sum(reported != observed[aid] for aid, reported in applications.items()),
                    Policy='Weekly native application evidence; unresolved relations retained without fabricated parents; not an instantaneous managed-device inventory')
    warnings = []
    for field, message in (
            ('UnresolvedDeviceRelationRows', 'Weekly application relations reference device IDs outside the current managed inventory'),
            ('UnresolvedApplicationRelationRows', 'Application relations reference IDs outside the collected application catalog'),
            ('CountMismatchApplications', 'Reported application device counts differ from collected relation counts')):
        if coverage[field]:
            warnings.append({'File': definitions['app_relations']['file'], 'Kind': field,
                             'Count': coverage[field], 'Message': message + f'; {field}={coverage[field]}. '
                             'Evidence is retained and qualified; no parent identity is fabricated.'})
    return coverage, warnings


def validate_sources(source, contract, tenant, now=None, identity=None):
    now = now or dt.datetime.now(UTC)
    if not identity or identity.get('TenantKey') != tenant:
        raise ValueError('Complete source identity is required')
    records, receipts, registry_hash = producer_records(source, identity)
    if contract['producerRegistry'] != REGISTRY.name:
        raise ValueError('Preparation producer registry mismatch')
    for receipt in receipts:
        start, end = utc(receipt['StartedAtUtc']), utc(receipt['CompletedAtUtc'])
        if start > end or end > now + dt.timedelta(minutes=5) or start > now + dt.timedelta(minutes=5):
            raise ValueError('Invalid producer acquisition interval')
        rule = cmdb_freshness.producer_policy(contract, [file for file, record in records.items()
                                                       if record['Producer'] == receipt['Producer']])
        if now - start > dt.timedelta(hours=rule['MaxAgeHours']):
            raise ValueError('Stale acquisition evidence: ' + receipt['Producer'])
    if set(records) != {s['file'] for s in contract['sources']}:
        raise ValueError('Producer registry and source contract disagree')
    checks, starts, ends = [], [], []
    for definition in contract['sources']:
        name = definition['file']
        if name not in records:
            raise ValueError('Missing producer completion proof: ' + name)
        record = records[name]
        coverage = record.get('DomainCoverage')
        if record.get('Status') != 'Success' or record.get('IsPartialInventory') is not bool(coverage) or type(record.get('Errors')) is not int or record['Errors'] != 0:
            raise ValueError('Incomplete producer result: ' + name)
        if not record.get('Producer') or not record.get('ScriptVersion') or not record.get('RunId'):
            raise ValueError('Missing producer lineage: ' + name)
        start, end = utc(record['StartedAtUtc']), utc(record['CompletedAtUtc'])
        if start > end or end > now + dt.timedelta(minutes=5) or start > now + dt.timedelta(minutes=5):
            raise ValueError('Invalid acquisition interval: ' + name)
        starts.append(start); ends.append(end)
        path = source / name
        if path.is_symlink() or not path.is_file():
            raise ValueError('Missing or linked source file: ' + name)
        before = sha(path)
        if before != str(record['SHA256']).upper():
            raise ValueError('Producer hash mismatch: ' + name)
        count = check_rows(path, definition, tenant, identity=identity)
        if coverage:
            collected = {v.casefold() for v in coverage['CollectedDomains']}
            if definition['name'] == 'ad_domains':
                if {normalized(row['DNSRoot']) for row in rows(path)} != collected:
                    raise ValueError('AD domain export differs from declared collected coverage')
            elif definition['name'] in {'ad_users', 'ad_computers', 'ad_groups', 'ad_objects'}:
                if (count and 'DomainName' not in header(path)) or any(normalized(row['DomainName']) not in collected for row in rows(path)):
                    raise ValueError('AD export includes an unavailable or unqualified domain: ' + name)
        if type(record['Rows']) is not int or count != record['Rows']:
            raise ValueError('Producer row count mismatch: ' + name)
        if sha(path) != before:
            raise ValueError('Source changed during validation: ' + name)
        checks.append(dict(record, SHA256=before))
    freshness = cmdb_freshness.evaluate(contract, checks, now)
    evidence = {'ProducerReceipts':receipts, 'RegistrySHA256':registry_hash, 'Files':checks,
                'StartedAtUtc':min(starts).isoformat(), 'CompletedAtUtc':max(ends).isoformat(),
                'Freshness':freshness, 'CoverageWarnings':[
                    {'Producer':r['Producer'], 'Message':'Partial AD coverage accepted; unavailable domains: '
                        + ', '.join(r['DomainCoverage']['UnavailableDomains']) + '; no historical exports substituted.'}
                    for r in receipts if r.get('DomainCoverage')]}
    validate_license_assignment_parents(source, contract)
    evidence['ApplicationCoverage'], evidence['ApplicationWarnings'] = assess_application_coverage(source, contract)
    recheck_sources(source, evidence, contract)
    return evidence


def recheck_sources(source, evidence, contract):
    if sha(REGISTRY) != evidence['RegistrySHA256']:
        raise ValueError('Producer registry changed during preparation')
    for receipt in evidence['ProducerReceipts']:
        check_publication(source / receipt['File'], load_json(source / receipt['File']))
        if sha(source / receipt['File']) != receipt['SHA256']:
            raise ValueError('Source proof changed during preparation')
    for record in evidence['Files']:
        if sha(source / record['File']) != record['SHA256']:
            raise ValueError('Source changed during preparation: ' + record['File'])


def validate_current(output, contract, tenant):
    if any(item.is_symlink() or not item.is_file() for item in output.iterdir()):
        raise ValueError('Linked or non-file output artifact')
    manifest = load_json(output / MANIFEST)
    if manifest.get('Owner') != OWNER or manifest.get('Status') != 'Validated' or manifest.get('TenantKey') != tenant or manifest.get('ContractVersion') != contract['version']:
        raise ValueError('Output is not the owned, validated tenant snapshot')
    expected = {item['name'] + '.csv' for item in contract['tables']}
    if {item.name for item in output.iterdir()} != expected | {MANIFEST}:
        raise ValueError('Unexpected or missing output artifacts')
    for definition in contract['tables']:
        file = definition['name'] + '.csv'
        if sha(output / file) != manifest['OutputFiles'][file]['SHA256']:
            raise ValueError('Output hash mismatch: ' + file)
        count = check_rows(output / file, definition, tenant, exact=True)
        if count != manifest['OutputFiles'][file]['Rows']:
            raise ValueError('Output row count mismatch: ' + file)
        for row in rows(output / file):
            for field, value in manifest['Identity'].items():
                if field in row and row[field] != value:
                    raise ValueError('Reporting identity mismatch: ' + file)
    validate_relationships(output, contract)
    return manifest


def validate_relationships(output, contract):
    parents = {'TenantDeviceKey':'DimDevice', 'TenantUserKey':'DimUser',
               'TenantGroupKey':'DimGroup', 'TenantSkuKey':'DimLicenseSku',
               'TenantServicePlanKey':'DimLicenseServicePlan',
               'TenantApplicationKey':'DimDetectedApplication', 'TenantTeamKey':'DimTeam',
               'TenantADGroupKey':'ADGroupSource', 'TenantADObjectKey':'ADDirectoryObject'}
    keys = {field: {row[field] for row in rows(output / (table+'.csv'))}
            for field, table in parents.items()}
    managed_ids = {row['ManagedDeviceId'] for row in rows(output / 'DimIntuneManagedDevice.csv')}
    applications = {normalized(row['AppId']): row for row in rows(output / 'DimDetectedApplication.csv')}
    application_counts, unresolved_devices = collections.Counter(), collections.Counter()
    mailboxes = {row['TenantMailboxKey'] for row in rows(output / 'FactMailbox.csv')}
    for definition in contract['tables']:
        name = definition['name']
        for row in rows(output / (name+'.csv')):
            for field, table in parents.items():
                if name != table and row.get(field) and row[field] not in keys[field]:
                    raise ValueError('Orphan output relationship: ' + name + '.' + field)
            if name == 'DeviceHardware' and row['ManagedDeviceId'] not in managed_ids:
                raise ValueError('Orphan output managed device: ' + name)
            if name == 'FactDeviceApplication':
                aid = normalized(row['AppId'])
                application = applications.get(aid)
                if (row['ApplicationLinkStatus'] != ('Resolved' if application else 'Unresolved')
                        or row['TenantApplicationKey'] != (application['TenantApplicationKey'] if application else '')):
                    raise ValueError('Inconsistent application link qualification')
                resolved = row['ManagedDeviceId'] in managed_ids
                if row['DeviceLinkStatus'] != ('Resolved' if resolved else 'Unresolved'):
                    raise ValueError('Inconsistent application device link qualification')
                application_counts[aid] += 1
                unresolved_devices[aid] += not resolved
            if name == 'FactMailboxHosting' and row['MailboxHostingKey'] not in mailboxes:
                raise ValueError('Orphan output mailbox hosting')
            for field in ('ManagerADObjectKey','ManagedByADObjectKey'):
                if row.get(field) and row[field] not in keys['TenantADObjectKey']:
                    raise ValueError('Orphan AD source owner/manager relationship')
            if row.get('NestedTenantGroupKey') and row['NestedTenantGroupKey'] not in keys['TenantGroupKey']:
                raise ValueError('Orphan nested Entra group relationship')
    for aid, row in applications.items():
        count, missing = application_counts[aid], unresolved_devices[aid]
        expected = application_coverage_status(int(row['ReportedDeviceCount']), count, missing)
        if (row['RelationshipCoverageStatus'] != expected or row['DeviceCount'] != str(count)
                or row['ExactRelatedDeviceCount'] != str(count)
                or row['ResolvedDeviceCount'] != str(count - missing)
                or row['UnresolvedDeviceCount'] != str(missing)):
            raise ValueError('Inconsistent application relationship coverage')


class PublicationLock:
    def __init__(self, path):
        self.path = path

    def __enter__(self):
        self.stream = self.path.open('a+b')
        if self.path.stat().st_size == 0:
            self.stream.write(b'0'); self.stream.flush()
        self.stream.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self.stream.close()
            raise ValueError('Another CMDB preparation owns the publication lock') from None
        return self

    def __exit__(self, *_):
        self.stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(self.stream.fileno(), msvcrt.LK_UNLCK, 1)
        self.stream.close()


def paths(source, output):
    source, output = Path(source).absolute(), Path(output).absolute()
    if source.name != 'DATA-LAST' or output.name != 'DATA-POWERBI-CMDB' or source.parent != output.parent:
        raise ValueError('CMDB preparation requires sibling DATA-LAST and DATA-POWERBI-CMDB')
    if source.is_symlink() or output.is_symlink() or source.resolve().parent != output.resolve().parent:
        raise ValueError('Linked or redirected output/source root')
    return source, output


def prepare(source, output, tenant, identity, validate_only=False, contract_path=CONTRACT, now=None, fault=None):
    source, output = paths(source, output)
    contract = load_json(contract_path)
    contract_hash = sha(contract_path)
    if not tenant or identity.get('TenantKey') != tenant or any(not identity.get(k) for k in ('OrganizationKey','EnvironmentKey','TenantId')):
        raise ValueError('Complete reporting tenant identity is required')
    evidence = validate_sources(source, contract, tenant, now, identity)
    if validate_only:
        return {'Status': 'ValidatedSources', 'SourceFiles': len(evidence['Files']), 'GeneratedTables': 0,
                'FreshnessWarnings':evidence['Freshness']['Warnings'], 'CoverageWarnings':evidence['CoverageWarnings'],
                'ApplicationWarnings':evidence['ApplicationWarnings'], 'ApplicationCoverage':evidence['ApplicationCoverage']}
    # No source copies, persistent Raw adapter, dated output or history.
    from cmdb_tables import build_tables
    with PublicationLock(output.parent / '.cmdb-preparation.lock'):
        orphan_stages = list(output.parent.glob('.cmdb-stage-*')) + list(output.parent.glob('.cmdb-rollback-*'))
        if orphan_stages:
            raise ValueError('Interrupted preparation artifacts require explicit recovery; nothing replaced')
        previous_manifest = validate_current(output, contract, tenant) if output.exists() else None
        if previous_manifest and previous_manifest['Identity'] != identity:
            raise ValueError('Reporting identity changed; explicit migration required')
        # Inherit the data-root ACL. Python 3.13+ mkdtemp uses an owner-only
        # Windows DACL that also excludes some sandbox / scheduled identities.
        stage = output.parent / ('.cmdb-stage-' + uuid.uuid4().hex)
        stage.mkdir()
        previous = output.parent / ('.cmdb-rollback-' + uuid.uuid4().hex)
        promoted = moved = False
        try:
            qualifications = build_tables(source, stage, contract, identity, evidence, now)
            output_files = {}
            for definition in contract['tables']:
                name = definition['name'] + '.csv'
                count = check_rows(stage / name, definition, tenant, exact=True)
                output_files[name] = {'Rows': count, 'SHA256': sha(stage / name)}
            manifest = {'Owner': OWNER, 'Status': 'Validated', 'ScriptVersion': VERSION,
                        'ContractVersion': contract['version'], 'ContractSHA256': contract_hash,
                        'TenantKey': tenant, 'Identity': identity,
                        'GeneratedAtUtc': (now or dt.datetime.now(UTC)).isoformat(),
                        'SourceRoot': str(source), 'OutputRoot': str(output),
                        'SourceEvidence': evidence, 'OutputFiles': output_files,
                        'PreparationQualifications': qualifications,
                        'MetricDefinitions': {
                        'TopApplication.ReportedDeviceCount':'Distinct device IDs observed in weekly application relations per name/publisher/platform product across versions, including unresolved current-inventory links; not a current managed-device count',
                        'FactDeviceApplication':'Every native application/device relation is retained; unresolved parent links are qualified, never fabricated',
                            'FactHybridIdentityCoverage.OnPremisesOnlyCount':'Unavailable: unmatched identity does not establish on-premises-only existence',
                            'FactUserActivity.HasAnyM365Activity':'Observation within the source report, not proof of lifetime use or licence waste',
                            'EndpointAnalyticsScore':'Score on a 0-100 scale, not a proportion; -1/-2 mean unavailable (blank), never zero',
                            'SourceFreshness':'Producer acquisition interval; workload report refresh dates are separate',
                            'HardwareCoverage':'Managed-device list properties only; not a full detailed hardware export'}}
            (stage / MANIFEST).write_text(json.dumps(manifest, indent=2), encoding='utf-8')
            validate_current(stage, contract, tenant)
            recheck_sources(source, evidence, contract)
            cmdb_freshness.evaluate(contract, evidence['Files'], now or dt.datetime.now(UTC))
            if sha(contract_path) != contract_hash:
                raise ValueError('Preparation contract changed during execution')
            if fault:
                fault('before-swap', source, stage)
            if output.exists():
                output.rename(previous); moved = True
            stage.rename(output); promoted = True
            if fault:
                fault('after-swap', source, output)
            validate_current(output, contract, tenant)
            recheck_sources(source, evidence, contract)
            cmdb_freshness.evaluate(contract, evidence['Files'], now or dt.datetime.now(UTC))
        except BaseException:
            if promoted:
                shutil.rmtree(output)
            if moved:
                previous.rename(output); moved = False
            raise
        finally:
            if stage.exists():
                shutil.rmtree(stage)
        if moved:
            # The only old snapshot is temporary rollback evidence, not history.
            try:
                shutil.rmtree(previous)
            except OSError:
                # Publication is already validated. Never undo it after partial
                # removal of the temporary previous snapshot.
                return {'Status': 'PreparedWithCleanupWarning', 'GeneratedTables': len(output_files),
                        'OutputRoot': str(output), 'CleanupRequired': str(previous),
                        'FreshnessWarnings':evidence['Freshness']['Warnings'], 'CoverageWarnings':evidence['CoverageWarnings'],
                        'ApplicationWarnings':evidence['ApplicationWarnings']}
        return {'Status': 'Prepared', 'GeneratedTables': len(output_files), 'OutputRoot': str(output),
            'FreshnessWarnings':evidence['Freshness']['Warnings'], 'CoverageWarnings':evidence['CoverageWarnings'],
            'ApplicationWarnings':evidence['ApplicationWarnings'], 'ApplicationCoverage':evidence['ApplicationCoverage']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', required=True)
    parser.add_argument('--output', required=True)
    for name in ('tenant-key','organization-key','environment-key','tenant-id'):
        parser.add_argument('--' + name, required=True)
    parser.add_argument('--validate-only', action='store_true')
    args = parser.parse_args()
    identity = {'TenantKey': args.tenant_key, 'OrganizationKey': args.organization_key,
                'EnvironmentKey': args.environment_key, 'TenantId': args.tenant_id}
    print(json.dumps(prepare(args.source, args.output, args.tenant_key, identity, args.validate_only)))


if __name__ == '__main__':
    main()
