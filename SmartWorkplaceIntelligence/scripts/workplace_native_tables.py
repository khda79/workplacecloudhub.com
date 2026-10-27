"""Native SmartInventory facets for the unified Workplace evidence batch; no collector or Power BI runtime.

Unknown evidence stays blank/qualified. Native IDs, never display names, drive
identity. Large application relations and service-plan facts are streamed.
"""
import base64
import collections
import csv
import datetime as dt
import hashlib
import json
import math
import re
import time
from pathlib import Path

from workplace_native_sources import rows, normalized as key, utc, application_coverage_status
from workplace_mailbox_identity import reconcile as reconcile_mailboxes
from workplace_native_types import count64, source_datetime, validate_value
from workplace_country import normalize_country
from workplace_device_identity import qualify as qualify_device_identity

VERSION = '0.5.0'
UNKNOWN = 'Unknown / unassigned'
ZERO = '00000000-0000-0000-0000-000000000000'


def get(row, *names):
    for name in names:
        if name in row and row[name] is not None:
            return row[name].strip()
    for name in names:
        for field, value in row.items():
            if field.casefold() == name.casefold() and value is not None:
                return value.strip()
    return ''


def boolean(value):
    if not value:
        return ''
    if key(value) not in ('true','false'):
        raise ValueError('Invalid boolean evidence')
    return key(value)


def typed_source(row, field, table, output_field=None, kind='dateTime'):
    try:
        return source_datetime(date_evidence(row, field)) if kind == 'dateTime' else count64(get(row, field))
    except (ValueError, OverflowError) as error:
        raise ValueError(f'{table}.csv column {output_field or field}, rule {kind}: invalid source value') from error


def number(value):
    if value == '':
        return ''
    result = float(value)
    if not math.isfinite(result) or result < 0:
        raise ValueError('Invalid non-negative numeric evidence')
    return result


def endpoint_score(value):
    """Only documented Endpoint Analytics sentinels become unavailable, never zero."""
    if value == '':
        return ''
    result = float(value)
    if result in (-1, -2):
        return ''
    if not math.isfinite(result) or not 0 <= result <= 100:
        raise ValueError('Endpoint Analytics score outside 0-100')
    return result


def date_evidence(row, field):
    """Prefer an additive round-trip value; never reinterpret the legacy column's zone."""
    return get(row, field + 'Iso8601') or get(row, field)


def qualified(value):
    if not value:
        return '', 'Not provided'
    try:
        normalized = source_datetime(value)
        if not normalized:
            return '', 'Sentinel / unknown'
        date = dt.datetime.fromisoformat(normalized.replace('Z', '+00:00'))
        if date.tzinfo is None:
            return '', 'Timezone not established'
        return utc(normalized).isoformat(), 'UTC qualified'
    except ValueError:
        return '', 'Invalid date format'


def lifecycle_date(row, prefix, *fields):
    """Preserve an explicit source property; never substitute activity dates."""
    field = next((field for field in fields if date_evidence(row, field)), fields[0])
    raw = get(row, field)
    value, status = qualified(date_evidence(row, field))
    return {prefix + 'Raw':raw, prefix + 'UtcDateTime':value,
            prefix + 'Status':status}


def password_expiry(row):
    """AD PasswordExpired is an observation, not a password-age inference."""
    supplied = any(field.casefold() == 'passwordexpired' for field in row)
    raw = get(row, 'PasswordExpired')
    value = boolean(raw)
    return {'PasswordExpired':value,
            'PasswordExpiryStatus':'Observed' if value else 'Not provided' if supplied else 'Not collected'}


def filetime(value):
    if not value or value == '0':
        return '', 'Not observed'
    try:
        if not value.isascii() or not value.isdecimal():
            raise ValueError()
        ticks = int(value)
        date = dt.datetime(1601,1,1,tzinfo=dt.timezone.utc) + dt.timedelta(microseconds=ticks // 10)
        if date.year < 1900:
            return '', 'Sentinel / unknown'
        return date.isoformat(), 'UTC from replicated AD FILETIME; approximate logon'
    except (ValueError, OverflowError):
        return '', 'Invalid FILETIME'


def byte_evidence(row, field):
    raw, status = get(row,field), get(row,field+'Status')
    if status == 'Missing' and raw == '':
        return '', status
    if not raw.isascii() or not raw.isdecimal() or int(raw) > 9223372036854775807:
        raise ValueError('Invalid detailed hardware byte evidence')
    value = int(raw)
    if status != ('ZeroReported' if value == 0 else 'Reported'):
        raise ValueError('Inconsistent detailed hardware byte status')
    return value, status


def activity_state(raw, collected, user_enabled=None):
    if user_enabled == 'false':
        return 'Disabled'
    if not raw:
        return 'No observed sign-in' if user_enabled is not None else 'Unknown'
    try:
        if user_enabled is not None and 'T' not in raw:
            return 'Timezone not established'
        if 'T' in raw and dt.datetime.fromisoformat(raw.replace('Z', '+00:00')).tzinfo is None:
            return 'Timezone not established'
        observed = utc(raw) if 'T' in raw else dt.datetime.combine(dt.date.fromisoformat(raw), dt.time(), dt.timezone.utc)
        age = (utc(collected)-observed).total_seconds()/86400
    except ValueError:
        return 'Invalid date'
    if age < -1:
        return 'Future date'
    if user_enabled is None:
        return 'Active (90d)' if age <= 90 else 'Inactive (>90d)'
    return 'Active 30d' if age <= 30 else 'Inactive 31-90d' if age <= 90 else 'Inactive over 90d'


def native_key(tenant, kind, *parts):
    # Digest a tuple rather than concatenate potentially colliding separators.
    digest = hashlib.sha256(json.dumps([key(p) for p in parts], separators=(',',':')).encode()).hexdigest()
    return tenant + '|' + kind + '|' + digest


def index(records, field):
    result = {}
    for record in records:
        identity = key(get(record, field))
        if not identity or identity in result:
            raise ValueError('Missing or ambiguous native identity: ' + field)
        result[identity] = record
    return result


def log_timing(label, started):
    """Operational duration on stdout; never includes values or identities."""
    print(f'[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] Native build timing: {label}; {time.perf_counter() - started:.1f}s', flush=True)


def build_tables(source, output, contract, identity, evidence, now=None):
    tenant = identity['TenantKey']
    reference = now or dt.datetime.now(dt.timezone.utc)
    definitions = {t['name']: t for t in contract['tables']}
    typed_contract = json.loads((Path(__file__).resolve().parents[1] / 'config/native-output-contract.json.txt').read_text(encoding='utf-8-sig'))
    typed_columns = {Path(t['file']).stem: [(c['name'], c['type']) for c in t['columns'] if c['type'] != 'string'] for t in typed_contract['tables']}
    source_defs = {s['name']: s for s in contract['sources']}
    proof = {f['File']: f for f in evidence['Files']}
    table_counts = {}
    findings = {}
    # Each table duration includes the preparation done since the previous table.
    table_started = [time.perf_counter()]

    def read(name):
        return rows(source / source_defs[name]['file'])

    def acquired(name):
        return proof[source_defs[name]['file']]['CompletedAtUtc']

    def emit(name, records):
        definition = definitions[name]
        count = 0
        with (output / (name+'.csv')).open('w', encoding='utf-8-sig', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=definition['columns'])
            writer.writeheader()
            for record in records:
                result = dict(identity, **record)
                # Display labels must never alter the underlying evidence value.
                for column in definition['columns']:
                    if column.endswith('Label') and column not in result and column[:-5] in result:
                        result[column] = result[column[:-5]] or 'Not provided'
                for column, kind in typed_columns[name]:
                    try:
                        validate_value(result.get(column, ''), kind)
                    except (ValueError, OverflowError) as error:
                        raise ValueError(f'{name}.csv record {count + 1}, column {column}, rule {kind}: invalid value') from error
                writer.writerow({column: result.get(column, '') for column in definition['columns']})
                count += 1
        table_counts[name] = count
        log_timing(f'table {name} rows={count}', table_started[0])
        table_started[0] = time.perf_counter()

    def finding(entity, entity_id, finding_type, message, severity='Warning', discriminator='', **links):
        fid = native_key(tenant, 'finding', entity, entity_id, finding_type, discriminator)
        findings[fid] = {'TenantFindingKey':fid,'FindingId':fid,'EntityType':entity,
                         'EntityId':entity_id,'Severity':severity,'FindingType':finding_type,
                         'Description':message,'RecommendedAction':'Review native source evidence before taking action.',
                         'DetectedDateTime':reference.isoformat(),'SourceSystem':'SmartInventory',**links}

    freshness = {item['File']:item for item in evidence['Freshness']['Sources']}
    for warning in evidence['Freshness']['Warnings']:
        finding('Source', warning['File'], 'SourceFreshnessWarning', warning['Message'])
    for warning in evidence['ApplicationWarnings']:
        finding('Source', warning['File'], 'ApplicationCoverageWarning', warning['Message'],
                severity='Information', discriminator=warning['Kind'])
    for warning in evidence['LicenseWarnings']:
        finding('Source', warning['File'], 'LicenseCoverageWarning', warning['Message'],
                discriminator=warning['Kind'])

    users, users_by_native = [], {}
    for row in read('users'):
        sid = get(row,'Object Id')
        location = get(row,'Usage location').upper()
        country_evidence = normalize_country(get(row, 'CountryOrRegion'))
        country = country_evidence['Code']
        enabled = boolean(get(row,'AccountEnabled'))
        sign_in = {name: typed_source(row, name, 'DimUser') for name in
                   ('LastSignInDateTime', 'LastNonInteractiveSignInDateTime', 'LastSuccessfulSignInDateTime')}
        item = {'TenantUserKey':native_key(tenant,'user',sid),'CmdbUserId':'entra-user|'+key(sid),
                'SourceUserId':sid,'UserPrincipalName':get(row,'User principal name'),
                'DisplayName':get(row,'Display name'),'AccountEnabled':enabled,'UserType':get(row,'UserType'),
                'Department':get(row,'Department'),'JobTitle':get(row,'Title'),
                'CountryCode':country,'CountryLabel':country or UNKNOWN,
                'CountryStatus':country_evidence['Status'],
                'CountryRaw':country_evidence['Raw'],'CountryName':country_evidence['Name'],
                'CountrySource':'Entra country','UsageLocation':location,
                **sign_in,
                'ActivityState':activity_state(sign_in['LastSuccessfulSignInDateTime'] or sign_in['LastSignInDateTime'], acquired('users'), enabled),
                'AccountStatusLabel':'Enabled' if enabled=='true' else 'Disabled' if enabled=='false' else 'Unknown',
                **lifecycle_date(row, 'Creation', 'When created'),
                **lifecycle_date(row, 'Modification', 'lastModifiedDateTime', 'ModifiedDateTime'),
                'SourceCollectedDateTime':acquired('users'),'ManagerStatus':'Not collected',
                'UserActivityStatus':'Separate workload evidence','UserSelection':get(row,'Display name')+' | '+sid}
        users.append(item); users_by_native[key(sid)] = item
        if not country:
            finding('User',item['CmdbUserId'],'UserCountryUnknown','Entra country is missing or unrecognized; no usage location or AD fallback.')
        if item['ActivityState'] in ('Invalid date','Future date','Timezone not established'):
            finding('User',item['CmdbUserId'],'InvalidActivityDate','User sign-in evidence is unqualified or in the future.')
    user_addresses = collections.defaultdict(list)
    for item in users:
        if item['UserPrincipalName']:
            user_addresses[key(item['UserPrincipalName'])].append(item)
    def user_by_address(address):
        candidates = user_addresses.get(key(address), [])
        return candidates[0] if len(candidates) == 1 else None
    emit('DimUser', users)

    group_scope = index(list(read('group_scope')), 'GroupId')
    # Licensing's independent catalog is comparison evidence, not a membership
    # cohort. New/deleted groups between two successful scans are not empty groups.
    licensing_groups = index(list(read('groups')), 'GroupId')
    native_groups = group_scope
    scope_proof = proof[source_defs['group_scope']['file']]
    # EvidenceRuntime and the shared receipt have distinct run identifiers.
    # Hashes bind files to the receipt; row run IDs bind this catalog to members.
    catalog_runs = {get(row, 'RunId') for row in native_groups.values()}
    if '' in catalog_runs or len(catalog_runs) > 1:
        raise ValueError('Entra group catalog contains missing or mixed row run identities')
    for row in native_groups.values():
        for field in ('GroupCollectedAtUtc', 'CollectedAtUtc'):
            if not utc(scope_proof['StartedAtUtc']) <= utc(row[field]) <= utc(scope_proof['CompletedAtUtc']):
                raise ValueError('Entra group catalog acquisition outside producer interval')
        if utc(row['GroupCollectedAtUtc']) > utc(row['CollectedAtUtc']):
            raise ValueError('Entra group catalog acquired after membership evidence')
    member_counts = collections.Counter()
    for member in read('group_members'):
        gid = key(member['GroupId'])
        if gid not in native_groups or member['MembershipKind'] != 'Direct':
            raise ValueError('Unqualified Entra group membership')
        if (get(member, 'RunId') != get(group_scope[gid], 'RunId')
                or utc(member['CollectedAtUtc']) != utc(group_scope[gid]['CollectedAtUtc'])):
            raise ValueError('Entra membership run or acquisition differs from group scope')
        member_counts[gid] += 1
    groups = {}
    for gid, row in native_groups.items():
        scope = group_scope[gid]
        if scope['MemberCollectionStatus'] != 'Collected' or int(scope['MemberCount']) != member_counts[gid]:
            raise ValueError('Entra membership count/completion mismatch')
        groups[gid] = {'TenantGroupKey':native_key(tenant,'group',gid),'CmdbGroupId':'entra-group|'+gid,
                       'DisplayName':get(row,'DisplayName'),'SourceGroupId':row['GroupId'],
                       'MailEnabled':boolean(get(row,'MailEnabled')),'SecurityEnabled':boolean(get(row,'SecurityEnabled')),
                       'GroupTypes':get(row,'GroupTypes'),'MemberStatus':'Collected',
                       'OwnerStatus':'Not collected','GroupSelection':get(row,'DisplayName')+' | '+gid,
                       'SourceCollectedDateTime':row['GroupCollectedAtUtc']}

    entra = index(list(read('entra_devices')), 'ObjectId')
    managed = index(list(read('managed')), 'ManagedDeviceId')
    identity_conflicts, identity_summary = qualify_device_identity(source, contract, identity, evidence)
    disputed_correlations = {key(r['AzureADDeviceId']) for mid,r in managed.items() if mid in identity_conflicts}
    def managed_correlation(row):
        mid = key(row['ManagedDeviceId'])
        corr = key(row['AzureADDeviceId'])
        return 'intune:' + mid if mid in identity_conflicts or not corr or corr == ZERO else corr
    # Multiple Intune records can share a correlation ID. Retain all source rows;
    # choose a deterministic candidate only with qualified acquisition dates.
    candidates = collections.defaultdict(list)
    for mid, row in managed.items():
        corr = managed_correlation(row)
        candidates[corr].append(row)
    selected = {}
    for corr, records in candidates.items():
        def rank(row):
            minimum = dt.datetime.min.replace(tzinfo=dt.timezone.utc)
            def rank_date(field):
                return utc(row[field]) if get(row,field) else minimum
            return (-rank_date('LastSyncDateTime').toordinal(), -rank_date('LastSyncDateTime').timestamp() if rank_date('LastSyncDateTime') != minimum else 0,
                    -rank_date('EnrolledDateTime').toordinal(), -rank_date('EnrolledDateTime').timestamp() if rank_date('EnrolledDateTime') != minimum else 0, key(row['ManagedDeviceId']))
        selected[corr] = sorted(records, key=rank)[0]
    entra_by_corr = {}
    for row in entra.values():
        corr = key(row['DeviceId'])
        if not corr or corr == ZERO or corr in entra_by_corr:
            raise ValueError('Missing or repeated native Entra device correlation ID')
        entra_by_corr[corr] = row
    devices = {}
    for corr in sorted(set(selected) | set(entra_by_corr)):
        md, ed = selected.get(corr), entra_by_corr.get(corr)
        row = md or ed
        conflict = bool(md and key(md['ManagedDeviceId']) in identity_conflicts)
        primary = users_by_native.get(key(get(md or {},'UserId'))) if not conflict else None
        sync_raw = get(md or {},'LastSyncDateTime')
        sync, sync_status = qualified(sync_raw)
        owner = key(get(md or {},'ManagedDeviceOwnerType'))
        owner = 'Corporate' if owner in ('company','corporate') else 'Personal' if owner == 'personal' else ''
        compliance = get(md or {},'ComplianceState')
        compliance = {'compliant':'Compliant','noncompliant':'NonCompliant','ingraceperiod':'InGracePeriod'}.get(key(compliance), compliance)
        item = {'TenantDeviceKey':native_key(tenant,'device',corr),'CmdbDeviceId':'device|'+corr,
                'SourceDeviceId':corr,'DeviceName':get(row,'DeviceName','DisplayName'),
                'OperatingSystem':get(row,'OperatingSystem'),'OperatingSystemVersion':get(row,'OsVersion','OperatingSystemVersion'),
                'Ownership':owner,'ComplianceState':compliance,'EncryptionState':boolean(get(md or {},'IsEncrypted')),
                'ManagementState':'Managed' if md else 'Not managed by Intune',
                'SourceSystems':'Entra;Intune' if md and ed else 'Intune' if md else 'Entra',
                'AssociatedAccount':primary['UserPrincipalName'] if primary else 'Not resolved',
                'AssociationStatus':'Identity conflict' if conflict else 'Resolved' if primary else 'Unresolved' if get(md or {},'UserId') else 'Not provided',
                'PrimaryUserSourceId':get(md or {},'UserId'), 'CountryCode':primary['CountryCode'] if primary else '',
                'CountryLabel':primary['CountryLabel'] if primary else UNKNOWN,
                'CountryStatus':'Identity conflict; attribution indeterminate' if conflict else ('Primary user Entra country' if primary['CountryCode'] else 'Primary user Entra country unavailable') if primary else 'No resolved primary user',
                'SyncRaw':sync_raw,'SyncUtcDateTime':sync,'SyncStatus':sync_status,
                **lifecycle_date(ed or {}, 'Creation', 'createdDateTime'),
                **lifecycle_date(ed or {}, 'Modification', 'lastModifiedDateTime', 'ModifiedDateTime'),
                'SourceCollectedDateTime':acquired('managed' if md else 'entra_devices')}
        item['DeviceSelection'] = item['DeviceName']+' | '+corr
        devices[corr] = item
        if conflict:
            finding('Device',item['CmdbDeviceId'],'DeviceIdentityConflict','Independent Intune observations disagree after three bounded attempts; Entra/AD and derived attribution remain indeterminate.', TenantDeviceKey=item['TenantDeviceKey'])
        if md and get(md,'UserId') and not primary and not conflict:
            finding('Device',item['CmdbDeviceId'],'OrphanPrimaryUserReference','Native Intune user ID is outside the collected Entra user population.')
        elif md and not get(md,'UserId'):
            finding('Device',item['CmdbDeviceId'],'DeviceWithoutPrimaryUser','Native Intune evidence does not provide a primary-user ID.')
        if not item['CountryCode']:
            finding('Device',item['CmdbDeviceId'],'DeviceCountryUnknown','Device country cannot be established from a uniquely resolved primary user.')
        else:
            finding('Device',item['CmdbDeviceId'],'DeviceCountryDerived','Device country comes from the primary user Entra country, not device geolocation.', 'Information')
        if len(candidates.get(corr,[])) > 1:
            finding('Device',item['CmdbDeviceId'],'MultipleIntuneCandidates','Multiple native Intune records share a correlation ID; all evidence is retained and selection is deterministic.','Information')
    emit('DimDevice', devices.values())
    hardware = index(list(read('hardware')), 'ManagedDeviceId')
    if set(hardware) != set(managed):
        raise ValueError('Detailed hardware and managed-device identity scope differ')
    for mid, row in hardware.items():
        if row['CollectionStatus'] != 'Collected':
            raise ValueError('Incomplete detailed hardware acquisition')
        if mid not in identity_conflicts and key(get(row,'azureADDeviceId')) != key(get(managed[mid],'AzureADDeviceId')):
            raise ValueError('Detailed hardware correlation differs from managed inventory')
        collected=utc(row['CollectedAtUtc'])
        receipt=proof[source_defs['hardware']['file']]
        if not utc(receipt['StartedAtUtc']) <= collected <= utc(receipt['CompletedAtUtc']):
            raise ValueError('Hardware observation outside producer acquisition interval')
    serials = collections.Counter(key(get(r,'serialNumber')) for r in hardware.values() if get(r,'serialNumber'))
    device_sources, bridges, hardware_rows, user_devices = [], [], [], []
    for system, records in [('Entra',entra.values()),('Intune',managed.values())]:
        for row in records:
            sid = get(row,'ObjectId','ManagedDeviceId')
            corr = key(get(row,'DeviceId','AzureADDeviceId'))
            conflict = system=='Intune' and key(sid) in identity_conflicts
            if system=='Intune':
                corr=managed_correlation(row)
            device = devices[corr]
            activity = get(row,'LastSyncDateTime','ApproximateLastSignInDateTime')
            av, ast = qualified(activity); ev, est = qualified(get(row,'EnrolledDateTime'))
            if activity and ast != 'UTC qualified':
                finding('Device',device['CmdbDeviceId'],'UnqualifiedActivityDate','Native device activity date has no established format/timezone.',discriminator=system+'|'+sid)
            chosen = system=='Entra' or key(selected[corr]['ManagedDeviceId'])==key(sid)
            device_sources.append({'TenantDeviceKey':device['TenantDeviceKey'], 'SourceRecordKey':native_key(tenant,'source',system,sid),
                'SourceSystem':system,'SourceObjectId':sid,'SourceDeviceId':corr,'DeviceName':device['DeviceName'],
                'ManagementAgent':get(row,'ManagementAgent'),'EnrollmentRaw':get(row,'EnrolledDateTime'),
                'EnrollmentUtcDateTime':ev,'EnrollmentStatus':est,'EnrollmentType':get(row,'DeviceEnrollmentType'),
                'ActivityKind':'Intune sync' if system=='Intune' else 'Approximate device sign-in',
                'ActivityRaw':activity,'ActivityUtcDateTime':av,'ActivityStatus':ast,
                'SourceCollectedDateTime':acquired('managed' if system=='Intune' else 'entra_devices'),
                'SelectionStatus':'Identity conflict' if conflict else 'Selected' if chosen else 'Other candidate','SelectionReason':'TenantKey + ManagedDeviceId; Entra/AD attribution indeterminate' if conflict else 'Native ID; latest sync/enrollment, then native ID'})
            if system!='Intune':
                continue
            bridges.append({'TenantIntuneDeviceKey':native_key(tenant,'intune',sid),'TenantDeviceKey':device['TenantDeviceKey'],'ManagedDeviceId':sid})
            detail = hardware[key(sid)]
            capacity, storage_status = byte_evidence(detail,'totalStorageSpaceInBytes')
            free, free_status = byte_evidence(detail,'freeStorageSpaceInBytes')
            memory, memory_status = byte_evidence(detail,'physicalMemoryInBytes')
            if capacity != '' and free != '' and capacity > 0 and free > capacity:
                raise ValueError('Detailed hardware free storage exceeds capacity')
            hardware_rows.append({'TenantDeviceKey':device['TenantDeviceKey'],'ManagedDeviceId':sid,
                'SerialNumber':get(detail,'serialNumber') or 'Not provided','Manufacturer':get(detail,'manufacturer') or 'Not provided',
                'Model':get(detail,'model') or 'Not provided','Storage':f'{capacity / 1073741824:.2f} GiB' if capacity else 'Unknown (reported 0)' if capacity==0 else 'Not provided',
                'PhysicalMemoryGiB':memory / 1073741824 if memory else '', 'StorageStatus':storage_status,'MemoryStatus':memory_status,
                'HardwareCollectedDateTime':get(detail,'CollectedAtUtc'),'InventoryCollectedDateTime':get(row,'CollectedAtUtc'),
                'CollectionCoverage':'Identity conflict; direct Intune hardware retained' if conflict else 'All native managed-device identities','CollectionMode':'Per-device GET with explicit hardware select'})
            for field in ('serialNumber','manufacturer','model'):
                if not get(detail,field):
                    finding('Device',device['CmdbDeviceId'],'HardwareAttributeMissing','Detailed hardware attribute not provided: '+field,discriminator=sid+'|'+field)
            if serials[key(get(detail,'serialNumber'))] > 1:
                finding('Device',device['CmdbDeviceId'],'RepeatedSerialValue','A reported serial appears on multiple native records; serial never defines identity.','Information',sid)
            if storage_status != 'Reported' or memory_status != 'Reported':
                finding('Device',device['CmdbDeviceId'],'HardwareCapacityUnknown','Missing and reported-zero storage/RAM remain distinct unknown evidence.',discriminator=sid)
            user = users_by_native.get(key(get(row,'UserId')))
            if chosen and user and not conflict:
                relationship = native_key(tenant,'primary-user',sid,user['SourceUserId'])
                user_devices.append({'TenantRelationshipKey':relationship,'TenantUserKey':user['TenantUserKey'],
                    'TenantDeviceKey':device['TenantDeviceKey'],'CmdbRelationshipId':relationship,
                    'CmdbUserId':user['CmdbUserId'],'CmdbDeviceId':device['CmdbDeviceId'],'RelationshipType':'PrimaryUser',
                    'SourceSystem':'Intune','Account':user['UserPrincipalName'],'Device':device['DeviceName'],
                    'DeviceSelection':device['DeviceSelection'],'DeviceCompliance':device['ComplianceState'],'DeviceSyncStatus':device['SyncStatus']})
    emit('DeviceSource',device_sources); emit('DimIntuneManagedDevice',bridges); emit('DeviceHardware',hardware_rows)
    emit('HardwareCoverage',[{'Status':'SuccessWithWarnings' if identity_conflicts else 'Validated','Coverage':'All native managed-device identities',
                             'Mode':'Per-device GET with explicit hardware select','RecordCount':len(hardware_rows)}])
    emit('FactUserDeviceRelationship',user_devices)
    emit('FactDeviceCompliance',({'TenantDeviceKey':d['TenantDeviceKey'],'CmdbDeviceId':d['CmdbDeviceId'],
         'ComplianceState':d['ComplianceState'],'LastSyncDateTime':d['SyncRaw'],'SourceSystem':d['SourceSystems']} for d in devices.values()))

    skus = index(list(read('skus')),'Id')
    emit('DimLicenseSku',({'TenantSkuKey':native_key(tenant,'sku',sid),'SkuId':sid,
         'SkuPartNumber':row['TenantSkuPartNumber'],'ConsumedUnits':typed_source(row,'TenantConsumedUnits','DimLicenseSku','ConsumedUnits','int64'),
         'EnabledUnits':typed_source(row,'TenantPrepaidEnabled','DimLicenseSku','EnabledUnits','int64')} for sid,row in skus.items()))
    paths, pairs, group_path_counts = [], collections.defaultdict(list), collections.Counter()
    for row in read('license_paths'):
        uid, sid, gid = key(row['UserId']), key(row['SkuId']), key(row['AssignedByGroupId'])
        user, sku, group = users_by_native.get(uid), skus.get(sid), groups.get(gid)
        if not uid or not sku:
            raise ValueError('License assignment SKU parent or native user identity missing')
        entity = 'User' if user else 'LicenseSourceUser'
        entity_id = user['CmdbUserId'] if user else 'license-source-user|' + uid
        updated, status = qualified(get(row,'LastUpdatedDateTime'))
        paths.append({'TenantAssignmentPathKey':native_key(tenant,'license-path',uid,sid,gid),
            'TenantUserKey':user['TenantUserKey'] if user else '', 'TenantSkuKey':native_key(tenant,'sku',sid),
            'TenantGroupKey':group['TenantGroupKey'] if group else '', 'SourceUserId':uid,'SkuId':sid,
            'AssignedByGroupId':gid,'Account':user['UserPrincipalName'] if user else '', 'Product':sku['TenantSkuPartNumber'],
            'AssignmentRoute':'Group' if gid else 'Direct','GroupName':group['DisplayName'] if group else 'Unresolved group' if gid else 'Direct assignment',
            'AssignmentState':row['AssignmentState'],'AssignmentError':row['AssignmentError'],
            'ErrorStatus':'Error reported' if key(row['AssignmentError']) not in ('','none') else 'No error reported' if row['AssignmentError'] else 'Not provided',
            'DisabledPlanIds':get(row,'DisabledPlanIds'),'AssignmentUpdatedRaw':get(row,'LastUpdatedDateTime'),
            'AssignmentUpdatedUtcDateTime':updated,'AssignmentUpdatedStatus':status,'SourceCollectedDateTime':acquired('license_paths'),
            'GroupLinkStatus':'Resolved' if group else 'Unresolved' if gid else 'Not applicable',
            'UserLinkStatus':'Resolved' if user else 'Unresolved','SkuLinkStatus':'Resolved'})
        if not user:
            finding(entity,entity_id,'UnresolvedLicenseUser','Native licensing user is absent from the independently collected user inventory; evidence retained without inferring deletion or fabricating a user.')
        if key(row['AssignmentError']) not in ('','none'):
            finding(entity,entity_id,'ObservedLicenseAssignmentError','Native license assignment reports an error; this does not prove non-use or overspend.',
                    discriminator=native_key(tenant,'license-path',uid,sid,gid), TenantGroupKey=group['TenantGroupKey'] if group else '')
        if gid and not group:
            finding(entity,entity_id,'UnresolvedAssignmentGroup','License assignment group is outside the collected WorkplaceScope cohort; membership is unknown, not empty.',discriminator=gid)
        pairs[uid,sid].append(row['AssignmentState']); group_path_counts[gid]+=1
    for gid, group in groups.items():
        group.update(ObservedPathCount=group_path_counts[gid],ObservedPathStatus='Observed paths' if group_path_counts[gid] else 'No observed license path')
    emit('DimGroup',groups.values()); emit('LicenseAssignmentPath',paths)
    memberships=[]
    for row in read('group_members'):
        group=groups[key(row['GroupId'])]; mid=key(row['MemberId']); kind=key(get(row,'MemberType'))
        if get(row,'CollectionStatus') != 'Collected':
            raise ValueError('Incomplete Entra member evidence')
        user=users_by_native.get(mid) if kind=='#microsoft.graph.user' else None
        ed=entra.get(mid) if kind=='#microsoft.graph.device' else None
        device=devices.get(key(ed['DeviceId'])) if ed else None
        nested=groups.get(mid) if kind=='#microsoft.graph.group' else None
        resolved=bool(user or device or nested)
        memberships.append({'TenantMembershipKey':native_key(tenant,'entra-member',row['GroupId'],mid),
            'TenantGroupKey':group['TenantGroupKey'],'TenantUserKey':user['TenantUserKey'] if user else '',
            'TenantDeviceKey':device['TenantDeviceKey'] if device else '', 'NestedTenantGroupKey':nested['TenantGroupKey'] if nested else '',
            'MemberId':row['MemberId'],'MemberType':get(row,'MemberType'),'MembershipKind':'Direct',
            'LinkStatus':'Resolved' if resolved else 'Outside collected entity scope','SourceCollectedDateTime':acquired('group_members')})
        if not resolved and kind in ('#microsoft.graph.user','#microsoft.graph.device','#microsoft.graph.group'):
            finding('Group',group['CmdbGroupId'],'UnresolvedGroupMember','Native direct member retained without a matching report entity.',discriminator=mid)
    emit('EntraGroupMembership',memberships)
    assignment_rows=[]
    state_rank={'Active':0,'ActiveWithError':1,'Error':2,'Disabled':3}
    for (uid,sid), states in pairs.items():
        state=min(states,key=lambda s:state_rank.get(s,4))
        user = users_by_native.get(uid)
        assignment_rows.append({'TenantUserKey':user['TenantUserKey'] if user else '',
            'TenantSkuKey':native_key(tenant,'sku',sid),'CmdbUserId':user['CmdbUserId'] if user else '',
            'SkuId':sid,'AssignmentState':state,'SourceSystem':'Entra',
            'SourceUserId':uid,'UserLinkStatus':'Resolved' if user else 'Unresolved'})
    emit('FactUserLicense',assignment_rows)
    plans={}
    for row in read('plans'):
        sid,pid=key(row['SkuId']),key(row['PlanId'])
        if sid not in skus:
            raise ValueError('Service-plan SKU missing')
        plans[sid,pid]=row
    emit('DimLicenseServicePlan',({'TenantServicePlanKey':native_key(tenant,'plan',sid,pid),
        'TenantSkuKey':native_key(tenant,'sku',sid),'SkuId':sid,'SkuPartNumber':get(row,'SkuPartNumber'),
        'ServicePlanId':pid,'ServicePlanName':row['PlanName'],'ProvisioningStatus':row['TenantProvisioningStatus'],
        'AppliesTo':get(row,'AppliesTo')} for (sid,pid),row in plans.items()))
    def service_plan_facts():
        state_codes={'A':('true','Success'),'D':('false','Disabled'),'PA':('true','PendingActivation'),
                     'PI':('true','PendingInput'),'PP':('true','PendingProvisioning'),'E':('true','Error')}
        for row in read('user_plans'):
            uid,sid,pid=key(row['UserId']),key(row['SkuId']),key(row['PlanId'])
            if not uid or (sid,pid) not in plans or (uid,sid) not in pairs:
                raise ValueError('User service-plan parent missing')
            state=row['StateCode']; decoded=state_codes.get(state)
            if decoded is None and state.startswith(('EN:','DIS:')):
                decoded=('true' if state.startswith('EN:') else 'false',state.split(':',1)[1])
            if decoded is None:
                raise ValueError('Unknown compact service-plan state code')
            user = users_by_native.get(uid)
            yield {'TenantUserKey':user['TenantUserKey'] if user else '', 'TenantSkuKey':native_key(tenant,'sku',sid),
                   'TenantServicePlanKey':native_key(tenant,'plan',sid,pid),'CmdbUserId':user['CmdbUserId'] if user else '',
                   'SkuId':sid,'ServicePlanId':pid,'ServicePlanName':plans[sid,pid]['PlanName'],
                   'IsEnabled':decoded[0],'AssignmentState':decoded[1],'SourceSystem':'Entra',
                   'SourceUserId':uid,'UserLinkStatus':'Resolved' if user else 'Unresolved'}
    emit('FactUserServicePlan',service_plan_facts())

    applications=index(list(read('apps')),'AppId'); product_devices=collections.defaultdict(set)
    application_devices=collections.Counter(); unresolved_app_devices=collections.Counter()
    product_versions=collections.defaultdict(set)
    product_names={}; app_product={}
    for aid,row in applications.items():
        if row['CollectionScope']!='AllPlatforms' or row['RelationCollectionScope']!='All':
            raise ValueError('All-platform, All-mode applications are required')
        product=tuple(key(row[field]) for field in ('AppName','AppPublisher','Platform'))
        app_product[aid]=product; product_versions[product].add(get(row,'AppVersion'))
        product_names.setdefault(product,tuple(row[field] for field in ('AppName','AppPublisher','Platform')))
    def installations():
        for row in read('app_relations'):
            aid,mid=key(row['AppId']),key(row['DeviceId'])
            application_devices[aid]+=1
            unresolved_app_devices[aid]+=mid not in managed
            if aid in applications:
                product_devices[app_product[aid]].add(mid)
            yield {'TenantDeviceApplicationKey':native_key(tenant,'app-device',aid,mid),
                   'TenantApplicationKey':native_key(tenant,'app',aid) if aid in applications else '',
                   'AppId':aid,'ManagedDeviceId':mid,
                   'ApplicationLinkStatus':'Resolved' if aid in applications else 'Unresolved',
                   'DeviceLinkStatus':'Resolved' if mid in managed else 'Unresolved',
                   'SourceCollectedDateTime':acquired('app_relations')}
    emit('FactDeviceApplication',installations())
    emit('DimDetectedApplication',({'TenantApplicationKey':native_key(tenant,'app',aid),'AppId':aid,
        'SourceApplicationKey':aid,'DisplayName':row['AppName'],'Version':row['AppVersion'],
        'Publisher':row['AppPublisher'],'Platform':row['Platform'],'DeviceCount':application_devices[aid],
        'ReportedDeviceCount':row['DeviceCount'],'ExactRelatedDeviceCount':application_devices[aid],
        'ResolvedDeviceCount':application_devices[aid]-unresolved_app_devices[aid],
        'UnresolvedDeviceCount':unresolved_app_devices[aid],
        'RelationshipCoverageStatus':application_coverage_status(int(row['DeviceCount']), application_devices[aid], unresolved_app_devices[aid]),
        'SourceCollectedDateTime':acquired('apps')} for aid,row in applications.items()))
    top=[]
    for product,names in product_names.items():
        name,publisher,platform=names
        top.append({'ApplicationProduct':' · '.join(names),'DisplayName':name,'Publisher':publisher,'Platform':platform,
                    'VersionCount':len(product_versions[product]),'ReportedDeviceCount':len(product_devices[product])})
    emit('TopApplication',sorted(top,key=lambda r:(-r['ReportedDeviceCount'],r['ApplicationProduct']))[:5])

    policies=list(read('policies')); policies_by_key={(key(r['PolicyFamily']),key(r['PolicyId'])):r for r in policies}
    for row in read('policy_assignments'):
        if (key(row['PolicyFamily']),key(row['PolicyId'])) not in policies_by_key:
            raise ValueError('Policy assignment parent missing')
        if not isinstance(json.loads(row['NativeTargetJson']),dict):
            raise ValueError('Invalid policy target payload')
    for row in policies:
        if row['AssignmentCollectionStatus']!='Collected' or not isinstance(json.loads(row['NativeEvidenceJson']),dict):
            raise ValueError('Incomplete policy assignment or native metadata')
    emit('DimIntuneConfigurationPolicy',({'TenantPolicyKey':native_key(tenant,'policy',r['PolicyFamily'],r['PolicyId']),
         'PolicyId':r['PolicyId'],'DisplayName':get(r,'DisplayName'),'Description':get(r,'Description'),
         'Platforms':get(r,'Platforms'),'Technologies':get(r,'Technologies'),
         'CreatedDateTime':typed_source(r,'CreatedDateTime','DimIntuneConfigurationPolicy'),'LastModifiedDateTime':typed_source(r,'LastModifiedDateTime','DimIntuneConfigurationPolicy'),
         'SourceCollectedDateTime':acquired('policies')} for r in policies if r['PolicyFamily'] not in ('WindowsFeatureUpdate','WindowsQualityUpdate')))
    emit('DimWindowsUpdatePolicy',({'TenantUpdatePolicyKey':native_key(tenant,'policy',r['PolicyFamily'],r['PolicyId']),
         'PolicyType':r['PolicyFamily'],'PolicyId':r['PolicyId'],'DisplayName':get(r,'DisplayName'),
         'TargetVersion':get(r,'FeatureUpdateVersion'),'CreatedDateTime':typed_source(r,'CreatedDateTime','DimWindowsUpdatePolicy'),
         'LastModifiedDateTime':typed_source(r,'LastModifiedDateTime','DimWindowsUpdatePolicy'),'SourceCollectedDateTime':acquired('policies')}
         for r in policies if r['PolicyFamily'] in ('WindowsFeatureUpdate','WindowsQualityUpdate')))

    mailbox_rows,hosting={},{}
    mailbox_entities,mailbox_observations,ambiguous_smtp = reconcile_mailboxes(
        {family:list(read(family)) for family in ('local_mailboxes','remote_mailboxes','mailboxes')})
    mailbox_keys = {mailbox_id:native_key(tenant,'mailbox',*mailbox_id) for mailbox_id in mailbox_entities}
    mailbox_identity_audit = []
    for mailbox_id,item in sorted(mailbox_observations.items()):
        target = mailbox_keys.get(item['target'], '')
        source_key = native_key(tenant,'mailbox-observation',*mailbox_id)
        mailbox_identity_audit.append({'SourceSystem':mailbox_id[0],'NativeId':mailbox_id[1],
            'TenantMailboxKey':target,'LinkBasis':item['basis'],'Issues':item['issues']})
        for issue in item['issues']:
            finding('Mailbox' if target else 'Mailbox recipient',target or source_key,issue,
                'Source='+mailbox_id[0]+'; native ID='+mailbox_id[1]+'; SMTP='+item['smtp']
                +'; '+item['basis']+'. Native recipient retained; SMTP is not an identity or proof of a physical mailbox.',
                discriminator=source_key)
    for mailbox_id,item in sorted(mailbox_entities.items()):
        family,r = mailbox_id[0],item['row']
        smtp=get(r,'PrimarySmtpAddress','PrimarySMTPaddress')
        recipient=get(r,'RecipientType') if family=='local_mailboxes' else get(r,'RecipientTypeDetails')
        external_id=get(r,'ExternalDirectoryObjectId')
        # Never override a supplied but unmatched technical user ID with SMTP.
        user=users_by_native.get(key(external_id)) if external_id else (
            user_by_address(smtp) if key(smtp) not in ambiguous_smtp else None)
        hkey=mailbox_keys[mailbox_id]
        mailbox_rows[mailbox_id]={'TenantMailboxKey':hkey,'CmdbMailboxId':hkey,'TenantUserKey':user['TenantUserKey'] if user else '',
            'CmdbUserId':user['CmdbUserId'] if user else '', 'RecipientTypeDetails':recipient,'ArchiveStatus':get(r,'ArchiveStatus','ArchiveState'),
            'SourceSystem':family,'DisplayName':get(r,'DisplayName'),'PrimarySmtpAddress':smtp,
            'LinkStatusLabel':'Resolved' if user else 'Unresolved','ExternalDirectoryObjectId':external_id}
        recipient_kind=key(recipient)
        hosting[mailbox_id]={'MailboxHostingKey':hkey,'CountryLabel':user['CountryLabel'] if user else UNKNOWN,
            'HostingLocation':'Exchange Online' if family=='mailboxes' else 'Exchange On-premises','RecipientTypeDetails':recipient,
            'MailboxTypeGroup':'Shared mailboxes' if 'sharedmailbox' in recipient_kind else 'User mailboxes' if recipient_kind=='usermailbox' else 'Other types',
            'EvidenceSource':family}
        if not smtp:
            finding('mailbox',hkey,'Missing SMTP','Native mailbox retained without an address; native identity does not depend on SMTP.')
    for m in mailbox_rows.values():
        if not m['TenantUserKey']:
            technical=key(m['RecipientTypeDetails'])=='discoverymailbox' and not m['ExternalDirectoryObjectId']
            finding('Mailbox',m['CmdbMailboxId'],'TechnicalMailboxWithoutUser' if technical else 'UnlinkedMailbox',
                    'Mailbox retained without a uniquely resolved Entra user link.', 'Information' if technical else 'Warning')
    emit('FactMailbox',mailbox_rows.values());emit('FactMailboxHosting',hosting.values())

    teams=index(list(read('teams')),'TeamId'); team_members=[]; team_member_counts=collections.Counter()
    for r in read('team_members'):
        tid,uid=key(r['TeamId']),key(r['UserId'])
        if tid not in teams:
            raise ValueError('Team membership parent missing')
        user=users_by_native.get(uid); team_member_counts[tid]+=1
        team_members.append({'TenantTeamMemberKey':native_key(tenant,'team-member',tid,uid,r['Role']),
            'TenantTeamKey':native_key(tenant,'team',tid),'TeamId':tid,'TenantUserKey':user['TenantUserKey'] if user else '',
            'UserId':uid,'UserPrincipalName':get(r,'UserPrincipalName'),'UserType':get(r,'UserType'),
            'Role':r['Role'],'SourceCollectedDateTime':acquired('team_members')})
    for tid,r in teams.items():
        if r['MemberCollectionStatus']!='Collected' or int(get(r,'MemberCount'))!=team_member_counts[tid]:
            raise ValueError('Team child-collection completion/count mismatch')
    emit('FactTeamMember',team_members)
    emit('DimTeam',({'TenantTeamKey':native_key(tenant,'team',tid),'TeamId':tid,'DisplayName':get(r,'TeamDisplayName'),
        'Visibility':get(r,'Visibility'),'CreatedDateTime':get(r,'CreatedDateTimeUtc'),
        'LastActivityDate':get(r,'LastActivityDateUtc'),'ActivityState':activity_state(get(r,'LastActivityDateUtc'), acquired('teams')),
        'OwnerCount':get(r,'OwnerCount'),'MemberCount':get(r,'MemberCount'),'GuestCount':get(r,'GuestCount'),
        'UnresolvedMemberCount':sum(not m['TenantUserKey'] for m in team_members if m['TeamId']==tid),
        'MembershipCoverageStatus':'Complete','IsArchived':boolean(get(r,'IsArchived')),'SourceCollectedDateTime':acquired('teams')} for tid,r in teams.items()))
    emit('DimSharePointSite',({'TenantSiteKey':native_key(tenant,'site',r['SiteId']),'SiteId':r['SiteId'],
        'SiteUrl':r['SiteUrl'],'SiteName':get(r,'Title'),'OwnerPrincipalName':get(r,'Owner'),
        'LastActivityDate':get(r,'LastActivityUtc'),'ActivityState':activity_state(get(r,'LastActivityUtc'), acquired('sites')),
        'StorageUsedBytes':number(get(r,'StorageUsedMB'))*1048576 if get(r,'StorageUsedMB') else '',
        'StorageAllocatedBytes':number(get(r,'StorageQuotaMB'))*1048576 if get(r,'StorageQuotaMB') else '',
        'RootWebTemplate':get(r,'Template'),'SourceCollectedDateTime':acquired('sites')} for r in read('sites')))
    activities=[]
    for r in read('activity'):
        user=user_by_address(r['UserPrincipalName'])
        fields={workload:workload+'LastActivityDate' for workload in
                ('Exchange','OneDrive','SharePoint','Teams','SkypeForBusiness','Yammer')}
        dates=[(r.get(field,''),workload) for workload,field in fields.items() if r.get(field,'')]
        for value,_ in dates:
            dt.date.fromisoformat(value)
        latest=max((value for value,_ in dates),default='')
        workloads=';'.join(sorted(workload for value,workload in dates if value==latest))
        item={'TenantUserActivityKey':native_key(tenant,'activity',r['UserPrincipalName']),
            'TenantUserKey':user['TenantUserKey'] if user else '', 'UserPrincipalName':r['UserPrincipalName'],
            'MatchStatus':'Resolved UPN' if user else 'Unresolved or ambiguous UPN','ReportRefreshDate':r['ReportRefreshDate'],
            'IsDeleted':boolean(get(r,'IsDeleted')),'LastActivityDate':latest,'LastActivityWorkload':workloads,
            'HasAnyM365Activity':'true' if dates else 'false','AssignedProducts':get(r,'AssignedProducts'),
            'SourceCollectedDateTime':acquired('activity')}
        for workload in ('Exchange','OneDrive','SharePoint','Teams'):
            field=fields[workload]
            item[workload+'LastActivityDate']=get(r,field)
        activities.append(item)
    emit('FactUserActivity',activities)
    emit('DimVerifiedDomain',({'TenantDomainKey':native_key(tenant,'domain',r['Id']),'DomainId':r['Id'],
        **{field:get(r,field) for field in ('IsDefault','IsInitial','AuthenticationType','SupportedServices','AvailabilityStatus')}} for r in read('domains')))
    emit('FactAutopilotDevice',({'TenantAutopilotDeviceKey':native_key(tenant,'autopilot',r['Autopilot ID']),
        'AutopilotDeviceId':r['Autopilot ID'],'DisplayName':get(r,'Display name'),'SerialNumber':get(r,'Serial number'),
        'Manufacturer':get(r,'Manufacturer'),'Model':get(r,'Model'),'GroupTag':get(r,'Group tag'),
        'EnrollmentState':get(r,'Enrollment state'),'LastContactedDateTime':typed_source(r,'Last contacted','FactAutopilotDevice','LastContactedDateTime'),
        'AzureAdDeviceId':get(r,'Azure AD Device ID'),'ManagedDeviceId':get(r,'Managed device ID'),
        'SourceCollectedDateTime':acquired('autopilot')} for r in read('autopilot')))
    analytics=[]
    unavailable_scores=collections.Counter()
    unavailable_devices=set()
    for r in read('analytics'):
        scores={}
        for field,source_field in [('EndpointAnalyticsScore','EndpointAnalyticsScore'),('StartupPerformanceScore','StartupScore'),('AppReliabilityScore','AppReliabilityScore'),('WorkFromAnywhereScore','WorkFromAnywhereScore')]:
            raw=get(r,source_field)
            scores[field]=endpoint_score(raw)
            if raw and scores[field]=='':
                unavailable_scores[(field,int(float(raw)))]+=1
                unavailable_devices.add(key(r['DeviceId']))
        analytics.append({'TenantEndpointAnalyticsDeviceKey':native_key(tenant,'analytics',r['ReportName'],r['DeviceId']),
            'DeviceId':r['DeviceId'],'DeviceName':get(r,'DeviceName'),'Manufacturer':get(r,'Manufacturer'),
            'Model':get(r,'Model'),**scores,'SourceCollectedDateTime':acquired('analytics'),'SourceSystem':r['ReportName']})
    emit('FactEndpointAnalyticsDevice',analytics)
    readiness=[]
    for r in read('readiness'):
        if r['UpgradeEligibility'] not in ('capable','notCapable','upgraded','unknown','unknownFutureValue'):
            raise ValueError('Unknown readiness evidence state')
        readiness.append({'TenantUpgradeEligibilityDeviceKey':native_key(tenant,'readiness',r['GraphId']),
            'MetricId':'Windows11Readiness','MetricDeviceId':r['GraphId'],'DeviceId':r['GraphId'],
            'DeviceIdSource':'Intune GraphId','DeviceName':get(r,'DeviceName'),'UpgradeEligibility':r['UpgradeEligibility'],
            'SourceCollectedDateTime':acquired('readiness'),'SourceSystem':'Intune'})
    emit('FactEndpointAnalyticsUpgradeEligibility',readiness)
    emit('FactWindowsUpdateAlert',({'TenantUpdateAlertKey':native_key(tenant,'update-alert',r['SourceReport'],r['DeviceId'],get(r,'PolicyId'),get(r,'EventDateUtc'),get(r,'AlertName')),
        'SourceReport':r['SourceReport'],'DeviceId':r['DeviceId'],'DeviceName':get(r,'DeviceName'),
        'PolicyId':get(r,'PolicyId'),'EventDateTimeUTC':get(r,'EventDateUtc'),'LastWUScanTimeUTC':get(r,'LastScanUtc'),
        'AggregateState':get(r,'AggregateState'),'CurrentDeviceUpdateStatus':get(r,'CurrentStatus'),
        'LatestAlertMessage':get(r,'AlertName'),'SourceCollectedDateTime':acquired('alerts'),'SourceSystem':'Intune'} for r in read('alerts')))

    ad_computers=list(read('ad_computers')); ad_users=list(read('ad_users')); coverage=[]
    ad_groups=list(read('ad_groups')); ad_objects=index(list(read('ad_objects')),'ObjectGUID')
    ad_dns=collections.defaultdict(list)
    for obj in ad_objects.values():
        ad_dns[key(obj['DistinguishedName'])].append(obj)
    if any(len(values)>1 for values in ad_dns.values()):
        raise ValueError('Ambiguous native AD distinguished name')
    def ad_object_by_dn(dn):
        matches=ad_dns.get(key(dn),[])
        return matches[0] if len(matches)==1 else None
    def ad_object_key(obj):
        return native_key(tenant,'ad-object',obj['ObjectGUID']) if obj else ''
    def sid_index(records, field):
        result=collections.defaultdict(list)
        for row in records:
            sid=key(get(row,field,'SID'))
            if sid: result[sid].append(row)
        return result
    ad_pc_sids=sid_index(ad_computers,'ObjectSID'); entra_sids=sid_index(entra.values(),'OnPremisesSecurityIdentifier')
    ad_user_sids=sid_index(ad_users,'ObjectSID'); raw_users=list(read('users'))
    cloud_user_sids=sid_index(raw_users,'OnPremisesSecurityIdentifier')
    def cloud_match(row, local, cloud):
        sid=key(get(row,'ObjectSID','SID'))
        if not sid: return None,'Native SID unavailable'
        if len(local[sid])>1 or len(cloud.get(sid,[]))>1: return None,'Ambiguous native SID'
        candidates=cloud.get(sid,[])
        return (candidates[0],'Unique native SID') if candidates else (None,'No exact cloud SID match')
    ad_user_links={}; ad_device_links={}
    for r in ad_computers:
        # Native synchronized SID match only. Names and enriched legacy booleans
        # cannot independently prove current Intune management.
        ed, match_status=cloud_match(r,ad_pc_sids,entra_sids)
        corr=key(ed['DeviceId']) if ed else ''
        md=selected.get(corr); device=devices.get(corr)
        state='Managed in Intune' if md else 'Entra only' if ed else 'No exact Entra SID match' if match_status=='No exact cloud SID match' else 'Native SID unavailable' if match_status=='Native SID unavailable' else 'Ambiguous Entra match'
        if corr in disputed_correlations and not md:
            state='Identity conflict; Intune coverage indeterminate'
            match_status+='; contested Intune identity excluded'
        ad_device_links[key(r['ObjectGUID'])]=(device,match_status)
        coverage.append({'TenantADComputerKey':native_key(tenant,'ad-computer',r['ObjectGUID']),
            'CmdbAdComputerId':'ad-computer|'+key(r['ObjectGUID']),'DeviceName':get(r,'Name'),
            'Enabled':boolean(get(r,'Enabled')),'OperatingSystem':r['OperatingSystem'],
            'OperatingSystemVersion':get(r,'OperatingSystemVersion'),'TenantDeviceKey':device['TenantDeviceKey'] if device else '',
            'EntraDeviceId':corr,'IntuneManagedDeviceId':get(md or {},'ManagedDeviceId'),'CoverageState':state,
            'MatchMethod':'Unique native SID' if ed else 'Unresolved native identity','SourceCollectedDateTime':acquired('ad_computers')})
    emit('FactADIntuneCoverage',coverage)
    def dates(row):
        last,status=filetime(get(row,'LastLogonTimestamp'))
        return {'LastLogonTimestampRaw':get(row,'LastLogonTimestamp'),'LastLogonUtcDateTime':last,'LastLogonStatus':status,
                **lifecycle_date(row, 'Creation', 'WhenCreated'),
                **lifecycle_date(row, 'Modification', 'WhenChanged'), **password_expiry(row)}
    ad_user_rows=[]
    for r in ad_users:
        cloud,status=cloud_match(r,ad_user_sids,cloud_user_sids)
        user=users_by_native.get(key(cloud['Object Id'])) if cloud else None
        ad_user_links[key(r['ObjectGUID'])]=(user,status)
        manager_dn=get(r,'manager'); manager=ad_object_by_dn(manager_dn) if manager_dn else None
        ad_user_rows.append({'TenantADUserKey':native_key(tenant,'ad-user',r['ObjectGUID']),
            **{field:get(r,field) for field in ('ObjectGUID','ObjectSID','DomainName','DistinguishedName','SamAccountName','UserPrincipalName','DisplayName','Department','Title','Country','AccountType')},
            'Enabled':boolean(get(r,'Enabled')),'ManagerDistinguishedName':manager_dn,'ManagerADObjectKey':ad_object_key(manager),
            'TenantUserKey':user['TenantUserKey'] if user else '', 'CloudMatchStatus':status,**dates(r),'SourceCollectedDateTime':acquired('ad_users')})
        if status=='Ambiguous native SID':
            finding('User' if user else 'ADUser',user['CmdbUserId'] if user else r['ObjectGUID'],'AmbiguousADCloudIdentity','A native SID is repeated; no cloud relationship is fabricated.')
    emit('ADUserSource',ad_user_rows)
    emit('ADComputerSource',({'TenantADComputerKey':native_key(tenant,'ad-computer',r['ObjectGUID']),
        **{field:get(r,field,'SID') if field=='ObjectSID' else get(r,field) for field in ('ObjectGUID','ObjectSID','DomainName','DistinguishedName','Name','DNSHostName','OperatingSystem','OperatingSystemVersion')},
        'Enabled':boolean(get(r,'Enabled')),'TenantDeviceKey':ad_device_links[key(r['ObjectGUID'])][0]['TenantDeviceKey'] if ad_device_links[key(r['ObjectGUID'])][0] else '',
        'CloudMatchStatus':ad_device_links[key(r['ObjectGUID'])][1],**dates(r),'SourceCollectedDateTime':acquired('ad_computers')} for r in ad_computers))
    emit('ADDirectoryObject',({'TenantADObjectKey':ad_object_key(r),
        **{field:get(r,field) for field in ('ObjectGUID','ObjectSID','DomainName','DistinguishedName','ObjectClass','PrimaryGroupID')},
        'SourceCollectedDateTime':acquired('ad_objects')} for r in ad_objects.values()))
    emit('ADDomainSource',({'TenantADDomainKey':native_key(tenant,'ad-domain',r['DNSRoot']),
        **{field:get(r,field) for field in ('DNSRoot','DomainSID','NetBIOSName','Forest','DistinguishedName','DomainMode','PDCEmulator','RIDMaster','InfrastructureMaster')},
        'SourceCollectedDateTime':acquired('ad_domains')} for r in read('ad_domains')))
    ad_group_index=index(ad_groups,'ObjectGUID')
    ad_group_sids=sid_index(ad_groups,'ObjectSID'); cloud_group_sids=sid_index(native_groups.values(),'OnPremisesSecurityIdentifier')
    ad_group_links={}
    for r in ad_groups:
        cloud,status=cloud_match(r,ad_group_sids,cloud_group_sids)
        ad_group_links[key(r['ObjectGUID'])]=(groups[key(cloud['GroupId'])] if cloud else None,status)
    emit('ADGroupSource',({'TenantADGroupKey':native_key(tenant,'ad-group',r['ObjectGUID']),
        **{field:get(r,field) for field in ('ObjectGUID','ObjectSID','DomainName','DistinguishedName','Name','DisplayName','GroupCategory','GroupScope','ManagedBy')},
        'ManagedByADObjectKey':ad_object_key(ad_object_by_dn(get(r,'ManagedBy'))) if get(r,'ManagedBy') else '',
        'TenantGroupKey':ad_group_links[key(r['ObjectGUID'])][0]['TenantGroupKey'] if ad_group_links[key(r['ObjectGUID'])][0] else '',
        'CloudMatchStatus':ad_group_links[key(r['ObjectGUID'])][1],
        'SourceCollectedDateTime':acquired('ad_groups')} for r in ad_groups))
    memberships=[]; observed_members=collections.defaultdict(set)
    for r in read('ad_members'):
        group=ad_group_index.get(key(get(r,'GroupObjectGUID')))
        unresolved_primary=r['ResolutionStatus']=='UnresolvedPrimaryGroup' and r['MembershipKind']=='PrimaryGroup' and not get(r,'GroupObjectGUID')
        if (not group and not unresolved_primary) or (group and key(get(group,'ObjectSID'))!=key(r['GroupSID'])):
            raise ValueError('AD membership group identity is missing or inconsistent')
        member=ad_objects.get(key(get(r,'MemberObjectGUID'))) if get(r,'MemberObjectGUID') else None
        if r['ResolutionStatus'] in ('Resolved','UnresolvedPrimaryGroup'):
            if not member or key(member['DistinguishedName'])!=key(r['MemberDistinguishedName']) or key(get(member,'ObjectSID'))!=key(get(r,'MemberSID')):
                raise ValueError('AD membership resolved identity is inconsistent')
        elif r['ResolutionStatus']!='UnresolvedOrExternal' or member:
            raise ValueError('Unresolved AD membership carries a resolved object identity')
        if r['MembershipKind'] not in ('Direct','PrimaryGroup'):
            raise ValueError('Unknown AD membership kind')
        if member and key(get(member,'ObjectClass'))!=key(get(r,'MemberObjectClass')):
            raise ValueError('AD member object class differs from native object evidence')
        if r['MembershipKind']=='PrimaryGroup' and (not member or not get(member,'PrimaryGroupID') or
            re.sub(r'-\d+$','-'+get(member,'PrimaryGroupID'),get(member,'ObjectSID')) != r['GroupSID']):
            raise ValueError('AD primary group differs from native SID/RID evidence')
        if r['MembershipKind']=='Direct': observed_members[key(group['ObjectGUID'])].add(key(r['MemberDistinguishedName']))
        uid=key(get(member or {},'ObjectGUID'))
        cloud_group,_=ad_group_links.get(key(get(group or {},'ObjectGUID')),(None,''))
        user,user_status=ad_user_links.get(uid,(None,'Outside collected AD user scope'))
        device,device_status=ad_device_links.get(uid,(None,'Outside collected AD workstation scope'))
        memberships.append({'TenantADMembershipKey':native_key(tenant,'ad-membership',get(r,'GroupObjectGUID'),r['GroupSID'],r['MemberDistinguishedName'],r['MembershipKind']),
            'TenantADGroupKey':native_key(tenant,'ad-group',group['ObjectGUID']) if group else '', 'TenantADObjectKey':ad_object_key(member),
            'TenantGroupKey':cloud_group['TenantGroupKey'] if cloud_group else '',
            'TenantUserKey':user['TenantUserKey'] if user else '', 'TenantDeviceKey':device['TenantDeviceKey'] if device else '',
            **{field:get(r,field) for field in ('GroupObjectGUID','GroupSID','MemberDistinguishedName','MemberObjectGUID','MemberSID','MemberObjectClass','MembershipKind','ResolutionStatus')},
            'CloudMatchStatus':user_status if uid in ad_user_links else device_status,'SourceCollectedDateTime':acquired('ad_members')})
    for gid,r in ad_group_index.items():
        payload=json.loads(r['MembersJson'])
        if not isinstance(payload,list) or any(not isinstance(v,str) or not v.strip() for v in payload):
            raise ValueError('Invalid native AD members payload')
        if {key(v) for v in payload} != observed_members[gid]:
            raise ValueError('AD direct membership evidence differs from MembersJson')
    emit('ADMembership',memberships)
    hybrid=[]
    for entity,ad,cloud,ad_field,cloud_field in [('User',ad_users,list(read('users')),'ObjectSID','OnPremisesSecurityIdentifier'),('Device',ad_computers,list(entra.values()),'ObjectSID','OnPremisesSecurityIdentifier')]:
        ad_keys=collections.Counter(key(get(r,ad_field,'SID')) for r in ad if get(r,ad_field,'SID'))
        cloud_keys=collections.Counter(key(get(r,cloud_field)) for r in cloud if get(r,cloud_field))
        matched={v for v in ad_keys if ad_keys[v]==1 and cloud_keys[v]==1}
        hybrid.append({'EntityType':entity,'OnPremisesCount':len(ad),'CloudCount':len(cloud),'MatchedCount':len(matched),
            'OnPremisesOnlyCount':'','CloudOnlyCount':'',
            'DuplicateOnPremisesKeyCount':sum(v>1 for v in ad_keys.values()),'DuplicateCloudKeyCount':sum(v>1 for v in cloud_keys.values()),
            'MatchRatio':len(matched)/len(ad) if ad else '', 'MatchMethod':'Unique native SID','SourceCollectedDateTime':acquired('ad_users' if entity=='User' else 'ad_computers')})
    emit('FactHybridIdentityCoverage',hybrid)
    emit('FactDataQuality',findings.values())
    entity_index={key(d['CmdbDeviceId']):{'TenantDeviceKey':d['TenantDeviceKey']} for d in devices.values()}
    entity_index.update({key(u['CmdbUserId']):{'TenantUserKey':u['TenantUserKey']} for u in users})
    entity_index.update({key(g['CmdbGroupId']):{'TenantGroupKey':g['TenantGroupKey']} for g in groups.values()})
    emit('EntityFinding',({**f,**entity_index.get(key(f['EntityId']),{}),
                          'LinkStatus':'Resolved' if key(f['EntityId']) in entity_index else 'Outside 360 entity scope'} for f in findings.values()))
    emit('DimCountry',({'CountryLabel':country} for country in sorted({UNKNOWN}|{u['CountryLabel'] for u in users}|{h['CountryLabel'] for h in hosting.values()})))
    emit('DimTenant',[{'TenantDisplayName':tenant,'Environment':identity['EnvironmentKey'],'LastRefreshDateTime':reference.isoformat()}])
    dates=[]; day=dt.date(reference.year,1,1)
    while day.year==reference.year:
        dates.append({'Date':day.isoformat(),'Year':day.year,'Quarter':(day.month-1)//3+1,'Month':day.month,'MonthName':day.strftime('%B'),'Day':day.day})
        day+=dt.timedelta(days=1)
    emit('DimDate',dates)
    app_files = {source_defs[name]['file'] for name in ('apps', 'app_relations')}
    app_coverage = evidence['ApplicationCoverage']
    license_counts = {source_defs['license_paths']['file']:evidence['LicenseCoverage']['UnresolvedAssignmentRows'],
                      source_defs['user_plans']['file']:evidence['LicenseCoverage']['UnresolvedServicePlanRows']}
    emit('SourceHealth',({'SourceName':r['File'],'Status':('SuccessWithWarnings' if r.get('DomainCoverage') or
        license_counts.get(r['File'], 0) or (identity_conflicts and r['File'] in {source_defs[n]['file'] for n in ('managed','hardware')}) or
        (r['File'] in app_files and (evidence['ApplicationWarnings'] or freshness[r['File']]['State']=='Aging')) else r['Status']),
        'Coverage':('Identity conflict; disputed Entra/AD attribution excluded; count=' + str(len(identity_conflicts)) if identity_conflicts and r['File'] in {source_defs[n]['file'] for n in ('managed','hardware')} else 'Partial AD coverage; unavailable domains: ' + ', '.join(r['DomainCoverage']['UnavailableDomains']) if r.get('DomainCoverage') else
                    'Weekly application evidence; unresolved device relations=' + str(app_coverage['UnresolvedDeviceRelationRows'])
                    + '; unresolved application relations=' + str(app_coverage['UnresolvedApplicationRelationRows'])
                    + '; application count mismatches=' + str(app_coverage['CountMismatchApplications'])
                    if r['File'] in app_files else 'Complete producer file scope; unresolved native licensing user links=' + str(license_counts[r['File']])
                    if license_counts.get(r['File'], 0) else 'Complete producer file scope'),
        'SourceRows':r['Rows'],'MaxItems':0,'StartedDateTime':r['StartedAtUtc'],'CompletedDateTime':r['CompletedAtUtc'],
        'Evidence':('Native producer success; exact hash and logical row count; acquisition age '
                    + str(freshness[r['File']]['AgeHours']) + 'h; '
                    + freshness[r['File']]['State'] + '; rejection after '
                    + str(freshness[r['File']]['MaxAgeHours']) + 'h')} for r in evidence['Files']))
    relationships=[('PrimaryUser',len(user_devices),'Native Intune primary user'),('HasMailbox',sum(bool(m['TenantUserKey']) for m in mailbox_rows.values()),'Unique native ID or SMTP/user match'),
        ('AssignedLicense',len(assignment_rows),'Native Entra user/SKU pairs, including explicitly unresolved user links'),('MemberOfGroup',sum(member_counts.values()),'Direct Entra memberships, including non-user objects'),
        ('DeviceHasApplication',table_counts['FactDeviceApplication'],'Weekly native app/device observations; includes qualified unresolved links'),('DeviceInAutopilot',sum(bool(get(r,'Managed device ID')) for r in read('autopilot')),'Native Autopilot managed device link'),
        ('ADMemberOfGroup',sum(1 for _ in read('ad_members')),'Native direct and primary-group evidence'),
        ('PolicyAssignmentTarget',sum(1 for _ in read('policy_assignments')),'Configured target; not proof of effective device assignment')]
    emit('FactRelationshipOverview',({'RelationshipType':name,'RelationshipCount':count,'EvidenceSource':text} for name,count,text in relationships))
    if set(table_counts)!=set(definitions):
        raise ValueError('Builder did not cover every reporting contract table')
    # Qualification plus native mailbox identity lineage; raw rows stay in hashed sources.
    return {'DeviceIdentity':identity_summary, 'ApplicationCoverage':app_coverage, 'LicenseCoverage':evidence['LicenseCoverage'],
        'MailboxIdentity': {'Observations':mailbox_identity_audit,
            'ObservedMailboxEntities':len(mailbox_entities),
            'UnresolvedRemoteRecipients':sum(i['identity'][0]=='remote_mailboxes' and i['target'] is None for i in mailbox_observations.values()),
            'Policy':'Native mailbox keys; unique technical cross-source links only. Remote recipients without an observed mailbox are findings, not extra mailbox counts. All source identities are retained in this audit.'},
        'ADDomainCoverage':[r['DomainCoverage'] for r in evidence['ProducerReceipts'] if r.get('DomainCoverage')], 'EntraGroupCatalogComparison': {
        'CatalogSource': source_defs['group_scope']['file'],
        'CatalogGroups': len(native_groups),
        'LicensingCatalogSource': source_defs['groups']['file'],
        'LicensingCatalogGroups': len(licensing_groups),
        'OnlyInMembershipCatalog': len(set(native_groups) - set(licensing_groups)),
        'OnlyInLicensingCatalog': len(set(licensing_groups) - set(native_groups)),
        'Policy': 'Group properties and direct memberships use one WorkplaceScope cohort; independent licensing catalog differences do not imply empty membership'},
        'EndpointAnalyticsUnavailableScores': {
        'DistinctDevices':len(unavailable_devices),
        'ScoreCells':sum(unavailable_scores.values()),
        'ByFieldAndSentinel':[{'Field':field,'Sentinel':sentinel,'Count':count}
            for (field,sentinel),count in sorted(unavailable_scores.items())],
        'Policy':'Documented -1/-2 sentinels mapped to blank; device rows retained; raw sources unchanged'}}
