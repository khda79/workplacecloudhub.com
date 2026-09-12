"""Additive V1 report data preparation. Local CSV only; no collection.

The summary contracts and their grain are preserved. Unknown timestamps are
kept as raw text and never parsed using the host's locale or assumed timezone.
"""
import datetime as dt
import json
from collections import Counter
from pathlib import Path

EXTRA_RELATIONSHIPS = [
    ('DeviceSource', 'TenantDeviceKey', 'DimDevice', 'TenantDeviceKey'),
    ('LicenseAssignmentPath', 'TenantUserKey', 'DimUser', 'TenantUserKey'),
    ('LicenseAssignmentPath', 'TenantSkuKey', 'DimLicenseSku', 'TenantSkuKey'),
    ('LicenseAssignmentPath', 'TenantGroupKey', 'DimGroup', 'TenantGroupKey'),
    ('EntityFinding', 'TenantDeviceKey', 'DimDevice', 'TenantDeviceKey'),
    ('EntityFinding', 'TenantUserKey', 'DimUser', 'TenantUserKey'),
    ('EntityFinding', 'TenantGroupKey', 'DimGroup', 'TenantGroupKey'),
]


def key(value):
    return (value or '').strip().casefold()


def timestamp(value):
    """Return qualified UTC text and an explicit status; keep raw elsewhere."""
    if not value or not value.strip():
        return '', 'Not provided'
    try:
        parsed = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
        if parsed.tzinfo is None:
            return '', 'Timezone not established'
        if parsed.year < 1900:
            return '', 'Sentinel / unknown'
        return parsed.astimezone(dt.timezone.utc).isoformat(), 'UTC qualified'
    except (ValueError, OverflowError):
        return '', 'Format or timezone not established'


def unique(rows, column, name):
    result = {}
    for row in rows:
        k = key(row[column])
        if not k or k in result:
            raise ValueError('Blank or duplicate 360 key: ' + name + '.' + column)
        result[k] = row
    return result


def enrich(root, identity, data, columns, hashes, read_csv, sha, product):
    contracts = {}
    for file in ('SmartWorkplaceCMDB.tables.json', 'SmartWorkplaceCMDB.raw.tables.json'):
        for t in json.loads((product/'Schema'/file).read_text(encoding='utf-8-sig'))['tables']:
            contracts[Path(t['name']).stem] = t

    def load(name, pk):
        t = contracts[name]
        path = root.joinpath(*t['area'].replace('\\', '/').split('/'), t['name'])
        if not path.is_file():
            raise ValueError('360 requires existing local source: ' + name)
        before = sha(path)
        cs, rows = read_csv(path)
        if cs != t['columns']:
            raise ValueError('360 source contract mismatch: ' + name)
        if any(any(r[k] != v for k, v in identity.items()) for r in rows):
            raise ValueError('360 source tenant mismatch: ' + name)
        lookup = unique(rows, pk, name)
        if sha(path) != before:
            raise ValueError('360 source changed during read: ' + name)
        hashes[str(path.relative_to(root))] = before
        return rows, lookup

    users, user_cmdb = load('CMDB_Users', 'CmdbUserId')
    devices, device_cmdb = load('CMDB_Devices', 'CmdbDeviceId')
    groups, group_cmdb = load('CMDB_Groups', 'CmdbGroupId')
    entra, _ = load('Entra_Devices', 'SourceObjectId')
    intune, _ = load('Intune_ManagedDevices', 'ManagedDeviceId')
    assignments, _ = load('M365_UserLicenseAssignments', 'RawAssignmentKey')
    findings, _ = load('CMDB_DataQuality', 'FindingId')

    def add_fields(name, fields, enrich_row):
        if any(c in columns[name] for c in fields):
            raise ValueError('360 column collision: ' + name)
        columns[name].extend(fields)
        for r in data[name]:
            extra = enrich_row(r)
            if set(extra) != set(fields):
                raise ValueError('360 detail field mismatch: ' + name)
            r.update(extra)

    def require_same(parent, name, c):
        vals = unique(data[name], c, name)
        if set(parent) != set(vals):
            raise ValueError('360 entity set mismatch: ' + name)

    for parent, name, c in [(user_cmdb, 'DimUser', 'CmdbUserId'),
                            (device_cmdb, 'DimDevice', 'CmdbDeviceId'),
                            (group_cmdb, 'DimGroup', 'CmdbGroupId')]:
        require_same(parent, name, c)

    def user_detail(r):
        c = user_cmdb[key(r['CmdbUserId'])]
        utc, status = timestamp(c['CreatedDateTime'])
        return dict(UserSelection=f"{r['DisplayName']} | {r['UserPrincipalName']} | {c['SourceUserId']}",
                    SourceUserId=c['SourceUserId'], CreationRaw=c['CreatedDateTime'],
                    CreationUtcDateTime=utc, CreationStatus=status,
                    SourceCollectedDateTime=c['SourceCollectedDateTime'],
                    ManagerStatus='Not collected', UserActivityStatus='Not collected')
    add_fields('DimUser', ['UserSelection','SourceUserId','CreationRaw','CreationUtcDateTime','CreationStatus',
                          'SourceCollectedDateTime','ManagerStatus','UserActivityStatus'], user_detail)
    user_by_source = {key(c['SourceUserId']):r for r in data['DimUser'] for c in [user_cmdb[key(r['CmdbUserId'])]]}
    if len(user_by_source) != len(users) or '' in user_by_source:
        raise ValueError('Ambiguous 360 source user mapping')

    def device_detail(r):
        c = device_cmdb[key(r['CmdbDeviceId'])]
        u = user_by_source.get(key(c['PrimaryUserId']))
        if u:
            country_code = u['CountryCode']
            country_label = u['CountryLabel'] or 'Unknown / unassigned'
            country_status = u['CountryStatus'] or 'Not provided'
        elif c['PrimaryUserId']:
            country_code = ''
            country_label = 'Unknown / unassigned'
            country_status = 'User unresolved'
        else:
            country_code = ''
            country_label = 'Unknown / unassigned'
            country_status = 'No primary user'
        utc, status = timestamp(c['LastSyncDateTime'])
        return dict(DeviceSelection=f"{r['DeviceName']} | {c['SourceDeviceId']}",
                    SourceDeviceId=c['SourceDeviceId'], SourceSystems=c['SourceSystem'],
                    AssociatedAccount=u['UserPrincipalName'] if u else 'Not resolved' if c['PrimaryUserId'] else 'Not provided',
                    AssociationStatus='Resolved' if u else 'Unresolved' if c['PrimaryUserId'] else 'Not provided',
                    CountryCode=country_code, CountryLabel=country_label, CountryStatus=country_status,
                    PrimaryUserSourceId=c['PrimaryUserId'], SyncRaw=c['LastSyncDateTime'],
                    SyncUtcDateTime=utc, SyncStatus=status, SourceCollectedDateTime=c['SourceCollectedDateTime'])
    add_fields('DimDevice', ['DeviceSelection','SourceDeviceId','SourceSystems','AssociatedAccount','AssociationStatus',
                            'CountryCode','CountryLabel','CountryStatus','PrimaryUserSourceId','SyncRaw',
                            'SyncUtcDateTime','SyncStatus','SourceCollectedDateTime'], device_detail)
    device_by_source = {key(c['SourceDeviceId']):r for r in data['DimDevice'] for c in [device_cmdb[key(r['CmdbDeviceId'])]]}
    if len(device_by_source) != len(devices) or '' in device_by_source:
        raise ValueError('Ambiguous 360 source device mapping')

    def group_detail(r):
        c = group_cmdb[key(r['CmdbGroupId'])]
        return dict(GroupSelection=f"{r['DisplayName']} | {c['SourceGroupId']}", SourceGroupId=c['SourceGroupId'],
                    MemberStatus='Not collected', OwnerStatus='Not collected', SourceCollectedDateTime=c['SourceCollectedDateTime'])
    add_fields('DimGroup', ['GroupSelection','SourceGroupId','MemberStatus','OwnerStatus','SourceCollectedDateTime'], group_detail)
    group_by_source = {key(r['SourceGroupId']):r for r in data['DimGroup']}
    if len(group_by_source) != len(groups) or '' in group_by_source:
        raise ValueError('Ambiguous 360 source group mapping')
    sku_by_source = {key(r['SkuId']):r for r in data['DimLicenseSku']}
    if len(sku_by_source) != len(data['DimLicenseSku']) or '' in sku_by_source:
        raise ValueError('Ambiguous 360 source SKU mapping')

    # Keep EVERY source row, including repeated Intune correlation candidates.
    # The CMDB-selected candidate is informational; do not re-normalize devices.
    candidates = {}
    zero = '00000000-0000-0000-0000-000000000000'
    for r in intune:
        corr = key(r['AzureAdDeviceId'])
        if not corr or corr == zero:
            corr = 'intune:' + key(r['ManagedDeviceId'])
        candidates.setdefault(corr, []).append(r)
    def rank_time(text):
        try:
            value = dt.datetime.fromisoformat(text.replace('Z', '+00:00'))
            if value.tzinfo is None: raise ValueError()
            return value.astimezone(dt.timezone.utc)
        except (ValueError, OverflowError):
            raise ValueError('Unqualified Intune candidate timestamp; selection cannot be verified')
    selected = {}
    for corr, rs in candidates.items():
        # Same tie-break ordering as the existing PowerShell normalizer.
        ordered = sorted(rs, key=lambda r: key(r['ManagedDeviceId']))
        ordered.sort(key=lambda r: rank_time(r['EnrolledDateTime']) if r['EnrolledDateTime'] else dt.datetime.min.replace(tzinfo=dt.timezone.utc), reverse=True)
        ordered.sort(key=lambda r: rank_time(r['LastSyncDateTime']) if r['LastSyncDateTime'] else dt.datetime.min.replace(tzinfo=dt.timezone.utc), reverse=True)
        selected[corr] = ordered[0]
    source_fields = ['TenantDeviceKey','SourceRecordKey','SourceSystem','SourceObjectId','SourceDeviceId','DeviceName',
                     'ManagementAgent','EnrollmentRaw','EnrollmentUtcDateTime','EnrollmentStatus','EnrollmentType',
                     'ActivityKind','ActivityRaw','ActivityUtcDateTime','ActivityStatus','SourceCollectedDateTime',
                     'SelectionStatus','SelectionReason']
    sources = []
    for system, records in [('Entra', entra), ('Intune', intune)]:
        for r in records:
            source_id = r['SourceObjectId'] if system == 'Entra' else r['ManagedDeviceId']
            corr = key(r['SourceDeviceId'] if system == 'Entra' else r['AzureAdDeviceId'])
            if system == 'Intune' and (not corr or corr == zero): corr = 'intune:' + key(source_id)
            device = device_by_source.get(corr)
            if not device: raise ValueError('Unmapped 360 device source record')
            activity = r['ApproximateLastSignInDateTime'] if system == 'Entra' else r['LastSyncDateTime']
            utc, status = timestamp(activity)
            enrolled = r.get('EnrolledDateTime', '')
            eu, es = timestamp(enrolled)
            chosen = system == 'Entra' or key(selected[corr]['ManagedDeviceId']) == key(source_id)
            if system == 'Intune' and chosen:
                c = device_cmdb[key(device['CmdbDeviceId'])]
                if key(c['PrimaryUserId']) != key(r['UserId']) or c['LastSyncDateTime'] != r['LastSyncDateTime']:
                    raise ValueError('Intune selected source differs from curated device')
            sources.append(dict(identity, TenantDeviceKey=device['TenantDeviceKey'],
                SourceRecordKey=identity['TenantKey']+'|'+system+'|'+source_id, SourceSystem=system,
                SourceObjectId=source_id, SourceDeviceId=r.get('SourceDeviceId', r.get('AzureAdDeviceId','')),
                DeviceName=r['DeviceName'], ManagementAgent=r.get('ManagementAgent','Not applicable'),
                EnrollmentRaw=enrolled, EnrollmentUtcDateTime=eu, EnrollmentStatus=es if system=='Intune' else 'Not applicable',
                EnrollmentType=r.get('DeviceEnrollmentType','Not applicable'), ActivityKind='Approximate device sign-in' if system=='Entra' else 'Intune sync',
                ActivityRaw=activity, ActivityUtcDateTime=utc, ActivityStatus=status,
                SourceCollectedDateTime=r['SourceCollectedDateTime'], SelectionStatus='Selected' if chosen else 'Other candidate',
                SelectionReason='Entra identity baseline' if system=='Entra' else 'Latest sync, latest enrollment, then source ID'))

    paths = []
    for a in assignments:
        u, g, s = user_by_source.get(key(a['SourceUserId'])), group_by_source.get(key(a['AssignedByGroupId'])), sku_by_source.get(key(a['SkuId']))
        utc, status = timestamp(a['LastUpdatedDateTime'])
        paths.append(dict(identity, TenantAssignmentPathKey=identity['TenantKey']+'|path|'+a['RawAssignmentKey'],
            TenantUserKey=u['TenantUserKey'] if u else '', TenantGroupKey=g['TenantGroupKey'] if g else '',
            TenantSkuKey=s['TenantSkuKey'] if s else '', SourceUserId=a['SourceUserId'], AssignedByGroupId=a['AssignedByGroupId'],
            SkuId=a['SkuId'], Account=u['UserPrincipalName'] if u else 'Unresolved user',
            Product=s['SkuPartNumber'] if s else 'Unresolved SKU',
            AssignmentRoute='Group' if a['AssignedByGroupId'] else 'Direct',
            GroupName=g['DisplayName'] if g else 'Unresolved group' if a['AssignedByGroupId'] else 'Direct assignment',
            AssignmentState=a['AssignmentState'], AssignmentError=a['AssignmentError'],
            ErrorStatus='Error reported' if key(a['AssignmentError']) not in ('', 'none') else 'No error reported' if a['AssignmentError'] else 'Not provided',
            DisabledPlanIds=a['DisabledPlanIds'], AssignmentUpdatedRaw=a['LastUpdatedDateTime'],
            AssignmentUpdatedUtcDateTime=utc, AssignmentUpdatedStatus=status, SourceCollectedDateTime=a['SourceCollectedDateTime'],
            GroupLinkStatus='Resolved' if g else 'Unresolved' if a['AssignedByGroupId'] else 'Not applicable',
            UserLinkStatus='Resolved' if u else 'Unresolved', SkuLinkStatus='Resolved' if s else 'Unresolved'))
    observed = Counter(key(r['TenantGroupKey']) for r in paths if r['TenantGroupKey'])
    add_fields('DimGroup',['ObservedPathCount','ObservedPathStatus'],
               lambda r:dict(ObservedPathCount=str(observed[key(r['TenantGroupKey'])]),
                             ObservedPathStatus='Observed license paths' if observed[key(r['TenantGroupKey'])] else 'No observed license path'))
    raw_pairs = {(key(r['TenantUserKey']), key(r['TenantSkuKey'])) for r in paths}
    summary_pairs = {(key(r['TenantUserKey']), key(r['TenantSkuKey'])) for r in data['FactUserLicense']}
    if raw_pairs != summary_pairs:
        raise ValueError('Assignment paths do not reconcile with summary user/SKU pairs')

    lookup_types = {'user': ('TenantUserKey', {key(r['CmdbUserId']):r['TenantUserKey'] for r in data['DimUser']}),
                    'device': ('TenantDeviceKey', {key(r['CmdbDeviceId']):r['TenantDeviceKey'] for r in data['DimDevice']}),
                    'group': ('TenantGroupKey', {key(r['CmdbGroupId']):r['TenantGroupKey'] for r in data['DimGroup']})}
    entity_findings = []
    for f in findings:
        row = dict(identity, TenantFindingKey=identity['TenantKey']+'|finding|'+f['FindingId'],
                   TenantUserKey='', TenantDeviceKey='', TenantGroupKey='', EntityType=f['EntityType'],
                   EntityId=f['EntityId'], Severity=f['Severity'], FindingType=f['FindingType'], Description=f['Description'],
                   RecommendedAction=f['RecommendedAction'], DetectedDateTime=f['DetectedDateTime'], LinkStatus='Outside 360 entity scope')
        if key(f['EntityType']) in lookup_types:
            field, lookup = lookup_types[key(f['EntityType'])]
            row[field] = lookup.get(key(f['EntityId']), '')
            row['LinkStatus'] = 'Resolved' if row[field] else 'Unresolved entity'
        entity_findings.append(row)
    new_tables = [('DeviceSource',sources,source_fields), ('LicenseAssignmentPath',paths,
        ['TenantAssignmentPathKey','TenantUserKey','TenantGroupKey','TenantSkuKey','SourceUserId','AssignedByGroupId','SkuId',
         'Account','Product','AssignmentRoute','GroupName','AssignmentState','AssignmentError','ErrorStatus','DisabledPlanIds',
         'AssignmentUpdatedRaw','AssignmentUpdatedUtcDateTime','AssignmentUpdatedStatus','SourceCollectedDateTime',
         'GroupLinkStatus','UserLinkStatus','SkuLinkStatus']),
        ('EntityFinding',entity_findings,['TenantFindingKey','TenantUserKey','TenantDeviceKey','TenantGroupKey','EntityType','EntityId',
                                        'Severity','FindingType','Description','RecommendedAction','DetectedDateTime','LinkStatus'])]
    for name, rows, fs in new_tables:
        data[name], columns[name] = rows, list(identity)+fs
    for name, pk in [('DeviceSource','SourceRecordKey'),('LicenseAssignmentPath','TenantAssignmentPathKey'),('EntityFinding','TenantFindingKey')]:
        unique(data[name],pk,name)
    # Existing associations gain display context; underlying keys remain intact.
    du = {key(r['TenantUserKey']):r for r in data['DimUser']}
    dd = {key(r['TenantDeviceKey']):r for r in data['DimDevice']}
    add_fields('FactUserDeviceRelationship',['Account','Device','DeviceSelection','DeviceCompliance','DeviceSyncStatus'],
               lambda r:dict(Account=du[key(r['TenantUserKey'])]['UserPrincipalName'], Device=dd[key(r['TenantDeviceKey'])]['DeviceName'],
                             DeviceSelection=dd[key(r['TenantDeviceKey'])]['DeviceSelection'],
                             DeviceCompliance=dd[key(r['TenantDeviceKey'])]['ComplianceState'] or 'Not provided',
                             DeviceSyncStatus=dd[key(r['TenantDeviceKey'])]['SyncStatus']))
    for f, fk, d, dk in EXTRA_RELATIONSHIPS:
        parent = {key(r[dk]) for r in data[d]}
        if any(r[fk] and key(r[fk]) not in parent for r in data[f]):
            raise ValueError('Orphan 360 model relationship: '+f+'.'+fk)
    return EXTRA_RELATIONSHIPS
