"""Native SmartInventory -> CMDB report tables; no collector or Power BI runtime.

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
from pathlib import Path

from cmdb_prepare import rows, normalized as key, utc

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


def number(value):
    if value == '':
        return ''
    result = float(value)
    if not math.isfinite(result) or result < 0:
        raise ValueError('Invalid non-negative numeric evidence')
    return result


def qualified(value):
    if not value:
        return '', 'Not provided'
    try:
        date = utc(value)
        if date.year < 1900:
            return '', 'Sentinel / unknown'
        return date.isoformat(), 'UTC qualified'
    except ValueError:
        return '', 'Format or timezone not established'


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


def build_tables(source, output, contract, identity, evidence, now=None):
    tenant = identity['TenantKey']
    reference = now or dt.datetime.now(dt.timezone.utc)
    definitions = {t['name']: t for t in contract['tables']}
    source_defs = {s['name']: s for s in contract['sources']}
    proof = {f['File']: f for f in evidence['Files']}
    table_counts = {}
    findings = {}

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
                writer.writerow({column: result.get(column, '') for column in definition['columns']})
                count += 1
        table_counts[name] = count

    def finding(entity, entity_id, finding_type, message, severity='Warning', discriminator='', **links):
        fid = native_key(tenant, 'finding', entity, entity_id, finding_type, discriminator)
        findings[fid] = {'TenantFindingKey':fid,'FindingId':fid,'EntityType':entity,
                         'EntityId':entity_id,'Severity':severity,'FindingType':finding_type,
                         'Description':message,'RecommendedAction':'Review native source evidence before taking action.',
                         'DetectedDateTime':reference.isoformat(),'SourceSystem':'SmartInventory',**links}

    users, users_by_native = [], {}
    for row in read('users'):
        sid = get(row,'Object Id')
        created, creation_status = qualified(get(row,'When created'))
        location = get(row,'Usage location').upper()
        country = location if re.fullmatch('[A-Z]{2}',location) else ''
        enabled = boolean(get(row,'AccountEnabled'))
        item = {'TenantUserKey':native_key(tenant,'user',sid),'CmdbUserId':'entra-user|'+key(sid),
                'SourceUserId':sid,'UserPrincipalName':get(row,'User principal name'),
                'DisplayName':get(row,'Display name'),'AccountEnabled':enabled,'UserType':get(row,'UserType'),
                'Department':get(row,'Department'),'JobTitle':get(row,'Title'),
                'CountryCode':country,'CountryLabel':country or UNKNOWN,
                'CountryStatus':'Reported' if country else 'Invalid' if location else 'Not provided',
                'LastSignInDateTime':get(row,'LastSignInDateTime'),
                'LastNonInteractiveSignInDateTime':get(row,'LastNonInteractiveSignInDateTime'),
                'LastSuccessfulSignInDateTime':get(row,'LastSuccessfulSignInDateTime'),
                'ActivityState':activity_state(get(row,'LastSuccessfulSignInDateTime') or get(row,'LastSignInDateTime'), acquired('users'), enabled),
                'AccountStatusLabel':'Enabled' if enabled=='true' else 'Disabled' if enabled=='false' else 'Unknown',
                'CreationRaw':get(row,'When created'),'CreationUtcDateTime':created,'CreationStatus':creation_status,
                'SourceCollectedDateTime':acquired('users'),'ManagerStatus':'Not collected',
                'UserActivityStatus':'Separate workload evidence','UserSelection':get(row,'Display name')+' | '+sid}
        users.append(item); users_by_native[key(sid)] = item
        if not country:
            finding('User',item['CmdbUserId'],'UserCountryUnknown','User usage location is missing or invalid; country is not inferred from AD.')
        if item['ActivityState'] in ('Invalid date','Future date'):
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
    native_groups = index(list(read('groups')), 'GroupId')
    if set(group_scope) != set(native_groups):
        raise ValueError('Full Entra group scope and membership scope differ')
    member_counts = collections.Counter()
    for member in read('group_members'):
        gid = key(member['GroupId'])
        if gid not in native_groups or member['MembershipKind'] != 'Direct':
            raise ValueError('Unqualified Entra group membership')
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
                       'SourceCollectedDateTime':acquired('groups')}

    entra = index(list(read('entra_devices')), 'ObjectId')
    managed = index(list(read('managed')), 'ManagedDeviceId')
    # Multiple Intune records can share a correlation ID. Retain all source rows;
    # choose a deterministic candidate only with qualified acquisition dates.
    candidates = collections.defaultdict(list)
    for mid, row in managed.items():
        corr = key(row['AzureADDeviceId'])
        if not corr or corr == ZERO:
            corr = 'intune:' + mid
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
        primary = users_by_native.get(key(get(md or {},'UserId')))
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
                'AssociationStatus':'Resolved' if primary else 'Unresolved' if get(md or {},'UserId') else 'Not provided',
                'PrimaryUserSourceId':get(md or {},'UserId'), 'CountryCode':primary['CountryCode'] if primary else '',
                'CountryLabel':primary['CountryLabel'] if primary else UNKNOWN,
                'CountryStatus':'Primary user usage location' if primary else 'No resolved primary user',
                'SyncRaw':sync_raw,'SyncUtcDateTime':sync,'SyncStatus':sync_status,
                'SourceCollectedDateTime':acquired('managed' if md else 'entra_devices')}
        item['DeviceSelection'] = item['DeviceName']+' | '+corr
        devices[corr] = item
        if md and get(md,'UserId') and not primary:
            finding('Device',item['CmdbDeviceId'],'OrphanPrimaryUserReference','Native Intune user ID is outside the collected Entra user population.')
        elif md and not get(md,'UserId'):
            finding('Device',item['CmdbDeviceId'],'DeviceWithoutPrimaryUser','Native Intune evidence does not provide a primary-user ID.')
        if not item['CountryCode']:
            finding('Device',item['CmdbDeviceId'],'DeviceCountryUnknown','Device country cannot be established from a uniquely resolved primary user.')
        else:
            finding('Device',item['CmdbDeviceId'],'DeviceCountryDerived','Device country comes from the primary user usage location, not device geolocation.', 'Information')
        if len(candidates.get(corr,[])) > 1:
            finding('Device',item['CmdbDeviceId'],'MultipleIntuneCandidates','Multiple native Intune records share a correlation ID; all evidence is retained and selection is deterministic.','Information')
    emit('DimDevice', devices.values())
    hardware = index(list(read('hardware')), 'ManagedDeviceId')
    if set(hardware) != set(managed):
        raise ValueError('Detailed hardware and managed-device identity scope differ')
    for mid, row in hardware.items():
        if row['CollectionStatus'] != 'Collected':
            raise ValueError('Incomplete detailed hardware acquisition')
        if key(get(row,'azureADDeviceId')) != key(get(managed[mid],'AzureADDeviceId')):
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
            if system=='Intune' and (not corr or corr==ZERO):
                corr='intune:'+key(sid)
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
                'SelectionStatus':'Selected' if chosen else 'Other candidate','SelectionReason':'Native ID; latest sync/enrollment, then native ID'})
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
                'HardwareCollectedDateTime':get(detail,'CollectedAtUtc'),'InventoryCollectedDateTime':acquired('managed'),
                'CollectionCoverage':'All native managed-device identities','CollectionMode':'Per-device GET with explicit hardware select'})
            for field in ('serialNumber','manufacturer','model'):
                if not get(detail,field):
                    finding('Device',device['CmdbDeviceId'],'HardwareAttributeMissing','Detailed hardware attribute not provided: '+field,discriminator=sid+'|'+field)
            if serials[key(get(detail,'serialNumber'))] > 1:
                finding('Device',device['CmdbDeviceId'],'RepeatedSerialValue','A reported serial appears on multiple native records; serial never defines identity.','Information',sid)
            if storage_status != 'Reported' or memory_status != 'Reported':
                finding('Device',device['CmdbDeviceId'],'HardwareCapacityUnknown','Missing and reported-zero storage/RAM remain distinct unknown evidence.',discriminator=sid)
            user = users_by_native.get(key(get(row,'UserId')))
            if chosen and user:
                relationship = native_key(tenant,'primary-user',sid,user['SourceUserId'])
                user_devices.append({'TenantRelationshipKey':relationship,'TenantUserKey':user['TenantUserKey'],
                    'TenantDeviceKey':device['TenantDeviceKey'],'CmdbRelationshipId':relationship,
                    'CmdbUserId':user['CmdbUserId'],'CmdbDeviceId':device['CmdbDeviceId'],'RelationshipType':'PrimaryUser',
                    'SourceSystem':'Intune','Account':user['UserPrincipalName'],'Device':device['DeviceName'],
                    'DeviceSelection':device['DeviceSelection'],'DeviceCompliance':device['ComplianceState'],'DeviceSyncStatus':device['SyncStatus']})
    emit('DeviceSource',device_sources); emit('DimIntuneManagedDevice',bridges); emit('DeviceHardware',hardware_rows)
    emit('HardwareCoverage',[{'Status':'Validated','Coverage':'All native managed-device identities',
                             'Mode':'Per-device GET with explicit hardware select','RecordCount':len(hardware_rows)}])
    emit('FactUserDeviceRelationship',user_devices)
    emit('FactDeviceCompliance',({'TenantDeviceKey':d['TenantDeviceKey'],'CmdbDeviceId':d['CmdbDeviceId'],
         'ComplianceState':d['ComplianceState'],'LastSyncDateTime':d['SyncRaw'],'SourceSystem':d['SourceSystems']} for d in devices.values()))

    skus = index(list(read('skus')),'Id')
    emit('DimLicenseSku',({'TenantSkuKey':native_key(tenant,'sku',sid),'SkuId':sid,
         'SkuPartNumber':row['TenantSkuPartNumber'],'ConsumedUnits':number(row['TenantConsumedUnits']),
         'EnabledUnits':number(row['TenantPrepaidEnabled'])} for sid,row in skus.items()))
    paths, pairs, group_path_counts = [], collections.defaultdict(list), collections.Counter()
    for row in read('license_paths'):
        uid, sid, gid = key(row['UserId']), key(row['SkuId']), key(row['AssignedByGroupId'])
        user, sku, group = users_by_native.get(uid), skus.get(sid), groups.get(gid)
        if not user or not sku:
            raise ValueError('License assignment parent identity missing')
        updated, status = qualified(get(row,'LastUpdatedDateTime'))
        paths.append({'TenantAssignmentPathKey':native_key(tenant,'license-path',uid,sid,gid),
            'TenantUserKey':user['TenantUserKey'],'TenantSkuKey':native_key(tenant,'sku',sid),
            'TenantGroupKey':group['TenantGroupKey'] if group else '', 'SourceUserId':uid,'SkuId':sid,
            'AssignedByGroupId':gid,'Account':user['UserPrincipalName'],'Product':sku['TenantSkuPartNumber'],
            'AssignmentRoute':'Group' if gid else 'Direct','GroupName':group['DisplayName'] if group else 'Unresolved group' if gid else 'Direct assignment',
            'AssignmentState':row['AssignmentState'],'AssignmentError':row['AssignmentError'],
            'ErrorStatus':'Error reported' if key(row['AssignmentError']) not in ('','none') else 'No error reported' if row['AssignmentError'] else 'Not provided',
            'DisabledPlanIds':get(row,'DisabledPlanIds'),'AssignmentUpdatedRaw':get(row,'LastUpdatedDateTime'),
            'AssignmentUpdatedUtcDateTime':updated,'AssignmentUpdatedStatus':status,'SourceCollectedDateTime':acquired('license_paths'),
            'GroupLinkStatus':'Resolved' if group else 'Unresolved' if gid else 'Not applicable','UserLinkStatus':'Resolved','SkuLinkStatus':'Resolved'})
        if key(row['AssignmentError']) not in ('','none'):
            finding('User',user['CmdbUserId'],'ObservedLicenseAssignmentError','Native license assignment reports an error; this does not prove non-use or overspend.',
                    discriminator=native_key(tenant,'license-path',uid,sid,gid), TenantGroupKey=group['TenantGroupKey'] if group else '')
        if gid and not group:
            finding('User',user['CmdbUserId'],'UnresolvedAssignmentGroup','License assignment group is absent from the current complete Entra group population.',discriminator=gid)
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
        assignment_rows.append({'TenantUserKey':users_by_native[uid]['TenantUserKey'],
            'TenantSkuKey':native_key(tenant,'sku',sid),'CmdbUserId':users_by_native[uid]['CmdbUserId'],
            'SkuId':sid,'AssignmentState':state,'SourceSystem':'Entra'})
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
            if uid not in users_by_native or (sid,pid) not in plans or (uid,sid) not in pairs:
                raise ValueError('User service-plan parent missing')
            state=row['StateCode']; decoded=state_codes.get(state)
            if decoded is None and state.startswith(('EN:','DIS:')):
                decoded=('true' if state.startswith('EN:') else 'false',state.split(':',1)[1])
            if decoded is None:
                raise ValueError('Unknown compact service-plan state code')
            yield {'TenantUserKey':users_by_native[uid]['TenantUserKey'],'TenantSkuKey':native_key(tenant,'sku',sid),
                   'TenantServicePlanKey':native_key(tenant,'plan',sid,pid),'CmdbUserId':users_by_native[uid]['CmdbUserId'],
                   'SkuId':sid,'ServicePlanId':pid,'ServicePlanName':plans[sid,pid]['PlanName'],
                   'IsEnabled':decoded[0],'AssignmentState':decoded[1],'SourceSystem':'Entra'}
    emit('FactUserServicePlan',service_plan_facts())

    applications=index(list(read('apps')),'AppId'); product_devices=collections.defaultdict(set)
    application_devices=collections.Counter(); product_versions=collections.defaultdict(set)
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
            if aid not in applications or mid not in managed:
                raise ValueError('Orphan application/device relation')
            application_devices[aid]+=1; product_devices[app_product[aid]].add(mid)
            yield {'TenantDeviceApplicationKey':native_key(tenant,'app-device',aid,mid),
                   'TenantApplicationKey':native_key(tenant,'app',aid),'AppId':aid,'ManagedDeviceId':mid,
                   'SourceCollectedDateTime':acquired('app_relations')}
    emit('FactDeviceApplication',installations())
    for aid,row in applications.items():
        if number(row['DeviceCount']) != application_devices[aid]:
            raise ValueError('Application per-version relation coverage differs from DeviceCount')
    emit('DimDetectedApplication',({'TenantApplicationKey':native_key(tenant,'app',aid),'AppId':aid,
        'SourceApplicationKey':aid,'DisplayName':row['AppName'],'Version':row['AppVersion'],
        'Publisher':row['AppPublisher'],'Platform':row['Platform'],'DeviceCount':application_devices[aid],
        'ReportedDeviceCount':row['DeviceCount'],'ExactRelatedDeviceCount':application_devices[aid],
        'RelationshipCoverageStatus':'Complete','SourceCollectedDateTime':acquired('apps')} for aid,row in applications.items()))
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
         'CreatedDateTime':get(r,'CreatedDateTime'),'LastModifiedDateTime':get(r,'LastModifiedDateTime'),
         'SourceCollectedDateTime':acquired('policies')} for r in policies if r['PolicyFamily'] not in ('WindowsFeatureUpdate','WindowsQualityUpdate')))
    emit('DimWindowsUpdatePolicy',({'TenantUpdatePolicyKey':native_key(tenant,'policy',r['PolicyFamily'],r['PolicyId']),
         'PolicyType':r['PolicyFamily'],'PolicyId':r['PolicyId'],'DisplayName':get(r,'DisplayName'),
         'TargetVersion':get(r,'FeatureUpdateVersion'),'CreatedDateTime':get(r,'CreatedDateTime'),
         'LastModifiedDateTime':get(r,'LastModifiedDateTime'),'SourceCollectedDateTime':acquired('policies')}
         for r in policies if r['PolicyFamily'] in ('WindowsFeatureUpdate','WindowsQualityUpdate')))

    mailbox_rows,hosting={},{}
    for family,hosting_location in [('local_mailboxes','Exchange On-premises'),('remote_mailboxes','Exchange Online'),('mailboxes','Exchange Online')]:
        source_addresses=set()
        for r in read(family):
            native=get(r,'MailboxGuid') if family=='mailboxes' else get(r,'ObjectGUID')
            smtp=get(r,'PrimarySmtpAddress','PrimarySMTPaddress')
            recipient=get(r,'RecipientType') if family=='local_mailboxes' else get(r,'RecipientTypeDetails')
            location='Exchange Online' if 'remote' in key(recipient) else hosting_location
            mkey=key(smtp) if smtp else family+':'+key(native)
            if mkey in source_addresses:
                raise ValueError('Conflicting native mailbox identities share SMTP within one source')
            source_addresses.add(mkey)
            old=hosting.get(mkey)
            if old and old['HostingLocation']=='Exchange Online' and location!='Exchange Online':
                continue
            user=users_by_native.get(key(get(r,'ExternalDirectoryObjectId'))) or user_by_address(smtp)
            hkey=native_key(tenant,'mailbox',mkey)
            mailbox_rows[mkey]={'TenantMailboxKey':hkey,'CmdbMailboxId':hkey,'TenantUserKey':user['TenantUserKey'] if user else '',
                'CmdbUserId':user['CmdbUserId'] if user else '', 'RecipientTypeDetails':recipient,'ArchiveStatus':get(r,'ArchiveStatus','ArchiveState'),
                'SourceSystem':family,'DisplayName':get(r,'DisplayName'),'PrimarySmtpAddress':smtp,
                'LinkStatusLabel':'Resolved' if user else 'Unresolved','ExternalDirectoryObjectId':get(r,'ExternalDirectoryObjectId')}
            recipient_kind=key(recipient)
            hosting[mkey]={'MailboxHostingKey':hkey,'CountryLabel':user['CountryLabel'] if user else UNKNOWN,
                'HostingLocation':location,'RecipientTypeDetails':recipient,
                'MailboxTypeGroup':'Shared mailboxes' if 'sharedmailbox' in recipient_kind else 'User mailboxes' if recipient_kind in ('usermailbox','remoteusermailbox') else 'Other types',
                'EvidenceSource':family}
            if not smtp:
                finding('mailbox',hkey,'Missing SMTP','Native mailbox retained without an address; hosting deduplication cannot use SMTP.')
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
        'EnrollmentState':get(r,'Enrollment state'),'LastContactedDateTime':get(r,'Last contacted'),
        'AzureAdDeviceId':get(r,'Azure AD Device ID'),'ManagedDeviceId':get(r,'Managed device ID'),
        'SourceCollectedDateTime':acquired('autopilot')} for r in read('autopilot')))
    analytics=[]
    for r in read('analytics'):
        scores={field:number(get(r,source_field)) for field,source_field in [('EndpointAnalyticsScore','EndpointAnalyticsScore'),('StartupPerformanceScore','StartupScore'),('AppReliabilityScore','AppReliabilityScore'),('WorkFromAnywhereScore','WorkFromAnywhereScore')]}
        if any(value!='' and value>100 for value in scores.values()):
            raise ValueError('Endpoint Analytics score outside 0-100')
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
        created,creation_status=qualified(get(row,'WhenCreated'))
        return {'LastLogonTimestampRaw':get(row,'LastLogonTimestamp'),'LastLogonUtcDateTime':last,'LastLogonStatus':status,
                'CreationRaw':get(row,'WhenCreated'),'CreationUtcDateTime':created,'CreationStatus':creation_status}
    ad_user_rows=[]
    for r in ad_users:
        cloud,status=cloud_match(r,ad_user_sids,cloud_user_sids)
        user=users_by_native.get(key(cloud['Object Id'])) if cloud else None
        ad_user_links[key(r['ObjectGUID'])]=(user,status)
        manager_dn=get(r,'manager'); manager=ad_object_by_dn(manager_dn) if manager_dn else None
        ad_user_rows.append({'TenantADUserKey':native_key(tenant,'ad-user',r['ObjectGUID']),
            **{field:get(r,field) for field in ('ObjectGUID','ObjectSID','DomainName','DistinguishedName','SamAccountName','UserPrincipalName','DisplayName','Department','Title','Country')},
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
        memberships.append({'TenantADMembershipKey':native_key(tenant,'ad-membership',r['GroupSID'],r['MemberDistinguishedName'],r['MembershipKind']),
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
    emit('SourceHealth',({'SourceName':r['File'],'Status':r['Status'],'Coverage':'Complete producer file scope',
        'SourceRows':r['Rows'],'MaxItems':0,'StartedDateTime':r['StartedAtUtc'],'CompletedDateTime':r['CompletedAtUtc'],
        'Evidence':'Native producer success; exact hash and logical row count'} for r in evidence['Files']))
    relationships=[('PrimaryUser',len(user_devices),'Native Intune primary user'),('HasMailbox',sum(bool(m['TenantUserKey']) for m in mailbox_rows.values()),'Unique native ID or SMTP/user match'),
        ('AssignedLicense',len(assignment_rows),'Native Entra user/SKU pairs'),('MemberOfGroup',sum(member_counts.values()),'Direct Entra memberships, including non-user objects'),
        ('DeviceHasApplication',table_counts['FactDeviceApplication'],'Native Intune app/device links'),('DeviceInAutopilot',sum(bool(get(r,'Managed device ID')) for r in read('autopilot')),'Native Autopilot managed device link'),
        ('ADMemberOfGroup',sum(1 for _ in read('ad_members')),'Native direct and primary-group evidence'),
        ('PolicyAssignmentTarget',sum(1 for _ in read('policy_assignments')),'Configured target; not proof of effective device assignment')]
    emit('FactRelationshipOverview',({'RelationshipType':name,'RelationshipCount':count,'EvidenceSource':text} for name,count,text in relationships))
    if set(table_counts)!=set(definitions):
        raise ValueError('Builder did not cover every reporting contract table')
