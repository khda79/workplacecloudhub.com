"""Version 0.5.0: offline, exact coverage of independently observed Intune conflicts."""
import hashlib
import json

VERSION = '0.5.0'
FILE = 'Intune_DeviceIdentityEvidence.csv'
PRODUCER = 'SmartM365-Devices-Inventory.ps1'
IDENTITY_FIELDS = {'TenantKey','OrganizationKey','EnvironmentKey','TenantId'}


def text(value):
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'True' if value else 'False'
    return str(value)


def key(value):
    return text(value).strip().casefold()


def qualify(source, contract, identity, evidence):
    # Import at call time to keep the source reader independent of this validator.
    from workplace_native_sources import rows, utc, check_rows, sha
    definitions = {d['name']: d for d in contract['sources']}
    records = {r['File']: r for r in evidence['Files']}
    def indexed(name):
        result = {}
        for row in rows(source / definitions[name]['file']):
            mid = key(row['ManagedDeviceId'])
            if not mid or mid in result:
                raise ValueError('Missing or duplicate native device identity')
            result[mid] = row
        return result
    managed, hardware = indexed('managed'), indexed('hardware')
    if set(managed) != set(hardware):
        raise ValueError('Detailed hardware and managed-device identity scope differ')
    mp, hp = (records[definitions[n]['file']] for n in ('managed', 'hardware'))
    if any(mp[f] != hp[f] for f in ('RunId', 'Producer', 'StartedAtUtc', 'Scope')):
        raise ValueError('Hardware and managed observations belong to different runs')
    start, end = utc(mp['StartedAtUtc']), utc(hp['CompletedAtUtc'])
    artifacts = evidence.get('SemanticFiles', [])
    for collection in ((managed, hardware) if artifacts else (hardware,)):
        for row in collection.values():
            if not start <= utc(row['CollectedAtUtc']) <= end:
                raise ValueError('Device observation outside producer acquisition interval')
    if any(row['CollectionStatus'] != 'Collected' for row in hardware.values()):
        raise ValueError('Incomplete detailed hardware acquisition')
    mismatches = {mid for mid in managed if key(managed[mid]['AzureADDeviceId']) != key(hardware[mid]['azureADDeviceId'])}
    if not artifacts:
        if mismatches:
            raise ValueError('Detailed hardware correlation differs from managed inventory')
        return set(), {'Status':'Legacy consistent', 'ManagedIdentities':len(managed), 'ConflictCount':0, 'ConvergedCount':0, 'ReReadCount':0}
    if len(artifacts) != 1 or artifacts[0]['File'] != FILE:
        raise ValueError('Unexpected device identity evidence artifact')
    record = artifacts[0]
    if any(record[f] != mp[f] for f in ('RunId', 'Producer', 'StartedAtUtc', 'Scope')):
        raise ValueError('Device identity evidence belongs to another run')
    definition = contract['deviceIdentity']
    path = source / FILE
    if path.is_symlink() or not path.is_file() or sha(path) != record['SHA256']:
        raise ValueError('Device identity evidence SHA256 mismatch or unsafe artifact')
    if check_rows(path, definition, identity['TenantKey'], exact=True, identity=identity) != record['Rows']:
        raise ValueError('Device identity evidence row count mismatch')
    conflicts, seen, converged, attempts = set(), set(), 0, 0
    for row in rows(path):
        mid = key(row['ManagedDeviceId'])
        if mid in seen or mid not in managed:
            raise ValueError('Duplicate or unknown device identity evidence')
        seen.add(mid)
        if row['ContractVersion'] != VERSION or row['RunId'] != mp['RunId']:
            raise ValueError('Device identity evidence version or run mismatch')
        raw = row['EvidenceJson']
        if hashlib.sha256(raw.encode('utf-8')).hexdigest().upper() != row['EvidenceSHA256']:
            raise ValueError('Device identity observation fingerprint mismatch')
        def unique_object(pairs):
            obj = {}
            for k, v in pairs:
                if k in obj:
                    raise ValueError('Duplicate device evidence JSON key')
                obj[k] = v
            return obj
        doc = json.loads(raw, object_pairs_hook=unique_object)
        if any(doc.get(f) != row[f] for f in ('ContractVersion', 'TenantKey', 'RunId', 'ManagedDeviceId', 'Resolution')):
            raise ValueError('Device identity observation lineage mismatch')
        observations = doc.get('Observations')
        if not isinstance(observations, list) or not 2 <= len(observations) <= 4:
            raise ValueError('Incomplete bounded identity observations')
        last_time = start
        for number, observation in enumerate(observations):
            if type(observation.get('Attempt')) is not int or observation['Attempt'] != number:
                raise ValueError('Invalid identity observation attempt sequence')
            m, h = observation['Managed'], observation['Hardware']
            required_managed = set(managed[mid]) - IDENTITY_FIELDS
            required_hardware = set(hardware[mid]) - IDENTITY_FIELDS
            if set(m) != required_managed or set(h) != required_hardware:
                raise ValueError('Incomplete device identity observation fields')
            if key(m['ManagedDeviceId']) != mid or key(h['ManagedDeviceId']) != mid or h['CollectionStatus'] != 'Collected':
                raise ValueError('Foreign or failed device identity observation')
            byte_values = {}
            for field in ('totalStorageSpaceInBytes','freeStorageSpaceInBytes','physicalMemoryInBytes'):
                value, status = text(h[field]), h[field+'Status']
                if value == '' and status == 'Missing':
                    continue
                if (not value.isascii() or not value.isdecimal() or int(value) > 9223372036854775807
                        or status != ('ZeroReported' if int(value) == 0 else 'Reported')):
                    raise ValueError('Invalid hardware value in identity observation')
                byte_values[field] = int(value)
            if ('totalStorageSpaceInBytes' in byte_values and 'freeStorageSpaceInBytes' in byte_values
                    and byte_values['freeStorageSpaceInBytes'] > byte_values['totalStorageSpaceInBytes']):
                raise ValueError('Invalid storage capacity in identity observation')
            mt, ht = utc(m['CollectedAtUtc']), utc(h['CollectedAtUtc'])
            if not last_time <= mt <= ht <= end:
                raise ValueError('Invalid device identity observation acquisition interval')
            last_time = ht
            agrees = key(m['AzureADDeviceId']) == key(h['azureADDeviceId'])
            if agrees and number != len(observations)-1:
                raise ValueError('Identity attempts continued after convergence')
            native = observation.get('NativeManaged')
            ids = [v for k,v in native.items() if k.casefold() == 'id'] if isinstance(native,dict) else []
            if len(ids) != 1 or key(ids[0]) != mid:
                raise ValueError('Missing full native managed observation')
        resolution = row['Resolution']
        agrees = key(observations[-1]['Managed']['AzureADDeviceId']) == key(observations[-1]['Hardware']['azureADDeviceId'])
        if resolution == 'Identity conflict' and len(observations) == 4 and not agrees:
            conflicts.add(mid)
            published = observations[0]
        elif resolution == 'Converged' and agrees:
            converged += 1
            published = observations[-1]
        else:
            raise ValueError('Invalid device identity resolution or incomplete attempts')
        for kind, actual in (('Managed', managed[mid]), ('Hardware', hardware[mid])):
            if any(text(published[kind][f]) != actual[f] for f in actual if f not in IDENTITY_FIELDS):
                raise ValueError('Published device observation differs from fingerprinted evidence')
        attempts += len(observations)-1
    if conflicts != mismatches:
        raise ValueError('Device identity quarantine does not exactly cover divergent observations')
    if sha(path) != record['SHA256']:
        raise ValueError('Device identity evidence changed during validation')
    return conflicts, {'Status':'SuccessWithWarnings' if conflicts else 'Success', 'ManagedIdentities':len(managed),
        'ConflictCount':len(conflicts), 'ConsistentIdentities':len(managed)-len(conflicts), 'InitialDivergentCount':len(seen), 'ConvergedCount':converged, 'ReReadCount':attempts,
        'Artifact':{'File':FILE, 'BatchFile':'Source-'+FILE, 'SHA256':record['SHA256'], 'RunId':record['RunId']},
        'Policy':'TenantKey plus ManagedDeviceId; disputed Entra/AD links are indeterminate; all Intune observations retained; unified entities are not a physical asset count'}


def validate_windows(source):
    """Standalone derived readers verify the same receipt and Windows-view binding."""
    from workplace_native_sources import load_json, CONTRACT, REGISTRY, sha, rows, check_publication, utc
    contract = load_json(CONTRACT)
    producer = next(p for p in load_json(REGISTRY)['Producers'] if p['Script'] == PRODUCER)
    receipt_path = source / producer['Receipt']
    if receipt_path.is_symlink():
        raise ValueError('Linked device receipt')
    digest = sha(receipt_path)
    receipt = load_json(receipt_path)
    check_publication(receipt_path, receipt)
    identity = {f:receipt.get(f) for f in ('TenantKey','OrganizationKey','EnvironmentKey','TenantId')}
    required = producer['Files'] + producer['SemanticFiles']['Files']
    if (receipt.get('Owner') != 'SmartInventory-SourceReceipt' or receipt.get('ContractVersion') != '1.2'
            or receipt.get('Producer') != PRODUCER or receipt.get('Status') != 'Completed'
            or receipt.get('IsPartialInventory') is not False or type(receipt.get('Errors')) is not int or receipt['Errors'] != 0
            or receipt.get('ConsumerScopeQualified') is not True or receipt.get('ScopeQualification') != 'ConsumerScope'
            or receipt.get('Scope') != producer['Scope'] or set(receipt.get('RequiredFiles', [])) != set(required)
            or any(not isinstance(v,str) or not v for v in identity.values())):
        raise ValueError('Unqualified device identity receipt')
    record_index = {}
    for record in receipt['Files']:
        name = record['File']
        if name in record_index or '/' in name or '\\' in name:
            raise ValueError('Unsafe or duplicate device receipt file')
        record_index[name] = record
    selected = required + ['Intune_Devices_Inventory.csv']
    for name in selected:
        record, path = record_index[name], source / name
        if (path.is_symlink() or sha(path) != record['SHA256'] or record.get('Status') != 'Success'
                or type(record.get('Errors')) is not int or record['Errors'] != 0 or record.get('IsPartialInventory') is not False
                or any(record.get(f) != receipt[f] for f in ('RunId','Producer','ScriptVersion','StartedAtUtc','Scope'))
                or not utc(receipt['StartedAtUtc']) <= utc(record['CompletedAtUtc']) <= utc(receipt['CompletedAtUtc'])):
            raise ValueError('Device derived-reader source fingerprint or lineage mismatch')
        data = list(rows(path))
        if type(record['Rows']) is not int or len(data) != record['Rows'] or any(any(r[f] != v for f,v in identity.items() if f in r) for r in data):
            raise ValueError('Device derived-reader row count or tenant mismatch')
    evidence = {'Files':[record_index[n] for n in producer['Files']], 'SemanticFiles':[record_index[FILE]]}
    conflicts, summary = qualify(source, contract, identity, evidence)
    windows = list(rows(source / 'Intune_Devices_Inventory.csv'))
    managed = {key(r['ManagedDeviceId']):r for r in rows(source / 'Intune_ManagedDevices_All.csv')}
    expected = {mid for mid,r in managed.items() if key(r['OperatingSystem']) == 'windows'}
    seen = set()
    for row in windows:
        mid = key(row['Device ID'])
        if mid in seen or mid not in expected or row.get('Identity RunId') != receipt['RunId'] or row.get('Identity Contract') != VERSION:
            raise ValueError('Windows view identity scope, run or version mismatch')
        seen.add(mid)
        if row.get('Identity Status') != ('Identity conflict' if mid in conflicts else 'Consistent'):
            raise ValueError('Windows view identity qualification mismatch')
        if key(row['Azure AD Device ID']) != key(managed[mid]['AzureADDeviceId']):
            raise ValueError('Windows view differs from observed managed identity')
        if mid in conflicts and any(row.get(f) for f in ('Entra ObjectId','Entra DeviceId','Entra DisplayName','Entra ApproximateLastSignInDateTime')):
            raise ValueError('Windows view retains a disputed Entra join')
    if seen != expected:
        raise ValueError('Windows view lost native managed identities')
    if sha(receipt_path) != digest or any(sha(source/n) != record_index[n]['SHA256'] for n in selected):
        raise ValueError('Device derived-reader evidence changed during validation')
    return dict(summary, ConflictIds=sorted(conflicts))


if __name__ == '__main__':
    import argparse
    from pathlib import Path
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--windows-source', required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(validate_windows(args.windows_source)))
