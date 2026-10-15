"""Receipt-qualified Teams / SPO projection. No network, collection or publication.

Receipt run IDs and CSV business run IDs are distinct namespaces. Hashes bind
the exact bytes; business run consistency is checked within each producer.
"""
import csv
import datetime as dt
import hashlib
import json
import math
from pathlib import Path

FILES = {'Teams': ['M365_Teams_Teams.csv', 'M365_Teams_Channels.csv'],
         'SPO': ['M365_SPO_Sites.csv', 'M365_SPO_Tenant.csv']}
EXTRA_HEADERS = {
    'M365_Teams_Channels.csv': ['TeamId','ChannelId','MembershipType','SharePointSiteId','SharePointSiteLinkStatus'],
    'M365_SPO_Tenant.csv': ['StorageUsedGB','StorageCapacityGB','CapacitySource','IsPartialInventory']}
SCHEMAS = {
    'TeamsSiteLinkEvidence': ('Teams Site Links', {
        'TenantKey':'string', 'Link Key':'string', 'Team ID':'string', 'Team':'string',
        'Channel ID':'string', 'Channel':'string', 'Link Type':'string', 'Site ID':'string',
        'Link Status':'string', 'Resolution':'string', 'Site URL':'string',
        'Receipt Run ID':'string', 'Business Run ID':'string', 'Collected UTC':'dateTime'}),
    'SharePointSiteLinkEvidence': ('SharePoint Site Links', {
        'TenantKey':'string', 'Site ID':'string', 'Site':'string', 'Site URL':'string',
        'Teams Link State':'string', 'Link Count':'int64', 'Storage Used GB':'double',
        'Collection Scope':'string', 'Usage Report Date':'dateTime',
        'Collected UTC':'dateTime', 'Link Coverage Qualified':'bool'}),
    'SharePointTenantCapacityEvidence': ('SharePoint Tenant Capacity', {
        'TenantKey':'string', 'Storage Used GB':'double', 'Capacity GB':'double',
        'Utilization':'double', 'Capacity State':'string', 'Capacity Source':'string',
        'Collected UTC':'dateTime'}),
    'CollaborationSourceQualification': ('Collaboration Source Qualification', {
        'TenantKey':'string', 'Source':'string', 'Receipt Version':'string',
        'Receipt Run ID':'string', 'Completed UTC':'dateTime', 'Source Rows':'int64',
        'Link Coverage Qualified':'bool', 'Qualification':'string'})}


def read(path, required=()):
    if Path(path).is_symlink() or not Path(path).is_file():
        raise ValueError('Missing or linked Teams/SPO source file: ' + Path(path).name)
    with Path(path).open(encoding='utf-8-sig', newline='') as stream:
        reader = csv.DictReader(stream)
        headers = reader.fieldnames or []
        if len(headers) != len(set(headers)) or not set(required).issubset(headers):
            raise ValueError('Invalid Teams/SPO source schema: ' + Path(path).name)
        return list(reader)


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def utc(value):
    result = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
    if result.tzinfo is None:
        raise ValueError('Timezone required for collection evidence')
    return result.astimezone(dt.timezone.utc)


def norm(value):
    return (value or '').strip().casefold()


def number(value):
    if value in ('', None):
        return ''
    result = float(value)
    if not math.isfinite(result) or result < 0:
        raise ValueError('Invalid non-negative storage evidence')
    return result


def validate(root, identity, reference, max_age_hours=168, max_skew_hours=24):
    root = Path(root)
    proofs, fingerprints, starts = {}, {}, []
    for producer, names in FILES.items():
        receipt_name = f'SmartInventory_SmartM365-{producer}-Inventory.current.json.txt'
        if (root / receipt_name).is_symlink():
            raise ValueError('Linked Teams/SPO receipt')
        proof = json.loads((root / receipt_name).read_text(encoding='utf-8-sig'))
        if any(proof.get(k) != v for k, v in identity.items()):
            raise ValueError('Teams/SPO receipt tenant mismatch')
        if (not proof.get('RunId') or not proof.get('ScriptVersion')
                or proof.get('Status') != 'Completed' or type(proof.get('Errors')) is not int or proof.get('Errors') != 0
                or proof.get('IsPartialInventory') is not False
                or proof.get('ConsumerScopeQualified') is not True):
            raise ValueError('Teams/SPO receipt is not consumer-qualified')
        start, end = utc(proof['StartedAtUtc']), utc(proof['CompletedAtUtc'])
        if start > end or end > reference + dt.timedelta(minutes=5) or reference-start > dt.timedelta(hours=max_age_hours):
            raise ValueError('Stale or invalid Teams/SPO acquisition interval')
        starts.append(start)
        fingerprints[receipt_name] = digest(root / receipt_name)
        run_name = receipt_name.replace('.current.', '.run.')
        if (root / run_name).is_symlink():
            raise ValueError('Linked Teams/SPO run state')
        state = json.loads((root / run_name).read_text(encoding='utf-8-sig'))
        if state.get('Status') != 'Completed' or state.get('RunId') != proof.get('RunId') or state.get('UnqualifiedPublication') is not False:
            raise ValueError('Teams/SPO source publication is not complete')
        fingerprints[run_name] = digest(root / run_name)
        business_runs = set()
        for name in names:
            records = [r for r in proof['Files'] if r['File'] == name]
            if len(records) != 1:
                raise ValueError('Missing or duplicate Teams/SPO file proof: ' + name)
            record = records[0]
            if (record.get('Status') != 'Success' or type(record.get('Errors')) is not int or record.get('Errors') != 0
                    or record.get('IsPartialInventory') is not False
                    or record.get('RunId') != proof['RunId'] or record.get('Delimiter') != ','):
                raise ValueError('Unqualified Teams/SPO file: ' + name)
            if (root / name).is_symlink():
                raise ValueError('Linked Teams/SPO source file: ' + name)
            fingerprint = digest(root / name)
            if fingerprint != record['SHA256'].upper():
                raise ValueError('Teams/SPO hash mismatch: ' + name)
            data = read(root / name, EXTRA_HEADERS.get(name, []) + list(identity) + ['RunId'])
            if type(record.get('Rows')) is not int or len(data) != record['Rows']:
                raise ValueError('Teams/SPO row count mismatch: ' + name)
            for row in data:
                if any(row.get(k) != v for k, v in identity.items()) or not row.get('RunId'):
                    raise ValueError('Teams/SPO row identity mismatch: ' + name)
                business_runs.add(row['RunId'])
            fingerprints[name] = fingerprint
        if len(business_runs) > 1:
            raise ValueError('Mixed Teams/SPO business runs')
        proofs[producer] = proof
    return {'Fingerprints': fingerprints, 'Proofs': proofs,
            'Coherent': max(starts)-min(starts) <= dt.timedelta(hours=max_skew_hours)}


def recheck(root, audit):
    for name, expected in audit['Fingerprints'].items():
        if digest(Path(root) / name) != expected:
            raise ValueError('Teams/SPO source changed after validation: ' + name)


def project(teams, channels, sites, capacities, tenant, audit):
    def index(data, fields):
        result = {}
        for row in data:
            key = tuple(norm(row.get(f)) for f in fields)
            if not all(key) or key in result:
                raise ValueError('Missing or duplicate Teams/SPO native key')
            result[key] = row
        return result
    ti, ci, si = index(teams, ['TeamId']), index(channels, ['TeamId', 'ChannelId']), index(sites, ['SiteId'])
    if any((key[0],) not in ti for key in ci):
        raise ValueError('Teams channel has no observed parent')
    if len(capacities) != 1:
        raise ValueError('Expected one tenant capacity row')
    coverage = audit['Coherent'] and all(norm(t.get('IsPartialInventory')) == 'false' and
                t.get('ChannelCollectionStatus') == 'Collected' and
                t.get('SharePointSiteLinkStatus') == 'Collected' and t.get('SharePointSiteId') for t in teams)
    for key, team in ti.items():
        actual = [c for ck, c in ci.items() if ck[0] == key[0]]
        for kind, field in [('standard','StandardChannelCount'), ('private','PrivateChannelCount'), ('shared','SharedChannelCount')]:
            try:
                coverage = coverage and int(team.get(field, '-1')) == sum(norm(c.get('MembershipType')) == kind for c in actual)
            except ValueError:
                coverage = False
    links = []
    proof = audit['Proofs']['Teams']
    def link(row, team, kind):
        nonlocal coverage
        status, site_id = row.get('SharePointSiteLinkStatus', 'Not collected'), row.get('SharePointSiteId', '')
        if kind == 'standard':
            status, site_id = 'Not applicable', ''
        elif status != 'Collected' or not site_id:
            coverage = False
        match = si.get((norm(site_id),)) if status == 'Collected' else None
        resolution = ('Not applicable' if kind == 'standard' else 'Matched' if match else
                      'Outside observed SPO scope' if status == 'Collected' and site_id else 'Unknown')
        channel_id = row.get('ChannelId', '')
        key = hashlib.sha256(json.dumps([tenant, norm(team['TeamId']), norm(channel_id), kind]).encode()).hexdigest()
        links.append(dict(zip(SCHEMAS['TeamsSiteLinkEvidence'][1], [tenant, key, team['TeamId'],
            team.get('TeamDisplayName',''), channel_id, row.get('ChannelDisplayName',''), kind, site_id,
            status, resolution, match.get('SiteUrl','') if match else row.get('SharePointSiteUrl',''),
            proof['RunId'], row.get('RunId',''), proof['CompletedAtUtc']])))
    for team in teams:
        link(team, team, 'Team root')
    for row in channels:
        kind = norm(row.get('MembershipType'))
        if kind not in ('standard','private','shared'):
            coverage = False
        link(row, ti[(norm(row['TeamId']),)], kind)
    site_rows = []
    for site in sites:
        if norm(site.get('IsOneDrive')) != 'false':
            raise ValueError('Expected SPO site scope excluding OneDrive')
        matches = [r for r in links if r['Resolution'] == 'Matched' and norm(r['Site ID']) == norm(site['SiteId'])]
        site_rows.append(dict(zip(SCHEMAS['SharePointSiteLinkEvidence'][1], [tenant, site['SiteId'],
            site.get('Title',''), site.get('SiteUrl',''), 'Observed Teams link' if matches else
            'No observed Teams link' if coverage else 'Unknown', len(matches), number(site.get('StorageUsedGB')),
            site.get('CollectionScope','Unknown'), site.get('ReportRefreshDate',''), site.get('CollectedAtUtc',''), coverage])))
    capacity = capacities[0]
    used = number(capacity.get('StorageUsedGB'))
    cap = number(capacity.get('StorageCapacityGB'))
    source = capacity.get('CapacitySource','Unknown')
    measured = cap != '' and cap > 0 and source == 'Get-PnPTenant.StorageQuota' and norm(capacity.get('IsPartialInventory')) == 'false'
    capacity_row = dict(zip(SCHEMAS['SharePointTenantCapacityEvidence'][1], [tenant, used,
        cap if measured else '', used/cap if measured and used != '' else '',
        'Measured' if measured else 'Not measured', source, audit['Proofs']['SPO']['CompletedAtUtc']]))
    qualification = [dict(zip(SCHEMAS['CollaborationSourceQualification'][1], [tenant, producer,
        p['ScriptVersion'], p['RunId'], p['CompletedAtUtc'], sum(r['Rows'] for r in p['Files'] if r['File'] in FILES[producer]),
        coverage, 'Observed scope only; not a full tenant census. ' +
        ('Link coverage qualified.' if coverage else 'Link coverage incomplete or acquisition skew exceeds 24 hours.')]))
        for producer, p in audit['Proofs'].items()]
    return dict(zip(SCHEMAS, [links, site_rows, [capacity_row], qualification]))


def build(root, output, tenant, audit):
    recheck(root, audit)
    data = project(read(root/'M365_Teams_Teams.csv'), read(root/'M365_Teams_Channels.csv'),
                   read(root/'M365_SPO_Sites.csv'), read(root/'M365_SPO_Tenant.csv'), tenant, audit)
    for name, items in data.items():
        with (output/(name+'.csv')).open('x', encoding='utf-8', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=SCHEMAS[name][1])
            writer.writeheader(); writer.writerows(items)
    recheck(root, audit)
    return {name:len(items) for name, items in data.items()}
