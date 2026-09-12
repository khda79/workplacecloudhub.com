"""Additive local CMDB hardware report preparation for stable V1; never collects data."""
import argparse
import copy
import csv
import datetime as dt
import hashlib
import json
from decimal import Decimal
from pathlib import Path

IDENTITY = ['TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId']
PRODUCT = Path(__file__).resolve().parents[1]
DETAIL_COLUMNS = IDENTITY + ['TenantDeviceKey', 'ManagedDeviceId', 'SerialNumber', 'Manufacturer', 'Model',
    'Storage', 'HardwareCollectedDateTime', 'InventoryCollectedDateTime', 'CollectionCoverage', 'CollectionMode']
COVERAGE_COLUMNS = IDENTITY + ['Status', 'Coverage', 'Mode', 'RecordCount']


def read(path):
    with Path(path).open(encoding='utf-8-sig', newline='') as stream:
        reader = csv.DictReader(stream)
        rows = list(reader)
        if any(None in row or any(v is None for v in row.values()) for row in rows):
            raise ValueError('Malformed hardware report input')
        return reader.fieldnames, rows


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def prepare_data(dim_path, hardware_path=None):
    hashes = {str(dim_path): sha(dim_path)}
    _, devices = read(dim_path)
    if not devices:
        raise ValueError('An identified device inventory is required')
    identity = {k: devices[0][k] for k in IDENTITY}
    by_ci, keys = {}, set()
    for row in devices:
        if any(row[k] != identity[k] for k in IDENTITY) or not row['CmdbDeviceId'] or row['CmdbDeviceId'] in by_ci or not row['TenantDeviceKey'] or row['TenantDeviceKey'] in keys:
            raise ValueError('Invalid or mixed device inventory identity')
        by_ci[row['CmdbDeviceId']] = row
        keys.add(row['TenantDeviceKey'])
    coverage = dict(identity, Status='NotProvided', Coverage='Not provided', Mode='Not provided', RecordCount='0')
    result = []
    if hardware_path is not None:
        hardware_path = Path(hardware_path)
        manifest_path = hardware_path.parent / 'CIRegistry.manifest.json'
        hashes[str(hardware_path)] = sha(hardware_path)
        hashes[str(manifest_path)] = sha(manifest_path)
        manifest = load_json(manifest_path)
        evidence = manifest['Hardware']
        if manifest['Channel'] != 'stable' or manifest['Status'] != 'Exported' or evidence['Status'] != 'Validated':
            raise ValueError('A validated exported stable V1 CI manifest is required')
        source_hashes = dict(evidence['InputHashes'])
        source_hashes.update(manifest['SourceEvidence']['InputHashes'])
        for file, digest in source_hashes.items():
            if sha(file) != digest.upper():
                raise ValueError('CI source evidence changed')
            hashes[file] = digest.upper()
        raw_paths = [Path(p) for p in source_hashes if Path(p).name == 'Intune_DeviceHardware.csv']
        inventory_paths = [Path(p) for p in source_hashes if Path(p).name == 'Intune_ManagedDevices.csv']
        if len(raw_paths) != 1 or len(inventory_paths) != 1:
            raise ValueError('Native hardware and inventory evidence are required')
        raw_rows = read(raw_paths[0])[1]
        inventory_rows = read(inventory_paths[0])[1]
        def native_index(rows):
            index = {}
            for row in rows:
                key = row['ManagedDeviceId'].strip().casefold()
                if not key or key in index or any(row[k] != identity[k] for k in IDENTITY):
                    raise ValueError('Duplicate or foreign native evidence')
                index[key] = row
            return index
        native, inventory = native_index(raw_rows), native_index(inventory_rows)
        columns, rows = read(hardware_path)
        expected = load_json(PRODUCT / 'Schema/SmartWorkplaceCMDB.ci.hardware.json')['columns']
        if columns != expected or len(rows) != evidence['RowCount'] or len(rows) != len(native):
            raise ValueError('CI hardware header or count mismatch')
        seen = set()
        for row in rows:
            key = row['ManagedDeviceId'].strip().casefold()
            if any(row[k] != identity[k] for k in IDENTITY) or key in seen or key not in native or key not in inventory or row['CI_ID'] not in by_ci:
                raise ValueError('Invalid hardware tenant, native ID or CI mapping')
            seen.add(key)
            source, original, device = native[key], inventory[key], by_ci[row['CI_ID']]
            correlation = source['AzureAdDeviceId'].strip().casefold()
            if not correlation or correlation == '00000000-0000-0000-0000-000000000000':
                correlation = 'intune:' + key
            if correlation != device['SourceDeviceId'].strip().casefold():
                raise ValueError('Hardware CI correlation differs from report inventory')
            for field in ['SourceSystem', 'ManagedDeviceId', 'AzureAdDeviceId', 'SerialNumber', 'SerialNumberStatus', 'Manufacturer', 'ManufacturerStatus', 'Model', 'ModelStatus', 'TotalStorageSpaceInBytes', 'StorageStatus']:
                if row[field] != source[field]:
                    raise ValueError('CI hardware differs from its raw evidence')
            if row['HardwareCollectedDateTime'] != source['SourceCollectedDateTime'] or row['InventoryCollectedDateTime'] != original['SourceCollectedDateTime'] or row['CollectionCoverage'] != evidence['Coverage'] or row['CollectionMode'] != evidence['Mode']:
                raise ValueError('Hardware provenance mismatch')
            prepared = dict(identity, TenantDeviceKey=device['TenantDeviceKey'], ManagedDeviceId=row['ManagedDeviceId'])
            for field in ['SerialNumber', 'Manufacturer', 'Model']:
                status = row[field + 'Status']
                if status not in ('Reported', 'Missing') or (status == 'Missing' and row[field]) or (status == 'Reported' and not row[field].strip()):
                    raise ValueError('Inconsistent hardware text status')
                prepared[field] = row[field] if status == 'Reported' else 'Not provided'
            status, value = row['StorageStatus'], row['TotalStorageSpaceInBytes']
            if status == 'Missing' and value == '':
                prepared['Storage'] = 'Not provided'
            elif status == 'ZeroReported' and value == '0':
                prepared['Storage'] = 'Unknown (reported 0)'
            elif status == 'Reported' and value.isascii() and value.isdecimal() and 0 < int(value) <= 9223372036854775807:
                prepared['Storage'] = f'{Decimal(value) / Decimal(1073741824):,.2f} GiB'
            else:
                raise ValueError('Invalid hardware capacity')
            for field in ['HardwareCollectedDateTime', 'InventoryCollectedDateTime']:
                parsed = dt.datetime.fromisoformat(row[field].replace('Z', '+00:00'))
                if parsed.tzinfo is None:
                    raise ValueError('Unqualified hardware source date')
                prepared[field] = parsed.astimezone(dt.timezone.utc).isoformat()
            prepared.update(CollectionCoverage=row['CollectionCoverage'], CollectionMode=row['CollectionMode'])
            result.append(prepared)
        coverage.update(Status='Validated', Coverage=evidence['Coverage'], Mode=evidence['Mode'], RecordCount=str(len(result)))
    for file, digest in hashes.items():
        if sha(file) != digest:
            raise ValueError('Input changed during hardware preparation')
    return identity, result, coverage, hashes


def model_definitions(report_data, identity):
    # This is an MCP payload generator, not a TMDL writer.
    from build_report import m_query
    definitions = []
    for name, columns in [('DeviceHardware', DETAIL_COLUMNS), ('HardwareCoverage', COVERAGE_COLUMNS)]:
        query = '\n'.join(m_query(Path(report_data) / (name + '.csv'), columns, identity))
        # The shared query generator treats RecordCount as text; explicitly type it.
        if name == 'HardwareCoverage':
            query = query.replace('in Typed', 'in Table.TransformColumnTypes(Typed, {{"RecordCount", Int64.Type}})')
        definitions.append(dict(name=name, mode='Import', partitionName=name, mExpression=query,
            description='Stable V1 optional CI hardware evidence; missing source never becomes hardware values.',
            columns=[dict(name=c, sourceColumn=c, dataType='DateTime' if c.endswith('DateTime') else 'Int64' if c == 'RecordCount' else 'String',
                          summarizeBy='None', isHidden=c in IDENTITY + ['TenantDeviceKey', 'RecordCount'],
                          **({'formatString': 'yyyy-MM-dd HH:mm:ss'} if c.endswith('DateTime') else {})) for c in columns]))
    return definitions


def measures():
    common = 'Stable V1. Current device filter; separate native records are retained. Coverage does not prove hardware freshness.'
    expressions = [
        ('Hardware records', "COALESCE(COUNTROWS('DeviceHardware'),0)", '#,0', common),
        ('Hardware fleet rows', "COUNTROWS('DeviceHardware')", '#,0',
         'Stable V1. Internal nonempty gate for the fleet detail visual; blank combinations are excluded without changing record counts.'),
        ('Devices with hardware', "COALESCE(DISTINCTCOUNT('DeviceHardware'[TenantDeviceKey]),0)", '#,0', common),
        ('Devices without hardware', "COUNTROWS('DimDevice') - [Devices with hardware]", '#,0', common),
        ('Hardware covered devices', "CALCULATE([Devices with hardware], REMOVEFILTERS('DeviceHardware'[Manufacturer]), REMOVEFILTERS('DeviceHardware'[Model]), REMOVEFILTERS('DeviceHardware'[SerialNumber]), REMOVEFILTERS('DeviceHardware'[Storage]))", '#,0',
         'Stable V1. Distinct devices with a hardware record in the current DimDevice scope. Hardware-attribute selections are excluded from this coverage numerator.'),
        ('Hardware uncovered devices', "MAX(0, [Devices] - [Hardware covered devices])", '#,0',
         'Stable V1. Devices in the current DimDevice scope minus covered devices. The denominator is the filtered device inventory, not hardware records.'),
        ('Hardware coverage rate', "DIVIDE([Hardware covered devices], [Devices])", '0.0%',
         'Stable V1. Covered devices divided by devices in the current DimDevice scope. Hardware-attribute selections do not change the coverage numerator.'),
        ('Hardware data gap rate', "DIVIDE([Hardware uncovered devices], [Devices])", '0.0%',
         'Stable V1. Devices without a hardware record divided by all devices in the current DimDevice scope.'),
        ('Missing serial records', "CALCULATE([Hardware records], KEEPFILTERS('DeviceHardware'[SerialNumber] = \"Not provided\"))", '#,0',
         'Stable V1. Hardware records whose source serial status was Missing. This is distinct from a reported value and never defines identity or ownership.'),
        ('Missing manufacturer records', "CALCULATE([Hardware records], KEEPFILTERS('DeviceHardware'[Manufacturer] = \"Not provided\"))", '#,0',
         'Stable V1. Hardware records whose source manufacturer status was Missing.'),
        ('Missing model records', "CALCULATE([Hardware records], KEEPFILTERS('DeviceHardware'[Model] = \"Not provided\"))", '#,0',
         'Stable V1. Hardware records whose source model status was Missing.'),
        ('Zero-reported storage records', "CALCULATE([Hardware records], KEEPFILTERS('DeviceHardware'[Storage] = \"Unknown (reported 0)\"))", '#,0',
         'Stable V1. Hardware records where the source explicitly reported zero storage. This is not a missing value.'),
        ('Repeated serial values', "VAR _serials = FILTER(ADDCOLUMNS(VALUES('DeviceHardware'[SerialNumber]), \"_RecordCount\", CALCULATE(COUNTROWS('DeviceHardware'))), 'DeviceHardware'[SerialNumber] <> \"Not provided\" && [_RecordCount] > 1) RETURN COUNTROWS(_serials)", '#,0',
         'Stable V1. Distinct non-missing serial values appearing on more than one hardware record in the current filter. Records remain separate; serial never defines identity or ownership.'),
        ('Hardware detail rows', "IF(COUNTROWS(ALLSELECTED('DimDevice')) = 1, COUNTROWS('DeviceHardware'))", '#,0', common),
        ('Hardware source message', 'VAR _status = SELECTEDVALUE(\'HardwareCoverage\'[Status], "NotProvided") RETURN IF(_status = "NotProvided", "Hardware source not supplied. Existing inventory is retained; hardware attributes are unavailable.", IF(COUNTROWS(ALLSELECTED(\'DimDevice\')) <> 1, "Select one device to inspect its hardware. Coverage: " & SELECTEDVALUE(\'HardwareCoverage\'[Coverage]) & ".", IF([Hardware records] = 0, "No hardware record for this device in the supplied snapshot.", "Source-reported hardware. Missing values remain unknown; each source retains its own collection date.")))', '@', common),
    ]
    return [dict(tableName='DeviceHardware', name=n, expression=e, formatString=f, displayFolder='Hardware', description=d,
                 isHidden=n == 'Hardware fleet rows')
            for n, e, f, d in expressions]


def build_fleet_page(report):
    """Build the additive fleet-analysis page without changing report themes."""
    from build_report import lit, projection, exclude_synthetic_blank
    report = Path(report)
    pages = report / 'definition/pages'
    device360 = pages / 'device360'
    devices = pages / 'devices'

    def template(page, name):
        return load_json(page / 'visuals' / name / 'visual.json')

    page = copy.deepcopy(load_json(device360 / 'page.json'))
    page.update(name='hardwarefleet', displayName='10  Hardware fleet')
    page.pop('pageBinding', None)
    page.pop('filterConfig', None)
    visuals = []

    def put(v, x, y, w, h):
        v['name'] = 'hardwarefleetv' + str(len(visuals))
        v['position'] = dict(x=x, y=y, width=w, height=h, z=len(visuals), tabOrder=len(visuals))
        visuals.append(v)
        return v

    for name, text, y, h in [
        ('device360v0', 'Smart Workplace CMDB — Hardware fleet', 14, 42),
        ('device360v1', 'VERSION 1.0.0', 60, 26),
        ('device360v2', 'Understand source-reported fleet composition and coverage gaps. Missing attributes, reported zero storage and repeated serial values are separate checks; serial never defines identity or ownership.', 96, 44),
        ('device360v3', 'Frozen snapshot · Coverage denominator: filtered device inventory · Attribute checks use hardware records · Collection time is not observation time', 864, 28),
    ]:
        visual = template(device360, name)
        visual['visual']['objects']['general'][0]['properties']['paragraphs'][0]['textRuns'][0]['value'] = text
        put(visual, 24, y, 1232, h)

    def slicer(template_name, column, label, x):
        visual = template(devices, template_name)
        visual.pop('filterConfig', None)
        visual['visual']['query'] = {'queryState': {'Values': {'projections': [projection('DimDevice', column, label)]}}}
        visual['visual']['visualContainerObjects']['title'][0]['properties']['text'] = lit(label)
        put(visual, x, 144, 608, 76)
        exclude_synthetic_blank(visual, 'DimDevice', column)
        return visual

    slicer('devicesv4', 'OperatingSystem', 'Operating system', 24)
    slicer('devicesv5', 'Ownership', 'Ownership', 648)

    def card(measure, title, x, y, w=296, h=96, precision=0):
        visual = template(devices, 'devicesv7')
        visual['visual']['query'] = {'queryState': {'Data': {'projections': [projection('DeviceHardware', measure, title, True)]}}}
        visual['visual']['visualContainerObjects']['title'][0]['properties']['text'] = lit(title)
        value = visual['visual']['objects']['value'][0]['properties']
        value['fontSize'] = lit(22)
        value['labelPrecision'] = lit(precision)
        value['showBlankAs'] = lit('0')
        return put(visual, x, y, w, h)

    card('Hardware records', 'Hardware records', 24, 232)
    card('Hardware covered devices', 'Covered devices', 336, 232)
    card('Hardware uncovered devices', 'Devices without hardware', 648, 232)
    card('Hardware coverage rate', 'Hardware coverage', 960, 232, precision=1)

    def bar(column, label, title, x):
        visual = template(devices, 'devicesv11')
        visual['visual']['query'] = {
            'queryState': {
                'Category': {'projections': [dict(projection('DeviceHardware', column, label), active=True)]},
                'Y': {'projections': [projection('DeviceHardware', 'Hardware records', 'Hardware records', True)]},
            },
            'sortDefinition': {'sort': [{'field': projection('DeviceHardware', 'Hardware records', measure=True)['field'], 'direction': 'Descending'}], 'isDefaultSort': True},
        }
        visual['visual']['visualContainerObjects']['title'][0]['properties']['text'] = lit(title)
        put(visual, x, 336, 608, 208)
        exclude_synthetic_blank(visual, 'DeviceHardware', column)
        return visual

    bar('Manufacturer', 'Manufacturer', 'Hardware records by manufacturer', 24)
    bar('Model', 'Model', 'Hardware records by model — scroll for the full distribution', 648)

    anomaly_cards = [
        ('Missing serial records', 'Missing serial'),
        ('Missing manufacturer records', 'Missing manufacturer'),
        ('Missing model records', 'Missing model'),
        ('Zero-reported storage records', 'Reported zero storage'),
        ('Repeated serial values', 'Repeated serial values'),
    ]
    for index, (measure, title) in enumerate(anomaly_cards):
        card(measure, title, 24 + index * 248, 552, 240, 96)

    table = template(devices, 'devicesv13')
    table.pop('filterConfig', None)
    fields = [
        ('DimDevice', 'DeviceSelection', 'Open Device 360', False),
        ('DeviceHardware', 'Manufacturer', 'Manufacturer', False),
        ('DeviceHardware', 'Model', 'Model', False),
        ('DeviceHardware', 'SerialNumber', 'Serial number', False),
        ('DeviceHardware', 'Storage', 'Storage', False),
        ('DeviceHardware', 'Hardware fleet rows', 'Fleet row gate', True),
    ]
    projections = [projection(*field) for field in fields]
    projections[-1]['hidden'] = True
    table['visual']['query'] = {'queryState': {'Values': {'projections': projections}}}
    table['visual']['objects'].pop('columnWidth', None)
    headers = table['visual']['objects']['columnHeaders'][0]['properties']
    headers.update(autoSizeColumnWidth=lit(True), columnAdjustment=lit('growToFit'))
    table['visual']['visualContainerObjects']['title'][0]['properties']['text'] = lit('Hardware records — right-click a device to open Device 360; use Hardware detail for source evidence')
    put(table, 24, 656, 1232, 200)
    return page, visuals


def build_page(report):
    from build_report import lit, projection, exclude_synthetic_blank
    report = Path(report)
    device = report / 'definition/pages/device360'
    def template(name):
        return load_json(device / 'visuals' / name / 'visual.json')
    page = copy.deepcopy(load_json(device / 'page.json'))
    page.update(name='hardware', displayName='11  Hardware detail')
    page.pop('pageBinding', None)
    page.pop('filterConfig', None)
    visuals = []
    def put(v, x, y, w, h):
        v['name'] = 'hardwarev' + str(len(visuals))
        v['position'] = dict(x=x, y=y, width=w, height=h, z=len(visuals), tabOrder=len(visuals))
        visuals.append(v)
        return v
    for name, text, y, h in [('device360v0', 'Smart Workplace CMDB — Hardware detail', 14, 42),
                              ('device360v1', 'VERSION 1.0.0', 60, 26),
                              ('device360v2', 'Inspect source-reported equipment. Serial numbers never define identity or ownership.', 96, 44),
                              ('device360v3', 'Frozen snapshot · Native records are retained · Collection time is not observation time', 864, 28)]:
        v = template(name)
        v['visual']['objects']['general'][0]['properties']['paragraphs'][0]['textRuns'][0]['value'] = text
        put(v, 24, y, 1232, h)
    slicer = template('device360v4')
    slicer.pop('filterConfig', None)
    for entry in slicer['visual']['objects'].get('general', []):
        entry['properties'].pop('filter', None)
        entry['properties'].pop('selfFilter', None)
    put(slicer, 24, 144, 1232, 80)
    exclude_synthetic_blank(slicer, 'DimDevice', 'DeviceSelection')
    def table(fields, title, y, h):
        v = template('device360v5')
        v.pop('filterConfig', None)
        projections = [projection(t, c, label, measure) for t, c, label, measure in fields]
        v['visual']['query'] = {'queryState': {'Values': {'projections': projections}}}
        v['visual']['objects'].pop('columnWidth', None)
        v['visual']['objects']['columnHeaders'][0]['properties'].update(autoSizeColumnWidth=lit(True), columnAdjustment=lit('growToFit'))
        v['visual']['visualContainerObjects']['title'][0]['properties']['text'] = lit(title)
        return put(v, 24, y, 1232, h)
    table([('DeviceHardware', n, n, True) for n in ['Hardware records', 'Devices with hardware', 'Devices without hardware']], 'Coverage in the current device filter', 236, 130)
    table([('DeviceHardware', 'Hardware source message', 'Source status', True)], 'Availability', 378, 110)
    table([('DeviceHardware', c, label, False) for c, label in [('ManagedDeviceId', 'Source record'), ('SerialNumber', 'Serial number'), ('Manufacturer', 'Manufacturer'), ('Model', 'Model'), ('Storage', 'Storage')]] + [('DeviceHardware', 'Hardware detail rows', 'Detail rows', True)], 'Equipment — select one device above', 500, 150)
    table([('DeviceHardware', c, label, False) for c, label in [('ManagedDeviceId', 'Source record'), ('HardwareCollectedDateTime', 'Hardware collected UTC'), ('InventoryCollectedDateTime', 'Inventory collected UTC'), ('CollectionCoverage', 'Coverage'), ('CollectionMode', 'Mode')]] + [('DeviceHardware', 'Hardware detail rows', 'Detail rows', True)], 'Source evidence — dates are not interchangeable', 662, 186)
    for v in visuals:
        qs = v.get('visual', {}).get('query', {}).get('queryState', {})
        for p in qs.get('Values', {}).get('projections', []):
            if p['queryRef'] == 'DeviceHardware.Hardware detail rows':
                p['hidden'] = True
    return page, visuals


def prepare(project, output, hardware_path=None):
    project, output = Path(project), Path(output)
    if output.exists():
        raise ValueError('Preparation output must be new')
    identity, rows, coverage, hashes = prepare_data(project / 'ReportData/DimDevice.csv', hardware_path)
    page, visuals = build_page(project / 'CMDB-REPORTS.Report')
    fleet_page, fleet_visuals = build_fleet_page(project / 'CMDB-REPORTS.Report')
    output.mkdir(parents=True)
    for name, columns, data in [('DeviceHardware', DETAIL_COLUMNS, rows), ('HardwareCoverage', COVERAGE_COLUMNS, [coverage])]:
        with (output / (name + '.csv')).open('w', encoding='utf-8-sig', newline='') as stream:
            writer = csv.DictWriter(stream, columns); writer.writeheader(); writer.writerows(data)
    payload = {'tables': model_definitions(project / 'ReportData', identity), 'measures': measures(),
               'relationship': dict(name='cmdb-hardware-device', fromTable='DeviceHardware', fromColumn='TenantDeviceKey', fromCardinality='Many', toTable='DimDevice', toColumn='TenantDeviceKey', toCardinality='One', crossFilteringBehavior='OneDirection', isActive=True)}
    def write(path, data):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf-8')
    write(output / 'model-payload.json', payload)
    write(output / 'page/page.json', page)
    for visual in visuals:
        write(output / 'page/visuals' / visual['name'] / 'visual.json', visual)
    write(output / 'fleet-page/page.json', fleet_page)
    for visual in fleet_visuals:
        write(output / 'fleet-page/visuals' / visual['name'] / 'visual.json', visual)
    write(output / 'preparation.json', {'channel': 'stable', 'project': str(project), 'sourceStatus': coverage['Status'], 'hardwareRows': len(rows), 'inputHashes': hashes})
    return {'status': 'Prepared', 'hardwareRows': len(rows), 'sourceStatus': coverage['Status'],
            'visuals': len(visuals), 'detailVisuals': len(visuals), 'fleetVisuals': len(fleet_visuals)}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--ci-hardware', type=Path)
    args = parser.parse_args()
    print(json.dumps(prepare(args.project, args.output, args.ci_hardware)))
