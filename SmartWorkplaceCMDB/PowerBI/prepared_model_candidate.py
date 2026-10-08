"""Offline candidate schema and reference semantics; never edits a model.

These Python results test business expectations, NOT execution of DAX or M.
"""
import json
import math

import load_prepared_snapshot as reader
from prepared_read_session import PRODUCT_KEY, product_key

NEW_TABLES = ('EntraGroupMembership', 'ADUserSource', 'ADComputerSource',
              'ADGroupSource', 'ADDomainSource', 'ADDirectoryObject', 'ADMembership')
DATE_FIELDS = {'SourceCollectedDateTime', 'LastLogonUtcDateTime', 'CreationUtcDateTime'}


def model_plan():
    contract = reader._json(reader.CONTRACT.read_bytes())
    tables = []
    for definition in contract['tables']:
        if definition['name'] not in NEW_TABLES:
            continue
        tables.append({
            'name': definition['name'], 'source': definition['name'] + '.csv',
            'columns': [{'name': field, 'dataType': 'DateTime' if field in DATE_FIELDS
                         else 'Boolean' if field == 'Enabled' else 'String',
                         'summarizeBy': 'None', 'isHidden': field in reader.IDENTITY_FIELDS
                         or field.startswith('Tenant'),
                         'datePolicy': 'Qualified UTC only; null remains null' if field in DATE_FIELDS else None}
                        for field in definition['columns']],
            'key': definition['key'], 'allowEmpty': definition['allowEmpty'],
        })
    # Only analytical relationships are materialized. Exact 360 cloud links
    # use TREATAS measures, not unused inactive AD/cloud filter loops.
    relations = []
    for child, field, parent, active in (
        ('ADMembership', 'TenantADObjectKey', 'ADDirectoryObject', True),
        ('ADMembership', 'TenantADGroupKey', 'ADGroupSource', True),
        ('EntraGroupMembership', 'TenantGroupKey', 'DimGroup', True),
    ):
        relations.append(dict(fromTable=child, fromColumn=field, fromCardinality='Many',
                              toTable=parent, toColumn=field, toCardinality='One',
                              isActive=active, crossFilteringBehavior='OneDirection',
                              relyOnReferentialIntegrity=False))
    return {'Status': 'CandidateOnly', 'ContractVersion': contract['version'], 'tables': tables,
            'hardwareColumns': [{'name': 'PhysicalMemoryGiB', 'dataType': 'Double',
                                 'summarizeBy': 'None', 'formatString': '0.0',
                                 'description': 'Observed physical memory in GiB; missing is blank, not zero.'},
                                {'name': 'StorageStatus', 'dataType': 'String', 'summarizeBy': 'None'},
                                {'name': 'MemoryStatus', 'dataType': 'String', 'summarizeBy': 'None'}],
            'applicationProjection': {'table': 'DimDetectedApplication', 'column': PRODUCT_KEY,
                                      'dataType': 'String', 'summarizeBy': 'None', 'isHidden': True},
            'calculatedColumnUpdates': [
                {'tableName': 'FactDeviceApplication', 'name': 'Tenant Intune device key',
                 'expression': "LOOKUPVALUE ( 'DimIntuneManagedDevice'[TenantIntuneDeviceKey], "
                     "'DimIntuneManagedDevice'[ManagedDeviceId], 'FactDeviceApplication'[ManagedDeviceId], "
                     "'DimIntuneManagedDevice'[TenantKey], 'FactDeviceApplication'[TenantKey] )"},
                {'tableName': 'DimDetectedApplication', 'name': 'Application product',
                 'groupByColumns': [PRODUCT_KEY]},
            ],
            'relationships': relations,
            'preserveCalculatedColumns': ['DimUser[License review band]',
                'DimDevice[Device form factor]', 'DimDevice[Windows version group]',
                'DimDevice[Device ownership group]', 'DimDetectedApplication[Application product]',
                'DimCountry[Country footprint label]', 'FactDeviceApplication[Tenant Intune device key]'],
            'qualification': 'No model modifications, DAX execution, M execution or live refresh performed.'}


def application_counts(apps, links, managed, devices, *, countries=None,
                       ownership=None, selected_device=None):
    scoped = {row['TenantDeviceKey'] for row in devices
              if (countries is None or row['CountryLabel'] in countries)
              and (ownership is None or row['Ownership'] == ownership)
              and (selected_device is None or row['TenantDeviceKey'] == selected_device)}
    scoped_query = countries is not None or ownership is not None or selected_device is not None
    managed_ids = {row['ManagedDeviceId'] for row in managed if row['TenantDeviceKey'] in scoped}
    catalog = {row['TenantApplicationKey']: row for row in apps}
    products = {}
    for row in links:
        if scoped_query and row['ManagedDeviceId'] not in managed_ids:
            continue
        app = catalog[row['TenantApplicationKey']]
        products.setdefault(product_key(app), set()).add(row['ManagedDeviceId'])
    return {key: len(value) for key, value in products.items()}


def top_five(counts):
    return sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:5]


def imported_coverage(devices, managed, *, countries=None, ownership=None):
    scoped = {row['TenantDeviceKey'] for row in devices
              if (countries is None or row['CountryLabel'] in countries)
              and (ownership is None or row['Ownership'] == ownership)}
    enrolled = scoped & {row['TenantDeviceKey'] for row in managed}
    return len(enrolled) / len(scoped) if scoped else None


def selected_rows(rows, field, selected_keys):
    keys = set(selected_keys)
    if len(keys) != 1:
        return None  # Both 360 pages require one exact tenant-aware identity.
    return [row for row in rows if row[field] in keys]


def ad_summary(rows):
    def workstation(row):
        os = row['OperatingSystem'].casefold()
        return 'server' not in os and any('windows ' + version in os for version in ('7', '8', '10', '11'))
    eligible = [row for row in rows if row['Enabled'] is True and workstation(row)]
    managed = sum(row['CoverageState'] == 'Managed in Intune' for row in eligible)
    return {'without_dns': sum(not row['DNSHostName'].strip() for row in rows),
            'enabled_windows_objects': len(eligible), 'managed_objects': managed,
            'coverage': managed / len(eligible) if eligible else None}


def memory_gib(value):
    if value is None or value == '':
        return None
    result = float(value)
    if not math.isfinite(result) or result <= 0:
        raise ValueError('Physical memory must be a positive observed value or blank')
    return result


if __name__ == '__main__':
    print(json.dumps(model_plan(), indent=2))
