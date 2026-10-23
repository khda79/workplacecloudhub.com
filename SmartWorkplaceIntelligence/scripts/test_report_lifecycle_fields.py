"""Offline lifecycle-display and M365 slicer contract checks.

These checks inspect the actual PBIR/TMDL definitions; they do not refresh a
model or evaluate DAX. Run this file directly or through unittest discovery.
SMARTWORKPLACE_PBIP can select another complete PBIP folder.
"""
import json
import os
from pathlib import Path
import re
import unittest

from test_model_dax_references import object_name, parse_model


PBIP = Path(os.environ.get('SMARTWORKPLACE_PBIP') or (
    Path(__file__).resolve().parents[1] / 'pbip'))
MODEL = PBIP / 'SmartWorkplaceIntelligence.SemanticModel' / 'definition'
REPORT = PBIP / 'SmartWorkplaceIntelligence.Report' / 'definition'
NATIVE_AD = ('AD User Evidence', 'AD Computer Evidence')
INVENTORIES = ('User Inventory Evidence', 'Device Inventory Evidence')
DETAIL_PAGES = {
    '0443c39bb6eba584da07': INVENTORIES[0],
    '3a0fb4ff8eb05c8835b3': INVENTORIES[0],
    'c8ce269962e9453ea038': INVENTORIES[1],
    'a431090c3f25c4a369d0': INVENTORIES[1],
}
SOURCE_DATES = (
    'AD Created (source)', 'Entra Created (source)',
    'AD Modified (source)', 'Entra Modified (source)',
)
LIFECYCLE_NATIVE = {
    'CreationRaw', 'CreationUtcDateTime', 'CreationStatus',
    'ModificationRaw', 'ModificationUtcDateTime', 'ModificationStatus',
    'PasswordExpired', 'PasswordExpiryStatus', 'Password Expired (display)',
}


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def column_blocks(path):
    """Read one-level column declarations, including their TOM properties."""
    blocks = {}
    current = None
    for line in path.read_text(encoding='utf-8-sig').splitlines():
        if re.match(r'^\tcolumn ', line):
            current = object_name(line.split('column ', 1)[1])
            blocks[current] = [line]
        elif line.strip() and len(line) - len(line.lstrip('\t')) <= 1:
            current = None
        elif current is not None:
            blocks[current].append(line)
    return {name: '\n'.join(lines) for name, lines in blocks.items()}


def projection_columns(visual):
    """Find column expressions only in data projections, including aggregates."""
    def walk(node):
        if isinstance(node, dict):
            col = node.get('Column')
            if isinstance(col, dict):
                entity = col.get('Expression', {}).get('SourceRef', {}).get('Entity')
                if entity:
                    yield entity, col['Property']
            for value in node.values():
                yield from walk(value)
        elif isinstance(node, list):
            for value in node:
                yield from walk(value)

    query = visual.get('visual', {}).get('query', {}).get('queryState', {})
    return {pair for role in query.values() for projection in role.get('projections', [])
            for pair in walk(projection.get('field', {}))}


def true_property(block, name):
    return bool(re.search(r'^\t\t' + re.escape(name) + r'(?:\s*:\s*true)?\s*$',
                          block, re.M | re.I))


def relationship_endpoints(text):
    endpoints = set()
    for line in text.splitlines():
        match = re.match(r"^\t(?:fromColumn|toColumn):\s*('(?:[^']|'')+'|[^.]+)\.(.+)$", line)
        if match:
            endpoints.add((object_name(match.group(1)), object_name(match.group(2))))
    return endpoints


class ReportLifecycleFieldsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tables, expressions = parse_model(MODEL)
        cls.dax = {(item['table'], item['object']): item['dax'] for item in expressions}
        cls.blocks = {name: column_blocks(MODEL / 'tables' / (name + '.tmdl'))
                      for name in (*INVENTORIES, *NATIVE_AD, 'Native User Evidence',
                                   'Native Device Evidence', 'M365 License Selection',
                                   'M365 User License Membership',
                                   'M365 Device License Membership')}
        cls.visuals = [(path, read(path)) for path in
                       sorted((REPORT / 'pages').glob('*/visuals/*/visual.json'))]
        cls.relationships = (MODEL / 'relationships.tmdl').read_text(encoding='utf-8-sig')
        cls.endpoints = relationship_endpoints(cls.relationships)

    def test_native_expiry_is_hidden_and_display_is_observation_guarded_text(self):
        for table in NATIVE_AD:
            with self.subTest(table=table):
                technical = self.blocks[table]['PasswordExpired']
                self.assertRegex(technical, r'dataType:\s*boolean')
                self.assertTrue(true_property(technical, 'isHidden'))
                display = self.blocks[table]['Password Expired (display)']
                self.assertRegex(display, r'dataType:\s*string')
                dax = self.dax[table, 'column Password Expired (display)']
                self.assertRegex(dax, r'\[PasswordExpiryStatus\]\s*=\s*"Observed"')
                for state in ('"Unknown"', '"true"', '"false"'):
                    self.assertIn(state, dax)

    def test_inventory_expiry_has_unknown_observation_and_unique_batch_guards(self):
        for table in INVENTORIES:
            with self.subTest(table=table):
                self.assertRegex(self.blocks[table]['AD Password Expired'], r'dataType:\s*string')
                dax = self.dax[table, 'column AD Password Expired']
                for state in ('"Observed"', '"Unknown"', '"true"', '"false"'):
                    self.assertIn(state, dax)
                self.assertIn('[PasswordExpiryStatus]', dax)
                self.assertIn('[Loaded Batch ID]', dax)
                self.assertRegex(dax, r'_nativeCount\s*=\s*1')
                self.assertRegex(dax, r'COUNTROWS\s*\(\s*_adRows\s*\)\s*=\s*1')
                self.assertRegex(dax, r'NOT\s+ISBLANK\s*\(\s*_batch\s*\)')

    def test_report_never_projects_the_technical_password_boolean(self):
        forbidden = {(table, 'PasswordExpired') for table in NATIVE_AD}
        for path, visual in self.visuals:
            with self.subTest(visual=str(path.relative_to(REPORT))):
                self.assertFalse(projection_columns(visual) & forbidden)

    def test_native_ad_detail_facets_expose_display_status_and_observed_dates(self):
        for table in NATIVE_AD:
            required = {(table, field) for field in (
                'Password Expired (display)', 'PasswordExpiryStatus',
                'CreationRaw', 'CreationStatus', 'ModificationRaw', 'ModificationStatus')}
            details = [visual for _, visual in self.visuals
                       if visual.get('visual', {}).get('visualType') in ('tableEx', 'pivotTable') and
                       (table, 'ObjectGUID') in projection_columns(visual)]
            with self.subTest(table=table):
                self.assertTrue(details, 'Expected a native AD account detail facet')
                for visual in details:
                    self.assertTrue(required <= projection_columns(visual), visual['name'])

    def test_boolean_projection_detection_also_covers_aggregates(self):
        column = {'Column': {'Expression': {'SourceRef': {'Entity': NATIVE_AD[0]}},
                             'Property': 'PasswordExpired'}}
        visual = {'visual': {'query': {'queryState': {'Values': {'projections': [
            {'field': {'Aggregation': {'Expression': column, 'Function': 2}}}
        ]}}}}}
        self.assertEqual(projection_columns(visual), {(NATIVE_AD[0], 'PasswordExpired')})

    def test_user_and_pc_detail_pages_expose_source_dates_and_expiry_status(self):
        for page, table in DETAIL_PAGES.items():
            required = {(table, field) for field in
                        (*SOURCE_DATES, 'AD Password Expired', 'AD Password Expiry Status')}
            identity = 'User Principal Name' if table == INVENTORIES[0] else 'Device Name'
            details = [visual for path, visual in self.visuals
                       if path.parents[2].name == page and
                       visual.get('visual', {}).get('visualType') in ('tableEx', 'pivotTable') and
                       (table, identity) in projection_columns(visual) and
                       sum(entity == table for entity, _ in projection_columns(visual)) >= 4]
            with self.subTest(page=page):
                self.assertTrue(details, 'Expected at least one account detail view')
                for visual in details:
                    self.assertTrue(required <= projection_columns(visual),
                                    (page, visual['name'], 'Missing dates or explicit expiry status'))

    def test_source_dates_remain_text_with_qualification_not_inferred_utc(self):
        for table in INVENTORIES:
            for column in SOURCE_DATES:
                with self.subTest(table=table, column=column):
                    self.assertRegex(self.blocks[table][column], r'dataType:\s*string')
                    dax = self.dax[table, 'column ' + column]
                    self.assertIn('[CreationRaw]' if 'Created' in column else '[ModificationRaw]', dax)
                    self.assertNotRegex(dax, r'\b(?:DATEVALUE|TIMEVALUE|FORMAT)\s*\(')
                    qualifier = column.split()[0] + (' Creation' if 'Created' in column else ' Modification') + ' Qualification'
                    self.assertIn(qualifier, self.blocks[table])
                    self.assertIn('[Loaded Batch ID]', dax)

    def test_hidden_utc_dates_require_explicit_qualification(self):
        for table in INVENTORIES:
            for prefix in ('AD', 'Entra'):
                for event in ('Created', 'Modified'):
                    column = prefix + ' ' + event + ' UTC'
                    with self.subTest(table=table, column=column):
                        block = self.blocks[table][column]
                        self.assertRegex(block, r'dataType:\s*dateTime')
                        self.assertTrue(true_property(block, 'isHidden'))
                        dax = self.dax[table, 'column ' + column]
                        self.assertIn('"UTC qualified"', dax)
                        self.assertRegex(dax, r'\bBLANK\s*\(')

    def test_lifecycle_fields_are_not_keys_unique_columns_or_relationship_endpoints(self):
        for table, columns in self.blocks.items():
            for name, block in columns.items():
                if name not in LIFECYCLE_NATIVE and not (
                        table in INVENTORIES and any(term in name for term in
                                                     ('Created', 'Modified', 'Qualification', 'Password'))):
                    continue
                with self.subTest(table=table, column=name):
                    self.assertFalse(true_property(block, 'isKey'))
                    self.assertFalse(true_property(block, 'isUnique'))
                    self.assertNotIn((table, name), self.endpoints)
        contract = read(Path(__file__).resolve().parents[1] / 'config/native-evidence-contract.json.txt')
        for table in contract['tables']:
            self.assertFalse(set(table.get('key', [])) & LIFECYCLE_NATIVE, table['name'])

    def test_suite_membership_bridges_preserve_pairs_not_unique_individual_columns(self):
        for name, identity in (('M365 User License Membership', 'User Source ID'),
                               ('M365 Device License Membership', 'Device Source ID')):
            with self.subTest(table=name):
                self.assertEqual(self.tables[name]['columns'], {identity, 'License Product'})
                text = (MODEL / 'tables' / (name + '.tmdl')).read_text(encoding='utf-8-sig')
                self.assertRegex(text, r'\bDISTINCT\s*\(\s*UNION\s*\(')
                for column in (identity, 'License Product'):
                    block = self.blocks[name][column]
                    self.assertTrue(true_property(block, 'isHidden'))
                    self.assertFalse(true_property(block, 'isKey'))
                    self.assertFalse(true_property(block, 'isUnique'))
                self.assertIn('"Unknown"', text)

    def test_suite_selection_and_membership_routes_remain_explicit(self):
        selection = (MODEL / 'tables/M365 License Selection.tmdl').read_text(encoding='utf-8-sig')
        expected = ('Microsoft 365 F3', 'Microsoft 365 E3', 'Microsoft 365 E5',
                    'No F3/E3/E5 observed', 'Unknown')
        self.assertCountEqual(re.findall(r'\{\s*"([^"]+)"\s*\}', selection), expected)
        device = (MODEL / 'tables/M365 Device License Membership.tmdl').read_text(encoding='utf-8-sig')
        self.assertIn('LOOKUPVALUE', device)
        self.assertIn('[Primary User UPN]', device)
        self.assertIn('[Loaded Batch ID]', device)
        self.assertRegex(device, r'COUNTROWS\s*\(\s*_userBatches\s*\)\s*=\s*1')
        self.assertRegex(device, r'COUNTROWS\s*\(\s*_deviceBatches\s*\)\s*=\s*1')
        for identity, inventory, bridge in (
                ('User Source ID', INVENTORIES[0], 'M365 User License Membership'),
                ('Device Source ID', INVENTORIES[1], 'M365 Device License Membership')):
            self.assertIn("fromColumn: '" + bridge + "'.'" + identity + "'", self.relationships)
            self.assertIn("toColumn: '" + inventory + "'.'" + identity + "'", self.relationships)
            self.assertIn("fromColumn: '" + bridge + "'.'License Product'", self.relationships)
        self.assertEqual(self.relationships.count("toColumn: 'M365 License Selection'.'License Product'"), 2)

    def test_three_requested_m365_slicers_use_the_shared_suite_selection(self):
        for page in ('9ae7c13f821d4a608e52', '3a0fb4ff8eb05c8835b3', 'c8ce269962e9453ea038'):
            with self.subTest(page=page):
                slicers = [visual for path, visual in self.visuals if path.parents[2].name == page and
                           visual.get('visual', {}).get('visualType') == 'slicer' and
                           ('M365 License Selection', 'License Product') in projection_columns(visual)]
                self.assertEqual(len(slicers), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
