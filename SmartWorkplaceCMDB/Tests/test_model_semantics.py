"""Synthetic protocol and semantics tests; no production data or model writes."""
import datetime as dt
import http.client
import json
import re
from pathlib import Path
import sys
import threading
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'Scripts'))
import load_prepared_snapshot as reader
import reference_semantics as model
import prepared_read_session as session_reader
import test_prepared_snapshot as fixtures


class SessionTests(unittest.TestCase):
    def setUp(self):
        # Reuse only the fixture, not inheritance (which would repeat 28 tests).
        self.fixture = fixtures.SnapshotTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.clock = 100.0
        self.now = self.fixture.now
        self.session = session_reader.ReadSession(self.fixture.load(), self.fixture.contract,
            ttl_seconds=60, monotonic=lambda: self.clock, utc_now=lambda: self.now)
        self.addCleanup(self.session.close)

    def test_original_46_tables_same_pinned_batch(self):
        for name in self.fixture.data:
            status, data = self.session.get(self.session.prefix + name)
            self.assertEqual(status, 200)
            if name != 'DimDetectedApplication.csv':
                self.assertEqual(data, self.fixture.data[name][0])

    def test_later_file_replacement_cannot_change_session(self):
        path = self.session.prefix + 'DimCountry.csv'
        before = self.session.get(path)
        (self.fixture.root / 'DimCountry.csv').write_bytes(b'changed')
        self.fixture.manifest['GeneratedAtUtc'] = 'later-invalid'
        self.fixture.write_manifest()
        self.assertEqual(before, self.session.get(path))

    def test_wrong_batch_token_traversal_and_query_rejected(self):
        self.assertEqual(self.session.get(self.session.prefix.replace(self.session.snapshot.batch_sha256, '0' * 64) + 'DimCountry.csv')[0], 409)
        for path in (self.session.prefix.replace(self.session.token, 'wrong') + 'DimCountry.csv',
                     self.session.prefix + '../current.json.txt',
                     self.session.prefix + '%2e%2e/current.json.txt',
                     self.session.prefix + 'DimCountry.csv?other=1',
                     self.session.prefix + 'current.json.txt'):
            self.assertEqual(self.session.get(path)[0], 404)

    def test_expiry_close_and_acquisition_freshness_are_fail_closed(self):
        path = self.session.prefix + 'DimCountry.csv'
        self.clock = 160
        self.assertEqual(self.session.get(path)[0], 410)
        self.clock = 100
        self.now = self.session.fresh_until + dt.timedelta(seconds=1)
        self.assertEqual(self.session.get(path)[0], 410)
        self.now = self.fixture.now
        self.session.close()
        self.assertEqual(self.session.get(path)[0], 410)

    def test_weekly_apps_session_uses_earliest_individual_expiry(self):
        f = self.fixture
        f.age_apps(216)
        session = session_reader.ReadSession(f.load(), f.contract, utc_now=lambda: f.now)
        self.addCleanup(session.close)
        self.assertEqual(session.fresh_until, f.now + dt.timedelta(hours=24))
        self.assertEqual(len(session.freshness['Warnings']), 2)
        metadata = json.loads(session.resources['model-contract.json.txt'])
        self.assertEqual(metadata['SourceFreshness']['ExpiresAtUtc'], session.fresh_until.isoformat())

    def test_metadata_has_exact_schema_identity_and_no_producer_paths(self):
        status, data = self.session.get(self.session.prefix + 'model-contract.json.txt')
        self.assertEqual(status, 200)
        metadata = json.loads(data)
        self.assertEqual(metadata['Identity'], self.fixture.identity)
        self.assertEqual(len(metadata['Tables']), 46)
        self.assertNotIn('SourceRoot', metadata)
        self.assertNotIn('ProducerReceipts', metadata)
        self.assertEqual(metadata['Tables']['DimDetectedApplication.csv']['Columns'][-1], session_reader.PRODUCT_KEY)

    def test_model_product_projection_matches_retained_catalog(self):
        f = self.fixture
        app = f.row('DimDetectedApplication', TenantApplicationKey='a1', AppId='a1',
                    DisplayName='  Editor ', Publisher='Vendor', Platform='Windows', Version='1',
                    DeviceCount='0', ReportedDeviceCount='0', ExactRelatedDeviceCount='0',
                    RelationshipCoverageStatus='Complete')
        if 'ResolvedDeviceCount' in app:
            app.update(ResolvedDeviceCount='0', UnresolvedDeviceCount='0')
        f.write_table('DimDetectedApplication', [app])
        f.save_manifest()
        session = session_reader.ReadSession(f.load(), f.contract, utc_now=lambda: f.now)
        self.addCleanup(session.close)
        data = session.get(session.prefix + 'DimDetectedApplication.csv')[1]
        rows = list(reader._csv_rows(data, 'projection'))
        self.assertEqual(rows[0][session_reader.PRODUCT_KEY], session_reader.product_key(
            dict(DisplayName='editor', Publisher='vendor', Platform='windows')))

    def test_real_loopback_http_no_cache_browser_or_writes(self):
        server = session_reader.make_server(self.session)
        self.assertEqual(server.server_address[0], '127.0.0.1')
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            def request(method='GET', headers=None):
                connection = http.client.HTTPConnection('127.0.0.1', server.server_port, timeout=5)
                try:
                    connection.request(method, self.session.prefix + 'DimCountry.csv', headers=headers or {})
                    response = connection.getresponse()
                    return response.status, dict(response.getheaders()), response.read()
                finally:
                    connection.close()
            status, headers, _ = request()
            self.assertEqual(status, 200)
            self.assertIn('no-store', headers['Cache-Control'])
            self.assertEqual(request(headers={'Origin': 'https://example.invalid'})[0], 403)
            self.assertEqual(request(headers={'Host': 'example.invalid'})[0], 403)
            self.assertEqual(request('POST')[0], 501)
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)

    def test_invalid_lifetime_and_changed_contract_rejected(self):
        for ttl in (0, 7201, True):
            with self.assertRaises(ValueError):
                session_reader.ReadSession(self.fixture.load(), self.fixture.contract, ttl_seconds=ttl)
        contract = dict(self.fixture.contract, maxAgeHours=999)
        with self.assertRaisesRegex(ValueError, 'contract mismatch'):
            session_reader.ReadSession(self.fixture.load(), contract)


class ModelCandidateTests(unittest.TestCase):
    def test_dax_column_references_resolve_to_candidate_contract(self):
        contract = reader._json(reader.CONTRACT.read_bytes())
        columns = {x['name']: set(x['columns']) for x in contract['tables']}
        columns['DimDetectedApplication'].add(session_reader.PRODUCT_KEY)
        dax = (Path(__file__).parent / 'References/expected-measures.dax').read_text(encoding='utf-8')
        measures = re.findall(r"MEASURE\s+'([^']+)'\[([^]]+)\]", dax)
        for table, name in re.findall(r"'([^']+)'\[([^]]+)\]", dax):
            self.assertTrue(name in columns.get(table, set()) or (table, name) in measures,
                            'Unresolved candidate reference: ' + table + '[' + name + ']')
        self.assertEqual(len(measures), len(set(measures)))

    def test_seven_exact_schemas_and_three_hardware_columns(self):
        plan = model.model_plan()
        contract = reader._json(reader.CONTRACT.read_bytes())
        self.assertEqual({x['name'] for x in plan['tables']}, set(model.NEW_TABLES))
        for table in plan['tables']:
            original = next(x for x in contract['tables'] if x['name'] == table['name'])
            self.assertEqual([x['name'] for x in table['columns']], original['columns'])
            self.assertTrue(all(x['summarizeBy'] == 'None' for x in table['columns']))
        self.assertEqual(len(plan['hardwareColumns']), 3)
        self.assertEqual(len(plan['preserveCalculatedColumns']), 7)

    def test_relationship_endpoints_and_single_direction_no_cloud_loops(self):
        plan = model.model_plan()
        self.assertEqual(len(plan['relationships']), 3)
        self.assertTrue(all(x['isActive'] for x in plan['relationships']))
        contract = reader._json(reader.CONTRACT.read_bytes())
        columns = {x['name']: set(x['columns']) for x in contract['tables']}
        for relationship in plan['relationships']:
            self.assertIn(relationship['fromColumn'], columns[relationship['fromTable']])
            self.assertIn(relationship['toColumn'], columns[relationship['toTable']])
            self.assertEqual(relationship['crossFilteringBehavior'], 'OneDirection')
            if relationship['toTable'] in ('DimUser', 'DimDevice'):
                self.assertFalse(relationship['isActive'])

    def test_existing_application_bridge_uses_exact_prepared_key_lookup(self):
        updates = model.model_plan()['calculatedColumnUpdates']
        bridge = next(x for x in updates if x['name'] == 'Tenant Intune device key')
        self.assertIn('LOOKUPVALUE', bridge['expression'])
        self.assertIn("'DimIntuneManagedDevice'[ManagedDeviceId]", bridge['expression'])
        self.assertIn("'DimIntuneManagedDevice'[TenantKey]", bridge['expression'])
        self.assertNotIn('|intune-device|', bridge['expression'])
        label = next(x for x in updates if x['name'] == 'Application product')
        self.assertEqual(label['groupByColumns'], [session_reader.PRODUCT_KEY])

    def test_top_five_candidate_uses_grouped_device_counts(self):
        dax = (Path(__file__).parent / 'References/expected-measures.dax').read_text(encoding='utf-8')
        top = dax.split("[Application Top 5 candidate] =", 1)[1].split('// Global AD', 1)[0]
        self.assertIn('SUMMARIZECOLUMNS', top)
        self.assertIn('ALLSELECTED', top)
        self.assertNotIn('ADDCOLUMNS', top)

    def test_product_normalization_versions_unicode_and_delimiters(self):
        def key(name, publisher='Vendor', platform='Windows'):
            return session_reader.product_key(dict(DisplayName=name, Publisher=publisher, Platform=platform))
        self.assertEqual(key(' Editor '), key('editor'))
        self.assertEqual(key('Straße'), key('STRASSE'))
        self.assertNotEqual(key('a|b', 'c'), key('a', 'b|c'))
        self.assertNotEqual(key('editor', ''), key('editor'))
        self.assertNotEqual(key('editor', platform='iOS'), key('editor'))

    def test_streamed_m_reader_does_not_buffer_large_csvs(self):
        m = (Path(__file__).resolve().parents[1] / 'PowerBI/Queries/prepared-read-session.pq').read_text(encoding='utf-8')
        self.assertNotIn('Binary.Buffer(', m)
        self.assertIn('Table.RowCount(Rows) = Definition[Rows]', m)
        self.assertIn('Contract[BatchSHA256] = BatchSHA256', m)

    def test_unresolved_weekly_app_devices_retained_globally_not_guessed_into_fleet(self):
        apps = [dict(TenantApplicationKey='a', DisplayName='Editor', Publisher='Vendor', Platform='Windows')]
        links = [dict(TenantApplicationKey='a', ManagedDeviceId='unresolved')]
        self.assertEqual(list(model.application_counts(apps, links, [], []).values()), [1])
        self.assertEqual(model.application_counts(apps, links, [], [], ownership='Corporate'), {})

    def test_country_corporate_versions_device360_and_top5(self):
        devices = [dict(TenantDeviceKey='d1', CountryLabel='Alpha', Ownership='Corporate'),
                   dict(TenantDeviceKey='d2', CountryLabel='Beta', Ownership='Corporate'),
                   dict(TenantDeviceKey='d3', CountryLabel='Alpha', Ownership='Personal')]
        managed = [dict(TenantDeviceKey='d' + str(i), ManagedDeviceId='m' + str(i)) for i in (1, 2, 3)]
        apps = [dict(TenantApplicationKey='a' + str(i), DisplayName='Editor' if i < 3 else 'Other',
                     Publisher='Vendor', Platform='Windows', Version=str(i)) for i in (1, 2, 3)]
        links = [dict(TenantApplicationKey=a, ManagedDeviceId=m) for a, m in
                 [('a1', 'm1'), ('a2', 'm1'), ('a1', 'm2'), ('a3', 'm3')]]
        counts = model.application_counts(apps, links, managed, devices, ownership='Corporate')
        self.assertEqual(list(counts.values()), [2])  # Same product, two versions, two devices.
        alpha = model.application_counts(apps, links, managed, devices, countries={'Alpha'}, ownership='Corporate')
        self.assertEqual(list(alpha.values()), [1])
        self.assertEqual(model.application_counts(apps, links, managed, devices, selected_device='missing'), {})
        self.assertEqual(model.application_counts(apps, links, managed, devices, selected_device='d1'), alpha)
        self.assertEqual(len(model.top_five({str(i): 1 for i in range(7)})), 5)
        self.assertEqual(model.top_five({'b': 1, 'a': 1}), [('a', 1), ('b', 1)])

    def test_ad_without_dns_unresolved_disabled_servers_and_no_denominator(self):
        rows = [dict(DNSHostName='', Enabled=True, OperatingSystem='Windows 11', CoverageState='Managed in Intune'),
                dict(DNSHostName=' ', Enabled=True, OperatingSystem='Windows 10', CoverageState='No exact Entra SID match'),
                dict(DNSHostName='pc', Enabled=False, OperatingSystem='Windows 7', CoverageState='Entra only'),
                dict(DNSHostName='srv', Enabled=True, OperatingSystem='Windows Server 2022', CoverageState='Managed in Intune')]
        self.assertEqual(model.ad_summary(rows), dict(without_dns=2, enabled_windows_objects=2, managed_objects=1, coverage=0.5))
        self.assertIsNone(model.ad_summary([])['coverage'])

    def test_imported_coverage_country_ownership_duplicates_and_no_denominator(self):
        devices = [dict(TenantDeviceKey='d1', CountryLabel='Alpha', Ownership='Corporate'),
                   dict(TenantDeviceKey='d2', CountryLabel='Alpha', Ownership='Corporate'),
                   dict(TenantDeviceKey='d3', CountryLabel='Beta', Ownership='Personal')]
        managed = [dict(TenantDeviceKey='d1'), dict(TenantDeviceKey='d1'), dict(TenantDeviceKey='d3')]
        self.assertEqual(model.imported_coverage(devices, managed), 2 / 3)
        self.assertEqual(model.imported_coverage(devices, managed, countries={'Alpha'}, ownership='Corporate'), 0.5)
        self.assertEqual(model.imported_coverage(devices, managed, ownership='Personal'), 1)
        self.assertIsNone(model.imported_coverage(devices, managed, countries={'Missing'}))

    def test_missing_memory_is_not_zero_or_nonfinite(self):
        self.assertIsNone(model.memory_gib(''))
        self.assertEqual(model.memory_gib('16'), 16)
        for value in ('0', '-1', 'nan', 'inf'):
            with self.assertRaises(ValueError):
                model.memory_gib(value)

    def test_both_360_pages_require_one_tenant_key_not_display_identity(self):
        for field in ('TenantUserKey', 'TenantDeviceKey'):
            rows = [{field: 'example-test|one', 'DisplayName': 'Same name'},
                    {field: 'other-test|one', 'DisplayName': 'Same name'},
                    {field: '', 'DisplayName': 'Same name'}]
            self.assertEqual(model.selected_rows(rows, field, ['example-test|one']), [rows[0]])
            self.assertEqual(model.selected_rows(rows, field, ['not-present']), [])
            self.assertIsNone(model.selected_rows(rows, field, []))
            self.assertIsNone(model.selected_rows(rows, field, ['example-test|one', 'other-test|one']))


if __name__ == '__main__':
    unittest.main()
