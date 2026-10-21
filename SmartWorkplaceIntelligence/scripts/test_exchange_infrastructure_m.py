"""Synthetic PQTest checks for the governed Exchange infrastructure reader.

PREPARED_READERS_FACTORY_PATH optionally tests a staged, MCP-exported factory
before it is serialized to the checkout. No model or production data is read.
"""
import csv
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

from test_prepared_latest_m import binary, expression, m_text

PQTEST = os.environ.get("PQTEST_PATH") or shutil.which("PQTest.exe")
SOURCE = "Source-Exchange_OnPrem_Servers_Inventory.csv"


def csv_bytes(columns, rows):
    stream = io.StringIO(newline="")
    writer = csv.writer(stream, delimiter=";", lineterminator="\n")
    writer.writerow(columns)
    writer.writerows(rows)
    return stream.getvalue().encode("utf-8")


def factory():
    staged = os.environ.get("PREPARED_READERS_FACTORY_PATH")
    return Path(staged).read_text(encoding="utf-8-sig") if staged else expression("CreatePreparedReaders")


def fixture(files=None, *, pin="Latest", change=None, duplicate=False,
            missing_payload=False, bad_size=False, tenant="synthetic", batch="batch-a"):
    files = files or {}
    pointer = {"SchemaVersion": 1, "BatchId": "batch-a", "TenantKey": "synthetic"}
    artifacts = [{"File": name, "Bytes": len(payload) + int(bad_size)} for name, payload in files.items()]
    if duplicate and artifacts:
        artifacts.append(dict(artifacts[0]))
    manifest = {"SchemaVersion": 1, "BatchId": batch, "TenantKey": tenant,
                "Files": [], "Artifacts": artifacts}
    root = '#table({"Name","Content"}, {{"current.json.txt",' + binary(pointer) + '}})'
    changed = binary(json.dumps(pointer).encode() + b" ")
    payload = [] if missing_payload else [(name, content) for name, content in files.items()]
    content_rows = [('batch.json.txt', json.dumps(manifest).encode())] + payload
    selected = '#table({"Name","Content"}, {' + ','.join('{' + m_text(name) + ',' + binary(value) + '}'
                                                            for name, value in content_rows) + '})'
    reader = '(segments as list, phase as text) as table => if List.IsEmpty(segments) then '
    if change:
        reader += 'if phase = ' + m_text(change) + ' then #table({"Name","Content"}, {{"current.json.txt",' + changed + '}}) else '
    reader += root + ' else ' + selected
    return ["PreparedBatchId = " + m_text(pin), 'PreparedExpectedTenantKey = "synthetic"',
            'PreparedReadFolder = ' + reader,
            'NativeEvidenceTypes = (Rows as table, Types as list) as table => Rows',
            'CreatePreparedReaders = ' + factory(), 'Readers = CreatePreparedReaders()',
            'Infrastructure = Readers[GetExchangeOnPremInfrastructure]()']


@unittest.skipUnless(PQTEST, "Set PQTEST_PATH for actual M evaluation")
class ExchangeInfrastructureMTests(unittest.TestCase):
    def evaluate(self, bindings, result):
        query = 'let\n' + ',\n'.join(bindings) + ',\nResult = ' + result + '\nin Result'
        with tempfile.TemporaryDirectory() as scratch:
            path = Path(scratch) / 'infrastructure.query.pq'
            path.write_text(query, encoding='utf-8')
            run = subprocess.run([PQTEST, 'run-test', '--queryFile', str(path)],
                                 capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=180)
        self.assertEqual(run.returncode, 0, run.stdout[-3500:] + run.stderr[-1000:])
        self.assertIn('"Status":"Passed"', run.stdout.replace(' ', ''), run.stdout[-3500:])
        self.assertIn('"Value":true', run.stdout.replace(' ', ''), run.stdout[-3500:])

    def inventory(self, tenant='synthetic', extra=True):
        columns = ['TenantKey', 'Name', 'AdminDisplayVersion', 'ServerRole']
        values = [tenant, 'mail-one', 'Version 15.1', 'Mailbox']
        if extra:
            columns += ['FutureNumeric', 'FutureText', 'FutureBlank']
            values += ['64', 'quoted; value', '']
        return {SOURCE: csv_bytes(columns, [values])}

    def test_missing_group_has_twelve_explicit_unknown_rows(self):
        self.evaluate(fixture(), '''[OK = Table.RowCount(Infrastructure) = 12 and
            List.AllTrue(List.Transform(Infrastructure[Source Available], each _ = false)) and
            List.AllTrue(List.Transform(Infrastructure[Numeric Value], each _ = null)) and
            List.Distinct(Infrastructure[Loaded Batch ID]) = {"batch-a"}]''')

    def test_semicolon_unknown_columns_quoting_and_lineage_are_retained(self):
        self.evaluate(fixture(self.inventory()), '''let
            Rows = Table.SelectRows(Infrastructure, each [Dataset] = "Servers"),
            Numeric = Table.SelectRows(Rows, each [Attribute] = "FutureNumeric"){0},
            Quoted = Table.SelectRows(Rows, each [Attribute] = "FutureText"){0},
            Blank = Table.SelectRows(Rows, each [Attribute] = "FutureBlank"){0}
            in [OK = Table.RowCount(Rows) = 7 and Numeric[Numeric Value] = 64 and
                Quoted[Value] = "quoted; value" and Blank[Numeric Value] = null and
                List.AllTrue(Rows[Source Available]) and List.Distinct(Rows[Server Name]) = {"mail-one"} and
                List.Distinct(Rows[Loaded Batch ID]) = {"batch-a"} and
                List.Count(List.Distinct(Rows[Source Row Key])) = 1]''')

    def test_empty_captured_csv_is_distinct_from_absence(self):
        source = {SOURCE: csv_bytes(['TenantKey', 'Name', 'AdminDisplayVersion', 'ServerRole'], [])}
        self.evaluate(fixture(source), '''let R = Table.SelectRows(Infrastructure, each [Dataset] = "Servers"){0}
            in [OK = R[Source Available] = true and R[Collection Status] = "Captured - no rows" and R[Numeric Value] = null]''')

    def test_foreign_blank_or_missing_tenant_and_bad_header_are_rejected(self):
        invalid = [self.inventory('foreign'), self.inventory(''),
                   {SOURCE: csv_bytes(['Name', 'AdminDisplayVersion', 'ServerRole'], [['mail-one', '15.1', 'Mailbox']])},
                   {SOURCE: csv_bytes(['TenantKey', 'TenantKey', 'Name', 'AdminDisplayVersion', 'ServerRole'],
                                      [['synthetic', 'synthetic', 'mail-one', '15.1', 'Mailbox']])},
                   {SOURCE: csv_bytes(['TenantKey', 'AdminDisplayVersion', 'ServerRole'], [['synthetic', '15.1', 'Mailbox']])}]
        for files in invalid:
            with self.subTest(columns=list(files)):
                self.evaluate(fixture(files), '[OK = (try Table.RowCount(Infrastructure))[HasError]]')

    def test_declared_duplicate_missing_payload_size_or_manifest_identity_is_rejected(self):
        for kwargs in ({'duplicate': True}, {'missing_payload': True}, {'bad_size': True},
                       {'tenant': 'foreign'}, {'batch': 'batch-b'}):
            with self.subTest(case=kwargs):
                self.evaluate(fixture(self.inventory(), **kwargs), '[OK = (try Table.RowCount(Infrastructure))[HasError]]')

    def test_latest_checks_pointer_during_payload_and_after_unpivot(self):
        for phase in ('verify-payload', 'verify-infrastructure'):
            with self.subTest(phase=phase):
                self.evaluate(fixture(self.inventory(), change=phase), '[OK = (try Table.RowCount(Infrastructure))[HasError]]')

    def test_explicit_pin_does_not_follow_pointer_movement(self):
        self.evaluate(fixture(self.inventory(), pin='batch-a', change='verify-payload'),
                      '[OK = Table.RowCount(Infrastructure) = 18 and List.Distinct(Infrastructure[Loaded Batch ID]) = {"batch-a"}]')

    def test_error_status_in_summary_is_preserved(self):
        name = 'Source-Exchange_OnPrem_Infrastructure_PerServerSummary.csv'
        content = csv_bytes(['TenantKey', 'ExchangeServerName', 'ComputeCollectionStatus',
                             'DiskDriveCollectionStatus', 'LogicalDiskCollectionStatus'],
                            [['synthetic', 'mail-one', 'ERROR', 'OK', 'OK']])
        self.evaluate(fixture({name: content}), '''let R = Table.SelectRows(Infrastructure, each [Dataset] = "Per-server summary")
            in [OK = List.Distinct(R[Collection Status]) = {"ERROR"} and List.AllTrue(R[Source Available])]''')

    def test_numeric_projection_does_not_reinterpret_localized_or_unit_text(self):
        content = csv_bytes(['TenantKey', 'Name', 'AdminDisplayVersion', 'ServerRole',
                             'Localized', 'WithUnits', 'NonFinite', 'Scientific'],
                            [['synthetic', 'mail-one', 'Version 15.1', 'Mailbox', '16,5', '64 GB', 'NaN', '1.5e2']])
        self.evaluate(fixture({SOURCE: content}), '''let
            R = Table.SelectRows(Infrastructure, each [Dataset] = "Servers"),
            Invalid = Table.SelectRows(R, each List.Contains({"Localized","WithUnits","NonFinite"}, [Attribute])),
            Scientific = Table.SelectRows(R, each [Attribute] = "Scientific"){0}
            in [OK = List.AllTrue(List.Transform(Invalid[Numeric Value], each _ = null)) and
                Scientific[Numeric Value] = 150]''')


if __name__ == '__main__':
    unittest.main()
