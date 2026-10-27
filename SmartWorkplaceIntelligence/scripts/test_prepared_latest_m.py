"""Synthetic PQTest regressions for the actual checkout readers; never refresh a model."""
import base64
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import textwrap
import unittest

DEFINITION = Path(__file__).resolve().parents[1] / "pbip" / "SmartWorkplaceIntelligence.SemanticModel" / "definition"
PQTEST = os.environ.get("PQTEST_PATH") or shutil.which("PQTest.exe")


def expression(name):
    source = (DEFINITION / "expressions.tmdl").read_text(encoding="utf-8-sig")
    match = re.search(r"^expression " + re.escape(name) + r" =([^\n]*(?:\n(?!\tlineageTag:|expression |///).*)*)", source, re.M)
    if not match:
        raise AssertionError("Missing expression: " + name)
    # Stop at TOM metadata (new expressions may have no lineageTag).
    body = re.split(r"\n\t(?:lineageTag:|annotation |queryGroup:)", match.group(1))[0]
    return textwrap.dedent(body.strip("\r\n")).strip()


def m_text(value):
    return '"' + value.replace("#", "#(#)").replace('"', '""').replace("\n", "#(lf)").replace("\r", "#(cr)") + '"'


def binary(value):
    if not isinstance(value, bytes):
        value = json.dumps(value).encode()
    return 'Binary.FromText("' + base64.b64encode(value).decode() + '", BinaryEncoding.Base64)'


def fixture(pin="Latest", pointer=None, manifest=None, change=None, missing=False,
            ambiguous=False, legacy=False, broken_preferred=False, size=7, second_batch="batch-b"):
    pointer = pointer or {"SchemaVersion": 1, "BatchId": "batch-a", "TenantKey": "synthetic"}
    manifest = manifest or {"SchemaVersion": 1, "BatchId": "batch-a", "TenantKey": "synthetic",
                            "Files": [{"File": "test.csv", "Bytes": size}],
                            "Artifacts": [{"File": "Source-test.csv", "Bytes": size}]}
    selected = (b"not-json" if broken_preferred else json.dumps(pointer).encode())
    root_rows = [] if missing else [('current.json' if legacy else 'current.json.txt', binary(selected))]
    if ambiguous:
        root_rows *= 2
    if broken_preferred:
        root_rows.append(("current.json", binary(pointer)))
    root_table = '#table({"Name","Content"}, {' + ",".join("{" + m_text(n) + "," + b + "}" for n, b in root_rows) + "})"
    changed = dict(pointer, BatchId=second_batch)
    # A whitespace-only change must also be rejected.
    changed_bytes = (json.dumps(pointer) + " ").encode() if change == "bytes" else json.dumps(changed).encode()
    phase = "verify-manifest" if change == "manifest" else "verify-payload"
    folder = '''(segments as list, phase as text) as table =>
        if List.IsEmpty(segments) then
            if CHANGE and phase = PHASE then #table({"Name","Content"}, {{"current.json.txt", CHANGED}})
            else ROOT
        else #table({"Name","Content"}, {{"batch.json.txt", MANIFEST},
            {"test.csv", Text.ToBinary("x#(lf)1#(lf)2#(lf)3")}, {"Source-test.csv", Text.ToBinary("x#(lf)1#(lf)2#(lf)3")}})'''
    for key, value in {"CHANGE": "true" if change else "false", "PHASE": m_text(phase), "CHANGED": binary(changed_bytes),
                       "ROOT": root_table, "MANIFEST": binary(manifest)}.items():
        folder = folder.replace(key, value)
    return [
        "PreparedBatchId = " + m_text(pin),
        'PreparedExpectedTenantKey = "synthetic"',
        "PreparedReadFolder = " + folder,
        "NativeEvidenceTypes = " + expression("NativeEvidenceTypes"),
        "CreatePreparedReaders = " + expression("CreatePreparedReaders"),
        "Readers = CreatePreparedReaders()",
    ]


@unittest.skipUnless(PQTEST, "Set PQTEST_PATH for actual M evaluation")
class PreparedLatestMTests(unittest.TestCase):
    def evaluate(self, bindings, result):
        query = "let\n" + ",\n".join(bindings) + ",\nResult = " + result + "\nin Result"
        with tempfile.TemporaryDirectory() as scratch:
            path = Path(scratch) / "latest.query.pq"
            path.write_text(query, encoding="utf-8")
            run = subprocess.run([PQTEST, "run-test", "--queryFile", str(path)],
                                 capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
        self.assertEqual(run.returncode, 0, run.stdout[-3500:] + run.stderr[-1000:])
        self.assertIn('"Status":"Passed"', run.stdout.replace(" ", ""), run.stdout[-3500:])
        self.assertIn('"Value":true', run.stdout.replace(" ", ""), run.stdout[-3500:])

    def test_latest_selects_pointer_and_labels_imported_rows(self):
        self.evaluate(fixture(), '''[OK =
            Readers[GetPreparedBatchManifest]()[BatchId] = "batch-a" and
            Table.First(Readers[Finish](Csv.Document(Readers[GetPreparedEvidenceFile]("test.csv"))))[Loaded Batch ID] = "batch-a"]''')

    def test_pin_overrides_pointer_and_ignores_pointer_movement(self):
        self.evaluate(fixture(pin="batch-a", pointer={"SchemaVersion": 1, "BatchId": "batch-b", "TenantKey": "synthetic"}, change="payload"),
                      '[OK = Readers[GetPreparedBatchManifest]()[BatchId] = "batch-a" and Binary.Length(Readers[GetPreparedEvidenceFile]("test.csv")) = 7]')

    def test_pointer_changes_during_manifest_read(self):
        self.evaluate(fixture(change="manifest"), '[OK = (try Readers[GetPreparedBatchManifest]())[HasError]]')

    def test_pointer_changes_during_payload_read_including_source_artifacts(self):
        for name in ("test.csv", "Source-test.csv"):
            with self.subTest(name=name):
                self.evaluate(fixture(change="payload"),
                              '[OK = (try Binary.Length(Readers[GetPreparedEvidenceFile](' + m_text(name) + ')))[HasError]]')

    def test_byte_change_with_same_identity_is_rejected(self):
        self.evaluate(fixture(change="bytes"), '[OK = (try Binary.Length(Readers[GetPreparedEvidenceFile]("test.csv")))[HasError]]')

    def test_wrong_tenant_manifest_batch_schema_or_size(self):
        cases = [
            fixture(pointer={"SchemaVersion": 1, "BatchId": "batch-a", "TenantKey": "wrong"}),
            fixture(manifest={"SchemaVersion": 1, "BatchId": "batch-b", "TenantKey": "synthetic", "Files": []}),
            fixture(manifest={"SchemaVersion": 1, "BatchId": "batch-a", "TenantKey": "wrong", "Files": []}),
            fixture(pointer={"SchemaVersion": 2, "BatchId": "batch-a", "TenantKey": "synthetic"}),
            fixture(size=8),
        ]
        for i, bindings in enumerate(cases):
            with self.subTest(case=i):
                self.evaluate(bindings, '[OK = (try Binary.Length(Readers[GetPreparedEvidenceFile]("test.csv")))[HasError]]')

    def test_absent_ambiguous_or_unsafe_pointer(self):
        for kwargs in ({"missing": True}, {"ambiguous": True}, {"pin": ""},
                       {"pointer": {"SchemaVersion": 1, "BatchId": "../escape", "TenantKey": "synthetic"}}):
            with self.subTest(case=kwargs):
                self.evaluate(fixture(**kwargs), '[OK = (try Readers[GetPreparedBatchManifest]())[HasError]]')

    def test_legacy_only_on_absence_not_parse_error(self):
        self.evaluate(fixture(legacy=True), '[OK = Readers[GetPreparedBatchManifest]()[BatchId] = "batch-a"]')
        self.evaluate(fixture(broken_preferred=True), '[OK = (try Readers[GetPreparedBatchManifest]())[HasError]]')

    def test_mailbox_factory_keeps_stability_guard(self):
        self.evaluate(fixture(change="payload"), '[OK = (try Binary.Length(Readers[CreateMailboxMigrationReader]()("Source-test.csv")))[HasError]]')

    def test_strict_native_projection_precedes_lineage_and_keeps_errors(self):
        self.evaluate(fixture(), '''let
            Types = {[name="Count", dataType="Int64"]},
            Good = Readers[Finish](NativeEvidenceTypes(#table({"Count"}, {{"42"}}), Types)),
            Bad = Readers[Finish](NativeEvidenceTypes(#table({"Count"}, {{"not-a-number"}}), Types))
            in [OK = Good{0}[Count] = 42 and Good{0}[Loaded Batch ID] = "batch-a"
                and (try Bad{0}[Count])[HasError]]''')

    def test_sharepoint_pointer_transport_bypasses_cache_and_falls_back_only_on_404(self):
        for absent in (False, True):
            bindings = [
                'PreparedSourceMode = "SharePoint"',
                'SharePointSiteUrl = "https://example.invalid"',
                'SharePointLibraryName = "Documents"',
                'SharePointDataFolder = "DATA"',
                'PreparedSharePointFolder = "DATA-POWERBI"',
                'Web.Contents = (base as text, options as record) as binary => '
                'if base <> SharePointSiteUrl or options[IsRetry] <> true then error "Cached or dynamic source" '
                'else if options[RelativePath] = "_api/web/lists" then Json.FromValue([value={[Title="Documents", RootFolder=[Name="Shared Documents", ServerRelativeUrl="/Shared Documents"]]}]) '
                'else if ' + ('true' if absent else 'false') + ' and Text.Contains(options[RelativePath], "current.json.txt") '
                'then Value.ReplaceMetadata(Text.ToBinary(""), [Response.Status=404]) '
                'else Text.ToBinary("pointer")',
                'PreparedReadFolder = ' + expression("PreparedReadFolder"),
                'Root = PreparedReadFolder({}, "verify-payload")',
            ]
            expected = "current.json" if absent else "current.json.txt"
            self.evaluate(bindings, '[OK = Root{0}[Name] = ' + m_text(expected) + ' and Text.FromBinary(Root{0}[Content]) = "pointer"]')
        bindings[5] = 'Web.Contents = (base as text, options as record) as binary => error "HTTP 403"'
        self.evaluate(bindings, '[OK = (try Root{0}[Content])[HasError]]')

    def test_library_navigation_name_can_differ_from_rest_title(self):
        adapters = [expression("PreparedReadFolder"),
                    (DEFINITION.parents[2] / "powerquery/PreparedReadFolder.SharePoint.pq").read_text()]
        for adapter in adapters:
            for count in (0, 1, 2):
                for paged in (False, True):
                    with self.subTest(adapter="service" if "PreparedSourceMode" not in adapter else "checkout", count=count, paged=paged):
                        bindings = [
                            'PreparedSourceMode = "SharePoint"',
                            'SharePointSiteUrl = "https://example.invalid"',
                            'SharePointLibraryName = "Shared Documents"',
                            'SharePointDataFolder = "DATA"',
                            'PreparedSharePointFolder = "DATA-POWERBI"',
                            'Web.Contents = (base as text, options as record) as binary => '
                            'if options[RelativePath] = "_api/web/lists" then Json.FromValue('
                            '[value=List.Repeat({[Title="Renamed Library", RootFolder=[Name="Shared Documents", ServerRelativeUrl="/Shared Documents"]]}, '
                            + str(count) + ')' + (', #"odata.nextLink"="not-followed"' if paged else '') + ']) '
                            'else if Text.Contains(options[RelativePath], "Shared%20Documents") then Text.ToBinary("pointer") else error "Wrong folder"',
                            'PreparedReadFolder = ' + adapter,
                            'Attempt = try Text.FromBinary(PreparedReadFolder({}){0}[Content])',
                        ]
                        result = '[OK = not Attempt[HasError] and Attempt[Value] = "pointer"]' if count == 1 and not paged else '[OK = Attempt[HasError]]'
                        self.evaluate(bindings, result)

    def test_two_query_results_expose_mixture_and_empty_table_limit(self):
        bindings = fixture()
        other = fixture(pointer={"SchemaVersion": 1, "BatchId": "batch-b", "TenantKey": "synthetic"},
                        manifest={"SchemaVersion": 1, "BatchId": "batch-b", "TenantKey": "synthetic", "Files": []})
        # Use the same production factory with a second injected storage view.
        other_folder = other[2].split(" = ", 1)[1]
        bindings += ["Other = CreatePreparedReaders(" + other_folder + ")",
                     'A = Readers[Finish](#table({"Value"}, {{1}, {2}}))',
                     'B = Other[Finish](#table({"Value"}, {{3}}))',
                     'Empty = Other[Finish](#table(type table [Value = number], {}))']
        self.evaluate(bindings, '''[OK = List.Count(List.Distinct(Table.Column(Table.Combine({A,B}), "Loaded Batch ID"))) = 2
            and Table.RowCount(Empty) = 0 and List.Contains(Table.ColumnNames(Empty), "Loaded Batch ID")]''')


def has_current_batch_provenance(table, source, expressions):
    """Conservative static contract for the returned partition, not a dead Finish call."""
    source = re.sub(r'\n\tannotation PBI_ResultType = (?:Table|Exception)\s*$', '', source)
    if table == 'Exchange On-prem Infrastructure':
        if not re.search(r'source = let Readers = CreatePreparedReaders\(\) in Readers\[GetExchangeOnPremInfrastructure\]\(\)\s*$', source):
            return False
        match = re.search(r'GetExchangeOnPremInfrastructure = \(\) as table =>(.*?)GetExchangeHybridFreeBusy =', expressions, re.S)
        if not match:
            return False
        body = match[1]
        # Both placeholder and observed rows must carry the selected batch, with stability checking.
        placeholder, observed = body.split('ReadSource =', 1)
        return all(token in body for token in [
            '#"Loaded Batch ID"=text',
            'Result = Table.Buffer(Table.FromRecords(List.Combine(List.Transform(Sources, ReadSource)), Schema, MissingField.Error))',
            'if ForcedRows >= 0 and CheckStable("verify-infrastructure") then Result else error'
        ]) and '#"Loaded Batch ID"=BatchId' in placeholder and '#"Loaded Batch ID"=BatchId' in observed
    expected = {
        'User Inventory Evidence': r'Table\.AddColumn\(Readers\[Finish\]\(BusinessRows\), "TenantKey", each PreparedExpectedTenantKey, type text\)',
        'Data Trust Evidence': r'Readers\[Finish\]\(Combined\)',
        'Exchange Mailbox Databases': r'if Forced >= 0 then Readers\[Finish\]\(Buffered\) else error "Database observations unavailable"',
        'Exchange Mailbox Server Reference': r'Readers\[Finish\]\(Display\)',
        'Exchange Health Checker Findings': r'if Forced >= 0 then Readers\[Finish\]\(BusinessRows\) else error "Health Checker findings unavailable"',
        'Exchange Health Checker Servers': r'if Forced >= 0 then Readers\[Finish\]\(BusinessRows\) else error "Health Checker server observations unavailable"',
    }.get(table, r'Readers\[Finish\]\(BusinessRows\)')
    return bool(re.search(r'\bin\s+' + expected + r'\s*$', source))


class PreparedLatestModelTests(unittest.TestCase):
    def test_provenance_is_hidden_not_a_business_key_or_linguistic_field(self):
        monitored = []
        expressions = (DEFINITION / "expressions.tmdl").read_text(encoding="utf-8-sig")
        for path in (DEFINITION / "tables").glob("*.tmdl"):
            source = path.read_text(encoding="utf-8-sig")
            if "Readers = CreatePreparedReaders()" not in source:
                continue
            # Derived mailbox projections retain provenance from their already monitored inputs.
            # Their nested transport validates detail artifacts; restamping would hide mixed inputs.
            if path.stem in {'Mailbox Exploration Inventory','Mailbox Delegation Detail','Mailbox License Membership'}:
                self.assertIn('Output=if not BatchSafe then error "Mailbox exploration requires one common prepared batch"', source)
                self.assertIn('"Loaded Batch ID"', source)
                continue
            monitored.append(path.stem)
            column = re.search(r"\tcolumn 'Loaded Batch ID'\n(.*?)(?=\n\t(?:partition|column|measure)|\Z)", source, re.S)
            self.assertIsNotNone(column, path.name)
            self.assertIn("\n\t\tisHidden", column[1])
            self.assertIn("dataType: string", column[1])
            self.assertIn("summarizeBy: none", column[1])
            self.assertNotRegex(column[1], r"\n\t\t(?:isKey|isUnique)(?:\s|$)")
            self.assertTrue(has_current_batch_provenance(path.stem, source, expressions), path.name)
        self.assertGreater(len(monitored), 100)
        audit = (DEFINITION / "tables/Loaded Batch Audit.tmdl").read_text(encoding="utf-8-sig")
        rows = re.findall(r'ROW\("Table Name", "([^"]+)"', audit)
        readiness = {'Hybrid Identity Findings','Device Readiness Gap Evidence','Readiness Source Status'}
        for name in readiness:
            source = (DEFINITION / ('tables/' + name + '.tmdl')).read_text(encoding='utf-8-sig')
            self.assertIn('R=CreatePreparedReaders()', source)
            self.assertIn('R[GetReadinessSource]', source)
            if name == 'Readiness Source Status':
                self.assertIn('#"Loaded Batch ID"=V[BatchId]', source)
            else:
                self.assertRegex(source, r'in R\[Finish\]\(')
        self.assertEqual(sorted(rows), sorted(set(monitored) | readiness))
        self.assertNotIn("Loaded Batch ID", (DEFINITION / "relationships.tmdl").read_text(encoding="utf-8-sig"))
        for path in (DEFINITION / "cultures").glob("*.tmdl"):
            self.assertNotIn("Loaded Batch", path.read_text(encoding="utf-8-sig"))

    def test_provenance_contract_rejects_unstamped_returns(self):
        expressions = (DEFINITION / "expressions.tmdl").read_text(encoding="utf-8-sig")
        for name, call, replacement in [
            ('AD Computer Evidence', 'Readers[Finish](BusinessRows)', 'BusinessRows'),
            ('Data Trust Evidence', 'Readers[Finish](Combined)', 'Combined'),
            ('Exchange Health Checker Findings', 'Readers[Finish](BusinessRows)', 'BusinessRows'),
            ('Exchange Health Checker Servers', 'Readers[Finish](BusinessRows)', 'BusinessRows'),
            ('Exchange Mailbox Databases', 'Readers[Finish](Buffered)', 'Buffered'),
            ('Exchange Mailbox Server Reference', 'Readers[Finish](Display)', 'Display'),
            ('Exchange On-prem Infrastructure', 'Readers[GetExchangeOnPremInfrastructure]()', 'Readers[GetOtherReader]()'),
        ]:
            source = (DEFINITION / ('tables/' + name + '.tmdl')).read_text(encoding='utf-8-sig')
            with self.subTest(table=name):
                self.assertTrue(has_current_batch_provenance(name, source, expressions))
                self.assertFalse(has_current_batch_provenance(name, source.replace(call, replacement), expressions))
        source = (DEFINITION / 'tables/Exchange On-prem Infrastructure.tmdl').read_text(encoding='utf-8-sig')
        for token in ['#"Loaded Batch ID"=BatchId', 'CheckStable("verify-infrastructure")']:
            with self.subTest(missing=token):
                self.assertFalse(has_current_batch_provenance('Exchange On-prem Infrastructure', source, expressions.replace(token, 'null')))
        self.assertFalse(has_current_batch_provenance('Example', 'let Unused = Readers[Finish](BusinessRows) in BusinessRows', expressions))

    def test_service_adapter_has_no_local_transport_and_keeps_fresh_reads(self):
        adapter = (DEFINITION.parents[2] / "powerquery/PreparedReadFolder.SharePoint.pq").read_text()
        self.assertNotRegex(adapter, r"\b(?:File|Folder)\.Contents|PreparedEvidenceRoot|PreparedSourceMode")
        self.assertIn("IsRetry=true", adapter)
        self.assertIn("ManualStatusHandling={404}", adapter)


if __name__ == "__main__":
    unittest.main()
