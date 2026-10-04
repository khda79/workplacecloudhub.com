"""Offline contract checks for ShareGate CSV diagnostics."""

__version__ = "1.0.5"

import csv
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import shutil
import unittest
import uuid


SCRIPT = Path(__file__).resolve().parents[1] / "Scripts" / "Diagnostics" / "analyze_sharegate_reports.py"
SPEC = importlib.util.spec_from_file_location("sharegate_diagnostics", SCRIPT)
DIAG = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DIAG)


class ShareGateDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).resolve().parent / (".diag-test-" + uuid.uuid4().hex)
        self.assertTrue(self.root.is_relative_to(Path(__file__).resolve().parent))
        self.root.mkdir()
        self.addCleanup(shutil.rmtree, self.root)
        self.project = self.root / "Migration"
        self.project.mkdir()
        self.output = self.root / "output"
        self.config = self.root / "Config"
        self.config.mkdir()
        self.old_config = DIAG.CONFIG
        DIAG.CONFIG = self.config
        self.addCleanup(setattr, DIAG, "CONFIG", self.old_config)
        aliases = {
            "SessionId": ["Session ID"], "RowId": ["ID"], "Timestamp": ["Date"],
            "Status": ["Status"], "ObjectType": ["Type"], "ItemName": ["Title"],
            "SourceUrl": ["Source site address"], "SourceList": ["Source list title"],
            "SourceItemId": ["Source ID"], "DestinationUrl": ["Destination site address"], "Messages": ["Messages"],
            "Warnings": ["Warnings"], "Errors": ["Errors"], "Details": ["Details"],
            "CopyOptions": ["Copy options"], "HelpLinks": ["Help links"]
        }
        rules = {"Rules": [
            {"Id": "SG-ACCESS-SOURCE", "Regex": "(?i)source access denied", "Category": "Source access",
             "Severity": "Error", "ProbableCause": "Source", "Remediation": "Review source",
             "ShareGateAction": "Review", "DefaultState": "To fix"},
            {"Id": "SG-ACCESS-TARGET", "Regex": "(?i)destination access denied", "Category": "Destination access",
             "Severity": "Error", "ProbableCause": "Destination", "Remediation": "Review destination",
             "ShareGateAction": "Review", "DefaultState": "To fix"},
            {"Id": "SG-ACCESS-UNKNOWN", "Regex": "(?i)(?:access denied|unauthori[sz]ed|not authorized)", "Category": "Access side undetermined",
             "Severity": "Error", "ProbableCause": "Unknown side", "Remediation": "Review both sides",
             "ShareGateAction": "Review", "DefaultState": "To fix"},
            {"Id": "FEATURE", "Regex": "(?i)feature unavailable", "Category": "Unavailable site feature",
             "Severity": "Warning", "ProbableCause": "Not available", "Remediation": "Review feature",
             "ShareGateAction": "Manual", "DefaultState": "Accepted"}
        ]}
        (self.config / "sharegate-diagnostics.columns.json.template").write_text(json.dumps(aliases), encoding="utf-8")
        (self.config / "sharegate-diagnostics.rules.json.template").write_text(json.dumps(rules), encoding="utf-8")
        self.report = self.root / "report.csv"
        fields = ["Session ID", "ID", "Date", "Status", "Type", "Title", "Source site address",
                  "Source list title", "Source ID", "Messages", "Warnings", "Errors", "Details", "Copy options", "Help links"]
        records = [
            ["s1", "1", "2026-10-02", "Error", "File", "<script>", "https://source.example/sites/a",
             "Documents", "42", "", "", "Access denied for <script>", "", "Incremental", "https://help.example/item"],
            ["s1", "2", "2026-10-02", "Warning", "File version", "File version", "https://source.example/sites/a",
             "Documents", "42", "", "Version warning", "", "", "Incremental", ""],
            ["s1", "3", "2026-10-02", "Warning", "Site feature", "Feature", "https://source.example/sites/a",
             "", "", "", "Feature unavailable", "", "", "", ""],
            ["s1", "4", "2026-10-02", "Success", "File", "Other file", "https://source.example/sites/a",
             "Documents", "43", "Copied", "", "", "", "Incremental", ""]
        ]
        with self.report.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(fields)
            writer.writerows(records)

    def test_duplicate_rows_item_versions_and_persistent_state(self):
        DIAG.analyze([self.report, self.report], self.output, self.project)
        summary = json.loads((self.output / "Summary.json.txt").read_text(encoding="utf-8"))
        self.assertEqual(summary["Lines"], 4)
        self.assertEqual(summary["DuplicateRowsSuppressed"], 4)
        self.assertEqual(summary["DistinctItems"], 2)
        self.assertEqual(summary["UnkeyedRows"], 1)
        self.assertEqual(summary["IssueLineState"], {"To fix": 2, "Accepted": 1})
        self.assertEqual(summary["ResidualLineRate"], 66.67)
        self.assertEqual(summary["ResidualItemRate"], 50.0)
        self.assertTrue((self.config / "sharegate-diagnostics.columns.json.txt").exists())
        self.assertTrue((self.config / "sharegate-diagnostics.rules.json.txt").exists())
        self.assertIn("&lt;script&gt;", (self.output / "MigrationDiagnostics-Report.html").read_text(encoding="utf-8"))
        access = next(value for value in summary["Patterns"] if value["RuleId"] == "SG-ACCESS-UNKNOWN")
        DIAG.set_state(self.project, access["PatternKey"], "Accepted")
        with self.assertRaisesRegex(ValueError, "another operator"):
            DIAG.set_state(self.project, access["PatternKey"], "Fixed", expected="To fix")
        DIAG.analyze([self.report], self.root / "second", self.project)
        updated = json.loads((self.root / "second" / "Summary.json.txt").read_text(encoding="utf-8"))
        self.assertEqual(updated["SelectedSessionId"], "")
        self.assertEqual(updated["InputEvidence"], [{
            "Path": str(self.report), "Size": self.report.stat().st_size,
            "LastWriteUtc": updated["InputEvidence"][0]["LastWriteUtc"],
            "Sha256": hashlib.sha256(self.report.read_bytes()).hexdigest(),
        }])
        self.assertEqual(updated["IssueLineState"], {"Accepted": 2, "To fix": 1})
        self.assertEqual(updated["ResidualLineRate"], 50.0)

    def test_raw_columns_and_manual_actions_are_exported(self):
        DIAG.analyze([self.report], self.output, self.project)
        summary = json.loads((self.output / "Summary.json.txt").read_text(encoding="utf-8"))
        self.assertEqual(summary["ManualActionCategories"], {"Unavailable site feature": 1})
        self.assertNotIn("Access side undetermined", summary["ManualActionCategories"])
        with (self.output / "ClassifiedRows.csv").open(encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.DictReader(stream))
        self.assertEqual(rows[0]["Raw: Copy options"], "Incremental")
        self.assertEqual(rows[0]["HelpLinks"], "https://help.example/item")
        with (self.output / "Remediation-Actions.csv").open(encoding="utf-8-sig", newline="") as stream:
            actions = list(csv.DictReader(stream))
        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["State"], "Accepted")

    def test_inspection_lists_sessions_and_reports_missing_column(self):
        self.assertEqual(DIAG.inspect_csv(self.report), ["s1"])
        bad = self.root / "bad.csv"
        bad.write_text("Status,ID\nError,1\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "Session ID column"):
            DIAG.inspect_csv(bad)

    def test_default_rules_use_import_throttle_and_explicit_access_side(self):
        template = SCRIPT.parents[2] / "Config" / "sharegate-diagnostics.rules.json.template"
        catalog = json.loads(template.read_text(encoding="utf-8"))
        rules = [(rule, re.compile(rule["Regex"])) for rule in catalog["Rules"]]
        cases = (
            ("Error", "Import failed", "Failed", "", "SG-IMPORT"),
            ("Warning", "Service delay", "", "429 requests", "SG-THROTTLE"),
            ("Error", "Source access denied", "", "", "SG-ACCESS-UNKNOWN"),
            ("Error", "Destination access denied", "", "", "SG-ACCESS-UNKNOWN"),
            ("Error", "Access denied", "", "", "SG-ACCESS-UNKNOWN"),
        )
        for status, message, import_status, throttle, expected in cases:
            row = {"Status": status, "Message": message, "Details": "", "ImportStatus": import_status,
                   "ThrottlingStatistics": throttle}
            DIAG.classify(row, rules, self.project)
            self.assertEqual(row["RuleId"], expected)

    def test_access_trace_host_and_five_minute_breakdown(self):
        source = "https://source.example/sites/a"
        target = "https://target.example/sites/a"
        report = self.root / "access.csv"
        fields = ["Session ID", "ID", "Date", "Status", "Type", "Title", "Source site address",
                  "Source list title", "Source ID", "Destination site address", "Errors"]
        cases = [
            ("1", "2026-09-30T22:01:26Z", "401 Unauthorized. URL 'https://source.example/sites/a/_vti_bin/client.svc/ProcessQuery' was not authorized"),
            ("2", "2026-09-30T22:03:01Z", "401 Unauthorized. WebUri : 'https://source.example/sites/a/_vti_bin/client.svc/ProcessQuery'"),
            ("3", "2026-09-30T22:06:01Z", "401 Unauthorized. WebUri : 'https://target.example/sites/a/_vti_bin/client.svc/ProcessQuery'"),
            ("4", "2026-09-30T22:09:01Z", "You are not authorized; no failing URL was recorded"),
        ]
        with report.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(fields)
            for row_id, timestamp, error in cases:
                writer.writerow(["s1", row_id, timestamp, "Error", "File", row_id, source,
                                 "Documents", row_id, target, error])
        DIAG.analyze([report], self.output, self.project)
        summary = json.loads((self.output / "Summary.json.txt").read_text(encoding="utf-8"))
        self.assertEqual(summary["AccessLines"], 4)
        self.assertEqual(summary["AccessItems"], 4)
        self.assertEqual([(bucket["Lines"], bucket["Source"], bucket["Destination"], bucket["Undetermined"])
                          for bucket in summary["AccessTimeBuckets"]], [(2, 2, 0, 0), (2, 0, 1, 1)])
        self.assertEqual(len(summary["AccessBySiteList"]), 1)
        with (self.output / "ClassifiedRows.csv").open(encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.DictReader(stream))
        self.assertEqual([row["AccessSide"] for row in rows], ["Source", "Source", "Destination", "Undetermined"])
        self.assertIn("source.example", rows[0]["AccessEvidence"])
        self.assertEqual(rows[3]["AccessEvidence"], "")
        html = (self.output / "MigrationDiagnostics-Report.html").read_text(encoding="utf-8")
        self.assertIn("Access failures by window, site and list", html)
        self.assertIn("22:00 UTC", html)
        self.assertTrue((self.output / "AccessFailures-5min-BySiteList.csv").exists())
        ambiguous = {"SourceUrl": "https://same.example/source", "DestinationUrl": "https://same.example/target",
                     "Message": "401 Unauthorized. WebUri : 'https://same.example/source/_vti_bin/client.svc'", "Details": ""}
        self.assertEqual(DIAG.failing_endpoint(ambiguous)[0], "Undetermined")


if __name__ == "__main__":
    unittest.main()
