"""Offline tests with fictional principals; no tenant access."""
import csv
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Scripts" / "Compare"))
import cross_tenant_identity as identity
import scan_evidence


class CrossTenantIdentityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "mapping.csv"
        self.row = {
            "SourceWebUrl": "https://source.example.test/sites/a",
            "SourcePrincipalType": "User",
            "SourcePrincipalLoginName": "user@source.example.test",
            "TargetWebUrl": "https://target.example.test/sites/a",
            "TargetPrincipalType": "User",
            "TargetPrincipalLoginName": "person@target.example.test",
            "Status": "Confirmed",
        }

    def write(self, rows):
        with self.path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=identity.REQUIRED_COLUMNS)
            writer.writeheader()
            writer.writerows(rows)

    def test_confirmed_mapping_is_scoped_by_source_web(self):
        other = dict(self.row, SourceWebUrl="https://source.example.test/sites/b", TargetPrincipalLoginName="other@target.example.test")
        self.write([self.row, other])
        mappings = identity.load_confirmed_mappings(self.path)
        row = {"WebUrl": "https://SOURCE.example.test/sites/a/", "PrincipalType": "user", "PrincipalLoginName": "USER@source.example.test"}
        self.assertEqual(identity.resolve_source_row(row, mappings)["PrincipalLoginName"], "person@target.example.test")
        row["WebUrl"] = "https://source.example.test/sites/unknown"
        self.assertIsNone(identity.resolve_source_row(row, mappings))
        self.assertEqual(row["PrincipalLoginName"], "USER@source.example.test")

    def test_duplicate_or_unconfirmed_mapping_is_rejected(self):
        for rows in ([self.row, dict(self.row)], [dict(self.row, Status="Proposed")]):
            self.write(rows)
            with self.assertRaises(ValueError):
                identity.load_confirmed_mappings(self.path)

    def test_distinct_source_principals_cannot_silently_merge(self):
        other = dict(self.row, SourcePrincipalLoginName="different@source.example.test")
        self.write([self.row, other])
        with self.assertRaises(ValueError):
            identity.load_confirmed_mappings(self.path)

    def test_missing_identity_or_invalid_web_is_rejected(self):
        for change in ({"TargetPrincipalLoginName": ""}, {"SourcePrincipalLoginName": ""}, {"SourceWebUrl": "http://source.example.test/a"}, {"SourceWebUrl": "https://source.example.test:bad/a"}, {"SourceWebUrl": "https://source.example.test/a b"}, {"TargetWebUrl": ""}):
            self.write([dict(self.row, **change)])
            with self.assertRaises(ValueError):
                identity.load_confirmed_mappings(self.path)

    def test_empty_mapping_is_rejected(self):
        self.write([])
        with self.assertRaises(ValueError):
            identity.load_confirmed_mappings(self.path)

    def test_resolution_keeps_unresolved_and_original_evidence(self):
        self.write([self.row])
        mappings = identity.load_confirmed_mappings(self.path)
        source = {"WebUrl": self.row["SourceWebUrl"], "PrincipalType": "User", "PrincipalLoginName": self.row["SourcePrincipalLoginName"], "PermissionLevels": "Read"}
        unknown = dict(source, PrincipalLoginName="unknown@source.example.test")
        invalid = dict(source, PrincipalLoginName="")
        resolved, unresolved = identity.resolve_inventory([source, unknown, invalid], mappings)
        self.assertEqual(len(resolved), 1)
        self.assertEqual([reason for _, reason in unresolved], ["No confirmed mapping", "Invalid source principal or web URL"])
        self.assertEqual(resolved[0][0], source)
        self.assertEqual(resolved[0][1]["PrincipalLoginName"], self.row["TargetPrincipalLoginName"])
        target = {"WebUrl": self.row["TargetWebUrl"], "PrincipalType": "User", "PrincipalLoginName": self.row["TargetPrincipalLoginName"]}
        self.assertTrue(identity.matches_target_row(resolved[0][1], target))
        self.assertFalse(identity.matches_target_row(resolved[0][1], dict(target, WebUrl="https://target.example.test/sites/other")))
        self.assertFalse(identity.matches_target_row(resolved[0][1], dict(target, PrincipalType="SecurityGroup")))

    def test_comparison_tokens_require_confirmed_matching_web_scope(self):
        self.write([self.row])
        mappings = identity.load_confirmed_mappings(self.path)
        source = {"WebUrl": self.row["SourceWebUrl"], "PrincipalType": "User", "PrincipalLoginName": self.row["SourcePrincipalLoginName"]}
        target = {"WebUrl": self.row["TargetWebUrl"], "PrincipalType": "User", "PrincipalLoginName": self.row["TargetPrincipalLoginName"]}
        map_source = lambda row: '/sites/a'
        normalize_target = lambda url: '/sites/a'
        prepared_source, prepared_target = identity.prepare_comparison_rows([source], [target], mappings, map_source, normalize_target)
        self.assertEqual(prepared_source[0]['CrossTenantComparisonPrincipal'], prepared_target[0]['CrossTenantComparisonPrincipal'])
        self.assertNotIn('CrossTenantComparisonPrincipal', source)
        with self.assertRaisesRegex(ValueError, 'incomplete'):
            identity.prepare_comparison_rows([dict(source, PrincipalLoginName='unknown@example.test')], [target], mappings, map_source, normalize_target)
        with self.assertRaisesRegex(ValueError, 'conflicts'):
            identity.prepare_comparison_rows([source], [target], mappings, lambda row: '/wrong', normalize_target)
        with self.assertRaisesRegex(ValueError, 'not evidenced'):
            identity.prepare_comparison_rows([source], [dict(target, WebUrl='https://other.example.test/sites/a')], mappings, map_source, normalize_target)

    def test_missing_or_duplicate_columns_are_rejected(self):
        for header in ("SourceWebUrl,Status\n", ",".join(identity.REQUIRED_COLUMNS) + ",Status\n"):
            self.path.write_text(header, encoding="utf-8")
            with self.assertRaises(ValueError):
                identity.load_confirmed_mappings(self.path)

    def test_web_routes_enforce_host_and_most_specific_scope(self):
        routes_path = Path(self.temp.name) / 'routes.txt'
        routes_path.write_text('https://source.example.test/sites/a https://target.example.test/sites/a\n'
                               'https://source.example.test/sites/a/sub https://target.example.test/sites/other\n', encoding='utf-8')
        routes = identity.load_web_mappings(routes_path)
        self.assertEqual(identity.expected_target_web('https://source.example.test/sites/a/sub/child', routes),
                         'https://target.example.test/sites/other/child')
        with self.assertRaises(ValueError):
            identity.expected_target_web('https://impostor.example.test/sites/a', routes)
        self.write([dict(self.row, TargetWebUrl='https://impostor.example.test/sites/a')])
        source = {'WebUrl': self.row['SourceWebUrl'], 'PrincipalType': 'User', 'PrincipalLoginName': self.row['SourcePrincipalLoginName']}
        target = {'WebUrl': 'https://impostor.example.test/sites/a', 'PrincipalType': 'User', 'PrincipalLoginName': self.row['TargetPrincipalLoginName']}
        with self.assertRaisesRegex(ValueError, 'host-aware'):
            identity.prepare_comparison_rows([source], [target], identity.load_confirmed_mappings(self.path),
                                             lambda row: '/sites/a', lambda url: '/sites/a',
                                             lambda url: identity.expected_target_web(url, routes))

    def test_preflight_reports_coverage_without_claiming_permission_comparison(self):
        self.write([self.row])
        source_path = Path(self.temp.name) / "source.csv"
        source_path.write_text(
            "WebUrl;PrincipalType;PrincipalLoginName\n"
            "https://source.example.test/sites/a;User;user@source.example.test\n"
            "https://source.example.test/sites/a;User;unknown@source.example.test\n",
            encoding="utf-8",
        )
        unresolved_path = Path(self.temp.name) / "unresolved.csv"
        report = identity.preflight_inventory(source_path, self.path, unresolved_path)
        self.assertEqual((report["SourceRows"], report["ResolvedRows"], report["UnresolvedRows"]), (2, 1, 1))
        self.assertTrue(report["PreflightOnly"])
        self.assertFalse(report["PermissionsCompared"])
        self.assertFalse(report["GroupMembershipCompared"])
        self.assertEqual(report["SourceScanEvidence"], "Unknown")
        self.assertEqual(len(report["SourceSha256"]), 64)
        with unresolved_path.open(encoding="utf-8-sig", newline="") as handle:
            unresolved = list(csv.DictReader(handle))
        self.assertEqual(len(unresolved), 1)
        self.assertEqual(unresolved[0]["PrincipalLoginName"], "unknown@source.example.test")
        original_output = unresolved_path.read_bytes()
        with self.assertRaises(FileExistsError):
            identity.preflight_inventory(source_path, self.path, unresolved_path)
        self.assertEqual(unresolved_path.read_bytes(), original_output)
        process = subprocess.run(
            [sys.executable, identity.__file__, "--source-csv", str(source_path), "--mapping-csv", str(self.path)],
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(process.returncode, 2)
        self.assertEqual(json.loads(process.stdout)["UnresolvedRows"], 1)
        self.assertNotIn("unknown@source", process.stdout)

    def test_preflight_rejects_changed_inventory_after_manifest(self):
        self.write([self.row])
        source_path = Path(self.temp.name) / "source.csv"
        source_path.write_text("WebUrl,PrincipalType,PrincipalLoginName\nhttps://source.example.test/sites/a,User,user@source.example.test\n", encoding="utf-8")
        scan_evidence.write_manifest(source_path, "Source", "Permission", "https://source.example.test/sites/a")
        with source_path.open("a", encoding="utf-8") as handle:
            handle.write("https://source.example.test/sites/a,User,another@source.example.test\n")
        with self.assertRaisesRegex(ValueError, "manifest does not match"):
            identity.preflight_inventory(source_path, self.path)

    def test_target_observation_is_scoped_and_does_not_claim_absence(self):
        self.write([self.row])
        source_path = Path(self.temp.name) / "source.csv"
        target_path = Path(self.temp.name) / "target.csv"
        source_path.write_text("WebUrl,PrincipalType,PrincipalLoginName\nhttps://source.example.test/sites/a,User,user@source.example.test\n", encoding="utf-8")
        target_path.write_text("WebUrl,PrincipalType,PrincipalLoginName\nhttps://target.example.test/sites/other,User,person@target.example.test\nhttps://target.example.test/sites/a,User,\n", encoding="utf-8")
        report = identity.preflight_inventory(source_path, self.path, target_path=target_path)
        self.assertTrue(report["TargetInventoryChecked"])
        self.assertEqual(report["TargetRowsWithInvalidIdentity"], 1)
        self.assertEqual(report["MappedSourceRowsWithTargetPrincipalNotObserved"], 1)
        self.assertFalse(report["PermissionsCompared"])
        target_path.write_text("WebUrl,PrincipalType,PrincipalLoginName\nhttps://target.example.test/sites/a,User,person@target.example.test\n", encoding="utf-8")
        report = identity.preflight_inventory(source_path, self.path, target_path=target_path)
        self.assertEqual(report["MappedSourceRowsWithTargetPrincipalObserved"], 1)


if __name__ == "__main__":
    unittest.main()
