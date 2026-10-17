"""Offline checks for the source-only migration complexity score."""

import csv
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT / "Scripts" / "Compare"))
import migration_complexity as complexity

FILE_FIELDS = ["SiteCollectionUrl", "WebUrl", "LibraryUrl", "FileName", "ServerRelativeUrl",
               "SizeBytes", "VersionsCount", "CheckedOutBy"]
PERMISSION_FIELDS = ["ObjectScope", "ObjectId", "ObjectUrl", "HasUniqueRoleAssignments", "PrincipalType",
                     "PrincipalName", "PrincipalLoginName", "PrincipalMemberLoginNames", "IsLimitedAccessOnly"]


def write_csv(path, fields, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


class MigrationComplexityTests(unittest.TestCase):
    @staticmethod
    def temporary_directory():
        return tempfile.TemporaryDirectory(prefix=".migration-complexity-test-", dir=Path(__file__).parent)

    def config_dir(self, root):
        config = root / "Config"
        config.mkdir()
        shutil.copyfile(PROJECT / "Config" / "migration-complexity.json.template",
                        config / "migration-complexity.json.template")
        return config

    def test_boundaries_belong_to_the_lower_band(self):
        criterion = {"OnePointFrom": 10, "TwoPointsAbove": 100, "ThreePointsAbove": 1000}
        self.assertEqual([complexity.points_for(value, criterion) for value in (9, 10, 100, 101, 1000, 1001)],
                         [0, 1, 1, 2, 2, 3])

    def test_runtime_config_self_heals_without_overwriting_local_values(self):
        with self.temporary_directory() as directory:
            config_dir = self.config_dir(Path(directory))
            template = json.loads((config_dir / "migration-complexity.json.template").read_text(encoding="utf-8"))
            local = {"SchemaVersion": 1, "Criteria": [dict(template["Criteria"][0], Weight=40)],
                     "Levels": template["Levels"], "FloorRules": []}
            (config_dir / "migration-complexity.json.txt").write_text(json.dumps(local), encoding="utf-8")
            config, runtime = complexity.load_config(config_dir)
            self.assertEqual(config["Criteria"][0]["Weight"], 40)
            self.assertEqual(len(config["Criteria"]), len(template["Criteria"]))
            self.assertIn("Settings", config)
            self.assertEqual(len(config["FloorRules"]), len(template["FloorRules"]))
            self.assertEqual(json.loads(runtime.read_text(encoding="utf-8")), config)

    def test_floor_rule_raises_level_and_partial_inputs_keep_denominator_honest(self):
        with self.temporary_directory() as directory:
            config, _ = complexity.load_config(self.config_dir(Path(directory)))
            values = {criterion["Id"]: (0, "") for criterion in config["Criteria"]
                      if criterion["Id"] in complexity.FILE_CRITERIA}
            values["PathLimitFiles"] = (500, "")
            result = complexity.score_migration(config, values)
            self.assertEqual(result["Score"], 14)
            self.assertEqual(result["ScoreLevel"], "Low")
            self.assertEqual(result["Level"], "High")
            self.assertEqual(result["FloorApplied"], ["PathLimitFiles"])
            self.assertEqual(result["EvaluatedWeight"], 70)
            self.assertEqual(result["TopDrivers"], ["PathLimitFiles"])
            inactive = next(row for row in result["Criteria"] if row["Id"] == "InactiveAccounts")
            self.assertFalse(inactive["Evaluated"])

    def test_files_outside_the_path_mapping_are_skipped(self):
        mappings = [("/sites/a", "/sites/target-a")]
        self.assertEqual(complexity.projected_target_path("/sites/a/Docs/x.docx", mappings, "/sites/ignored"),
                         "sites/target-a/Docs/x.docx")
        self.assertEqual(complexity.projected_target_path("/sites/b/Docs/x.docx", mappings, "/sites/ignored"), "")
        self.assertEqual(complexity.projected_target_path("/Docs/x.docx", [], "/sites/t"), "sites/t/Docs/x.docx")

    def test_disabled_criterion_and_floor_rule_are_ignored(self):
        with self.temporary_directory() as directory:
            config, _ = complexity.load_config(self.config_dir(Path(directory)))
            for criterion in config["Criteria"]:
                if criterion["Id"] == "LibraryCount":
                    criterion["Enabled"] = False
            for rule in config["FloorRules"]:
                rule["Enabled"] = False
            values = {criterion["Id"]: (0, "") for criterion in config["Criteria"]
                      if criterion["Id"] in complexity.FILE_CRITERIA}
            values["LibraryCount"] = (500, "")
            values["PathLimitFiles"] = (500, "")
            result = complexity.score_migration(config, values)
            library = next(row for row in result["Criteria"] if row["Id"] == "LibraryCount")
            self.assertFalse(library["Evaluated"])
            self.assertEqual(library["ValueDisplay"], "Disabled")
            self.assertEqual(result["EvaluatedWeight"], 65)
            self.assertEqual(result["FloorApplied"], [])
            self.assertEqual(result["Level"], "Low")

    def test_end_to_end_metrics_from_source_scans(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            config_dir = self.config_dir(root)
            long_name = "x" * 260 + ".docx"
            files = root / "SP2019-FileInventory-Fixture-20261001-120000.csv"
            write_csv(files, FILE_FIELDS, [
                dict(WebUrl="https://source.test", LibraryUrl="/Docs", FileName="a.docx",
                     ServerRelativeUrl="/Docs/a.docx", SizeBytes=3 * 1024 ** 3, VersionsCount=4),
                dict(WebUrl="https://source.test/sub", LibraryUrl="/sub/Docs", FileName="b.docx",
                     ServerRelativeUrl="/sub/Docs/b.docx", SizeBytes=10, VersionsCount=0, CheckedOutBy="someone"),
                dict(WebUrl="https://source.test/sub", LibraryUrl="/sub/Docs", FileName="Thumbs.db",
                     ServerRelativeUrl="/sub/Docs/Thumbs.db", SizeBytes=10, VersionsCount=0),
                dict(WebUrl="https://source.test/sub", LibraryUrl="/sub/Docs", FileName=long_name,
                     ServerRelativeUrl=f"/sub/Docs/{long_name}", SizeBytes=10, VersionsCount=2),
            ])
            permissions = root / "SP2019-PermissionInventory-Fixture-20261001-120000.csv"
            write_csv(permissions, PERMISSION_FIELDS, [
                dict(ObjectScope="Item", ObjectId="1", HasUniqueRoleAssignments="True", PrincipalType="User",
                     PrincipalLoginName="i:0e.t|adfs|active@contoso.test"),
                dict(ObjectScope="Item", ObjectId="1", HasUniqueRoleAssignments="True", PrincipalType="User",
                     PrincipalLoginName="i:0e.t|adfs|disabled@contoso.test"),
                dict(ObjectScope="List", ObjectId="2", HasUniqueRoleAssignments="True", PrincipalType="User",
                     PrincipalLoginName="i:0e.t|adfs|limited@contoso.test", IsLimitedAccessOnly="True"),
                dict(ObjectScope="Web", ObjectId="3", HasUniqueRoleAssignments="True", PrincipalType="SharePointGroup",
                     PrincipalLoginName="Owners",
                     PrincipalMemberLoginNames="i:0e.t|adfs|ghost@contoso.test || i:0#.w|sharepoint\\system"),
            ])
            entra = root / "entra.csv"
            write_csv(entra, ["UserPrincipalName", "AccountEnabled"], [
                dict(UserPrincipalName="active@contoso.test", AccountEnabled="True"),
                dict(UserPrincipalName="disabled@contoso.test", AccountEnabled="False"),
                dict(UserPrincipalName="limited@contoso.test", AccountEnabled="True"),
            ])
            output = root / "complexity.json.txt"
            complexity.main(["--migration-name", "Fixture", "--file-inventory", str(files),
                             "--permission-inventory", str(permissions), "--entra-users-csv", str(entra),
                             "--target-site-url", "https://target.test/sites/fixture",
                             "--config-dir", str(config_dir), "--output", str(output)])
            result = json.loads(output.read_text(encoding="utf-8"))
            values = {row["Id"]: row["Value"] for row in result["Criteria"]}
            self.assertEqual(result["Status"], "Complete")
            self.assertEqual(values["FileCount"], 4)
            self.assertEqual(values["SubsiteCount"], 2)
            self.assertEqual(values["LibraryCount"], 2)
            self.assertEqual(values["AverageVersions"], 1.5)
            self.assertEqual(values["LargeFiles"], 1)
            self.assertEqual(values["RiskyFiles"], 2)
            self.assertEqual(values["PathLimitFiles"], 1)
            self.assertEqual(values["BrokenInheritance"], 1)
            self.assertEqual(values["DirectUsers"], 2)
            self.assertEqual(values["InactiveAccounts"], 2)
            self.assertEqual(result["Inputs"]["FileInventory"]["Name"], files.name)
            self.assertEqual(len(result["Inputs"]["ConfigSha256"]), 64)


if __name__ == "__main__":
    unittest.main()
