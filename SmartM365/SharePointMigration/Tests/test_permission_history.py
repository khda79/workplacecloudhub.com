"""Offline checks for permission scan history comparisons."""

import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Scripts" / "Compare"))
from compare_sp_permission_inventory_history import read_inventory, run


class PermissionHistoryTests(unittest.TestCase):
    @staticmethod
    def temporary_directory():
        return tempfile.TemporaryDirectory(prefix=".permission-history-test-", dir=Path(__file__).parent)

    def test_grants_are_classified_by_identity_and_effective_details(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            old = root / "SP2019-PermissionInventory-Fixture-20261001-120000.csv"
            new = root / "SP2019-PermissionInventory-Fixture-20261002-120000.csv"
            fields = ["ObjectScope", "ObjectUrl", "ObjectServerRelativeUrl", "ItemId",
                      "PrincipalType", "PrincipalName", "PrincipalLoginName",
                      "PermissionLevels", "HasUniqueRoleAssignments", "InheritedFrom",
                      "PrincipalMemberLoginNames"]

            def grant(item, principal, levels, members=""):
                return dict(ObjectScope="Item", ObjectUrl=f"https://source.test/sites/a/docs/{item}",
                            ObjectServerRelativeUrl=f"/sites/a/docs/{item}", ItemId=item,
                            PrincipalType="User", PrincipalName=principal,
                            PrincipalLoginName=principal, PermissionLevels=levels,
                            HasUniqueRoleAssignments="True", InheritedFrom="",
                            PrincipalMemberLoginNames=members)

            def write(path, rows):
                with path.open("w", encoding="utf-8-sig", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=fields, delimiter=";")
                    writer.writeheader()
                    writer.writerows(rows)

            write(old, [grant("1", "alice", "Read"), grant("2", "bob", "Read"),
                        grant("3", "carol", "Read"), grant("4", "dave", "Read")])
            write(new, [grant("1", "alice", "Read"), grant("2", "bob", "Edit"),
                        grant("4", "dave", "Read", "member-a"), grant("5", "erin", "Read")])
            self.assertEqual(len(read_inventory(old)), 4)
            output = root / "report"
            counts = run(old, new, output, "Fixture history")
            self.assertEqual({status: counts[status] for status in
                              ("Unchanged", "Changed", "Removed", "Added", "Ambiguous")},
                             {"Unchanged": 1, "Changed": 2, "Removed": 1, "Added": 1, "Ambiguous": 0})
            with (output / "PermissionChanges.csv").open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
            self.assertEqual(len(rows), 5)
            self.assertIn("PermissionLevels", next(row for row in rows if row["Principal"] == "bob")["ChangedFields"])
            self.assertIn("PrincipalMemberLoginNames", next(row for row in rows if row["Principal"] == "dave")["ChangedFields"])
            self.assertIn("href=\"PermissionChanges.csv\"", (output / "PermissionHistory-Report.html").read_text(encoding="utf-8"))

    def test_reversed_scans_are_rejected(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            old = root / "SPO-PermissionInventory-Fixture-20261002-120000.csv"
            new = root / "SPO-PermissionInventory-Fixture-20261001-120000.csv"
            old.write_text("ObjectScope,ObjectUrl,PrincipalType,PrincipalName,PermissionLevels\n", encoding="utf-8")
            new.write_text("ObjectScope,ObjectUrl,PrincipalType,PrincipalName,PermissionLevels\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "earlier"):
                run(old, new, root / "report", "Fixture history")


if __name__ == "__main__":
    unittest.main()
