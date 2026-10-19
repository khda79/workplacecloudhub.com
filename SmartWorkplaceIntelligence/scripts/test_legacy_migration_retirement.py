"""Guard retired staging tables against model and report dependencies.

Offline only. SMARTWORKPLACE_PBIP may select another PBIP project folder.
"""
import json
import os
from pathlib import Path
import re
import tempfile
import unittest


RETIRED = (
    "Mailbox Migration Evidence",
    "Mailbox Migration History",
    "Windows AD Unlinked Evidence",
)
REPLACEMENTS = (
    "Mailbox Migration Inventory",
    "Mailbox Migration Weekly History",
    "Windows Migration AD Observations",
)
PBIP = Path(os.environ.get("SMARTWORKPLACE_PBIP") or (
    Path(__file__).resolve().parents[1] / "pbip"))
MODEL = "SmartWorkplaceIntelligence.SemanticModel"
REPORT = "SmartWorkplaceIntelligence.Report"


def retired_dependencies(definition, report):
    """Check exact table tokens; retained '... Rebuilt' partition names are valid."""
    issues = []
    token = re.compile(r"(['\"])(" + "|".join(map(re.escape, RETIRED)) + r")\1")
    for name in RETIRED:
        if (definition / "tables" / (name + ".tmdl")).exists():
            issues.append("Retired table file: " + name)
    for file in sorted(definition.rglob("*.tmdl")):
        for number, line in enumerate(file.read_text(encoding="utf-8-sig").splitlines(), 1):
            if token.search(line) and not line.lstrip().startswith(("///", "//")):
                issues.append(f"Model reference: {file.relative_to(definition)}:{number}")

    def walk(value, file):
        if isinstance(value, dict):
            for item in value.values():
                walk(item, file)
        elif isinstance(value, list):
            for item in value:
                walk(item, file)
        elif isinstance(value, str) and (value in RETIRED or token.search(value)):
            issues.append("Report reference: " + str(file.relative_to(report)))

    for file in sorted(report.rglob("*.json")):
        walk(json.loads(file.read_text(encoding="utf-8-sig")), file)
    return issues


class LegacyMigrationRetirementTests(unittest.TestCase):
    def test_model_and_report_dependencies(self):
        issues = retired_dependencies(PBIP / MODEL / "definition", PBIP / REPORT / "definition")
        self.assertEqual(issues, [], "\n".join(issues))

    def test_replacements_are_registered_with_import_partitions(self):
        definition = PBIP / MODEL / "definition"
        model = (definition / "model.tmdl").read_text(encoding="utf-8-sig")
        for name in REPLACEMENTS:
            with self.subTest(table=name):
                self.assertIn("ref table '" + name + "'", model)
                table = (definition / "tables" / (name + ".tmdl")).read_text(encoding="utf-8-sig")
                self.assertIn("table '" + name + "'", table)
                self.assertRegex(table, r"(?m)^\tpartition .+ = m$")
                self.assertIn("mode: import", table)
                self.assertNotIn("excludeFromModelRefresh", table)

    def test_checker_rejects_model_dependencies(self):
        examples = (
            "ref table 'Mailbox Migration Evidence'",
            "\tmeasure X = COUNTROWS ( 'Mailbox Migration History' )",
            '    Source = #"Windows AD Unlinked Evidence",',
            "\tfromColumn: 'Mailbox Migration Evidence'.Country",
            "\ttablePermission 'Mailbox Migration History'",
            'annotation PBI_QueryOrder = ["Windows AD Unlinked Evidence"]',
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for example in examples:
                with self.subTest(reference=example):
                    (root / "model.tmdl").write_text(example, encoding="utf-8")
                    self.assertTrue(retired_dependencies(root, root / "report"))

    def test_checker_rejects_nested_report_dependencies(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in RETIRED:
                with self.subTest(table=name):
                    value = {"filter": [{"Expression": {"SourceRef": {"Entity": name}}}]}
                    (root / "visual.json").write_text(json.dumps(value), encoding="utf-8")
                    self.assertTrue(retired_dependencies(root / "model", root))

    def test_checker_rejects_retired_table_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "tables").mkdir()
            (root / "tables" / (RETIRED[0] + ".tmdl")).touch()
            self.assertTrue(retired_dependencies(root, root / "report"))

    def test_rebuilt_partition_names_are_not_table_dependencies(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "model.tmdl").write_text("\n".join(
                "\tpartition '" + name + " Rebuilt' = m" for name in RETIRED), encoding="utf-8")
            self.assertEqual(retired_dependencies(root, root / "report"), [])


if __name__ == "__main__":
    unittest.main()
