"""Execute the actual PBIP license-label steps against synthetic rows."""
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

MODEL = Path(__file__).resolve().parents[1] / "pbip" / "SmartWorkplaceIntelligence.SemanticModel" / "definition" / "tables" / "Mailbox Shared Conversion Evidence.tmdl"
PQTEST = os.environ.get("PQTEST_PATH") or shutil.which("PQTest.exe")


@unittest.skipUnless(PQTEST, "Set PQTEST_PATH for actual M evaluation")
class SharedConversionLabels(unittest.TestCase):
    def test_qualified_zero_and_unknown_are_distinct(self):
        source = MODEL.read_text(encoding="utf-8-sig")
        match = re.search(r"(LicenseLabel = Table.AddColumn\(CountryRenamed,.*?\n\s*OwnerLicenses = Table.ReorderColumns\([^\n]+)", source, re.S)
        self.assertIsNotNone(match)
        query = '''let
CountryRenamed = #table(type table [Owner Licenses = nullable text, License Source Status = nullable text, Owner License Count = nullable number],
    {{"", "Observed", 0}, {"", "Not collected", null}, {"SPE_F1", "Observed", 1}, {"", "Observed", null}, {null, null, null}, {"SPE_F1", "Stale", null}}),
''' + match.group(1) + ''',
Rows = Table.ToRecords(OwnerLicenses),
Result = [OK = Rows{0}[Owner Licenses] = "No license" and Rows{1}[Owner Licenses] = "Unknown" and
    Rows{2}[Owner Licenses] = "SPE_F1" and Rows{3}[Owner Licenses] = "Unknown" and
    Rows{4}[Owner Licenses] = "Unknown" and Rows{5}[Owner Licenses] = "Unknown" and
    Table.ColumnNames(OwnerLicenses) = Table.ColumnNames(CountryRenamed)]
in Result'''
        with tempfile.TemporaryDirectory() as scratch:
            path = Path(scratch) / "shared-conversion.query.pq"
            path.write_text(query, encoding="utf-8")
            run = subprocess.run([PQTEST, "run-test", "--queryFile", str(path)], capture_output=True,
                                 text=True, encoding="utf-8", errors="replace", timeout=90)
        self.assertEqual(run.returncode, 0, run.stdout[-2500:] + run.stderr[-1000:])
        compact = run.stdout.replace(" ", "")
        self.assertIn('"Status":"Passed"', compact, run.stdout[-2500:])
        self.assertIn('"Value":true', compact, run.stdout[-2500:])


if __name__ == "__main__":
    unittest.main()
