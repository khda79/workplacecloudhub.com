"""Offline compatibility checks for folder cleanup artifacts."""

import csv
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EXPORT = ROOT / "Scripts" / "Export" / "export_folders_to_delete_excel.py"
GENERATE = ROOT / "Scripts" / "Generate" / "generate_sp_target_delete_folders_script.py"


class FolderDeleteArtifactTests(unittest.TestCase):
    def test_generator_accepts_prefixed_and_legacy_folder_csv(self):
        for name, has_folder in (
            ("Example-ExtraFoldersInTarget.csv", False),
            ("Example-ExtraFoldersInTarget.csv", True),
            ("ExtraFoldersInTarget.csv", True),
        ):
            with self.subTest(name=name, has_folder=has_folder), tempfile.TemporaryDirectory() as directory:
                comparison = Path(directory)
                with (comparison / name).open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(
                        handle,
                        fieldnames=["TargetWebUrl", "TargetFolderServerRelativeUrl", "Depth"],
                        delimiter=";",
                    )
                    writer.writeheader()
                    if has_folder:
                        writer.writerow({
                            "TargetWebUrl": "https://target.example/sites/a",
                            "TargetFolderServerRelativeUrl": "/sites/a/Documents/Extra",
                            "Depth": "1",
                        })

                output = comparison / "SmartM365-Example-SharePointTarget-ExtraFoldersRemove.ps1"
                result = subprocess.run(
                    [sys.executable, str(GENERATE), "--comparison-directory", str(comparison),
                     "--output-ps1", str(output)],
                    capture_output=True, text=True,
                )
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
                self.assertEqual(output.exists(), has_folder)

                if has_folder:
                    workbook = comparison / "Example-Folders-To-Delete.xlsx"
                    export = subprocess.run(
                        [sys.executable, str(EXPORT), "--comparison-directory", str(comparison),
                         "--output-xlsx", str(workbook)],
                        capture_output=True, text=True,
                    )
                    self.assertEqual(export.returncode, 0, export.stdout + export.stderr)
                    self.assertTrue(workbook.is_file())


if __name__ == "__main__":
    unittest.main()
