"""Offline checks for generated cleanup launchers."""

import csv
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
GENERATE = ROOT / "Scripts" / "Generate"


class GeneratedLauncherTests(unittest.TestCase):
    def test_each_cleanup_generator_creates_and_removes_companion_launcher(self):
        cases = (
            ("files", "ExtraInTarget.csv", ["WebUrl", "ServerRelativeUrl", "FileName"],
             ["https://target.example/sites/a", "/sites/a/Docs/extra.txt", "extra.txt"]),
            ("libraries", "LibrarySummary.csv", ["TargetWebUrl", "WebPath", "LibraryPath", "ExtraInTarget"],
             ["https://target.example/sites/a", "sites/a", "sites/a/Docs", "1"]),
            ("folders", "ExtraFoldersInTarget.csv", ["TargetWebUrl", "TargetFolderServerRelativeUrl", "Depth"],
             ["https://target.example/sites/a", "/sites/a/Docs/Extra", "1"]),
        )
        for kind, csv_name, fields, values in cases:
            with self.subTest(kind=kind), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                source = root / csv_name
                script = root / f"SmartM365-Example-SharePointTarget-Extra{kind.title()}Remove.ps1"
                launcher = script.with_suffix(".cmd")

                def write_csv(include_row):
                    with source.open("w", encoding="utf-8", newline="") as handle:
                        writer = csv.DictWriter(handle, fieldnames=fields, delimiter=";")
                        writer.writeheader()
                        if include_row:
                            writer.writerow(dict(zip(fields, values)))

                def generate():
                    result = subprocess.run(
                        [sys.executable, str(GENERATE / f"generate_sp_target_delete_{kind}_script.py"),
                         "--comparison-directory", str(root), "--output-ps1", str(script)],
                        capture_output=True, text=True,
                    )
                    self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

                write_csv(True)
                generate()
                self.assertTrue(script.is_file())
                self.assertTrue(launcher.is_file())
                content = launcher.read_text(encoding="ascii")
                self.assertIn("%~dpn0.ps1", content)
                self.assertIn('-ExecutionPolicy Bypass -File "%TARGET_SCRIPT%" -Execute', content)
                self.assertNotIn("-Force", content)
                self.assertNotIn("%*", content)

                write_csv(False)
                generate()
                self.assertFalse(script.exists())
                self.assertFalse(launcher.exists())

    @unittest.skipUnless(os.name == "nt", "Windows CMD launcher")
    def test_launcher_requires_generated_script_confirmation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "SmartM365-SharePointTarget-ExtraFilesRemove-20261011-102759-example.ps1"
            marker = root / "executed.txt"
            script.write_text(
                "param([switch]$Execute)\n"
                "if (-not $Execute) { exit 2 }\n"
                "$answer = Read-Host 'Type DELETE to continue'\n"
                "if ($answer -ne 'DELETE') { exit 3 }\n"
                "Set-Content -LiteralPath (Join-Path $PSScriptRoot 'executed.txt') -Value 'done'\n",
                encoding="utf-8",
            )
            backfill = subprocess.run(
                [sys.executable, str(GENERATE / "generated_ps7_launcher.py"), str(script)],
                capture_output=True, text=True,
            )
            self.assertEqual(backfill.returncode, 0, backfill.stdout + backfill.stderr)
            launcher = script.with_suffix(".cmd")
            restrictive_env = os.environ.copy()
            restrictive_env["PSExecutionPolicyPreference"] = "AllSigned"
            denied = subprocess.run(["cmd.exe", "/d", "/c", str(launcher)], input="NO\n",
                                    capture_output=True, text=True, timeout=30, env=restrictive_env)
            self.assertEqual(denied.returncode, 3, denied.stdout + denied.stderr)
            self.assertFalse(marker.exists())
            accepted = subprocess.run(["cmd.exe", "/d", "/c", str(launcher)], input="DELETE\n",
                                      capture_output=True, text=True, timeout=30, env=restrictive_env)
            self.assertEqual(accepted.returncode, 0, accepted.stdout + accepted.stderr)
            self.assertEqual(marker.read_text(encoding="utf-8").strip(), "done")


if __name__ == "__main__":
    unittest.main()
