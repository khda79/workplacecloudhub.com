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
    @unittest.skipUnless(os.name == "nt", "PowerShell 7 on Windows")
    def test_direct_library_operation_has_recycle_default_without_file_guard(self):
        operation = ROOT / "Scripts" / "Operations" / "SmartM365-SharePointTarget-ExtraLibrariesRemove.ps1"
        content = operation.read_text(encoding="utf-8")
        self.assertIn("[switch]$Recycle = $true", content)
        self.assertIn("Remove-PnPList -Identity $list.Id -Recycle:$Recycle", content)
        self.assertNotIn("Test-LibraryContainsHomePage", content)
        self.assertNotIn("ProtectedHomePage", content)
        environment = os.environ.copy()
        environment["LIBRARY_OPERATION"] = str(operation)
        parsed = subprocess.run(
            ["pwsh.exe", "-NoProfile", "-Command", r"""
$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($env:LIBRARY_OPERATION, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { $errors | ForEach-Object { Write-Error $_.Message }; exit 1 }
exit 0
"""],
            capture_output=True, text=True, env=environment, timeout=30,
        )
        self.assertEqual(parsed.returncode, 0, parsed.stdout + parsed.stderr)

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
                self.assertIn('-ExecutionPolicy Bypass -File "%TARGET_SCRIPT%" -Execute -Recycle', content)
                self.assertNotIn("-Force", content)
                self.assertNotIn("%*", content)
                generated = script.read_text(encoding="utf-8")
                self.assertIn("SmartM365-CleanupSafety: RecycleDefault", generated)
                self.assertIn("[switch]$Recycle = $true", generated)
                if os.name == "nt":
                    environment = os.environ.copy()
                    environment["GENERATED_PS1"] = str(script)
                    parsed = subprocess.run(
                        ["pwsh.exe", "-NoProfile", "-Command", r"""
$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($env:GENERATED_PS1, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { $errors | ForEach-Object { Write-Error $_.Message }; exit 1 }
"""],
                        capture_output=True, text=True, env=environment, timeout=30,
                    )
                    self.assertEqual(parsed.returncode, 0, parsed.stdout + parsed.stderr)
                if kind == "files":
                    self.assertIn("SmartM365-CleanupSafety: HomeAspxProtected", generated)
                    self.assertIn("$actionStatus = 'ProtectedHomePage'", generated)
                    if os.name == "nt":
                        environment = os.environ.copy()
                        environment["GENERATED_PS1"] = str(script)
                        probe = subprocess.run(
                            ["pwsh.exe", "-NoProfile", "-Command", r"""
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:GENERATED_PS1, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { throw 'Generated script has PowerShell parse errors.' }
$function = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-ProtectedHomePage' }, $true)
if ($null -eq $function) { throw 'Home.aspx guard function is missing.' }
Invoke-Expression $function.Extent.Text
$removeFunction = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Remove-TargetFile' }, $true)
if ($null -eq $removeFunction) { throw 'File deletion function is missing.' }
Invoke-Expression $removeFunction.Extent.Text
$cases = @(
    @{ Path = '/sites/a/SitePages/Home.aspx'; Expected = $true },
    @{ Path = '/sites/a/SitePages/%48ome.aspx'; Expected = $true },
    @{ Path = '/sites/a/SitePages/home.aspx?web=1'; Expected = $true },
    @{ Path = '/sites/a/Docs/MyHome.aspx'; Expected = $false }
)
foreach ($case in $cases) {
    $item = [pscustomobject]@{ TargetServerRelativeUrl = $case.Path; TargetFileUrl = ''; FileName = '' }
    if ((Test-ProtectedHomePage -Item $item) -ne $case.Expected) { throw "Unexpected guard result for $($case.Path)" }
}
$item = [pscustomobject]@{ TargetServerRelativeUrl = '/sites/a/Docs/other.aspx'; TargetFileUrl = ''; FileName = 'HOME.ASPX' }
if (-not (Test-ProtectedHomePage -Item $item)) { throw 'FileName guard failed.' }
$script:RemoveCalled = $false
function Remove-PnPFile { $script:RemoveCalled = $true }
try {
    Remove-TargetFile -Connection ([pscustomobject]@{}) -ServerRelativeUrl '/sites/a/SitePages/Home.aspx' -Recycle
    throw 'Home.aspx deletion was allowed.'
}
catch { if ($_.Exception.Message -eq 'Home.aspx deletion was allowed.') { throw } }
if ($script:RemoveCalled) { throw 'Remove-PnPFile was called for Home.aspx.' }
exit 0
"""],
                            capture_output=True, text=True, env=environment, timeout=30,
                        )
                        self.assertEqual(probe.returncode, 0, probe.stdout + probe.stderr)
                if kind == "libraries":
                    self.assertIn("Remove-PnPList -Identity $list.Id -Recycle:$Recycle", generated)
                    self.assertNotIn("Test-LibraryContainsHomePage", generated)
                    self.assertNotIn("ProtectedHomePage", generated)

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
                "param([switch]$Execute, [switch]$Recycle)\n"
                "if (-not $Execute -or -not $Recycle) { exit 2 }\n"
                "$answer = Read-Host 'Type RECYCLE to continue'\n"
                "if ($answer -ne 'RECYCLE') { exit 3 }\n"
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
            unsafe = subprocess.run(["cmd.exe", "/d", "/c", str(launcher)], input="RECYCLE\n",
                                    capture_output=True, text=True, timeout=30, env=restrictive_env)
            self.assertEqual(unsafe.returncode, 1, unsafe.stdout + unsafe.stderr)
            self.assertIn("Cleanup safety checks are missing", unsafe.stdout)
            self.assertFalse(marker.exists())

            script.write_text(
                "# SmartM365-CleanupSafety: RecycleDefault\n"
                "# SmartM365-CleanupSafety: HomeAspxProtected\n"
                + script.read_text(encoding="utf-8"),
                encoding="utf-8",
            )
            denied = subprocess.run(["cmd.exe", "/d", "/c", str(launcher)], input="NO\n",
                                    capture_output=True, text=True, timeout=30, env=restrictive_env)
            self.assertEqual(denied.returncode, 3, denied.stdout + denied.stderr)
            self.assertFalse(marker.exists())
            accepted = subprocess.run(["cmd.exe", "/d", "/c", str(launcher)], input="RECYCLE\n",
                                      capture_output=True, text=True, timeout=30, env=restrictive_env)
            self.assertEqual(accepted.returncode, 0, accepted.stdout + accepted.stderr)
            self.assertEqual(marker.read_text(encoding="utf-8").strip(), "done")


if __name__ == "__main__":
    unittest.main()
