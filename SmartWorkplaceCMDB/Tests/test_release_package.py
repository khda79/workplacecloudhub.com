"""Offline release-allowlist safety and dependency-closure tests."""

import json
from pathlib import Path
import re
import unittest


PRODUCT = Path(__file__).resolve().parents[1]
ALLOWLIST_PATH = PRODUCT / "Release" / "Files.json"


def load_allowlist():
    return json.loads(ALLOWLIST_PATH.read_text(encoding="utf-8-sig"))


class ReleasePackageTests(unittest.TestCase):
    def test_allowlist_entries_are_unique_existing_files(self):
        entries = load_allowlist()
        self.assertEqual(len(entries), len(set(entries)))
        missing = [entry for entry in entries if not (PRODUCT / entry).is_file()]
        self.assertEqual(missing, [])

    def test_allowlist_excludes_private_runtime_artifacts(self):
        forbidden = re.compile(
            r"(?:^|/)(?:Data|DATA-[^/]*|LOG-ALL)(?:/|$)"
            r"|\.(?:pbip|pbix|csv|log|transcript\.txt)$"
            r"|\.local\.json$|(?:^|/)AUDIT-[^/]*\.md$",
            flags=re.IGNORECASE,
        )
        violations = [entry for entry in load_allowlist() if forbidden.search(entry)]
        self.assertEqual(violations, [])

    def test_module_psscriptroot_dot_sources_are_allowlisted(self):
        entries = set(load_allowlist())
        pattern = re.compile(
            r"\.\s*\(\s*Join-Path\s+\$PSScriptRoot\s+['\"]([^'\"]+)['\"]\s*\)",
            flags=re.IGNORECASE,
        )
        missing = []
        for module in sorted((PRODUCT / "Modules").rglob("*.psm1")):
            source = module.read_text(encoding="utf-8-sig")
            for relative_dependency in pattern.findall(source):
                dependency = (module.parent / relative_dependency).relative_to(PRODUCT)
                package_path = dependency.as_posix()
                if package_path not in entries:
                    missing.append(f"{module.relative_to(PRODUCT).as_posix()} -> {package_path}")
        self.assertEqual(missing, [])

    def test_orchestrator_static_project_dependencies_are_allowlisted(self):
        entries = set(load_allowlist())
        orchestrator = PRODUCT / "Orchestration" / "SmartWorkplaceCMDB-Orchestrator.ps1"
        pattern = re.compile(
            r"Join-Path\s+\$projectRoot\s+['\"]([^'\"]+)['\"]",
            flags=re.IGNORECASE,
        )
        dependencies = {
            Path(match.replace("\\", "/")).as_posix()
            for match in pattern.findall(orchestrator.read_text(encoding="utf-8-sig"))
            if Path(match).suffix.lower() in {".json", ".ps1", ".psd1"}
        }
        missing = sorted(dependency for dependency in dependencies if dependency not in entries)
        self.assertEqual(missing, [])

    def test_orchestrator_fixtures_are_allowlisted(self):
        entries = set(load_allowlist())
        orchestrator = PRODUCT / "Orchestration" / "SmartWorkplaceCMDB-Orchestrator.ps1"
        fixture_pattern = re.compile(
            r"FixtureName\s*=\s*['\"]([^'\"]+)['\"]",
            flags=re.IGNORECASE,
        )
        fixture_names = set(
            fixture_pattern.findall(orchestrator.read_text(encoding="utf-8-sig"))
        )
        expected = {f"Tests/Fixtures/{name}" for name in fixture_names}
        missing = sorted(expected - entries)
        self.assertEqual(missing, [])

    def test_cloud_launchers_are_allowlisted(self):
        entries = set(load_allowlist())
        launchers = {
            launcher.relative_to(PRODUCT).as_posix()
            for launcher in (PRODUCT / "Launchers" / "Cloud").glob("*.cmd")
        }
        missing = sorted(launchers - entries)
        self.assertEqual(missing, [])


if __name__ == "__main__":
    unittest.main()
