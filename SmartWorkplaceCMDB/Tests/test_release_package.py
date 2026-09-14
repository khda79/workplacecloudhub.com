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


if __name__ == "__main__":
    unittest.main()
