"""Offline public-package closure and collector-retirement tests."""
import json
from pathlib import Path
import re
import unittest

PRODUCT = Path(__file__).resolve().parents[1]


def load_allowlist():
    return json.loads((PRODUCT / 'Release/Files.json').read_text(encoding='utf-8-sig'))


class ReleasePackageTests(unittest.TestCase):
    def test_allowlist_entries_are_unique_existing_files(self):
        entries = load_allowlist()
        self.assertEqual(len(entries), len(set(entries)))
        self.assertEqual([p for p in entries if not (PRODUCT / p).is_file()], [])

    def test_allowlist_excludes_private_runtime_artifacts(self):
        forbidden = re.compile(
            r'(?:^|/)(?:Data|DATA-[^/]*|LOG-ALL|CMDB-REPORTS|\.local-review|\.pbi)(?:/|$)'
            r'|\.(?:pbip|pbix|bim|csv|log|transcript\.txt)$'
            r'|\.local\.json$|(?:^|/)AUDIT-[^/]*\.md$', re.IGNORECASE)
        self.assertEqual([p for p in load_allowlist() if forbidden.search(p)], [])

    def test_retired_collection_chain_absent(self):
        entries = load_allowlist()
        for directory in ('Collectors', 'Modules', 'Orchestration', 'Reports'):
            # Empty ignored folders are harmless; no executable runtime remains.
            self.assertEqual(list((PRODUCT / directory).rglob('*.ps1')), [])
            self.assertEqual(list((PRODUCT / directory).rglob('*.psm1')), [])
            self.assertFalse(any(p.startswith(directory + '/') for p in entries))
        self.assertFalse((PRODUCT / 'PowerBI/prepare_current_report_data.py').exists())

    def test_refresh_dependencies_are_allowlisted(self):
        expected = {'Launchers/Start-SmartWorkplaceCMDB-Refresh.ps1',
                    'Config/refresh.local.json.template', 'PowerBI/refresh_prepared_report.py',
                    'PowerBI/prepared_read_session.py', 'PowerBI/load_prepared_snapshot.py',
                    'PowerBI/Queries/prepared-read-session.pq', 'PowerBI/Queries/prepared-types-candidate.pq'}
        self.assertEqual(sorted(expected - set(load_allowlist())), [])

    def test_monorepo_source_contract_dependencies_exist(self):
        repo = PRODUCT.parent
        for path in ('SmartM365/SmartInventory/PreparedEvidence/cmdb_freshness.py',
                     'SmartM365/SmartInventory/PreparedEvidence/cmdb-prepared-contract.json.txt',
                     'SmartM365/Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt'):
            self.assertTrue((repo / path).is_file(), path)

    def test_native_date_regression_assets_are_allowlisted(self):
        self.assertIn('Tests/prepared_dates.query.pq', load_allowlist())
        self.assertIn('Tests/test_refresh_prepared_report.py', load_allowlist())

    def test_release_does_not_claim_global_live_qualification(self):
        release = json.loads((PRODUCT / 'RELEASE.json').read_text(encoding='utf-8-sig'))
        self.assertFalse(release['liveQualified'])
        self.assertRegex(release['version'], r'^\d+\.\d+\.\d+$')


if __name__ == '__main__':
    unittest.main()
