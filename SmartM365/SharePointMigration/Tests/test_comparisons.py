"""Offline regression tests; all evidence is synthetic and temporary."""
import csv
import importlib.util
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'Scripts' / 'Compare' / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


FILES = load('compare_sp_source_target_file_inventories')
PERMISSIONS = load('compare_sp_source_target_permissions')


class MappingTests(unittest.TestCase):
    def test_invalid_mappings_fail_in_both_engines(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'mapping.txt'
            for module in (FILES, PERMISSIONS):
                for content in ('/source /target ignored', '/source', '# empty', '/source web /target web'):
                    with self.subTest(engine=module.__name__, content=content):
                        path.write_text(content, encoding='utf-8-sig')
                        with self.assertRaises(ValueError):
                            module.load_path_mappings(path)

    def test_valid_mapping_delimiters_and_longest_path(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'mapping.txt'
            for module in (FILES, PERMISSIONS):
                for delimiter in (' ', '\t', ';', ','):
                    with self.subTest(engine=module.__name__, delimiter=delimiter):
                        path.write_text('# comment\n/source /target\n/source/sub' + delimiter + '/special\n', encoding='utf-8-sig')
                        mappings = module.load_path_mappings(path)
                        self.assertEqual(module.apply_path_mappings('/source/sub/file.docx', mappings), '/special/file.docx')
                        self.assertEqual(module.apply_path_mappings('/source2/file.docx', mappings), '/source2/file.docx')
                        self.assertEqual(module.load_path_mappings(None), [])


class FileComparisonTests(unittest.TestCase):
    def test_versions_dates_duplicates_scope_and_dry_run_generation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fields = ['WebUrl', 'LibraryTitle', 'LibraryRootFolderUrl', 'FileName', 'ServerRelativeUrl', 'SizeBytes', 'Modified', 'Version']

            def row(name, version='1.0', modified='2026-09-08 10:00:00', web='/source'):
                return dict(zip(fields, ['https://workplacecloudhub.sharepoint.com' + web, 'Docs', web + '/Docs', name, web + '/Docs/' + name, '100', modified, version]))

            source = [row('same.docx'), row('old.docx', '2.0'), row('date.docx'), row('dup.docx'), row('dup.docx'), row('excluded.docx', web='/outside')]
            target = [row('same.docx', web='/target'), row('old.docx', web='/target'), row('date.docx', modified='2026-09-08 09:00:00', web='/target'), row('dup.docx', web='/target'), row('extra.docx', web='/target')]
            for name, rows in [('source', source), ('target', target)]:
                with (root / (name + '.csv')).open('w', newline='', encoding='utf-8') as handle:
                    writer = csv.DictWriter(handle, fieldnames=fields); writer.writeheader(); writer.writerows(rows)
                (root / (name + '.txt')).write_text('https://workplacecloudhub.sharepoint.com/' + name, encoding='utf-8')
            (root / 'mapping.txt').write_text('/source /target', encoding='utf-8')
            output = root / 'output'
            run = subprocess.run([sys.executable, str(ROOT / 'Scripts/Compare/compare_sp_source_target_file_inventories.py'),
                '--source-csv', str(root / 'source.csv'), '--target-csv', str(root / 'target.csv'), '--output-directory', str(output),
                '--path-mapping-file', str(root / 'mapping.txt'), '--source-web-urls-file', str(root / 'source.txt'),
                '--target-web-urls-file', str(root / 'target.txt'), '--modified-date-tolerance-minutes', '0',
                '--source-modified-time-zone', 'UTC', '--target-modified-time-zone', 'UTC'], capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)

            def read(name):
                with (output / name).open(encoding='utf-8-sig', newline='') as handle:
                    return list(csv.DictReader(handle, delimiter=';'))

            self.assertEqual(len(read('ChangedVersion.csv')), 1)
            self.assertEqual(len(read('TargetOlderThanSource.csv')), 1)
            self.assertEqual(len(read('DuplicateKeys.csv')), 1)
            self.assertTrue(any(r['FileName'] == 'old.docx' for r in read('MissingInTarget.csv')))
            self.assertFalse(any(r['FileName'] == 'excluded.docx' for r in read('MissingInTarget.csv')))
            script = ROOT / 'Scripts/Generate/generate_sp_target_delete_files_script.py'
            command = [sys.executable, str(script), '--comparison-directory', str(output), '--output-ps1', str(root / 'delete.ps1')]
            blocked = subprocess.run(command, capture_output=True, text=True)
            self.assertNotEqual(blocked.returncode, 0, 'Duplicate guard did not block generation.')
            generated = subprocess.run(command + ['--allow-duplicate-keys'], capture_output=True, text=True)
            self.assertEqual(generated.returncode, 0, generated.stdout + generated.stderr)
            text = (root / 'delete.ps1').read_text(encoding='utf-8-sig')
            self.assertIn('WhatIf', text)
            self.assertIn('Execute', text)
            workbook = root / 'comparison.xlsx'
            exported = subprocess.run([sys.executable, str(ROOT / 'Scripts/Export/export_comparison_to_excel.py'),
                '--comparison-directory', str(output), '--output-xlsx', str(workbook)], capture_output=True, text=True)
            self.assertEqual(exported.returncode, 0, exported.stdout + exported.stderr)
            with zipfile.ZipFile(workbook) as package:
                self.assertIsNone(package.testzip())
                sheets = ET.fromstring(package.read('xl/workbook.xml'))
                self.assertTrue(any(node.get('name') == 'TargetOlderThanSource' for node in sheets.iter()))


class PermissionComparisonTests(unittest.TestCase):
    def test_permission_differences_disabled_users_and_limited_access(self):
        def row(name, levels='Read', web='/source'):
            return {'ObjectScope': 'Web', 'ObjectServerRelativeUrl': web, 'WebUrl': 'https://workplacecloudhub.sharepoint.com' + web,
                    'PrincipalType': 'User', 'PrincipalLoginName': name + '@workplacecloudhub.com', 'PermissionLevels': levels}

        source = [row('matched'), row('disabled'), row('missing'), row('reduced', 'Read|Edit'), row('limited', 'Limited Access')]
        target = [row('matched', web='/target'), row('reduced', web='/target')]
        aliases = {name + '@workplacecloudhub.com': name + '@workplacecloudhub.com' for name in ['matched','disabled','missing','reduced','limited']}
        with tempfile.TemporaryDirectory() as directory:
            report = PERMISSIONS.write_permission_comparison_report(source, target, Path(directory), 'Synthetic', '/source', '/target',
                path_mappings=[('/source','/target')], entra_user_aliases=aliases, disabled_entra_users={'disabled@workplacecloudhub.com'})
            self.assertEqual(report['MatchedPermissions'], 1)
            self.assertEqual(report['DisabledEntraUsersNotInSPO'], 1)
            # MissingInSPO includes the identity with reduced permission levels.
            self.assertEqual(report['MissingInSPO'], 2)
            self.assertEqual(report['TargetHasLessPermissions'], 1)
            with report['Summary'].open(encoding='utf-8-sig') as handle:
                summary = next(csv.DictReader(handle))
            self.assertEqual(summary['SourceLimitedAccessOnlyIgnored'], '1')
            self.assertTrue(report['Html'].is_file())
            with zipfile.ZipFile(report['Excel']) as package:
                self.assertIsNone(package.testzip())


if __name__ == '__main__':
    unittest.main(verbosity=2)
