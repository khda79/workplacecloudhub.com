"""Offline regression tests; all evidence is synthetic and temporary."""
import csv
import json
import importlib.util
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
COMPARE_DIR = ROOT / 'Scripts' / 'Compare'
if str(COMPARE_DIR) not in sys.path:
    sys.path.insert(0, str(COMPARE_DIR))


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'Scripts' / 'Compare' / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


FILES = load('compare_sp_source_target_file_inventories')
PERMISSIONS = load('compare_sp_source_target_permissions')
GLOBAL = load('build_global_report')
GLOBAL_PERMISSIONS = load('build_global_permissions_report')
EVIDENCE = load('scan_evidence')
HTML = load('report_html')


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
    def test_root_scope_includes_descendants_but_not_sibling_prefix(self):
        roots = {'https://example.com/sites/bu'}
        self.assertTrue(FILES.web_is_in_scope('https://example.com/sites/bu', roots))
        self.assertTrue(FILES.web_is_in_scope('https://example.com/sites/bu/subsite', roots))
        self.assertFalse(FILES.web_is_in_scope('https://example.com/sites/bu-other', roots))

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
            summary = read('Summary.csv')[0]
            self.assertEqual(summary['SuccessPercent'], '75.00%')
            self.assertIn('75.00%', Path(summary['HtmlSummary']).read_text(encoding='utf-8'))
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


class HtmlReportTests(unittest.TestCase):
    def test_public_branding_template_has_no_client_logo(self):
        template = json.loads((ROOT / 'Config' / 'report-branding.json.template').read_text(encoding='utf-8'))
        self.assertEqual(template['ClientLogoPath'], '')

    def test_report_header_has_one_logo_with_client_preferred_and_fallback(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            branding = project / 'Config'
            branding.mkdir(parents=True)
            logo = branding / 'client-logo.png'
            logo.write_bytes(b'\x89PNG\r\n\x1a\n' + b'example')
            product_logo = project / 'product-logo.png'
            product_logo.write_bytes(b'\x89PNG\r\n\x1a\n' + b'product')
            def header():
                rendered = HTML.render_report('Title', 'Migration diagnostics', 'now', 'Review', 'note', '', '', 'Footer')
                return rendered.split('</header>', 1)[0]
            (branding / 'report-branding.json.txt').write_text(json.dumps({'ClientLogoPath': logo.name}), encoding='utf-8')
            with patch.object(HTML, '__file__', str(project / 'Scripts' / 'Compare' / 'report_html.py')):
                self.assertIn('alt="Client logo"', header())
                self.assertEqual(header().count('<img '), 1)
                (branding / 'report-branding.json.txt').write_text(json.dumps({'ClientLogoPath': '../client-logo.png'}), encoding='utf-8')
                self.assertEqual(header().count('<img '), 0)
                (branding / 'report-branding.json.txt').write_text(json.dumps({'ClientLogoPath': '..\\client-logo.png'}), encoding='utf-8')
                self.assertEqual(header().count('<img '), 0)
                (branding / 'report-branding.json.txt').write_text(json.dumps({'ClientLogoPath': '../client-logo.png', 'WorkplaceCloudHubLogoPath': product_logo.name}), encoding='utf-8')
                self.assertIn('alt="WorkplaceCloudHub"', header())
                self.assertEqual(header().count('<img '), 1)

    def test_global_report_uses_latest_summary_and_exposes_metrics(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            migration = root / 'Migrations' / 'Example'
            migration.mkdir(parents=True)
            (migration / 'migration.config.psd1').write_text("@{Name='Example'}", encoding='utf-8')
            (migration / 'migration.mapping.txt').write_text('https://source/sites/example https://target/sites/example\n', encoding='utf-8')
            for stamp, matched in [('20261001-100000', '5'), ('20261002-100000', '8')]:
                folder = migration / 'comparisons' / 'files' / ('Example-files-' + stamp)
                folder.mkdir(parents=True)
                with (folder / 'Summary.csv').open('w', encoding='utf-8', newline='') as handle:
                    writer = csv.DictWriter(handle, fieldnames=['SourceCsv', 'TargetCsv', 'SourceUniqueKeys', 'TargetUniqueKeys', 'MatchedKeys', 'MissingInTarget', 'ExtraInTarget'], delimiter=';')
                    writer.writeheader()
                    writer.writerow({'SourceCsv': 'SP2019-FileInventory-Example-20261002-080000.csv', 'TargetCsv': 'SPO-FileInventory-Example-20261002-090000.csv',
                                     'SourceUniqueKeys': '10', 'TargetUniqueKeys': '9', 'MatchedKeys': matched, 'MissingInTarget': '2', 'ExtraInTarget': '1'})
            output = GLOBAL.build(root / 'Migrations', root / 'out')
            self.assertTrue(output.is_file())
            with output.with_suffix('.csv').open(encoding='utf-8-sig', newline='') as handle:
                row = next(csv.DictReader(handle, delimiter=';'))
            self.assertEqual(row['SuccessPercent'], '80.00%')
            self.assertEqual(row['SourceScannedAt'], '2026-10-02 08:00:00')
            self.assertEqual(row['MissingInTarget'], '2')
            self.assertIn('https://source/sites/example', output.read_text(encoding='utf-8'))
            page = output.read_text(encoding='utf-8')
            self.assertIn('Open Excel report', page)
            self.assertLess(page.index('Open Excel report'), page.index('<p class="intro">'))
            with zipfile.ZipFile(output.with_suffix('.xlsx')) as package:
                self.assertIsNone(package.testzip())
                workbook = ET.fromstring(package.read('xl/workbook.xml'))
                self.assertEqual([node.get('name') for node in workbook if node.tag.endswith('sheets') for node in node],
                                 ['Latest comparisons', 'Definitions'])
                namespace = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                sheet = ET.fromstring(package.read('xl/worksheets/sheet1.xml'))
                self.assertEqual(sheet.find(".//x:c[@r='B2']/x:v", namespace).text, '0.8')
                self.assertEqual(sheet.find(".//x:c[@r='B2']", namespace).get('s'), '2')
                self.assertEqual(sheet.find(".//x:c[@r='E2']", namespace).get('s'), '1')

    def test_global_permissions_report_uses_latest_summary_and_excel_link(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            migration = root / 'Migrations' / 'Example'
            migration.mkdir(parents=True)
            (migration / 'migration.config.psd1').write_text("@{Name='Example'}", encoding='utf-8')
            (migration / 'migration.mapping.txt').write_text('https://source/sites/example https://target/sites/example\n', encoding='utf-8')
            for run_stamp, matched in [('20261001-100000', '5'), ('20261002-100000', '8')]:
                folder = migration / 'comparisons' / 'permissions' / ('Example-permissions-' + run_stamp)
                folder.mkdir(parents=True)
                with (folder / 'Summary.csv').open('w', encoding='utf-8', newline='') as handle:
                    writer = csv.DictWriter(handle, fieldnames=['SourceCsv', 'TargetCsv', 'SourceUniqueKeys', 'TargetUniqueKeys', 'MatchedPermissions', 'MissingInSPO', 'ExtraInSPO'])
                    writer.writeheader()
                    writer.writerow({'SourceCsv': 'SP2019-PermissionInventory-Example-20261002-080000.csv',
                                     'TargetCsv': 'SPO-PermissionInventory-Example-20261002-090000.csv',
                                     'SourceUniqueKeys': '10', 'TargetUniqueKeys': '9', 'MatchedPermissions': matched,
                                     'MissingInSPO': '2', 'ExtraInSPO': '1'})
            output = GLOBAL_PERMISSIONS.build(root / 'Migrations', root / 'out')
            self.assertTrue(output.is_file())
            with output.with_suffix('.csv').open(encoding='utf-8-sig', newline='') as handle:
                row = next(csv.DictReader(handle, delimiter=';'))
            self.assertEqual(row['SuccessPercent'], '80.00%')
            self.assertEqual(row['MatchedPermissions'], '8')
            page = output.read_text(encoding='utf-8')
            self.assertIn('Global permissions comparison report', page)
            self.assertLess(page.index('Open Excel report'), page.index('<p class="intro">'))
            with zipfile.ZipFile(output.with_suffix('.xlsx')) as package:
                self.assertIsNone(package.testzip())
                workbook = ET.fromstring(package.read('xl/workbook.xml'))
                self.assertEqual([node.get('name') for group in workbook if group.tag.endswith('sheets') for node in group],
                                 ['Latest permissions', 'Definitions'])

    def test_manifest_records_true_row_count_and_sha256(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'scan.csv'
            path.write_text('Name,Value\nfile,"line 1\nline 2"\n', encoding='utf-8')
            receipt = EVIDENCE.write_manifest(path, 'Source', 'File', 'https://source/sites/example')
            data = json.loads(receipt.read_text(encoding='utf-8'))
            self.assertEqual(data['Rows'], 1)
            self.assertEqual(data['Side'], 'Source')
            self.assertEqual(len(data['Sha256']), 64)
            self.assertEqual(EVIDENCE.describe_pair(path, path)['ScanEvidenceStatus'], 'Verified')
            path.write_text('Name,Value\nchanged,1\n', encoding='utf-8')
            with self.assertRaises(ValueError):
                EVIDENCE.describe_pair(path, path)

    def test_comparators_find_adjacent_report_module_with_portable_python(self):
        with tempfile.TemporaryDirectory() as directory:
            shutil.copy2(COMPARE_DIR / 'report_html.py', Path(directory) / 'report_html.py')
            shutil.copy2(COMPARE_DIR / 'scan_evidence.py', Path(directory) / 'scan_evidence.py')
            shutil.copy2(ROOT / 'Scripts' / 'console_lifecycle.py', Path(directory) / 'console_lifecycle.py')
            portable_python = ROOT / 'Tools' / 'Python' / 'python.exe'
            executable = str(portable_python if portable_python.is_file() else sys.executable)
            for name in ('compare_sp_source_target_file_inventories.py', 'compare_sp_source_target_permissions.py'):
                copied = Path(directory) / name
                shutil.copy2(COMPARE_DIR / name, copied)
                run = subprocess.run([executable, str(copied), '--help'], capture_output=True, text=True)
                self.assertEqual(run.returncode, 0, f'{name}: {run.stderr}')

    def test_branded_file_report_is_portable_and_escapes_inventory_text(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            detail = root / 'details & more.csv'
            detail.write_text('sample', encoding='utf-8')
            rows = [
                {'Status': 'Review', 'MissingInTarget': 1, 'WebPath': '/site',
                 'SourceLibraryTitle': '<script>alert(1)</script>' if index == 0 else f'Library {index}'}
                for index in range(21)
            ]
            path = root / 'summary.html'
            FILES.create_file_html_summary(
                path, 'Synthetic <file> comparison', {'MissingInTarget': 21}, rows,
                [('Details', detail, 'Full detail')],
            )
            report = path.read_text(encoding='utf-8')
            self.assertIn('WorkplaceCloudHub', report)
            self.assertIn('data:image/png;base64,', report)
            self.assertIn('Showing 20 of 21 libraries with differences', report)
            self.assertIn('href="details%20%26%20more.csv"', report)
            self.assertIn('&lt;script&gt;alert(1)&lt;/script&gt;', report)
            self.assertNotIn('<script>alert(1)</script>', report)
            self.assertIn('Synthetic &lt;file&gt; comparison', report)

    def test_permission_report_keeps_scope_warning_and_skips_missing_links(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / 'summary.html'
            PERMISSIONS.create_permission_html_summary(
                path, 'Synthetic permissions',
                {'ScopeWarning': 'Source < target', 'DisabledEntraUsersNotInSPO': 1},
                [], [], [('Missing export', root / 'absent.csv', 'Not generated')],
            )
            report = path.read_text(encoding='utf-8')
            self.assertIn('WorkplaceCloudHub', report)
            self.assertIn('Source &lt; target', report)
            self.assertIn('No report files found.', report)
            self.assertNotIn('href="absent.csv"', report)
            self.assertIn('Showing 0 of 0 objects with differences', report)


if __name__ == '__main__':
    unittest.main(verbosity=2)
