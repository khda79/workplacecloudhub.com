"""Synthetic regression tests for master-to-thematic report generation, v1.2.4."""

import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

import generate_thematic_reports as gen


class ThematicReportsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=os.environ.get('THEMATIC_TEST_ROOT'))
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'master/Main.Report'
        self.source.mkdir(parents=True)
        (self.root / 'master/Main.SemanticModel').mkdir()
        self.output = self.root / 'generated'
        self.manifest = self.root / 'themes.json.txt'
        self.manifest.write_bytes(gen.encode({'version': 1, 'reports': [
            {'key': 'licensing', 'order': 3, 'displayName': 'SWI - Licensing', 'pages': ['capacity', 'recovery']},
            {'key': 'other', 'order': 4, 'displayName': 'SWI - Other', 'pages': ['other']}]}))
        self.put('definition.pbir', {'version': '4.0', 'datasetReference': {'byPath': {'path': '../Main.SemanticModel'}}})
        self.put('.platform', {'metadata': {'type': 'Report', 'displayName': 'Master'},
                               'config': {'logicalId': '7a73d6e2-6551-42d1-8ea3-4d32c03d3354'}})
        self.put('definition/version.json', {'version': '2.0.0'})
        self.put('definition/report.json', {'annotations': [{'name': 'defaultPage', 'value': 'other'}]})
        self.put('definition/pages/pages.json', {'pageOrder': ['other', 'capacity', 'recovery'], 'activePageName': 'other'})
        for page in ('capacity', 'recovery', 'other'):
            self.put(f'definition/pages/{page}/page.json', {'name': page, 'displayName': page})
            self.put(f'definition/pages/{page}/visuals/chart/visual.json', {'name': 'chart', 'visual': {'visualType': 'tableEx'}})
            self.put(f'definition/bookmarks/{page}.bookmark.json', {
                'name': page, 'displayName': page, 'options': {'applyOnlyToTargetVisuals': True,
                    'targetVisualNames': ['chart'], 'suppressData': True},
                'explorationState': {'activeSection': page, 'sections': {page: {'visualContainers': {'chart': {}}}}}})
        self.put('definition/bookmarks/bookmarks.json', {'items': [
            {'name': 'license-tabs', 'displayName': 'License tabs', 'children': ['capacity', 'recovery']}, 'other']})
        self.put('StaticResources/RegisteredResources/logo.svg', b'<svg/>')
        self.put('.pbi/localSettings.json', b'{"private":true}')
        self.put('.pbi/cache.abf', b'PRIVATE CACHE')

    def put(self, path, value):
        f = self.source / path
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_bytes(value if isinstance(value, bytes) else gen.encode(value))

    def read(self, path):
        return gen.decode((self.source / path).read_bytes())

    def run_generation(self, **kwargs):
        return gen.generate(self.source, self.manifest, 'licensing', self.output, **kwargs)

    def fault(self, path, edit):
        document = self.read(path)
        edit(document)
        self.put(path, document)

    def test_default_plan_writes_nothing(self):
        result = self.run_generation()
        self.assertFalse(self.output.exists())
        self.assertFalse(result['written'])
        self.assertEqual(result['sourcePreservation'], 'verified-byte-identical')

    def test_default_profile_selection_and_explicit_override(self):
        self.assertEqual(gen.select_manifest(self.source), gen.DEFAULT_MANIFEST)
        document = self.read('definition/pages/pages.json')
        document['pageOrder'].extend(sorted(gen.MOBILE_PAGES))
        self.put('definition/pages/pages.json', document)
        self.assertEqual(gen.select_manifest(self.source), gen.PC_MOBILE_MANIFEST)
        self.assertEqual(gen.select_manifest(self.source, self.manifest), self.manifest)
        document['pageOrder'].pop()
        self.put('definition/pages/pages.json', document)
        with self.assertRaisesRegex(gen.GenerationError, 'both mobile pages'):
            gen.select_manifest(self.source)
        self.assertEqual(gen.select_manifest(self.source, self.manifest), self.manifest)

    def test_profiles_cover_30_and_32_pages_without_changing_existing_theme_keys(self):
        legacy = gen.decode(gen.DEFAULT_MANIFEST.read_bytes())
        modern = gen.decode(gen.PC_MOBILE_MANIFEST.read_bytes())
        old_pages = {p for t in legacy['reports'] for p in t['pages']}
        new_pages = {p for t in modern['reports'] for p in t['pages']}
        self.assertEqual((len(legacy['reports']), len(old_pages)), (11, 30))
        self.assertEqual((len(modern['reports']), len(new_pages)), (12, 32))
        self.assertEqual(new_pages - old_pages, gen.MOBILE_PAGES)
        for theme in legacy['reports']:
            replacement = next(t for t in modern['reports'] if t['key'] == theme['key'])
            self.assertEqual(replacement['pages'], theme['pages'])
        gen.read_manifest(gen.DEFAULT_MANIFEST, old_pages)
        gen.read_manifest(gen.PC_MOBILE_MANIFEST, new_pages)
        with self.assertRaisesRegex(gen.GenerationError, 'Unknown configured pages'):
            gen.read_manifest(gen.PC_MOBILE_MANIFEST, old_pages)
        self.assertEqual(modern['navigation']['layoutAdjustments'], [])
        self.assertEqual(modern['navigation']['explorerSelectors'], [])
        self.assertEqual(modern['navigation']['protectedButtonPages'], [])

    def test_root_projects_bundle_has_valid_paths_and_distinct_receipts(self):
        document = gen.decode(self.manifest.read_bytes())
        document['outputLayout'] = 'root-projects'
        self.manifest.write_bytes(gen.encode(document))
        result = gen.generate(self.source, self.manifest, 'all', self.output, write=True)
        self.assertEqual(len(list(self.output.glob('*.pbip'))), 2)
        self.assertEqual(len(list((self.output / 'receipts').glob('*.json.txt'))), 2)
        for receipt in result['reports']:
            project = gen.decode((self.output / receipt['projectFile']).read_bytes())
            report = self.output / project['artifacts'][0]['report']['path']
            self.assertEqual(report, self.output / receipt['reportFolder'])
            self.assertEqual(gen.fingerprints(gen.snapshot(report)), receipt['generatedReportFiles'])
            binding = gen.decode((report / 'definition.pbir').read_bytes())
            self.assertEqual((report / binding['datasetReference']['byPath']['path']).resolve(),
                             (self.root / 'master/Main.SemanticModel').resolve())
            self.assertTrue((self.output / receipt['receiptFile']).is_file())
            self.assertFalse((self.output / receipt['theme'] / 'generation-receipt.json.txt').exists())
        self.assertTrue((self.output / 'generation-manifest.json.txt').is_file())

    def test_invalid_bundle_layout_fails_before_writing(self):
        with self.assertRaisesRegex(gen.GenerationError, 'require root-projects'):
            gen.generate(self.source, self.manifest, 'all', self.output, write=True)
        self.assertFalse(self.output.exists())
        document = gen.decode(self.manifest.read_bytes())
        document['outputLayout'] = '../unsafe'
        self.manifest.write_bytes(gen.encode(document))
        with self.assertRaisesRegex(gen.GenerationError, 'Unsupported output layout'):
            self.run_generation(write=True)
        self.assertFalse(self.output.exists())

    def test_friendly_file_name_preserves_content_and_logical_identity(self):
        baseline = self.run_generation(write=True)
        original = self.output / 'SWI-licensing.Report'
        document = gen.decode(self.manifest.read_bytes())
        document['reports'][0].update(fileName='03 · Licensing', displayName='03 · Licensing')
        self.manifest.write_bytes(gen.encode(document))
        target = self.root / 'friendly'
        result = gen.generate(self.source, self.manifest, 'licensing', target, write=True)
        project = gen.decode((target / '03 · Licensing.pbip').read_bytes())
        self.assertEqual(project['artifacts'][0]['report']['path'], '03 · Licensing.Report')
        report = target / project['artifacts'][0]['report']['path']
        for path in baseline['generatedReportFiles']:
            if path not in {'.platform', 'definition.pbir'}:
                self.assertEqual((original / path).read_bytes(), (report / path).read_bytes())
        self.assertEqual(gen.decode((original / '.platform').read_bytes())['config'],
                         gen.decode((report / '.platform').read_bytes())['config'])
        self.assertEqual(gen.decode((report / '.platform').read_bytes())['metadata']['displayName'], '03 · Licensing')
        self.assertEqual((report / gen.decode((report / 'definition.pbir').read_bytes())['datasetReference']['byPath']['path']).resolve(),
                         (self.root / 'master/Main.SemanticModel').resolve())
        self.assertEqual(result['projectFile'], '03 · Licensing.pbip')

    def test_unsafe_or_duplicate_file_names_are_refused(self):
        original = gen.decode(self.manifest.read_bytes())
        for name in ['../escape', 'bad/name', 'bad\\name', 'CON', 'NUL.txt', 'bad.', 'bad?', '', ' leading']:
            with self.subTest(name=name):
                document = gen.decode(gen.encode(original))
                document['reports'][0]['fileName'] = name
                self.manifest.write_bytes(gen.encode(document))
                with self.assertRaisesRegex(gen.GenerationError, 'Invalid report file name'):
                    self.run_generation()
                self.assertFalse(self.output.exists())
        document = gen.decode(gen.encode(original))
        document['reports'][0]['fileName'] = 'Same'
        document['reports'][1]['fileName'] = 'same'
        self.manifest.write_bytes(gen.encode(document))
        with self.assertRaisesRegex(gen.GenerationError, 'Duplicate report file name'):
            self.run_generation()

    def test_generate_preserves_pages_and_visuals_byte_for_byte(self):
        before = {p.relative_to(self.source).as_posix(): p.read_bytes() for p in self.source.rglob('*') if p.is_file()}
        result = self.run_generation(write=True)
        report = self.output / 'SWI-licensing.Report'
        for path, data in before.items():
            if path.startswith(('definition/pages/capacity/', 'definition/pages/recovery/')):
                self.assertEqual((report / path).read_bytes(), data)
        after = {p.relative_to(self.source).as_posix(): p.read_bytes() for p in self.source.rglob('*') if p.is_file()}
        self.assertEqual(before, after)
        self.assertFalse((report / '.pbi').exists())
        self.assertFalse((report / 'definition/pages/other').exists())
        self.assertFalse(any(p.name.endswith('.SemanticModel') for p in self.output.rglob('*')))
        self.assertEqual(result['bookmarks'], 2)
        self.assertFalse(result['publicationPerformed'])

    def test_binding_points_to_original_model_and_active_page_is_reset(self):
        self.run_generation(write=True)
        report = self.output / 'SWI-licensing.Report'
        binding = gen.decode((report / 'definition.pbir').read_bytes())
        self.assertEqual((report / binding['datasetReference']['byPath']['path']).resolve(),
                         (self.root / 'master/Main.SemanticModel').resolve())
        pages = gen.decode((report / 'definition/pages/pages.json').read_bytes())
        self.assertEqual(pages, {'pageOrder': ['capacity', 'recovery'], 'activePageName': 'capacity'})
        self.assertEqual(gen.decode((report / 'definition/report.json').read_bytes())['annotations'][0]['value'], 'capacity')

    def test_bookmark_groups_pruned_without_changing_retained_bookmarks(self):
        self.run_generation(write=True)
        report = self.output / 'SWI-licensing.Report'
        meta = gen.decode((report / 'definition/bookmarks/bookmarks.json').read_bytes())
        self.assertEqual(meta['items'], [{'name': 'license-tabs', 'displayName': 'License tabs',
                                         'children': ['capacity', 'recovery']}])
        for name in ('capacity', 'recovery'):
            path = f'definition/bookmarks/{name}.bookmark.json'
            self.assertEqual((self.source / path).read_bytes(), (report / path).read_bytes())

    def test_repeated_output_is_refused_even_when_generated(self):
        self.run_generation(write=True)
        before = (self.output / 'generation-receipt.json.txt').read_bytes()
        with self.assertRaisesRegex(gen.GenerationError, 'already exists'):
            self.run_generation(write=True)
        self.assertEqual(before, (self.output / 'generation-receipt.json.txt').read_bytes())

    def test_foreign_output_is_never_overwritten(self):
        self.output.mkdir()
        (self.output / 'keep').write_bytes(b'keep')
        with self.assertRaises(gen.GenerationError):
            self.run_generation(write=True)
        self.assertEqual((self.output / 'keep').read_bytes(), b'keep')

    def test_output_in_master_or_ancestor_is_refused(self):
        for output in (self.source / 'derived', self.source.parent / 'derived', self.root):
            with self.assertRaises(gen.GenerationError):
                gen.generate(self.source, self.manifest, 'licensing', output, write=True)

    def test_external_page_dependency_is_refused(self):
        self.put('definition/pages/capacity/visuals/link/visual.json', {
            'name': 'link', 'navigationSection': {'expr': {'Literal': {'Value': "'other'"}}}})
        with self.assertRaisesRegex(gen.GenerationError, 'External page'):
            self.run_generation(write=True)
        self.assertFalse(self.output.exists())

    def test_dynamic_navigation_is_refused(self):
        self.put('definition/pages/capacity/visuals/link/visual.json', {
            'name': 'link', 'navigationSection': {'expr': {'Measure': {'Property': 'Target'}}}})
        with self.assertRaisesRegex(gen.GenerationError, 'Dynamic dependency'):
            self.run_generation()

    def configure_support(self):
        data = gen.decode(self.manifest.read_bytes())
        data['reports'][0]['supportingPages'] = ['other']
        self.manifest.write_bytes(gen.encode(data))
        self.put('definition/pages/capacity/visuals/link/visual.json', {
            'name': 'link', 'navigationSection': {'expr': {'Literal': {'Value': "'other'"}}}})

    def test_explicit_support_preserves_navigation_and_hides_only_support(self):
        self.configure_support()
        before = gen.fingerprints(gen.snapshot(self.source))
        result = self.run_generation(write=True)
        report = self.output / 'SWI-licensing.Report'
        self.assertEqual(result['supportingPages'], ['other'])
        self.assertEqual(result['bookmarks'], 3)
        meta = gen.decode((report / 'definition/pages/pages.json').read_bytes())
        self.assertEqual(meta['pageOrder'], ['capacity', 'recovery', 'other'])
        for page in ('capacity', 'recovery'):
            path = f'definition/pages/{page}/page.json'
            self.assertEqual((report / path).read_bytes(), (self.source / path).read_bytes())
        hidden = gen.decode((report / 'definition/pages/other/page.json').read_bytes())
        self.assertEqual(hidden['visibility'], 'HiddenInViewMode')
        link = 'definition/pages/capacity/visuals/link/visual.json'
        self.assertEqual((report / link).read_bytes(), (self.source / link).read_bytes())
        self.assertEqual(before, gen.fingerprints(gen.snapshot(self.source)))

    def test_support_does_not_silently_expand_to_unknown_destinations(self):
        self.configure_support()
        self.put('definition/pages/other/visuals/next/visual.json', {
            'name': 'next', 'pageName': 'missing'})
        with self.assertRaisesRegex(gen.GenerationError, 'External page'):
            self.run_generation(write=True)
        self.assertFalse(self.output.exists())

    def test_invalid_support_list_is_refused(self):
        for pages in (['capacity'], ['missing'], ['other', 'other']):
            data = gen.decode(self.manifest.read_bytes())
            data['reports'][0]['supportingPages'] = pages
            self.manifest.write_bytes(gen.encode(data))
            with self.assertRaises(gen.GenerationError):
                self.run_generation()

    def test_reviewed_orphan_target_removed_without_changing_bookmark_state(self):
        self.fault('definition/bookmarks/capacity.bookmark.json',
                   lambda d: d['options']['targetVisualNames'].append('deleted'))
        data = gen.decode(self.manifest.read_bytes())
        data['orphanBookmarkTargets'] = {'capacity': ['deleted']}
        self.manifest.write_bytes(gen.encode(data))
        result = self.run_generation(write=True)
        path = 'definition/bookmarks/capacity.bookmark.json'
        original = self.read(path)
        generated = gen.decode((self.output / 'SWI-licensing.Report' / path).read_bytes())
        self.assertEqual(generated['explorationState'], original['explorationState'])
        self.assertEqual(generated['options']['targetVisualNames'], ['chart'])
        self.assertEqual(original['options']['targetVisualNames'], ['chart', 'deleted'])
        self.assertEqual(result['bookmarkRepairs'], [{'bookmark': 'capacity', 'removedMissingTargets': ['deleted']}])

    def test_orphan_allowlist_does_not_remove_existing_target(self):
        data = gen.decode(self.manifest.read_bytes())
        data['orphanBookmarkTargets'] = {'capacity': ['chart']}
        self.manifest.write_bytes(gen.encode(data))
        result = self.run_generation(write=True)
        path = 'definition/bookmarks/capacity.bookmark.json'
        self.assertEqual((self.output / 'SWI-licensing.Report' / path).read_bytes(), (self.source / path).read_bytes())
        self.assertEqual(result['bookmarkRepairs'], [])

    def test_orphan_allowlist_does_not_mask_dangling_bookmark_state(self):
        self.fault('definition/bookmarks/capacity.bookmark.json',
                   lambda d: d['explorationState']['sections']['capacity']['visualContainers'].update({'deleted': {}}))
        data = gen.decode(self.manifest.read_bytes())
        data['orphanBookmarkTargets'] = {'capacity': ['deleted']}
        self.manifest.write_bytes(gen.encode(data))
        with self.assertRaisesRegex(gen.GenerationError, 'state targets missing'):
            self.run_generation()

    def test_missing_bookmark_is_refused(self):
        self.put('definition/pages/capacity/visuals/link/visual.json', {
            'name': 'link', 'bookmark': {'expr': {'Literal': {'Value': "'missing'"}}}})
        with self.assertRaisesRegex(gen.GenerationError, 'Missing or external bookmarks'):
            self.run_generation()

    def test_cross_theme_bookmark_is_not_silently_trimmed(self):
        self.fault('definition/bookmarks/capacity.bookmark.json',
                   lambda d: d['explorationState']['sections'].update({'other': {}}))
        with self.assertRaisesRegex(gen.GenerationError, 'Cross-theme bookmark'):
            self.run_generation()

    def test_unknown_bookmark_target_is_refused(self):
        self.fault('definition/bookmarks/capacity.bookmark.json',
                   lambda d: d['options']['targetVisualNames'].append('missing'))
        with self.assertRaisesRegex(gen.GenerationError, 'missing visuals'):
            self.run_generation()

    def test_parent_group_is_valid_even_when_sorted_before_group_file(self):
        self.put('definition/pages/capacity/visuals/chart/visual.json', {'name': 'chart', 'parentGroupName': 'zgroup'})
        self.put('definition/pages/capacity/visuals/zgroup/visual.json', {'name': 'zgroup', 'visualGroup': {}})
        self.fault('definition/bookmarks/capacity.bookmark.json',
                   lambda d: d['options']['targetVisualNames'].append('zgroup'))
        self.assertEqual(self.run_generation()['visuals'], 3)

    def test_unclassified_page_is_refused(self):
        data = gen.decode(self.manifest.read_bytes())
        data['reports'].pop()
        self.manifest.write_bytes(gen.encode(data))
        with self.assertRaisesRegex(gen.GenerationError, 'Unclassified'):
            self.run_generation()

    def test_duplicate_principal_assignment_is_refused(self):
        data = gen.decode(self.manifest.read_bytes())
        data['reports'][1]['pages'].append('capacity')
        self.manifest.write_bytes(gen.encode(data))
        with self.assertRaisesRegex(gen.GenerationError, 'multiple themes'):
            self.run_generation()

    def test_removed_page_is_refused(self):
        data = gen.decode(self.manifest.read_bytes())
        data['reports'][0]['pages'].append('absent')
        self.manifest.write_bytes(gen.encode(data))
        with self.assertRaisesRegex(gen.GenerationError, 'Unknown configured'):
            self.run_generation()

    def test_physical_pages_not_in_order_are_refused(self):
        self.put('definition/pages/unlisted/page.json', {'name': 'unlisted'})
        with self.assertRaisesRegex(gen.GenerationError, 'page order'):
            self.run_generation()

    def test_duplicate_json_keys_are_refused(self):
        self.put('definition/report.json', b'{"settings":{},"settings":{}}')
        with self.assertRaisesRegex(gen.GenerationError, 'Duplicate JSON'):
            self.run_generation()

    def test_service_binding_is_explicit_and_does_not_copy_a_model(self):
        result = self.run_generation(write=True, semantic_model_id='8bdc4589-4760-4a04-b23b-ad6db1a67163')
        binding = gen.decode((self.output / 'SWI-licensing.Report/definition.pbir').read_bytes())
        self.assertEqual(binding['datasetReference'], {'byConnection': {
            'connectionString': 'semanticmodelid=8bdc4589-4760-4a04-b23b-ad6db1a67163'}})
        self.assertEqual(result['modelBinding'], 'shared-service-model-reference-only')

    def test_generated_report_is_reproducible_with_stable_separate_identity(self):
        captured = gen.snapshot(self.source)
        path = self.output / 'SWI-licensing.Report'
        first, _, _ = gen.build(self.source, captured, self.manifest, 'licensing', path)
        second, _, _ = gen.build(self.source, captured, self.manifest, 'licensing', path)
        self.assertEqual(first, second)
        self.assertNotEqual(gen.decode(first['.platform'])['config']['logicalId'],
                            self.read('.platform')['config']['logicalId'])

    def test_master_drift_during_plan_aborts_before_writing(self):
        original = gen.snapshot
        calls = 0
        def racing_snapshot(source):
            nonlocal calls
            calls += 1
            result = original(source)
            if calls == 2:
                result['definition/report.json'] += b' '
            return result
        with mock.patch.object(gen, 'snapshot', side_effect=racing_snapshot):
            with self.assertRaisesRegex(gen.GenerationError, 'Master changed'):
                self.run_generation(write=True)
        self.assertFalse(self.output.exists())

    def test_next_generation_inherits_master_change_without_overwriting_previous(self):
        self.run_generation(write=True)
        previous = self.output / 'SWI-licensing.Report/definition/pages/capacity/visuals/chart/visual.json'
        old_bytes = previous.read_bytes()
        self.fault('definition/pages/capacity/visuals/chart/visual.json',
                   lambda d: d.update({'isHidden': True}))
        next_output = self.root / 'generated-next'
        gen.generate(self.source, self.manifest, 'licensing', next_output, write=True)
        self.assertEqual(previous.read_bytes(), old_bytes)
        changed = next_output / 'SWI-licensing.Report/definition/pages/capacity/visuals/chart/visual.json'
        self.assertTrue(gen.decode(changed.read_bytes())['isHidden'])
        unchanged = 'SWI-licensing.Report/definition/pages/recovery/visuals/chart/visual.json'
        self.assertEqual((self.output / unchanged).read_bytes(), (next_output / unchanged).read_bytes())

    def test_missing_visual_interaction_target_is_refused(self):
        self.fault('definition/pages/capacity/page.json',
                   lambda d: d.update({'visualInteractions': [{'source': 'chart', 'target': 'missing'}]}))
        with self.assertRaisesRegex(gen.GenerationError, 'Interaction'):
            self.run_generation()

    def test_missing_parent_group_is_refused(self):
        self.fault('definition/pages/capacity/visuals/chart/visual.json',
                   lambda d: d.update({'parentGroupName': 'absent'}))
        with self.assertRaisesRegex(gen.GenerationError, 'Missing parent'):
            self.run_generation()

    def test_link_guard_is_exercised_without_symlink_privileges(self):
        with mock.patch.object(gen, 'linked', side_effect=lambda p: p == self.source):
            with self.assertRaisesRegex(gen.GenerationError, 'Linked'):
                self.run_generation()

    def test_corrupt_staging_aborts_before_promotion(self):
        original = gen.snapshot
        def corrupt_snapshot(source):
            result = original(source)
            if '.staging-' in str(source):
                result['definition/report.json'] += b' '
            return result
        with mock.patch.object(gen, 'snapshot', side_effect=corrupt_snapshot):
            with self.assertRaisesRegex(gen.GenerationError, 'read-back failed'):
                self.run_generation(write=True)
        self.assertFalse(self.output.exists())

    def tabs(self, **policy):
        data = gen.decode(self.manifest.read_bytes())
        data['navigation'] = dict(mode='tabs', **policy)
        self.manifest.write_bytes(gen.encode(data))

    def nav_button(self, page, name, action, **properties):
        properties['type'] = {'expr': {'Literal': {'Value': "'" + action + "'"}}}
        self.put(f'definition/pages/{page}/visuals/{name}/visual.json', {
            'name': name, 'visual': {'visualType': 'actionButton',
                'visualContainerObjects': {'visualLink': [{'properties': properties}]}}})

    def test_tabs_remove_navigation_without_touching_master_or_views(self):
        self.tabs()
        self.nav_button('capacity', 'nav', 'PageNavigation', navigationSection={'expr': {'Literal': {'Value': "'other'"}}})
        self.nav_button('capacity', 'back', 'Back')
        self.nav_button('capacity', 'view', 'Bookmark', bookmark={'expr': {'Literal': {'Value': "'capacity'"}}})
        self.fault('definition/pages/capacity/page.json', lambda d: d.update({
            'pageBinding': {'type': 'Drillthrough', 'parameters': [{'boundFilter': 'drill'}]},
            'visibility': 'HiddenInViewMode',
            'filterConfig': {'filters': [{'name': 'drill'}, {'name': 'keep'}]},
            'visualInteractions': [{'source': 'nav', 'target': 'chart'}]}))
        self.fault('definition/bookmarks/capacity.bookmark.json', lambda d:
            d['options']['targetVisualNames'].extend(['nav', 'back']))
        before = gen.fingerprints(gen.snapshot(self.source))
        result = self.run_generation(write=True)
        report = self.output / 'SWI-licensing.Report'
        self.assertFalse((report / 'definition/pages/capacity/visuals/nav/visual.json').exists())
        self.assertFalse((report / 'definition/pages/capacity/visuals/back/visual.json').exists())
        self.assertEqual((report / 'definition/pages/capacity/visuals/view/visual.json').read_bytes(),
                         (self.source / 'definition/pages/capacity/visuals/view/visual.json').read_bytes())
        page = gen.decode((report / 'definition/pages/capacity/page.json').read_bytes())
        self.assertNotIn('pageBinding', page)
        self.assertNotIn('visibility', page)
        self.assertEqual(page['filterConfig']['filters'], [{'name': 'keep'}])
        self.assertEqual(page['visualInteractions'], [])
        bookmark = gen.decode((report / 'definition/bookmarks/capacity.bookmark.json').read_bytes())
        self.assertEqual(bookmark['options']['targetVisualNames'], ['chart'])
        self.assertEqual(before, gen.fingerprints(gen.snapshot(self.source)))
        self.assertEqual(before, result['masterFiles'])

    def test_tabs_keep_protected_buttons_byte_identical(self):
        self.tabs(protectedButtonPages=['capacity'])
        self.nav_button('capacity', 'protected-back', 'Back')
        self.nav_button('capacity', 'protected-nav', 'PageNavigation', navigationSection={'expr': {'Literal': {'Value': "'recovery'"}}})
        self.run_generation(write=True)
        for name in ['protected-back', 'protected-nav']:
            path = f'definition/pages/capacity/visuals/{name}/visual.json'
            self.assertEqual((self.output / 'SWI-licensing.Report' / path).read_bytes(), (self.source / path).read_bytes())

    def test_tabs_remove_bookmark_links_to_other_pages(self):
        self.tabs()
        self.nav_button('capacity', 'other-view', 'Bookmark', bookmark={'expr': {'Literal': {'Value': "'other'"}}})
        self.run_generation(write=True)
        self.assertFalse((self.output / 'SWI-licensing.Report/definition/pages/capacity/visuals/other-view/visual.json').exists())

    def test_tabs_do_not_mask_unknown_missing_visuals(self):
        self.tabs()
        self.fault('definition/bookmarks/capacity.bookmark.json', lambda d: d['options']['targetVisualNames'].append('unknown'))
        with self.assertRaisesRegex(gen.GenerationError, 'missing visuals'):
            self.run_generation()

    def orphan_navigator(self):
        self.tabs(orphanNavigatorVisuals=[{'page': 'capacity', 'visual': 'obsolete',
            'bookmarkGroup': 'retired-group', 'selectedBookmark': 'retired-bookmark'}])
        properties = {key: {'expr': {'Literal': {'Value': "'" + value + "'"}}}
                      for key, value in [('bookmarkGroup', 'retired-group'),
                                         ('selectedBookmark', 'retired-bookmark')]}
        self.put('definition/pages/capacity/visuals/obsolete/visual.json', {
            'name': 'obsolete', 'visual': {'visualType': 'bookmarkNavigator', 'objects': {
                'bookmarks': [{'properties': properties}]}}})

    def test_reviewed_orphan_navigator_prunes_only_its_references_and_preserves_source(self):
        self.orphan_navigator()
        self.fault('definition/pages/capacity/page.json', lambda d: d.update({
            'visualInteractions': [{'source': 'obsolete', 'target': 'chart'}]}))
        self.fault('definition/bookmarks/capacity.bookmark.json', lambda d: (
            d['options']['targetVisualNames'].append('obsolete'),
            d['explorationState']['sections']['capacity']['visualContainers'].update(obsolete={})))
        before = gen.fingerprints(gen.snapshot(self.source))
        result = self.run_generation(write=True)
        root = self.output / 'SWI-licensing.Report'
        self.assertFalse((root / 'definition/pages/capacity/visuals/obsolete/visual.json').exists())
        bookmark = gen.decode((root / 'definition/bookmarks/capacity.bookmark.json').read_bytes())
        self.assertEqual(bookmark['options']['targetVisualNames'], ['chart'])
        self.assertEqual(set(bookmark['explorationState']['sections']['capacity']['visualContainers']), {'chart'})
        self.assertEqual(gen.decode((root / 'definition/pages/capacity/page.json').read_bytes())['visualInteractions'], [])
        self.assertEqual(len(result['navigation']['removedOrphanNavigators']), 1)
        self.assertEqual(before, gen.fingerprints(gen.snapshot(self.source)))

    def test_orphan_adapter_preserves_a_working_navigator(self):
        self.orphan_navigator()
        self.put('definition/bookmarks/retired-bookmark.bookmark.json', {
            'name': 'retired-bookmark', 'explorationState': {'activeSection': 'capacity',
                'sections': {'capacity': {'visualContainers': {'chart': {}}}}}})
        self.fault('definition/bookmarks/bookmarks.json', lambda d: d['items'].append({
            'name': 'retired-group', 'children': ['retired-bookmark']}))
        result = self.run_generation(write=True)
        path = 'definition/pages/capacity/visuals/obsolete/visual.json'
        self.assertEqual((self.output / 'SWI-licensing.Report' / path).read_bytes(), (self.source / path).read_bytes())
        self.assertEqual(result['navigation']['removedOrphanNavigators'], [])

    def test_orphan_adapter_rejects_partial_resolution_and_binding_drift(self):
        self.orphan_navigator()
        self.fault('definition/bookmarks/bookmarks.json', lambda d: d['items'].append({
            'name': 'retired-group', 'children': []}))
        with self.assertRaisesRegex(gen.GenerationError, 'Partially resolved'):
            self.run_generation(write=True)
        self.assertFalse(self.output.exists())
        self.fault('definition/pages/capacity/visuals/obsolete/visual.json', lambda d:
            d['visual']['objects']['bookmarks'][0]['properties'].update(
                bookmarkGroup={'expr': {'Literal': {'Value': "'different-group'"}}}))
        with self.assertRaisesRegex(gen.GenerationError, 'bindings changed'):
            self.run_generation()

    def test_orphan_adapter_cannot_remove_another_visual_type(self):
        self.orphan_navigator()
        self.fault('definition/pages/capacity/visuals/obsolete/visual.json', lambda d:
            d['visual'].update(visualType='tableEx'))
        with self.assertRaisesRegex(gen.GenerationError, 'must target a bookmark navigator'):
            self.run_generation()

    def test_unlisted_orphan_group_is_still_rejected(self):
        self.orphan_navigator()
        self.tabs()  # No reviewed removal policy.
        with self.assertRaisesRegex(gen.GenerationError, 'Missing or external bookmark navigator group'):
            self.run_generation()

    def test_tabs_reject_hidden_support_pages(self):
        self.configure_support()
        self.tabs()
        with self.assertRaisesRegex(gen.GenerationError, 'hidden supporting'):
            self.run_generation()

    def test_tabs_add_explorer_picker_without_reusing_template_filter(self):
        field = {'Column': {'Expression': {'SourceRef': {'Entity': 'Inventory'}}, 'Property': 'Name'}}
        self.put('definition/pages/capacity/visuals/picker/visual.json', {
            'name': 'picker', 'position': {'x': 100, 'y': 100, 'width': 200, 'height': 80},
            'filterConfig': {'filters': [{'name': 'old'}]},
            'visual': {'visualType': 'slicer', 'objects': {'general': [{'properties': {'filter': {'old': True}}}]},
                'query': {'queryState': {'Values': {'projections': [{'field': field}]}}}}})
        self.tabs(explorerSelectors=[{'page': 'capacity', 'entity': 'Inventory', 'column': 'ID',
            'templateVisual': 'picker', 'position': {'x': 320, 'y': 100, 'width': 200, 'height': 80}}])
        result = self.run_generation(write=True)
        added = result['navigation']['addedExplorerSelectors'][0]
        picker = gen.decode((self.output / 'SWI-licensing.Report' / f"definition/pages/capacity/visuals/{added['visual']}/visual.json").read_bytes())
        self.assertNotIn('filterConfig', picker)
        self.assertNotIn('filter', picker['visual']['objects']['general'][0]['properties'])
        self.assertEqual(picker['visual']['query']['queryState']['Values']['projections'][0]['field']['Column']['Property'], 'ID')

    def layout_fixture(self):
        self.fault('definition/pages/capacity/page.json', lambda d: d.update(width=1920, height=1080))
        self.put('definition/pages/capacity/visuals/selection/visual.json', {
            'name': 'selection', 'position': {'x': 32, 'y': 392, 'width': 1856, 'height': 80, 'z': 12000},
            'visual': {'visualType': 'slicer', 'query': {'keep': 'original'}, 'objects': {'keep': 'style'}}})
        self.nav_button('capacity', 'protected-nav', 'Back')
        self.fault('definition/pages/capacity/visuals/protected-nav/visual.json', lambda d:
            d.update(position={'x': 32, 'y': 386.9, 'width': 1856, 'height': 37.5}))
        self.tabs(protectedButtonPages=['capacity'], layoutAdjustments=[{'page': 'capacity', 'visual': 'selection',
            'expectedPosition': {'x': 32, 'y': 392, 'width': 1856, 'height': 80}, 'position': {'y': 432}}])

    def test_layout_fix_preserves_buttons_filters_and_source(self):
        self.layout_fixture()
        before = gen.fingerprints(gen.snapshot(self.source))
        result = self.run_generation(write=True)
        folder = self.output / 'SWI-licensing.Report/definition/pages/capacity/visuals'
        selector = gen.decode((folder / 'selection/visual.json').read_bytes())
        original = self.read('definition/pages/capacity/visuals/selection/visual.json')
        self.assertEqual(selector['position']['y'], 432)
        selector['position']['y'] = 392
        self.assertEqual(selector, original)
        self.assertEqual((folder / 'protected-nav/visual.json').read_bytes(),
                         (self.source / 'definition/pages/capacity/visuals/protected-nav/visual.json').read_bytes())
        self.assertEqual(len(result['navigation']['layoutAdjustments']), 1)
        self.assertEqual(before, gen.fingerprints(gen.snapshot(self.source)))

    def test_layout_fix_refuses_changed_geometry_or_new_collision(self):
        self.layout_fixture()
        self.fault('definition/pages/capacity/visuals/selection/visual.json', lambda d: d['position'].update(y=400))
        with self.assertRaisesRegex(gen.GenerationError, 'geometry changed'):
            self.run_generation()
        self.fault('definition/pages/capacity/visuals/selection/visual.json', lambda d: d['position'].update(y=392))
        self.put('definition/pages/capacity/visuals/obstacle/visual.json', {
            'name': 'obstacle', 'position': {'x': 32, 'y': 500, 'width': 1856, 'height': 40},
            'visual': {'visualType': 'tableEx'}})
        with self.assertRaisesRegex(gen.GenerationError, 'overlaps'):
            self.run_generation()

    def test_symlink_source_is_refused(self):
        link = self.root / 'Linked.Report'
        try:
            link.symlink_to(self.source, target_is_directory=True)
        except OSError:
            self.skipTest('Symlink creation unavailable')
        with self.assertRaisesRegex(gen.GenerationError, 'Linked'):
            gen.generate(link, self.manifest, 'licensing', self.output)


if __name__ == '__main__':
    unittest.main()
