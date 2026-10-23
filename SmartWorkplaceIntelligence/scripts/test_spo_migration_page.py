"""Static PBIR acceptance checks. These do not qualify Desktop rendering."""
import json
import os
from pathlib import Path
import unittest
from test_model_dax_references import parse_model

ROOT = Path(os.environ.get('SMARTWORKPLACE_PBIP') or Path(__file__).resolve().parents[1] / 'pbip')
REPORT = ROOT / 'SmartWorkplaceIntelligence.Report/definition'
PAGE = 'migration-to-sharepoint-online'


def read(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))


def overlaps(a, b):
    return a['x'] < b['x'] + b['width'] and b['x'] < a['x'] + a['width'] and a['y'] < b['y'] + b['height'] and b['y'] < a['y'] + a['height']


class MigrationPageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.page = read(REPORT / 'pages' / PAGE / 'page.json')
        cls.visuals = [read(p) for p in (REPORT / 'pages' / PAGE / 'visuals').glob('*/visual.json')]
        cls.groups = {v['name']: v for v in cls.visuals if 'visualGroup' in v}
        cls.top = [v for v in cls.groups.values() if not v.get('parentGroupName')]
        cls.top.sort(key=lambda v: ['Progress', 'Schedule', 'Reconciliation', 'Quality'].index(v['visualGroup']['displayName']))
        cls.tables, _ = parse_model(ROOT / 'SmartWorkplaceIntelligence.SemanticModel/definition')

    def test_one_page_at_the_approved_position(self):
        order = read(REPORT / 'pages/pages.json')['pageOrder']
        self.assertEqual(len(order), 30)
        self.assertEqual(len(order), len(set(order)))
        self.assertEqual(order.count(PAGE), 1)
        anchor = next(p for p in order if read(REPORT / 'pages' / p / 'page.json')['displayName'] == 'Backup & Resilience')
        self.assertEqual(order[order.index(anchor) + 1:order.index(anchor) + 4],
                         [PAGE, '7b020122c7890360dc36', '736753e8cf8c034002ef'])
        self.assertEqual((self.page['width'], self.page['height']), (1920, 1280))
        self.assertEqual(self.page['displayOption'], 'FitToWidth')
        self.assertEqual([v.get('isHidden', False) for v in self.top], [False, True, True, True])

    def test_bookmarks_only_change_visibility_not_data(self):
        index = read(REPORT / 'bookmarks/bookmarks.json')['items']
        navs = [v for v in self.visuals if v.get('visual', {}).get('visualType') == 'bookmarkNavigator']
        self.assertEqual(len(navs), 2)
        for nav in navs:
            group_id = nav['visual']['objects']['bookmarks'][0]['properties']['bookmarkGroup']['expr']['Literal']['Value'].strip("'")
            group = next(g for g in index if g['name'] == group_id)
            for number, child in enumerate(group['children']):
                mark = read(REPORT / 'bookmarks' / (child + '.bookmark.json'))
                self.assertTrue(mark['options']['suppressData'])
                self.assertTrue(mark['options']['suppressActiveSection'])
                state = mark['explorationState']['sections'][PAGE]
                if nav['name'] == '639bcc87cc19419112e9':
                    ids = ['3da1f56d55390e9ee5f5', '2ef3803af6af59f2a42f']
                    self.assertEqual(state.get('visualContainerGroups', {}), {})
                    self.assertEqual(mark['options']['targetVisualNames'], ids)
                    self.assertFalse(mark['options']['applyOnlyToTargetVisuals'])
                    self.assertFalse(mark['options']['suppressDisplay'])
                    self.assertEqual(state['visualContainers'], {
                        name: {'singleVisual': {} if n == number else {'display': {'mode': 'hidden'}}}
                        for n, name in enumerate(ids)})
                    visuals = {v['name']: v for v in self.visuals}
                    self.assertTrue(all(visuals[name]['parentGroupName'] == 'b2deaadb9550a5b30275' for name in ids))
                    self.assertEqual([visuals[name].get('isHidden', False) for name in ids], [False, True])
                    self.assertEqual(nav['visual']['objects']['bookmarks'][0]['properties']['selectedBookmark']['expr']['Literal']['Value'], "'8b7de51565202c517aff'")
                else:
                    self.assertEqual(state['visualContainers'], {})
                    groups = state['visualContainerGroups']
                    self.assertEqual(set(groups), set(mark['options']['targetVisualNames']))
                    self.assertTrue(all(set(value) == {'isHidden'} for value in groups.values()))
                    self.assertEqual([x['isHidden'] for x in groups.values()], [n != number for n in range(len(groups))])

    def test_filters_are_shared_and_clear_keeps_the_active_view(self):
        slicers = [v for v in self.visuals if v.get('visual', {}).get('visualType') == 'slicer']
        self.assertEqual(len(slicers), 4)
        fields = {(f['field']['Column']['Expression']['SourceRef']['Entity'], f['field']['Column']['Property'])
                  for v in slicers for f in v['visual']['query']['queryState']['Values']['projections']}
        self.assertEqual(fields, {('SPO Migration Project', 'Project Key'), ('SPO Migration Phase', 'Phase'),
                                  ('SPO Migration Business Unit', 'Business Unit'), ('SPO Migration Source Type', 'Source Type')})
        self.assertNotIn('a57964e4b1454a706c62', {v['name'] for v in self.visuals})
        self.assertTrue(all(not s.get('parentGroupName') for s in slicers))
        self.assertTrue(all(s['position']['height'] >= 80 for s in slicers))
        button = next(v for v in self.visuals if v['name'] == '93f15cca80d95412e991')
        self.assertEqual(button['visual']['visualContainerObjects']['visualLink'][0]['properties']['type']['expr']['Literal']['Value'], "'ClearAllSlicers'")
        self.assertGreaterEqual(len(self.page['filterConfig']['filters']), 10)

    def test_visible_geometry_in_every_tab_and_schedule_mode(self):
        for top in self.top:
            nested = [g for g in self.groups.values() if g.get('parentGroupName') == top['name']]
            for mode in nested or [None]:
                selected = {top['name']} | ({mode['name']} if mode else set())
                hidden = {}
                if top['name'] == 'b2deaadb9550a5b30275':
                    bookmark = '8b7de51565202c517aff' if mode['name'] == '3b7aebbfb8249fe2aa74' else '9772dab5c1271a299bbf'
                    state = read(REPORT / 'bookmarks' / (bookmark + '.bookmark.json'))['explorationState']['sections'][PAGE]
                    hidden = {name: value['singleVisual'].get('display', {}).get('mode') == 'hidden'
                              for name, value in state['visualContainers'].items()}
                visible = [v for v in self.visuals if 'visual' in v and not hidden.get(v['name'], v.get('isHidden', False))
                           and (not v.get('parentGroupName') or v['parentGroupName'] in selected)]
                panel = next(v for v in visible if v['name'] == 'migration-style-panel')
                self.assertEqual(panel['visual']['visualType'], 'shape')
                for i, v in enumerate(visible):
                    p = v['position']
                    self.assertGreaterEqual(p['x'], 0)
                    self.assertGreaterEqual(p['y'], 0)
                    self.assertLessEqual(p['x'] + p['width'], self.page['width'])
                    self.assertLessEqual(p['y'] + p['height'], self.page['height'])
                    for other in visible[i + 1:]:
                        if panel['name'] in (v['name'], other['name']):
                            foreground = other if v is panel else v
                            if overlaps(p, other['position']):
                                # Child z is relative to its enclosing group(s).
                                effective_z = foreground['position']['z']
                                parent = foreground.get('parentGroupName')
                                while parent:
                                    group = self.groups[parent]
                                    effective_z += group['position']['z']
                                    parent = group.get('parentGroupName')
                                self.assertLess(panel['position']['z'], effective_z)
                            continue
                        self.assertFalse(overlaps(p, other['position']), (top['visualGroup']['displayName'], v['name'], other['name']))

    def test_field_bindings_exist_and_record_tables_do_not_mix_measures(self):
        tables = {name: item['columns'] for name, item in self.tables.items()}
        measures = {(name, measure) for name, item in self.tables.items() for measure in item['measures']}
        for v in self.visuals:
            visual = v.get('visual', {})
            for role in visual.get('query', {}).get('queryState', {}).values():
                for projection in role['projections']:
                    self.assertTrue(projection['nativeQueryRef'])
                    field = projection['field']
                    kind = next(iter(field))
                    value = field[kind]
                    table, prop = value['Expression']['SourceRef']['Entity'], value['Property']
                    self.assertIn(table, tables)
                    self.assertIn(kind, ('Measure', 'Column'))
                    if kind == 'Measure':
                        self.assertIn((table, prop), measures)
                    else:
                        self.assertIn(prop, tables[table])
            if visual.get('visualType') == 'tableEx':
                fields = visual['query']['queryState']['Values']['projections']
                self.assertTrue(all('Column' in f['field'] for f in fields))
                self.assertEqual(visual['objects']['columnHeaders'][0]['properties']['columnAdjustment']['expr']['Literal']['Value'], "'growToFit'")
                self.assertEqual(visual['objects']['values'][0]['properties']['wordWrap']['expr']['Literal']['Value'], 'false')

    def test_record_table_readability_minimum(self):
        # Keep the approved minimum. Desktop readability has not been qualified.
        for v in self.visuals:
            if v.get('visual', {}).get('visualType') == 'tableEx':
                with self.subTest(visual=v['name']):
                    self.assertGreaterEqual(v['position']['height'], 300)

    def test_groups_and_bottom_controls_fit_the_taller_canvas(self):
        for group in self.groups.values():
            p = group['position']
            self.assertEqual((p['x'], p['y'], p['width'], p['height']), (0, 0, self.page['width'], self.page['height']))
        visuals = {v['name']: v for v in self.visuals}
        tables = [v for v in self.visuals if v.get('visual', {}).get('visualType') == 'tableEx']
        bottom = max(v['position']['y'] + v['position']['height'] for v in tables)
        panel = visuals['migration-style-panel']['position']
        self.assertLess(bottom, panel['y'] + panel['height'])
        for name in ['93f15cca80d95412e991', '2756da8eb80be4bb858f']:
            p = visuals[name]['position']
            self.assertGreater(p['y'], panel['y'] + panel['height'])
            self.assertLessEqual(p['y'] + p['height'], self.page['height'])
        lower_tables = [v for v in tables if v['name'] != '7777ba38487007b4d9e7']
        for name in ['083edd8faa940f3dd4e5', '615330edde0979799652', '9a79746e4fcbe26e6705']:
            note = visuals[name]
            below = next(v for v in lower_tables if v['parentGroupName'] == note['parentGroupName'])
            self.assertLess(note['position']['y'] + note['position']['height'], below['position']['y'])

    def test_history_does_not_connect_unmeasured_days(self):
        histories = [v['visual'] for v in self.visuals if 'SPO Historical Reported Done' in json.dumps(v)]
        self.assertEqual(len(histories), 1)
        properties = histories[0]['objects']['lineStyles'][0]['properties']
        self.assertEqual(properties['strokeShow']['expr']['Literal']['Value'], 'false')
        self.assertEqual(properties['showMarker']['expr']['Literal']['Value'], 'true')
        self.assertIn('Historical Migrate Operations', json.dumps(histories[0]))


if __name__ == '__main__':
    unittest.main()
