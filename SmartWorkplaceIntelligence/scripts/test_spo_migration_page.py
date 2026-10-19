"""Static PBIR acceptance checks. These do not qualify Desktop rendering."""
import json
import os
from pathlib import Path
import unittest
import migration_model_definitions as model

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
        cls.defs = model.definitions()

    def test_one_page_at_the_approved_position(self):
        order = read(REPORT / 'pages/pages.json')['pageOrder']
        self.assertEqual(len(order), 29)
        self.assertEqual(order.count(PAGE), 1)
        anchor = next(p for p in order if read(REPORT / 'pages' / p / 'page.json')['displayName'] == 'On-prem Content & Storage')
        self.assertEqual(order[order.index(anchor) + 1], PAGE)
        self.assertEqual((self.page['width'], self.page['height']), (1920, 1080))
        self.assertEqual([v['isHidden'] for v in self.top], [False, True, True, True])

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
                self.assertEqual(state['visualContainers'], {})
                values = list(state['visualContainerGroups'].values())
                self.assertEqual([x['isHidden'] for x in values], [n != number for n in range(len(values))])
                self.assertEqual(set(state['visualContainerGroups']), set(mark['options']['targetVisualNames']))

    def test_filters_are_shared_and_clear_keeps_the_active_view(self):
        slicers = [v for v in self.visuals if v.get('visual', {}).get('visualType') == 'slicer']
        self.assertEqual(len(slicers), 4)
        self.assertTrue(all(not s.get('parentGroupName') for s in slicers))
        self.assertTrue(all(s['position']['height'] >= 80 for s in slicers))
        button = next(v for v in self.visuals if v.get('visual', {}).get('visualType') == 'actionButton')
        self.assertEqual(button['visual']['visualContainerObjects']['visualLink'][0]['properties']['type']['expr']['Literal']['Value'], "'ClearAllSlicers'")
        self.assertGreaterEqual(len(self.page['filterConfig']['filters']), 10)

    def test_visible_geometry_in_every_tab_and_schedule_mode(self):
        for top in self.top:
            nested = [g for g in self.groups.values() if g.get('parentGroupName') == top['name']]
            for mode in nested or [None]:
                selected = {top['name']} | ({mode['name']} if mode else set())
                visible = [v for v in self.visuals if 'visual' in v and (not v.get('parentGroupName') or v['parentGroupName'] in selected)]
                for i, v in enumerate(visible):
                    p = v['position']
                    self.assertGreaterEqual(p['x'], 0)
                    self.assertGreaterEqual(p['y'], 0)
                    self.assertLessEqual(p['x'] + p['width'], 1920)
                    self.assertLessEqual(p['y'] + p['height'], 1080)
                    for other in visible[i + 1:]:
                        self.assertFalse(overlaps(p, other['position']), (top['visualGroup']['displayName'], v['name'], other['name']))

    def test_field_bindings_exist_and_record_tables_do_not_mix_measures(self):
        tables = {t['Name']: {c['Name'] for c in t['Columns']} for t in self.defs['tables']}
        measures = {(m['tableName'], m['name']) for m in self.defs['measures']}
        for v in self.visuals:
            visual = v.get('visual', {})
            for role in visual.get('query', {}).get('queryState', {}).values():
                for projection in role['projections']:
                    self.assertTrue(projection['nativeQueryRef'])
                    field = projection['field']
                    kind = next(iter(field))
                    value = field[kind]
                    table, prop = value['Expression']['SourceRef']['Entity'], value['Property']
                    if kind == 'Measure':
                        self.assertIn((table, prop), measures)
                    elif table in tables:
                        self.assertIn(prop, tables[table])
            if visual.get('visualType') == 'tableEx':
                fields = visual['query']['queryState']['Values']['projections']
                self.assertTrue(all('Column' in f['field'] for f in fields))
                self.assertGreaterEqual(v['position']['height'], 300)
                self.assertEqual(visual['objects']['columnHeaders'][0]['properties']['columnAdjustment']['expr']['Literal']['Value'], "'growToFit'")
                self.assertEqual(visual['objects']['values'][0]['properties']['wordWrap']['expr']['Literal']['Value'], 'false')

    def test_history_does_not_connect_unmeasured_days(self):
        histories = [v['visual'] for v in self.visuals if 'SPO Historical Reported Done' in json.dumps(v)]
        self.assertEqual(len(histories), 1)
        properties = histories[0]['objects']['lineStyles'][0]['properties']
        self.assertEqual(properties['strokeShow']['expr']['Literal']['Value'], 'false')
        self.assertEqual(properties['showMarker']['expr']['Literal']['Value'], 'true')
        self.assertIn('Historical Migrate Operations', json.dumps(histories[0]))


if __name__ == '__main__':
    unittest.main()
