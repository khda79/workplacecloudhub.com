"""Offline checks for daily trends navigation, filtering, bindings and geometry."""
import json
import os
from pathlib import Path
import re
import unittest

PBIP = Path(os.environ.get('SMARTWORKPLACE_PBIP') or (Path(__file__).resolve().parents[1] / 'pbip'))
REPORT = PBIP / 'SmartWorkplaceIntelligence.Report' / 'definition'
MODEL = PBIP / 'SmartWorkplaceIntelligence.SemanticModel' / 'definition'
NAMES = ['User accounts', 'AD computer accounts', 'Computer accounts', 'Mailboxes', 'Domain detail']
BOOKMARK_NAMES = ['User accounts', 'AD computer accounts', 'Windows AD accounts', 'Mailboxes', 'Domain detail']


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def overlaps(a, b):
    return (a['x'] < b['x'] + b['width'] and b['x'] < a['x'] + a['width'] and
            a['y'] < b['y'] + b['height'] and b['y'] < a['y'] + a['height'])


class DailyTrendsLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.order = read(REPORT / 'pages/pages.json')['pageOrder']
        cls.all_pages = [read(REPORT / 'pages' / page / 'page.json') for page in cls.order]
        cls.page = next(page for page in cls.all_pages if page['name'] == 'daily-directory-mailbox-trends')
        cls.containers = [read(file) for file in sorted((REPORT / 'pages' / cls.page['name'] / 'visuals').glob('*/visual.json'))]
        cls.pages = [next(v for v in cls.containers if v.get('visualGroup', {}).get('displayName') == name) for name in NAMES]
        cls.shared = [v for v in cls.containers if 'visual' in v and not v.get('parentGroupName')]
        cls.visuals = {g['name']: [v for v in cls.containers if v.get('parentGroupName') == g['name']] for g in cls.pages}

    def test_navigation_and_canvas(self):
        self.assertEqual(len(self.order), 30)
        anchor = next(p['name'] for p in self.all_pages if p['displayName'] == 'Operational Overview')
        start = self.order.index(anchor) + 1
        self.assertEqual(self.order[start], self.page['name'])
        self.assertEqual(self.page['displayName'], 'Directory & Mailbox Daily Trends')
        self.assertEqual((self.page['width'], self.page['height'], self.page['displayOption']), (1920, 1080, 'FitToPage'))
        self.assertEqual([g.get('isHidden', False) for g in self.pages], [False, True, True, True, True])
        nav = next(v['visual'] for v in self.shared if v['visual']['visualType'] == 'bookmarkNavigator')
        layout = nav['objects']['layout'][0]['properties']
        self.assertEqual(layout['columnCount']['expr']['Literal']['Value'], '5D')
        self.assertEqual(layout['rowCount']['expr']['Literal']['Value'], '1D')
        self.assertLessEqual(float(layout['cellPadding']['expr']['Literal']['Value'].rstrip('D')), 4)
        group_name = nav['objects']['bookmarks'][0]['properties']['bookmarkGroup']['expr']['Literal']['Value'].strip("'")
        index = read(REPORT / 'bookmarks/bookmarks.json')
        group = next(g for g in index['items'] if g['name'] == group_name)
        self.assertEqual(len(group['children']), 5)
        for number, name in enumerate(group['children']):
            mark = read(REPORT / 'bookmarks' / (name + '.bookmark.json'))
            self.assertEqual(mark['displayName'], BOOKMARK_NAMES[number])
            self.assertTrue(mark['options']['suppressData'], 'Tabs must preserve slicer selections')
            self.assertTrue(mark['options']['suppressActiveSection'])
            self.assertTrue(mark['options']['applyOnlyToTargetVisuals'])
            self.assertEqual(set(mark['options']['targetVisualNames']), {g['name'] for g in self.pages})
            state = mark['explorationState']['sections'][self.page['name']]
            self.assertEqual(state['visualContainers'], {}, 'Bookmarks must not store filter/query state')
            self.assertEqual([state['visualContainerGroups'][g['name']]['isHidden'] for g in self.pages], [i != number for i in range(5)])

    def test_geometry_has_no_overlap_or_clipping(self):
        for page in self.pages:
            visuals = self.shared + self.visuals[page['name']]
            for number, visual in enumerate(visuals):
                a = visual['position']
                self.assertGreaterEqual(a['x'], 0)
                self.assertGreaterEqual(a['y'], 0)
                self.assertLessEqual(a['x'] + a['width'], self.page['width'])
                self.assertLessEqual(a['y'] + a['height'], self.page['height'])
                for other in visuals[number + 1:]:
                    self.assertFalse(overlaps(a, other['position']), (page['name'], visual['name'], other['name']))

    def test_one_native_chart_per_subject_and_matrix_only_in_details(self):
        for number, page in enumerate(self.pages):
            types = [v['visual']['visualType'] for v in self.visuals[page['name']]]
            self.assertEqual(types.count('lineChart'), 1 if number < 4 else 0)
            self.assertEqual(types.count('pivotTable'), 2 if number == 4 else 0)
            self.assertEqual(types.count('tableEx'), 1 if number == 4 else 0)

    def test_common_filters_are_shared_once_and_not_in_bookmarks(self):
        slicers = [v for v in self.shared if v['visual']['visualType'] == 'slicer']
        self.assertEqual(len(slicers), 4)
        for table in ['Daily Period', 'Daily Domain', 'Date', 'Country Selection']:
            self.assertEqual(sum(table in json.dumps(v['visual']['query']) for v in slicers), 1)
        for slicer in slicers:
            query = json.dumps(slicer['visual']['query'])
            expected = ('SharedCountry' if 'Country Selection' in query else
                        'SharedADDomain' if 'Daily Domain' in query else None)
            sync = slicer['visual'].get('syncGroup')
            if expected:
                self.assertEqual(sync, {'groupName': expected, 'fieldChanges': True, 'filterChanges': True})
            else:
                self.assertIsNone(sync, 'Period and Date must remain local to the daily page')
        extra = [v for g in self.pages for v in self.visuals[g['name']] if v['visual']['visualType'] == 'slicer']
        self.assertEqual(len(extra), 1)
        self.assertIn('Daily View', json.dumps(extra[0]))

    def test_chart_grain_gaps_and_reading_tooltips(self):
        expected = [('Axis Enabled Users', 'Axis Disabled Users', 'Axis AD Users Reading Day'),
                    ('Axis Total AD Computer Accounts', 'Axis Enabled AD Computer Accounts', 'Axis Disabled AD Computer Accounts', 'Axis All AD Computer Accounts Reading Day'),
                    ('Axis Windows 10 All AD Accounts Chart', 'Axis Windows 11 All AD Accounts Chart', 'Axis Windows All AD Accounts Reading Day'),
                    ('Axis On-premises Mailboxes', 'Axis Exchange Online Remote', 'Axis Exchange Online Share', 'Axis Mailbox Reading Day')]
        for page, required in zip(self.pages, expected):
            container = next(v for v in self.visuals[page['name']] if v['visual']['visualType'] == 'lineChart')
            chart = container['visual']
            value = json.dumps(container)
            for name in (*required, 'Axis Date', 'Axis Grain Subtitle', 'Axis Point Visible'):
                self.assertIn(name, value)
            self.assertEqual(chart['objects']['categoryAxis'][0]['properties']['axisType']['expr']['Literal']['Value'], "'Categorical'")
            self.assertIn('Quality Tooltip', value)
        mailbox = next(v['visual'] for v in self.visuals[self.pages[3]['name']] if v['visual']['visualType'] == 'lineChart')
        self.assertIn('y2Axis', mailbox['objects'])

    def test_detail_values_and_changes_are_separate(self):
        matrix = next(v['visual'] for v in self.visuals[self.pages[4]['name']] if v['name'] == 'be8bc7cefef068d1abd0')
        values = matrix['query']['queryState']['Values']['projections']
        self.assertEqual(len(values), 12)
        self.assertEqual(len({p['nativeQueryRef'] for p in values}), 12, 'Matrix aliases must be unique')
        self.assertEqual(len({p['displayName'] for p in values}), 12, 'Matrix headers must identify each value/change')
        self.assertEqual(sum(p['field']['Measure']['Property'].endswith('Value Text') for p in values), 6)
        self.assertEqual(sum(p['field']['Measure']['Property'].endswith('Change Text') for p in values), 6)

    def test_display_helpers_only_extract_existing_card_text(self):
        for table, count in [('Daily Directory Trends', 12), ('Daily Mailbox Trends', 4)]:
            text = (MODEL / 'tables' / (table + '.tmdl')).read_text(encoding='utf-8-sig')
            helpers = re.findall(r"^\tmeasure '([^']+(?:Value|Change) Text)' = (.+)$", text, re.M)
            self.assertEqual(len(helpers), count)
            for _, expression in helpers:
                self.assertRegex(expression, r'^VAR _text = \[[^\]]+ Card\] VAR _separator = SEARCH')
                self.assertNotRegex(expression, r'CALCULATE|SUMX|FILTER|ALLSELECTED')

    def test_planning_snapshot_and_history_are_explicit(self):
        mailbox = self.visuals[self.pages[3]['name']]
        cards = [v for v in mailbox if v['visual']['visualType'] == 'cardVisual' and v['position']['y'] == 312]
        self.assertEqual(len(cards), 5)
        self.assertEqual(sorted(v['position']['x'] for v in cards), [32, 408, 784, 1160, 1536])
        for name in ['Retained on-premises', 'Remaining in migration scope', 'Exceptions requiring review', 'Planning snapshot', 'Observed history']:
            self.assertIn(name, json.dumps(mailbox))
        snapshot = (MODEL / 'tables/Mailbox Migration Exception Summary.tmdl').read_text(encoding='utf-8-sig')
        self.assertNotRegex(snapshot, r"'Date'\[|'Daily Period'\[")
        self.assertIn('Pending review', snapshot)
        self.assertIn('Global Review Count', snapshot)


if __name__ == '__main__':
    unittest.main()
