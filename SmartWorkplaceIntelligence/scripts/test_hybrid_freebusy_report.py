"""Offline navigation and binding contract for the Messaging & Hybrid Free/Busy view; no source IO."""
import hashlib
import json
import os
from pathlib import Path
import unittest

ROOT = Path(os.environ.get('SMARTWORKPLACE_REPORT') or Path(__file__).resolve().parents[1] / 'pbip/SmartWorkplaceIntelligence.Report') / 'definition'
PID = '71a60a59e0eb52a7a7d0'
# Existing Messaging & Hybrid views in tab order: (tab visual, group, bookmark).
VIEWS = [('b76d39eb2e0d1857e671', '356f38365fae5fd56ff9', 'ce9b78a7cffd6236713e'),
         ('ca1cf7e7609e4c5d0707', '421d3be4994300ee3abf', '7a63252f7b186ab2b3a7'),
         ('2a6e59a203c50bbaa9ed', '44b410e354eea9b64e0c', '2c93bbad604972e3b7ce'),
         ('9e9b3db02c389e850393', '509da9ee86e35c2c4711', '5bb21234305c3967ada6'),
         ('a88f54fe3a8365079e53', 'daaf5c483406b36ac30c', '1490c66d05a332d9ed51')]


def uid(tag):
    return hashlib.sha256((PID + '|freebusy-' + tag).encode()).hexdigest()[:20]


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


NEW = (uid('tab'), uid('group'), uid('bookmark'))
ALL = VIEWS + [NEW]


class HybridFreeBusyReportTests(unittest.TestCase):
    def test_bookmarks_select_exactly_one_view(self):
        groups = {view[1] for view in ALL}
        for _, group, bookmark in ALL:
            b = read(ROOT / 'bookmarks' / (bookmark + '.bookmark.json'))
            self.assertTrue(b['options']['suppressData'])
            self.assertEqual(set(b['options']['targetVisualNames']), groups)
            states = b['explorationState']['sections'][PID]['visualContainerGroups']
            self.assertEqual([g for g, s in states.items() if not s['isHidden']], [group])
        children = [i for i in read(ROOT / 'bookmarks' / 'bookmarks.json')['items'] if i['displayName'] == 'Messaging & Hybrid views'][0]['children']
        self.assertEqual(children, [view[2] for view in ALL])

    def test_six_tabs_fit_and_link_to_their_bookmark(self):
        root = ROOT / 'pages' / PID / 'visuals'
        tabs = [read(root / view[0] / 'visual.json') for view in ALL]
        for a, b in zip(tabs, tabs[1:]):
            self.assertLess(a['position']['x'] + a['position']['width'], b['position']['x'])
        self.assertLessEqual(tabs[-1]['position']['x'] + tabs[-1]['position']['width'], 1888)
        for tab, (_, _, bookmark) in zip(tabs, ALL):
            link = tab['visual']['visualContainerObjects']['visualLink'][0]['properties']['bookmark']['expr']['Literal']['Value']
            self.assertEqual(link, "'" + bookmark + "'")

    def test_free_busy_bindings(self):
        root = ROOT / 'pages' / PID / 'visuals'
        k = read(root / uid('kpis') / 'visual.json')
        measures = [p['field']['Measure']['Property'] for p in k['visual']['query']['queryState']['Data']['projections']]
        self.assertEqual(measures, ['Exchange Free Busy Status', 'Exchange Free Busy Checks Completed', 'Exchange Free Busy Checks Total',
                                    'Exchange Free Busy Errors', 'Exchange Free Busy Warnings', 'Exchange Free Busy Not Authorized Checks'])
        summary = read(root / uid('summary') / 'visual.json')
        columns = [p['field']['Column']['Property'] for p in summary['visual']['query']['queryState']['Values']['projections']]
        self.assertEqual(columns[:3], ['Side', 'Auth', 'Side Status'])
        self.assertIn('Not Authorized Cmdlets', columns)
        self.assertIn('Missing Roles', columns)
        findings = read(root / uid('findings') / 'visual.json')
        self.assertEqual(findings['visual']['query']['queryState']['Values']['projections'][0]['field']['Column']['Expression']['SourceRef']['Entity'], 'Exchange Hybrid Free Busy Findings')
        values = findings['filterConfig']['filters'][0]['filter']['Where'][0]['Condition']['In']['Values']
        self.assertEqual([v[0]['Literal']['Value'] for v in values], ["'Error'", "'Warning'", "'NotAuthorized'"])
        for tag in ['label', 'kpis', 'summary', 'findings']:
            v = read(root / uid(tag) / 'visual.json')
            self.assertEqual(v['parentGroupName'], NEW[1])
            pos = v['position']
            self.assertLessEqual(pos['y'] + pos['height'], 1000)


if __name__ == '__main__':
    unittest.main()
