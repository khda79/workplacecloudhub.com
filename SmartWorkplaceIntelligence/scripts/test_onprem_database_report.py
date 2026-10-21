"""Offline navigation and semantic contract for the database views; no source IO."""
import hashlib
import json
import os
from pathlib import Path
import unittest

ROOT=Path(os.environ.get('SMARTWORKPLACE_REPORT') or Path(__file__).resolve().parents[1]/'pbip/SmartWorkplaceIntelligence.Report')/'definition'
PID='on-prem-messaging'
VIEWS=['Overview','Servers & compute','Storage','Databases','Readiness','Collection & details']
def uid(s):return hashlib.sha256((PID+'|'+s).encode()).hexdigest()[:20]
def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))

class DatabaseReportTests(unittest.TestCase):
    def test_bookmarks_preserve_filters_and_select_one_view(self):
        groups={uid('group-'+n) for n in VIEWS}
        for view in VIEWS:
            b=read(ROOT/'bookmarks'/(uid('bookmark-'+view)+'.bookmark.json'))
            self.assertTrue(b['options']['suppressData'])
            self.assertEqual(set(b['options']['targetVisualNames']),groups)
            states=b['explorationState']['sections'][PID]['visualContainerGroups']
            self.assertEqual([g for g,s in states.items() if not s['isHidden']],[uid('group-'+view)])

    def test_kpis_and_database_detail_bindings(self):
        root=ROOT/'pages'/PID/'visuals'
        k=read(root/'e9848fab86e52edadf2b/visual.json')
        measures=[p['field']['Measure']['Property'] for p in k['visual']['query']['queryState']['Data']['projections']]
        self.assertEqual(len(measures),9)
        self.assertNotIn('Exchange Infrastructure Captured Sources',measures)
        self.assertIn('Exchange Legacy Server Display',measures)
        self.assertIn('Exchange Observed Databases',measures)
        v=read(root/uid('database-details')/'visual.json')
        columns=[p['field']['Column']['Property'] for p in v['visual']['query']['queryState']['Values']['projections']]
        self.assertEqual(columns,['Database','Server Name','Database Status','Observed Mailboxes','Database Size GB','Reusable Internal Space GB','Data Notes'])

    def test_four_database_charts_and_removed_coverage_tables(self):
        root=ROOT/'pages'/PID/'visuals'
        for tag in ['db-mailboxes','db-sizes','db-status','db-server']:
            v=read(root/uid(tag)/'visual.json')
            self.assertEqual(v['visual']['visualType'],'barChart')
            self.assertEqual(v['visual']['query']['queryState']['Category']['projections'][0]['field']['Column']['Expression']['SourceRef']['Entity'],'Exchange Mailbox Databases')
            self.assertIn('Is Database',[f['field']['Column']['Property'] for f in v['filterConfig']['filters']])
        for old in ['bfbe9d175e2bf7e878fa','06bd5ae21628c075ad6d']:
            self.assertFalse((root/old).exists())

    def test_visible_visuals_fit_canvas_and_slicer_has_space(self):
        root=ROOT/'pages'/PID/'visuals'
        for p in root.glob('*/visual.json'):
            v=read(p);pos=v['position']
            self.assertGreaterEqual(pos['x'],0)
            self.assertLessEqual(pos['x']+pos['width'],1920)
            self.assertLessEqual(pos['y']+pos['height'],1080)
        s=read(root/uid('database-server')/'visual.json')
        self.assertGreaterEqual(s['position']['height'],76)

if __name__=='__main__':unittest.main()
