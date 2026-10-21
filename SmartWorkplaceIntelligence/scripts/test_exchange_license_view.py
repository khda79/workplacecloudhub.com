"""Version 1.1.0: retained populations and emphasized estimated purchases."""
import hashlib,json,os
from pathlib import Path
import unittest

ROOT=Path(os.environ.get('SMARTWORKPLACE_REPORT') or Path(__file__).resolve().parents[1]/'pbip/SmartWorkplaceIntelligence.Report')/'definition'
TARGET='736753e8cf8c034002ef'
def uid(s):return hashlib.sha256(('exchange-licensing|'+s).encode()).hexdigest()[:20]
def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))

class LicenseViewTests(unittest.TestCase):
    def test_old_page_removed_and_no_stale_navigation(self):
        self.assertFalse((ROOT/'pages/exchange-reconciliation').exists())
        self.assertNotIn('exchange-reconciliation',read(ROOT/'pages/pages.json')['pageOrder'])
        for p in ROOT.rglob('*.json'):self.assertNotIn('exchange-reconciliation',p.read_text(encoding='utf-8-sig'),str(p))
    def test_seven_views_keep_filter_state(self):
        bookmarks=[read(p) for p in (ROOT/'bookmarks').glob('*.bookmark.json')]
        marks=[b for b in bookmarks if b.get('explorationState',{}).get('activeSection')==TARGET]
        self.assertEqual(len(marks),7)
        for b in marks:
            self.assertTrue(b['options']['suppressData'])
            states=b['explorationState']['sections'][TARGET]['visualContainerGroups']
            licensed=b['name']==uid('bookmark')
            self.assertEqual(states[uid('group')]['isHidden'],not licensed)
            self.assertEqual(states[uid('common')]['isHidden'],licensed)
    def test_two_detail_tables_retain_native_id_and_job_titles(self):
        page=ROOT/'pages'/TARGET/'visuals'
        for i in range(2):
            v=read(page/uid('detail-'+str(i))/'visual.json')
            fields=v['visual']['query']['queryState']['Values']['projections']
            columns=[f['field'].get('Column',{}).get('Property') for f in fields]
            self.assertIn('Mailbox GUID',columns)
            self.assertEqual(v['parentGroupName'],uid('group'))
            if i==0:
                self.assertIn('AD Job Title',columns);self.assertIn('Entra Job Title',columns)
                self.assertTrue(any(f['field'].get('Measure',{}).get('Property')=='Exchange Scenario Row Target' for f in fields))
            else:self.assertIn('Retained On-premises',columns);self.assertIn('SE CAL Review',columns)
    def test_products_and_canvas(self):
        page=ROOT/'pages'/TARGET/'visuals'
        for product in ['E3','F3','SE']:
            v=read(page/uid('product-'+product)/'visual.json')
            names=[f['field']['Measure']['Property'] for f in v['visual']['query']['queryState']['Data']['projections']]
            self.assertEqual(len(names),2 if product=='SE' else 3)
            self.assertNotIn('Exchange Scenario '+product+' Buy Label',names)
            purchase=read(page/uid('purchase-'+product)/'visual.json')
            headline=purchase['visual']['query']['queryState']['Data']['projections']
            expected='Exchange Scenario SE Required Label' if product=='SE' else 'Exchange Scenario '+product+' Buy Label'
            self.assertEqual([f['field']['Measure']['Property'] for f in headline],[expected])
            size=lambda visual:float(visual['visual']['objects']['value'][0]['properties']['fontSize']['expr']['Literal']['Value'].rstrip('D'))
            self.assertGreater(size(purchase),size(v))
            self.assertEqual(purchase['parentGroupName'],uid('group'))
        for p in page.glob('*/visual.json'):
            pos=read(p)['position'];self.assertLessEqual(pos['x']+pos['width'],1920);self.assertLessEqual(pos['y']+pos['height'],1080)

    def test_mailbox_population_kpis_and_explicit_service_gap(self):
        v=read(ROOT/'pages'/TARGET/'visuals'/uid('population')/'visual.json')
        names=[f['field']['Measure']['Property'] for f in v['visual']['query']['queryState']['Data']['projections']]
        self.assertEqual(names,['Exchange Population Total','Exchange Population Register Retained',
            'Exchange Population Service Label','Exchange Population Potential Label'])
        self.assertEqual(v['parentGroupName'],uid('group'))

    def test_accessibility_describes_licensing_not_infrastructure(self):
        page=ROOT/'pages'/TARGET/'visuals'
        for tag in ['note','filter-0','filter-1','filter-2','detail-0','detail-1']:
            v=read(page/uid(tag)/'visual.json')
            label=v['visual']['visualContainerObjects']['general'][0]['properties']['altText']['expr']['Literal']['Value']
            self.assertNotIn('infrastructure',label.lower())
            self.assertTrue(any(word in label.lower() for word in ['licens','mailbox','retained']))

if __name__=='__main__':unittest.main()
