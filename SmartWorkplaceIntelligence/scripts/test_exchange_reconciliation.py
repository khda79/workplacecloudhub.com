"""Synthetic query-scoped fixtures for the actual reconciliation DAX.

query_cases() is executed by Modeling MCP against Desktop. Unit tests below
check contracts only; they do not claim execution of the DAX engine.
"""
import copy
import json
from pathlib import Path
import re
import unittest

from exchange_reconciliation_model import TABLE, measures, table_expression


def query_cases():
    expression = table_expression()
    schema = {}
    for table, column in re.findall(r"'([^']+)'\[([^\]]+)\]", expression):
        schema.setdefault(table, set()).add(column)
    batch = 'synthetic-batch-a'
    rows = {
        'Exchange Mailbox Server Reference': [dict(**{'Object GUID':'object-a','Server Name':'synthetic-server','Loaded Batch ID':batch})],
        'Mailbox Migration Inventory': [dict(**{'Mailbox GUID': 'mailbox-a', 'Object GUID': 'object-a',
            'Display Name': 'Synthetic person', 'Loaded Batch ID': batch, 'Mailbox Type': 'UserMailbox',
            'Mailbox Size GB': 1, 'Archive State': 'None', 'Hold and Retention': 'None observed',
            'Full Access Count': 0, 'Send As Count': 0, 'Send On Behalf': ''})],
        'AD User Evidence': [dict(ObjectGUID='object-a', TenantADUserKey='ad-a', TenantUserKey='user-a',
            CloudMatchStatus='Unique native SID', Title='Synthetic title', **{'Loaded Batch ID': batch})],
        'Native User Evidence': [dict(TenantUserKey='user-a', JobTitle='Synthetic title', **{'Loaded Batch ID': batch})],
        'Native User License': [dict(TenantUserKey='user-a',TenantSkuKey='sku-a',**{'Loaded Batch ID': batch})],
        'Native License SKU': [dict(TenantSkuKey='sku-a',SkuPartNumber='SYNTHETIC_E3',**{'Loaded Batch ID': batch})],
        'Native User Service Plan': [dict(TenantUserKey='user-a',ServicePlanName='EXCHANGE_S_ENTERPRISE',
            IsEnabled='true',AssignmentState='Success',**{'Loaded Batch ID': batch})],
        'Mailbox Migration Exceptions': [dict(**{'Mailbox Identity': 'local_mailboxes|object-a',
            'Status': 'Matched local', 'Retained': 1, 'Requires Review': 0, 'Loaded Batch ID': batch})],
    }
    numeric = {'Mailbox Size GB','Full Access Count','Send As Count','Retained','Requires Review'}
    def query(source):
        definitions=[]
        result=expression
        for i,(table,cols) in enumerate(schema.items()):
            name='__ExchangeFixture'+str(i)
            cols=sorted(cols)
            declarations=', '.join(json.dumps(c)+', '+('DOUBLE' if c in numeric else 'STRING') for c in cols)
            def literal(v):
                if v is None:return 'BLANK()'
                return json.dumps(v)
            records=['{'+', '.join(literal(row.get(c,0 if c in numeric else '')) for c in cols)+'}' for row in source[table]]
            data='{'+', '.join(records)+'}' if records else '{}'
            definitions.append("TABLE '"+name+"' = DATATABLE("+declarations+', '+data+')')
            result=result.replace("'"+table+"'", "'"+name+"'")
        return 'DEFINE\n'+'\n'.join(definitions)+'\nEVALUATE SELECTCOLUMNS(('+result+'), "Link", [Account Link], "Server", [Server Name], "Plan", [Mailbox Plan Observation], "F3", [F3 Review], "E3", [E3 Review], "Retained", [Retained On-premises], "CAL", [SE CAL Review])'
    cases=[]
    def add(name,source,expected):cases.append(dict(name=name,query=query(source),expected=expected))
    add('unique_plan2',rows,dict(Link='Unique AD GUID + native SID',Server='synthetic-server',Plan='Plan 2 observed',Retained='Matched retained decision'))
    for name,mutate,expected in [
        ('ambiguous_ad',lambda r:r['AD User Evidence'].append(copy.deepcopy(r['AD User Evidence'][0])),dict(Link='Ambiguous AD identity',Plan='Unknown - unresolved account')),
        ('ambiguous_cloud',lambda r:r['Native User Evidence'].append(copy.deepcopy(r['Native User Evidence'][0])),dict(Link='Ambiguous cloud identity')),
        ('different_batch',lambda r:r['AD User Evidence'][0].update({'Loaded Batch ID':'synthetic-batch-b'}),dict(Link='No same-batch AD identity',Plan='Unknown - unresolved account')),
        ('missing_identity',lambda r:r['Mailbox Migration Inventory'][0].update({'Object GUID':''}),dict(Link='No same-batch AD identity')),
        ('smtp_not_identity',lambda r:r['AD User Evidence'][0].update(CloudMatchStatus='No exact cloud SID match'),dict(Link='AD only; cloud unresolved')),
        ('plan_error',lambda r:r['Native User Service Plan'][0].update(AssignmentState='Error'),dict(Plan='No mailbox plan observed')),
        ('plan_pending',lambda r:r['Native User Service Plan'][0].update(AssignmentState='PendingActivation'),dict(Plan='No mailbox plan observed')),
        ('plan_disabled',lambda r:r['Native User Service Plan'][0].update(IsEnabled='false'),dict(Plan='No mailbox plan observed')),
        ('plan_other_batch',lambda r:r['Native User Service Plan'][0].update({'Loaded Batch ID':'synthetic-batch-b'}),dict(Plan='No mailbox plan observed')),
        ('foundation_only',lambda r:r['Native User Service Plan'][0].update(ServicePlanName='EXCHANGE_S_FOUNDATION'),dict(Plan='Foundation only - no mailbox right')),
        ('kiosk_over2',lambda r:(r['Native User Service Plan'][0].update(ServicePlanName='EXCHANGE_S_DESKLESS'),r['Mailbox Migration Inventory'][0].update({'Mailbox Size GB':2.01})),dict(Plan='Kiosk observed',F3='Above 2 GB - Kiosk alone insufficient')),
        ('exactly2',lambda r:r['Mailbox Migration Inventory'][0].update({'Mailbox Size GB':2}),dict(F3='At most 2 GB - frontline eligibility and usage unverified')),
        ('unknown_size',lambda r:r['Mailbox Migration Inventory'][0].update({'Mailbox Size GB':None}),dict(F3='Unknown mailbox size - review')),
        ('hold_review',lambda r:r['Mailbox Migration Inventory'][0].update({'Hold and Retention':'Review - preserve settings'}),dict(F3='Feature or delegation review required')),
        ('shared_mailbox',lambda r:r['Mailbox Migration Inventory'][0].update({'Mailbox Type':'SharedMailbox'}),dict(E3='Special mailbox rules; no automatic E3 requirement')),
        ('over50',lambda r:r['Mailbox Migration Inventory'][0].update({'Mailbox Size GB':50.01}),dict(E3='Above 50 GB - Plan 2 or equivalent review')),
        ('over100',lambda r:r['Mailbox Migration Inventory'][0].update({'Mailbox Size GB':100.01}),dict(E3='Above 100 GB - quota and archive design review')),
        ('exception_unreviewed',lambda r:r['Mailbox Migration Exceptions'][0].update({'Requires Review':1}),dict(Retained='No matched retained decision; review register')),
        ('server_other_batch',lambda r:r['Exchange Mailbox Server Reference'][0].update({'Loaded Batch ID':'synthetic-batch-b'}),dict(Server='Not observed')),
    ]:
        source=copy.deepcopy(rows);mutate(source);add(name,source,expected)
    return cases


class ReconciliationContracts(unittest.TestCase):
    def test_all_identity_edges_are_same_batch(self):
        e=table_expression()
        for table in ('AD User Evidence','Native User Evidence','Native User License','Native License SKU','Native User Service Plan','Mailbox Migration Exceptions','Exchange Mailbox Server Reference'):
            self.assertIn('"Loaded Batch ID", \''+table+'\'[Loaded Batch ID]',e)
        self.assertGreaterEqual(e.count('[Loaded Batch ID] = _batch'),7)
    def test_native_identity_only(self):
        e=table_expression()
        self.assertIn('COUNTROWS(_ad) = 1',e)
        self.assertIn('COUNTROWS(_cloud) = 1',e)
        self.assertIn('Unique native SID',e)
        self.assertNotIn('LOOKUPVALUE',e)
    def test_plans_not_capacity_or_foundation(self):
        e=table_expression()
        self.assertIn('[AssignmentState] = "Success"',e)
        self.assertIn('Foundation only - no mailbox right',e)
        self.assertNotIn('ConsumedUnits',e)
    def test_no_purchase_quantity_measure(self):
        self.assertTrue(all('Required' not in d['name'] for d in measures()))
        self.assertIn('contract, active coverage, accessors and frontline eligibility required',table_expression())
    def test_synthetic_cases_cover_failure_modes(self):
        cases=query_cases()
        self.assertEqual(len(cases),20)
        self.assertTrue(all('EVALUATE SELECTCOLUMNS' in c['query'] for c in cases))


if __name__ == '__main__':
    unittest.main()
