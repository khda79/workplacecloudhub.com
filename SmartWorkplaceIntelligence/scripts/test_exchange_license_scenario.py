"""Version 1.0.0: synthetic DAX cases; query_cases runs through Modeling MCP.

No tenant exports are loaded and no model refresh is performed by these tests.
"""
import copy
import json
import re
import unittest
from exchange_license_scenario import people_expression,scenario_measures,stock_expression


def fixture_query(expression,rows,scalar=False):
    schema={}
    for table,column in re.findall(r"'([^']+)'\[([^\]]+)\]",expression):schema.setdefault(table,set()).add(column)
    definitions=[];result=expression
    numeric={'Mailbox Size GB','Full Access Count','Send As Count','Requires Review','Available Units'}
    for i,(table,columns) in enumerate(schema.items()):
        name='__ScenarioFixture'+str(i);columns=sorted(columns)
        declarations=', '.join(json.dumps(c)+', '+('DOUBLE' if c in numeric else 'STRING') for c in columns)
        records=['{'+', '.join('BLANK()' if row.get(c) is None else json.dumps(row[c]) for c in columns)+'}' for row in rows.get(table,[])]
        data='DATATABLE('+declarations+', {'+', '.join(records)+'})' if records else (
            'FILTER(DATATABLE('+declarations+', {{'+', '.join('BLANK()' for _ in columns)+'}}),FALSE())')
        definitions.append("TABLE '"+name+"' = "+data)
        result=result.replace("'"+table+"'","'"+name+"'")
    definitions.append("TABLE '__ScenarioHarness' = DATATABLE(\"X\",INTEGER,{{1}})")
    definitions.append("MEASURE '__ScenarioHarness'[Dataset Loaded Batch ID] = \"batch-a\"")
    end='ROW("Value",('+result+'))' if scalar else 'SELECTCOLUMNS(('+result+'),"Target",[Product],"Covered",[Covered],"PlanReview",[Plan Review])'
    return 'DEFINE\n'+'\n'.join(definitions)+'\nEVALUATE '+end


def query_cases():
    table='Exchange Mailbox Reconciliation'
    base={table:[{'Cloud Account Key':'user-a','AD Account Key':'ad-a','Mailbox GUID':'mail-a','Object GUID':'object-a',
        'Primary SMTP Address':'a@contoso.example','Mailbox Type':'UserMailbox','Mailbox Size GB':1,
        'Archive State':'None','Hold and Retention':'None observed','Full Access Count':0,'Send As Count':0,
        'Send On Behalf':'','Loaded Batch ID':'batch-a','Retained On-premises':'No matched retained decision; review register',
        'Account Link':'Unique AD GUID + native SID'}],
        'Mailbox Migration Exceptions':[], 'Native User Service Plan':[], 'Native User License':[], 'Native License SKU':[]}
    cases=[]
    def add(name,mutate,target='F3',covered=0,repair=None):
        rows=copy.deepcopy(base);mutate(rows)
        expected={'Target':target,'Covered':covered}
        if repair is not None:expected['PlanReview']=repair
        cases.append(dict(name=name,query=fixture_query(people_expression(),rows),expected=[expected]))
    def change(**values):return lambda r:r[table][0].update(values)
    def plan(r,name='EXCHANGE_S_ENTERPRISE',batch='batch-a',state='Success'):
        r['Native User Service Plan'].append({'TenantUserKey':'user-a','Loaded Batch ID':batch,'ServicePlanName':name,'IsEnabled':'true','AssignmentState':state})
    def suite(r,name='SPE_E3',state='Active'):
        r['Native User License'].append({'TenantUserKey':'user-a','TenantSkuKey':'sku-a','Loaded Batch ID':'batch-a','AssignmentState':state})
        r['Native License SKU'].append({'TenantSkuKey':'sku-a','Loaded Batch ID':'batch-a','SkuPartNumber':name})
    def exception(r,**values):
        x={'Loaded Batch ID':'batch-a','Requires Review':1,'Mailbox Identity':'','Candidate Identity':'local_mailboxes|object-a','Primary SMTP Address':''};x.update(values);r['Mailbox Migration Exceptions'].append(x)
    add('small_unlicensed',lambda r:None)
    add('exactly_two',change(**{'Mailbox Size GB':2}))
    add('over_two',change(**{'Mailbox Size GB':2.01}),'E3')
    add('unknown_size',change(**{'Mailbox Size GB':None}),'Technical review')
    add('over_hundred',change(**{'Mailbox Size GB':100.01}),'Technical review')
    add('delegation',change(**{'Full Access Count':1}),'E3')
    add('unknown_archive',change(**{'Archive State':'Unknown'}),'Technical review')
    add('retained',change(**{'Retained On-premises':'Matched retained decision'}),'SE')
    add('retained_plan1',lambda r:(change(**{'Retained On-premises':'Matched retained decision'})(r),plan(r,'EXCHANGE_S_STANDARD')),'SE',1)
    add('retained_kiosk',lambda r:(change(**{'Retained On-premises':'Matched retained decision'})(r),plan(r,'EXCHANGE_S_DESKLESS')),'SE',0)
    add('confirmed_retained_over_pending_duplicate',lambda r:(change(**{'Retained On-premises':'Matched retained decision'})(r),exception(r)),'SE',0)
    add('retained_native_ad_without_cloud',change(**{'Retained On-premises':'Matched retained decision','Cloud Account Key':'','Account Link':'AD only; cloud unresolved'}),'SE',0)
    add('pending_candidate_quarantined',lambda r:exception(r),'Register review')
    add('pending_address_quarantines_only',lambda r:exception(r,**{'Candidate Identity':'','Primary SMTP Address':'A@contoso.example'}),'Register review')
    add('withdrawn_not_retained',lambda r:exception(r,**{'Requires Review':0}))
    add('other_batch_register_ignored',lambda r:exception(r,**{'Loaded Batch ID':'batch-b'}))
    add('unresolved_identity',change(**{'Cloud Account Key':'','Account Link':'AD only; cloud unresolved'}),'Identity review')
    add('current_suite_owned_other_batch_plan_ignored',lambda r:(suite(r),plan(r,batch='batch-b')),covered=1,repair=1)
    add('other_batch_suite_ignored',lambda r:(suite(r),r['Native User License'][0].update({'Loaded Batch ID':'batch-b'}),r['Native License SKU'][0].update({'Loaded Batch ID':'batch-b'})))
    add('se_other_batch_plan_ignored',lambda r:(change(**{'Retained On-premises':'Matched retained decision'})(r),plan(r,batch='batch-b')),'SE',0)
    add('e3_assigned_effective',lambda r:(change(**{'Mailbox Size GB':4})(r),suite(r),plan(r)),'E3',1)
    add('error_plan_suite_already_owned',lambda r:(change(**{'Mailbox Size GB':4})(r),suite(r),plan(r,state='Error')),'E3',1,1)
    add('error_assignment_quarantined',lambda r:(change(**{'Mailbox Size GB':4})(r),suite(r,state='Error'),plan(r)),'Assignment review',0)
    add('f3_assigned_effective',lambda r:(suite(r,'SPE_F1'),plan(r,'EXCHANGE_S_DESKLESS')),'F3',1)
    add('e5_covers_f3',lambda r:(suite(r,'SPE_E5'),plan(r)),'F3',1)
    def another(r,**values):
        x=copy.deepcopy(r[table][0]);x['Mailbox GUID']='mail-b';x.update(values);r[table].append(x)
    add('deduplicate_worst_case',lambda r:another(r,**{'Mailbox Size GB':4}),'E3')
    add('conflicting_targets_review',lambda r:another(r,**{'Retained On-premises':'Matched retained decision'}),'Identity review')
    add('shared_is_not_one_user_license',change(**{'Mailbox Type':'SharedMailbox'}),'Special')
    cap={'License SKU Capacity':[{'Loaded Batch ID':'batch-a','SKU Part Number':'SPE_E3','Capacity Class':'Governed','Excluded':'false',
        'Snapshot ID':'snapshot-a','SKU ID':'sku-a','Available Units':7}],
        'License Report Qualification':[{'Loaded Batch ID':'batch-a','Scope':'M365_Licenses_ReportSkuCapacity.csv','State':'Qualified'}]}
    def stock(name,mutate,expected):
        rows=copy.deepcopy(cap);mutate(rows);cases.append(dict(name=name,query=fixture_query(stock_expression('E3'),rows,True),expected=[{'Value':expected}]))
    stock('stock_observed',lambda r:None,7)
    stock('zero_is_known',lambda r:r['License SKU Capacity'][0].update({'Available Units':0}),0)
    stock('absent_is_unknown',lambda r:r.update({'License SKU Capacity':[]}),None)
    stock('different_batch_stock_unknown',lambda r:r['License SKU Capacity'][0].update({'Loaded Batch ID':'batch-b'}),None)
    stock('unqualified_stock_unknown',lambda r:r['License Report Qualification'][0].update(State='Not qualified'),None)
    stock('duplicate_sku_unknown',lambda r:r['License SKU Capacity'].append(copy.deepcopy(r['License SKU Capacity'][0])),None)
    stock('negative_stock_unknown',lambda r:r['License SKU Capacity'][0].update({'Available Units':-1}),None)
    stock('mixed_snapshots_unknown',lambda r:r['License SKU Capacity'].append({**r['License SKU Capacity'][0],'SKU ID':'sku-b','Snapshot ID':'snapshot-b'}),None)
    buy=next(m['expression'] for m in scenario_measures() if m['name']=='Exchange Scenario E3 Buy')
    for name,gap,available,expected in [('buy_after_free_stock',9,4,5),('buy_clamped',2,4,0),('unknown_stock_not_zero',9,None,None),('unknown_need_not_zero',None,4,None)]:
        val=lambda n:'BLANK()' if n is None else str(n)
        query='DEFINE TABLE \'__ScenarioHarness\'=DATATABLE("X",INTEGER,{{1}})\nMEASURE \'__ScenarioHarness\'[Exchange Scenario E3 Gap]='+val(gap)+'\nMEASURE \'__ScenarioHarness\'[Exchange Scenario E3 Available]='+val(available)+'\nEVALUATE ROW("Value",('+buy+'))'
        cases.append(dict(name=name,query=query,expected=[{'Value':expected}]))
    return cases


class ScenarioContracts(unittest.TestCase):
    def test_required_failure_modes_and_explicit_assumptions(self):
        self.assertEqual(len(query_cases()),40)
        expression=people_expression()
        self.assertIn('_local && NOT(_remote),"SE",_pending',expression)
        self.assertIn('_pending,"Register review"',expression)
        self.assertIn('DISTINCT(SELECTCOLUMNS(_scope',expression)
        self.assertIn('[PlanBatch]=_batch',expression)
    def test_no_contractual_cal_purchase_inference(self):
        defs={m['name']:m for m in scenario_measures()}
        self.assertEqual(defs['Exchange Scenario SE Buy Label']['expression'],'"To qualify"')
        self.assertIn('frontline eligibility assumed',defs['Exchange Scenario Note']['expression'])
        self.assertIn('stock assumed allocatable',defs['Exchange Scenario Note']['expression'])

if __name__=='__main__':unittest.main()
