"""Version 1.0.0: future classification wiring, synthetic DAX only."""
import copy,unittest
from exchange_license_scenario import people_expression
from exchange_mailbox_populations import population_expression
from test_exchange_license_scenario import fixture_query


def query_cases():
    table='Exchange Mailbox Reconciliation'
    base={table:[{'Cloud Account Key':'user-a','AD Account Key':'ad-a','Mailbox GUID':'mail-a','Object GUID':'object-a',
        'Primary SMTP Address':'a@contoso.example','Mailbox Type':'UserMailbox','Mailbox Size GB':1,
        'Archive State':'None','Hold and Retention':'None observed','Full Access Count':0,'Send As Count':0,
        'Send On Behalf':'','Loaded Batch ID':'batch-a','Retained On-premises':'No matched retained decision; review register',
        'Account Link':'Unique AD GUID + native SID'}],
        'Mailbox Migration Exceptions':[], 'Native User Service Plan':[], 'Native User License':[], 'Native License SKU':[],
        'AD Account Classification':[{'Object GUID':'object-a','Loaded Batch ID':'batch-a','Service Classification':'Service'}]}
    result=[]
    def add(name,mutate,target):
        rows=copy.deepcopy(base);mutate(rows)
        result.append(dict(name=name,query=fixture_query(people_expression(True),rows),expected=[{'Target':target,'Covered':0}]))
    add('service_user_retained_as_se',lambda r:None,'SE')
    add('service_and_register_one_person',lambda r:r[table][0].update({'Retained On-premises':'Matched retained decision'}),'SE')
    add('other_classified_user_can_migrate',lambda r:r['AD Account Classification'][0].update({'Service Classification':'Other classified account'}),'F3')
    add('cross_batch_classification_not_used',lambda r:r['AD Account Classification'][0].update({'Loaded Batch ID':'batch-b'}),'Identity review')
    add('duplicate_classification_not_used',lambda r:r['AD Account Classification'].append(copy.deepcopy(r['AD Account Classification'][0])),'Identity review')
    add('unknown_classification_quarantined',lambda r:r['AD Account Classification'][0].update({'Service Classification':'Unknown classification'}),'Identity review')
    for name,flag,predicate,expected in [
            ('service_is_excluded_from_potential',False,'[Population]="Potential before service review"',0),
            ('retention_union_has_no_double_count',True,'[Population] IN {"Register retained","Service retained"}',1),
            ('service_kpi_keeps_register_overlap',True,'[Service Retained]=1',1)]:
        rows=copy.deepcopy(base)
        if flag:rows[table][0]['Retained On-premises']='Matched retained decision'
        expr='VAR _p=('+population_expression(True)+') RETURN COALESCE(COUNTROWS(FILTER(_p,'+predicate+')),0)'
        result.append(dict(name=name,query=fixture_query(expr,rows,True),expected=[{'Value':expected}]))
    return result


class ClassificationModelTests(unittest.TestCase):
    def test_classification_is_optional_until_materialization(self):
        self.assertNotIn("'AD Account Classification'",people_expression())
        self.assertNotIn("'AD Account Classification'",population_expression())
        self.assertIn('[ClassBatch]=_batch',people_expression(True))
        self.assertIn('COUNTROWS(_matches)=1',people_expression(True))
        self.assertIn('_classificationReview ||',people_expression(True))
        self.assertEqual(len(query_cases()),9)


if __name__=='__main__':unittest.main()
