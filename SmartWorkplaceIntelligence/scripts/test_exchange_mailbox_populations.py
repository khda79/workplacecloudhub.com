"""Version 1.0.0: synthetic population DAX cases, no source refresh."""
import copy
import unittest
from exchange_mailbox_populations import population_expression,population_measures
from test_exchange_license_scenario import fixture_query


def query_cases():
    table='Exchange Mailbox Reconciliation'
    def mailbox(key,retained=False):
        return {'Mailbox GUID':key,'Object GUID':key,'Primary SMTP Address':key+'@contoso.example',
            'Loaded Batch ID':'batch-a','Retained On-premises':'Matched retained decision' if retained else 'No matched retained decision; review register'}
    base={table:[mailbox('one',True),mailbox('two'),mailbox('three')],
        'Mailbox Migration Exceptions':[{'Loaded Batch ID':'batch-a','Requires Review':1,
            'Mailbox Identity':'','Candidate Identity':'local_mailboxes|two','Primary SMTP Address':''}]}
    cases=[]
    def add(name,mutate,predicate,expected):
        rows=copy.deepcopy(base);mutate(rows)
        expression='VAR _population=('+population_expression()+') RETURN COALESCE(COUNTROWS(FILTER(_population,'+predicate+')),0)'
        cases.append(dict(name=name,query=fixture_query(expression,rows,True),expected=[{'Value':expected}]))
    add('total_mailbox_grain',lambda r:None,'TRUE()',3)
    add('confirmed_register',lambda r:None,'[Population]="Register retained"',1)
    add('pending_not_confirmed',lambda r:None,'[Population]="Register review"',1)
    add('potential_excludes_retained_and_pending',lambda r:None,'[Population]="Potential before service review"',1)
    add('duplicate_mailbox_not_double_counted',lambda r:r[table].append(copy.deepcopy(r[table][0])),'TRUE()',3)
    add('confirmed_over_pending_hint',lambda r:r['Mailbox Migration Exceptions'][0].update({'Candidate Identity':'local_mailboxes|one'}),'[Population]="Register retained"',1)
    add('confirmed_not_counted_in_review',lambda r:r['Mailbox Migration Exceptions'][0].update({'Candidate Identity':'local_mailboxes|one'}),'[Population]="Register review"',0)
    add('pending_duplicate_not_double_counted',lambda r:r['Mailbox Migration Exceptions'].append(copy.deepcopy(r['Mailbox Migration Exceptions'][0])),'[Population]="Register review"',1)
    add('other_batch_hint_ignored',lambda r:r['Mailbox Migration Exceptions'][0].update({'Loaded Batch ID':'batch-b'}),'[Population]="Potential before service review"',2)
    add('smtp_hint_excludes_but_never_confirms',lambda r:r['Mailbox Migration Exceptions'][0].update({'Candidate Identity':'','Primary SMTP Address':'THREE@contoso.example'}),'[Population]="Register review"',1)
    add('withdrawn_hint_ignored',lambda r:r['Mailbox Migration Exceptions'][0].update({'Requires Review':0}),'[Population]="Potential before service review"',2)
    add('blank_hints_do_not_match_blanks',lambda r:(r['Mailbox Migration Exceptions'][0].update({'Candidate Identity':''}),r[table][1].update({'Object GUID':'','Primary SMTP Address':''})),'[Population]="Register review"',0)
    add('mixed_batch_unknown',lambda r:r[table][1].update({'Loaded Batch ID':'batch-b'}),'TRUE()',0)
    add('blank_native_mailbox_unknown',lambda r:r[table][1].update({'Mailbox GUID':''}),'TRUE()',0)
    return cases


class PopulationContracts(unittest.TestCase):
    def test_native_same_batch_and_deduplication(self):
        expression=population_expression()
        self.assertIn('DISTINCT(SELECTCOLUMNS(_scope,"Mailbox",[Mailbox],"Batch",[Batch]))',expression)
        self.assertIn('[ExceptionBatch]=_batch',expression)
        self.assertIn('_retained,"Register retained",_pending,"Register review"',expression)
        self.assertEqual(len(query_cases()),14)

    def test_missing_service_is_not_zero_or_generic(self):
        defs={m['name']:m for m in population_measures()}
        self.assertEqual(defs['Exchange Population Service Label']['expression'],'"Not captured"')
        self.assertIn('Generic is not Service',defs['Exchange Population Note']['expression'])
        self.assertIn('At most ',defs['Exchange Population Potential Label']['expression'])

    def test_se_total_does_not_subtract_existing_rights(self):
        expr=next(m['expression'] for m in population_measures() if m['name']=='Exchange Scenario SE Required Label')
        self.assertIn('[Exchange Scenario SE Need]',expr)
        self.assertNotIn('Covered',expr)
        self.assertNotIn('Gap',expr)
        self.assertNotIn('Buy',expr)


if __name__=='__main__':unittest.main()
