"""UTC age boundary cases, executed against the production color expression."""
import unittest
from dataset_footer_model import definitions


def query_cases():
    color=definitions()[-1]['expression']
    cases=[]
    for age,expected in [(0,'#334155'),(26,'#334155'),(73.99,'#334155'),(74,'#334155'),(74.0001,'#B42318'),(96,'#B42318'),(None,'#805400')]:
        value='BLANK()' if age is None else str(age)
        cases.append(dict(name='footer_'+str(age),query='EVALUATE ROW("Color", '+color.replace('[Dataset Batch Age Hours]',value)+')',expected={'Color':expected}))
    stamp=definitions()[0]['expression']
    for batch,expected in [('20300101T230000123Z-a','2030-01-01 23:00:00'),('20301301T000000000Z-a','Unknown'),('invalid','Unknown')]:
        expr=stamp.replace('[Dataset Loaded Batch ID]', '"'+batch+'"')
        cases.append(dict(name='timestamp_'+batch,query='EVALUATE ROW("Timestamp", VAR _s=('+expr+') RETURN IF(ISBLANK(_s),"Unknown",FORMAT(_s,"yyyy-MM-dd HH:mm:ss")))',expected={'Timestamp':expected}))
    # Execute the real age expression too: date subtraction has floating-point
    # noise, so rounding must preserve exactly 74 h versus the next second.
    # DAX DateTime arithmetic does not guarantee millisecond resolution.
    for extra_ms,expected in [(0,'#334155'),(1000,'#B42318')]:
        age=definitions()[1]['expression'].replace('[Dataset Batch Timestamp UTC]','DATE(2030,1,1)+TIME(23,0,0)').replace('UTCNOW()','DATE(2030,1,5)+TIME(1,0,0)+DIVIDE('+str(extra_ms)+',86400000)')
        cases.append(dict(name='real_age_74h_plus_'+str(extra_ms),query='EVALUATE ROW("Color", '+color.replace('[Dataset Batch Age Hours]','('+age+')')+')',expected={'Color':expected}))
    return cases


class FooterTests(unittest.TestCase):
    def test_threshold_is_strictly_after_74_hours(self):
        self.assertIn('_age>74',definitions()[-1]['expression'])
        self.assertNotIn('Today Paris',definitions()[-1]['expression'])
    def test_age_uses_utc_and_preserves_milliseconds(self):
        self.assertIn('UTCNOW()',definitions()[1]['expression'])
        self.assertIn('86400000',definitions()[0]['expression'])
    def test_ambiguous_age_remains_explicit(self):
        self.assertIn('ISBLANK(_age)',definitions()[-1]['expression'])


if __name__=='__main__':unittest.main()
