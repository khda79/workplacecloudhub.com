"""Synthetic identity cases only. No source files, collectors or network."""
import base64
import copy
import unittest
import uuid

from workplace_mailbox_identity import reconcile

G1 = '00000000-0000-0000-0000-000000000001'
G2 = '00000000-0000-0000-0000-000000000002'
G3 = '00000000-0000-0000-0000-000000000003'
A1 = base64.b64encode(uuid.UUID(G1).bytes_le).decode()
A2 = base64.b64encode(uuid.UUID(G2).bytes_le).decode()


class MailboxIdentityTests(unittest.TestCase):
    def setUp(self):
        self.data = dict(local_mailboxes=[], remote_mailboxes=[], mailboxes=[])

    def remote(self, native='r1', **fields):
        row = dict(ObjectGuid=native, PrimarySmtpAddress='same@synthetic.invalid', **fields)
        self.data['remote_mailboxes'].append(row)
        return row

    def cloud(self, native=G1, **fields):
        row = dict(MailboxGuid=native, PrimarySmtpAddress='same@synthetic.invalid', **fields)
        self.data['mailboxes'].append(row)
        return row

    def run_case(self):
        before = copy.deepcopy(self.data)
        result = reconcile(self.data)
        self.assertEqual(self.data, before, 'Native evidence was modified')
        entities, observations, _ = result
        self.assertEqual(len(observations), sum(map(len, self.data.values())))
        self.assertTrue(all(o['target'] is None or o['target'] in entities for o in observations.values()))
        return result

    def test_conflicting_remote_recipients_kept_only_one_physical_mailbox(self):
        self.remote('r1', ExchangeGuid=G2, ImmutableId=A2)
        self.remote('r2', ExchangeGuid=str(uuid.UUID(int=0)), ImmutableId=A1)
        self.cloud(OnPremisesImmutableId=A1)
        entities, obs, ambiguous = self.run_case()
        self.assertEqual(len(entities), 1)
        self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])
        self.assertEqual(obs[('remote_mailboxes','r2')]['target'], ('mailboxes',G1))
        self.assertEqual(obs[('remote_mailboxes','r2')]['basis'], 'ImmutableId')
        self.assertIn('same@synthetic.invalid', ambiguous)
        self.assertTrue(all('ConflictingPrimarySmtpAddress' in o['issues'] for k,o in obs.items() if k[0]=='remote_mailboxes'))

    def test_smtp_never_creates_a_mailbox_link(self):
        self.remote(); self.cloud()
        entities, obs, _ = self.run_case()
        self.assertEqual(len(entities),1)
        self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])

    def test_remote_only_is_not_an_observed_cloud_mailbox(self):
        self.remote(ExchangeGuid=G1)
        entities, obs, _ = self.run_case()
        self.assertEqual(entities,{})
        self.assertIn('UnresolvedRemoteMailboxRecipient',obs[('remote_mailboxes','r1')]['issues'])

    def test_unique_exchange_guid_survives_different_smtp(self):
        self.remote(ExchangeGuid=G1)['PrimarySmtpAddress']='old@synthetic.invalid'
        self.cloud()
        _, obs, _ = self.run_case()
        self.assertEqual(obs[('remote_mailboxes','r1')]['target'],('mailboxes',G1))

    def test_both_signals_agree(self):
        self.remote(ExchangeGuid=G1,ImmutableId=A1);self.cloud(OnPremisesImmutableId=A1)
        _, obs, _ = self.run_case()
        self.assertEqual(obs[('remote_mailboxes','r1')]['basis'],'ExchangeGuid+ImmutableId')

    def test_anchor_cannot_override_different_known_exchange_guid(self):
        self.remote(ExchangeGuid=G2,ImmutableId=A1);self.cloud(OnPremisesImmutableId=A1)
        _, obs, _ = self.run_case()
        self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])
        self.assertIn('ConflictingMailboxTechnicalIdentity',obs[('remote_mailboxes','r1')]['issues'])

    def test_guid_cannot_override_different_supplied_anchor(self):
        self.remote(ExchangeGuid=G1,ImmutableId=A2);self.cloud(OnPremisesImmutableId=A1)
        _, obs, _ = self.run_case()
        self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])

    def test_signals_pointing_to_different_cloud_objects_stay_unresolved(self):
        self.remote(ExchangeGuid=G1,ImmutableId=A2)
        self.cloud(OnPremisesImmutableId=A1);self.cloud(G2,OnPremisesImmutableId=A2)
        _, obs, _ = self.run_case()
        self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])

    def test_duplicate_anchor_in_either_source_is_not_first_match(self):
        for duplicate_cloud in (True,False):
            with self.subTest(cloud=duplicate_cloud):
                self.setUp();self.remote(ImmutableId=A1);self.cloud(OnPremisesImmutableId=A1)
                if duplicate_cloud:
                    self.cloud(G2,OnPremisesImmutableId=A1)
                else:
                    self.remote('r2',ImmutableId=A1)
                _, obs, _ = self.run_case()
                self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])

    def test_duplicate_remote_exchange_guid_is_not_first_match(self):
        self.remote(ExchangeGuid=G1);self.remote('r2',ExchangeGuid=G1);self.cloud()
        _, obs, _ = self.run_case()
        self.assertTrue(all(o['target'] is None for k,o in obs.items() if k[0]=='remote_mailboxes'))

    def test_different_signals_converging_on_one_mailbox_do_not_pick_a_winner(self):
        self.remote('r1',ExchangeGuid=G1);self.remote('r2',ImmutableId=A1)
        self.cloud(OnPremisesImmutableId=A1)
        entities, obs, _ = self.run_case()
        self.assertEqual(len(entities),1)
        self.assertTrue(all(o['target'] is None for k,o in obs.items() if k[0]=='remote_mailboxes'))

    def test_local_fan_in_restores_every_observed_local_identity(self):
        self.data['local_mailboxes']=[dict(ObjectGuid='l1',ExchangeGuid=G1),dict(ObjectGuid='l2',ImmutableId=A1)]
        self.cloud(OnPremisesImmutableId=A1)
        entities, obs, _ = self.run_case()
        self.assertEqual(len(entities),3)
        self.assertTrue(all(o['target']==k for k,o in obs.items()))

    def test_invalid_empty_zero_guids_and_anchors_never_link(self):
        for bad in ('','not-a-guid',str(uuid.UUID(int=0))):
            self.setUp();self.remote(ExchangeGuid=bad,ImmutableId='invalid!');self.cloud()
            _, obs, _ = self.run_case()
            self.assertIsNone(obs[('remote_mailboxes','r1')]['target'])

    def test_object_guid_does_not_invent_a_source_anchor(self):
        self.remote(G1);self.cloud(OnPremisesImmutableId=A1)
        _, obs, _ = self.run_case()
        self.assertIsNone(obs[('remote_mailboxes',G1)]['target'])

    def test_same_smtp_distinct_physical_identities_never_overwrite(self):
        self.cloud();self.cloud(G2)
        self.data['local_mailboxes']=[dict(ObjectGUID='local1',PrimarySMTPaddress='SAME@synthetic.invalid')]
        entities, _, ambiguous = self.run_case()
        self.assertEqual(len(entities),3)
        self.assertIn('same@synthetic.invalid',ambiguous)

    def test_local_and_cloud_same_unique_exchange_guid_count_once(self):
        self.cloud()
        self.data['local_mailboxes']=[dict(ObjectGUID='local1',ExchangeGuid=G1)]
        entities, obs, _ = self.run_case()
        self.assertEqual(len(entities),1)
        self.assertEqual(obs[('local_mailboxes','local1')]['target'],('mailboxes',G1))

    def test_missing_duplicate_or_zero_native_identity_remains_blocking(self):
        for bad in ('', str(uuid.UUID(int=0)), G1):
            self.setUp();self.cloud();self.cloud(bad)
            with self.assertRaisesRegex(ValueError,'native mailbox identity'):
                self.run_case()

    def test_order_does_not_choose_a_winner(self):
        self.remote('r2',ImmutableId=A2);self.remote('r1',ImmutableId=A1)
        self.cloud(G2,OnPremisesImmutableId=A2);self.cloud(OnPremisesImmutableId=A1)
        first=self.run_case()
        for rows in self.data.values(): rows.reverse()
        self.assertEqual(first,self.run_case())

    def test_rename_does_not_change_native_identity(self):
        self.cloud();first=self.run_case()[0]
        self.data['mailboxes'][0]['PrimarySmtpAddress']='renamed@synthetic.invalid'
        self.assertEqual(set(first),set(self.run_case()[0]))


if __name__ == '__main__':
    unittest.main()
