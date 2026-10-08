"""Offline refresh lifecycle/configuration checks; no model or tenant writes."""
import copy
import io
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch, Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'PowerBI'))
import refresh_prepared_report as refresh
import prepared_read_session as session_reader
import test_prepared_snapshot as fixtures


class RefreshTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.SnapshotTests('test_all_46_contract_tables_and_only_retained_bytes_are_returned')
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.path = self.fixture.root.parent / 'refresh.local.json'
        self.config = dict(PreparedRoot=str(self.fixture.root), ExpectedIdentity=self.fixture.identity.copy())

    def read(self, config):
        self.path.write_text(json.dumps(config), encoding='utf-8')
        return refresh.read_config(self.path)

    def test_defaults(self):
        result = self.read(self.config)
        self.assertEqual(result['LifetimeSeconds'], 7200)
        self.assertEqual(result['MaxBytes'], 4294967296)

    def test_unknown_fields_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Unknown'):
            self.read(dict(self.config, Upload=True))

    def test_missing_identity_cannot_be_inferred(self):
        config = copy.deepcopy(self.config)
        del config['ExpectedIdentity']['TenantId']
        with self.assertRaisesRegex(ValueError, 'independently'):
            self.read(config)

    def test_invalid_roots_rejected(self):
        for value in ('', 'relative/DATA', None, 42):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, 'absolute'):
                self.read(dict(self.config, PreparedRoot=value))

    def test_invalid_lifetime_rejected(self):
        for value in (0, 7201, True, 2.5, '7200'):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, 'Lifetime'):
                self.read(dict(self.config, LifetimeSeconds=value))

    def test_invalid_budget_rejected(self):
        for value in (0, -1, True, '1024'):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, 'MaxBytes'):
                self.read(dict(self.config, MaxBytes=value))

    def test_summary_is_redacted_and_not_a_refresh_result(self):
        snapshot = self.fixture.load()
        with patch.object(refresh.reader, 'load_snapshot', return_value=snapshot), \
                patch.object(refresh, 'dt') as clock:
            clock.datetime.now.return_value = self.fixture.now
            actual, summary = refresh.validate(self.read(self.config))
        self.assertIs(actual, snapshot)
        self.assertEqual(summary['Tables'], 46)
        self.assertFalse(summary['RefreshStarted'])
        self.assertFalse(summary['ModelModified'])
        text = json.dumps(summary)
        for private in (str(self.fixture.root), self.fixture.identity['TenantKey'], 'SessionToken'):
            self.assertNotIn(private, text)

    def test_ctrl_c_closes_owned_reader(self):
        config = self.read(self.config)
        session = session_reader.ReadSession(self.fixture.load(), self.fixture.contract,
            utc_now=lambda: self.fixture.now)
        server = Mock(server_port=12345)
        server.handle_request.side_effect = KeyboardInterrupt
        # Keep the input thread alive only until this test releases it.
        import threading
        release = threading.Event()
        def wait():
            release.wait(3)
            return ''
        try:
            with patch.object(refresh, 'ReadSession', return_value=session), \
                    patch.object(refresh, 'make_server', return_value=server), \
                    patch('builtins.input', side_effect=wait), patch('sys.stdout', new=io.StringIO()):
                refresh.hold_session(session.snapshot, config)
        finally:
            release.set()
        self.assertTrue(session.closed)
        self.assertEqual(session.resources, {})
        server.server_close.assert_called_once()

    def test_bind_failure_clears_owned_session(self):
        session = session_reader.ReadSession(self.fixture.load(), self.fixture.contract,
            utc_now=lambda: self.fixture.now)
        with patch.object(refresh, 'ReadSession', return_value=session), \
                patch.object(refresh, 'make_server', side_effect=OSError('bind failed')), \
                patch('sys.stdout', new=io.StringIO()), self.assertRaises(OSError):
            refresh.hold_session(session.snapshot, self.read(self.config))
        self.assertTrue(session.closed)


if __name__ == '__main__':
    unittest.main()
