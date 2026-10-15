"""Synthetic regression tests: no production data, collectors or publication."""
import csv
import io
import unittest
from workplace_native_types import count64, source_datetime, validate_value
from workplace_native_tables import activity_state


class NativeTypeTests(unittest.TestCase):
    def test_integer_serialization_without_float_rounding(self):
        for raw in ('8794', '8794.0', '0', '9223372036854775807'):
            result = count64(raw)
            self.assertIsInstance(result, int)
            stream = io.StringIO(); csv.writer(stream).writerow([result])
            self.assertNotIn('.', stream.getvalue())
            validate_value(result, 'int64')
        self.assertEqual(count64(''), '')

    def test_invalid_counts_are_not_rounded_or_zeroed(self):
        for raw in ('-1', '1.5', 'NaN', 'Infinity', '9223372036854775808', '1e100', 'x'):
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                count64(raw)

    def test_legacy_dates_are_day_first_without_timezone_inference(self):
        self.assertEqual(source_datetime('05/10/2026 09:08:07'), '2026-10-05T09:08:07')
        self.assertEqual(source_datetime('31/07/2026 08:00:34'), '2026-07-31T08:00:34')
        self.assertEqual(activity_state(source_datetime('05/10/2026 09:08:07'), '2026-10-15T12:00:00Z', 'true'), 'Timezone not established')
        self.assertEqual(activity_state('2026-10-05', '2026-10-15T12:00:00Z', 'true'), 'Timezone not established')
        self.assertEqual(activity_state('2026-10-05T09:08:07Z', '2026-10-15T12:00:00Z', 'true'), 'Active 30d')

    def test_iso_precision_offset_and_missingness(self):
        for raw in ('2026-10-05T09:08:07.1234567Z', '2026-10-05T09:08:07+02:00', '2026-10-05', '2026-10-05T09:08:07'):
            self.assertEqual(source_datetime(raw), raw)
            validate_value(raw, 'dateTime')
        self.assertEqual(source_datetime(''), '')
        self.assertEqual(source_datetime('01/01/0001 00:00:00'), '')
        self.assertEqual(source_datetime('0001-01-01T00:00:00Z'), '')

    def test_bad_dates_fail_and_legacy_is_not_valid_output(self):
        for raw in ('10/31/2026 00:00:00', '31/02/2026 00:00:00', 'yesterday'):
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                source_datetime(raw)
        for raw in ('05/10/2026 09:08:07', '2026-02-31T00:00:00Z'):
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                validate_value(raw, 'dateTime')


if __name__ == '__main__':
    unittest.main()
