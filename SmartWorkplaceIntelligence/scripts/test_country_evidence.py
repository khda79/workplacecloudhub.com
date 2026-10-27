"""Synthetic country semantics and PowerShell/Python parity; no tenant inputs."""
import json
import subprocess
import unittest
from pathlib import Path
from unittest import mock

import workplace_country as country


class CountryTests(unittest.TestCase):
    def test_all_codes_names_and_aliases_match_in_both_runtimes(self):
        doc = json.loads(country.REFERENCE.read_text(encoding='utf-8-sig'))
        values = ['', None, 'Unknown', 'Unknown / unassigned', 'ZZ', 'Not a country']
        for row in doc['countries']:
            for value in [row['code'], row['name'], *row['aliases']]:
                values += [value, ' ' + value.lower() + ' ']
                self.assertEqual(country.normalize_country(value)['Code'], row['code'])
        module = Path(__file__).with_name('CountryEvidence.psm1')
        command = ("$ErrorActionPreference='Stop'; Import-Module '" + str(module).replace("'", "''") +
                   "' -Force; $v=([Console]::In.ReadToEnd() | ConvertFrom-Json); "
                   "@($v | ForEach-Object { Resolve-DirectoryCountry $_ }) | ConvertTo-Json -Compress")
        result = subprocess.run(['pwsh', '-NoProfile', '-Command', command],
                                input=json.dumps(values), text=True, encoding='utf-8', capture_output=True, check=True)
        self.assertEqual(json.loads(result.stdout), [country.normalize_country(v) for v in values])

    def test_directory_country_does_not_accept_email_domain_or_invalid_code(self):
        for value in ('user@contoso.com', 'es.contoso.com', 'ZZ', 'Atlantis', 'Error'):
            actual = country.normalize_country(value)
            self.assertEqual((actual['Code'], actual['Name'], actual['Status']), ('', 'Unknown', 'Unrecognized'))
            self.assertEqual(actual['Raw'], value)

    def test_ambiguous_alias_is_rejected(self):
        config = json.loads(country.REFERENCE.read_text(encoding='utf-8-sig'))
        config['countries'][0]['aliases'].append(config['countries'][1]['code'])
        country.reference.cache_clear()
        try:
            with mock.patch.object(Path, 'read_bytes', return_value=json.dumps(config).encode()):
                with self.assertRaisesRegex(ValueError, 'Ambiguous'):
                    country.reference()
        finally:
            country.reference.cache_clear()

    def test_native_contract_declares_break_and_exact_legacy_status(self):
        config = json.loads((country.REFERENCE.parent / 'native-evidence-contract.json.txt').read_text(encoding='utf-8-sig'))
        self.assertEqual(config['version'], '0.4.0')
        self.assertEqual(config['countrySemantics']['legacyStatus'], 'Entra country unavailable (legacy batch)')
        self.assertFalse(config['countrySemantics']['usageLocationFallback'])


if __name__ == '__main__':
    unittest.main()
