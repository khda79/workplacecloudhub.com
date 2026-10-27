"""Offline country normalization shared with CountryEvidence.psm1 via one JSON file.

Only explicit directory country belongs here. Licensing usageLocation is not a
fallback. No source data, network access or environment-dependent region lookup.
"""
import hashlib
import json
import re
from functools import lru_cache
from pathlib import Path

REFERENCE = Path(__file__).resolve().parents[1] / 'config/country-reference.json.txt'


@lru_cache(maxsize=1)
def reference():
    payload = REFERENCE.read_bytes()
    config = json.loads(payload.decode('utf-8-sig'))
    if config['schemaVersion'] != 1 or len(config['countries']) != 249:
        raise ValueError('Unsupported or incomplete country reference')
    aliases, codes = {}, set()
    for country in config['countries']:
        code = country['code']
        if not re.fullmatch('[A-Z]{2}', code) or code in codes or not country['name'].strip():
            raise ValueError('Invalid or duplicate ISO country')
        codes.add(code)
        for alias in [code, country['name'], *country['aliases']]:
            key = alias.strip().lower()
            if not key or key in ('unknown', 'unknown / unassigned'):
                raise ValueError('Invalid country alias')
            if key in aliases and aliases[key]['code'] != code:
                raise ValueError('Ambiguous country alias')
            aliases[key] = country
    return aliases, hashlib.sha256(payload).hexdigest().upper()


def normalize_country(value):
    raw = '' if value is None else str(value).strip()
    item = reference()[0].get(raw.lower())
    return {'Raw': raw, 'Code': item['code'] if item else '',
            'Name': item['name'] if item else 'Unknown',
            'Status': 'Reported' if item else 'Not provided' if not raw else
                      'Unknown' if raw.lower() in ('unknown', 'unknown / unassigned') else 'Unrecognized'}
