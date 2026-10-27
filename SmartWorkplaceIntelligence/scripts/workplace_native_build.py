"""Offline native facets inside one Workplace capture and publication lifecycle.

This module never collects, transfers, writes pointers, or opens Power BI.
The caller owns staging, locks, capture cleanup and atomic publication.
"""
import argparse
import datetime as dt
import json
import time
from pathlib import Path

import workplace_native_sources as sources
from workplace_native_tables import build_tables, log_timing
import workplace_teams_sites as teams_sites
import workplace_country as country


def identity_from_receipts(source, tenant):
    registry = sources.load_json(sources.REGISTRY)
    proof = sources.load_json(source / registry['Producers'][0]['Receipt'])
    identity = {name: proof.get(name) for name in
                ('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')}
    if identity['TenantKey'] != tenant or any(not isinstance(v, str) or not v.strip() for v in identity.values()):
        raise ValueError('Complete native source reporting identity is required')
    return identity


def validate(source, tenant, now=None, expected_identity=None):
    source = Path(source).resolve(strict=True)
    identity = identity_from_receipts(source, tenant)
    if expected_identity is not None and identity != expected_identity:
        raise ValueError('Native source identity differs from the effective tenant profile')
    contract = sources.load_json(sources.CONTRACT)
    reference = now or dt.datetime.now(dt.timezone.utc)
    started = time.perf_counter()
    evidence = sources.validate_sources(source, contract, tenant, now=reference, identity=identity)
    log_timing('source validation', started)
    return {'Version': 1, 'Identity': identity, 'ReferenceUtc': reference.isoformat(),
            'ContractSHA256': sources.sha(sources.CONTRACT), 'NativeContractVersion': contract['version'],
            'CountryReferenceSHA256': country.reference()[1], 'Evidence': evidence,
            'TeamsSites': teams_sites.validate(source, identity, reference)}


def build(source, output, audit):
    source, output = Path(source).resolve(strict=True), Path(output).resolve(strict=True)
    if output == source or source in output.parents or output in source.parents:
        raise ValueError('Native staging and captured sources must be separate sibling trees')
    contract = sources.load_json(sources.CONTRACT)
    if audit['Version'] != 1 or audit['ContractSHA256'] != sources.sha(sources.CONTRACT):
        raise ValueError('Native source validation contract changed')
    if audit.get('CountryReferenceSHA256') != sources.sha(country.REFERENCE):
        raise ValueError('Country reference changed after source validation')
    started = time.perf_counter()
    sources.recheck_sources(source, audit['Evidence'], contract)
    log_timing('source recheck before build', started)
    targets = [output / (t['name'] + '.csv') for t in contract['tables']]
    targets.extend(output / (name + '.csv') for name in teams_sites.SCHEMAS)
    if any(path.exists() for path in targets):
        raise ValueError('Native staging must not overwrite existing domain outputs')
    started = time.perf_counter()
    details = build_tables(source, output, contract, audit['Identity'], audit['Evidence'],
                          now=sources.utc(audit['ReferenceUtc']))
    log_timing('all native tables', started)
    counts = {}
    for table in contract['tables']:
        started = time.perf_counter()
        count = sources.check_rows(output / (table['name'] + '.csv'), table,
                                   audit['Identity']['TenantKey'], exact=True, identity=audit['Identity'])
        counts[table['name']] = count
        log_timing(f"row check {table['name']}", started)
    started = time.perf_counter()
    sources.validate_relationships(output, contract)
    log_timing('relationship validation', started)
    started = time.perf_counter()
    collaboration_counts = teams_sites.build(source, output, audit['Identity']['TenantKey'], audit['TeamsSites'])
    log_timing('Teams and sites', started)
    if audit.get('CountryReferenceSHA256') != sources.sha(country.REFERENCE):
        raise ValueError('Country reference changed after source validation')
    started = time.perf_counter()
    sources.recheck_sources(source, audit['Evidence'], contract)
    log_timing('source recheck after build', started)
    return dict(audit, OutputRows=counts, CollaborationRows=collaboration_counts, Details=details, Publication=False,
                Qualification='Native identities and observed links; not verified physical assets or effective assignments')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--tenant', required=True)
    parser.add_argument('--audit', type=Path, required=True)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--identity', type=Path)
    args = parser.parse_args()
    if args.output is None:
        expected = sources.load_json(args.identity) if args.identity else None
        result = validate(args.source, args.tenant, expected_identity=expected)
    else:
        audit = sources.load_json(args.audit)
        if audit['Identity']['TenantKey'] != args.tenant:
            raise ValueError('Native validation tenant mismatch')
        result = build(args.source, args.output, audit)
    args.audit.write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps({'Status': 'NativeEvidenceValidated', 'Publication': False,
                      'SourceFiles': len(result['Evidence']['Files']),
                      'OutputTables': len(result.get('OutputRows', {}))}))


if __name__ == '__main__':
    main()
