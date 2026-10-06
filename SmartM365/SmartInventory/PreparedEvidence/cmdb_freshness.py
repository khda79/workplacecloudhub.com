"""Shared acquisition-age policy for CMDB preparation and buffered readers.

No file access, fallback, collector invocation or transport-date substitution.
"""
import datetime as dt
import math
import re

VERSION = '0.1.1'
UTC = dt.timezone.utc


def partial_ad_coverage(producer, coverage):
    """Validate the explicit domain partition shared by preparation and readers."""
    if (producer != 'SmartM365-ActiveDirectory-Inventory.ps1'
            or not isinstance(coverage, dict) or coverage.get('Kind') != 'ADDomainCoverage'
            or coverage.get('Reason') != 'NonBlockingDomainErrors' or coverage.get('Status') != 'PartialAccepted'):
        raise ValueError('Invalid partial AD coverage declaration')
    sets = {}
    for field in ('ExpectedDomains', 'CollectedDomains', 'UnavailableDomains', 'NonBlockingDomainErrors'):
        values = coverage.get(field)
        if (not isinstance(values, list) or not values
                or any(not isinstance(v, str) or not re.fullmatch(r'[a-zA-Z0-9](?:[a-zA-Z0-9.-]*[a-zA-Z0-9])?', v) for v in values)
                or len({v.casefold() for v in values}) != len(values)):
            raise ValueError('Invalid partial AD coverage domain list: ' + field)
        sets[field] = {v.casefold() for v in values}
    if (sets['CollectedDomains'] & sets['UnavailableDomains']
            or sets['CollectedDomains'] | sets['UnavailableDomains'] != sets['ExpectedDomains']
            or not sets['UnavailableDomains'] <= sets['NonBlockingDomainErrors']):
        raise ValueError('Partial AD coverage is not explicitly tolerated or fully accounted for')
    return coverage


def utc(value):
    try:
        result = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
        if result.tzinfo is None:
            raise ValueError()
        return result.astimezone(UTC)
    except (ValueError, TypeError, AttributeError):
        raise ValueError('Acquisition timestamp requires an explicit timezone') from None


def hours(value):
    if type(value) not in (int, float) or not math.isfinite(value) or not 0 < value <= 8760:
        raise ValueError('Invalid freshness contract')
    return value


def policies(contract):
    maximum = hours(contract['maxAgeHours'])
    span = hours(contract['maxCollectionSpanHours'])
    names = [source['file'] for source in contract['sources']]
    if not names or len(set(names)) != len(names):
        raise ValueError('Missing or repeated freshness source')
    result = {name: dict(Group='Core', WarningAgeHours=maximum,
                         MaxAgeHours=maximum, MaxCollectionSpanHours=span) for name in names}
    groups = contract.get('freshnessGroups', [])
    if not isinstance(groups, list):
        raise ValueError('Invalid freshness groups')
    seen_names, seen_files = {'Core'}, set()
    for group in groups:
        name = group['name']
        files = group['sources']
        if (not isinstance(name, str) or not name.strip() or name in seen_names
                or not isinstance(files, list) or not files or len(set(files)) != len(files)
                or not set(files).issubset(result) or seen_files.intersection(files)):
            raise ValueError('Unknown, repeated or overlapping freshness group')
        warning = hours(group['warningAgeHours'])
        maximum = hours(group['maxAgeHours'])
        span = hours(group['maxCollectionSpanHours'])
        if warning > maximum:
            raise ValueError('Freshness warning cannot exceed the rejection limit')
        seen_names.add(name)
        seen_files.update(files)
        for file in files:
            result[file] = dict(Group=name, WarningAgeHours=warning,
                                MaxAgeHours=maximum, MaxCollectionSpanHours=span)
    return result


def producer_policy(contract, files):
    rules = policies(contract)
    if not files or any(file not in rules for file in files):
        raise ValueError('Producer freshness source mismatch')
    rule = rules[files[0]]
    if any(rules[file] != rule for file in files):
        raise ValueError('One producer cannot span different freshness groups')
    return rule


def evaluate(contract, records, now):
    """Inclusive hard boundary; warnings begin strictly after the target age."""
    now = utc(now.isoformat())
    rules = policies(contract)
    records = list(records)
    files = [record['File'] for record in records]
    if len(files) != len(set(files)) or set(files) != set(rules):
        raise ValueError('Freshness source set differs from the contract')
    future = now + dt.timedelta(minutes=5)
    sources, warnings, intervals, expiries = [], [], {}, []
    for record in records:
        file = record['File']
        rule = rules[file]
        start, end = utc(record['StartedAtUtc']), utc(record['CompletedAtUtc'])
        if start > end or end > future:
            raise ValueError('Invalid acquisition interval: ' + file)
        age = (now - start).total_seconds() / 3600
        if age > rule['MaxAgeHours']:
            raise ValueError('Stale acquisition evidence: ' + file)
        expires = start + dt.timedelta(hours=rule['MaxAgeHours'])
        expiries.append(expires)
        intervals.setdefault(rule['Group'], []).append((start, end, rule))
        state = 'Aging' if age > rule['WarningAgeHours'] else 'WithinTarget'
        sources.append(dict(File=file, **rule, StartedAtUtc=record['StartedAtUtc'],
                            CompletedAtUtc=record['CompletedAtUtc'], AgeHours=round(age, 6),
                            State=state, ExpiresAtUtc=expires.isoformat()))
        if state == 'Aging':
            warnings.append(dict(File=file, Group=rule['Group'], AgeHours=round(age, 6),
                                 Message=f'Acquisition age exceeds {rule["WarningAgeHours"]} hours; '
                                         f'accepted only up to {rule["MaxAgeHours"]} hours: {file}'))
    groups = []
    for name, items in intervals.items():
        span = max(item[1] for item in items) - min(item[0] for item in items)
        limit = items[0][2]['MaxCollectionSpanHours']
        if span > dt.timedelta(hours=limit):
            raise ValueError('Source collection interval exceeds the contract: ' + name)
        groups.append(dict(Name=name, SpanHours=round(span.total_seconds() / 3600, 6),
                           MaxCollectionSpanHours=limit))
    return dict(EvaluatedAtUtc=now.isoformat(), Sources=sources, Groups=groups,
                Warnings=warnings, ExpiresAtUtc=min(expiries).isoformat())
