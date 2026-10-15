"""Typed native output boundary. No source writes or timezone inference."""
import datetime as dt
import decimal
import math
import re


def count64(value):
    if value == '':
        return ''
    try:
        number = decimal.Decimal(str(value))
    except decimal.InvalidOperation as error:
        raise ValueError('Invalid non-negative int64 evidence') from error
    if not number.is_finite() or number < 0 or number > 9223372036854775807 or number != number.to_integral_value():
        raise ValueError('Invalid non-negative int64 evidence')
    return int(number)


def source_datetime(value):
    """ISO or explicit legacy dd/MM/yyyy HH:mm:ss; never infer a timezone."""
    value = str(value).strip()
    if not value:
        return ''
    if re.fullmatch(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}', value):
        parsed = dt.datetime.strptime(value, '%d/%m/%Y %H:%M:%S')
    elif re.fullmatch(r'\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})?)?', value):
        parsed = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
        # Preserve seven-digit source precision and explicit offsets verbatim.
        return '' if parsed.year < 1900 else value.replace(' ', 'T')
    else:
        raise ValueError('Unsupported source date format; expected ISO or legacy day-first timestamp')
    return '' if parsed.year < 1900 else parsed.isoformat()


def validate_value(value, kind):
    """Fail before graph expansion; publication validation remains authoritative."""
    raw = str(value).strip() if value is not None else ''
    if not raw or kind == 'string':
        return
    if kind == 'dateTime':
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})?)?', raw):
            raise ValueError('Expected ISO date/time')
        dt.datetime.fromisoformat(raw.replace('Z', '+00:00'))
    elif kind == 'int64':
        if not re.fullmatch(r'[+-]?\d+', raw) or not -9223372036854775808 <= int(raw) <= 9223372036854775807:
            raise ValueError('Expected int64')
    elif kind in ('double', 'decimal'):
        if not math.isfinite(float(raw)):
            raise ValueError('Expected finite number')
    elif kind == 'boolean':
        if raw.casefold() not in ('true', 'false'):
            raise ValueError('Expected boolean')
    else:
        raise ValueError('Unsupported contract type: ' + kind)
