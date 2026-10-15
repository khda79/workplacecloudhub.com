"""Pure mailbox reconciliation. No SMTP identity, APIs, files or publication.

Physical mailbox observations come from Get-Mailbox / EXO. A RemoteMailbox
recipient is an identity observation, not an additional observed cloud mailbox.
Unresolved recipients and every source-to-mailbox mapping remain inspectable.
"""
import base64
import collections
import uuid


def value(row, field):
    return next((str(v or '').strip() for k, v in row.items() if k.casefold() == field.casefold()), '')


def guid(raw):
    try:
        result = uuid.UUID(raw)
        return str(result) if result.int else ''
    except (ValueError, AttributeError):
        return ''


def anchor(raw):
    # Anchors are opaque case-sensitive bytes. Never infer one from an AD GUID.
    try:
        result = base64.b64decode(raw, validate=True)
        return result if result and any(result) else b''
    except ValueError:
        return b''


def reconcile(populations):
    observations = {}
    addresses = collections.defaultdict(list)
    by_guid = collections.defaultdict(list)
    by_anchor = collections.defaultdict(list)
    for family in ('local_mailboxes', 'remote_mailboxes', 'mailboxes'):
        for row in populations[family]:
            native = value(row, 'MailboxGuid' if family == 'mailboxes' else 'ObjectGuid')
            identity = (family, native.casefold())
            if not native or native == str(uuid.UUID(int=0)) or identity in observations:
                raise ValueError('Missing or repeated native mailbox identity: ' + family)
            smtp = value(row, 'PrimarySmtpAddress').casefold()
            exchange = guid(value(row, 'MailboxGuid' if family == 'mailboxes' else 'ExchangeGuid'))
            immutable = anchor(value(row, 'OnPremisesImmutableId' if family == 'mailboxes' else 'ImmutableId'))
            observations[identity] = dict(identity=identity, row=row, smtp=smtp, exchange=exchange,
                                          anchor=immutable, issues=[], target=None, basis='')
            if smtp:
                addresses[(family, smtp)].append(identity)
            if exchange:
                by_guid[(family, exchange)].append(identity)
            if immutable:
                by_anchor[(family, immutable)].append(identity)

    for identities in addresses.values():
        if len(identities) > 1:
            for identity in identities:
                observations[identity]['issues'].append('ConflictingPrimarySmtpAddress')

    # A source address duplicated in ANY family cannot be used to link a user.
    ambiguous_smtp = {smtp for (_, smtp), ids in addresses.items() if len(ids) > 1}

    def cloud_match(item):
        family = item['identity'][0]
        candidates = set()
        bases = []
        ambiguous = False
        for label, token, lookup in (('ExchangeGuid', item['exchange'], by_guid),
                                     ('ImmutableId', item['anchor'], by_anchor)):
            if not token:
                continue
            targets = lookup.get(('mailboxes', token), [])
            own = lookup.get((family, token), [])
            if len(own) > 1 or len(targets) > 1:
                ambiguous = True
            elif len(targets) == 1:
                candidates.add(targets[0]); bases.append(label)
        if ambiguous or len(candidates) > 1:
            item['issues'].append('AmbiguousMailboxTechnicalIdentity')
            return None, ''
        if not candidates:
            return None, ''
        target = next(iter(candidates))
        cloud = observations[target]
        # An anchor match cannot override a known, conflicting mailbox GUID,
        # nor can a GUID match override two different supplied anchors.
        if any(item[field] and cloud[field] and item[field] != cloud[field]
               for field in ('exchange', 'anchor')):
            item['issues'].append('ConflictingMailboxTechnicalIdentity')
            return None, ''
        return target, '+'.join(bases)

    for identity, item in observations.items():
        family = identity[0]
        if family == 'mailboxes':
            item.update(target=identity, basis='Native EXO mailbox identity')
        elif family == 'local_mailboxes':
            target, basis = cloud_match(item)
            item.update(target=target or identity, basis=basis or 'Native on-premises mailbox identity')
        else:
            target, basis = cloud_match(item)
            item.update(target=target, basis=basis or 'No unique technical mailbox link')
            if target is None:
                item['issues'].append('UnresolvedRemoteMailboxRecipient')

    # Independent unique signals can still converge from two distinct objects
    # in one source (one by GUID, another by anchor). Do not select a winner.
    target_sources = collections.defaultdict(list)
    for identity, item in observations.items():
        if item['target'] and item['target'] != identity:
            target_sources[(identity[0], item['target'])].append(identity)
    for identities in target_sources.values():
        if len(identities) > 1:
            for identity in identities:
                item = observations[identity]
                item['issues'].append('AmbiguousMailboxTechnicalIdentity')
                item['basis'] = 'Multiple source identities target one observed mailbox'
                item['target'] = identity if identity[0] == 'local_mailboxes' else None
                if item['target'] is None:
                    item['issues'].append('UnresolvedRemoteMailboxRecipient')
    entities = {identity:item for identity,item in observations.items()
                if identity[0] != 'remote_mailboxes' and item['target'] == identity}

    # Shared addresses without a technical link never collapse physical rows.
    physical_addresses = collections.defaultdict(list)
    for identity, item in entities.items():
        if item['smtp']:
            physical_addresses[item['smtp']].append(identity)
    for smtp, identities in physical_addresses.items():
        if len(identities) > 1:
            ambiguous_smtp.add(smtp)
            for identity in identities:
                if 'ConflictingPrimarySmtpAddress' not in observations[identity]['issues']:
                    observations[identity]['issues'].append('ConflictingPrimarySmtpAddress')
    return entities, observations, ambiguous_smtp
