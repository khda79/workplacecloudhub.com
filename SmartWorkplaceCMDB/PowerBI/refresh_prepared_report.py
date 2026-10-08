"""Hold a validated, immutable batch for a native Desktop refresh.

No collector, model write, XMLA processing, upload or persistent CSV copy.
Interactive parameter values are private, expire, and must not be logged.
"""
import argparse
import datetime as dt
import json
from pathlib import Path
import threading

import load_prepared_snapshot as reader
from prepared_read_session import ReadSession, make_server

VERSION = '1.0.0'
FIELDS = {'PreparedRoot', 'ExpectedIdentity', 'LifetimeSeconds', 'MaxBytes'}


def read_config(path):
    config = reader._json(Path(path).read_bytes())
    if not isinstance(config, dict) or set(config) - FIELDS:
        raise ValueError('Unknown refresh configuration field')
    root = config.get('PreparedRoot')
    if not isinstance(root, str) or not root.strip() or not Path(root).is_absolute():
        raise ValueError('PreparedRoot must be an absolute prepared-batch directory')
    identity = config.get('ExpectedIdentity')
    if (not isinstance(identity, dict) or set(identity) != set(reader.IDENTITY_FIELDS)
            or any(not isinstance(value, str) or not value.strip() for value in identity.values())):
        raise ValueError('ExpectedIdentity must independently specify all four tenant identity fields')
    ttl = config.get('LifetimeSeconds', 7200)
    budget = config.get('MaxBytes', 4 * 1024 * 1024 * 1024)
    if type(ttl) is not int or not 1 <= ttl <= 7200:
        raise ValueError('LifetimeSeconds must be between 1 and 7200')
    if type(budget) is not int or budget <= 0:
        raise ValueError('MaxBytes must be a positive integer')
    return dict(PreparedRoot=root, ExpectedIdentity=identity, LifetimeSeconds=ttl, MaxBytes=budget)


def validate(config):
    snapshot = reader.load_snapshot(config['PreparedRoot'], config['ExpectedIdentity'],
                                    max_bytes=config['MaxBytes'])
    policy = reader.freshness.evaluate(reader._json(snapshot.contract_bytes),
        snapshot.manifest['SourceEvidence']['Files'], dt.datetime.now(reader.UTC))
    summary = dict(Status='ValidatedForNativeDesktopRefresh', LauncherVersion=VERSION,
                   BatchSHA256=snapshot.batch_sha256, Tables=len(snapshot.tables),
                   ExpiresAtUtc=policy['ExpiresAtUtc'], FreshnessWarnings=policy['Warnings'],
                   ModelModified=False, RefreshStarted=False)
    return snapshot, summary


def hold_session(snapshot, config):
    session = ReadSession(snapshot, reader._json(snapshot.contract_bytes),
                          ttl_seconds=config['LifetimeSeconds'])
    server = None
    finished = threading.Event()
    try:
        server = make_server(session)
        server.timeout = 1
        print('PRIVATE session parameters; do not paste them into logs or commit them.', flush=True)
        print('In Desktop: Transform data > Edit parameters. Set these three text values:', flush=True)
        for name, value in (
            ('CMDBReadBaseUrl', 'http://127.0.0.1:' + str(server.server_port)),
            ('CMDBReadToken', session.token), ('CMDBReadBatch', snapshot.batch_sha256)):
            print(name + ' = ' + value, flush=True)
        print('Use the migrated CMDB project, not the historical report builders.', flush=True)
        print('Apply parameters, then refresh in Desktop. Use anonymous loopback credentials.', flush=True)
        print('Keep this console open until refresh completes. Save in Desktop after success.', flush=True)
        print('Press Enter here AFTER refresh completes to stop the reader (or Ctrl+C).', flush=True)
        print('The reader stops earlier if its lifetime or any source deadline expires.', flush=True)

        def wait_for_finish():
            try:
                input()
            except EOFError:
                pass
            finally:
                finished.set()

        threading.Thread(target=wait_for_finish, daemon=True).start()
        while (not finished.is_set() and session.monotonic() < session.deadline
               and session.utc_now() <= session.fresh_until):
            server.handle_request()
        if not finished.is_set():
            print('Read session expired. A pending refresh may fail; validate a new batch/session.', flush=True)
    except KeyboardInterrupt:
        pass
    finally:
        session.close()
        if server is not None:
            server.server_close()
        print('Read session stopped. No model save or refresh success is inferred.', flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', required=True, type=Path)
    parser.add_argument('--validate-only', action='store_true')
    args = parser.parse_args()
    try:
        config = read_config(args.config)
        snapshot, summary = validate(config)
        print(json.dumps(summary), flush=True)
        if not args.validate_only:
            hold_session(snapshot, config)
    except (ValueError, OSError, KeyError, TypeError) as error:
        parser.exit(1, 'Refresh preparation rejected: ' + str(error) + '\n')


if __name__ == '__main__':
    main()
