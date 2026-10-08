"""Ephemeral loopback reader; no report edits or persistent copies.

Run only for a controlled Desktop refresh after production qualification.
The random session parameters are private and must never be committed.
"""
import argparse
import csv
import datetime as dt
import hmac
import io
import json
import secrets
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import load_prepared_snapshot as reader

VERSION = '0.1.2'
PRODUCT_KEY = 'Application product key'


def product_key(row):
    """Match the preparer's strip/casefold tuple, excluding version.

    Hex prevents model case-insensitive collation from merging distinct keys;
    JSON tuple encoding prevents delimiter collisions. Missing publisher remains
    missing: no assumption that it equals an observed publisher.
    """
    values = [row[field].strip().casefold() for field in ('DisplayName', 'Publisher', 'Platform')]
    return json.dumps(values, ensure_ascii=False, separators=(',', ':')).encode('utf-8').hex()


class ReadSession:
    def __init__(self, snapshot, contract, *, ttl_seconds=3600, utc_now=None, monotonic=None):
        if type(ttl_seconds) is not int or not 1 <= ttl_seconds <= 7200:
            raise ValueError('Session lifetime must be between 1 and 7200 seconds')
        contract_bytes = snapshot.contract_bytes or reader.CONTRACT.read_bytes()
        if (contract != reader._json(contract_bytes)
                or reader._hash(snapshot.manifest['ContractSHA256']) != reader.digest(contract_bytes)):
            raise ValueError('Read-session contract mismatch')
        self.snapshot = snapshot
        self.token = secrets.token_hex(32)
        self.prefix = '/session/' + self.token + '/' + snapshot.batch_sha256 + '/'
        self.utc_now = utc_now or (lambda: dt.datetime.now(reader.UTC))
        self.monotonic = monotonic or time.monotonic
        self.deadline = self.monotonic() + ttl_seconds
        self.freshness = reader.freshness.evaluate(contract, snapshot.manifest['SourceEvidence']['Files'], self.utc_now())
        self.fresh_until = reader._utc(self.freshness['ExpiresAtUtc'])
        self.closed = False
        self.resources = dict(snapshot.tables)
        columns = {item['name'] + '.csv': list(item['columns']) for item in contract['tables']}
        name = 'DimDetectedApplication.csv'
        projected_columns = columns[name] + [PRODUCT_KEY]
        stream = io.StringIO(newline='')
        writer = csv.DictWriter(stream, fieldnames=projected_columns)
        writer.writeheader()
        for row in snapshot.iter_rows('DimDetectedApplication'):
            writer.writerow(dict(row, **{PRODUCT_KEY: product_key(row)}))
        self.resources[name] = stream.getvalue().encode('utf-8-sig')
        columns[name] = projected_columns
        self.resources['model-contract.json.txt'] = json.dumps({
            'ProtocolVersion': VERSION, 'Owner': reader.OWNER, 'Status': 'Validated',
            'BatchSHA256': snapshot.batch_sha256, 'Identity': snapshot.manifest['Identity'],
            'Tables': {name: {'Columns': columns[name], 'Rows': entry['Rows']}
                       for name, entry in snapshot.manifest['OutputFiles'].items()},
            'Projection': {name: PRODUCT_KEY},
            'SourceFreshness': self.freshness,
        }, ensure_ascii=False).encode('utf-8')

    def get(self, path):
        if self.closed or self.monotonic() >= self.deadline or self.utc_now() > self.fresh_until:
            return 410, b'Session expired; start a new validated read session.'
        if not path.startswith('/session/'):
            return 404, b'Not found'
        parts = path.split('/')
        if len(parts) != 5 or not hmac.compare_digest(parts[2], self.token):
            return 404, b'Not found'
        if parts[3] != self.snapshot.batch_sha256:
            return 409, b'Pinned batch mismatch'
        data = self.resources.get(parts[4])
        if data is None:
            return 404, b'Not found'
        return 200, data

    def close(self):
        self.closed = True
        self.resources.clear()


def make_server(session):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass  # Never log session tokens, paths, table rows or tenant identity.

        def do_GET(self):
            expected_host = '127.0.0.1:' + str(self.server.server_port)
            if self.headers.get('Host') != expected_host or 'Origin' in self.headers:
                status, body = 403, b'Loopback client required'
            else:
                status, body = session.get(self.path)
            self.send_response(status)
            self.send_header('Cache-Control', 'no-store, max-age=0')
            self.send_header('Content-Type', 'application/json; charset=utf-8' if self.path.endswith('.json.txt') and status == 200 else 'text/csv; charset=utf-8' if status == 200 else 'text/plain; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            try:
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionResetError):
                pass

    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    server.daemon_threads = True
    return server


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', required=True)
    parser.add_argument('--tenant-key', required=True)
    parser.add_argument('--organization-key', required=True)
    parser.add_argument('--environment-key', required=True)
    parser.add_argument('--tenant-id', required=True)
    parser.add_argument('--ttl-seconds', type=int, default=3600)
    parser.add_argument('--max-bytes', type=int, default=4 * 1024 * 1024 * 1024)
    parser.add_argument('--contract', type=reader.Path, default=reader.CONTRACT)
    parser.add_argument('--registry', type=reader.Path, default=reader.REGISTRY)
    parser.add_argument('--expected-manifest-sha256')
    args = parser.parse_args()
    identity = dict(TenantKey=args.tenant_key, OrganizationKey=args.organization_key,
                    EnvironmentKey=args.environment_key, TenantId=args.tenant_id)
    snapshot = reader.load_snapshot(args.root, identity, max_bytes=args.max_bytes,
        contract_path=args.contract, registry_path=args.registry,
        expected_manifest_sha256=args.expected_manifest_sha256)
    session = ReadSession(snapshot, reader._json(snapshot.contract_bytes), ttl_seconds=args.ttl_seconds)
    server = make_server(session)
    server.timeout = 1
    print(json.dumps({'Status': 'CandidateReadSessionReady',
                      'BaseUrl': 'http://127.0.0.1:' + str(server.server_port),
                      'SessionToken': session.token, 'BatchSHA256': snapshot.batch_sha256,
                      'LifetimeSeconds': args.ttl_seconds}), flush=True)
    try:
        while session.monotonic() < session.deadline and session.utc_now() <= session.fresh_until:
            server.handle_request()
    except KeyboardInterrupt:
        pass
    finally:
        session.close()
        server.server_close()


if __name__ == '__main__':
    main()
