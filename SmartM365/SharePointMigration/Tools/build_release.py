"""Build a local, source-only release candidate from an explicit Git file list."""
__version__ = "1.0.4"
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import zipfile

PROJECT = Path(__file__).resolve().parents[1]
REPO = PROJECT.parents[1]
PREFIX = 'SmartM365/SharePointMigration/'
VERSION = '1.0.64'
ROOT_FILES = ['LICENSE', 'NOTICE', 'Install-WorkplaceCloudHub-CodeSigningCertificate.cmd',
              'Install-WorkplaceCloudHub-CodeSigningCertificate.ps1',
              'Certificates/workplacecloudhub.com-CodeSigning-D70ECB7B00377EBFB76B304C08DFC6620584E114.cer']


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-directory', required=True)
    args = parser.parse_args()
    destination = Path(args.output_directory).resolve()
    destination.mkdir(parents=True, exist_ok=True)
    archive = destination / f'SmartM365-SharePointMigration-{VERSION}.zip'
    if archive.exists():
        raise SystemExit(f'Refusing to replace an existing candidate: {archive}')
    tracked = subprocess.check_output(['git', '-C', str(REPO), 'ls-files', '-z', '--', PREFIX], text=True).split('\0')
    files = {p for p in tracked if p} | set(ROOT_FILES)
    required_notes = PREFIX + f'RELEASE-NOTES-{VERSION}.md'
    if required_notes not in files:
        raise SystemExit(f'Release notes must be committed: {required_notes}')
    if subprocess.check_output(
            ['git', '-C', str(REPO), 'status', '--porcelain', '--untracked-files=all', '--', PREFIX, *ROOT_FILES], text=True).strip():
        raise SystemExit('Release inputs must be clean and committed before packaging.')
    records = []
    excluded_workspace_markers = []
    for relative in sorted(files):
        path = REPO / relative
        product_relative = relative.removeprefix(PREFIX)
        if relative.startswith(PREFIX):
            if '.local.' in product_relative or '__pycache__' in product_relative or product_relative.startswith(('Tools/Python/', 'Output/')):
                raise SystemExit(f'Runtime/private file rejected: {relative}')
            if product_relative in ('Migrations/.gitkeep', 'ArchivedMigrations/.gitkeep'):
                excluded_workspace_markers.append(relative)
                continue
            if product_relative.startswith('ArchivedMigrations/'):
                raise SystemExit(f'Archived migration rejected: {relative}')
            if product_relative.startswith('Migrations/'):
                parts = Path(product_relative).parts
                if len(parts) >= 4 and parts[0] == 'Migrations' and parts[2:] in (
                        ('ShareGate', 'README.md'), ('ShareGate', 'MigrationReport', '.gitkeep')):
                    excluded_workspace_markers.append(relative)
                    continue
                raise SystemExit(f'Local migration rejected: {relative}')
        if path.is_symlink() or not path.is_file():
            raise SystemExit(f'Invalid input: {relative}')
        data = path.read_bytes()
        records.append({'path': relative, 'sha256': hashlib.sha256(data).hexdigest().upper(), 'bytes': len(data)})
    manifest = {'product': 'Smart SharePoint Migration Toolkit', 'version': VERSION, 'status': 'release-build',
                'excludedWorkspaceMarkers': len(excluded_workspace_markers),
                'baseCommit': subprocess.check_output(['git', '-C', str(REPO), 'rev-parse', 'HEAD'], text=True).strip(),
                'includesUncommittedChanges': bool(subprocess.check_output(
                    ['git', '-C', str(REPO), 'status', '--porcelain', '--untracked-files=all', '--', PREFIX, *ROOT_FILES], text=True).strip()),
                'files': records}
    manifest_bytes = (json.dumps(manifest, indent=2) + '\n').encode('utf-8')
    with zipfile.ZipFile(archive, 'x', compression=zipfile.ZIP_DEFLATED) as bundle:
        for record in records:
            bundle.write(REPO / record['path'], record['path'])
        bundle.writestr('release-manifest.json', manifest_bytes)
    with zipfile.ZipFile(archive) as bundle:
        assert bundle.testzip() is None
        for record in records:
            assert hashlib.sha256(bundle.read(record['path'])).hexdigest().upper() == record['sha256']
    digest = hashlib.sha256(archive.read_bytes()).hexdigest().upper()
    archive.with_suffix('.sha256').write_text(f'{digest}  {archive.name}\n', encoding='ascii')
    archive.with_suffix('.manifest.json').write_bytes(manifest_bytes)
    print(json.dumps({'archive': str(archive), 'files': len(records) + 1, 'sha256': digest}))


if __name__ == '__main__':
    from pathlib import Path as _ScriptPath
    import sys as _script_sys
    _script_sys.path.insert(0, str(_ScriptPath(__file__).resolve().parents[1] / 'Scripts'))
    from console_lifecycle import run_console_script
    run_console_script(main, __file__, __version__)
