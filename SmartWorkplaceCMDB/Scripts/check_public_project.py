"""Validate public PBIP closure and privacy without refresh or model writes."""
import argparse
import json
from pathlib import Path
import re

PARAMETERS = {
    'CMDBReadBaseUrl': '"http://127.0.0.1:1"',
    'CMDBReadToken': '"CONFIGURE_PRIVATE_SESSION"',
    'CMDBReadBatch': '"CONFIGURE_PRIVATE_BATCH"',
    'CMDBLicenseReportRoot': '"CONFIGURE_PRIVATE_LICENSE_ROOT"',
    'CMDBReportDataRoot': '"CONFIGURE_PRIVATE_LEGACY_ROOT"',
}
IDENTITY = '[TenantKey="CONFIGURE_TENANT_KEY", OrganizationKey="CONFIGURE_ORGANIZATION_KEY", EnvironmentKey="CONFIGURE_ENVIRONMENT_KEY", TenantId="CONFIGURE_TENANT_ID"]'
PRIVATE_PATH = re.compile(r'(?<![A-Za-z0-9])[A-Za-z]:[\\/]|\\\\[A-Za-z0-9]|https?://[^\s"]*sharepoint\.com', re.I)
ENTITY_COLUMN = re.compile(r'Tenant(?:User|Device|Group|ADObject|ADGroup)Key|UserPrincipalName|UPN|Mail|DNSHostName|Hostname|DeviceName|DisplayName', re.I)
FORBIDDEN = {'.pbi', '.local', '.local-review', 'ReportData', 'Validation', 'Validation.Room'}
SUFFIXES = {'.csv', '.parquet', '.abf', '.pbix', '.zip', '.log', '.pyc'}


def read_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from strings(child)
    elif isinstance(value, list):
        for child in value:
            yield from strings(child)


def nodes(value):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from nodes(child)
    elif isinstance(value, list):
        for child in value:
            yield from nodes(child)


def expression(value):
    return '\n'.join(value) if isinstance(value, list) else value


def validate(project):
    project = Path(project)
    errors = []
    pbip = read_json(project / 'SmartWorkplaceCMDB.pbip')
    report = project / pbip['artifacts'][0]['report']['path']
    binding = read_json(report / 'definition.pbir')['datasetReference']
    if set(binding) != {'byPath'}:
        errors.append('Report must bind to a relative local semantic model')
        return errors, {}
    relative_model = binding['byPath']['path']
    if relative_model != '../SmartWorkplaceCMDB.SemanticModel':
        errors.append('Unexpected semantic-model binding')
        return errors, {}
    model_dir = project / 'SmartWorkplaceCMDB.SemanticModel'
    model = read_json(model_dir / 'model.bim')['model']
    expected = {table['name']: table for table in model['tables']}
    params = {entry['name']: expression(entry['expression']) for entry in model['expressions']}
    for name, placeholder in PARAMETERS.items():
        if params.get(name, '').split(' meta ', 1)[0] != placeholder:
            errors.append('Non-neutral public parameter: ' + name)
    if params.get('CMDBExpectedIdentity') != IDENTITY:
        errors.append('Public expected identity must be neutral')
    for path in project.rglob('*'):
        relative = path.relative_to(project)
        if path.is_symlink() or any(part in FORBIDDEN for part in relative.parts):
            errors.append('Private/link artifact: ' + str(relative))
        if not path.is_file():
            continue
        if path.suffix.lower() in SUFFIXES or path.name in ('localSettings.json', 'diagramLayout.json'):
            errors.append('Runtime artifact: ' + str(relative))
        if path.suffix in ('.json', '.bim', '.pbip', '.pbir', '.pbism') or path.name == '.platform':
            value = read_json(path)
            if any(PRIVATE_PATH.search(text) for text in strings(value)):
                errors.append('Private source path: ' + str(relative))
            # Real identity keys in saved filters/selections are not public defaults.
            for item in nodes(value):
                for key in ('filter', 'expansionStates'):
                    selected = item.get(key)
                    if selected is None:
                        continue
                    properties = [n['Column']['Property'] for n in nodes(selected)
                                  if isinstance(n.get('Column'), dict) and 'Property' in n['Column']]
                    literals = [n['Literal'].get('Value') for n in nodes(selected)
                                if isinstance(n.get('Literal'), dict)]
                    if any(ENTITY_COLUMN.search(p) for p in properties):
                        if any(v not in (None, 'null', 'false', 'true') for v in literals):
                            errors.append('Private entity selection: ' + str(relative))
            if path.name == 'visual.json':
                for item in nodes(value):
                    for kind, collection in (('Column', 'columns'), ('Measure', 'measures')):
                        ref = item.get(kind)
                        if not isinstance(ref, dict):
                            continue
                        table = ref.get('Expression', {}).get('SourceRef', {}).get('Entity')
                        if not table:  # Source aliases are validated by the PBIR CLI.
                            continue
                        names = {x['name'] for x in expected.get(table, {}).get(collection, [])}
                        if ref.get('Property') not in names:
                            errors.append('Unresolved visual binding: ' + str(relative))
    pages_root = report / 'definition/pages'
    order = read_json(pages_root / 'pages.json')['pageOrder']
    if len(order) != len(set(order)) or any(not (pages_root / p / 'page.json').is_file() for p in order):
        errors.append('Page order must contain unique existing pages')
    for relation in model.get('relationships', []):
        for end in ('from', 'to'):
            table = relation[end + 'Table']
            column = relation[end + 'Column']
            if column not in {x['name'] for x in expected.get(table, {}).get('columns', [])}:
                errors.append('Unresolved relationship endpoint: ' + relation['name'])
    # Reject embedded imported data rather than mistaking a BIM for a cache.
    for table in model['tables']:
        for partition in table.get('partitions', []):
            source = partition.get('source', {})
            text = expression(source.get('expression', ''))
            if source.get('type') not in ('m', 'calculated') or 'Binary.Decompress' in text:
                errors.append('Unexpected embedded/source partition: ' + table['name'])
    summary = dict(Pages=len(order), Tables=len(expected),
                   Measures=sum(len(t.get('measures', [])) for t in model['tables']),
                   Relationships=len(model.get('relationships', [])))
    return sorted(set(errors)), summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('project', type=Path)
    args = parser.parse_args()
    errors, summary = validate(args.project)
    print(json.dumps(dict(Status='Failed' if errors else 'VerifiedPublicProject',
                          **summary, Errors=errors)))
    if errors:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
