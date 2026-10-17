"""Static check of every DAX reference in the semantic model TMDL.

Offline TMDL loading does not resolve DAX, so a measure or column removed from one
table can leave broken references elsewhere (Broken_Filters only in Desktop).
This test parses all DAX expressions (measures, calculated columns, calculated
tables, calculation items, format string and detail rows expressions) and checks:
- 'Table'[Name] / Table[Name]: the table exists and Name is one of its columns or measures;
- [Name] without table: Name is a measure, a column of any table (row context), or a
  column alias created by a string literal in the same expression.
Run: python -B -m unittest test_model_dax_references -v
SMARTWORKPLACE_TMDL may point to another SemanticModel definition folder.
"""
import os
from pathlib import Path
import re
import tempfile
import unittest

DEFINITION = Path(os.environ.get('SMARTWORKPLACE_TMDL') or (
    Path(__file__).resolve().parents[1] / 'pbip' / 'SmartWorkplaceIntelligence.SemanticModel' / 'definition'))

NAME = r"(?:'(?:[^']|'')+'|[A-Za-z_][\w.\-]*)"


def unquote(token):
    token = token.strip()
    if token.startswith("'") and token.endswith("'"):
        return token[1:-1].replace("''", "'")
    return token


def object_name(rest):
    match = re.match(r"\s*('(?:[^']|'')+'|[^\s=]+)", rest)
    return unquote(match.group(1)) if match else ''


def tabs(line):
    return len(line) - len(line.lstrip('\t'))


def parse_model(folder):
    tables, expressions = {}, []
    for path in sorted((folder / 'tables').glob('*.tmdl')):
        lines = path.read_text(encoding='utf-8-sig').splitlines()
        table = None
        i = 0
        while i < len(lines):
            line = lines[i]
            m = re.match(r'^table (.+)$', line)
            if m:
                table = object_name(m.group(1))
                tables[table] = {'columns': set(), 'measures': set()}
                i += 1
                continue
            m = re.match(r'^(\t+)(measure|column|calculationItem|formatStringDefinition|detailRowsDefinition|source) ?(.*)$', line)
            if table and m:
                depth, kind, rest = len(m.group(1)), m.group(2), m.group(3)
                name = object_name(rest) if kind in ('measure', 'column', 'calculationItem') else kind
                if kind == 'measure':
                    tables[table]['measures'].add(name)
                elif kind == 'column':
                    tables[table]['columns'].add(name)
                expression = None
                if kind == 'source':
                    # Only calculated partitions hold DAX; M partitions are skipped.
                    header = next((lines[k] for k in range(i - 1, max(i - 4, -1), -1) if re.match(r'^\tpartition ', lines[k])), '')
                    if header.rstrip().endswith('= calculated'):
                        expression = ''
                elif '=' in rest or kind in ('formatStringDefinition', 'detailRowsDefinition'):
                    expression = rest.split('=', 1)[1] if '=' in rest else ''
                if expression is not None:
                    body = [expression]
                    k = i + 1
                    while k < len(lines) and (lines[k].strip() == '' or tabs(lines[k]) > depth + 1 or
                                              (tabs(lines[k]) == depth + 1 and not re.match(r'^\t+[A-Za-z]+(:| =|$)', lines[k]))):
                        if lines[k].strip().startswith('///'):
                            break
                        body.append(lines[k])
                        k += 1
                    text = '\n'.join(body).strip()
                    if text and not text.startswith('```') and kind != 'source' or (kind == 'source' and text):
                        expressions.append({'table': table, 'object': '%s %s' % (kind, name), 'dax': text, 'file': path.name})
            i += 1
    return tables, expressions


def strip_dax(dax):
    literals = re.findall(r'"((?:[^"]|"")*)"', dax)
    text = re.sub(r'"(?:[^"]|"")*"', '""', dax)
    text = re.sub(r'/\*.*?\*/', ' ', text, flags=re.S)
    text = re.sub(r'(//|--)[^\n]*', ' ', text)
    return text, {unquote(l.replace('""', '"')) for l in literals}


def check(folder):
    tables, expressions = parse_model(folder)
    measures = {m for t in tables.values() for m in t['measures']}
    columns = {c for t in tables.values() for c in t['columns']}
    lower_tables = {k.lower(): k for k in tables}
    problems = []
    for item in expressions:
        text, literals = strip_dax(item['dax'])
        qualified = set()
        for m in re.finditer(r"(" + NAME + r")\s*\[([^\]]+)\]", text):
            table_token, name = m.group(1), m.group(2)
            table = unquote(table_token)
            # Function-call or keyword prefixes are not table names (e.g. VAR x = [M]).
            if not table_token.startswith("'") and table.upper() in ('VAR', 'RETURN', 'IN', 'AND', 'OR', 'NOT', 'ELSE', 'THEN'):
                continue
            actual = lower_tables.get(table.lower())
            if actual is None:
                if table_token.startswith("'") or table in columns or table in measures:
                    problems.append('%s / %s: unknown table %s[%s]' % (item['file'], item['object'], table, name))
                    qualified.add(m.start(2))
                continue
            qualified.add(m.start(2))
            if name not in tables[actual]['columns'] and name not in tables[actual]['measures']:
                problems.append('%s / %s: unknown column or measure %s[%s]' % (item['file'], item['object'], actual, name))
        for m in re.finditer(r"\[([^\]]+)\]", text):
            if m.start(1) in qualified:
                continue
            name = m.group(1)
            if name in measures or name in columns or name in literals:
                continue
            problems.append('%s / %s: unknown reference [%s]' % (item['file'], item['object'], name))
    return tables, expressions, problems


class ModelDaxReferenceTests(unittest.TestCase):
    def test_every_dax_reference_resolves(self):
        tables, expressions, problems = check(DEFINITION)
        self.assertGreater(len(tables), 100)
        self.assertGreater(len(expressions), 500)
        self.assertEqual(problems, [], '\n'.join(problems))

    def test_checker_detects_broken_references(self):
        table_a = '\n'.join([
            'table A',
            "\tmeasure 'Kept' = SUM ( B[Units] )",
            '\t\tformatString: 0',
            '',
            "\tmeasure 'Broken' =",
            "\t\t\t[Gone] + 'B'[Missing] + 'Nope'[X] + [Kept] + SUMX ( B, [Units] )",
            '\t\t\t+ MAXX ( ADDCOLUMNS ( B, "Alias", 1 ), [Alias] )',
            '\t\tformatString: 0', ''])
        table_b = '\n'.join(['table B', '\tcolumn Units', '\t\tdataType: int64', '\t\tsourceColumn: Units', ''])
        with tempfile.TemporaryDirectory() as scratch:
            folder = Path(scratch)
            (folder / 'tables').mkdir()
            (folder / 'tables' / 'A.tmdl').write_text(table_a, encoding='utf-8')
            (folder / 'tables' / 'B.tmdl').write_text(table_b, encoding='utf-8')
            _, expressions, problems = check(folder)
        self.assertEqual(len(expressions), 2)
        self.assertEqual(sorted(problems), sorted([
            'A.tmdl / measure Broken: unknown column or measure B[Missing]',
            'A.tmdl / measure Broken: unknown reference [Gone]',
            'A.tmdl / measure Broken: unknown table Nope[X]']), problems)

if __name__ == '__main__':
    import sys
    _, found, issues = check(DEFINITION)
    print(len(found), 'DAX expressions;', len(issues), 'unresolved references')
    print('\n'.join(issues))
    sys.exit(1 if issues else 0)
