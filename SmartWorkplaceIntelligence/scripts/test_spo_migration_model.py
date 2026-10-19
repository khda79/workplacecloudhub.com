"""Offline model-contract and filter-path checks; no DAX execution or refresh."""
import re
import unittest
import migration_model_definitions as model


class MigrationModelTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.defs = model.definitions()
        cls.tables = {t['Name']: t for t in cls.defs['tables']}
        cls.measures = {m['name']: m for m in cls.defs['measures']}

    def test_output_schema_and_columns_are_not_inferred(self):
        for t in model.CONTRACT['tables']:
            actual = {c['Name']: c for c in self.tables[t['table']]['Columns']}
            self.assertTrue(set(c['name'] for c in t['columns']).issubset(actual))
            for c in t['columns']:
                self.assertEqual(actual[c['name']]['SourceColumn'], c['name'])
                self.assertNotIn(actual[c['name']]['DataType'], ['Double', 'Unknown'])

    def test_optional_group_fails_closed_without_hiding_errors(self):
        reader = self.defs['expressions'][0]['expression']
        for t in model.CONTRACT['tables']:
            self.assertIn(t['file'], reader)
        self.assertIn('Incomplete migration evidence group', reader)
        self.assertIn('Migration evidence schema mismatch', reader)
        self.assertIn('Foreign migration tenant', reader)
        self.assertIn('Not available in selected batch', reader)
        self.assertNotIn('try ', reader)
        self.assertNotIn('Excel.Workbook', reader)

    def test_no_active_date_path_to_current_plan(self):
        rels = [r for r in self.defs['relationships'] if r['fromTable'] == 'SPO Migration Plan' and r['toTable'] == 'Date']
        self.assertEqual({r['fromColumn'] for r in rels}, {'Planned Date', 'Actual Date', 'Validation Date'})
        self.assertTrue(all(not r['isActive'] for r in rels))
        for role in ['Planned', 'Actual', 'Validation']:
            eligible = [m for m in self.defs['measures'] if 'USERELATIONSHIP' in m['expression'] and f'[{role} Date]' in m['expression']]
            self.assertEqual(len(eligible), 1)
        snapshots = [r for r in self.defs['relationships'] if r['fromTable'] == 'SPO Migration History' and r['toTable'] == 'SPO Migration Snapshot Date']
        self.assertEqual(len(snapshots), 1)
        self.assertTrue(snapshots[0]['isActive'])

    def test_relationships_have_no_bidirectional_or_parallel_filter_paths(self):
        graph = {}
        for r in self.defs['relationships']:
            self.assertEqual(r['crossFilteringBehavior'], 'OneDirection')
            self.assertEqual((r['fromCardinality'], r['toCardinality']), ('Many', 'One'))
            if r['isActive']:
                graph.setdefault(r['toTable'], []).append(r['fromTable'])

        def walk(node, trail):
            self.assertNotIn(node, trail, 'Active filter cycle')
            result = [node]
            for child in graph.get(node, []):
                result.extend(walk(child, trail + [node]))
            return result

        for node in graph:
            reached = walk(node, [])
            self.assertEqual(len(reached), len(set(reached)), f'Parallel active paths from {node}')

    def test_historical_attributes_are_not_current_plan_lookups(self):
        for col in ['Source Type', 'Business Unit', 'Phase']:
            targets = [r for r in self.defs['relationships'] if r['toTable'] == 'SPO Migration ' + col]
            self.assertEqual({r['fromTable'] for r in targets}, {'SPO Migration Plan', 'SPO Migration History'})
        history = self.tables['SPO Migration History']['MExpression']
        self.assertIn('Is Latest Daily Snapshot', history)
        self.assertIn('Order.Descending', history)
        self.assertNotIn('SPO Migration Plan', history)
        self.assertIn('[Present In Register]=TRUE()', self.measures['SPO Historical Migrate Operations']['expression'])

    def test_reported_and_validated_use_one_denominator_with_cancellations(self):
        for name in ['SPO Reported Completion', 'SPO Validated Completion']:
            self.assertIn('[SPO Migrate Population]', self.measures[name]['expression'])
        denominator = self.measures['SPO Migrate Population']['expression']
        self.assertIn('REMOVEFILTERS', denominator)
        self.assertNotIn('CANCELLED', denominator)
        self.assertIn('[Disposition]="MIGRATE"', self.measures['SPO Migrate Operations']['expression'])
        self.assertIn('[Is Validated]=TRUE()', self.measures['SPO Validated Done']['expression'])
        self.assertIn('"n/a"', self.measures['SPO Reported Progress']['expression'])

    def test_storage_stays_at_collection_grain_and_shared_is_unmeasured(self):
        measure = self.measures['SPO Fully Qualified Migrating GiB']['expression']
        self.assertIn("SUM('SPO Migration Collection'[Storage Bytes])", measure)
        self.assertIn('[Is Fully Migrating]=TRUE()', measure)
        self.assertIn('Shared folders', measure)
        self.assertNotIn("SUM('SPO Migration Plan'", measure)
        self.assertIn('Not measurable', self.measures['SPO Volume Context']['expression'])

    def test_imported_fields_remain_hidden_and_resolved_target_is_clickable(self):
        cols = {c['Name']: c for c in self.tables['SPO Migration Plan']['Columns']}
        imported = [c for c in cols.values() if c['Name'].startswith('Imported ')]
        self.assertEqual(len(imported), 5)
        self.assertTrue(all(c['IsHidden'] for c in imported))
        self.assertEqual(cols['Resolved Target URL']['DataCategory'], 'WebUrl')
        self.assertNotIn('OnPremMatchStatus', cols)
        self.assertNotIn('EmptyFolder', cols)

    def test_measure_references_and_new_table_columns_exist(self):
        self.assertEqual(len(self.measures), len(self.defs['measures']))
        for m in self.defs['measures']:
            self.assertTrue(m['formatString'])
            self.assertTrue(m['description'])
            for table, column in re.findall(r"'([^']+)'\[([^]]+)\]", m['expression']):
                if table in self.tables:
                    self.assertIn(column, [c['Name'] for c in self.tables[table]['Columns']], (m['name'], table, column))
            for name in re.findall(r'(?<!\w)\[(SPO [^]]+)\]', m['expression']):
                self.assertIn(name, self.measures, m['name'])


if __name__ == '__main__':
    unittest.main()
