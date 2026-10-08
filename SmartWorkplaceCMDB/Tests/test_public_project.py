"""Offline checks of the actual generic project and its privacy boundaries."""
import copy
import json
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import uuid

PRODUCT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PRODUCT / 'Scripts'))
import check_public_project as check


class PublicProjectTests(unittest.TestCase):
    def test_actual_generic_project_is_complete_and_neutral(self):
        errors, summary = check.validate(PRODUCT / 'PowerBI/Project')
        self.assertEqual(errors, [])
        self.assertEqual(summary, dict(Pages=13, Tables=103, Measures=249, Relationships=80))

    def setUp(self):
        self.temp = Path(tempfile.gettempdir()) / ('cmdb-public-' + uuid.uuid4().hex)
        self.temp.mkdir()
        self.addCleanup(shutil.rmtree, self.temp)
        self.project = self.temp / 'Project'
        shutil.copytree(PRODUCT / 'PowerBI/Project', self.project)
        self.model_file = self.project / 'SmartWorkplaceCMDB.SemanticModel/model.bim'

    def change_model(self, transform):
        document = check.read_json(self.model_file)
        transform(document['model'])
        self.model_file.write_text(json.dumps(document), encoding='utf-8')

    def test_non_neutral_token_rejected(self):
        def change(model):
            next(x for x in model['expressions'] if x['name'] == 'CMDBReadToken')['expression'] = '"test-only-secret"'
        self.change_model(change)
        self.assertTrue(any('Non-neutral' in e for e in check.validate(self.project)[0]))

    def test_private_source_path_rejected(self):
        self.change_model(lambda m: m.update(description='C:\\Private\\Data'))
        self.assertTrue(any('Private source' in e for e in check.validate(self.project)[0]))

    def test_private_cache_rejected(self):
        cache = self.project / 'SmartWorkplaceCMDB.SemanticModel/.pbi'
        cache.mkdir()
        (cache / 'cache.abf').write_bytes(b'synthetic')
        self.assertTrue(any('artifact' in e for e in check.validate(self.project)[0]))

    def test_broken_relationship_rejected(self):
        self.change_model(lambda m: m['relationships'][0].update(fromColumn='MissingSyntheticColumn'))
        self.assertTrue(any('relationship endpoint' in e for e in check.validate(self.project)[0]))

    def test_fabric_binding_rejected(self):
        file = self.project / 'SmartWorkplaceCMDB.Report/definition.pbir'
        data = check.read_json(file)
        data['datasetReference'] = {'byConnection': {'connectionString': 'synthetic'}}
        file.write_text(json.dumps(data), encoding='utf-8')
        self.assertTrue(any('relative local' in e for e in check.validate(self.project)[0]))

    def test_embedded_partition_rejected(self):
        self.change_model(lambda m: m['tables'][0]['partitions'][0].update(source={'type': 'm', 'expression': 'Binary.Decompress(...)'}))
        self.assertTrue(any('embedded' in e for e in check.validate(self.project)[0]))

    def test_entity_filter_rejected(self):
        file = self.project / 'SmartWorkplaceCMDB.Report/definition/pages/overview/page.json'
        data = check.read_json(file)
        data['filterConfig'] = {'filters': [{'filter': {'Column': {'Property': 'TenantUserKey'},
                            'Literal': {'Value': "'synthetic-private-user'"}}}]}
        file.write_text(json.dumps(data), encoding='utf-8')
        self.assertTrue(any('entity selection' in e for e in check.validate(self.project)[0]))


if __name__ == '__main__':
    unittest.main()
