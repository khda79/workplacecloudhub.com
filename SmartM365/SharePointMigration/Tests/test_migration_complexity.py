"""Offline checks for the source-only migration complexity score."""

import csv
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT / "Scripts" / "Compare"))
import migration_complexity as complexity

FILE_FIELDS = ["SiteCollectionUrl", "WebUrl", "LibraryUrl", "FileName", "ServerRelativeUrl",
               "SizeBytes", "VersionsCount", "CheckedOutBy"]
PERMISSION_FIELDS = ["ObjectScope", "ObjectId", "ObjectUrl", "HasUniqueRoleAssignments", "PrincipalType",
                     "PrincipalName", "PrincipalLoginName", "PrincipalMemberLoginNames", "IsLimitedAccessOnly"]


def write_csv(path, fields, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


class MigrationComplexityTests(unittest.TestCase):
    def test_compute_continues_when_automatic_functional_json_is_malformed(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            files, perms = root/'files.csv', root/'permissions.csv'
            write_csv(files, FILE_FIELDS, [])
            write_csv(perms, PERMISSION_FIELDS, [])
            self.write_automatic(files, 'FunctionalComponents', [])
            self.write_automatic(perms, 'CustomPermissions', [])
            Path(str(files)+'.assessment.json.txt').write_text('[]')
            result = complexity.compute(SimpleNamespace(
                config_dir=self.config_dir(root), file_inventory=files, permission_inventory=perms,
                path_mapping_file=None, target_site_url='', entra_users_csv=None,
                source_assessment=None, assessment_target_platform='SPO', migration_name='DEMO'))
            criteria = {row['Id']: row for row in result['Criteria']}
            self.assertFalse(criteria['FunctionalComplexity']['Evaluated'])
            self.assertTrue(criteria['CustomPermissionLevels']['Evaluated'])
            self.assertEqual(criteria['CustomPermissionLevels']['Value'], 0)
            self.assertTrue(criteria['FileCount']['Evaluated'])
            self.assertEqual(result['Status'], 'Partial')

    def test_malformed_automatic_shapes_only_disable_the_affected_criterion(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            files, perms = root/'files.csv', root/'permissions.csv'
            files.write_text('DEMO source files')
            perms.write_text('DEMO source permissions')
            for source, section, criterion, other in (
                    (files, 'FunctionalComponents', 'FunctionalComplexity', 'CustomPermissionLevels'),
                    (perms, 'CustomPermissions', 'CustomPermissionLevels', 'FunctionalComplexity')):
                for case in ('root-array', 'root-null', 'section-array', 'section-null',
                             'inputs-array', 'source-hash-number', 'detail-hash-number',
                             'items-null', 'null-item', 'scalar-item', 'errors-null',
                             'detectors-scalar', 'date-number', 'invalid-json'):
                    with self.subTest(section=section, case=case):
                        self.write_automatic(files, 'FunctionalComponents', [])
                        self.write_automatic(perms, 'CustomPermissions', [])
                        sidecar = Path(str(source)+'.assessment.json.txt')
                        data = json.loads(sidecar.read_text())
                        if case == 'root-array': data = []
                        elif case == 'root-null': data = None
                        elif case == 'section-array': data[section] = []
                        elif case == 'section-null': data[section] = None
                        elif case == 'inputs-array': data['Inputs'] = []
                        elif case == 'source-hash-number':
                            data['Inputs'] = {key: 123 for key in data['Inputs']}
                        elif case == 'detail-hash-number': data['DetailsSha256'] = 123
                        elif case == 'items-null': data[section]['Items'] = None
                        elif case == 'null-item': data[section]['Items'] = [None]
                        elif case == 'scalar-item': data[section]['Items'] = ['bad item']
                        elif case == 'errors-null': data[section]['Errors'] = None
                        elif case == 'detectors-scalar': data[section]['Detectors'] = 'bad detectors'
                        elif case == 'date-number': data[section]['ReviewedAtUtc'] = 123
                        sidecar.write_text('{' if case == 'invalid-json' else json.dumps(data))
                        values, evidence, present, _ = complexity.automatic_assessment_metrics(files, perms)
                        self.assertNotIn(criterion, values)
                        self.assertIn(criterion, present)  # Never enable the manual fallback.
                        self.assertFalse(evidence[criterion]['Valid'])
                        self.assertTrue(evidence[criterion]['Error'])
                        self.assertEqual(values[other][0], 0)

    def write_automatic(self, csv_path, section, items, complete=True, qualified=True):
        detail = Path(str(csv_path)+'.assessment.json.txt.details.csv')
        detail.write_text('DEMO automatic detail',encoding='utf-8')
        key = 'FileInventorySha256' if section == 'FunctionalComponents' else 'PermissionInventorySha256'
        data = {'Complete':complete and qualified,'CollectionComplete':complete,'QualificationComplete':qualified,
                'Mode':'Automatic','ReviewedAtUtc':'2026-01-01T00:00:00Z','Evidence':'DEMO automatic source assessment',
                'Items':items,'Errors':[],'Detectors':[]}
        sidecar=Path(str(csv_path)+'.assessment.json.txt')
        sidecar.write_text(json.dumps({'SchemaVersion':1,'TargetPlatform':'SPO','RulesSha256':'DEMO',
            'Inputs':{key:complexity.sha256_file(csv_path)},'DetailsFile':detail.name,
            'DetailsSha256':complexity.sha256_file(detail),section:data}),encoding='utf-8')
        return sidecar

    def test_automatic_evidence_is_bound_and_never_hides_unknowns(self):
        with self.temporary_directory() as directory:
            root=Path(directory)
            files=root/'files.csv';files.write_text('DEMO source files')
            perms=root/'permissions.csv';perms.write_text('DEMO source permissions')
            self.write_automatic(files,'FunctionalComponents',[
                {'Id':'wf','Category':'Workflow','Effort':3,'Classification':'Estimated','Technology':'Workflow2013','Evidence':'DEMO workflow','Reason':'Replacement'}])
            self.write_automatic(perms,'CustomPermissions',[
                {'IsBuiltIn':False,'IsUsed':True,'EffectivePermissions':'7','Evidence':'DEMO custom role'},
                {'IsBuiltIn':False,'IsUsed':True,'EffectivePermissions':'7','Evidence':'DEMO alias'}])
            values,evidence,present,hashes=complexity.automatic_assessment_metrics(files,perms)
            self.assertEqual(values['FunctionalComplexity'][0],3)
            self.assertEqual(values['CustomPermissionLevels'][0],1)
            self.assertEqual(len(evidence['FunctionalComplexity']['KnownRisks']),1)
            self.assertEqual(present,{'FunctionalComplexity','CustomPermissionLevels'})
            self.assertTrue(hashes['File'])
            values,_,_,_=complexity.automatic_assessment_metrics(files,perms,'Other')
            self.assertNotIn('FunctionalComplexity',values)
            self.write_automatic(files,'FunctionalComponents',[
                {'Id':'wf','Category':'Workflow','Effort':3,'Classification':'Estimated','Technology':'Workflow2013','Evidence':'DEMO workflow'},
                {'Id':'unknown','Category':'Form','Effort':None,'Classification':'Unqualified','Evidence':'DEMO unknown form'}],qualified=False)
            values,evidence,_,_=complexity.automatic_assessment_metrics(files,perms)
            self.assertNotIn('FunctionalComplexity',values)
            self.assertTrue(evidence['FunctionalComplexity']['CollectionComplete'])
            self.assertFalse(evidence['FunctionalComplexity']['QualificationComplete'])
            self.assertEqual(evidence['FunctionalComplexity']['UnqualifiedCount'],1)
            self.assertEqual(len(evidence['FunctionalComplexity']['KnownRisks']),1)
            perms.write_text('changed source')
            values,evidence,_,_=complexity.automatic_assessment_metrics(files,perms)
            self.assertNotIn('CustomPermissionLevels',values)
            self.assertFalse(evidence['CustomPermissionLevels']['Valid'])

    def test_automatic_detail_corruption_and_empty_assessment(self):
        with self.temporary_directory() as directory:
            files=Path(directory)/'files.csv';files.write_text('DEMO source')
            sidecar=self.write_automatic(files,'FunctionalComponents',[])
            values,_,_,_=complexity.automatic_assessment_metrics(files,None)
            self.assertEqual(values['FunctionalComplexity'][0],0)
            Path(str(sidecar)+'.details.csv').write_text('corrupted detail')
            values,evidence,present,_=complexity.automatic_assessment_metrics(files,None)
            self.assertNotIn('FunctionalComplexity',values)
            self.assertIn('FunctionalComplexity',present)
            self.assertIn('SHA256 mismatch',evidence['FunctionalComplexity']['Error'])

    @staticmethod
    def temporary_directory():
        return tempfile.TemporaryDirectory(prefix=".migration-complexity-test-", dir=Path(__file__).parent)

    def config_dir(self, root):
        config = root / "Config"
        config.mkdir()
        shutil.copyfile(PROJECT / "Config" / "migration-complexity.json.template",
                        config / "migration-complexity.json.template")
        return config

    def test_boundaries_belong_to_the_lower_band(self):
        criterion = {"OnePointFrom": 10, "TwoPointsAbove": 100, "ThreePointsAbove": 1000}
        self.assertEqual([complexity.points_for(value, criterion) for value in (9, 10, 100, 101, 1000, 1001)],
                         [0, 1, 1, 2, 2, 3])

    def test_runtime_config_self_heals_without_overwriting_local_values(self):
        with self.temporary_directory() as directory:
            config_dir = self.config_dir(Path(directory))
            template = json.loads((config_dir / "migration-complexity.json.template").read_text(encoding="utf-8"))
            local = {"SchemaVersion": 1, "Criteria": [dict(template["Criteria"][0], Weight=40)],
                     "Levels": template["Levels"], "FloorRules": []}
            (config_dir / "migration-complexity.json.txt").write_text(json.dumps(local), encoding="utf-8")
            config, runtime = complexity.load_config(config_dir)
            self.assertEqual(config["Criteria"][0]["Weight"], 40)
            self.assertEqual(len(config["Criteria"]), len(template["Criteria"]))
            self.assertIn("Settings", config)
            self.assertEqual(len(config["FloorRules"]), len(template["FloorRules"]))
            self.assertEqual(json.loads(runtime.read_text(encoding="utf-8")), config)

    def test_floor_rule_raises_level_and_partial_inputs_keep_denominator_honest(self):
        with self.temporary_directory() as directory:
            config, _ = complexity.load_config(self.config_dir(Path(directory)))
            values = {criterion["Id"]: (0, "") for criterion in config["Criteria"]
                      if criterion["Id"] in complexity.FILE_CRITERIA}
            values["PathLimitFiles"] = (500, "")
            values.update(LibraryConcentration=(0, ""), VersionConcentration=(0, ""))
            result = complexity.score_migration(config, values)
            self.assertEqual(result["Score"], 15)
            self.assertEqual(result["ScoreLevel"], "Low")
            self.assertEqual(result["Level"], "High")
            self.assertEqual(result["FloorApplied"], ["PathLimitFiles"])
            self.assertEqual(result["EvaluatedWeight"], 65)
            self.assertEqual(result["TopDrivers"], ["PathLimitFiles"])
            inactive = next(row for row in result["Criteria"] if row["Id"] == "InactiveAccounts")
            self.assertFalse(inactive["Evaluated"])

    def test_files_outside_the_path_mapping_are_skipped(self):
        mappings = [("/sites/a", "/sites/target-a")]
        self.assertEqual(complexity.projected_target_path("/sites/a/Docs/x.docx", mappings, "/sites/ignored"),
                         "sites/target-a/Docs/x.docx")
        self.assertEqual(complexity.projected_target_path("/sites/b/Docs/x.docx", mappings, "/sites/ignored"), "")
        self.assertEqual(complexity.projected_target_path("/Docs/x.docx", [], "/sites/t"), "sites/t/Docs/x.docx")

    def test_disabled_criterion_and_floor_rule_are_ignored(self):
        with self.temporary_directory() as directory:
            config, _ = complexity.load_config(self.config_dir(Path(directory)))
            for criterion in config["Criteria"]:
                if criterion["Id"] == "LibraryCount":
                    criterion["Enabled"] = False
            for rule in config["FloorRules"]:
                rule["Enabled"] = False
            values = {criterion["Id"]: (0, "") for criterion in config["Criteria"]
                      if criterion["Id"] in complexity.FILE_CRITERIA}
            values["LibraryCount"] = (500, "")
            values["PathLimitFiles"] = (500, "")
            values.update(LibraryConcentration=(0, ""), VersionConcentration=(0, ""))
            result = complexity.score_migration(config, values)
            library = next(row for row in result["Criteria"] if row["Id"] == "LibraryCount")
            self.assertFalse(library["Evaluated"])
            self.assertEqual(library["ValueDisplay"], "Disabled")
            self.assertEqual(result["EvaluatedWeight"], 60)
            self.assertEqual(result["FloorApplied"], [])
            self.assertEqual(result["Level"], "Low")

    def test_end_to_end_metrics_from_source_scans(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            config_dir = self.config_dir(root)
            long_name = "x" * 260 + ".docx"
            files = root / "SP2019-FileInventory-Fixture-20261001-120000.csv"
            write_csv(files, FILE_FIELDS, [
                dict(WebUrl="https://source.test", LibraryUrl="/Docs", FileName="a.docx",
                     ServerRelativeUrl="/Docs/a.docx", SizeBytes=3 * 1024 ** 3, VersionsCount=4),
                dict(WebUrl="https://source.test/sub", LibraryUrl="/sub/Docs", FileName="b.docx",
                     ServerRelativeUrl="/sub/Docs/b.docx", SizeBytes=10, VersionsCount=0, CheckedOutBy="someone"),
                dict(WebUrl="https://source.test/sub", LibraryUrl="/sub/Docs", FileName="Thumbs.db",
                     ServerRelativeUrl="/sub/Docs/Thumbs.db", SizeBytes=10, VersionsCount=0),
                dict(WebUrl="https://source.test/sub", LibraryUrl="/sub/Docs", FileName=long_name,
                     ServerRelativeUrl=f"/sub/Docs/{long_name}", SizeBytes=10, VersionsCount=2),
            ])
            permissions = root / "SP2019-PermissionInventory-Fixture-20261001-120000.csv"
            write_csv(permissions, PERMISSION_FIELDS, [
                dict(ObjectScope="Item", ObjectId="1", HasUniqueRoleAssignments="True", PrincipalType="User",
                     PrincipalLoginName="i:0e.t|adfs|active@contoso.test"),
                dict(ObjectScope="Item", ObjectId="1", HasUniqueRoleAssignments="True", PrincipalType="User",
                     PrincipalLoginName="i:0e.t|adfs|disabled@contoso.test"),
                dict(ObjectScope="List", ObjectId="2", HasUniqueRoleAssignments="True", PrincipalType="User",
                     PrincipalLoginName="i:0e.t|adfs|limited@contoso.test", IsLimitedAccessOnly="True"),
                dict(ObjectScope="Web", ObjectId="3", HasUniqueRoleAssignments="True", PrincipalType="SharePointGroup",
                     PrincipalLoginName="Owners",
                     PrincipalMemberLoginNames="i:0e.t|adfs|ghost@contoso.test || i:0#.w|sharepoint\\system"),
            ])
            entra = root / "entra.csv"
            write_csv(entra, ["UserPrincipalName", "AccountEnabled"], [
                dict(UserPrincipalName="active@contoso.test", AccountEnabled="True"),
                dict(UserPrincipalName="disabled@contoso.test", AccountEnabled="False"),
                dict(UserPrincipalName="limited@contoso.test", AccountEnabled="True"),
            ])
            output = root / "complexity.json.txt"
            complexity.main(["--migration-name", "Fixture", "--file-inventory", str(files),
                             "--permission-inventory", str(permissions), "--entra-users-csv", str(entra),
                             "--target-site-url", "https://target.test/sites/fixture",
                             "--config-dir", str(config_dir), "--output", str(output)])
            result = json.loads(output.read_text(encoding="utf-8"))
            values = {row["Id"]: row["Value"] for row in result["Criteria"]}
            self.assertEqual(result["Status"], "Partial")
            self.assertEqual(result["CoveragePercent"], 85)
            self.assertEqual(values["FileCount"], 4)
            self.assertEqual(values["SubsiteCount"], 2)
            self.assertEqual(values["LibraryCount"], 2)
            self.assertEqual(values["AverageVersions"], 1.5)
            self.assertEqual(values["LargeFiles"], 1)
            self.assertEqual(values["RiskyFiles"], 2)
            self.assertEqual(values["PathLimitFiles"], 1)
            self.assertEqual(values["BrokenInheritance"], 1)
            self.assertEqual(values["DirectUsers"], 2)
            self.assertEqual(values["InactiveAccounts"], 2)
            self.assertEqual(result["Inputs"]["FileInventory"]["Name"], files.name)
            self.assertEqual(len(result["Inputs"]["ConfigSha256"]), 64)

    def test_concentration_boundaries_and_max_avoid_double_counting(self):
        config = json.loads((PROJECT / "Config/migration-complexity.json.template").read_text())
        by_id = {item["Id"]: item for item in config["Criteria"]}
        library = by_id["FileCount"]["Alternatives"][0]
        versions = by_id["AverageVersions"]["Alternatives"][0]
        self.assertEqual([complexity.points_for(x, library) for x in (19999, 20000, 50000, 50001, 100000, 100001)], [0,1,1,2,2,3])
        self.assertEqual([complexity.points_for(x, versions) for x in (0, .001, 1, 1.001, 5, 5.001)], [0,1,1,2,2,3])
        values = {item["Id"]: (0, "") for item in config["Criteria"]}
        values.update(FileCount=(120000,""), LibraryConcentration=(120000,""), AverageVersions=(12,""), VersionConcentration=(6,""))
        scored = complexity.score_migration(config, values)
        rows = {item["Id"]: item for item in scored["Criteria"]}
        self.assertEqual(len(rows), 13)
        self.assertEqual(scored["TotalWeight"], 100)
        self.assertEqual(rows["FileCount"]["WeightedPoints"], 45)
        self.assertEqual(rows["AverageVersions"]["WeightedPoints"], 30)
        self.assertEqual(scored["Score"], 25)
        self.assertEqual(scored["CoveragePercent"], 100)
        values.pop("VersionConcentration")
        missing = complexity.score_migration(config, values)
        self.assertEqual(missing["EvaluatedWeight"], 90)

    def test_revision_upgrades_old_defaults_and_preserves_custom_weights(self):
        with self.temporary_directory() as directory:
            config_dir = self.config_dir(Path(directory))
            old = json.loads((config_dir/'migration-complexity.json.template').read_text())
            old.pop('ScoringRevision')
            for item in old['Criteria']:
                if item['Id'] == 'TotalSizeGiB': item['Weight'] = 15
                if item['Id'] == 'BrokenInheritance': item['Weight'] = 22
                if item['Id'] == 'DirectUsers': item['Weight'] = 10
            (config_dir/'migration-complexity.json.txt').write_text(json.dumps(old))
            loaded, runtime = complexity.load_config(config_dir)
            weights = {item['Id']:item['Weight'] for item in loaded['Criteria']}
            self.assertEqual(weights['TotalSizeGiB'], 10)
            self.assertEqual(weights['DirectUsers'], 5)
            self.assertEqual(weights['BrokenInheritance'], 22)
            self.assertEqual(json.loads(runtime.read_text())['ScoringRevision'], 2)

    def test_demo_fixture_is_calculated_with_current_rules(self):
        config = json.loads((PROJECT/'Config/migration-complexity.json.template').read_text())
        fixture = json.loads((PROJECT/'Config/migration-complexity.demo.json.template').read_text())
        def verify(result):
            values = {row['Id']:(row['Value'],row['Detail']) for row in result['Criteria'] if row['Evaluated']}
            for row in result['Criteria']:
                for component in row['Components']:
                    if component['Value'] is not None: values[component['Id']] = (component['Value'],component['Detail'])
            actual = complexity.score_migration(config, values)
            for key in ('Score','Level','EvaluatedWeight','TotalWeight','CoveragePercent'):
                self.assertEqual(actual[key], result[key])
            self.assertEqual([r['Weight'] for r in actual['Criteria']], [r['Weight'] for r in result['Criteria']])
        verify(fixture); verify(fixture['PartialResult'])

    def test_source_assessment_deduplication_completeness_and_binding(self):
        with self.temporary_directory() as directory:
            root = Path(directory)
            files, permissions = root/'files.csv', root/'permissions.csv'
            files.write_text('source files'); permissions.write_text('source permissions')
            assessment = {
                'SchemaVersion': 1,
                'Inputs': {'FileInventorySha256': complexity.sha256_file(files), 'PermissionInventorySha256': complexity.sha256_file(permissions)},
                'FunctionalComponents': {'Complete': True, 'ReviewedAtUtc': '2026-01-01T00:00:00Z', 'Evidence':'DEMO assessment', 'Items': [
                    {'Id':'wf1','Category':'Workflow','Effort':2,'Evidence':'DEMO workflow'},
                    {'Id':'wf1','Category':'Form','Effort':3,'Evidence':'DEMO dependency'}]},
                'CustomPermissions': {'Complete': True, 'ReviewedAtUtc': '2026-01-01T00:00:00Z', 'Evidence':'DEMO definitions', 'Items': [
                    {'IsBuiltIn':False,'IsUsed':True,'EffectivePermissions':'123','Evidence':'DEMO A'},
                    {'IsBuiltIn':False,'IsUsed':True,'EffectivePermissions':'00123','Evidence':'DEMO alias'},
                    {'IsBuiltIn':True,'IsUsed':True}, {'IsBuiltIn':False,'IsUsed':False}]}}
            path = root/'assessment.json.txt'
            def evaluate():
                path.write_text(json.dumps(assessment))
                return complexity.assessment_metrics(path, files, permissions)
            result = evaluate()
            self.assertEqual(result['CustomPermissionLevels'][0], 1)
            self.assertEqual(result['FunctionalComplexity'][0], 3)
            assessment['FunctionalComponents']['Items'] = []
            self.assertEqual(evaluate()['FunctionalComplexity'][0], 0)
            assessment['FunctionalComponents']['Complete'] = False
            self.assertNotIn('FunctionalComplexity', evaluate())
            permissions.write_text('changed source')
            self.assertNotIn('CustomPermissionLevels', evaluate())
            assessment['Inputs']['PermissionInventorySha256'] = complexity.sha256_file(permissions)
            assessment['CustomPermissions']['Items'][0]['EffectivePermissions'] = 'not a mask'
            with self.assertRaises(ValueError): evaluate()

    def test_functional_and_custom_thresholds(self):
        config = json.loads((PROJECT / "Config/migration-complexity.json.template").read_text())
        by_id = {item['Id']: item for item in config['Criteria']}
        self.assertEqual([complexity.points_for(x, by_id['FunctionalComplexity']) for x in range(4)], [0,1,2,3])
        self.assertEqual([complexity.points_for(x, by_id['CustomPermissionLevels']) for x in (0,1,2,3,5,6)], [0,1,1,2,2,3])

    def test_missing_and_invalid_versions_are_not_zero_risk(self):
        with self.temporary_directory() as directory:
            path = Path(directory)/'files.csv'
            write_csv(path, FILE_FIELDS, [dict(WebUrl='https://source.test', LibraryUrl='/Docs', VersionsCount='unknown')])
            values = complexity.file_metrics(path, {}, [], '')
            self.assertNotIn('AverageVersions', values)
            self.assertNotIn('VersionConcentration', values)
            write_csv(path, FILE_FIELDS, [dict(WebUrl='https://source.test', LibraryUrl='/Docs', VersionsCount=100), dict(WebUrl='https://source.test',LibraryUrl='/Docs',VersionsCount=0)])
            values = complexity.file_metrics(path, {}, [], '')
            self.assertEqual(values['VersionConcentration'][0], 50)
            self.assertEqual(values['LibraryConcentration'][0], 2)


if __name__ == "__main__":
    unittest.main()
