"""Offline report regression tests: synthetic data only; no Power BI or network."""
import contextlib
import csv
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest

PRODUCT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("cmdb_report", PRODUCT / "PowerBI/build_report.py")
report = importlib.util.module_from_spec(spec)
spec.loader.exec_module(report)


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="cmdb-report-v1-test-")
        self.base = Path(self.tmp.name)
        self.root = self.base / "DATA-LAST"
        self.identity = dict(TenantKey="fictional-prod", OrganizationKey="fictional", EnvironmentKey="prod", TenantId="11111111-1111-1111-1111-111111111111")
        self.contract = {Path(t["name"]).stem: t for t in json.loads((PRODUCT / "Schema/SmartWorkplaceCMDB.tables.json").read_text())["tables"]}
        rows = {
            "DimTenant": [dict(TenantDisplayName="FICTIONAL TEST", Environment="test")],
            "DimUser": [dict(TenantUserKey="fictional-prod|u1", CmdbUserId="fictional-prod|u1", AccountEnabled="true"), dict(TenantUserKey="fictional-prod|u2", CmdbUserId="fictional-prod|u2", AccountEnabled="false")],
            "DimDevice": [dict(TenantDeviceKey="fictional-prod|d1", ComplianceState="compliant"), dict(TenantDeviceKey="fictional-prod|d2", ComplianceState="")],
            "DimLicenseSku": [dict(TenantSkuKey="fictional-prod|s1", SkuPartNumber="SKU-A", EnabledUnits="20", ConsumedUnits="1"), dict(TenantSkuKey="fictional-prod|s2", SkuPartNumber="SKU-B", EnabledUnits="30", ConsumedUnits="0")],
            "FactUserLicense": [dict(TenantUserKey="fictional-prod|u1", TenantSkuKey="fictional-prod|s1", AssignmentState="Active")],
            "FactMailbox": [dict(TenantMailboxKey="fictional-prod|m1", CmdbMailboxId="fictional-prod|m1", RecipientTypeDetails="DiscoveryMailbox")],
            "CMDB_Mailboxes": [dict(CmdbMailboxId="fictional-prod|m1", RecipientTypeDetails="DiscoveryMailbox", DisplayName="Fictional technical mailbox")],
            "FactDataQuality": [dict(TenantFindingKey="fictional-prod|q1", FindingId="q1", Severity="Information", FindingType="TechnicalMailboxWithoutUser")],
            "CMDB_DataQuality": [dict(FindingId="q1", Severity="Information", FindingType="TechnicalMailboxWithoutUser", Description="Fictional finding", RecommendedAction="Review locally")],
        }
        for name, t in self.contract.items():
            if t["area"] == "PowerBI" or name in ("CMDB_Mailboxes", "CMDB_DataQuality"):
                self.write(name, rows.get(name, []))

    def tearDown(self):
        self.tmp.cleanup()

    def path(self, name):
        t = self.contract[name]
        return self.root / t["area"] / t["name"]

    def write(self, name, rows):
        t = self.contract[name]
        path = self.path(name)
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=t["columns"])
            writer.writeheader()
            for row in rows:
                r = {c: "" for c in t["columns"]}
                if t.get("tenantScoped", True):
                    r.update(self.identity)
                r.update(row)
                writer.writerow(r)

    def write_first_raw_source(self):
        t = json.loads((PRODUCT / "Schema/SmartWorkplaceCMDB.raw.tables.json").read_text())["tables"][0]
        path = self.root.joinpath(*t["area"].replace("\\", "/").split("/"), t["name"])
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=t["columns"])
            writer.writeheader()
            writer.writerow(dict({c: "" for c in t["columns"]}, **self.identity))
        return path

    def write_snapshot_run(self, build_time="2026-09-13T04:45:22.1961983+02:00",
                           run_end="2026-09-13T02:45:22.4359793+00:00"):
        manifest = self.root / "CMDB/CMDB_BuildManifest.csv"
        manifest.parent.mkdir(parents=True, exist_ok=True)
        manifest_fields = list(self.identity) + ["BuildDateTime"]
        with manifest.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=manifest_fields)
            writer.writeheader()
            writer.writerow(dict(self.identity, BuildDateTime=build_time))
        run_path = self.root.parent / "LOG-ALL/Orchestration/Runs/run.csv"
        run_path.parent.mkdir(parents=True, exist_ok=True)
        fields = ["RunId", "Pipeline", "Mode", "Step", "Status", "StartedDateTime", "EndedDateTime"]
        rows = [
            dict(RunId="synthetic", Pipeline="Full", Mode="Collect", Step="Active Directory collection", Status="Completed",
                 StartedDateTime="2026-09-13T02:44:00+00:00", EndedDateTime="2026-09-13T02:44:30+00:00"),
            dict(RunId="synthetic", Pipeline="Full", Mode="Collect", Step="Contract build and manifest", Status="Completed",
                 StartedDateTime="2026-09-13T02:45:21+00:00", EndedDateTime=run_end),
        ]
        with run_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)
        return run_path

    def test_six_pages_have_working_bindings_and_keep_source_immutable(self):
        hashes = {str(p): report.sha(p) for p in self.root.rglob("*.csv")}
        output = self.base / "report"
        with contextlib.redirect_stdout(io.StringIO()):
            report.build(self.root, output)
        manifest = json.loads((output / "REPORT-MANIFEST.json").read_text(encoding="utf-8"))
        self.assertTrue((output / "SmartWorkplaceCMDB.pbip").is_file())
        self.assertFalse((output / "CMDB-REPORTS.pbip").exists())
        self.assertEqual(len(manifest["pages"]), 6)
        self.assertGreater(manifest["visualCount"], 60)
        self.assertEqual(manifest["version"], "1.0.0-local-report")
        self.assertEqual(hashes, {str(p): report.sha(p) for p in self.root.rglob("*.csv")})
        model = json.loads((output / "CMDB-REPORTS.SemanticModel/model.bim").read_text(encoding="utf-8"))
        self.assertEqual(len(model["model"]["relationships"]), 6)
        self.assertTrue(all(r["crossFilteringBehavior"] == "oneDirection" for r in model["model"]["relationships"]))
        self.assertFalse(any(r.get("toTable") == "SourceHealth" for r in model["model"]["relationships"]))
        for t in model["model"]["tables"]:
            query = "\n".join(t["partitions"][0]["source"]["expression"])
            self.assertIn("Unexpected tenant identity", query)
            self.assertNotIn("Web.Contents", query)
        with self.assertRaisesRegex(ValueError, "new directory"):
            report.build(self.root, output)

    def test_null_conformity_is_not_noncompliant(self):
        _, data, _, _ = report.prepare_data(self.root)
        ms = {m["name"]: m["expected"] for m in report.measures(data)}
        self.assertEqual(ms["Compliant devices"], 1)
        self.assertEqual(ms["Noncompliant devices"], 0)
        self.assertEqual(ms["Compliant device share"], 0.5)
        self.assertEqual(ms["Warnings"], 0)
        self.assertEqual(ms["Information"], 1)
        self.assertEqual(ms["Integrity issues"], 0)
        self.assertEqual(ms["Coverage gaps"], 0)
        self.assertEqual(ms["Derived country gaps"], 0)
        self.assertIsNone(ms["SKU consumed units"])
        self.assertTrue(all(r["Status"] == "Not collected" and r["SourceRows"] == "" for r in data["SourceHealth"]))

    def test_cross_tenant_row_rejected_before_output(self):
        self.write("DimUser", [dict(TenantUserKey="foreign|u1", TenantKey="foreign")])
        with self.assertRaisesRegex(ValueError, "Tenant identity mismatch"):
            report.build(self.root, self.base / "report")
        self.assertFalse((self.base / "report").exists())

    def test_quality_indicators_separate_integrity_coverage_and_derived_gaps(self):
        findings = [
            dict(TenantFindingKey="fictional-prod|q1", FindingId="q1", Severity="Warning", FindingType="OrphanPrimaryUserReference"),
            dict(TenantFindingKey="fictional-prod|q2", FindingId="q2", Severity="Warning", FindingType="ObservedLicenseAssignmentError"),
            dict(TenantFindingKey="fictional-prod|q3", FindingId="q3", Severity="Warning", FindingType="UserCountryUnknown"),
            dict(TenantFindingKey="fictional-prod|q4", FindingId="q4", Severity="Warning", FindingType="DeviceWithoutPrimaryUser"),
            dict(TenantFindingKey="fictional-prod|q5", FindingId="q5", Severity="Warning", FindingType="DeviceCountryUnknown"),
        ]
        self.write("FactDataQuality", findings)
        self.write(
            "CMDB_DataQuality",
            [{key: value for key, value in finding.items() if key != "TenantFindingKey"} for finding in findings],
        )
        _, data, _, _ = report.prepare_data(self.root)
        ms = {m["name"]: m for m in report.measures(data)}
        self.assertEqual(ms["Integrity issues"]["expected"], 2)
        self.assertEqual(ms["Coverage gaps"]["expected"], 2)
        self.assertEqual(ms["Derived country gaps"]["expected"], 1)
        self.assertIn("must not be added", ms["Derived country gaps"]["description"])

    def test_duplicate_case_insensitive_key_rejected(self):
        self.write("DimDevice", [dict(TenantDeviceKey="fictional-prod|A"), dict(TenantDeviceKey="FICTIONAL-PROD|a")])
        with self.assertRaisesRegex(ValueError, "duplicate key"):
            report.prepare_data(self.root)

    def test_orphan_relationship_rejected(self):
        self.write("FactUserLicense", [dict(TenantUserKey="fictional-prod|absent", TenantSkuKey="fictional-prod|s1")])
        with self.assertRaisesRegex(ValueError, "Orphan model relationship"):
            report.prepare_data(self.root)

    def test_duplicate_assignment_rejected(self):
        row = dict(TenantUserKey="fictional-prod|u1", TenantSkuKey="fictional-prod|s1")
        self.write("FactUserLicense", [row, row])
        with self.assertRaisesRegex(ValueError, "duplicate license pair"):
            report.prepare_data(self.root)

    def test_missing_mailbox_details_rejected(self):
        self.write("CMDB_Mailboxes", [])
        with self.assertRaisesRegex(ValueError, "Missing detail row"):
            report.prepare_data(self.root)

    def test_unknown_boolean_rejected(self):
        self.write("DimUser", [dict(TenantUserKey="fictional-prod|u1", AccountEnabled="maybe")])
        with self.assertRaisesRegex(ValueError, "Invalid boolean"):
            report.prepare_data(self.root)

    def test_unknown_account_is_not_disabled(self):
        self.write("DimUser", [dict(TenantUserKey="fictional-prod|u1", AccountEnabled=""), dict(TenantUserKey="fictional-prod|u2", AccountEnabled="false")])
        _, data, _, _ = report.prepare_data(self.root)
        ms = {m["name"]: m for m in report.measures(data)}
        self.assertEqual(ms["Disabled accounts"]["expected"], 1)
        self.assertIn("== FALSE()", ms["Disabled accounts"]["expression"])

    def test_source_hash_mismatch_remains_visible(self):
        t = json.loads((PRODUCT / "Schema/SmartWorkplaceCMDB.raw.tables.json").read_text())["tables"][0]
        path = self.root.joinpath(*t["area"].replace("\\", "/").split("/"), t["name"])
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=t["columns"])
            w.writeheader()
            w.writerow(dict({c: "" for c in t["columns"]}, **self.identity))
        sidecar = path.with_name(path.name + ".status.json")
        sidecar.write_text(json.dumps(dict(self.identity, Status="Completed", Coverage="Complete", SHA256="invalid", RowCount=1)), encoding="utf-8")
        _, data, _, _ = report.prepare_data(self.root)
        self.assertEqual(data["SourceHealth"][0]["Evidence"], "Inconsistent")
        self.assertEqual(data["SourceHealth"][0]["SourceRows"], "1")

    def test_matching_full_run_is_used_as_csv_collection_evidence(self):
        self.write_first_raw_source()
        run_path = self.write_snapshot_run()
        _, data, _, hashes = report.prepare_data(self.root)
        source = data["SourceHealth"][0]
        self.assertEqual(source["Status"], "Completed")
        self.assertEqual(source["Coverage"], "Not reported")
        self.assertEqual(source["Evidence"], "Orchestrator run + CSV")
        self.assertEqual(source["SourceRows"], "1")
        self.assertEqual(source["CompletedDateTime"], "2026-09-13T02:44:30+00:00")
        self.assertIn(str(run_path.relative_to(self.root.parent)), hashes)

    def test_unmatched_full_run_does_not_certify_csv(self):
        self.write_first_raw_source()
        self.write_snapshot_run(run_end="2026-09-13T03:45:22.4359793+00:00")
        _, data, _, _ = report.prepare_data(self.root)
        source = data["SourceHealth"][0]
        self.assertEqual(source["Status"], "CSV without evidence")
        self.assertEqual(source["Evidence"], "Missing")

    def test_sidecar_takes_precedence_over_matching_run(self):
        path = self.write_first_raw_source()
        self.write_snapshot_run()
        sidecar = path.with_name(path.name + ".status.json")
        sidecar.write_text(json.dumps(dict(self.identity, Status="Partial", Coverage="Partial", SHA256=report.sha(path), RowCount=1)), encoding="utf-8")
        _, data, _, _ = report.prepare_data(self.root)
        source = data["SourceHealth"][0]
        self.assertEqual(source["Status"], "Partial")
        self.assertEqual(source["Coverage"], "Partial")
        self.assertEqual(source["Evidence"], "Verified")

    def test_source_evidence_foreign_tenant_rejected(self):
        t = json.loads((PRODUCT / "Schema/SmartWorkplaceCMDB.raw.tables.json").read_text())["tables"][0]
        path = self.root.joinpath(*t["area"].replace("\\", "/").split("/"), t["name"] + ".status.json")
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(dict(self.identity, TenantKey="foreign")), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "evidence tenant mismatch"):
            report.prepare_data(self.root)

    def test_timestamp_requires_timezone(self):
        self.write("DimTenant", [dict(LastRefreshDateTime="2026-09-09T20:00:00")])
        with self.assertRaisesRegex(ValueError, "Missing timestamp timezone"):
            report.prepare_data(self.root)

    def test_header_mismatch_rejected(self):
        self.path("DimDevice").write_text("Unexpected\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "CSV contract mismatch"):
            report.prepare_data(self.root)

    def test_input_subfolder_output_rejected(self):
        with self.assertRaisesRegex(ValueError, "outside the source"):
            report.build(self.root, self.root / "report")


if __name__ == "__main__":
    unittest.main()
