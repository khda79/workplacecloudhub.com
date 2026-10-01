"""Offline checks for the private current-source report preparation guard."""

import csv
import json
from pathlib import Path
import sys
import tempfile
import unittest

POWERBI = Path(__file__).resolve().parents[1] / "PowerBI"
sys.path.insert(0, str(POWERBI))
import prepare_current_report_data as current  # noqa: E402


class CurrentReportDataTests(unittest.TestCase):
    def test_source_hash_path_handles_collection_and_orchestration_roots(self):
        source = Path("C:/Example/DATA-LAST")
        self.assertEqual(current.source_hash_path(source, "PowerBI/DimDevice.csv"),
                         source / "PowerBI/DimDevice.csv")
        self.assertEqual(current.source_hash_path(source, "DATA-LAST/CMDB/CMDB_BuildManifest.csv"),
                         source / "CMDB/CMDB_BuildManifest.csv")
        self.assertEqual(current.source_hash_path(source, "LOG-ALL/run.csv"),
                         source.parent / "LOG-ALL/run.csv")

    def test_hardware_hashes_survive_report_folder_promotion(self):
        with tempfile.TemporaryDirectory(prefix="cmdb-report-path-") as temporary:
            old = Path(temporary) / "Report.stage"
            new = Path(temporary) / "Report"
            old.mkdir()
            key = current.hardware_hash_key(old / "CMDB_CIDeviceHardware.csv", old)
            self.assertEqual(key, "@report/CMDB_CIDeviceHardware.csv")
            self.assertEqual(current.hardware_hash_path(key, new),
                             new / "CMDB_CIDeviceHardware.csv")

    def test_manifest_sidecar_uses_promoted_report_directory(self):
        with tempfile.TemporaryDirectory(prefix="cmdb-report-sidecar-") as temporary:
            root = Path(temporary)
            stage = root / "Report.stage.test"
            final = root / "Report"
            self.assertEqual(current.report_sidecar_path(stage, final), final)
            self.assertEqual(current.report_sidecar_path(final), final)
            with self.assertRaisesRegex(ValueError, "beside the staging directory"):
                current.report_sidecar_path(stage, root / "other" / "Report")

    def test_validate_current_rejects_changed_collection_source(self):
        with tempfile.TemporaryDirectory(prefix="cmdb-current-report-") as temporary:
            root = Path(temporary)
            source = root / "DATA-LAST"
            output = root / "Derived"
            (source / "CMDB").mkdir(parents=True)
            (source / "PowerBI").mkdir()
            output.mkdir()
            identity = dict(TenantKey="fictional-prod", OrganizationKey="fictional",
                            EnvironmentKey="prod", TenantId="11111111-1111-1111-1111-111111111111")
            build = source / "CMDB" / "CMDB_BuildManifest.csv"
            with build.open("w", newline="", encoding="utf-8") as stream:
                writer = csv.DictWriter(stream, fieldnames=[*identity, "BuildDateTime"])
                writer.writeheader()
                writer.writerow(dict(identity, BuildDateTime="2026-01-01T00:00:00Z"))
            input_csv = source / "PowerBI" / "DimDevice.csv"
            input_csv.write_text("Key\nfictional\n", encoding="utf-8")
            derived_csv = output / "Derived.csv"
            derived_csv.write_text("Key\nfictional\n", encoding="utf-8")
            exchange = root / "exchange.csv"
            exchange.write_text("PrimarySmtpAddress\nfictional@example.invalid\n", encoding="utf-8")
            hardware = root / "hardware.csv"
            hardware.write_text("Key\nfictional\n", encoding="utf-8")
            manifest = {
                "status": "Prepared", "sourceRoot": str(source),
                "sourceBuildDateTime": "2026-01-01T00:00:00Z", "identity": identity,
                "outputHashes": {"Derived.csv": current.base.sha(derived_csv)},
                "sourceHashes": {"PowerBI/DimDevice.csv": current.base.sha(input_csv)},
                "mailboxEvidence": {
                    "localPath": str(exchange), "localSha256": current.base.sha(exchange),
                    "remotePath": str(exchange), "remoteSha256": current.base.sha(exchange),
                },
                "hardwareInputHashes": {str(hardware): current.base.sha(hardware)},
            }
            (output / "report-data.manifest.json.txt").write_text(json.dumps(manifest), encoding="utf-8")
            self.assertEqual(current.validate_current(source, output)["status"], "Current")
            input_csv.write_text("Key\nchanged\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Collection source changed"):
                current.validate_current(source, output)


if __name__ == "__main__":
    unittest.main()
