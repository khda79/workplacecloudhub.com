"""The CMDB adapter must preserve the email counts and unavailable states."""

import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "PowerBI"))
import license_report_snapshot as report  # noqa: E402


def sample():
    products = []
    for name in ("F1", "F3", "E3", "E5"):
        available = name != "E5"
        products.append({"Product": "Microsoft 365 " + name, "Enabled": 10,
                         "Consumed": 8, "UsageAvailable": available,
                         "Counts": {"Assigned": 8, "RecoveryCandidates": 1 if name == "F3" else 0,
                                    "RecoveryUnknown": 2, "Disabled": 0,
                                    "M365Inactive": 1 if name == "F3" else 0,
                                    "Multiple": 0, "MultipleAll": 1, "SharedEligible": 0}
                         if available else None})
    return {"SchemaVersion": 1, "SnapshotId": "a" * 32, "TenantKey": "sample",
            "GeneratedAtUtc": "2026-10-07T10:00:00Z", "Products": products,
            "OtherProducts": [{"Product": name, "Enabled": 4, "Consumed": 2}
                              for name in ("Microsoft 365 Copilot", "Dynamics 365", "Power BI")],
            "E3ToF3Review": {"Available": True, "Candidates": 1, "Unknown": 3},
            "RecoveryCandidates": [{"License": "Microsoft 365 F3", "UserId": "u1",
                                    "UserPrincipalName": "user@example.invalid",
                                    "DisplayName": "Example", "RecoveryReason": "No M365 activity in 90 days",
                                    "LastAdActivityDate": "N/D", "LastM365ActivityDate": "2026-01-01",
                                    "PrimaryOnIntuneWindowsPc": "No", "TargetSuites": "Microsoft 365 F3"}],
            "DowngradeCandidates": [{"UserId": "u2", "UserPrincipalName": "user2@example.invalid",
                                     "DisplayName": "Example Two", "MailboxSizeGB": 1.4,
                                     "OneDriveUsedGB": 1.0}],
            "Sources": [{"Name": "M365_Licenses_Users.csv", "Ready": True,
                         "Provisional": True}]}


class LicenseReportSnapshotTests(unittest.TestCase):
    def test_preserves_counts_and_unknowns(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "snapshot.json.txt"
            path.write_text(json.dumps(sample()), encoding="utf-8")
            snapshot = report.load_snapshot(path, "sample")
            summary, candidates = report.flatten(snapshot, [{"SourceUserId": "u1", "TenantUserKey": "sample|u1"}])
            self.assertEqual(len(summary), 7)
            self.assertEqual(next(row for row in summary if row["Product"] == "Microsoft 365 F3")["RecoveryCandidates"], 1)
            self.assertIsNone(next(row for row in summary if row["Product"] == "Microsoft 365 E5")["RecoveryCandidates"])
            self.assertEqual(next(row for row in summary if row["Product"] == "Microsoft 365 E3")["E3ToF3ReviewCandidates"], 1)
            self.assertEqual(candidates[0]["TenantUserKey"], "sample|u1")
            self.assertEqual(candidates[0]["LastAdActivityDate"], "")
            self.assertEqual(candidates[0]["LastM365ActivityDate"], "2026-01-01")
            self.assertEqual(candidates[1]["IdentityJoinStatus"], "Not observed in CMDB snapshot")
            self.assertEqual(summary[1]["EvidenceStatus"], "Provisional")

    def test_gap_unavailable_is_blank(self):
        snapshot = sample()
        snapshot["MailboxGap"] = {"UserMailboxes": {"Available": True, "Total": 5,
                                                    "Universe": 20, "EntraEnabled": None},
                                  "NoUserMailbox": {"Available": False, "Reason": "missing source"}}
        snapshot["AdGapActivity"] = {"Available": False, "Reason": "missing AD"}
        rows = report.flatten_gaps(snapshot)
        self.assertEqual(next(row for row in rows if row["Section"] == "UserMailboxes"
                              and row["Metric"] == "Total")["Value"], 5)
        self.assertIsNone(next(row for row in rows if row["Section"] == "NoUserMailbox"
                               and row["Metric"] == "Total")["Value"])
        self.assertEqual(next(row for row in rows if row["Section"] == "AdGapActivity"
                              and row["Metric"] == "Members")["EvidenceStatus"], "N/D")

    def test_rejects_wrong_tenant_or_detail(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "snapshot.json.txt"
            path.write_text(json.dumps(sample()), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "contract or tenant"):
                report.load_snapshot(path, "other")
            broken = sample()
            broken["RecoveryCandidates"] = []
            path.write_text(json.dumps(broken), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Recovery details differ"):
                report.load_snapshot(path, "sample")


if __name__ == "__main__":
    unittest.main()
