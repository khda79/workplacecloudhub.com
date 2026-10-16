import csv
import json
import os
import sys
import tempfile
import unittest
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from new_daily_trends_evidence import IDENTITY, SOURCES, generate_source, read_source, run


class DailyTrendsEvidenceTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("SWI_TEST_TMPDIR", Path(__file__).resolve().parents[1] / "_private"))
        temp_root.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=temp_root)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.raw = self.root / "raw"
        self.out = self.root / "staging"
        self.identity = ("tenant-example", "org-example", "prod", "00000000-0000-0000-0000-000000000001")

    def write_source(self, source, data):
        spec = SOURCES[source]
        path = self.raw / "DATA-ALL" / spec["file"]
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=[*IDENTITY, "Date", "DomainName", *spec["metrics"]], delimiter=";")
            writer.writeheader()
            for day, domain, values, hour in data:
                row = dict(zip(IDENTITY, self.identity))
                row.update({"Date": f"{day}T{hour:02d}:00:00", "DomainName": domain})
                row.update({metric: values.get(metric, "0") for metric in spec["metrics"]})
                writer.writerow(row)

    def test_latest_sum_blank_gap_and_stable_keys(self):
        self.write_source("AD Users", [
            ("2026-01-01", "a.contoso.example", {"TotalUsers": "10", "EnabledUsers": "8"}, 8),
            ("2026-01-01", "a.contoso.example", {"TotalUsers": "12", "EnabledUsers": "9"}, 9),
            ("2026-01-01", "b.contoso.example", {"TotalUsers": "3", "EnabledUsers": "2", "Inactive90DaysUsers": ""}, 8),
            ("2026-01-03", "a.contoso.example", {"TotalUsers": "14"}, 8),
        ])
        identity, days = read_source(self.raw / "DATA-ALL" / SOURCES["AD Users"]["file"], SOURCES["AD Users"])
        rows = generate_source("AD Users", SOURCES["AD Users"], identity, days, Decimal("1"), 7, {"manualDays": [], "persistentDomains": []})
        tenant = [r for r in rows if r["Domain Key"] == "__all__"]
        self.assertEqual({r["Date"] for r in tenant}, {"2026-01-01", "2026-01-03"})
        self.assertEqual(next(r["Metric Value"] for r in tenant if r["Date"] == "2026-01-01" and r["Metric Name"] == "TotalUsers"), "15")
        self.assertEqual(next(r["Metric Status"] for r in tenant if r["Metric Name"] == "Inactive90DaysUsers"), "Not measured")
        self.assertEqual(len({r["Domain Key"] for r in rows if r["Domain Name"] == "a.contoso.example"}), 1)

    def test_partial_persistence_and_manual_day(self):
        base = date(2026, 1, 1)
        data = []
        for index in range(9):
            day = (base + timedelta(days=index)).isoformat()
            data.append((day, "a.contoso.example", {"TotalMailboxCount": "90"}, 8))
            if index == 0:
                data.append((day, "retired.contoso.example", {"TotalMailboxCount": "10"}, 8))
        self.write_source("Exchange On-premises", data)
        identity, days = read_source(self.raw / "DATA-ALL" / SOURCES["Exchange On-premises"]["file"], SOURCES["Exchange On-premises"])
        rows = generate_source("Exchange On-premises", SOURCES["Exchange On-premises"], identity, days, Decimal("1"), 7, {"manualDays": [{"source": "Exchange On-premises", "date": "2026-01-09", "reason": "Synthetic capture issue"}], "persistentDomains": []})
        tenant = {r["Date"]: r for r in rows if r["Domain Key"] == "__all__" and r["Metric Name"] == "TotalMailboxCount"}
        self.assertEqual(tenant["2026-01-02"]["Source Day Status"], "Partial capture")
        self.assertEqual(tenant["2026-01-07"]["Source Day Status"], "Partial capture")
        self.assertEqual(tenant["2026-01-08"]["Quality Event"], "Domain absent (persistent)")
        self.assertEqual(tenant["2026-01-08"]["Source Day Status"], "Complete")
        self.assertEqual(tenant["2026-01-09"]["Source Day Status"], "Partial capture")
        self.assertEqual(tenant["2026-01-09"]["Quality Reason"], "Synthetic capture issue")

    def test_small_absent_domain_is_not_unmeasured_or_partial(self):
        self.write_source("AD Computers", [
            ("2026-01-01", "a.contoso.example", {"TotalComputers": "10000", "EnabledAccounts": "8000"}, 8),
            ("2026-01-01", "small.contoso.example", {"TotalComputers": "1", "EnabledAccounts": "0"}, 8),
            ("2026-01-02", "a.contoso.example", {"TotalComputers": "10010", "EnabledAccounts": "8010"}, 8),
        ])
        identity, days = read_source(self.raw / "DATA-ALL" / SOURCES["AD Computers"]["file"], SOURCES["AD Computers"])
        rows = generate_source("AD Computers", SOURCES["AD Computers"], identity, days, Decimal("1"), 7, {"manualDays": [], "persistentDomains": []})
        point = next(r for r in rows if r["Domain Key"] == "__all__" and r["Date"] == "2026-01-02" and r["Metric Name"] == "EnabledAccounts")
        self.assertEqual(point["Source Day Status"], "Complete")
        self.assertEqual(point["Metric Status"], "Measured")
        self.assertEqual(point["Metric Value"], "8010")
        self.assertEqual(point["Absent Domain Count"], 1)

    def test_tenant_mismatch_rejected_without_outputs(self):
        for source in SOURCES:
            self.write_source(source, [("2026-01-01", "a.contoso.example", {}, 8)])
        path = self.raw / "DATA-ALL" / SOURCES["AD Users"]["file"]
        text = path.read_text(encoding="utf-8").replace("tenant-example", "other-tenant")
        path.write_text(text, encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "same|identity"):
            run(self.raw, self.out, Decimal("1"), 7, None)
        self.assertFalse(self.out.exists())


if __name__ == "__main__":
    unittest.main()
