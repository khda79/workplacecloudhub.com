"""Offline navigation-contract tests for the compact V1 Power BI cockpit."""

import importlib.util
from pathlib import Path
import sys
import unittest


PRODUCT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PRODUCT / "PowerBI"))
spec = importlib.util.spec_from_file_location(
    "cmdb_report_cockpit", PRODUCT / "PowerBI/report_cockpit.py"
)
cockpit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cockpit)


class CockpitNavigationTests(unittest.TestCase):
    def test_decision_pages_precede_three_drillthrough_capable_detail_pages(self):
        self.assertEqual(len(cockpit.PAGE_ORDER), 10)
        self.assertEqual(
            cockpit.PAGE_ORDER[:7],
            ["overview", "risk", "lifecycle", "licenses", "devices", "users", "businessservices"],
        )
        self.assertEqual(set(cockpit.PAGE_ORDER[7:]), cockpit.DRILLTHROUGH_PAGES)

    def test_display_names_are_contiguous_and_match_page_order(self):
        self.assertEqual(set(cockpit.DISPLAY_NAMES), set(cockpit.PAGE_ORDER))
        for index, page_id in enumerate(cockpit.PAGE_ORDER, start=1):
            self.assertTrue(cockpit.DISPLAY_NAMES[page_id].startswith(f"{index:02d}  "))

    def test_retired_pages_do_not_overlap_the_compact_navigation(self):
        self.assertTrue(cockpit.RETIRED_PAGES.isdisjoint(cockpit.PAGE_ORDER))
        self.assertEqual(
            cockpit.RETIRED_PAGES,
            {"quality", "transformation", "impact", "mailboxes", "hardwarefleet", "hardware"},
        )

    def test_fleet_inventory_keeps_one_device_grain(self):
        self.assertTrue(cockpit.FLEET_TABLE_FIELDS)
        self.assertEqual(
            {table for table, _field, _label, _is_measure in cockpit.FLEET_TABLE_FIELDS},
            {"DimDevice"},
        )
        self.assertIn(
            ("DimDevice", "DeviceSelection", "Open Device 360", False),
            cockpit.FLEET_TABLE_FIELDS,
        )

    def test_executive_license_summary_keeps_five_distinct_products(self):
        self.assertEqual(
            cockpit.LICENSE_SUMMARY_SKUS,
            [
                ("Microsoft 365 F1", "M365_F1"),
                ("Microsoft 365 F3", "SPE_F1"),
                ("Microsoft 365 E3", "SPE_E3"),
                ("Microsoft 365 E5", "SPE_E5"),
                ("Microsoft 365 Copilot", "Microsoft_365_Copilot"),
            ],
        )

    def test_page_navigation_targets_workplace_health(self):
        visual = {"visual": {}}
        cockpit.set_page_navigation(visual, "risk", "Open warnings")
        properties = visual["visual"]["visualContainerObjects"]["visualLink"][0]["properties"]
        self.assertEqual(properties["type"]["expr"]["Literal"]["Value"], "'PageNavigation'")
        self.assertEqual(properties["navigationSection"]["expr"]["Literal"]["Value"], "'risk'")

    def test_mailbox_hosting_reconciliation_prioritizes_online_and_hashes_identity(self):
        identity = {
            "TenantKey": "tenant-key",
            "OrganizationKey": "org-key",
            "EnvironmentKey": "env-key",
            "TenantId": "tenant-id",
        }
        fact = [{"TenantUserKey": "user-1", "PrimarySmtpAddress": "Alice@example.test", "RecipientTypeDetails": "UserMailbox"}]
        remote = [{"PrimarySmtpAddress": "bob@example.test", "RecipientTypeDetails": "RemoteUserMailbox"}]
        local = [
            {"PrimarySmtpAddress": "alice@example.test", "RecipientType": "UserMailbox"},
            {"PrimarySmtpAddress": "carol@example.test", "RecipientType": "SharedMailbox"},
        ]
        rows = cockpit.reconcile_mailboxes(
            fact, remote, local, identity,
            {"user-1": "France"}, {"alice@example.test": "France"},
        )
        self.assertEqual(sum(row["HostingLocation"] == "Exchange Online" for row in rows), 2)
        self.assertEqual(sum(row["HostingLocation"] == "Exchange On-premises" for row in rows), 1)
        self.assertEqual(len(rows), 3)
        self.assertNotIn("PrimarySmtpAddress", rows[0])
        self.assertTrue(all(len(row["MailboxHostingKey"]) == 64 for row in rows))


if __name__ == "__main__":
    unittest.main()
