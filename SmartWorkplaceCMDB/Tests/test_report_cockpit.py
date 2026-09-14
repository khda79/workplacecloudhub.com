"""Offline navigation-contract tests for the compact V1 Power BI cockpit."""

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch


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

    def test_all_three_360_pages_are_managed_by_the_repeatable_restyle(self):
        self.assertTrue(cockpit.DRILLTHROUGH_PAGES.issubset(cockpit.MANAGED_PAGES))

    def test_each_visible_page_has_a_registered_header_icon(self):
        self.assertEqual(set(cockpit.PAGE_ICON_RESOURCES), set(cockpit.PAGE_ORDER))
        self.assertTrue(all(name.endswith(".svg") for name in cockpit.PAGE_ICON_RESOURCES.values()))

    def test_each_executive_ratio_card_has_a_distinct_vector_illustration(self):
        self.assertEqual(
            set(cockpit.OVERVIEW_RATIO_ICON_RESOURCES),
            {"formfactor", "ownership", "windows", "accounts", "mailboxes"},
        )
        self.assertEqual(
            len(set(cockpit.OVERVIEW_RATIO_ICON_RESOURCES.values())),
            len(cockpit.OVERVIEW_RATIO_ICON_RESOURCES),
        )
        self.assertTrue(
            all(name.endswith(".svg") for name in cockpit.OVERVIEW_RATIO_ICON_RESOURCES.values())
        )

    def test_kpi_and_detail_section_icons_use_the_same_vector_asset_system(self):
        self.assertEqual(
            set(cockpit.KPI_ICON_RESOURCES),
            {"devices", "management", "compliance", "users", "mailboxes", "licenses", "quality", "analytics", "relationships"},
        )
        self.assertEqual(
            set(cockpit.SECTION_ICON_RESOURCES),
            {"identity", "enrollment", "hardware", "activity", "findings"},
        )
        resources = [*cockpit.KPI_ICON_RESOURCES.values(), *cockpit.SECTION_ICON_RESOURCES.values()]
        self.assertTrue(all(name.endswith(".svg") for name in resources))
        self.assertEqual(len(resources), len(set(resources)))

    def test_copyright_is_exact_and_targets_the_public_website(self):
        self.assertEqual(
            cockpit.COPYRIGHT_TEXT,
            "© 2026 WorkplaceCloudHub — https://workplacecloudhub.com/",
        )
        self.assertEqual(cockpit.COPYRIGHT_URL, "https://workplacecloudhub.com/")

    def test_searchable_slicer_keeps_all_as_the_neutral_state(self):
        visual = {"visual": {"objects": {}}}
        cockpit.configure_searchable_slicer(visual)
        objects = visual["visual"]["objects"]

        self.assertEqual(objects["data"][0]["properties"]["mode"]["expr"]["Literal"]["Value"], "'Dropdown'")
        self.assertEqual(objects["general"][0]["properties"]["selfFilterEnabled"]["expr"]["Literal"]["Value"], "true")
        self.assertNotIn("selection", objects)
        self.assertNotIn("filter", objects["general"][0]["properties"])

    def test_default_device_ownership_page_filter_is_corporate_and_repeat_safe(self):
        page = {
            "filterConfig": {
                "filters": [
                    {"name": "preserved", "type": "Categorical"},
                    {
                        "name": cockpit.DEFAULT_OWNERSHIP_FILTER_NAME,
                        "field": {"Column": {
                            "Expression": {"SourceRef": {"Entity": "DimDevice"}},
                            "Property": "OwnershipLabel",
                        }},
                        "type": "Categorical",
                    },
                ],
                "filterSortOrder": "Custom",
            }
        }
        for _ in range(2):
            cockpit.set_page_categorical_filter(
                page,
                cockpit.DEFAULT_OWNERSHIP_FILTER_NAME,
                "DimDevice",
                "OwnershipLabel",
                cockpit.DEFAULT_OWNERSHIP,
            )

        filters = page["filterConfig"]["filters"]
        self.assertEqual([item["name"] for item in filters], ["preserved", "defaultdeviceownership"])
        ownership = filters[1]
        self.assertEqual(ownership["field"]["Column"]["Property"], "OwnershipLabel")
        predicate = ownership["filter"]["Where"][0]["Condition"]["In"]
        self.assertEqual(
            predicate["Expressions"][0]["Column"]["Expression"]["SourceRef"],
            {"Source": "s"},
        )
        self.assertEqual(predicate["Values"][0][0]["Literal"]["Value"], "'Corporate'")

    def test_version_is_centered_in_the_shared_footer(self):
        with tempfile.TemporaryDirectory() as root:
            pages = Path(root)
            path = pages / "risk" / "visuals" / "riskv1" / "visual.json"
            cockpit.write(path, {
                "position": {"x": 24, "y": 60, "width": 600, "height": 24},
                "visual": {"objects": {"general": [{"properties": {
                    "paragraphs": [{"textRuns": [{"value": "VERSION 1.0.0"}]}]
                }}]}},
            })

            cockpit.position_version_in_footer(pages, "risk")
            version = cockpit.load(path)

        self.assertEqual(
            {key: version["position"][key] for key in ("x", "y", "width", "height")},
            {"x": 536, "y": 864, "width": 208, "height": 28},
        )
        paragraph = version["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]
        self.assertEqual(paragraph["horizontalTextAlignment"], "center")

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

    def test_license_country_footprint_keeps_five_distinct_products_and_colors(self):
        self.assertEqual(
            [(title, sku) for title, sku, _measure, _label, _color in cockpit.LICENSE_COUNTRY_SERIES],
            cockpit.LICENSE_SUMMARY_SKUS,
        )
        self.assertEqual(len({color for *_rest, color in cockpit.LICENSE_COUNTRY_SERIES}), 5)
        self.assertEqual(
            [label for _title, _sku, _measure, label, _color in cockpit.LICENSE_COUNTRY_SERIES],
            ["F1", "F3", "E3", "E5", "Copilot"],
        )

    def test_executive_country_footprints_use_the_compact_country_label(self):
        template = {
            "name": "template",
            "visual": {
                "objects": {"labels": [{"properties": {}}]},
                "visualContainerObjects": {"title": [{"properties": {"text": {}}}]},
            }
        }
        for builder in (cockpit.add_country_bar, cockpit.add_license_country_bar):
            visuals = []
            with patch.object(cockpit, "clone_visual_by_type", return_value=template):
                visual = builder(Path("unused"), visuals, "overview", 0, 0, 100, 100)
            category = visual["visual"]["query"]["queryState"]["Category"]["projections"][0]
            self.assertEqual(
                category["field"]["Column"]["Property"],
                cockpit.COUNTRY_FOOTPRINT_COLUMN,
            )
        self.assertEqual(cockpit.COUNTRY_FOOTPRINT_UNKNOWN, "Unknown")

    def test_country_footprint_label_shortens_only_the_display_value(self):
        table = {"columns": [{"name": "CountryLabel", "dataType": "string"}]}
        cockpit.add_or_replace_calculated_column(
            table,
            cockpit.COUNTRY_FOOTPRINT_COLUMN,
            (
                "IF('DimCountry'[CountryLabel] = \"Unknown / unassigned\", "
                f'\"{cockpit.COUNTRY_FOOTPRINT_UNKNOWN}\", \'DimCountry\'[CountryLabel])'
            ),
            "Compact footprint label.",
        )
        compact = next(
            column for column in table["columns"]
            if column["name"] == cockpit.COUNTRY_FOOTPRINT_COLUMN
        )
        self.assertIn('"Unknown"', compact["expression"])
        self.assertIn("'DimCountry'[CountryLabel]", compact["expression"])

    def test_refreshable_derived_tables_use_the_private_data_root_parameter(self):
        source = (PRODUCT / "PowerBI/report_cockpit.py").read_text(encoding="utf-8")
        self.assertIn('hosting_path = data_dir / "FactMailboxHosting.csv"', source)
        self.assertIn('country_path = data_dir / "DimCountry.csv"', source)
        self.assertIn("report, local_path, remote_path, data_dir", source)
        self.assertIn("File.Contents(CMDBDataRoot &", source)

    def test_explicit_report_data_directory_precedes_parameter_discovery(self):
        with tempfile.TemporaryDirectory() as root:
            root = Path(root)
            report = root / "project" / "SmartWorkplaceCMDB.Report"
            local_data = report.parent / "ReportData"
            explicit_data = root / "explicit"
            parameter_root = root / "parameter"
            parameter_data = parameter_root / "PowerBI" / "CMDB-REPORTS" / "ReportData"
            for path in (local_data, explicit_data, parameter_data):
                path.mkdir(parents=True)
            model = {"expressions": [{
                "name": "CMDBDataRoot",
                "expression": f'"{parameter_root}" meta [IsParameterQuery=true]',
            }]}
            self.assertEqual(
                cockpit.resolve_report_data_dir(report, model, explicit_data),
                explicit_data.resolve(),
            )
            self.assertEqual(
                cockpit.resolve_report_data_dir(report, model),
                parameter_data,
            )
            with self.assertRaisesRegex(ValueError, "Explicit ReportData"):
                cockpit.resolve_report_data_dir(report, model, root / "missing")

    def test_quality_indicators_are_explicit_and_non_additive(self):
        self.assertEqual(
            [name for name, _expression, _description in cockpit.QUALITY_INDICATORS],
            ["Integrity issues", "Coverage gaps", "Derived country gaps"],
        )
        derived = next(item for item in cockpit.QUALITY_INDICATORS if item[0] == "Derived country gaps")
        self.assertIn("must not be added", derived[2])

    def test_series_colors_use_metadata_selectors(self):
        visual = {"visual": {}}
        cockpit.set_series_colors(visual, cockpit.POPULATION_COUNTRY_SERIES)
        entries = visual["visual"]["objects"]["dataPoint"]
        self.assertEqual(
            [entry["selector"]["metadata"] for entry in entries],
            [f"{table}.{measure}" for table, measure, _label, _color in cockpit.POPULATION_COUNTRY_SERIES],
        )
        self.assertTrue(all("fill" in entry["properties"] for entry in entries))

    def test_ratio_category_colors_use_scope_identity_selectors(self):
        visual = {"visual": {}}
        colors = cockpit.EXECUTIVE_RATIO_CATEGORY_COLORS["accounts"]
        cockpit.set_category_colors(visual, "DimUser", "AccountStatusLabel", colors)
        entries = visual["visual"]["objects"]["dataPoint"]

        self.assertIn("defaultColor", entries[0]["properties"])
        self.assertEqual(len(entries), len(colors) + 1)
        selected = {}
        for entry in entries[1:]:
            comparison = entry["selector"]["data"][0]["scopeId"]["Comparison"]
            self.assertEqual(comparison["ComparisonKind"], 0)
            self.assertEqual(comparison["Left"]["Column"]["Expression"]["SourceRef"]["Entity"], "DimUser")
            self.assertEqual(comparison["Left"]["Column"]["Property"], "AccountStatusLabel")
            value = comparison["Right"]["Literal"]["Value"].strip("'")
            selected[value] = entry["properties"]["fill"]["solid"]["color"]["expr"]["Literal"]["Value"].strip("'")

        self.assertEqual(selected, colors)
        self.assertNotEqual(selected["Enabled"], selected["Disabled"])

    def test_each_executive_ratio_category_palette_has_distinct_colors(self):
        for palette in cockpit.EXECUTIVE_RATIO_CATEGORY_COLORS.values():
            self.assertEqual(len(palette), len(set(palette.values())))

    def test_executive_measures_add_quality_groups_and_license_country_ratios(self):
        tables = [
            {"name": "DimCountry", "measures": []},
            {"name": "FactDataQuality", "measures": []},
            {"name": "DimUser", "measures": []},
        ]
        cockpit.add_executive_measures(tables)
        country_measures = {measure["name"]: measure for measure in tables[0]["measures"]}
        quality_measures = {measure["name"]: measure for measure in tables[1]["measures"]}
        self.assertTrue({item[2] for item in cockpit.LICENSE_COUNTRY_SERIES}.issubset(country_measures))
        self.assertIn("Executive Windows 10 devices", country_measures)
        self.assertIn("_build < 22000", country_measures["Executive Windows 10 devices"]["expression"])
        compatibility_expression = country_measures["Executive Windows 10 not compatible devices"]["expression"]
        self.assertIn("FactEndpointAnalyticsUpgradeEligibility", compatibility_expression)
        self.assertIn('= "notCapable"', compatibility_expression)
        self.assertIn("RETURN IF(_observed = 0, BLANK(), _notCapable)", compatibility_expression)
        self.assertIn(
            "unknown is never treated as incompatible",
            country_measures["Executive Windows 10 not compatible devices"]["description"],
        )
        self.assertIn("REMOVEFILTERS('DimCountry')", country_measures["Executive M365 E3 country share"]["expression"])
        endpoint_score = country_measures["Executive Endpoint Analytics score"]
        self.assertEqual(endpoint_score["formatString"], "0.0")
        self.assertIn("not a population percentage", endpoint_score["description"])
        self.assertEqual(cockpit.ENDPOINT_ANALYTICS_CARD_TITLE, "Endpoint Analytics Score")
        self.assertTrue({item[0] for item in cockpit.QUALITY_INDICATORS}.issubset(quality_measures))
        self.assertIn("Data quality score", quality_measures)
        self.assertIn("Data quality status", quality_measures)
        self.assertIn("Data quality health", quality_measures)
        self.assertIn("MINX", quality_measures["Data quality score"]["expression"])
        self.assertIn("REMOVEFILTERS('DimDevice')", quality_measures["Data quality score"]["expression"])
        self.assertIn("REMOVEFILTERS('DimUser')", quality_measures["Data quality score"]["expression"])
        self.assertIn("Overlapping findings", quality_measures["Data quality score"]["description"])
        self.assertEqual(tables[2]["measures"][0]["name"], "Account status share")

    def test_application_measures_distinguish_exact_relations_from_source_counts(self):
        self.assertTrue(
            {
                "SourceApplicationKey",
                "ReportedDeviceCount",
                "ExactRelatedDeviceCount",
                "RelationshipCoverageStatus",
            }.issubset(cockpit.DIM_DETECTED_APPLICATION_COLUMNS)
        )
        tables = [
            {"name": "DimDevice", "measures": []},
            {"name": "DimUser", "measures": []},
            {"name": "FactMailboxHosting", "measures": []},
            {"name": "DimDetectedApplication", "measures": []},
            {"name": "DeviceSource", "measures": []},
            {"name": "FactDeviceApplication", "measures": []},
        ]
        cockpit.add_operational_measures(tables)
        app_measures = {measure["name"]: measure for measure in tables[3]["measures"]}
        relation_measures = {measure["name"]: measure for measure in tables[5]["measures"]}
        self.assertIn("SourceApplicationKey", app_measures["Application products"]["expression"])
        self.assertIn("Reported application-device occurrences", app_measures)
        self.assertIn("ReportedDeviceCount", app_measures["Reported application-device occurrences"]["expression"])
        self.assertIn("SourceApplicationKey", relation_measures["Installed application products"]["expression"])
        self.assertIn("TREATAS", relation_measures["Installed application products"]["expression"])
        self.assertIn("Application-device installations", relation_measures)
        self.assertIn("COUNTROWS('FactDeviceApplication')", relation_measures["Application-device installations"]["expression"])

    def test_sharepoint_count_card_is_not_mistaken_for_a_share_metric(self):
        def card():
            return {"visual": {
                "objects": {"value": [{"properties": {}}]},
                "visualContainerObjects": {
                    "title": [{"properties": {"text": cockpit.lit("old")}}],
                },
            }}

        site_card = card()
        cockpit.set_card(site_card, "DimSharePointSite", "SharePoint sites", "SharePoint sites")
        self.assertEqual(
            site_card["visual"]["objects"]["value"][0]["properties"]["labelDisplayUnits"]["expr"]["Literal"]["Value"],
            "1D",
        )
        self.assertEqual(
            site_card["visual"]["objects"]["value"][0]["properties"]["labelPrecision"]["expr"]["Literal"]["Value"],
            "0D",
        )
        rate_card = card()
        cockpit.set_card(rate_card, "DimDevice", "Compliance rate", "Compliance rate")
        self.assertEqual(
            rate_card["visual"]["objects"]["value"][0]["properties"]["labelPrecision"]["expr"]["Literal"]["Value"],
            "1D",
        )

    def test_large_fact_row_count_is_streamed_and_validated(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "fact.csv"
            path.write_text("Id,State\n1,Keep\n2,Skip\n3,Keep\n", encoding="utf-8-sig")
            self.assertEqual(cockpit.count_csv_rows(path, lambda row: row["State"] == "Keep"), 2)

    def test_intune_device_bridge_keeps_one_exact_managed_device_key(self):
        identity = {
            "TenantKey": "tenant",
            "OrganizationKey": "organization",
            "EnvironmentKey": "prod",
            "TenantId": "tenant-id",
        }
        rows = [
            {
                **identity,
                "SourceSystem": "Entra",
                "SourceObjectId": "entra-id",
                "TenantDeviceKey": "tenant|device|one",
            },
            {
                **identity,
                "SourceSystem": "Intune",
                "SourceObjectId": "ABC-123",
                "TenantDeviceKey": "tenant|device|one",
            },
        ]
        bridge = cockpit.build_intune_device_bridge(rows, identity)
        self.assertEqual(len(bridge), 1)
        self.assertEqual(bridge[0]["TenantIntuneDeviceKey"], "tenant|intune-device|abc-123")
        self.assertEqual(bridge[0]["ManagedDeviceId"], "ABC-123")
        self.assertEqual(bridge[0]["TenantDeviceKey"], "tenant|device|one")

    def test_intune_device_bridge_rejects_duplicate_managed_device_ids(self):
        identity = {
            "TenantKey": "tenant",
            "OrganizationKey": "organization",
            "EnvironmentKey": "prod",
            "TenantId": "tenant-id",
        }
        row = {
            **identity,
            "SourceSystem": "Intune",
            "SourceObjectId": "ABC-123",
            "TenantDeviceKey": "tenant|device|one",
        }
        with self.assertRaisesRegex(ValueError, "Duplicate Intune device bridge key"):
            cockpit.build_intune_device_bridge([row, dict(row)], identity)

    def test_upgrade_eligibility_source_requires_exact_ids_and_supported_states(self):
        row = {column: "value" for column in cockpit.UPGRADE_ELIGIBILITY_COLUMNS}
        row.update({
            "TenantKey": "tenant",
            "OrganizationKey": "organization",
            "EnvironmentKey": "prod",
            "TenantId": "tenant-id",
            "DeviceId": "intune-device-id",
            "UpgradeEligibility": "capable",
            "SourceCollectedDateTime": "2026-09-13T00:00:00Z",
        })
        source = Path("FactEndpointAnalyticsUpgradeEligibility.csv")
        with patch.object(
            cockpit, "read_csv", return_value=(cockpit.UPGRADE_ELIGIBILITY_COLUMNS, [row])
        ):
            self.assertEqual(len(cockpit.validate_upgrade_eligibility_source(source)), 1)

        invalid = dict(row, UpgradeEligibility="inventedState")
        with patch.object(
            cockpit, "read_csv", return_value=(cockpit.UPGRADE_ELIGIBILITY_COLUMNS, [invalid])
        ):
            with self.assertRaisesRegex(ValueError, "unsupported state"):
                cockpit.validate_upgrade_eligibility_source(source)

    def test_upgrade_eligibility_source_rejects_conflicting_exact_device_ids(self):
        base = {column: "value" for column in cockpit.UPGRADE_ELIGIBILITY_COLUMNS}
        base.update({
            "TenantKey": "tenant",
            "OrganizationKey": "organization",
            "EnvironmentKey": "prod",
            "TenantId": "tenant-id",
            "DeviceId": "intune-device-id",
            "SourceCollectedDateTime": "2026-09-13T00:00:00Z",
        })
        rows = [
            dict(base, UpgradeEligibility="capable"),
            dict(base, UpgradeEligibility="notCapable"),
        ]
        source = Path("FactEndpointAnalyticsUpgradeEligibility.csv")
        with patch.object(
            cockpit, "read_csv", return_value=(cockpit.UPGRADE_ELIGIBILITY_COLUMNS, rows)
        ):
            with self.assertRaisesRegex(ValueError, "Conflicting upgrade eligibility"):
                cockpit.validate_upgrade_eligibility_source(source)

    def test_operational_category_palettes_distinguish_semantic_states(self):
        compliance = cockpit.BAR_CATEGORY_COLORS[("DimDevice", "ComplianceStateLabel")]
        updates = cockpit.BAR_CATEGORY_COLORS[("FactWindowsUpdateAlert", "AggregateState")]
        self.assertNotEqual(compliance["Compliant"], compliance["NonCompliant"])
        self.assertNotEqual(updates["Success"], updates["Error"])
        self.assertEqual(compliance["NonCompliant"], updates["Error"])

    def test_mailbox_type_group_preserves_user_shared_and_other_families(self):
        self.assertEqual(cockpit.mailbox_type_group({"RecipientTypeDetails": "UserMailbox"}), "User mailbox")
        self.assertEqual(cockpit.mailbox_type_group({"RecipientTypeDetails": "RemoteUserMailbox"}), "User mailbox")
        self.assertEqual(cockpit.mailbox_type_group({"RecipientTypeDetails": "SharedMailbox"}), "Shared mailbox")
        self.assertEqual(cockpit.mailbox_type_group({"RecipientTypeDetails": "RemoteSharedMailbox"}), "Shared mailbox")
        self.assertEqual(cockpit.mailbox_type_group({"RecipientTypeDetails": "RoomMailbox"}), "Other mailbox types")

    def test_ratio_bar_uses_percentage_measure_and_count_tooltip(self):
        visual = {
            "name": "testratiobar",
            "visual": {
                "objects": {},
                "visualContainerObjects": {"title": [{"properties": {}}]},
            }
        }
        cockpit.set_ratio_bar(
            visual,
            "DimUser", "AccountStatusLabel", "DimUser",
            "Account status share", "Users", "Enabled vs disabled users",
        )
        query = visual["visual"]["query"]["queryState"]
        self.assertEqual(visual["visual"]["visualType"], "clusteredBarChart")
        self.assertEqual(query["Y"]["projections"][0]["queryRef"], "DimUser.Account status share")
        self.assertEqual(query["Tooltips"]["projections"][0]["queryRef"], "DimUser.Users")
        self.assertEqual(
            visual["visual"]["objects"]["valueAxis"][0]["properties"]["end"]["expr"]["Literal"]["Value"],
            "1D",
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
        self.assertEqual(
            {row["RecipientTypeDetails"]: row["MailboxTypeGroup"] for row in rows},
            {
                "UserMailbox": "User mailbox",
                "RemoteUserMailbox": "User mailbox",
                "SharedMailbox": "Shared mailbox",
            },
        )

    def test_mailbox_hosting_baseline_retains_only_unmatched_onpremises_rows(self):
        identity = {
            "TenantKey": "tenant-key", "OrganizationKey": "org-key",
            "EnvironmentKey": "env-key", "TenantId": "tenant-id",
        }
        def mailbox_key(address):
            return cockpit.hashlib.sha256((identity["TenantKey"] + "|" + address).encode("utf-8")).hexdigest().upper()
        baseline = [
            dict(identity, MailboxHostingKey=mailbox_key("alice@example.test"), CountryLabel="France",
                 HostingLocation="Exchange On-premises", RecipientTypeDetails="UserMailbox",
                 MailboxTypeGroup="User mailbox", EvidenceSource="Dated local evidence"),
            dict(identity, MailboxHostingKey=mailbox_key("carol@example.test"), CountryLabel="France",
                 HostingLocation="Exchange On-premises", RecipientTypeDetails="SharedMailbox",
                 MailboxTypeGroup="Shared mailbox", EvidenceSource="Dated local evidence"),
        ]
        rows = cockpit.reconcile_mailboxes(
            [{"PrimarySmtpAddress": "alice@example.test", "RecipientTypeDetails": "UserMailbox"}],
            [], [], identity, {}, {}, baseline,
        )
        self.assertEqual(len(rows), 2)
        self.assertEqual(sum(row["HostingLocation"] == "Exchange Online" for row in rows), 1)
        self.assertEqual(sum(row["HostingLocation"] == "Exchange On-premises" for row in rows), 1)
        self.assertEqual(next(row for row in rows if row["HostingLocation"] == "Exchange Online")["MailboxHostingKey"], mailbox_key("alice@example.test"))

    def test_workplace_health_uses_windows_update_attention_evidence(self):
        with (
            patch.object(cockpit, "new_page", return_value=({}, [])),
            patch.object(cockpit, "add_slicer", return_value={}),
            patch.object(cockpit, "compact_page_header"),
            patch.object(cockpit, "add_card") as add_card,
            patch.object(cockpit, "add_ratio_bar") as add_ratio_bar,
            patch.object(cockpit, "add_table"),
        ):
            cockpit.build_risk(Path("ignored"))
        add_card.assert_any_call(
            Path("ignored"), [], "risk", "FactWindowsUpdateAlert",
            "Windows update alerts needing attention", "Update alerts needing attention", 336, 136, h=88,
        )
        add_ratio_bar.assert_any_call(
            Path("ignored"), [], "risk", "FactWindowsUpdateAlert", "AggregateState",
            "FactWindowsUpdateAlert", "Update aggregate state rate", "Windows update alert records",
            "Windows Update evidence by state", 648, 240, w=608, h=184,
        )

    def test_card_value_font_override_is_written_to_the_visual(self):
        template = {
            "visual": {
                "visualType": "cardVisual",
                "objects": {"value": [{"properties": {"fontSize": cockpit.lit(24)}}]},
            }
        }
        with (
            patch.object(cockpit, "clone_visual_by_type", return_value=template),
            patch.object(cockpit, "set_card"),
            patch.object(cockpit, "put", side_effect=lambda _items, visual, *_args: visual),
        ):
            visual = cockpit.add_card(
                Path("ignored"), [], "licenses", "DimLicenseSku", "License SKUs",
                "Observed SKUs", 24, 136, h=88, value_font_size=21,
            )

        font_size = visual["visual"]["objects"]["value"][0]["properties"]["fontSize"]
        self.assertEqual(font_size["expr"]["Literal"]["Value"], "21D")

    def test_licensing_kpis_request_the_compact_value_font(self):
        with (
            patch.object(cockpit, "new_page", return_value=({}, [])),
            patch.object(cockpit, "compact_page_header"),
            patch.object(cockpit, "add_slicer"),
            patch.object(cockpit, "add_card") as add_card,
            patch.object(cockpit, "add_ratio_bar"),
            patch.object(cockpit, "add_table", return_value={}),
            patch.object(cockpit, "set_categorical_values_filter"),
        ):
            cockpit.build_licensing(Path("ignored"))

        self.assertEqual(add_card.call_count, 4)
        self.assertTrue(
            all(call.kwargs.get("value_font_size") == 21 for call in add_card.call_args_list)
        )

    def test_lifecycle_uses_autopilot_and_endpoint_analytics(self):
        with (
            patch.object(cockpit, "new_page", return_value=({}, [])),
            patch.object(cockpit, "add_slicer", return_value={}),
            patch.object(cockpit, "compact_page_header"),
            patch.object(cockpit, "add_card") as add_card,
            patch.object(cockpit, "add_ratio_bar") as add_ratio_bar,
            patch.object(cockpit, "add_compact_card"),
            patch.object(cockpit, "add_table"),
        ):
            cockpit.build_lifecycle(Path("ignored"))
        add_card.assert_any_call(
            Path("ignored"), [], "lifecycle", "FactADIntuneCoverage",
            "AD to Intune coverage rate", "AD → Intune coverage", 336, 136, h=88,
        )
        add_card.assert_any_call(
            Path("ignored"), [], "lifecycle", "FactAutopilotDevice",
            "Selected Autopilot devices", "Autopilot · exact matches", 648, 136, h=88,
        )
        add_card.assert_any_call(
            Path("ignored"), [], "lifecycle", "FactEndpointAnalyticsDevice",
            "Selected Endpoint Analytics score", "Endpoint Analytics score", 960, 136, h=88,
        )
        add_ratio_bar.assert_any_call(
            Path("ignored"), [], "lifecycle", "FactAutopilotDevice", "EnrollmentState",
            "FactAutopilotDevice", "Autopilot enrollment rate", "Selected Autopilot devices",
            "Autopilot — exact Intune matches", 960, 240, w=296, h=204,
        )

    def test_exact_coverage_and_relationship_measures_are_source_bounded(self):
        tables = [
            {"name": "DimDevice", "measures": []},
            {"name": "DimUser", "measures": []},
            {"name": "FactADIntuneCoverage", "measures": []},
            {"name": "FactRelationshipOverview", "measures": []},
            {"name": "FactMailboxHosting", "measures": []},
            {
                "name": "FactUserDeviceRelationship",
                "measures": [{"name": "Relationship type rate", "expression": "legacy"}],
            },
        ]
        cockpit.add_operational_measures(tables)
        coverage = {item["name"]: item for item in tables[2]["measures"]}
        relationships = {item["name"]: item for item in tables[3]["measures"]}
        user_device = {item["name"]: item for item in tables[5]["measures"]}
        self.assertIn("AD to Intune coverage rate", coverage)
        self.assertIn("ObjectSid", coverage["AD Windows workstations managed in Intune"]["description"])
        self.assertIn("Relationship edges", relationships)
        self.assertNotIn("Relationship type rate", user_device)
        self.assertIn("User-device relationship type rate", user_device)
        self.assertEqual(
            set(cockpit.BAR_CATEGORY_COLORS[("FactRelationshipOverview", "RelationshipType")]),
            {"PrimaryUser", "HasMailbox", "AssignedLicense", "MemberOfGroup", "DeviceHasApplication", "DeviceInAutopilot"},
        )

    def test_fleet_page_uses_a_schema_safe_precomputed_top_five(self):
        with (
            patch.object(cockpit, "new_page", return_value=({}, [])),
            patch.object(cockpit, "add_slicer", return_value={}),
            patch.object(cockpit, "set_categorical_selection"),
            patch.object(cockpit, "compact_page_header"),
            patch.object(cockpit, "add_card"),
            patch.object(cockpit, "add_ratio_bar") as add_ratio_bar,
            patch.object(cockpit, "add_table"),
        ):
            cockpit.build_fleet_hardware(Path("ignored"))
        add_ratio_bar.assert_any_call(
            Path("ignored"), [], "devices", "TopApplication", "ApplicationProduct",
            "TopApplication", "Top application occurrence rate", "Top application occurrences",
            "Top 5 apps — global collected occurrences", 856, 240, w=400, h=184,
        )

    def test_people_page_uses_collected_activity_state(self):
        with (
            patch.object(cockpit, "new_page", return_value=({}, [])),
            patch.object(cockpit, "add_slicer", return_value={}),
            patch.object(cockpit, "compact_page_header"),
            patch.object(cockpit, "add_card"),
            patch.object(cockpit, "add_ratio_bar") as add_ratio_bar,
            patch.object(cockpit, "add_table"),
        ):
            cockpit.build_people_messaging(Path("ignored"))
        add_ratio_bar.assert_any_call(
            Path("ignored"), [], "users", "DimUser", "ActivityState",
            "DimUser", "Activity state rate", "Users", "Observed sign-in activity", 336, 240, w=296, h=184,
        )


if __name__ == "__main__":
    unittest.main()
