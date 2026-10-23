"""Offline contracts for all-state AD computer trends and their report views.

Synthetic source fixtures exercise the existing qualification producer. DAX
execution is a separate Modeling MCP check; these tests do not refresh data.
"""
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
import json
import os
from pathlib import Path
import re
import unittest

import new_daily_trends_evidence as daily


PROJECT = Path(__file__).resolve().parents[1]
REPORT = Path(os.environ.get("SMARTWORKPLACE_REPORT", PROJECT / "pbip/SmartWorkplaceIntelligence.Report"))
MODEL = Path(os.environ.get("SMARTWORKPLACE_MODEL", PROJECT / "pbip/SmartWorkplaceIntelligence.SemanticModel/definition"))
DEFINITIONS = os.environ.get("SMARTWORKPLACE_DAILY_MEASURES")
PAGE = "daily-directory-mailbox-trends"


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def expressions():
    if DEFINITIONS:
        return {item["name"]: item["expression"] for item in read(Path(DEFINITIONS))}
    result = {}
    for name in ("Daily Directory Trends", "Daily Axis"):
        text = (MODEL / "tables" / (name + ".tmdl")).read_text(encoding="utf-8-sig")
        for match in re.finditer(r"^\tmeasure '([^']+)' =([\s\S]*?)(?=\n\t(?:measure|column|partition|hierarchy|annotation) |\Z)", text, re.M):
            result[match[1]] = match[2]
    return result


class DailyComputerAccountsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.measures = expressions()
        cls.containers = [read(path) for path in (REPORT / "definition/pages" / PAGE / "visuals").glob("*/visual.json")]

    def test_five_views_preserve_slicers_and_show_one_group(self):
        groups = {item["name"] for item in self.containers if "visualGroup" in item}
        added = next(item["name"] for item in self.containers if item.get("visualGroup", {}).get("displayName") == "AD computer accounts")
        self.assertEqual(len(groups), 5)
        index = read(REPORT / "definition/bookmarks/bookmarks.json")
        group = next(item for item in index["items"] if item["name"] == "80d83479f6abbd1e246d")
        self.assertEqual(len(group["children"]), 5)
        labels = []
        visible = set()
        for name in group["children"]:
            mark = read(REPORT / "definition/bookmarks" / (name + ".bookmark.json"))
            labels.append(mark["displayName"])
            self.assertTrue(mark["options"]["suppressData"])
            self.assertTrue(mark["options"]["suppressActiveSection"])
            self.assertEqual(set(mark["options"]["targetVisualNames"]), groups)
            state = mark["explorationState"]["sections"][PAGE]
            self.assertEqual(state["visualContainers"], {})
            shown = [key for key, value in state["visualContainerGroups"].items() if not value["isHidden"]]
            self.assertEqual(len(shown), 1)
            visible.update(shown)
        self.assertIn("AD computer accounts", labels)
        self.assertIn("Windows AD accounts", labels)
        self.assertEqual(visible, groups)
        self.assertIn(added, visible)

    def test_all_os_view_exposes_total_enabled_and_disabled_at_one_reading(self):
        group = next(item["name"] for item in self.containers if item.get("visualGroup", {}).get("displayName") == "AD computer accounts")
        view = [item for item in self.containers if item.get("parentGroupName") == group]
        chart = next(item for item in view if item.get("visual", {}).get("visualType") == "lineChart")
        fields = [item["field"]["Measure"]["Property"] for item in chart["visual"]["query"]["queryState"]["Y"]["projections"]]
        self.assertEqual(fields, ["Axis Total AD Computer Accounts", "Axis Enabled AD Computer Accounts", "Axis Disabled AD Computer Accounts"])
        for name in fields:
            self.assertIn("[Axis All AD Computer Accounts Reading Day]", self.measures[name])
        cards = [item for item in view if item.get("visual", {}).get("visualType") == "cardVisual" and item["position"]["y"] == 296]
        self.assertEqual(len(cards), 3)
        self.assertEqual([item["position"]["height"] for item in cards], [128] * 3)
        for item in view:
            position = item["position"]
            self.assertGreaterEqual(position["x"], 0)
            self.assertGreaterEqual(position["y"], 0)
            self.assertLessEqual(position["x"] + position["width"], 1920)
            self.assertLessEqual(position["y"] + position["height"], 1080)

    def test_windows_missing_state_is_not_coerced_to_zero(self):
        for version in ("10", "11"):
            expression = self.measures[f"Windows {version} All AD Accounts Daily"]
            self.assertIn(f"[Windows {version} Enabled Daily]", expression)
            self.assertIn(f"[Windows {version} Disabled Daily]", expression)
            self.assertIn("ISBLANK(_enabled) || ISBLANK(_disabled)", expression)
            self.assertIn("BLANK()", expression)
            self.assertNotIn("COALESCE", expression)
            share = self.measures[f"Windows {version} All AD Accounts Share Daily"]
            self.assertIn("[Total AD Computers Daily]", share)
            self.assertIn("ISBLANK(_windows) || ISBLANK(_total)", share)
        chart = next(item for item in self.containers if item["name"] == "fdb15a422357d5010529")
        values = chart["visual"]["query"]["queryState"]["Y"]["projections"]
        self.assertEqual([item["field"]["Measure"]["Property"] for item in values], ["Axis Windows 10 All AD Accounts Chart", "Axis Windows 11 All AD Accounts Chart"])
        self.assertEqual(values[0]["displayName"], "Windows 10 AD computer accounts")

    def test_computer_card_captions_match_their_bound_measures(self):
        expected = {
            "Total AD Computer Accounts Card": "All AD computer accounts (all OS)",
            "Enabled AD Computer Accounts Card": "Enabled AD computer accounts (all OS)",
            "Disabled AD Computer Accounts Card": "Disabled AD computer accounts (all OS)",
            "Windows 10 All AD Accounts Card": "Windows 10 AD computer accounts",
            "Windows 11 All AD Accounts Share and Count Change Card": "Windows 11 AD computer accounts share",
            "Windows 11 Overtakes Windows 10 All AD Accounts Day": "First observed day: Windows 11 > Windows 10",
        }
        found = set()
        for item in self.containers:
            visual = item.get("visual", {})
            if visual.get("visualType") != "cardVisual":
                continue
            measure = visual["query"]["queryState"]["Data"]["projections"][0]["field"]["Measure"]["Property"]
            if measure not in expected:
                continue
            objects = visual["visualContainerObjects"]
            title = objects["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"]
            alt = objects["general"][0]["properties"]["altText"]["expr"]["Literal"]["Value"]
            self.assertEqual(title, "'" + expected[measure] + "'")
            self.assertIn("AD computer accounts", alt)
            self.assertNotIn("AD user", alt)
            if measure.startswith("Windows") or measure.startswith("Total"):
                self.assertIn("enabled and disabled", alt)
            elif measure.startswith("Enabled"):
                self.assertIn("Enabled AD computer accounts", alt)
            else:
                self.assertIn("Disabled AD computer accounts", alt)
            found.add(measure)
        self.assertEqual(found, set(expected))

    def test_computer_chart_captions_and_axes_do_not_inherit_user_or_enabled_only_scope(self):
        expected = {
            "eb190fea0be0ed0e8a7c": ("AD computer accounts - all operating systems and account states", "AD computer accounts (all OS)"),
            "fdb15a422357d5010529": ("Windows 10 / Windows 11 - AD computer accounts", "AD computer accounts (count or share)"),
        }
        for visual_id, (title, axis) in expected.items():
            visual = next(item["visual"] for item in self.containers if item["name"] == visual_id)
            objects = visual["visualContainerObjects"]
            self.assertEqual(objects["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"], "'" + title + "'")
            self.assertEqual(visual["objects"]["valueAxis"][0]["properties"]["titleText"]["expr"]["Literal"]["Value"], "'" + axis + "'")
            alt = objects["general"][0]["properties"]["altText"]["expr"]["Literal"]["Value"]
            self.assertIn("enabled and disabled", alt)
            self.assertNotIn("AD user", alt)

    def test_daily_prose_quality_cards_keep_readable_wrapped_diagnostics(self):
        quality_cards = []
        for item in self.containers:
            visual = item.get("visual", {})
            if visual.get("visualType") != "cardVisual":
                continue
            measures = [entry["field"]["Measure"]["Property"] for entry in visual["query"]["queryState"]["Data"]["projections"]]
            if not any(name.endswith("Quality Text") for name in measures):
                continue
            props = visual["objects"]["value"][0]["properties"]
            font = props["fontSize"]["expr"]["Literal"]["Value"]
            self.assertIn(font, ("10D", "11D"))
            self.assertEqual(props["textWrap"]["expr"]["Literal"]["Value"], "true")
            quality_cards.append(item["name"])
        self.assertEqual(len(quality_cards), 6)

    def test_windows_reading_is_common_and_separate_from_all_os_reading(self):
        windows = self.measures["Axis Windows All AD Accounts Reading Day"]
        all_os = self.measures["Axis All AD Computer Accounts Reading Day"]
        for name in ("Total AD Computers Daily", "Windows 10 All AD Accounts Daily", "Windows 11 All AD Accounts Daily"):
            self.assertIn("NOT ISBLANK([" + name + "])", windows)
        for name in ("Total AD Computers Daily", "Enabled AD Computer Accounts Daily", "Disabled AD Computer Accounts Daily"):
            self.assertIn("NOT ISBLANK([" + name + "])", all_os)
        for expression in (windows, all_os):
            self.assertIn("[Axis Point Visible]", expression)
            self.assertIn("[Axis Is Weekly]", expression)
            self.assertIn("ALLSELECTED('Date'[Date])", expression)
        self.assertNotIn("Windows", all_os)

    def test_synthetic_source_retains_other_os_and_disabled_windows(self):
        spec = daily.SOURCES["AD Computers"]
        identity = ("test-tenant", "test-org", "test-env", "test-id")
        by_day = defaultdict(dict)
        for domain, total, enabled, w10e, w10d, w11e, w11d in [("alpha.example", 50, 30, 5, 3, 7, 2), ("beta.example", 40, 25, 4, 2, 1, 1)]:
            when = datetime(2026, 1, 5, 8)
            row = {name: "0" for name in spec["metrics"]}
            row.update({"DomainName": domain, "TotalComputers": str(total), "EnabledAccounts": str(enabled), "DisabledAccounts": str(total - enabled), "Windows 10 Enabled": str(w10e), "Windows 10 Disabled": str(w10d), "Windows 11 Enabled": str(w11e), "Windows 11 Disabled": str(w11d)})
            by_day[when.date()][domain] = (when, row)
        rows = daily.generate_source("AD Computers", spec, identity, by_day, Decimal(20), 7, {"manualDays": [], "persistentDomains": []})
        totals = {row["Metric Name"]: row for row in rows if row["Domain Key"] == "__all__"}
        self.assertEqual(totals["TotalComputers"]["Metric Value"], "90")
        self.assertEqual(totals["DisabledAccounts"]["Metric Value"], "35")
        self.assertEqual(totals["Windows 10 Disabled"]["Metric Value"], "5")
        self.assertEqual(totals["Windows 11 Disabled"]["Metric Value"], "3")
        # A missing state in one domain stays unmeasured at tenant level.
        by_day[when.date()]["beta.example"][1]["Windows 10 Disabled"] = ""
        rows = daily.generate_source("AD Computers", spec, identity, by_day, Decimal(20), 7, {"manualDays": [], "persistentDomains": []})
        missing = next(row for row in rows if row["Domain Key"] == "__all__" and row["Metric Name"] == "Windows 10 Disabled")
        self.assertEqual(missing["Metric Status"], "Not measured")
        self.assertEqual(missing["Metric Value"], "")


if __name__ == "__main__":
    unittest.main()
