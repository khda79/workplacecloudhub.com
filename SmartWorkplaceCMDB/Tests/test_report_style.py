"""Synthetic tests for the report-only Smart Workplace CMDB restyle."""

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "PowerBI"))
import restyle_report as style


class ReportStyle(unittest.TestCase):
    def test_restyle_removes_beta_from_visible_context_and_is_idempotent(self):
        with tempfile.TemporaryDirectory(prefix="cmdb-report-style-") as temp:
            report = Path(temp) / "Example.Report"
            page_dir = report / "definition/pages/overview"
            visual_dir = page_dir / "visuals"
            page_dir.mkdir(parents=True)
            page = {"name": "overview", "displayName": "01  Overview · BETA",
                    "displayOption": "FitToPage", "height": 900, "width": 1280}
            (page_dir / "page.json").write_text(json.dumps(page), encoding="utf-8")

            def write_visual(name, kind, y, text=None, query=None):
                payload = {"name": name, "position": {"x": 24, "y": y, "width": 1232, "height": 28,
                                                              "z": y, "tabOrder": y},
                           "visual": {"visualType": kind, "objects": {}, "visualContainerObjects": {},
                                      "query": query}}
                if text is not None:
                    payload["visual"]["objects"] = {"general": [{"properties": {"paragraphs": [{"textRuns": [
                        {"value": text, "textStyle": {"fontFamily": "Segoe UI", "fontSize": "14px", "color": "#000000"}}
                    ]}]}}]}
                target = visual_dir / name
                target.mkdir(parents=True)
                (target / "visual.json").write_text(json.dumps(payload), encoding="utf-8")

            write_visual("brand", "textbox", 12, "SMARTWORKPLACECMDB / BETA")
            write_visual("title", "textbox", 44, "01  Overview")
            write_visual("subtitle", "textbox", 100, "Frozen source context")
            write_visual("footer", "textbox", 864, "BETA · Frozen snapshot")
            query = {"queryState": {"Data": {"projections": [{"queryRef": "Facts.Count"}]}}}
            write_visual("card", "cardVisual", 224, query=query)

            style.restyle(report)
            result_page = json.loads((page_dir / "page.json").read_text(encoding="utf-8"))
            self.assertEqual(result_page["displayName"], "01  Overview")
            self.assertEqual(result_page["width"], 1280)
            self.assertEqual(result_page["objects"]["background"][0]["properties"]["color"], style.color(style.PAGE_BACKGROUND))

            title = json.loads((visual_dir / "brand/visual.json").read_text(encoding="utf-8"))
            run = title["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]
            self.assertEqual(run["value"], "Smart Workplace CMDB — Overview")
            version = json.loads((visual_dir / "title/visual.json").read_text(encoding="utf-8"))
            self.assertEqual(version["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"],
                             "VERSION 1.0.0")
            card = json.loads((visual_dir / "card/visual.json").read_text(encoding="utf-8"))
            self.assertEqual(card["visual"]["query"], query)
            self.assertEqual(card["position"]["y"], 224)
            self.assertEqual(card["visual"]["visualContainerObjects"]["border"][0]["properties"]["radius"], style.literal(8))

            style.restyle(report)
            version = json.loads((visual_dir / "title/visual.json").read_text(encoding="utf-8"))
            self.assertEqual(version["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"],
                             "VERSION 1.0.0")
            footer = json.loads((visual_dir / "footer/visual.json").read_text(encoding="utf-8"))
            self.assertEqual(footer["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"],
                             "V1 · Frozen snapshot")


if __name__ == "__main__":
    unittest.main()
