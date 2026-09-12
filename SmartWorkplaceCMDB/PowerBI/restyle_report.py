"""Apply the Smart Workplace light visual identity to an existing CMDB PBIR.

The transformation is report-only: it preserves data bindings, filters,
interactions, page dimensions, visual ids and semantic-model files.
"""

import argparse
import json
import re
from pathlib import Path


PAGE_BACKGROUND = "#F5F8FB"
WALLPAPER = "#E8EEF5"
SURFACE = "#FFFFFF"
BORDER = "#DDE7F0"
TITLE = "#1F2937"
MUTED = "#526577"
VERSION = "1.0.0"


def literal(value):
    if isinstance(value, bool):
        encoded = "true" if value else "false"
    elif isinstance(value, (int, float)):
        encoded = f"{value}D"
    else:
        encoded = "'" + str(value).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": encoded}}}


def color(value):
    return {"solid": {"color": literal(value)}}


def load(path):
    return json.loads(path.read_text(encoding="utf-8"))


def save(path, payload):
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def set_textbox(visual, value, x, y, width, height, size, ink, font="Segoe UI", weight=None):
    run = visual["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]
    run["value"] = value
    style = run.setdefault("textStyle", {})
    style.update(fontFamily=font, fontSize=f"{size}px", color=ink)
    if weight:
        style["fontWeight"] = weight
    else:
        style.pop("fontWeight", None)
    visual["position"].update(x=x, y=y, width=width, height=height)


def set_vco(visual):
    vco = visual["visual"].setdefault("visualContainerObjects", {})
    vco["background"] = [{"properties": {
        "show": literal(True), "color": color(SURFACE), "transparency": literal(0)}}]
    vco["border"] = [{"properties": {
        "show": literal(True), "color": color(BORDER), "radius": literal(8)}}]
    vco["visualHeader"] = [{"properties": {"show": literal(False)}}]
    vco["dropShadow"] = [{"properties": {"show": literal(False)}}]
    title = vco.get("title", [{"properties": {}}])
    props = title[0].setdefault("properties", {})
    if props.get("show", {}).get("expr", {}).get("Literal", {}).get("Value") != "false":
        props.update(fontSize=literal(12), fontColor=color(TITLE), bold=literal(False))
    vco["title"] = title


def restyle(report):
    report = Path(report)
    pages = report / "definition" / "pages"
    changed = []
    for page_dir in sorted(path for path in pages.iterdir() if path.is_dir()):
        page_path = page_dir / "page.json"
        if not page_path.exists():
            continue
        page = load(page_path)
        page["displayName"] = re.sub(r"\s*·\s*BETA\s*$", "", page["displayName"])
        objects = page.setdefault("objects", {})
        objects["background"] = [{"properties": {"color": color(PAGE_BACKGROUND), "transparency": literal(0)}}]
        objects["outspace"] = [{"properties": {"color": color(WALLPAPER), "transparency": literal(0)}}]
        save(page_path, page)
        changed.append(str(page_path))

        short_title = re.sub(r"^\d+\s+", "", page["displayName"])
        for visual_path in sorted((page_dir / "visuals").glob("*/visual.json")):
            visual = load(visual_path)
            kind = visual.get("visual", {}).get("visualType")
            position = visual.get("position", {})
            current_text = None
            if kind == "textbox":
                current_text = visual["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"]
            if kind == "textbox" and position.get("y") in (12, 14):
                set_textbox(visual, f"Smart Workplace CMDB — {short_title}", 24, 14, 900, 42,
                            26, TITLE, "Segoe UI Semibold", "bold")
            elif kind == "textbox" and (position.get("y") in (44, 60) or
                                         re.fullmatch(r"(?:BETA|VERSION)\s+[^·]+", current_text or "")):
                set_textbox(visual, f"VERSION {VERSION}", 24, 60, 900, 26, 11, MUTED)
            elif kind == "textbox" and position.get("y") in (96, 100):
                set_textbox(visual, current_text, 24, 96, 1232, 44, 13, MUTED)
            elif kind == "textbox" and position.get("y") == 864:
                footer = re.sub(r"^BETA\b", "V1", current_text or "")
                footer = footer.replace("0.3.0-beta.1", VERSION)
                set_textbox(visual, footer, 24, 864, 1232, 28, 10, MUTED)
            elif kind == "textbox" and current_text and "BETA" in current_text.upper():
                stable_text = re.sub(r"\bBETA\b", "V1", current_text, flags=re.IGNORECASE)
                stable_text = stable_text.replace("0.3.0-beta.1", VERSION)
                set_textbox(visual, stable_text, position.get("x", 24), position.get("y", 0),
                            position.get("width", 1232), position.get("height", 28), 11, MUTED)
            elif kind != "textbox":
                set_vco(visual)
            else:
                continue
            save(visual_path, visual)
            changed.append(str(visual_path))
    return changed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path, help="Path to the target .Report directory")
    args = parser.parse_args()
    changed = restyle(args.report)
    print(json.dumps({"status": "Restyled", "files": len(changed), "report": str(args.report)}))


if __name__ == "__main__":
    main()
