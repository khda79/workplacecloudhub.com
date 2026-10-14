"""Portable, evidence-aware history charts for migration comparison reports."""

import csv
import html
import math
import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from comparison_artifacts import find_artifact


START_MARKER = "<!-- comparison-history:start -->"
END_MARKER = "<!-- comparison-history:end -->"
RUN_CONTEXT_MARKER = '<section class="section" aria-labelledby="run-context">'
KINDS = {
    "files": (";", "MatchedKeys", "File"),
    "permissions": (",", "MatchedPermissions", "Permission"),
}


def _escape(value):
    return html.escape(str(value or ""), quote=True)


def _integer(value):
    try:
        number = int(str(value).strip())
        return number if number >= 0 else None
    except (TypeError, ValueError):
        return None


def _run_time(run_directory, kind, summary_path):
    match = re.search(rf"-{re.escape(kind)}-(\d{{8}})-(\d{{6}})(?:-|$)", run_directory.name, re.IGNORECASE)
    if match:
        try:
            return datetime.strptime("".join(match.groups()), "%Y%m%d%H%M%S")
        except ValueError:
            pass
    return datetime.fromtimestamp(summary_path.stat().st_mtime)


def _history_context(report_path, kind):
    report_path = Path(report_path).resolve()
    for run_directory in report_path.parents:
        if (run_directory.parent.name.lower() == kind
                and run_directory.parent.parent.name.lower() == "comparisons"):
            return run_directory.parent, run_directory, report_path.parent.relative_to(run_directory)
    return None


def _report_link(report_path, candidate_directory, summary):
    from_summary = str(summary.get("HtmlSummary") or "").replace("\\", "/").rsplit("/", 1)[-1]
    candidate = candidate_directory / from_summary if from_summary else None
    if not candidate or not candidate.is_file():
        reports = sorted(candidate_directory.glob("*summary*.html"))
        candidate = reports[-1] if reports else None
    if not candidate or not candidate.is_file():
        return ""
    relative = os.path.relpath(candidate, Path(report_path).parent).replace("\\", "/")
    return quote(relative, safe="/.-_")


def load_history(report_path, kind):
    """Read completed sibling comparisons up to and including this report's run."""
    if kind not in KINDS:
        raise ValueError(f"Unknown comparison kind: {kind}")
    context = _history_context(report_path, kind)
    if not context:
        return []
    history_root, current_run, scope = context
    current_summary = find_artifact(current_run / scope, "Summary.csv")
    if not current_summary.is_file():
        return []
    current_time = _run_time(current_run, kind, current_summary)
    delimiter, matched_field, _ = KINDS[kind]
    points = []

    for run_directory in history_root.iterdir():
        if not run_directory.is_dir():
            continue
        report_directory = run_directory / scope
        try:
            summary_path = find_artifact(report_directory, "Summary.csv")
        except ValueError:
            continue
        if not summary_path.is_file():
            continue
        run_time = _run_time(run_directory, kind, summary_path)
        if (run_time, run_directory.name) > (current_time, current_run.name):
            continue
        try:
            with summary_path.open(encoding="utf-8-sig", newline="") as handle:
                summary = next(csv.DictReader(handle, delimiter=delimiter), {})
        except (OSError, csv.Error, UnicodeError):
            summary = {}

        source = _integer(summary.get("SourceUniqueKeys"))
        target = _integer(summary.get("TargetUniqueKeys"))
        matched = _integer(summary.get(matched_field))
        validation = str(summary.get("ValidationStatus") or "Unknown")
        evidence = str(summary.get("ScanEvidenceStatus") or "Unknown")
        percent = None
        if source and validation != "InconclusiveEmptyInventory":
            if kind == "permissions" and target and matched is not None:
                percent = 100 * matched / source
            elif kind == "files":
                value = str(summary.get("SuccessPercent") or "").strip()
                if re.fullmatch(r"\d+(?:\.\d+)?%", value):
                    percent = float(value[:-1])
        if percent is not None and (not math.isfinite(percent) or not 0 <= percent <= 100):
            percent = None
        points.append({
            "time": run_time,
            "run": run_directory.name,
            "percent": percent,
            "matched": matched,
            "source": source,
            "evidence": evidence,
            "validation": validation,
            "href": _report_link(report_path, report_directory, summary),
            "current": run_directory == current_run,
        })

    points.sort(key=lambda point: (point["time"], point["run"]))
    return points


def _evidence_style(evidence):
    if evidence == "Verified":
        return "#1266a3", "Verified"
    if evidence == "LegacyFilename":
        return "#a66b00", "Legacy"
    if evidence in ("Stale", "Unverified"):
        return "#b13c43", "Stale or unverified"
    return "#687b8a", "Unknown evidence"


def render_history_section(report_path, kind):
    points = load_history(report_path, kind)
    if not points:
        return ""
    _, _, label = KINDS[kind]
    width = max(760, 110 + 82 * max(0, len(points) - 1))
    height = 305
    left, right, top, bottom = 58, 28, 22, 225
    plot_width = width - left - right
    x_positions = [left + plot_width * (index / (len(points) - 1) if len(points) > 1 else 0.5)
                   for index in range(len(points))]

    svg = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" role="img" aria-labelledby="history-title history-description" '
        'style="display:block;font-family:Segoe UI,Arial,sans-serif;font-size:12px">',
        f'<title id="history-title">{label} comparison success rate history</title>',
        '<desc id="history-description">One position per comparison. Evidence warnings are marked in amber or red. '
        'Unavailable percentages are labelled N/A. The table below contains every value.</desc>',
    ]
    for tick in (0, 25, 50, 75, 100):
        y = bottom - (bottom - top) * tick / 100
        svg.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width-right}" y2="{y:.1f}" '
                   'stroke="#dbe5ed" stroke-width="1"/>')
        svg.append(f'<text x="{left-9}" y="{y+4:.1f}" text-anchor="end" fill="#617386">{tick}%</text>')

    for index in range(1, len(points)):
        previous, current = points[index - 1], points[index]
        if previous["percent"] is None or current["percent"] is None:
            continue
        y1 = bottom - (bottom - top) * previous["percent"] / 100
        y2 = bottom - (bottom - top) * current["percent"] / 100
        qualified = previous["evidence"] == current["evidence"] == "Verified"
        dash = "" if qualified else ' stroke-dasharray="5 4"'
        svg.append(f'<line x1="{x_positions[index-1]:.1f}" y1="{y1:.1f}" '
                   f'x2="{x_positions[index]:.1f}" y2="{y2:.1f}" '
                   f'stroke="{ "#1266a3" if qualified else "#8b9ba9" }" stroke-width="2"{dash}/>')

    label_step = max(1, math.ceil(len(points) / 12))
    for index, point in enumerate(points):
        x = x_positions[index]
        timestamp = point["time"].strftime("%d/%m/%Y %H:%M")
        percent_text = f'{point["percent"]:.2f}%' if point["percent"] is not None else "N/A"
        matched_text = str(point["matched"]) if point["matched"] is not None else "N/A"
        source_text = str(point["source"]) if point["source"] is not None else "N/A"
        tooltip = (f'{timestamp}: {percent_text}; matched/source {matched_text}/{source_text}; '
                   f'scan evidence {point["evidence"]}; status {point["validation"]}')
        if point["percent"] is not None:
            y = bottom - (bottom - top) * point["percent"] / 100
            color, _ = _evidence_style(point["evidence"])
            marker = (f'<circle class="history-point" cx="{x:.1f}" cy="{y:.1f}" '
                      f'r="{6 if point["current"] else 5}" fill="{color}" stroke="white" '
                      f'stroke-width="2"><title>{_escape(tooltip)}</title></circle>')
        else:
            marker = (f'<text x="{x:.1f}" y="{bottom+50}" text-anchor="middle" '
                      f'fill="#687b8a" font-weight="600">N/A<title>{_escape(tooltip)}</title></text>')
        if point["href"]:
            marker = f'<a href="{_escape(point["href"])}">{marker}</a>'
        svg.append(marker)
        if index in (0, len(points)-1) or index % label_step == 0:
            svg.append(f'<text x="{x:.1f}" y="{bottom+25}" text-anchor="middle" fill="#51677b">'
                       f'{point["time"]:%d/%m %H:%M}</text>')
    svg.append("</svg>")

    table_rows = []
    for point in points:
        percent_text = f'{point["percent"]:.2f}%' if point["percent"] is not None else "N/A"
        matched_text = str(point["matched"]) if point["matched"] is not None else "N/A"
        source_text = str(point["source"]) if point["source"] is not None else "N/A"
        link = f'<a href="{_escape(point["href"])}">Open report</a>' if point["href"] else "—"
        table_rows.append(
            "<tr>"
            f'<td>{point["time"]:%Y-%m-%d %H:%M}</td>'
            f'<td class="num">{percent_text}</td>'
            f'<td class="num">{matched_text} / {source_text}</td>'
            f'<td>{_escape(point["evidence"])}</td>'
            f'<td>{_escape(point["validation"])}</td>'
            f'<td>{link}</td>'
            "</tr>"
        )
    unavailable = sum(point["percent"] is None for point in points)
    unavailable_note = f"; {unavailable} unavailable" if unavailable else ""
    return (
        f'{START_MARKER}\n'
        '<section class="section" aria-labelledby="comparison-history">'
        '<div class="section-heading"><h2 id="comparison-history">Success rate over time</h2>'
        f'<span class="section-note">{len(points)} comparisons{unavailable_note}</span></div>'
        '<p style="color:#617386;margin:0 0 14px">Success rate uses each run\'s own source key count. '
        'The comparison scope can change between runs. Dashed segments include less qualified scan evidence. '
        'Hover over a point for details; select it to open that report.</p>'
        '<div style="max-width:100%;overflow-x:auto" role="region" aria-label="Comparison history chart" tabindex="0">'
        f'{"".join(svg)}</div>'
        '<div style="display:flex;flex-wrap:wrap;gap:8px 18px;color:#51677b;font-size:12px;margin-top:8px">'
        '<span style="color:#1266a3">● Verified</span><span style="color:#a66b00">● Legacy evidence</span>'
        '<span style="color:#b13c43">● Stale or unverified</span><span>N/A = unavailable, not 0%</span></div>'
        '<details style="margin-top:16px"><summary style="cursor:pointer;font-weight:650">All comparison values</summary>'
        '<div class="table-scroll" role="region" aria-label="Comparison history values" tabindex="0">'
        '<table><thead><tr><th scope="col">Run date</th><th scope="col">Success rate</th>'
        '<th scope="col">Matched / source keys</th><th scope="col">Scan evidence</th>'
        '<th scope="col">Status</th><th scope="col">Report</th></tr></thead>'
        f'<tbody>{"".join(table_rows)}</tbody></table></div></details></section>\n'
        f'{END_MARKER}'
    )


def refresh_history_in_existing_html(report_path, kind):
    """Update only the history block in an existing, generated HTML report."""
    report_path = Path(report_path)
    history = render_history_section(report_path, kind)
    if not history or not report_path.is_file():
        return False
    content = report_path.read_text(encoding="utf-8")
    if START_MARKER in content and END_MARKER in content:
        before, rest = content.split(START_MARKER, 1)
        _, after = rest.split(END_MARKER, 1)
        updated = before + history + after
    elif RUN_CONTEXT_MARKER in content:
        updated = content.replace(RUN_CONTEXT_MARKER, history + "\n  " + RUN_CONTEXT_MARKER, 1)
    else:
        return False
    if updated == content:
        return True
    temp_path = report_path.with_name(report_path.name + ".history.tmp")
    try:
        temp_path.write_text(updated, encoding="utf-8", newline="\n")
        temp_path.replace(report_path)
    finally:
        temp_path.unlink(missing_ok=True)
    return True
