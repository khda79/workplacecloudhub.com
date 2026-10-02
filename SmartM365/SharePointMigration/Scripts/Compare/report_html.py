"""Shared, self-contained HTML presentation for migration comparison reports."""

import base64
import html
import json
from pathlib import Path


REPORT_CSS = """
:root {
  color-scheme: light;
  --navy: #071527;
  --navy-2: #0d2747;
  --blue: #3aa0ff;
  --blue-soft: #eaf5ff;
  --violet: #8367ff;
  --ink: #111827;
  --muted: #637083;
  --line: #dfe8f4;
  --surface: #ffffff;
  --background: #f7faff;
  --danger: #a82435;
  --warning: #825600;
  --success: #14663c;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
  margin: 0;
  background: var(--background);
  color: var(--ink);
  font: 14px/1.5 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
.skip-link { position: absolute; left: -10000px; top: 10px; z-index: 10; }
.skip-link:focus { left: 10px; padding: 8px 12px; background: white; color: var(--navy); }
.brand-band { height: 5px; background: linear-gradient(90deg, var(--blue), var(--violet)); }
main { width: min(1440px, calc(100% - 40px)); margin: 28px auto 48px; }
.hero {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 24px;
  padding: 26px 30px;
  border-radius: 18px;
  background: linear-gradient(125deg, var(--navy), var(--navy-2));
  color: white;
  box-shadow: 0 14px 34px rgba(7, 21, 39, .12);
}
.hero-main { min-width: 0; }
.brand { display: flex; align-items: center; gap: 18px; margin-bottom: 18px; }
.brand img { display: block; width: 145px; max-height: 100px; object-fit: contain; object-position: left center; }
.client-brand { margin: 0 0 16px; display: flex; align-items: center; gap: 10px; }
.client-brand img { max-width: 180px; max-height: 72px; object-fit: contain; background: white; padding: 5px; border-radius: 6px; }
.download-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 14px; margin: 16px 0; padding: 13px 18px; background: white; border: 1px solid var(--line); border-radius: 12px; }
.download-bar strong { margin-right: auto; }
.download-primary { display: inline-block; padding: 9px 15px; border-radius: 7px; background: #086db8; color: white; font-weight: 700; text-decoration: none; }
.download-primary:hover { background: #005a9e; color: white; }
.download-secondary { font-weight: 600; }
.brand-name { font-weight: 750; font-size: 16px; letter-spacing: .01em; }
.brand-product { color: #c9e3ff; font-size: 12px; }
.eyebrow { color: #9ed3ff; font-size: 11px; font-weight: 750; letter-spacing: .14em; text-transform: uppercase; }
h1 { max-width: 900px; margin: 7px 0 4px; font-size: clamp(24px, 3vw, 34px); line-height: 1.17; }
.subtitle { color: #c9d9eb; font-size: 13px; }
.badge { display: inline-flex; align-items: center; border-radius: 999px; padding: 7px 12px; font-weight: 700; white-space: nowrap; }
.badge.ok { color: #d9ffea; background: rgba(30, 150, 92, .25); border: 1px solid #55bd87; }
.badge.warn { color: #ffe3e7; background: rgba(168, 36, 53, .28); border: 1px solid #f09ca8; }
.badge.note { color: #fff0c9; background: rgba(167, 112, 16, .28); border: 1px solid #ecc779; }
.intro { margin: 18px 2px; color: var(--muted); font-size: 13px; }
.metrics { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; margin: 18px 0; }
.metric { min-width: 0; padding: 17px 18px; border: 1px solid var(--line); border-top: 4px solid var(--blue); border-radius: 12px; background: var(--surface); box-shadow: 0 4px 14px rgba(7, 21, 39, .035); }
.metric.bad { border-top-color: var(--danger); }
.metric.note { border-top-color: var(--warning); }
.metric.muted { border-top-color: var(--muted); }
.metric-label { color: var(--muted); font-size: 12px; font-weight: 650; }
.metric-value { margin-top: 8px; font-size: 29px; line-height: 1.1; font-weight: 750; font-variant-numeric: tabular-nums; }
.metric.bad .metric-value { color: var(--danger); }
.metric.note .metric-value { color: var(--warning); }
.metric.ok .metric-value { color: var(--navy-2); }
.metric.muted .metric-value { color: var(--muted); }
.section { margin: 18px 0; padding: 21px 23px; border: 1px solid var(--line); border-radius: 12px; background: var(--surface); }
.section-heading { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: baseline; gap: 8px 16px; margin-bottom: 13px; }
h2 { margin: 0; color: var(--navy); font-size: 18px; }
.section-note { color: var(--muted); font-size: 12px; }
.context { display: grid; grid-template-columns: minmax(180px, 250px) minmax(0, 1fr); gap: 9px 16px; margin: 0; }
.context dt { color: var(--muted); }
.context dd { margin: 0; overflow-wrap: anywhere; }
a { color: #075d9f; text-decoration: underline; text-underline-offset: 2px; }
a:hover { color: var(--navy); }
a:focus-visible { outline: 3px solid var(--violet); outline-offset: 3px; }
.table-scroll { max-width: 100%; overflow-x: auto; }
table { width: 100%; min-width: 700px; border-collapse: collapse; }
th, td { padding: 10px 12px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }
th { color: var(--navy-2); background: var(--blue-soft); font-size: 12px; font-weight: 750; white-space: nowrap; }
tbody tr:nth-child(even) { background: #fafcff; }
tbody tr:hover { background: #f1f8ff; }
td { overflow-wrap: anywhere; }
.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.empty { color: var(--muted); text-align: center; padding: 20px; }
.callout { margin: 18px 0; padding: 14px 18px; border-radius: 10px; }
.callout.warn { color: #5f4100; background: #fff5dc; border: 1px solid #e7c46e; }
.footer { margin: 22px 2px; color: var(--muted); font-size: 12px; }
@media (max-width: 720px) {
  main { width: min(100% - 24px, 1440px); margin-top: 12px; }
  .hero { display: block; padding: 20px; }
  .hero .badge { margin-top: 16px; }
  .section { padding: 17px 15px; }
  .context { grid-template-columns: 1fr; gap: 2px; }
  .context dd { margin-bottom: 10px; }
}
@media print {
  body { background: white; }
  main { width: 100%; margin: 0; }
  .hero { box-shadow: none; border: 1px solid var(--line); print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .metric, .section { box-shadow: none; break-inside: avoid; }
  .table-scroll { overflow: visible; }
  table { min-width: 0; font-size: 10px; }
  th, td { padding: 5px 6px; }
  a { color: var(--navy); }
}
"""


def escape(value):
    return html.escape(str(value or ""), quote=True)


def metric_card(label, value, tone):
    return (
        f'<div class="metric {escape(tone)}">'
        f'<div class="metric-label">{escape(label)}</div>'
        f'<div class="metric-value">{escape(value)}</div></div>'
    )


def _branding_config():
    config = Path(__file__).resolve().parents[2] / "Config" / "report-branding.json.txt"
    if not config.is_file():
        return {}
    try:
        data = json.loads(config.read_text(encoding="utf-8-sig"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def _logo_data_uri(file_name, max_bytes, directory):
    if not isinstance(file_name, str) or not file_name or any(char in file_name for char in ("/", "\\", ":")) or file_name in (".", ".."):
        return ""
    logo = Path(__file__).resolve().parents[2] / directory / file_name
    try:
        if not logo.is_file() or logo.stat().st_size > max_bytes:
            return ""
        data = logo.read_bytes()
    except OSError:
        return ""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        mime = "image/png"
    elif data.startswith(b"\xff\xd8\xff"):
        mime = "image/jpeg"
    else:
        return ""
    return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"


def logo_html():
    name = _branding_config().get("WorkplaceCloudHubLogoPath", "WorkplaceCloudHub-lockup-WPF.png")
    uri = _logo_data_uri(name, 1024 * 1024, "")
    return f'<img src="{uri}" alt="WorkplaceCloudHub">' if uri else ""


def client_logo_html():
    """Embed the logo named by the local SharePointMigration configuration."""
    uri = _logo_data_uri(_branding_config().get("ClientLogoPath", ""), 200 * 1024, "Config")
    return f'<div class="client-brand"><img src="{uri}" alt="Client logo"></div>' if uri else ""


def render_report(title, report_kind, generated_at, status_text, status_class, cards_html, body_html, footer, alert_html="", download_html=""):
    return f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{escape(title)} | WorkplaceCloudHub</title>
<style>{REPORT_CSS}</style>
</head>
<body>
<a class="skip-link" href="#report-content">Skip to report</a>
<div class="brand-band"></div>
<main id="report-content">
  <header class="hero">
    <div class="hero-main">
      <div class="brand">{logo_html()}<div><div class="brand-name">WorkplaceCloudHub</div><div class="brand-product">SmartM365 · SharePoint Migration</div></div></div>
      {client_logo_html()}
      <div class="eyebrow">{escape(report_kind)}</div>
      <h1>{escape(title)}</h1>
      <div class="subtitle">Generated at {escape(generated_at)}</div>
    </div>
    <div class="badge {escape(status_class)}" role="status">{escape(status_text)}</div>
  </header>
  {download_html}
  <p class="intro">This summary uses the selected inventory files. Review their scan logs before accepting the comparison as complete.</p>
  {alert_html}
  <div class="metrics" aria-label="Comparison metrics">{cards_html}</div>
  {body_html}
  <footer class="footer">{escape(footer)} · WorkplaceCloudHub</footer>
</main>
</body>
</html>
'''
