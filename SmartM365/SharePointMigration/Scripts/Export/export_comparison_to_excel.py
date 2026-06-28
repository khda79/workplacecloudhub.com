import argparse
import builtins
import csv
import html
import shutil
import re
import uuid
import zipfile
from datetime import datetime
from pathlib import Path


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


SHEETS = [
    ("Summary", "Summary.csv"),
    ("LibrarySummary", "LibrarySummary.csv"),
    ("MissingInTarget", "MissingInTarget.csv"),
    ("ExtraInTarget", "ExtraInTarget.csv"),
    ("ExtraFoldersInTarget", "ExtraFoldersInTarget.csv"),
    ("DifferentSize", "DifferentSize.csv"),
    ("TargetOlderThanSource", "TargetOlderThanSource.csv"),
    ("ChangedModifiedDate", "ChangedModifiedDate.csv"),
    ("ChangedVersion", "ChangedVersion.csv"),
    ("DuplicateKeys", "DuplicateKeys.csv"),
]

SP2019_CHANGES_SHEETS = [
    ("Summary", "Summary.csv"),
    ("LibrarySummary", "LibrarySummary.csv"),
    ("RemovedFromNewScan", "MissingInTarget.csv"),
    ("AddedInNewScan", "ExtraInTarget.csv"),
    ("ChangedSize", "DifferentSize.csv"),
    ("TargetOlderThanSource", "TargetOlderThanSource.csv"),
    ("ChangedModifiedDate", "ChangedModifiedDate.csv"),
    ("ChangedVersion", "ChangedVersion.csv"),
    ("DuplicateKeys", "DuplicateKeys.csv"),
]
CSV_DELIMITERS = ",;\t"

MAX_EXCEL_ROWS = 1_048_576
INVALID_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
INTEGER_RE = re.compile(r"^-?\d+$")
DECIMAL_RE = re.compile(r"^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$")
NUMERIC_COLUMNS = {
    "SourceRows",
    "SourceFilteredRows",
    "SourceUniqueKeys",
    "SourceDuplicateKeysIgnored",
    "TargetRows",
    "TargetFilteredRows",
    "TargetUniqueKeys",
    "TargetDuplicateKeysIgnored",
    "MatchedKeys",
    "MissingInTarget",
    "ExtraInTarget",
    "ExtraFoldersInTarget",
    "TargetMatchingPathCount",
    "ExtraInTargetBytes",
    "ExtraFileCount",
    "ExtraFileBytes",
    "ExtraFileMB",
    "Depth",
    "DifferentSize",
    "ChangedModifiedDate",
    "TargetOlderThanSource",
    "ChangedVersion",
    "SizeToleranceBytes",
    "ModifiedDateToleranceMinutes",
    "SourceFiles",
    "TargetFiles",
    "MatchedFiles",
    "SourceBytes",
    "TargetBytes",
    "DifferentSizeSourceBytes",
    "DifferentSizeTargetBytes",
    "DifferentSizeDeltaBytes",
    "DeltaModifiedMinutes",
    "SizeBytes",
    "SourceSizeBytes",
    "TargetSizeBytes",
    "DeltaBytes",
    "TargetOlderByMinutes",
    "VersionsCount",
}
PERCENT_COLUMNS = {
    "MatchedFilesPercent",
    "MissingInTargetPercent",
    "ExtraInTargetPercent",
    "DifferentSizePercent",
    "ChangedModifiedDatePercent",
    "TargetOlderThanSourcePercent",
    "ChangedVersionPercent",
}
DATE_COLUMNS = {
    "Created",
    "Modified",
    "SourceModified",
    "TargetModified",
    "SourceModifiedNormalizedUtc",
    "TargetModifiedNormalizedUtc",
}
DATE_FORMATS = (
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d",
)
EXCEL_EPOCH = datetime(1899, 12, 30)
DEFAULT_COLUMN_WIDTH = 14
MAX_COLUMN_WIDTH = 80
COLUMN_WIDTHS = {
    "Status": 26,
    "LibraryKey": 28,
    "WebPath": 18,
    "LibraryPath": 34,
    "SourceWebUrl": 58,
    "TargetWebUrl": 58,
    "SourceWebTitle": 22,
    "TargetWebTitle": 22,
    "SourceLibraryTitle": 24,
    "TargetLibraryTitle": 24,
    "SourceFiles": 13,
    "TargetFiles": 13,
    "MatchedFiles": 14,
    "MatchedFilesPercent": 19,
    "MissingInTarget": 16,
    "MissingInTargetPercent": 22,
    "ExtraInTarget": 15,
    "ExtraInTargetPercent": 20,
    "DifferentSize": 15,
    "DifferentSizePercent": 20,
    "ChangedModifiedDate": 22,
    "ChangedModifiedDatePercent": 26,
    "TargetOlderThanSource": 24,
    "TargetOlderThanSourcePercent": 30,
    "ChangedVersion": 18,
    "ChangedVersionPercent": 22,
    "MissingReason": 18,
    "TargetMatchingPathCount": 24,
    "TargetVersions": 24,
    "TargetVersionComparisons": 42,
    "ServerRelativeUrl": 76,
    "TargetServerRelativeUrl": 76,
    "SourceServerRelativeUrl": 76,
    "FileUrl": 76,
    "TargetFileUrl": 76,
    "SourceFileUrl": 76,
    "FileName": 36,
    "Directory": 58,
    "LibraryTitle": 24,
    "WebUrl": 58,
    "Created": 20,
    "Modified": 20,
    "SourceModified": 20,
    "TargetModified": 20,
    "SourceModifiedNormalizedUtc": 29,
    "TargetModifiedNormalizedUtc": 29,
    "CreatedBy": 24,
    "ModifiedBy": 24,
    "SizeBytes": 16,
    "SourceSizeBytes": 18,
    "TargetSizeBytes": 18,
    "DeltaBytes": 16,
    "TargetOlderByMinutes": 22,
    "SourceVersion": 16,
    "TargetVersion": 16,
    "Version": 14,
    "VersionsCount": 16,
    "VersionComparison": 22,
    "SourceBytes": 18,
    "TargetBytes": 18,
    "DifferentSizeSourceBytes": 24,
    "DifferentSizeTargetBytes": 24,
    "DifferentSizeDeltaBytes": 23,
    "DeltaModifiedMinutes": 22,
}


def clean_xml_text(value):
    if value is None:
        return ""
    return INVALID_XML_RE.sub("", str(value))


def detect_csv_dialect(handle):
    sample = handle.read(65536)
    handle.seek(0)
    if not sample:
        return csv.excel

    try:
        return csv.Sniffer().sniff(sample, delimiters=CSV_DELIMITERS)
    except csv.Error:
        return csv.excel


def column_name(index):
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def parse_excel_datetime(value):
    text = clean_xml_text(value).strip()
    if not text:
        return None

    text = text.replace("Z", "").split(".", 1)[0]
    for date_format in DATE_FORMATS:
        try:
            parsed = datetime.strptime(text, date_format)
            delta = parsed - EXCEL_EPOCH
            return delta.days + (delta.seconds / 86400) + (delta.microseconds / 86400000000)
        except ValueError:
            continue

    return None


def cell_xml(row_index, column_index, value, is_numeric=False, is_date=False, is_percent=False):
    ref = f"{column_name(column_index)}{row_index}"
    text = clean_xml_text(value).strip()
    if is_date:
        excel_datetime = parse_excel_datetime(text)
        if excel_datetime is not None:
            return f'<c r="{ref}" s="1"><v>{excel_datetime:.10f}</v></c>'

    if is_percent and DECIMAL_RE.match(text):
        return f'<c r="{ref}" s="2"><v>{text}</v></c>'

    if is_numeric and DECIMAL_RE.match(text):
        return f'<c r="{ref}"><v>{text}</v></c>'

    text = html.escape(clean_xml_text(value), quote=False)
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def infer_column_width(header):
    if header in COLUMN_WIDTHS:
        return COLUMN_WIDTHS[header]
    if header in NUMERIC_COLUMNS:
        return 14
    if header in PERCENT_COLUMNS:
        return 18
    if header in DATE_COLUMNS:
        return 20
    if header.endswith("Url"):
        return 58
    if "Title" in header:
        return 24
    return min(max(len(header) + 3, DEFAULT_COLUMN_WIDTH), MAX_COLUMN_WIDTH)


def write_column_widths(writer, header):
    if not header:
        return
    writer.write("<cols>")
    for column_index, column_header in enumerate(header, start=1):
        width = infer_column_width(column_header)
        writer.write(
            f'<col min="{column_index}" max="{column_index}" '
            f'width="{width}" customWidth="1"/>'
        )
    writer.write("</cols>\n")


def sheet_xml_start(writer, header=None):
    writer.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    writer.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ')
    writer.write('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n')
    writer.write(
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>\n'
    )
    write_column_widths(writer, header or [])
    writer.write("<sheetData>\n")


def sheet_xml_end(writer, row_count, column_count):
    writer.write("</sheetData>\n")
    if row_count > 0 and column_count > 0:
        last_ref = f"{column_name(column_count)}{row_count}"
        writer.write(f'<autoFilter ref="A1:{last_ref}"/>\n')
    writer.write("</worksheet>\n")


def write_sheet_from_csv(csv_path, xml_path):
    row_count = 0
    column_count = 0
    truncated = False

    with csv_path.open("r", encoding="utf-8-sig", newline="") as csv_handle, xml_path.open(
        "w", encoding="utf-8", newline=""
    ) as xml_handle:
        reader = csv.reader(csv_handle, dialect=detect_csv_dialect(csv_handle))
        header = next(reader, [])
        column_count = len(header)
        sheet_xml_start(xml_handle, header)

        if header:
            row_count = 1
            xml_handle.write(f'<row r="{row_count}">')
            for column_index, value in enumerate(header, start=1):
                xml_handle.write(cell_xml(row_count, column_index, value))
            xml_handle.write("</row>\n")

        for row in reader:
            row_count += 1
            if row_count > MAX_EXCEL_ROWS:
                truncated = True
                row_count -= 1
                break

            column_count = max(column_count, len(row))
            xml_handle.write(f'<row r="{row_count}">')
            for column_index, value in enumerate(row, start=1):
                column_name_value = header[column_index - 1] if column_index <= len(header) else ""
                is_numeric = row_count > 1 and column_name_value in NUMERIC_COLUMNS
                is_date = row_count > 1 and column_name_value in DATE_COLUMNS
                is_percent = row_count > 1 and column_name_value in PERCENT_COLUMNS
                xml_handle.write(cell_xml(row_count, column_index, value, is_numeric, is_date, is_percent))
            xml_handle.write("</row>\n")

        sheet_xml_end(xml_handle, row_count, column_count)

    return row_count, column_count, truncated


def workbook_xml(sheet_names):
    sheets_xml = []
    for index, sheet_name in enumerate(sheet_names, start=1):
        escaped_name = html.escape(sheet_name, quote=True)
        sheets_xml.append(f'<sheet name="{escaped_name}" sheetId="{index}" r:id="rId{index}"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        + "".join(sheets_xml)
        + "</sheets></workbook>"
    )


def workbook_rels(sheet_count):
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + "</Relationships>"
    )


def content_types(sheet_count):
    overrides = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(overrides)
        + "</Types>"
    )


def root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def styles_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="2">'
        '<numFmt numFmtId="164" formatCode="dd/mm/yyyy hh:mm:ss"/>'
        '<numFmt numFmtId="165" formatCode="0.00%"/>'
        '</numFmts>'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '</cellXfs>'
        '</styleSheet>'
    )


def doc_props_app(sheet_names):
    titles = "".join(f"<vt:lpstr>{html.escape(name)}</vt:lpstr>" for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Python</Application>'
        f'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
        f'<TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>'
        '</Properties>'
    )


def doc_props_core():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:creator>SharePointInventory</dc:creator>'
        '<cp:lastModifiedBy>SharePointInventory</cp:lastModifiedBy>'
        '</cp:coreProperties>'
    )


def export_workbook(comparison_directory, output_path, sheets=None):
    comparison_directory = Path(comparison_directory)
    output_path = Path(output_path)
    if sheets is None:
        sheets = SHEETS
    sheet_names = []
    sheet_files = []
    stats = []
    temp_files = []

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_directory = output_path.parent / ".excel-temp"
    if temp_directory.exists():
        shutil.rmtree(temp_directory, ignore_errors=True)
    temp_directory.mkdir(parents=True, exist_ok=True)
    temp_prefix = f"{output_path.stem}-{uuid.uuid4().hex}"
    try:
        for index, (sheet_name, csv_name) in enumerate(sheets, start=1):
            csv_path = comparison_directory / csv_name
            if not csv_path.exists():
                continue
            xml_path = temp_directory / f"{temp_prefix}-sheet{index}.xml.tmp"
            row_count, column_count, truncated = write_sheet_from_csv(csv_path, xml_path)
            sheet_names.append(sheet_name)
            sheet_files.append(xml_path)
            temp_files.append(xml_path)
            stats.append((sheet_name, row_count, column_count, truncated))

        if not sheet_files:
            raise FileNotFoundError(f"No comparison CSV files found in {comparison_directory}")

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            archive.writestr("[Content_Types].xml", content_types(len(sheet_files)))
            archive.writestr("_rels/.rels", root_rels())
            archive.writestr("xl/workbook.xml", workbook_xml(sheet_names))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(len(sheet_files)))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app(sheet_names))
            archive.writestr("docProps/core.xml", doc_props_core())
            for index, sheet_file in enumerate(sheet_files, start=1):
                archive.write(sheet_file, f"xl/worksheets/sheet{index}.xml")
    finally:
        for temp_file in temp_files:
            try:
                temp_file.unlink()
            except (FileNotFoundError, PermissionError):
                pass
        try:
            temp_directory.rmdir()
        except (FileNotFoundError, OSError):
            pass

    return stats


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-xlsx", required=True)
    parser.add_argument(
        "--sheet-profile",
        choices=["default", "sp2019-changes"],
        default="default",
        help="Use default SP2019-vs-SPO sheet names or SP2019 change-oriented sheet names.",
    )
    args = parser.parse_args()

    sheets = SP2019_CHANGES_SHEETS if args.sheet_profile == "sp2019-changes" else SHEETS
    stats = export_workbook(args.comparison_directory, args.output_xlsx, sheets=sheets)
    print(f"Excel file created: {args.output_xlsx}")
    for sheet_name, row_count, column_count, truncated in stats:
        marker = " (truncated at Excel row limit)" if truncated else ""
        print(f"{sheet_name}: {row_count} rows, {column_count} columns{marker}")


if __name__ == "__main__":
    main()
