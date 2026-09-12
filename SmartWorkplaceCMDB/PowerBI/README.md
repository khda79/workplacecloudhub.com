# SmartWorkplaceCMDB Power BI — V1

The canonical DSI cockpit is maintained in place with `report_cockpit.py`.
The approved V1 navigation exposes seven decision pages and retains three
hidden 360 drill-through pages. The transformation preserves the three stable
drill-through identifiers, consolidates overlapping content and retires six
redundant presentation pages.
All visible release labels and the canonical project name are V1; the project
is stored as `PowerBI/CMDB-REPORTS/CMDB-REPORTS.pbip`. See
[COCKPIT-DSI.md](COCKPIT-DSI.md) for metric and source boundaries.

Power BI is a native deliverable of SmartWorkplaceCMDB. The report builder
creates six English pages: Overview, Devices and compliance, Users and
relationships, Licenses and assignments, Mailboxes, and Quality and coverage.
All report labels, measure names and descriptions, generated status values,
and operator instructions are English. The model culture is en-US. Business
data such as names and department values retain their original source text.

Add `--include-360` to generate nine pages, including Device 360, User 360 and
Group 360 from existing local raw and CMDB exports. This optional enrichment
requires the sibling `report_360.py`; the default six-page build remains
compatible with existing curated-only inputs.

## Local report preparation

Run PowerBI/build_report.py with Python 3.10 or later; only the standard library
is required. The builder reads a frozen DATA-LAST directory and creates a new
private PBIP project. It never connects to a tenant, installs dependencies,
collects data, publishes a report or overwrites an existing destination.

The following paths are **fictional placeholders**. Replace them with authorized
local paths. Run from the repository root:

    python .\SmartWorkplaceCMDB\PowerBI\build_report.py --data-root ".\SmartWorkplaceCMDB\Data\Tenants\prod\DATA-LAST" --output ".\SmartWorkplaceCMDB\PowerBI\CMDB-REPORTS-NEW"

For the 360 version (the paths are still fictional placeholders):

    python .\SmartWorkplaceCMDB\PowerBI\build_report.py --include-360 --data-root ".\SmartWorkplaceCMDB\Data\Tenants\prod\DATA-LAST" --output ".\SmartWorkplaceCMDB\PowerBI\CMDB-REPORTS-NEW"

The output must be a new directory outside the source DATA-LAST. Open the
generated `.pbip` in Power BI Desktop, refresh the local CSV files, check all
generated pages, then save a `.pbix` if that distribution form is required. Preserve the previous
PBIX until the replacement has been validated; replace it when previous
versions are no longer needed. The modeling MCP cannot save the report as a PBIX; that Desktop step
requires the operator. Computer Use is not required.

ReportData is a frozen copy for this report. Refresh only rereads those local
files; it does not collect new tenant data or follow changes in the original
DATA-LAST. Generate a new output folder for a new authorized collection.
If a generated project is moved, update and validate its Power Query source
paths before opening it in Desktop.

## Model and source contract

Each curated table except DimDate starts with TenantKey, OrganizationKey,
EnvironmentKey and TenantId. Relationships use tenant-scoped keys such as
TenantUserKey, TenantDeviceKey and TenantSkuKey. The builder checks contracts,
all four identity fields, unique keys, types and all six relationships.
Power Query rechecks headers and tenant identity on import. Blank values stay
null; timestamp fields require a timezone and are converted to UTC.

The eleven business tables are retained. Mailbox and quality details are joined
from matching CMDB CSV rows. SourceHealth adds a disconnected source-evidence
table. No CSV source contract is changed. Source evidence compares local CSV
hashes and row counts; missing or inconsistent evidence remains visible.
SourceHealth ignores finding filters. DimDate and DimTenant remain hidden and
disconnected: no historical trend, global tenant selector or RLS is claimed.

## Readability

Cards, chart data labels and axes use explicit full-number display units.
Cards use the supported value.labelDisplayUnits property set to 1 (None),
not the discarded value.displayUnits property. Internal card outlines are
disabled and content margins are explicit. Textboxes have zero container
padding, larger fonts and enough space for their estimated line count.
Table values and headers wrap onto multiple lines.

Display-only Label columns render missing text as Not provided. Original
columns, nulls, numeric types, keys and all 23 measure formulas remain
unchanged. Source rows and timestamps also have readable display columns.
Charts and slicers exclude only the synthetic blank relationship member;
every physical row has a nonblank display label, so missing source values
remain visible as Not provided. This prevents a zero-count blank category
from appearing as a localized empty label. No global data filter is added.

Official schema and data checks do not replace an operator check of Desktop
rendering. Review full numbers, long labels, all six pages and a Not provided
selection after opening and refreshing the prepared project.

## Device, User and Group 360

The optional enrichment adds three report-owned tables, for 15 tables and 13
single-direction relationships. It preserves the eleven business tables,
SourceHealth and the original 23 measure definitions. Sixteen additional
measures count source/assignment paths or gate detail tables to one selected
entity. These tables are local report transformations, not new collector CSV
contracts or new tenant queries.

- Device 360 shows identity, compliance, the associated account and its link
  status, source identifiers, management agent, enrollment, source-specific
  activity dates and entity-linked findings. Repeated Intune source candidates
  remain visible; selection follows the existing normalizer's ordering and is
  checked against its output. All source candidates are retained.
- User 360 shows job title, department, account creation text, associated
  devices/mailboxes and every license assignment path, including group origin,
  state, errors, last update and disabled plan IDs. These are assignment paths,
  not unique user/SKU pairs or evidence of effective usage.
- Group 360 shows identity and observed license paths, plus a coverage filter
  to find groups with exported paths. Cloud members and owners remain Not
  collected. No observed path never means unused or safe to delete.

Search the unique selector on each 360 page. Single selection is enabled, and
detail measures use ALLSELECTED to avoid displaying all objects when there is
no single selected entity. Names alone are never keys. Right-click the unique
selection column in the device/user inventories to drill through. The User 360
device table can open Device 360; its group selection column can open Group
360; the Group 360 user selection can open User 360. Page tabs return to the
overview. Native navigation still requires operator validation in Desktop.

New timestamp columns preserve the original text beside a qualified UTC value
and status. Dates without an explicit timezone stay unqualified, and pre-1900
sentinels stay unknown. No local-culture guess or activity threshold is used.
Existing source columns remain unchanged. Collection time and activity time
are separate. Account creation text without a timezone cannot establish an
exact instant. Assignment last updated is not an initial assignment date.

Missing associations, unresolved references and missing collection are distinct
states. Only exact tenant-qualified entity mappings link quality findings;
dataset findings and other entity types remain on the global quality page.
ConfidenceScore remains the unmodified configuration value, not a measured
trust percentage. No change journal, hardware enrichment, membership inventory
or business dependency graph is introduced in this increment.

The enriched build fails before creating output if required local source files
are absent, source identity differs, source keys are duplicated, entity sets
drift, selected Intune evidence disagrees with CMDB, or assignment paths cannot
reconcile to the existing user/SKU summary. Source hashes are checked before
writing the generated project. Raw IDs and names remain private in ReportData.

## Metric definitions and filters

- Enabled accounts do not prove activity or sign-ins. Unknown account status
  is not classified as disabled.
- Users with assignments include all assignment states. This is not a count
  of active or used licenses.
- Capacity is displayed only for a single SKU. Do not sum different products
  as a count of people. The assignment-state filter does not filter capacity.
- Compliant device share is explicitly compliant devices divided by all
  filtered devices, including unknown states in the denominator. Missing
  compliance is not classified as noncompliant.
- User and device filters propagate through the six single-direction
  relationships. Charts and long tables use native scrolling.
- A DiscoveryMailbox without an external identifier may produce an
  informational finding under the corroborated local quality rule. That
  finding remains visible with its stable key.
- Complete means the collector reported complete coverage for the granted
  permissions. It does not certify tenant completeness or AD access.

REPORT-MANIFEST.json includes all 23 measure definitions, CSV-derived expected
values and input/output hashes. VALIDATION-MEASURES.dax uses DEFINE MEASURE to
check those calculations without changing the imported model.

## Validation and handling

The hardware reporting adapter and the canonical cockpit's hardware pages are
documented in [HARDWARE-REPORT.md](HARDWARE-REPORT.md). Hardware evidence is joined
through native CI/device keys and retains separate source dates. The current
pilot includes a validated real hardware snapshot. Devices without a matching
hardware record still show an explicit unavailable message. The generic six/nine-page generator below remains
separate from this additive hardware page.

The canonical cockpit also has an additive `10  Hardware fleet` page;
the source-evidence page is retained as `11  Hardware detail`. The report and
its visible context labels use V1 and no page title contains BETA. Fleet
coverage uses filtered `DimDevice` rows as its denominator and removes direct
manufacturer/model/serial/storage filters from its numerator. The distribution
charts and anomaly cards remain in hardware-record scope. Missing attributes,
source-reported zero storage and repeated reported serial values are separate
checks; none proves identity, ownership, physical capacity or a safe merge.
Manufacturer/model charts use native scrolling, and the fleet detail table
retains the DeviceSelection field for the existing Device 360 drillthrough.

The canonical private snapshot has passed prior native Desktop reload and
screenshot review. A fresh post-finalization reload is required before release
closure. Rendering evidence qualifies only the reviewed snapshot; screenshots
do not qualify native clicks.
When performing an operator check, target only the exact canonical PBIP/PID and
verify page filters, chart cross-filtering, table scrolling and the three 360
drill-through targets. Device 360 includes hardware equipment and source-date
evidence. Preserve the current detail selection unless the operator explicitly
changes it.

Run the synthetic offline suite without network access:

    python -B -m unittest discover -s .\SmartWorkplaceCMDB\Tests -p "test_report*.py" -v

Apply the cockpit only to an authorized private PBIR folder after making a
backup. The following path is a fictional placeholder:

    python .\SmartWorkplaceCMDB\PowerBI\report_cockpit.py --report ".\SmartWorkplaceCMDB\PowerBI\CMDB-REPORTS\CMDB-REPORTS.Report"

Schema validation and DAX checks do not prove Desktop rendering. After import,
check every generated page, visual errors, exact unfiltered totals and one filter on
each applicable page. Stop on import errors, unexpected sources, another
tenant or unexplained totals. Preserve evidence before changing anything.

For 360 pages, validate with an actual filter on DeviceSelection,
UserSelection or GroupSelection and group by every column in each detail
visual. Filtering an entire dimension row is not an equivalent regression
test: it previously concealed an ALLSELECTED key-column defect that returned
blank detail measures despite a valid selector choice. Detail gates now use
ALLSELECTED on the whole dimension. Test one selection, two selections and
no selection; only a single selected entity may show detail records. Missing
related records must remain empty. Report projections set displayName for
English captions; nativeQueryRef alone does not rename a visible header.
Dropdown containers reserve room for a full row below their titles.

The 360 layout uses taller profiles, fixed business-column widths and no
table totals. Detail gates and long navigation/path keys remain in the
visual query as hidden projections, preserving selection logic and row
identity. Names and accounts are displayed beside those hidden navigation
fields. Hiding a field is presentation only, not a security boundary.
Check the rendered rows and cross-page drillthrough in Desktop before
replacing the PBIX; schema and DAX validation do not validate the UI.

The CSV, model, PBIP, PBIX and caches contain personal data. Keep them in the
authorized private folder, outside Git and publishing/sharing workflows.
Hiding technical keys in the model is not data security. Keep ReportData while
the project is in use. After retention is agreed, cleanup must target only
the approved review folder with Desktop closed; preserve original snapshots.

The V1 generators, tests and documentation are included in the stable local
release allowlist. No report, dataset or private source data is published by
the build or by this finalization workflow.
