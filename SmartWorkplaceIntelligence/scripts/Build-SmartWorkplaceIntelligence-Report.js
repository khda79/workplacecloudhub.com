#!/usr/bin/env node
'use strict';

// Deterministic PBIR generator for the one canonical Smart Workplace
// Intelligence report. It writes report metadata only; the semantic model is
// authored and exported separately through the Power BI Modeling MCP.

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const reportRoot = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report');
const definitionRoot = path.join(reportRoot, 'definition');
const pagesRoot = path.join(definitionRoot, 'pages');
const resourcesRoot = path.join(reportRoot, 'StaticResources', 'RegisteredResources');

const schemas = {
  report: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json',
  pages: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json',
  page: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json',
  visual: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json',
  version: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json'
};

function ensureInside(target) {
  const resolved = path.resolve(target);
  if (resolved !== root && !resolved.startsWith(`${root}${path.sep}`)) throw new Error(`Refusing to write outside product root: ${resolved}`);
}

function writeJson(target, value) {
  ensureInside(target);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const content = `${JSON.stringify(value, null, 2)}\r\n`;
  if (fs.existsSync(target)) {
    const existing = fs.readFileSync(target, 'utf8');
    if (existing === content) return;
    try {
      if (JSON.stringify(JSON.parse(existing)) === JSON.stringify(value)) return;
    } catch { /* Rewrite invalid or non-JSON content below. */ }
  }
  fs.writeFileSync(target, content, 'utf8');
}

function stableId(seed) { return crypto.createHash('sha1').update(seed).digest('hex').slice(0, 20); }
function literal(value, suffix = '') {
  if (typeof value === 'boolean') return { expr: { Literal: { Value: String(value) } } };
  if (typeof value === 'number') return { expr: { Literal: { Value: `${value}${suffix || 'D'}` } } };
  return { expr: { Literal: { Value: `'${String(value).replaceAll("'", "''")}'` } } };
}
function color(value) { return { solid: { color: literal(value) } }; }
function position(x, y, width, height, z) { return { x, y, z, width, height, tabOrder: z }; }

function measureField(table, name, label = name) {
  return { field: { Measure: { Expression: { SourceRef: { Entity: table } }, Property: name } }, queryRef: `${table}.${name}`, nativeQueryRef: label };
}
function columnField(table, name, label = name) {
  return { field: { Column: { Expression: { SourceRef: { Entity: table } }, Property: name } }, queryRef: `${table}.${name}`, nativeQueryRef: label };
}
function fieldProjection(field) { return field.kind === 'measure' ? measureField(field.table, field.name, field.label) : columnField(field.table, field.name, field.label); }
const M = (table, name, label = name) => ({ kind: 'measure', table, name, label });
const C = (table, name, label = name) => ({ kind: 'column', table, name, label });

function chrome(title, options = {}) {
  return {
    ...(options.altText ? { general: [{ properties: { altText: literal(options.altText) } }] } : {}),
    title: [{ properties: { show: literal(Boolean(title)), ...(title ? { text: literal(title), fontSize: literal(options.titleSize || 13), fontColor: color('#1F2937'), bold: literal(true) } : {}) } }],
    background: [{ properties: { show: literal(options.showBackground !== false), color: color(options.background || '#FFFFFF'), transparency: literal(0) } }],
    border: [{ properties: { show: literal(options.showBorder !== false), color: color(options.border || '#CBD8E5'), width: literal(options.borderWidth || 1), radius: literal(options.radius || 12) } }],
    visualHeader: [{ properties: { show: literal(false) } }],
    padding: [{ properties: { top: literal(options.padding ?? 8), bottom: literal(options.padding ?? 8), left: literal(options.padding ?? 10), right: literal(options.padding ?? 10) } }]
  };
}

function image(seed, itemName, altText, x, y, width, height, z) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'image',
      objects: {
        general: [{ properties: { imageUrl: { expr: { ResourcePackageItem: { PackageName: 'RegisteredResources', PackageType: 1, ItemName: itemName } } } } }]
      },
      visualContainerObjects: {
        general: [{ properties: { altText: literal(altText) } }],
        background: [{ properties: { show: literal(false) } }],
        border: [{ properties: { show: literal(false) } }],
        visualHeader: [{ properties: { show: literal(false) } }],
        padding: [{ properties: { top: literal(0), bottom: literal(0), left: literal(0), right: literal(0) } }]
      },
      drillFilterOtherVisuals: true
    }
  };
}

const iconFiles = [
  'smartworkplace-device-20260913.svg',
  'smartworkplace-health-20260913.svg',
  'smartworkplace-kpi-analytics-20260913.svg',
  'smartworkplace-kpi-compliance-20260913.svg',
  'smartworkplace-kpi-devices-20260913.svg',
  'smartworkplace-kpi-licenses-20260913.svg',
  'smartworkplace-kpi-mailboxes-20260913.svg',
  'smartworkplace-kpi-quality-20260913.svg',
  'smartworkplace-kpi-relationships-20260913.svg',
  'smartworkplace-kpi-users-20260913.svg',
  'smartworkplace-licenses-20260913.svg',
  'smartworkplace-lifecycle-20260913.svg',
  'smartworkplace-overview-20260913.svg',
  'smartworkplace-people-20260913.svg',
  'smartworkplace-ratio-mailboxes-20260913.svg',
  'smartworkplace-section-activity-20260913.svg',
  'smartworkplace-section-findings-20260913.svg',
  'smartworkplace-section-hardware-20260913.svg',
  'smartworkplace-section-identity-20260913.svg',
  'smartworkplace-services-20260913.svg'
];

const pageIcons = {
  'Executive Overview': 'smartworkplace-overview-20260913.svg',
  'Workforce & Identity': 'smartworkplace-people-20260913.svg',
  'Licensing & Cost': 'smartworkplace-licenses-20260913.svg',
  'Devices & Compliance': 'smartworkplace-device-20260913.svg',
  'Windows Lifecycle': 'smartworkplace-lifecycle-20260913.svg',
  'Endpoint Experience': 'smartworkplace-kpi-analytics-20260913.svg',
  'Applications & Standardization': 'smartworkplace-section-hardware-20260913.svg',
  'Collaboration Adoption': 'smartworkplace-services-20260913.svg',
  'Content & Storage': 'smartworkplace-section-activity-20260913.svg',
  'Messaging & Hybrid': 'smartworkplace-ratio-mailboxes-20260913.svg',
  'Backup & Resilience': 'smartworkplace-health-20260913.svg',
  'Security Posture': 'smartworkplace-health-20260913.svg',
  'Data Trust': 'smartworkplace-kpi-quality-20260913.svg',
  'Risk & Incident Signals': 'smartworkplace-section-findings-20260913.svg',
  'User Explorer': 'smartworkplace-people-20260913.svg',
  'Device Explorer': 'smartworkplace-device-20260913.svg',
  'Application Explorer': 'smartworkplace-section-hardware-20260913.svg',
  'Service & Content Explorer': 'smartworkplace-services-20260913.svg',
  'Mailbox Explorer': 'smartworkplace-ratio-mailboxes-20260913.svg',
  'Control & Evidence Explorer': 'smartworkplace-kpi-quality-20260913.svg'
};

function textbox(seed, text, x, y, width, height, z, style = {}) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'textbox',
      objects: { general: [{ properties: { paragraphs: [{ textRuns: [{ value: text, textStyle: { fontFamily: style.bold ? 'Segoe UI Semibold' : 'Segoe UI', fontSize: style.size || '12px', color: style.color || '#52606D', ...(style.bold ? { fontWeight: 'bold' } : {}) } }], horizontalTextAlignment: style.align || 'left' }] } }] },
      visualContainerObjects: {
        background: [{ properties: { show: literal(false) } }],
        border: [{ properties: { show: literal(false), radius: literal(0) } }],
        visualHeader: [{ properties: { show: literal(false) } }],
        padding: [{ properties: { top: literal(0), bottom: literal(0), left: literal(0), right: literal(0) } }]
      }
    }
  };
}

function card(seed, fields, title, x, y, width, height, z, options = {}) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    ...(options.filters?.length ? { filterConfig: { filters: options.filters } } : {}),
    visual: {
      visualType: 'cardVisual',
      query: { queryState: { Data: { projections: fields.map(fieldProjection) } } },
      objects: {
        value: [{ properties: { show: literal(true), fontSize: literal(options.valueSize || 24), bold: literal(true), fontColor: color(options.valueColor || '#005A9E'), labelDisplayUnits: literal(options.displayUnits ?? 0), textWrap: literal(true) }, selector: { id: 'default' } }],
        label: [{ properties: { show: literal(options.showLabel !== false), fontSize: literal(options.labelSize || 12), fontColor: color('#52606D'), textWrap: literal(true) }, selector: { id: 'default' } }],
        outline: [{ properties: { show: literal(false) }, selector: { id: 'default' } }],
        ...(options.accentColor ? { accentBar: [{ properties: { show: literal(true), position: literal('Left'), width: literal(4), color: color(options.accentColor), transparency: literal(0, 'L') }, selector: { id: 'default' } }] } : {}),
        padding: [{ properties: { paddingIndividual: literal(true), topMargin: literal(0, 'L'), bottomMargin: literal(0, 'L'), leftMargin: literal(8, 'L'), rightMargin: literal(8, 'L') }, selector: { id: 'default' } }],
        layout: [{ properties: { topOuterMargin: literal(0, 'L'), bottomOuterMargin: literal(0, 'L'), leftOuterMargin: literal(0, 'L'), rightOuterMargin: literal(0, 'L'), paddingUniform: literal(0, 'L') }, selector: { id: 'default' } }],
        spacing: [{ properties: { verticalSpacing: literal(0) }, selector: { id: 'default' } }],
        ...(fields.length > 1 ? { cardCalloutArea: [{ properties: { show: literal(true), paddingUniform: literal(8, 'L'), rectangleRoundedCurve: literal(6, 'L'), backgroundFillColor: color('#F7FAFC'), backgroundTransparency: literal(0) } }] } : {})
      },
      visualContainerObjects: chrome(title, options.bare
        ? { ...options, showBackground: false, showBorder: false, padding: 4, altText: options.altText || `${title || 'Key indicators'}: ${fields.map(field => field.label || field.name).join(', ')}.` }
        : { ...options, altText: options.altText || `${title || 'Key indicators'}: ${fields.map(field => field.label || field.name).join(', ')}.` })
    }
  };
}

function shapeContainer(seed, x, y, width, height, z, options = {}) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'shape',
      objects: {
        shape: [{ properties: { tileShape: literal('rectangle'), roundEdge: literal(options.radius || 12, 'L') }, selector: { id: 'default' } }],
        fill: [{ properties: { fillColor: color(options.background || '#FFFFFF'), transparency: literal(0) }, selector: { id: 'default' } }],
        outline: [{ properties: { show: literal(true), lineColor: color(options.border || '#CBD8E5'), weight: literal(options.borderWidth || 1), transparency: literal(0) }, selector: { id: 'default' } }]
      },
      visualContainerObjects: {
        background: [{ properties: { show: literal(false) } }],
        border: [{ properties: { show: literal(false) } }],
        visualHeader: [{ properties: { show: literal(false) } }],
        padding: [{ properties: { top: literal(0), bottom: literal(0), left: literal(0), right: literal(0) } }]
      }
    }
  };
}

function sparkline(seed, category, value, tooltips, x, y, width, height, z, accentColor, options = {}) {
  const name = stableId(seed);
  const valueProjection = measureField(value.table, value.name);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'lineChart',
      query: {
        queryState: {
          Category: { projections: [columnField(category.table, category.name)] },
          Y: { projections: [valueProjection] },
          Tooltips: { projections: tooltips.map(fieldProjection) }
        },
        sortDefinition: { sort: [{ field: columnField(category.table, category.name).field, direction: 'Ascending' }], isDefaultSort: false }
      },
      objects: {
        categoryAxis: [{ properties: { show: literal(Boolean(options.showCategoryAxis)), showAxisTitle: literal(false), gridlineShow: literal(false), axisType: literal(options.axisType || 'Categorical'), fontSize: literal(options.axisFontSize || 9), labelColor: color('#52606D') } }],
        valueAxis: [{ properties: { show: literal(false), showAxisTitle: literal(false), gridlineShow: literal(false), scaleToFit: literal(true) } }],
        legend: [{ properties: { show: literal(false) } }],
        dataPoint: [{ properties: { defaultColor: color(accentColor) } }],
        lineStyles: [{ properties: { strokeWidth: literal(3), showMarker: literal(false) } }]
      },
      visualContainerObjects: {
        general: [{ properties: { altText: literal(options.altText || 'Trend over time for the current filter context.') } }],
        title: [{ properties: { show: literal(false) } }],
        subTitle: [{ properties: { show: literal(false) } }],
        background: [{ properties: { show: literal(false) } }],
        border: [{ properties: { show: literal(false) } }],
        visualHeader: [{ properties: { show: literal(false) } }],
        padding: [{ properties: { top: literal(0), bottom: literal(0), left: literal(0), right: literal(0) } }],
        visualTooltip: [{ properties: { show: literal(true) } }]
      },
      drillFilterOtherVisuals: true
    }
  };
}

function azureMap(seed, category, size, tooltips, title, x, y, width, height, z) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'azureMap',
      query: {
        queryState: {
          Category: { projections: [{ ...columnField(category.table, category.name), active: true }] },
          Size: { projections: [{ ...measureField(size.table, size.name), active: true }] },
          Tooltips: { projections: tooltips.map(field => ({ ...measureField(field.table, field.name), active: true })) }
        }
      },
      objects: {
        bubbleLayer: [{ properties: { show: literal(true), sizeByValue: literal(true), fillColor: color('#1565C0'), mapTransparency: literal(12), borderShow: literal(true), autoStrokeColor: literal(true), minBubbleRadius: literal(8, 'L'), maxRadius: literal(28, 'L'), clusteringEnabled: literal(false) } }],
        categoryLabels: [{ properties: { show: literal(true), fontFamily: literal('Segoe UI'), fontSize: literal(10), bold: literal(true), color: color('#172B4D'), enableBackground: literal(false) } }],
        mapControls: [{ properties: { defaultStyle: literal('grayscale_light'), showStylePicker: literal(false), showLabels: literal(true), showNavigationControls: literal(true), autoZoom: literal(true), worldWrap: literal(false), showSelectionControl: literal(false), showCountryRegionBorders: literal(true), showAdminDistrictBorders: literal(false), showAdminDistrict2Borders: literal(false), showBuildingFootprints: literal(false), showRoadDetails: literal(false), geocodingCulture: literal('en-GB') } }]
      },
      visualContainerObjects: chrome(title, { altText: `${title}. Geographic distribution for the current filter context.` })
    }
  };
}

function slicer(seed, table, column, title, x, y, width, z) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, 80, z),
    visual: {
      visualType: 'slicer',
      query: { queryState: { Values: { projections: [columnField(table, column)] } } },
      objects: {
        data: [{ properties: { mode: literal('Dropdown') } }],
        header: [{ properties: { show: literal(true), text: literal(title), textSize: literal(10), fontColor: color('#1F2937') } }]
      },
      visualContainerObjects: chrome('', { altText: `Filter by ${title}.` })
    }
  };
}

function bar(seed, category, value, title, x, y, width, height, z, options = {}) {
  const name = stableId(seed);
  const valueProjection = measureField(value.table, value.name);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'barChart',
      query: {
        queryState: { Category: { projections: [columnField(category.table, category.name)] }, Y: { projections: [valueProjection] } },
        sortDefinition: { sort: [{ field: valueProjection.field, direction: 'Descending' }], isDefaultSort: false }
      },
      objects: {
        dataPoint: [{ properties: { defaultColor: color('#0078D4') } }],
        categoryAxis: [{ properties: { labelColor: color('#334155'), fontSize: literal(options.categoryFontSize || 10), innerPadding: literal(24, 'L'), ...(options.categoryLabelMaxMargin ? { maxMarginFactor: literal(options.categoryLabelMaxMargin, 'L') } : {}) } }],
        valueAxis: [{ properties: { start: literal(0), labelColor: color('#52606D'), gridlineColor: color('#E7EEF5') } }],
        labels: [{ properties: { show: literal(true), fontSize: literal(10), color: color('#1F2937'), labelDisplayUnits: literal(0) } }]
      },
      visualContainerObjects: chrome(title, { altText: `${title}. Ranked comparison for the current filter context.` }),
      drillFilterOtherVisuals: true
    }
  };
}

function line(seed, category, value, title, x, y, width, height, z) {
  const name = stableId(seed);
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'lineChart',
      query: {
        queryState: { Category: { projections: [columnField(category.table, category.name)] }, Y: { projections: [measureField(value.table, value.name)] } },
        sortDefinition: { sort: [{ field: columnField(category.table, category.name).field, direction: 'Ascending' }], isDefaultSort: false }
      },
      objects: {
        dataPoint: [{ properties: { defaultColor: color('#0078D4') } }],
        categoryAxis: [{ properties: { labelColor: color('#334155'), fontSize: literal(10) } }],
        valueAxis: [{ properties: { start: literal(0), labelColor: color('#52606D'), gridlineColor: color('#E7EEF5') } }],
        lineStyles: [{ properties: { strokeWidth: literal(3), showMarker: literal(true), markerShape: literal('circle'), markerSize: literal(7) } }]
      },
      visualContainerObjects: chrome(title, { altText: `${title}. Time series for the current filter context.` }),
      drillFilterOtherVisuals: true
    }
  };
}

function categoricalFilter(seed, table, column, values) {
  const alias = 'f';
  return {
    name: `Filter${stableId(seed).padEnd(24, '0').slice(0, 24)}`,
    field: columnField(table, column).field,
    type: 'Categorical',
    filter: {
      Version: 2,
      From: [{ Name: alias, Entity: table, Type: 0 }],
      Where: [{
        Condition: {
          In: {
            Expressions: [{ Column: { Expression: { SourceRef: { Source: alias } }, Property: column } }],
            Values: values.map(value => [{ Literal: { Value: `'${String(value).replaceAll("'", "''")}'` } }])
          }
        }
      }]
    },
    howCreated: 'User'
  };
}

function tableVisual(seed, fields, title, x, y, width, height, z, options = {}) {
  const name = stableId(seed);
  const query = { queryState: { Values: { projections: fields.map(fieldProjection) } } };
  if (options.sortBy) {
    query.sortDefinition = { sort: [{ field: fieldProjection(options.sortBy).field, direction: options.sortDirection || 'Ascending' }], isDefaultSort: false };
  }
  if (options.sorts?.length) {
    query.sortDefinition = { sort: options.sorts.map(({ field, direction }) => ({ field: fieldProjection(field).field, direction })), isDefaultSort: false };
  }
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    ...(options.filters?.length ? { filterConfig: { filters: options.filters } } : {}),
    visual: {
      visualType: 'tableEx',
      query,
      objects: {
        columnHeaders: [{ properties: { columnAdjustment: literal('growToFit'), autoSizeColumnWidth: literal(true), wordWrap: literal(true), fontColor: color('#FFFFFF'), backColor: color('#005A9E'), bold: literal(true), fontSize: literal(options.headerFontSize || 10) } }],
        values: [{ properties: { wordWrap: literal(true), backColorPrimary: color('#FFFFFF'), backColorSecondary: color('#F3F7FA'), fontColorPrimary: color('#1F2937'), fontColorSecondary: color('#1F2937'), fontSize: literal(options.valueFontSize || 9) } }],
        total: [{ properties: { totals: literal(false) } }]
      },
      visualContainerObjects: { ...chrome(title, { altText: `${title}. Detailed records for the current filter context.` }), stylePreset: [{ properties: { name: literal('None') } }] },
      drillFilterOtherVisuals: true
    }
  };
}

function matrix(seed, rows, columns, values, title, x, y, width, height, z) {
  const name = stableId(seed);
  const queryState = { Rows: { projections: rows.map(fieldProjection) } };
  if (columns.length) queryState.Columns = { projections: columns.map(fieldProjection) };
  queryState.Values = { projections: values.map(fieldProjection) };
  return {
    $schema: schemas.visual,
    name,
    position: position(x, y, width, height, z),
    visual: {
      visualType: 'pivotTable',
      query: { queryState },
      objects: {
        columnHeaders: [{ properties: { columnAdjustment: literal('growToFit'), autoSizeColumnWidth: literal(true), fontColor: color('#FFFFFF'), backColor: color('#005A9E'), bold: literal(true), fontSize: literal(10) } }],
        rowHeaders: [{ properties: { fontColor: color('#1F2937'), backColor: color('#F3F7FA') } }],
        values: [{ properties: { backColorPrimary: color('#FFFFFF'), backColorSecondary: color('#F3F7FA'), fontColorPrimary: color('#1F2937'), fontColorSecondary: color('#1F2937') } }]
      },
      visualContainerObjects: { ...chrome(title, { altText: `${title}. Summary grouped for the current filter context.` }), stylePreset: [{ properties: { name: literal('None') } }] },
      drillFilterOtherVisuals: true
    }
  };
}

const ribbon = M('Source Coverage', 'Evidence Ribbon Text');
const operationalSignalFields = [
  C('Device Operational Signals', 'Severity Label', 'Severity'),
  C('Device Operational Signals', 'Signal'),
  M('Device Operational Signals', 'Affected Devices', 'Affected'),
  M('Device Operational Signals', 'Covered Devices', 'Covered'),
  M('Device Operational Signals', 'Affected Among Covered (%)', '% Affected'),
  M('Device Operational Signals', 'Source Coverage vs Managed Estate (%)', 'Source Coverage'),
  C('Device Operational Signals', 'Recommended Action')
];
const workforceSignalFields = [
  C('Workforce Operational Signals', 'Severity Label', 'Severity'),
  C('Workforce Operational Signals', 'Signal'),
  C('Workforce Operational Signals', 'Risk Category'),
  M('Workforce Operational Signals', 'Workforce Signal Affected Users', 'Affected users'),
  M('Workforce Operational Signals', 'Workforce Signal Covered Population', 'Covered population'),
  M('Workforce Operational Signals', 'Workforce Signal Affected Estate', 'Affected estate (%)'),
  C('Workforce Operational Signals', 'Evidence Status'),
  C('Workforce Operational Signals', 'Recommended Action')
];
const managedOperationalSignalFields = [
  C('Device Operational Signals', 'Severity Label', 'Severity'),
  C('Device Operational Signals', 'Signal'),
  M('Device Operational Signals', 'Affected'),
  M('Device Operational Signals', 'Covered'),
  M('Device Operational Signals', 'Affected (%)'),
  C('Device Operational Signals', 'Recommended Action')
];
const managedOperationalSignalCompactFields = [
  C('Device Operational Signals', 'Severity Label', 'Severity'),
  C('Device Operational Signals', 'Signal'),
  M('Device Operational Signals', 'Affected'),
  M('Device Operational Signals', 'Affected (%)'),
  C('Device Operational Signals', 'Recommended Action')
];
const applicationOperationalSignalFields = [
  C('Application Operational Signals', 'Severity Label', 'Severity'),
  C('Application Operational Signals', 'Signal'),
  C('Application Operational Signals', 'Risk Category'),
  M('Application Operational Signals', 'Application Signal Affected', 'Affected'),
  M('Application Operational Signals', 'Application Signal Covered', 'Covered'),
  M('Application Operational Signals', 'Application Signal Affected (%)', 'Affected (%)'),
  C('Application Operational Signals', 'Recommended Action')
];
const collaborationOperationalSignalFields = [
  C('Collaboration Operational Signals', 'Severity Label', 'Severity'),
  C('Collaboration Operational Signals', 'Signal'),
  C('Collaboration Operational Signals', 'Risk Category'),
  M('Collaboration Operational Signals', 'Collaboration Signal Affected', 'Affected'),
  M('Collaboration Operational Signals', 'Collaboration Signal Covered', 'Covered'),
  M('Collaboration Operational Signals', 'Collaboration Signal Affected (%)', 'Affected (%)'),
  C('Collaboration Operational Signals', 'Recommended Action')
];
const contentStorageOperationalSignalFields = [
  C('Content Storage Operational Signals', 'Severity Label', 'Severity'),
  C('Content Storage Operational Signals', 'Signal'),
  C('Content Storage Operational Signals', 'Risk Category'),
  M('Content Storage Operational Signals', 'Affected Content'),
  M('Content Storage Operational Signals', 'Covered Content'),
  M('Content Storage Operational Signals', 'Affected Content (%)'),
  C('Content Storage Operational Signals', 'Recommended Action')
];
const backupOperationalSignalFields = [
  C('Backup Operational Signals', 'Severity Label', 'Severity'),
  C('Backup Operational Signals', 'Signal'),
  C('Backup Operational Signals', 'Risk Category'),
  M('Backup Operational Signals', 'Backup Signal Affected', 'Affected'),
  M('Backup Operational Signals', 'Backup Signal Covered', 'Covered'),
  M('Backup Operational Signals', 'Backup Signal Affected (%)', 'Affected (%)'),
  C('Backup Operational Signals', 'Recommended Action')
];
const dataTrustOperationalSignalFields = [
  C('Data Trust Operational Signals', 'Severity Label', 'Severity'),
  C('Data Trust Operational Signals', 'Signal'),
  C('Data Trust Operational Signals', 'Risk Category'),
  M('Data Trust Operational Signals', 'Data Trust Signal Affected', 'Affected'),
  M('Data Trust Operational Signals', 'Data Trust Signal Covered', 'Covered'),
  M('Data Trust Operational Signals', 'Data Trust Signal Affected (%)', 'Affected (%)'),
  C('Data Trust Operational Signals', 'Recommended Action')
];
const enterpriseOperationalSignalFields = [
  C('Enterprise Operational Signals', 'Domain'),
  C('Enterprise Operational Signals', 'Severity Label', 'Severity'),
  C('Enterprise Operational Signals', 'Signal'),
  C('Enterprise Operational Signals', 'Risk Category'),
  M('Enterprise Operational Signals', 'Affected Signal Observations', 'Affected observations'),
  M('Enterprise Operational Signals', 'Enterprise Signal Covered', 'Covered'),
  M('Enterprise Operational Signals', 'Enterprise Signal Affected (%)', 'Affected (%)'),
  C('Enterprise Operational Signals', 'Evidence Status'),
  C('Enterprise Operational Signals', 'Recommended Action')
];
const pages = [
  { name: 'Executive Overview', question: 'Digital Workplace health, adoption and protection at a glance', kpis: [M('User Snapshots', '# Active Workforce'), M('Device Snapshots', '# Devices'), M('License Capacity', 'License Capacity Utilization (%)'), M('License Capacity', 'M365 License Capacity Utilization (%)'), M('User Activities', 'M365 Active Use (30D) (%)'), M('Device Compliance Snapshots', 'Device Compliance (%)'), M('Endpoint Experience Measurements', 'Endpoint Analytics Score'), M('Security Control Evidence', 'Microsoft Secure Score (%)')], hero: ['bar', C('Service', 'Service Domain'), M('Findings', '# Actionable Findings'), 'Domain pulse'], detail: [C('Findings', 'Severity Label'), C('Findings', 'Finding Category'), M('Findings', '# Affected Entities'), C('Findings', 'Recommended Action')], support: [C('Source Coverage', 'Freshness State'), M('Source Coverage', '# Sources')] },
  { name: 'Workforce & Identity', question: 'Which workforce and identity records need attention?', kpis: [M('User Snapshots', '# Users'), M('User Snapshots', '# Active Workforce'), M('User Snapshots', '# Enabled Inactive Accounts'), M('Identity Matches', 'Identity Match (%)')], hero: ['bar', C('Geography', 'Country'), M('User Snapshots', '# Users'), 'Workforce state by country'], detail: [C('User', 'User Surrogate Key'), C('User', 'Account State'), C('User', 'Workforce Status'), C('Identity Matches', 'Match Status'), C('Identity Matches', 'Match Method')], support: [C('Identity Matches', 'Match Status'), M('Identity Matches', '# Identities')] },
  {
    name: 'Licensing & Cost',
    question: 'Where is licensed capacity underused or reclaimable?',
    subtitle: 'Licensing & Cost · BETA 1.0.0-beta.1 · current private tenant capacity and user-assignment evidence · non-finite, free, and trial entitlements excluded from capacity KPIs',
    kpis: [M('License Evidence', 'Observed Potentially Reclaimable Assignments', 'Potentially Reclaimable Assignments'), M('License Evidence', 'Available License Units'), M('License Evidence', 'Observed License Capacity Utilization (%)', 'Capacity Utilization (%)'), M('License Evidence', 'Enabled License Units'), M('License Evidence', 'Consumed License Units'), M('License Evidence', 'Governed License Products'), M('License Evidence', 'Observed Dormant 31-90D Assignments', 'Dormant 31-90D Assignments')],
    hero: ['bar', C('License Evidence', 'License Product'), M('License Evidence', 'Consumed License Units'), 'Consumed units by license product'],
    detail: [C('License Evidence', 'License Product'), C('License Evidence', 'SKU Part Number', 'SKU'), C('License Evidence', 'Capacity Class'), C('License Evidence', 'Enabled Units'), C('License Evidence', 'Consumed Units'), C('License Evidence', 'Available Units'), C('License Evidence', 'Capacity Utilization'), C('License Evidence', 'Assigned Users'), C('License Evidence', 'Potentially Reclaimable Assignments', 'Reclaim Candidates'), C('License Evidence', 'Risk Priority'), C('License Evidence', 'Recommended Action')],
    support: [C('License Evidence', 'Risk Priority'), M('License Evidence', 'License Products')],
    topSlicers: [[C('License Evidence', 'License Product'), 'License product'], [C('License Evidence', 'Capacity Class'), 'Capacity class'], [C('License Evidence', 'Risk Priority'), 'Risk priority']],
    bottomSource: 'Enterprise Operational Signals',
    bottom: enterpriseOperationalSignalFields,
    operationalSignals: ['Potentially reclaimable license assignments', 'Underused governed license products'],
    signalTitle: 'License optimization signals',
    evidenceRibbon: M('License Evidence', 'License Evidence Ribbon Text')
  },
  {
    name: 'Devices & Compliance',
    question: 'Which devices require operational attention?',
    subtitle: 'Devices & Compliance · BETA 1.0.0-beta.1 · current private inventory, compliance, update, and security-control evidence',
    kpis: [
      M('Device Inventory Evidence', 'Managed Devices', '# Managed Devices'),
      M('Device Inventory Evidence', 'Device Manufacturers', '# Manufacturers'),
      M('Device Inventory Evidence', 'Device Models', '# Models'),
      M('Device Inventory Evidence', 'Device Compliance Rate', 'Device Compliance Rate'),
      M('Device Inventory Evidence', 'Noncompliant Devices', '# Noncompliant Devices'),
      M('Device Inventory Evidence', 'AD to Intune Match Rate', 'AD to Intune Coverage (%)'),
      M('Device Inventory Evidence', 'Devices Low on Disk', '# Devices Low on Disk Space'),
      M('Device Inventory Evidence', 'Windows 10 Devices Remaining', '# Windows 10 Remaining'),
      M('Device Inventory Evidence', 'Windows 10 Devices Not Capable', '# Windows 10 Incompatible')
    ],
    hero: ['bar', C('Device Inventory Evidence', 'Platform'), M('Device Inventory Evidence', 'Managed Devices', '# Devices'), 'Managed device estate by platform'],
    detail: [C('Device Inventory Evidence', 'Device Name', 'Device'), C('Device Inventory Evidence', 'Management State', 'Management'), C('Device Inventory Evidence', 'Compliance State', 'Compliance'), C('Device Inventory Evidence', 'Ownership'), C('Device Inventory Evidence', 'Operating System Family', 'Operating System'), C('Device Inventory Evidence', 'Windows Release')],
    support: [C('Device Inventory Evidence', 'Compliance State'), M('Device Inventory Evidence', 'Managed Devices', '# Devices')],
    contextSlicer: [C('Device Inventory Evidence', 'Ownership'), 'Device ownership'],
    topSlicers: [[C('Device Inventory Evidence', 'Country'), 'Country'], [C('Device Inventory Evidence', 'Device Category'), 'Device category'], [C('Device Inventory Evidence', 'Manufacturer'), 'Manufacturer'], [C('Device Inventory Evidence', 'Compliance State'), 'Compliance state']],
    manufacturerShare: [C('Device Inventory Evidence', 'Manufacturer'), M('Device Inventory Evidence', 'Device Share by Manufacturer', 'Device share'), 'Device share by manufacturer (%)'],
    windowsVersionShare: [C('Device Inventory Evidence', 'Windows Release Short'), M('Device Inventory Evidence', 'Windows 11 Device Share by Release', 'Device share'), 'Windows 11 devices by release (%)'],
    deviceSecurityKpis: [[M('Security Control Evidence', 'Device Compliance Healthy (%)'), M('Security Control Evidence', 'Disk Encryption Healthy (%)')], [M('Security Control Evidence', 'Secure Boot Healthy (%)'), M('Security Control Evidence', 'Code Integrity Healthy (%)')]],
    bottomSource: 'Device Operational Signals',
    bottom: managedOperationalSignalFields,
    operationalSignals: ['Noncompliant devices', 'Devices low on disk space', 'Devices approaching low disk threshold', 'Devices not synced with Intune for more than 30 days', 'Secure Boot disabled', 'Device removal candidates', 'AD-to-Intune unmatched devices'],
    evidenceRibbon: M('Report Metadata', 'Device Evidence Ribbon Text')
  },
  {
    name: 'Windows Lifecycle',
    question: 'How exposed is the Windows estate to readiness and support risk?',
    kpis: [M('Device Inventory Evidence', 'Windows 10 Devices Remaining'), M('Device Inventory Evidence', 'Windows 10 Devices Not Capable'), M('Device Inventory Evidence', 'Older Windows 11 Release Devices'), M('Device Inventory Evidence', 'Current Windows 11 Adoption (%)', 'Windows 11 Adoption (%)'), M('Device Inventory Evidence', 'Windows 10 Readiness Assessed (%)', 'Windows 10 Readiness Assessed (%)'), M('Device Inventory Evidence', 'Secure Boot Disabled Devices')],
    hero: ['bar', C('Device Inventory Evidence', 'Windows Release Short'), M('Device Inventory Evidence', 'Managed Devices'), 'Windows estate by release'],
    detail: [C('Device Inventory Evidence', 'Device Name', 'Device'), C('Device Inventory Evidence', 'Windows Release'), C('Device Inventory Evidence', 'Windows 11 Upgrade Eligibility', 'Eligibility'), C('Device Inventory Evidence', 'Windows 11 Blocking Reasons', 'Blockers'), C('Device Inventory Evidence', 'Secure Boot State')],
    support: [C('Device Inventory Evidence', 'Windows 11 Release State'), M('Device Inventory Evidence', 'Managed Devices')],
    bottomSource: 'Device Operational Signals',
    bottom: managedOperationalSignalFields,
    operationalSignals: ['Windows 10 devices not capable of Windows 11', 'Windows Update hard failures', 'Windows 10 devices capable but not migrated', 'Windows 11 eligibility not assessed', 'Secure Boot disabled', 'Older Windows 11 releases'],
    topSlicers: [[C('Device Inventory Evidence', 'Country'), 'Country'], [C('Device Inventory Evidence', 'Device Category'), 'Device category'], [C('Device Inventory Evidence', 'Windows Release Short'), 'Windows release']],
    evidenceRibbon: M('Report Metadata', 'Device Evidence Ribbon Text')
  },
  {
    name: 'Endpoint Experience',
    question: 'Where is endpoint experience weakest or incomplete?',
    kpis: [M('Device Inventory Evidence', 'Current Endpoint Analytics Score', 'Endpoint Analytics Score'), M('Device Inventory Evidence', 'Endpoint Analytics Below 50 Devices'), M('Device Inventory Evidence', 'Devices with Stop Errors'), M('Device Inventory Evidence', 'Endpoint Analytics Managed Coverage (%)', 'Endpoint Analytics Coverage (%)'), M('Device Inventory Evidence', 'Current Startup Score', 'Startup Score'), M('Device Inventory Evidence', 'Current App Reliability Score', 'App Reliability Score'), M('Device Inventory Evidence', 'Managed Devices')],
    hero: ['bar', C('Device Inventory Evidence', 'Endpoint Analytics State'), M('Device Inventory Evidence', 'Managed Devices'), 'Managed devices by Endpoint Analytics state'],
    detail: [C('Device Inventory Evidence', 'Device Name', 'Device'), C('Device Inventory Evidence', 'Manufacturer'), C('Device Inventory Evidence', 'Model'), C('Device Inventory Evidence', 'Endpoint Analytics Score'), C('Device Inventory Evidence', 'Startup Score'), C('Device Inventory Evidence', 'App Reliability Score'), C('Device Inventory Evidence', 'Core Boot Time (s)'), C('Device Inventory Evidence', 'Core Sign-in Time (s)'), C('Device Inventory Evidence', 'Stop Error Count'), C('Device Inventory Evidence', 'Stop Error State')],
    support: [C('Device Inventory Evidence', 'Startup Score State'), M('Device Inventory Evidence', 'Managed Devices')],
    bottomSource: 'Device Operational Signals',
    bottom: managedOperationalSignalFields,
    operationalSignals: ['Endpoint Analytics score below 50', 'Startup score below 50', 'Boot time over 60 seconds', 'Sign-in time over 30 seconds', 'App reliability score below 50', 'Devices with stop errors'],
    topSlicers: [[C('Device Inventory Evidence', 'Country'), 'Country'], [C('Device Inventory Evidence', 'Device Category'), 'Device category'], [C('Device Inventory Evidence', 'Endpoint Analytics State'), 'Endpoint Analytics state']],
    evidenceRibbon: M('Report Metadata', 'Device Evidence Ribbon Text')
  },
  {
    name: 'Applications & Standardization',
    question: 'Which applications create the most footprint or version fragmentation?',
    subtitle: 'Applications & Standardization · BETA 1.0.0-beta.1 · current private Intune discovered-app evidence · 11 observed DATA-ALL weeks',
    kpis: [
      M('Application Inventory Evidence', 'Application Install Observations', 'Install Observations'),
      M('Application Inventory Evidence', 'Fragmented Applications'),
      M('Application Inventory Evidence', 'Installations Outside Dominant Version', 'Outside Dominant Version'),
      M('Application Inventory Evidence', 'Applications'),
      M('Application Inventory Evidence', 'Application Versions'),
      M('Application Inventory Evidence', 'Application Publishers', 'Publishers'),
      M('Application Inventory Evidence', 'Dominant Version Coverage (%)'),
      M('Application Inventory Evidence', 'Highly Fragmented Applications')
    ],
    hero: ['bar', C('Application Inventory Evidence', 'Application Name'), M('Application Inventory Evidence', 'Application Install Observations', 'Install Observations'), 'Application footprint'],
    detail: [C('Application Inventory Evidence', 'Application Name', 'Application'), C('Application Inventory Evidence', 'Publisher'), C('Application Inventory Evidence', 'Platform'), C('Application Inventory Evidence', 'Product Version Count', 'Versions'), C('Application Inventory Evidence', 'Standardization State'), M('Application Inventory Evidence', 'Application Install Observations', 'Install Observations'), M('Application Inventory Evidence', 'Installations Outside Dominant Version', 'Outside Dominant')],
    support: [C('Application Inventory Evidence', 'Standardization State'), M('Application Inventory Evidence', 'Applications')],
    topSlicers: [[C('Application Inventory Evidence', 'Platform'), 'Platform'], [C('Application Inventory Evidence', 'Publisher'), 'Publisher'], [C('Application Inventory Evidence', 'Standardization State'), 'Standardization']],
    bottomSource: 'Application Operational Signals',
    bottom: applicationOperationalSignalFields,
    signalTitle: 'Application standardization signals',
    evidenceRibbon: M('Application Inventory Evidence', 'Application Evidence Ribbon Text')
  },
  {
    name: 'Collaboration Adoption',
    question: 'Where are collaboration adoption and governance gaps?',
    subtitle: 'Collaboration Adoption · BETA 1.0.0-beta.1 · current private Teams, SharePoint, OneDrive and Copilot evidence · source freshness shown',
    kpis: [
      M('Collaboration Adoption Evidence', 'Service-weighted Activity Rate (%)', 'Service-weighted Activity Rate (30D)'),
      M('Collaboration Adoption Evidence', 'Active User-Service Observations 30D', 'Active User-Service Observations (30D)'),
      M('Collaboration Adoption Evidence', 'Stale Collaboration Sources', 'Stale Usage Sources'),
      M('Collaboration Adoption Evidence', 'Collaboration Eligible Users', 'Eligible Users'),
      M('Collaboration Adoption Evidence', 'Collaboration Service Objects', 'Service Objects'),
      M('Collaboration Adoption Evidence', 'Collaboration Inactive Objects', 'Inactive Objects'),
      M('Collaboration Adoption Evidence', 'Collaboration Orphaned Objects', 'Orphaned Objects')
    ],
    hero: ['bar', C('Collaboration Adoption Evidence', 'Service'), M('Collaboration Adoption Evidence', 'Service-weighted Activity Rate (%)', 'Active User Rate'), 'Thirty-day active-user rate by service'],
    detail: [C('Collaboration Adoption Evidence', 'Service'), M('Collaboration Adoption Evidence', 'Collaboration Eligible Users', 'Eligible Users'), M('Collaboration Adoption Evidence', 'Active User-Service Observations 30D', 'Active Users 30D'), M('Collaboration Adoption Evidence', 'Service-weighted Activity Rate (%)', 'Active Rate'), C('Collaboration Adoption Evidence', 'Usage Refresh Date'), C('Collaboration Adoption Evidence', 'Evidence Age Days'), C('Collaboration Adoption Evidence', 'Freshness State'), M('Collaboration Adoption Evidence', 'Collaboration Service Objects', 'Objects'), M('Collaboration Adoption Evidence', 'Collaboration Inactive Objects', 'Inactive'), M('Collaboration Adoption Evidence', 'Collaboration Orphaned Objects', 'Orphaned')],
    support: [C('Collaboration Adoption Evidence', 'Freshness State'), M('Collaboration Adoption Evidence', 'Collaboration Services', 'Services')],
    topSlicers: [[C('Collaboration Adoption Evidence', 'Service'), 'Service'], [C('Collaboration Adoption Evidence', 'Freshness State'), 'Freshness']],
    bottomSource: 'Collaboration Operational Signals',
    bottom: collaborationOperationalSignalFields,
    signalTitle: 'Collaboration adoption and governance signals',
    evidenceRibbon: M('Collaboration Adoption Evidence', 'Collaboration Evidence Ribbon Text')
  },
  {
    name: 'Content & Storage',
    question: 'Which SharePoint and OneDrive containers need lifecycle or storage attention?',
    subtitle: 'Content & Storage · BETA 1.0.0-beta.1 · current private SharePoint inventory · OneDrive usage freshness shown · SharePoint weekly history',
    kpis: [
      M('Content Storage Evidence', 'Content Storage Used (TB)', 'Storage Used (TB)'),
      M('Content Storage Evidence', 'Inactive Content Storage (TB)', 'Inactive Storage (TB)'),
      M('Content Storage Evidence', 'Containers at 80%+ Quota', 'At 80%+ Quota'),
      M('Content Storage Evidence', 'SharePoint Sites'),
      M('Content Storage Evidence', 'OneDrive Accounts'),
      M('Content Storage Evidence', 'Inactive Content Containers', 'Inactive 180D+'),
      M('Content Storage Evidence', 'Orphaned SharePoint Sites', 'Orphaned Sites'),
      M('Content Storage Evidence', 'OneDrive Accounts Without Activity Evidence', 'OneDrive Without Activity')
    ],
    hero: ['bar', C('Content Storage Evidence', 'Service'), M('Content Storage Evidence', 'Content Storage Used (TB)', 'Storage Used (TB)'), 'Storage used by service (TB)'],
    detail: [
      C('Content Storage Evidence', 'Content Object Name', 'Content Object'),
      C('Content Storage Evidence', 'Service'),
      C('Content Storage Evidence', 'Object Type'),
      C('Content Storage Evidence', 'Activity State'),
      C('Content Storage Evidence', 'Days Since Last Activity', 'Inactive Days'),
      C('Content Storage Evidence', 'Owner State'),
      C('Content Storage Evidence', 'Storage Used GB', 'Used (GB)'),
      C('Content Storage Evidence', 'Storage Allocated GB', 'Allocated (GB)'),
      C('Content Storage Evidence', 'Storage Utilization (%)', 'Utilization (%)'),
      C('Content Storage Evidence', 'Storage Pressure'),
      C('Content Storage Evidence', 'File Count', 'Files'),
      C('Content Storage Evidence', 'Risk Priority'),
      C('Content Storage Evidence', 'Risk Signal'),
      C('Content Storage Evidence', 'Recommended Action')
    ],
    support: [C('Content Storage Evidence', 'Freshness State'), M('Content Storage Evidence', 'Content Containers', 'Containers')],
    topSlicers: [[C('Content Storage Evidence', 'Service'), 'Service'], [C('Content Storage Evidence', 'Activity State'), 'Activity state'], [C('Content Storage Evidence', 'Storage Pressure'), 'Storage pressure'], [C('Content Storage Evidence', 'Risk Priority'), 'Risk priority']],
    bottomSource: 'Content Storage Operational Signals',
    bottom: contentStorageOperationalSignalFields,
    signalTitle: 'Content storage and lifecycle signals',
    evidenceRibbon: M('Content Storage Evidence', 'Content Evidence Ribbon Text'),
    drillthrough: C('Content Storage Evidence', 'Content Object Key')
  },
  {
    name: 'Messaging & Hybrid',
    question: 'Which mailbox records require hosting, migration, archive, or delegation attention?',
    subtitle: 'Messaging & Hybrid · BETA 1.0.0-beta.1 · current private Exchange Online and on-premises mailbox evidence with server-version reconciliation',
    kpis: [M('Mailbox Evidence', 'On-premises Mailboxes'), M('Mailbox Evidence', 'Large Mailboxes Without Archive'), M('Mailbox Evidence', 'Mailboxes With Extensive Delegation'), M('Mailbox Evidence', 'Mailboxes'), M('Mailbox Evidence', 'Exchange Online Mailboxes'), M('Mailbox Evidence', 'Observed Exchange Online Adoption (%)'), M('Mailbox Evidence', 'Mailboxes with Archive'), M('Mailbox Evidence', 'Mailboxes with Delegations'), M('Mailbox Evidence', 'Total Mailbox Size (TB)')],
    hero: ['bar', C('Mailbox Exchange Version Catalog', 'Exchange Version'), M('Mailbox Exchange Version Catalog', 'Mailboxes by Exchange Version', 'Mailboxes'), 'Mailboxes by Exchange version'],
    heroOptions: { categoryFontSize: 9, categoryLabelMaxMargin: 48 },
    detail: [C('Mailbox Evidence', 'Mailbox Key', 'Mailbox'), C('Mailbox Evidence', 'Primary SMTP Address', 'Primary SMTP'), C('Mailbox Evidence', 'Hosting Location', 'Hosting'), C('Mailbox Evidence', 'Exchange Version'), C('Mailbox Evidence', 'Mailbox Size GB', 'Size (GB)'), C('Mailbox Evidence', 'Archive State'), C('Mailbox Evidence', 'Archive Size GB', 'Archive (GB)'), C('Mailbox Evidence', 'Delegation Count', 'Delegations'), C('Mailbox Evidence', 'Delegation Types'), C('Mailbox Evidence', 'Mailbox Type'), C('Mailbox Evidence', 'Recipient Type'), C('Mailbox Evidence', 'Operational State'), C('Mailbox Evidence', 'Forwarding State'), C('Mailbox Evidence', 'Last Activity Date'), C('Mailbox Evidence', 'Risk Priority'), C('Mailbox Evidence', 'Risk Signal')],
    support: [C('Mailbox Evidence', 'Hosting Location'), M('Mailbox Evidence', 'Mailboxes')],
    topSlicers: [[C('Mailbox Evidence', 'Country'), 'Country'], [C('Mailbox Evidence', 'Hosting Location'), 'Hosting location'], [C('Mailbox Evidence', 'Exchange Version'), 'Exchange version'], [C('Mailbox Evidence', 'Mailbox Type'), 'Mailbox type']],
    bottomSource: 'Enterprise Operational Signals',
    bottom: enterpriseOperationalSignalFields,
    operationalSignals: ['Large mailboxes without archive', 'Mailboxes with extensive delegation', 'Mailboxes inactive for more than 180 days', 'Mailboxes with forwarding configured', 'On-premises mailboxes'],
    signalTitle: 'Messaging and hybrid attention signals',
    evidenceRibbon: M('Mailbox Evidence', 'Mailbox Evidence Ribbon Text'),
    drillthrough: C('Mailbox Evidence', 'Mailbox Key')
  },
  {
    name: 'Backup & Resilience',
    question: 'Which mailbox protection gaps, scope anomalies, or stale evidence require action?',
    subtitle: 'Backup & Resilience · BETA 1.0.0-beta.1 · private mailbox protection and policy-scope reconciliation · SharePoint and OneDrive backup evidence not collected',
    kpis: [
      M('Backup Mailbox Evidence', 'Expected Mailboxes Protected (%)'),
      M('Backup Mailbox Evidence', 'Expected Mailboxes Not Protected'),
      M('Backup Mailbox Evidence', 'Backup Policy Scope Anomalies'),
      M('Backup Mailbox Evidence', 'Expected Mailboxes'),
      M('Backup Mailbox Evidence', 'Protected Mailboxes'),
      M('Backup Mailbox Evidence', 'Disabled Scope Members'),
      M('Backup Mailbox Evidence', 'Stale Backup Sources'),
      M('Backup Mailbox Evidence', 'Backup Workload Evidence Gaps')
    ],
    hero: ['bar', C('Backup Mailbox Evidence', 'Protection State'), M('Backup Mailbox Evidence', 'Backup Evidence Records', 'Records'), 'Mailbox evidence by protection state'],
    detail: [
      C('Backup Mailbox Evidence', 'Display Name', 'Mailbox'),
      C('Backup Mailbox Evidence', 'Primary Address'),
      C('Backup Mailbox Evidence', 'Mailbox Type'),
      C('Backup Mailbox Evidence', 'Account State'),
      C('Backup Mailbox Evidence', 'Protection State'),
      C('Backup Mailbox Evidence', 'Reconciliation State'),
      C('Backup Mailbox Evidence', 'Protection Policy'),
      C('Backup Mailbox Evidence', 'Risk Priority'),
      C('Backup Mailbox Evidence', 'Evidence Age Days', 'Age (days)'),
      C('Backup Mailbox Evidence', 'Risk Signal'),
      C('Backup Mailbox Evidence', 'Recommended Action')
    ],
    support: [C('Backup Mailbox Evidence', 'Reconciliation State'), M('Backup Mailbox Evidence', 'Backup Evidence Records', 'Records')],
    topSlicers: [[C('Backup Mailbox Evidence', 'Protection State'), 'Protection state'], [C('Backup Mailbox Evidence', 'Reconciliation State'), 'Reconciliation'], [C('Backup Mailbox Evidence', 'Account State'), 'Account state'], [C('Backup Mailbox Evidence', 'Risk Priority'), 'Risk priority']],
    bottomSource: 'Backup Operational Signals',
    bottom: backupOperationalSignalFields,
    operationalSignals: ['Expected mailboxes not protected', 'Protected mailboxes outside expected policy scope', 'Policy scope members without a mailbox', 'Disabled accounts retained in expected policy scope', 'Protection units reporting errors', 'Stale backup evidence sources', 'SharePoint backup evidence not collected', 'OneDrive backup evidence not collected'],
    signalTitle: 'Backup protection and evidence signals',
    evidenceRibbon: M('Backup Mailbox Evidence', 'Backup Evidence Ribbon Text'),
    drillthrough: C('Backup Mailbox Evidence', 'Backup Entity Key')
  },
  {
    name: 'Security Posture',
    question: 'Which security controls have evidence-backed gaps?',
    subtitle: 'Security Posture · BETA 1.0.0-beta.1 · current private device, identity, access, endpoint protection, and Secure Score evidence',
    kpis: [M('Security Control Evidence', 'Microsoft Secure Score (%)'), M('Security Control Evidence', 'MFA Capability (%)'), M('Security Control Evidence', 'Defender Protection Active (%)'), M('Security Control Evidence', 'Conditional Access Enforcement (%)'), M('Security Control Evidence', 'Firewall Enabled (%)'), M('Security Control Evidence', 'Device Compliance Healthy (%)'), M('Security Control Evidence', 'Secure Boot Healthy (%)'), M('Security Control Evidence', 'Disk Encryption Healthy (%)')],
    hero: ['bar', C('Security Control Evidence', 'Control Name'), M('Security Control Evidence', 'Observed Security Health (%)'), 'Control health by security control'],
    detail: [C('Security Control Evidence', 'Control Name'), C('Security Control Evidence', 'Security Domain'), C('Security Control Evidence', 'Severity Label'), C('Security Control Evidence', 'Evidence Status'), C('Security Control Evidence', 'Covered Entities'), C('Security Control Evidence', 'Affected Entities'), C('Security Control Evidence', 'Metric Unit'), C('Security Control Evidence', 'Health Rate'), C('Security Control Evidence', 'Gap State'), C('Security Control Evidence', 'Evidence Source'), C('Security Control Evidence', 'Recommended Action')],
    support: [C('Security Control Evidence', 'Evidence Status'), M('Security Control Evidence', 'Security Controls')],
    topSlicers: [[C('Security Control Evidence', 'Security Domain'), 'Security domain'], [C('Security Control Evidence', 'Evidence Status'), 'Evidence status'], [C('Security Control Evidence', 'Severity Label'), 'Severity']],
    secondaryKpis: [[M('Security Control Evidence', 'Noncompliant Devices (#)'), M('Security Control Evidence', 'Devices Without Active Defender Protection (#)'), M('Security Control Evidence', 'Devices Without Enabled Firewall (#)')], [M('Security Control Evidence', 'Users Not MFA Capable (#)'), M('Security Control Evidence', 'Conditional Access Policies Not Enforced (#)')]],
    bottomSource: 'Enterprise Operational Signals',
    bottom: enterpriseOperationalSignalFields,
    operationalSignals: ['Secure Boot gaps', 'Disk encryption gaps', 'Code integrity gaps', 'Active Directory health warnings or critical checks', 'Entra Connect synchronization health issues', 'Users not MFA capable', 'Devices without active Defender protection', 'Devices without enabled firewall', 'Conditional Access policies not enforced'],
    signalTitle: 'Security control and evidence signals',
    evidenceRibbon: M('Security Control Evidence', 'Observed Security Evidence Ribbon Text'),
    drillthrough: C('Security Control Evidence', 'Control Name')
  },
  {
    name: 'Data Trust',
    question: 'Which evidence quality, freshness, or reconciliation gaps block trusted decisions?',
    subtitle: 'Data Trust · BETA 1.0.0-beta.1 · private source inventory, schema qualification, business-snapshot freshness, and cross-domain reconciliation evidence',
    kpis: [
      M('Data Trust Evidence', 'Decision Ready Sources (%)'),
      M('Data Trust Evidence', 'Sources Requiring Refresh'),
      M('Data Trust Evidence', 'Blocked Sources'),
      M('Data Trust Evidence', 'Evidence Sources'),
      M('Data Trust Evidence', 'Source Schema Compliance (%)', 'Schema Compliance (%)'),
      M('Data Trust Evidence', 'Entra Devices in Hardware ID Conflicts', 'Hardware ID Conflicts'),
      M('Device Inventory Evidence', 'AD-to-Intune Unmatched Devices'),
      M('Backup Mailbox Evidence', 'Backup Workload Evidence Gaps', 'Backup Evidence Gaps')
    ],
    hero: ['bar', C('Data Trust Evidence', 'Freshness State'), M('Data Trust Evidence', 'Evidence Sources', 'Sources'), 'Evidence sources by freshness'],
    detail: [
      C('Data Trust Evidence', 'Source Name'),
      C('Data Trust Evidence', 'Evidence Domain', 'Domain'),
      C('Data Trust Evidence', 'Snapshot Date'),
      C('Data Trust Evidence', 'Evidence Age Days', 'Age (days)'),
      C('Data Trust Evidence', 'Freshness State'),
      C('Data Trust Evidence', 'Schema State'),
      C('Data Trust Evidence', 'File State'),
      C('Data Trust Evidence', 'Row Count', 'Rows'),
      C('Data Trust Evidence', 'Decision Readiness'),
      C('Data Trust Evidence', 'Recommended Action')
    ],
    support: [C('Data Trust Evidence', 'Evidence Domain'), M('Data Trust Evidence', 'Evidence Sources', 'Sources')],
    topSlicers: [[C('Data Trust Evidence', 'Evidence Domain'), 'Domain'], [C('Data Trust Evidence', 'Freshness State'), 'Freshness'], [C('Data Trust Evidence', 'Decision Readiness'), 'Decision readiness'], [C('Data Trust Evidence', 'Schema State'), 'Schema state']],
    bottomSource: 'Data Trust Operational Signals',
    bottom: dataTrustOperationalSignalFields,
    operationalSignals: ['Stale evidence sources', 'Blocked or missing evidence sources', 'Source schema mismatches', 'Unexpectedly empty required sources', 'Entra hardware ID conflicts', 'AD-to-Intune unmatched devices', 'Device removal candidates', 'Backup workload evidence gaps', 'Mailbox backup scope anomalies', 'AD-to-Entra identity conflicts', 'Users missing activity evidence'],
    signalTitle: 'Data trust and reconciliation signals',
    evidenceRibbon: M('Data Trust Evidence', 'Data Trust Evidence Ribbon Text'),
    drillthrough: C('Data Trust Evidence', 'Source Key')
  },
  { name: 'Action Center', question: 'What should the Digital Workplace team address next?', kpis: [M('Findings', '# Critical Findings'), M('Findings', '# High Findings'), M('Findings', '# Medium Findings'), M('Findings', '# Evidence-Blocked Findings')], hero: ['bar', C('Service', 'Service Domain'), M('Findings', '# Actionable Findings'), 'Prioritized action queue by service domain'], detail: [C('Findings', 'Severity Label'), C('Findings', 'Finding Category'), C('Findings', 'Evidence Status'), M('Findings', '# Affected Entities'), C('Findings', 'Recommended Action')], support: [C('Findings', 'Evidence Status'), M('Findings', '# Actionable Findings')] },
  { name: 'User 360', question: "What evidence explains this user's workplace state?", kpis: [M('User Snapshots', '# Users'), M('User Activities', '# User Activity Signals'), M('User License Assignments', '# Potentially Reclaimable Assignments'), M('Identity Matches', 'Identity Match (%)')], hero: ['line', C('Date', 'Date'), M('User Activities', '# User Activity Signals'), 'Supported user activity timeline'], detail: [C('User', 'User Surrogate Key'), C('User', 'Account State'), C('User', 'Workforce Status'), C('Identity Matches', 'Match Status'), C('Identity Matches', 'Match Method')], support: [C('Service Relationships', 'Relationship Type'), M('Service Relationships', '# Relationships')], localSlicer: [C('User', 'User Surrogate Key'), 'User'] },
  { name: 'Device 360', question: "What evidence explains this device's workplace state?", kpis: [M('Device Snapshots', '# Devices'), M('Device Snapshots', '# Device Health Signals'), M('Device Compliance Snapshots', 'Device Compliance (%)'), M('Device Matches', 'Device Reconciliation (%)')], hero: ['line', C('Date', 'Date'), M('Device Snapshots', '# Device Health Signals'), 'Supported device health timeline'], detail: [C('Device', 'Device Surrogate Key'), C('Device', 'Management State'), C('Device', 'Compliance State'), C('Device', 'Operating System Family'), C('Device', 'Readiness State')], support: [C('Service Relationships', 'Relationship Type'), M('Service Relationships', '# Relationships')], localSlicer: [C('Device', 'Device Surrogate Key'), 'Device'] },
  { name: 'Service 360', question: "What evidence explains this service's adoption, health, and impact?", kpis: [M('Collaboration Usage', 'Service Active User (%)'), M('Collaboration Usage', '# Service Objects'), M('Data Quality Results', 'Data Quality Pass (%)'), M('Source Coverage', 'Source Freshness Compliance (%)')], hero: ['line', C('Date', 'Date'), M('Collaboration Usage', 'Service Active User (%)'), 'Service adoption and health trend'], detail: [C('Service', 'Service Name'), C('Service', 'Service Domain'), C('Service', 'Service Category'), M('Findings', '# Actionable Findings'), M('Service Relationships', '# Relationships')], support: [C('Service Relationships', 'Evidence Status'), M('Service Relationships', '# Relationships')], localSlicer: [C('Service', 'Service Name'), 'Service'] }
];

// Approved information architecture: domain overviews, six shared explorers,
// and one consolidated evidence-backed risk and incident signal page.
const workforce = pages.find(page => page.name === 'Workforce & Identity');
Object.assign(workforce, {
  subtitle: 'Workforce & Identity · BETA 1.0.0-beta.1 · private current Entra, M365 activity, AD identity and license evidence',
  kpis: [
    M('User Inventory Evidence', 'Active Workforce'),
    M('User Inventory Evidence', 'Accounts Without Activity >30D'),
    M('User Inventory Evidence', 'Activity Evidence Coverage'),
    M('User Inventory Evidence', 'Enabled Human Users'),
    M('Identity Reconciliation Evidence', 'AD to Entra Coverage'),
    M('Identity Reconciliation Evidence', 'Identity Conflicts'),
    M('User Inventory Evidence', 'Missing Activity Evidence'),
    M('User Inventory Evidence', 'Users With Missing Manager')
  ],
  hero: ['bar', C('User Inventory Evidence', 'Activity State'), M('User Inventory Evidence', 'Enabled Human Users'), 'Workforce activity state'],
  detail: [
    C('User Inventory Evidence', 'User Principal Name', 'User'),
    C('User Inventory Evidence', 'Display Name'),
    C('User Inventory Evidence', 'Department'),
    C('User Inventory Evidence', 'Manager'),
    C('User Inventory Evidence', 'Activity State'),
    C('User Inventory Evidence', 'Identity Match Status'),
    C('User Inventory Evidence', 'Has M365 License'),
    C('User Inventory Evidence', 'Attention Priority')
  ],
  topSlicers: [[C('User Inventory Evidence', 'Country'), 'Country'], [C('User Inventory Evidence', 'Department'), 'Department'], [C('User Inventory Evidence', 'Activity State'), 'Activity state']],
  map: [C('User Inventory Evidence', 'Country'), M('User Inventory Evidence', 'Enabled Human Users'), [M('User Inventory Evidence', 'Active Workforce'), M('User Inventory Evidence', 'Activity Evidence Coverage')], 'Workforce footprint by country'],
  identitySecurityKpis: [
    [M('Identity Reconciliation Evidence', 'AD to Entra Coverage'), M('Identity Reconciliation Evidence', 'Identity Conflicts')],
    [M('User Inventory Evidence', 'Users With Missing Manager'), M('User Inventory Evidence', 'Department Completeness')]
  ],
  identitySummaryTitle: 'Identity evidence posture',
  identitySummaryGroup1: 'AD to Entra reconciliation',
  identitySummaryGroup2: 'Directory data completeness',
  bottomSource: 'Workforce Operational Signals',
  bottom: workforceSignalFields,
  signalTitle: 'Workforce risk and attention signals',
  evidenceRibbon: M('Report Metadata', 'User Evidence Ribbon Text')
});

const actionCenter = pages.find(page => page.name === 'Action Center');
Object.assign(actionCenter, {
  name: 'Risk & Incident Signals',
  question: 'Which evidence-backed risk and incident signals require action now?',
  subtitle: 'Risk & Incident Signals · BETA 1.0.0-beta.1 · cross-domain priorities calculated from current private evidence only',
  kpis: [M('Enterprise Operational Signals', 'Affected Signal Observations'), M('Data Trust Evidence', 'Blocked Sources'), M('Security Control Evidence', 'Security Controls Without Evidence'), M('Device Inventory Evidence', 'Noncompliant Devices'), M('License Evidence', 'Observed Potentially Reclaimable Assignments', 'Potentially Reclaimable Assignments'), M('Mailbox Evidence', 'Large Mailboxes Without Archive')],
  hero: ['bar', C('Enterprise Operational Signals', 'Domain'), M('Enterprise Operational Signals', 'Affected Signal Observations'), 'Affected signal observations by domain'],
  detail: enterpriseOperationalSignalFields,
  support: [C('Enterprise Operational Signals', 'Severity Label'), M('Enterprise Operational Signals', 'Affected Signal Observations')],
  topSlicers: [[C('Enterprise Operational Signals', 'Domain'), 'Domain'], [C('Enterprise Operational Signals', 'Severity Label'), 'Severity'], [C('Enterprise Operational Signals', 'Evidence Status'), 'Evidence status']],
  bottomSource: 'Enterprise Operational Signals',
  bottom: enterpriseOperationalSignalFields,
  operationalSignals: ['Potentially reclaimable license assignments', 'Underused governed license products', 'Large mailboxes without archive', 'Mailboxes with extensive delegation', 'Mailboxes inactive for more than 180 days', 'Mailboxes with forwarding configured', 'On-premises mailboxes', 'Noncompliant devices', 'Secure Boot gaps', 'Disk encryption gaps', 'Code integrity gaps', 'Active Directory health warnings or critical checks', 'Entra Connect synchronization health issues', 'Users not MFA capable', 'Devices without active Defender protection', 'Devices without enabled firewall', 'Conditional Access policies not enforced', 'Stale evidence sources', 'Blocked or missing evidence sources', 'Accounts stale for more than 90 days', 'Highly fragmented applications', 'Content containers at 80 percent or more of quota', 'Expected mailboxes not protected'],
  signalTitle: 'Cross-domain risk and incident signals',
  evidenceRibbon: M('Data Trust Evidence', 'Data Trust Evidence Ribbon Text')
});

const userExplorer = pages.find(page => page.name === 'User 360');
Object.assign(userExplorer, {
  name: 'User Explorer',
  question: 'Which user evidence explains identity, activity, license, and risk state?',
  subtitle: 'User Explorer · BETA 1.0.0-beta.1 · private current workforce evidence · 11 observed DATA-ALL weeks',
  kpis: [
    M('User Inventory Evidence', 'Enabled Human Users'),
    M('User Inventory Evidence', 'Active Workforce'),
    M('User Inventory Evidence', 'Accounts Without Activity >30D'),
    M('User Inventory Evidence', 'Activity Evidence Coverage'),
    M('User Inventory Evidence', 'Missing Activity Evidence'),
    M('User Inventory Evidence', 'Users With Missing Manager'),
    M('User Inventory Evidence', 'Department Completeness')
  ],
  hero: ['bar', C('User Inventory Evidence', 'Attention Priority'), M('User Inventory Evidence', 'Enabled Human Users'), 'Users by attention priority'],
  detail: [
    C('User Inventory Evidence', 'User Principal Name', 'User'),
    C('User Inventory Evidence', 'Display Name'),
    C('User Inventory Evidence', 'Department'),
    C('User Inventory Evidence', 'Manager'),
    C('User Inventory Evidence', 'Activity State'),
    C('User Inventory Evidence', 'Days Since Last Activity'),
    C('User Inventory Evidence', 'Last Activity Date'),
    C('User Inventory Evidence', 'Last Activity Workload'),
    C('User Inventory Evidence', 'Identity Match Status'),
    C('User Inventory Evidence', 'M365 License Count'),
    C('User Inventory Evidence', 'Attention Priority')
  ],
  support: [C('User Inventory Evidence', 'Activity State'), M('User Inventory Evidence', 'Enabled Human Users')],
  topSlicers: [[C('User Inventory Evidence', 'User Principal Name'), 'User or mail'], [C('User Inventory Evidence', 'Department'), 'Department'], [C('User Inventory Evidence', 'Activity State'), 'Activity state']],
  drillthrough: C('User Inventory Evidence', 'User Principal Name'),
  bottomSource: 'Workforce Operational Signals',
  bottom: workforceSignalFields,
  signalTitle: 'Workforce risk and attention signals',
  evidenceRibbon: M('Report Metadata', 'User Evidence Ribbon Text')
});

const deviceExplorer = pages.find(page => page.name === 'Device 360');
Object.assign(deviceExplorer, {
  name: 'Device Explorer',
  question: "Which evidence explains each device's operational state?",
  subtitle: 'Device Explorer · BETA 1.0.0-beta.1 · current private inventory and update evidence',
  kpis: [M('Device Inventory Evidence', 'Managed Devices', '# Devices'), M('Device Inventory Evidence', 'Noncompliant Devices', 'Noncompliant Devices'), M('Device Inventory Evidence', 'Device Compliance Rate', 'Device Compliance (%)'), M('Device Inventory Evidence', 'AD to Intune Match Rate', 'AD to Intune Coverage (%)'), M('Device Inventory Evidence', 'Devices Low on Disk', 'Low Disk Devices'), M('Device Inventory Evidence', 'Average Observed Tenure (Years)'), M('Device Inventory Evidence', 'Devices Observed 4+ Years'), M('Device Inventory Evidence', 'Windows 10 Devices Remaining', '# Windows 10 Remaining'), M('Device Inventory Evidence', 'Windows 10 Devices Not Capable', '# Windows 10 Incompatible'), M('Device Inventory Evidence', 'Devices Not Synced 30+ Days', 'Not Synced >30 Days')],
  hero: ['bar', C('Device Inventory Evidence', 'Tenure Bucket'), M('Device Inventory Evidence', 'Managed Devices', '# Devices'), 'Managed estate by observed tenure'],
  support: [C('Device Inventory Evidence', 'Tenure Confidence'), M('Device Inventory Evidence', 'Managed Devices', '# Devices')],
  detail: [
    C('Device Inventory Evidence', 'Device Name', 'Device'),
    C('Device Inventory Evidence', 'Primary User UPN', 'Primary User'),
    C('Device Inventory Evidence', 'Manufacturer'),
    C('Device Inventory Evidence', 'Model'),
    C('Device Inventory Evidence', 'Ownership'),
    C('Device Inventory Evidence', 'Operating System Family', 'Operating System'),
    C('Device Inventory Evidence', 'Windows Release'),
    C('Device Inventory Evidence', 'Windows Generation'),
    C('Device Inventory Evidence', 'Windows 11 Upgrade Eligibility', 'Windows 11 Eligibility'),
    C('Device Inventory Evidence', 'Windows 11 Blocking Reasons', 'Windows 11 Blockers'),
    C('Device Inventory Evidence', 'Operating System Version', 'OS Version'),
    C('Device Inventory Evidence', 'Management State', 'Management'),
    C('Device Inventory Evidence', 'Compliance State', 'Compliance'),
    C('Device Inventory Evidence', 'AD to Intune Match', 'AD Match'),
    C('Device Inventory Evidence', 'Entra Match'),
    C('Device Inventory Evidence', 'Windows Update State', 'Update State'),
    C('Device Inventory Evidence', 'Windows Update Risk', 'Update Risk'),
    C('Device Inventory Evidence', 'Windows Update Action Code', 'Update Action'),
    C('Device Inventory Evidence', 'Secure Boot State'),
    C('Device Inventory Evidence', 'Endpoint Analytics Score', 'EA Score'),
    C('Device Inventory Evidence', 'Endpoint Analytics State', 'EA State'),
    C('Device Inventory Evidence', 'Stop Error Count'),
    C('Device Inventory Evidence', 'Stop Error State'),
    C('Device Inventory Evidence', 'Entra Hardware ID Conflict', 'Hardware ID Conflict'),
    C('Device Inventory Evidence', 'Removal Candidate'),
    C('Device Inventory Evidence', 'Windows 11 Release State'),
    C('Device Inventory Evidence', 'System Drive Capacity (GB)', 'Capacity (GB)'),
    C('Device Inventory Evidence', 'Free Disk Space (GB)', 'Free (GB)'),
    C('Device Inventory Evidence', 'Free Disk Space (%)', 'Free (%)'),
    C('Device Inventory Evidence', 'Disk Space State', 'Disk State'),
    C('Device Inventory Evidence', 'Last Sync DateTime', 'Last Sync'),
    C('Device Inventory Evidence', 'Observed Since'),
    C('Device Inventory Evidence', 'Managed Tenure (Years)', 'Observed Tenure (Years)'),
    C('Device Inventory Evidence', 'Tenure Confidence')
  ],
  topSlicers: [[C('Device Inventory Evidence', 'Device Name'), 'Device'], [C('Device Inventory Evidence', 'Device Category'), 'Device category'], [C('Device Inventory Evidence', 'Disk Space State'), 'Disk space state'], [C('Device Inventory Evidence', 'Secure Boot State'), 'Secure Boot']],
  windowsVersionShare: [C('Device Inventory Evidence', 'Windows Release Short'), M('Device Inventory Evidence', 'Windows 11 Device Share by Release', 'Device share'), 'Windows 11 devices by release (%)'],
  bottomSource: 'Device Operational Signals',
  bottom: managedOperationalSignalCompactFields,
  operationalSignals: ['Windows 10 devices not capable of Windows 11', 'Noncompliant devices', 'Devices low on disk space', 'Devices approaching low disk threshold', 'Devices not synced with Intune for more than 30 days', 'Windows Update hard failures', 'Windows 10 devices capable but not migrated', 'Windows 11 eligibility not assessed', 'Secure Boot disabled', 'Endpoint Analytics score below 50', 'Devices with stop errors', 'Entra hardware ID conflicts', 'Device removal candidates', 'Older Windows 11 releases', 'AD-to-Intune unmatched devices'],
  drillthrough: C('Device Inventory Evidence', 'Device Name'),
  evidenceRibbon: M('Report Metadata', 'Device Evidence Ribbon Text')
});

const serviceExplorer = pages.find(page => page.name === 'Service 360');
Object.assign(serviceExplorer, {
  name: 'Service & Content Explorer',
  question: 'Which collaboration objects require ownership, lifecycle, or access attention?',
  subtitle: 'Service & Content Explorer · BETA 1.0.0-beta.1 · current private Teams, SharePoint and OneDrive object evidence',
  kpis: [
    M('Collaboration Object Evidence', 'At-Risk Collaboration Objects', 'At-Risk Objects'),
    M('Collaboration Object Evidence', 'Inactive Collaboration Objects', 'Inactive Objects'),
    M('Collaboration Object Evidence', 'Orphaned Collaboration Objects', 'Orphaned Objects'),
    M('Collaboration Object Evidence', 'Collaboration Objects', 'Objects'),
    M('Collaboration Object Evidence', 'Active Collaboration Objects 30D', 'Active Objects 30D'),
    M('Collaboration Object Evidence', 'Collaboration Objects with External Access', 'External Access'),
    M('Collaboration Object Evidence', 'Collaboration Storage Used (GB)', 'Storage Used (GB)')
  ],
  hero: ['bar', C('Collaboration Object Evidence', 'Service'), M('Collaboration Object Evidence', 'At-Risk Collaboration Objects', 'At-Risk Objects'), 'At-risk objects by service'],
  detail: [C('Collaboration Object Evidence', 'Object Name', 'Object'), C('Collaboration Object Evidence', 'Service'), C('Collaboration Object Evidence', 'Object Type'), C('Collaboration Object Evidence', 'Activity State'), C('Collaboration Object Evidence', 'Days Since Last Activity', 'Inactive Days'), C('Collaboration Object Evidence', 'Owner State'), C('Collaboration Object Evidence', 'Guest Count', 'Guests'), C('Collaboration Object Evidence', 'External Access State'), C('Collaboration Object Evidence', 'Storage Used GB', 'Storage (GB)'), C('Collaboration Object Evidence', 'Risk Priority'), C('Collaboration Object Evidence', 'Risk Signal'), C('Collaboration Object Evidence', 'Recommended Action')],
  support: [C('Collaboration Object Evidence', 'Activity State'), M('Collaboration Object Evidence', 'Collaboration Objects', 'Objects')],
  topSlicers: [[C('Collaboration Object Evidence', 'Service'), 'Service'], [C('Collaboration Object Evidence', 'Object Type'), 'Object type'], [C('Collaboration Object Evidence', 'Risk Priority'), 'Risk priority'], [C('Collaboration Object Evidence', 'Activity State'), 'Activity state']],
  bottomSource: 'Collaboration Operational Signals',
  bottom: collaborationOperationalSignalFields,
  signalTitle: 'Collaboration object attention signals',
  evidenceRibbon: M('Collaboration Adoption Evidence', 'Collaboration Evidence Ribbon Text'),
  drillthrough: C('Collaboration Object Evidence', 'Object ID')
});

pages.push(
  {
    name: 'Application Explorer',
    question: 'Which application versions, publishers, and footprints require attention?',
    subtitle: 'Application Explorer · BETA 1.0.0-beta.1 · current private Intune discovered-app evidence · application-version grain',
    kpis: [
      M('Application Inventory Evidence', 'Applications'),
      M('Application Inventory Evidence', 'Application Versions'),
      M('Application Inventory Evidence', 'Application Publishers', 'Publishers'),
      M('Application Inventory Evidence', 'Application Install Observations', 'Install Observations'),
      M('Application Inventory Evidence', 'Highly Fragmented Applications'),
      M('Application Inventory Evidence', 'Long Tail Versions'),
      M('Application Inventory Evidence', 'Unknown Publisher Versions')
    ],
    hero: ['bar', C('Application Inventory Evidence', 'Publisher'), M('Application Inventory Evidence', 'Application Install Observations', 'Install Observations'), 'Application footprint by publisher'],
    detail: [C('Application Inventory Evidence', 'Application Name', 'Application'), C('Application Inventory Evidence', 'Application Version', 'Version'), C('Application Inventory Evidence', 'Publisher'), C('Application Inventory Evidence', 'Platform'), C('Application Inventory Evidence', 'Device Count', 'Devices'), C('Application Inventory Evidence', 'Product Version Count', 'Product Versions'), C('Application Inventory Evidence', 'Dominant Version Share (%)', 'Dominant Share'), C('Application Inventory Evidence', 'Version Footprint State', 'Footprint State'), C('Application Inventory Evidence', 'Risk Priority')],
    support: [C('Application Inventory Evidence', 'Risk Priority'), M('Application Inventory Evidence', 'Applications')],
    topSlicers: [[C('Application Inventory Evidence', 'Application Name'), 'Application'], [C('Application Inventory Evidence', 'Publisher'), 'Publisher'], [C('Application Inventory Evidence', 'Platform'), 'Platform'], [C('Application Inventory Evidence', 'Risk Priority'), 'Risk priority']],
    bottomSource: 'Application Operational Signals',
    bottom: applicationOperationalSignalFields,
    signalTitle: 'Application standardization signals',
    evidenceRibbon: M('Application Inventory Evidence', 'Application Evidence Ribbon Text'),
    drillthrough: C('Application Inventory Evidence', 'Application Name')
  },
  {
    name: 'Mailbox Explorer',
    question: 'Which mailbox evidence explains lifecycle, migration, identity, and protection state?',
    subtitle: 'Mailbox Explorer · BETA 1.0.0-beta.1 · current private Exchange Online and on-premises mailbox evidence',
    kpis: [M('Mailbox Evidence', 'Mailboxes'), M('Mailbox Evidence', 'Observed Exchange Online Adoption (%)'), M('Mailbox Evidence', 'Large Mailboxes Without Archive'), M('Mailbox Evidence', 'Exchange Online Mailboxes'), M('Mailbox Evidence', 'On-premises Mailboxes'), M('Mailbox Evidence', 'Mailboxes with Archive'), M('Mailbox Evidence', 'Mailboxes with Delegations'), M('Mailbox Evidence', 'Total Mailbox Size (TB)')],
    hero: ['bar', C('Mailbox Exchange Version Catalog', 'Exchange Version'), M('Mailbox Exchange Version Catalog', 'Mailboxes by Exchange Version', 'Mailboxes'), 'Mailboxes by Exchange version'],
    heroOptions: { categoryFontSize: 9, categoryLabelMaxMargin: 48 },
    detail: [C('Mailbox Evidence', 'Mailbox Key', 'Mailbox'), C('Mailbox Evidence', 'Primary SMTP Address', 'Primary SMTP'), C('Mailbox Evidence', 'Display Name'), C('Mailbox Evidence', 'Country'), C('Mailbox Evidence', 'Hosting Location', 'Hosting'), C('Mailbox Evidence', 'Exchange Version'), C('Mailbox Evidence', 'Mailbox Size GB', 'Size (GB)'), C('Mailbox Evidence', 'Archive State'), C('Mailbox Evidence', 'Archive Size GB', 'Archive (GB)'), C('Mailbox Evidence', 'Delegation Count', 'Delegations'), C('Mailbox Evidence', 'Delegation Types'), C('Mailbox Evidence', 'Mailbox Type'), C('Mailbox Evidence', 'Recipient Type'), C('Mailbox Evidence', 'Operational State'), C('Mailbox Evidence', 'Forwarding State'), C('Mailbox Evidence', 'Last Activity Date'), C('Mailbox Evidence', 'Risk Priority'), C('Mailbox Evidence', 'Risk Signal'), C('Mailbox Evidence', 'Recommended Action')],
    support: [C('Mailbox Evidence', 'Hosting Location'), M('Mailbox Evidence', 'Mailboxes')],
    topSlicers: [[C('Mailbox Evidence', 'Mailbox Key'), 'Mailbox'], [C('Mailbox Evidence', 'Hosting Location'), 'Hosting location'], [C('Mailbox Evidence', 'Exchange Version'), 'Exchange version'], [C('Mailbox Evidence', 'Mailbox Type'), 'Mailbox type']],
    bottomSource: 'Enterprise Operational Signals',
    bottom: enterpriseOperationalSignalFields,
    operationalSignals: ['Large mailboxes without archive', 'Mailboxes with extensive delegation', 'Mailboxes inactive for more than 180 days', 'Mailboxes with forwarding configured', 'On-premises mailboxes'],
    signalTitle: 'Mailbox lifecycle and access signals',
    evidenceRibbon: M('Mailbox Evidence', 'Mailbox Evidence Ribbon Text'),
    drillthrough: C('Mailbox Evidence', 'Mailbox Key')
  },
  {
    name: 'Control & Evidence Explorer',
    question: 'Which controls and sources support, weaken, or block a decision?',
    subtitle: 'Control & Evidence Explorer · BETA 1.0.0-beta.1 · current private security controls and cross-domain source qualification',
    kpis: [M('Security Control Evidence', 'Observed Security Evidence Coverage (%)'), M('Data Trust Evidence', 'Decision Ready Sources (%)'), M('Security Control Evidence', 'Security Controls Without Evidence'), M('Data Trust Evidence', 'Sources Requiring Refresh'), M('Data Trust Evidence', 'Blocked Sources'), M('Security Control Evidence', 'Security Gaps')],
    hero: ['bar', C('Security Control Evidence', 'Security Domain'), M('Security Control Evidence', 'Security Affected Entities'), 'Affected entities by security domain'],
    detail: [C('Security Control Evidence', 'Control Name'), C('Security Control Evidence', 'Security Domain'), C('Security Control Evidence', 'Severity Label'), C('Security Control Evidence', 'Evidence Status'), C('Security Control Evidence', 'Covered Entities'), C('Security Control Evidence', 'Affected Entities'), C('Security Control Evidence', 'Health Rate'), C('Security Control Evidence', 'Gap State'), C('Security Control Evidence', 'Evidence Source'), C('Security Control Evidence', 'Recommended Action')],
    support: [C('Security Control Evidence', 'Evidence Status'), M('Security Control Evidence', 'Security Controls')],
    topSlicers: [[C('Security Control Evidence', 'Security Domain'), 'Security domain'], [C('Security Control Evidence', 'Evidence Status'), 'Evidence status'], [C('Security Control Evidence', 'Severity Label'), 'Severity']],
    bottomSource: 'Enterprise Operational Signals',
    bottom: enterpriseOperationalSignalFields,
    operationalSignals: ['Stale evidence sources', 'Blocked or missing evidence sources'],
    signalTitle: 'Control and evidence qualification signals',
    evidenceRibbon: M('Data Trust Evidence', 'Data Trust Evidence Ribbon Text'),
    drillthrough: C('Security Control Evidence', 'Control Name')
  }
);

const geographicPages = {
  'Workforce & Identity': [C('User Inventory Evidence', 'Country'), M('User Inventory Evidence', 'Enabled Human Users'), [M('User Inventory Evidence', 'Active Workforce'), M('User Inventory Evidence', 'Activity Evidence Coverage')], 'Workforce footprint by country'],
  'Devices & Compliance': [C('Device Inventory Evidence', 'Country'), M('Device Inventory Evidence', 'Managed Devices', '# Devices'), [M('Device Inventory Evidence', 'Device Compliance Rate', 'Device Compliance (%)')], 'Managed devices by country'],
  'Windows Lifecycle': [C('Device Inventory Evidence', 'Country'), M('Device Inventory Evidence', 'Managed Devices'), [M('Device Inventory Evidence', 'Current Windows 11 Adoption (%)')], 'Windows estate by country'],
  'Endpoint Experience': [C('Device Inventory Evidence', 'Country'), M('Device Inventory Evidence', 'Managed Devices'), [M('Device Inventory Evidence', 'Current Endpoint Analytics Score', 'Endpoint Analytics Score')], 'Endpoint footprint by country'],
  'Messaging & Hybrid': [C('Mailbox Evidence', 'Country'), M('Mailbox Evidence', 'Mailboxes'), [M('Mailbox Evidence', 'Observed Exchange Online Adoption (%)')], 'Mailbox footprint by country']
};
pages.forEach(page => {
  if (geographicPages[page.name]) page.map = geographicPages[page.name];
});

const pageTrends = {
  'Workforce & Identity': { title: 'Active workforce trend', seriesLabel: 'Active users', category: C('Workforce Trend Evidence', 'Snapshot Date'), value: M('Workforce Trend Evidence', 'Trend Active Workforce'), change: M('Workforce Trend Evidence', 'Active Workforce Period Change'), tooltips: [M('Workforce Trend Evidence', 'Trend Enabled Human Users'), M('Workforce Trend Evidence', 'Trend Accounts Without Activity >30D'), M('Workforce Trend Evidence', 'Trend Activity Evidence Coverage')], color: '#1565C0' },
  'User Explorer': { title: 'Selected-scope activity trend', seriesLabel: 'Active users', category: C('User Activity History Evidence', 'Snapshot Date'), value: M('User Activity History Evidence', 'Historical Active Workforce'), change: M('User Activity History Evidence', 'Historical Active Workforce Period Change'), tooltips: [M('User Activity History Evidence', 'Historical Enabled Human Users'), M('User Activity History Evidence', 'Historical Accounts Without Activity >30D'), M('User Activity History Evidence', 'Historical Activity Evidence Coverage')], color: '#1565C0' },
  'Devices & Compliance': { title: 'Device estate trend', seriesLabel: '# Devices', value: M('Executive KPI Trends', '# Devices Trend'), delta: M('Executive KPI Trends', '# Devices Monthly Delta'), change: M('Executive KPI Trends', '# Devices 12M Change'), color: '#005A9E' },
  'Windows Lifecycle': { title: 'Windows adoption trend', seriesLabel: 'Windows 11 Adoption', category: C('Windows Lifecycle Trend Evidence', 'Snapshot Date'), value: M('Windows Lifecycle Trend Evidence', 'Trend Windows 11 Adoption'), change: M('Windows Lifecycle Trend Evidence', 'Windows 11 Adoption Period Change'), tooltips: [M('Windows Lifecycle Trend Evidence', 'Trend Windows 10 Devices'), M('Windows Lifecycle Trend Evidence', 'Trend Windows 11 Devices')], color: '#107C10' },
  'Endpoint Experience': { title: 'Endpoint Analytics score trend', seriesLabel: 'Endpoint Analytics Score', category: C('Endpoint Experience Trend Evidence', 'Snapshot Date'), value: M('Endpoint Experience Trend Evidence', 'Trend Endpoint Analytics Score'), change: M('Endpoint Experience Trend Evidence', 'Endpoint Analytics Score Period Change'), tooltips: [M('Endpoint Experience Trend Evidence', 'Trend Startup Score'), M('Endpoint Experience Trend Evidence', 'Trend App Reliability Score'), M('Endpoint Experience Trend Evidence', 'Trend Endpoint Analytics Covered Devices'), M('Endpoint Experience Trend Evidence', 'Trend Devices with Stop Errors')], color: '#005A9E' },
  'Applications & Standardization': { title: 'Application standardization trend', seriesLabel: 'Dominant Version Coverage', category: C('Application Trend Evidence', 'Snapshot Date'), value: M('Application Trend Evidence', 'Trend Dominant Version Coverage (%)'), change: M('Application Trend Evidence', 'Dominant Version Coverage Period Change'), tooltips: [M('Application Trend Evidence', 'Trend Application Install Observations'), M('Application Trend Evidence', 'Trend Fragmented Applications')], color: '#7A5AF8' },
  'Collaboration Adoption': { title: 'Weekly collaboration object activity trend', seriesLabel: 'Active Object Rate', category: C('Collaboration Trend Evidence', 'Snapshot Week'), value: M('Collaboration Trend Evidence', 'Trend Active Object Rate (%)'), change: M('Collaboration Trend Evidence', 'Active Object Rate Period Change'), tooltips: [M('Collaboration Trend Evidence', 'Trend Collaboration Objects'), M('Collaboration Trend Evidence', 'Trend Adoption Rate (%)')], color: '#00838F' },
  'Service & Content Explorer': { title: 'Weekly service object activity trend', seriesLabel: 'Active Object Rate', category: C('Collaboration Trend Evidence', 'Snapshot Week'), value: M('Collaboration Trend Evidence', 'Trend Active Object Rate (%)'), change: M('Collaboration Trend Evidence', 'Active Object Rate Period Change'), tooltips: [M('Collaboration Trend Evidence', 'Trend Collaboration Objects')], color: '#00838F' },
  'Content & Storage': { title: 'Weekly SharePoint storage trend', seriesLabel: 'SharePoint storage used (GB)', category: C('Content Storage Trend Evidence', 'Snapshot Date'), value: M('Content Storage Trend Evidence', 'Trend SharePoint Storage Used (GB)'), change: M('Content Storage Trend Evidence', 'SharePoint Storage Period Change'), tooltips: [M('Content Storage Trend Evidence', 'Trend SharePoint Sites'), M('Content Storage Trend Evidence', 'Trend Inactive SharePoint Sites')], color: '#00838F' },
  'Messaging & Hybrid': { title: 'Messaging adoption trend', seriesLabel: 'Exchange Online Adoption', value: M('Executive KPI Trends', 'Exchange Online Adoption Trend (%)'), delta: M('Executive KPI Trends', 'Exchange Online Adoption Monthly Delta (pp)'), change: M('Executive KPI Trends', 'Exchange Online Adoption 12M Change'), color: '#00838F' }
};
pages.forEach(page => {
  if (pageTrends[page.name]) page.trend = pageTrends[page.name];
});

const pageSignalFilters = {
  'Workforce & Identity': ['Affected Entity Type', ['User']],
  'User Explorer': ['Affected Entity Type', ['User']],
  'Licensing & Cost': ['Finding Category', ['Identity Hygiene', 'Ownership']],
  'Devices & Compliance': ['Affected Entity Type', ['Device']],
  'Device Explorer': ['Affected Entity Type', ['Device']],
  'Windows Lifecycle': ['Affected Entity Type', ['Device']],
  'Endpoint Experience': ['Affected Entity Type', ['Device']],
  'Applications & Standardization': ['Finding Category', ['Application Standardization']],
  'Application Explorer': ['Finding Category', ['Application Standardization']],
  'Collaboration Adoption': ['Finding Category', ['Collaboration Adoption']],
  'Service & Content Explorer': ['Finding Category', ['Collaboration Adoption', 'Content Governance']],
  'Content & Storage': ['Finding Category', ['Content Governance']],
  'Messaging & Hybrid': ['Affected Entity Type', ['Mailbox']],
  'Mailbox Explorer': ['Affected Entity Type', ['Mailbox']],
  'Backup & Resilience': ['Finding Category', ['Backup Coverage']],
  'Security Posture': ['Finding Category', ['Identity Hygiene', 'Identity Synchronization', 'Ownership', 'Device Compliance', 'Security Evidence']],
  'Control & Evidence Explorer': ['Finding Category', ['Security Evidence']],
  'Data Trust': ['Finding Category', ['Security Evidence']]
};
pages.forEach(page => {
  if (pageSignalFilters[page.name]) page.signalFilter = pageSignalFilters[page.name];
});

// Three decision KPIs receive the strongest visual emphasis on each page.
// The remaining measures stay visible as supporting context without competing
// with the page's immediate operational question.
const priorityKpiOrders = {
  'Workforce & Identity': [0, 1, 2],
  'User Explorer': [0, 1, 2],
  'Licensing & Cost': [0, 1, 2],
  'Devices & Compliance': [0, 3, 4],
  'Device Explorer': [1, 4, 9],
  'Windows Lifecycle': [0, 1, 2],
  'Endpoint Experience': [0, 1, 2],
  'Applications & Standardization': [0, 1, 2],
  'Application Explorer': [0, 1, 2],
  'Collaboration Adoption': [0, 1, 2],
  'Service & Content Explorer': [0, 1, 2],
  'Content & Storage': [0, 2, 1],
  'Messaging & Hybrid': [0, 1, 2],
  'Mailbox Explorer': [0, 1, 2],
  'Backup & Resilience': [0, 1, 2],
  'Security Posture': [0, 1, 2],
  'Control & Evidence Explorer': [0, 1, 2],
  'Data Trust': [0, 1, 2],
  'Risk & Incident Signals': [0, 1, 2]
};
pages.forEach(page => {
  if (priorityKpiOrders[page.name]) page.priorityKpiOrder = priorityKpiOrders[page.name];
});

function buildExecutivePage(page) {
  const pageId = stableId(`page:${page.name}`);
  const visuals = [];
  let z = 100;
  const pageIcon = pageIcons[page.name];

  visuals.push(image(`${pageId}:page-icon`, pageIcon, `${page.name} page icon`, 32, 16, 48, 48, 900));
  visuals.push(textbox(`${pageId}:title`, page.question, 96, 16, 1024, 48, z++, { size: '26px', color: '#172B4D', bold: true }));
  visuals.push(textbox(`${pageId}:subtitle`, 'Executive cockpit · BETA 1.0.0-beta.1 · private source-backed executive evidence · observed histories where available', 96, 64, 1024, 32, z++, { size: '13px', color: '#52606D' }));
  visuals.push(slicer(`${pageId}:workforce-country`, 'User Inventory Evidence', 'Country', 'Workforce country', 1184, 8, 168, z++));
  visuals.push(slicer(`${pageId}:device-country`, 'Device Inventory Evidence', 'Country', 'Device country', 1368, 8, 168, z++));
  visuals.push(slicer(`${pageId}:evidence-domain`, 'Data Trust Evidence', 'Evidence Domain', 'Evidence domain', 1552, 8, 168, z++));
  visuals.push(card(`${pageId}:evidence-header`, [M('Data Trust Evidence', 'Data Trust Evidence Ribbon Text')], 'Evidence & freshness', 1736, 8, 152, 80, z++, { valueSize: 9, labelSize: 8, titleSize: 9, valueColor: '#2F6B9A', accentColor: '#2F6B9A', showLabel: false, padding: 1 }));

  const groups = [
    { id: 'workforce', title: 'Workforce & Identity', x: 32, y: 112, height: 296, width: 560, color: '#1565C0', icon: 'smartworkplace-people-20260913.svg', fields: [M('User Inventory Evidence', 'Active Workforce'), M('User Inventory Evidence', 'Enabled Human Users', '# Users'), M('Identity Reconciliation Evidence', 'AD to Entra Coverage')], category: C('Workforce Trend Evidence', 'Snapshot Date'), trend: M('Workforce Trend Evidence', 'Trend Active Workforce'), change: M('Workforce Trend Evidence', 'Active Workforce Period Change'), tooltips: [M('Workforce Trend Evidence', 'Trend Enabled Human Users'), M('Workforce Trend Evidence', 'Trend Accounts Without Activity >30D'), M('Workforce Trend Evidence', 'Trend Activity Evidence Coverage')] },
    { id: 'devices', title: 'Devices & Endpoint', x: 616, y: 112, height: 296, width: 768, color: '#005A9E', icon: 'smartworkplace-kpi-devices-20260913.svg', fields: [M('Device Inventory Evidence', 'Managed Devices', '# Devices'), M('Device Inventory Evidence', 'AD to Intune Match Rate'), M('Device Inventory Evidence', 'Device Compliance Rate'), M('Device Inventory Evidence', 'Current Endpoint Analytics Score', 'Endpoint Analytics Score')], category: C('Endpoint Experience Trend Evidence', 'Snapshot Date'), trend: M('Endpoint Experience Trend Evidence', 'Trend Endpoint Analytics Score'), change: M('Endpoint Experience Trend Evidence', 'Endpoint Analytics Score Period Change'), tooltips: [M('Endpoint Experience Trend Evidence', 'Trend Startup Score'), M('Endpoint Experience Trend Evidence', 'Trend App Reliability Score'), M('Endpoint Experience Trend Evidence', 'Trend Endpoint Analytics Covered Devices')] },
    { id: 'licensing', title: 'Licensing & Cost', x: 1408, y: 112, height: 296, width: 480, color: '#6D5BD0', icon: 'smartworkplace-kpi-licenses-20260913.svg', fields: [M('License Evidence', 'Observed License Capacity Utilization (%)', 'Capacity Utilization (%)'), M('License Evidence', 'Observed Potentially Reclaimable Assignments', 'Potentially Reclaimable Assignments'), M('License Evidence', 'Available License Units')], status: M('License Evidence', 'License Evidence Ribbon Text') },
    { id: 'windows', title: 'Windows modernization', x: 32, y: 424, height: 304, width: 560, color: '#107C10', icon: 'smartworkplace-lifecycle-20260913.svg', fields: [M('Device Inventory Evidence', 'Current Windows 11 Adoption (%)'), M('Device Inventory Evidence', 'Windows 10 Devices Remaining'), M('Device Inventory Evidence', 'Windows 10 Devices Not Capable')], category: C('Windows Lifecycle Trend Evidence', 'Snapshot Date'), trend: M('Windows Lifecycle Trend Evidence', 'Trend Windows 11 Adoption'), change: M('Windows Lifecycle Trend Evidence', 'Windows 11 Adoption Period Change'), tooltips: [M('Windows Lifecycle Trend Evidence', 'Trend Windows 10 Devices'), M('Windows Lifecycle Trend Evidence', 'Trend Windows 11 Devices')] },
    { id: 'messaging', title: 'Messaging & protection', x: 616, y: 424, height: 304, width: 768, color: '#00838F', icon: 'smartworkplace-ratio-mailboxes-20260913.svg', fields: [M('Mailbox Evidence', 'Mailboxes'), M('Mailbox Evidence', 'Observed Exchange Online Adoption (%)'), M('Backup Mailbox Evidence', 'Expected Mailboxes Protected (%)'), M('Backup Mailbox Evidence', 'Backup Policy Scope Anomalies')], status: M('Mailbox Evidence', 'Mailbox Evidence Ribbon Text') },
    { id: 'security', title: 'Security posture', x: 1408, y: 424, height: 304, width: 480, color: '#2F6B9A', icon: 'smartworkplace-health-20260913.svg', fields: [M('Security Control Evidence', 'Microsoft Secure Score (%)'), M('Security Control Evidence', 'MFA Capability (%)'), M('Security Control Evidence', 'Defender Protection Active (%)')], status: M('Security Control Evidence', 'Observed Security Evidence Ribbon Text') }
  ];
  groups.forEach((group, i) => {
    visuals.push(shapeContainer(`${pageId}:group-frame:${group.id}`, group.x, group.y, group.width, group.height, z++));
    visuals.push(card(`${pageId}:group:${group.id}`, group.fields, group.title, group.x + 1, group.y + 1, group.width - 2, 140, z++, { valueSize: group.fields.length > 3 ? 17 : 19, labelSize: 10, valueColor: group.color, accentColor: group.color, bare: true }));
    if (group.trend) {
      const changeWidth = group.change ? 112 : 0;
      const trendWidth = group.width - changeWidth - (group.change ? 36 : 24);
      visuals.push(sparkline(`${pageId}:trend:${group.id}`, group.category, group.trend, group.tooltips || [], group.x + 12, group.y + 142, trendWidth, group.height - 154, z++, group.color, { showCategoryAxis: false, axisFontSize: 9 }));
      if (group.change) {
        visuals.push(card(`${pageId}:trend-change:${group.id}`, [group.change], 'Period change', group.x + group.width - changeWidth - 12, group.y + 186, changeWidth, 80, z++, { valueSize: 14, labelSize: 8, valueColor: group.color, showLabel: false, titleSize: 9, bare: true, altText: `Period change for ${group.title}.` }));
      }
    } else {
      visuals.push(card(`${pageId}:status:${group.id}`, [group.status], 'Evidence status', group.x + 12, group.y + 154, group.width - 24, group.height - 166, z++, { valueSize: 10, labelSize: 8, valueColor: group.color, showLabel: false, titleSize: 9, bare: true, altText: `Evidence status for ${group.title}.` }));
    }
    visuals.push(image(`${pageId}:group-icon:${group.id}`, group.icon, `${group.title} icon`, group.x + group.width - 40, group.y + 12, 24, 24, 920 + i));
  });

  visuals.push(azureMap(`${pageId}:country-map`, C('Device Inventory Evidence', 'Country'), M('Device Inventory Evidence', 'Managed Devices'), [
    M('Device Inventory Evidence', 'Device Compliance Rate'),
    M('Device Inventory Evidence', 'Current Windows 11 Adoption (%)'),
    M('Device Inventory Evidence', 'Current Endpoint Analytics Score')
  ], 'Managed estate by country', 32, 744, 560, 264, z++));
  visuals.push(image(`${pageId}:country-map-icon`, 'smartworkplace-overview-20260913.svg', 'Workplace footprint by country icon', 552, 756, 24, 24, 940));

  visuals.push(tableVisual(`${pageId}:priorities`, [
    C('Enterprise Operational Signals', 'Domain'),
    C('Enterprise Operational Signals', 'Severity Label'),
    C('Enterprise Operational Signals', 'Signal'),
    M('Enterprise Operational Signals', 'Affected Signal Observations'),
    C('Enterprise Operational Signals', 'Recommended Action')
  ], 'Leadership priorities', 616, 744, 1272, 264, z++, {
    headerFontSize: 12,
    valueFontSize: 12,
    sorts: [
      { field: C('Enterprise Operational Signals', 'Severity Label'), direction: 'Ascending' },
      { field: M('Enterprise Operational Signals', 'Affected Signal Observations'), direction: 'Descending' }
    ]
  }));
  visuals.push(image(`${pageId}:priorities-icon`, 'smartworkplace-section-findings-20260913.svg', 'Leadership priorities icon', 1848, 756, 24, 24, 942));
  visuals.push(textbox(`${pageId}:footer`, 'Product version · Data as of · Evidence status · © 2026 WorkplaceCloudHub — https://workplacecloudhub.com/', 32, 1020, 1856, 28, z++, { size: '11px', color: '#6B7280', align: 'center' }));

  const pageJson = {
    $schema: schemas.page,
    name: pageId,
    displayName: page.name,
    displayOption: 'FitToPage',
    height: 1080,
    width: 1920,
    objects: {
      background: [{ properties: { color: color('#F4F7FB'), transparency: literal(0) } }],
      outspace: [{ properties: { color: color('#E8EEF5'), transparency: literal(0) } }]
    }
  };

  const pageDir = path.join(pagesRoot, pageId);
  writeJson(path.join(pageDir, 'page.json'), pageJson);
  for (const visual of visuals) writeJson(path.join(pageDir, 'visuals', visual.name, 'visual.json'), visual);
  return { pageId, visualCount: visuals.length };
}

function buildPage(page, index) {
  if (page.name === 'Executive Overview') return buildExecutivePage(page);
  const pageId = stableId(`page:${page.name}`);
  const visuals = [];
  let z = 100;
  const pageIcon = pageIcons[page.name];
  visuals.push(image(`${pageId}:page-icon`, pageIcon, `${page.name} page icon`, 32, 16, 48, 48, 900));
  visuals.push(textbox(`${pageId}:title`, page.question, 96, 16, 1024, 48, z++, { size: '26px', color: '#172B4D', bold: true }));
  visuals.push(textbox(`${pageId}:subtitle`, page.subtitle || `${page.name} · BETA 1.0.0-beta.1 · current private source-backed evidence`, 96, 64, 1024, 32, z++, { size: '13px', color: '#52606D' }));
  const topSlicers = page.topSlicers || [
    [C('Date', 'Reporting Period'), 'Reporting period'],
    [C('Geography', 'Country'), 'Country'],
    [C('Tenant', 'Tenant Name'), 'Tenant']
  ];
  const fourTopSlicers = topSlicers.length > 3;
  const slicerXs = fourTopSlicers ? [1120, 1276, 1432, 1588] : [1184, 1368, 1552];
  const slicerWidths = fourTopSlicers ? [140, 140, 140, 140] : [168, 168, 168];
  topSlicers.slice(0, fourTopSlicers ? 4 : 3).forEach((entry, slicerIndex) => {
    visuals.push(slicer(`${pageId}:top-slicer:${slicerIndex}`, entry[0].table, entry[0].name, entry[1], slicerXs[slicerIndex], 8, slicerWidths[slicerIndex], z++));
  });
  visuals.push(card(`${pageId}:evidence-header`, [page.evidenceRibbon || ribbon], 'Evidence & freshness', fourTopSlicers ? 1744 : 1736, 8, fourTopSlicers ? 144 : 152, 80, z++, { valueSize: 9, labelSize: 8, titleSize: 9, valueColor: '#2F6B9A', accentColor: '#2F6B9A', showLabel: false, padding: 1 }));
  const priorityIndices = page.priorityKpiOrder || page.kpis.slice(0, 3).map((_, kpiIndex) => kpiIndex);
  const priorityKpis = priorityIndices.map(kpiIndex => page.kpis[kpiIndex]).filter(Boolean).slice(0, 3);
  const supportingKpis = page.kpis.filter((_, kpiIndex) => !priorityIndices.includes(kpiIndex));
  if (supportingKpis.length) {
    visuals.push(card(`${pageId}:priority-kpis`, priorityKpis, 'Priority KPIs', 32, 112, 1064, 120, z++, { valueSize: 26, labelSize: 12, displayUnits: 1, accentColor: '#005A9E', altText: `Three priority indicators for ${page.name}.` }));
    visuals.push(card(`${pageId}:supporting-kpis`, supportingKpis, 'Supporting KPIs', 1120, 112, 768, 120, z++, { valueSize: supportingKpis.length > 4 ? 17 : 20, labelSize: 10, displayUnits: 1, valueColor: '#2F6B9A', accentColor: '#7A9CB8', altText: `Supporting KPIs for ${page.name}.` }));
  } else {
    visuals.push(card(`${pageId}:priority-kpis`, priorityKpis, 'Priority KPIs', 32, 112, 1856, 120, z++, { valueSize: 26, labelSize: 12, accentColor: '#005A9E', altText: `Three priority indicators for ${page.name}.` }));
  }
  visuals.push(image(`${pageId}:kpi-icon`, 'smartworkplace-kpi-analytics-20260913.svg', 'Priority indicators icon', 1056, 124, 24, 24, 911));

  const hasExtraSlicers = Array.isArray(page.extraSlicers) && page.extraSlicers.length > 0;
  if (hasExtraSlicers) {
    page.extraSlicers.slice(0, 2).forEach((entry, slicerIndex) => {
      visuals.push(slicer(`${pageId}:extra-slicer:${slicerIndex}`, entry[0].table, entry[0].name, entry[1], 32 + (slicerIndex * 288), 248, 272, z++));
    });
  }
  const mainY = hasExtraSlicers ? 344 : 256;
  const mainHeight = hasExtraSlicers ? 312 : 400;
  const isRiskPage = page.name === 'Risk & Incident Signals';
  const isSecurityPage = page.name === 'Security Posture';
  const isDeviceExplorer = page.name === 'Device Explorer';
  const isWideDetailExplorer = isDeviceExplorer || page.name === 'User Explorer' || page.name === 'Mailbox Explorer' || page.name === 'Messaging & Hybrid';
  const contextWidth = isWideDetailExplorer ? 320 : 560;
  const [kind, category, value, title] = page.hero;
  if (page.map) {
    const mapWidth = isWideDetailExplorer ? 320 : 560;
    visuals.push(azureMap(`${pageId}:map`, page.map[0], page.map[1], page.map[2], page.map[3], 32, mainY, mapWidth, mainHeight, z++));
    visuals.push(image(`${pageId}:map-icon`, 'smartworkplace-overview-20260913.svg', `${page.map[3]} icon`, 32 + mapWidth - 40, mainY + 12, 24, 24, 912));
  } else if (!isSecurityPage) {
    const contextVisual = matrix(`${pageId}:context`, [page.support[0]], [], [page.support[1]], `${page.name} diagnostic context`, 32, mainY, contextWidth, mainHeight, z++);
    if (page.name === 'User Explorer') {
      contextVisual.visual.objects.columnFormatting = [{
        properties: { labelDisplayUnits: literal(1), labelPrecision: literal(0, 'L') },
        selector: { metadata: 'User Inventory Evidence.Enabled Workforce Accounts' }
      }];
    }
    visuals.push(contextVisual);
    visuals.push(image(`${pageId}:context-icon`, 'smartworkplace-kpi-relationships-20260913.svg', `${page.name} diagnostic context icon`, 32 + contextWidth - 40, mainY + 12, 24, 24, 912));
  } else {
    visuals.push(shapeContainer(`${pageId}:security-evidence-frame`, 32, mainY, 400, mainHeight, z++));
    visuals.push(textbox(`${pageId}:security-evidence-title`, 'Security evidence qualification', 44, mainY + 10, 300, 28, z++, { size: '13px', color: '#172B4D', bold: true }));
    visuals.push(card(`${pageId}:security-evidence-coverage`, [M('Security Control Evidence', 'Observed Security Evidence Coverage (%)')], '', 40, mainY + 48, 384, 104, z++, { valueSize: 24, labelSize: 11, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(card(`${pageId}:security-evidence-observed`, [M('Security Control Evidence', 'Observed Security Controls')], '', 40, mainY + 168, 184, 112, z++, { valueSize: 22, labelSize: 10, valueColor: '#107C10', accentColor: '#107C10', bare: true }));
    visuals.push(card(`${pageId}:security-evidence-missing`, [M('Security Control Evidence', 'Security Controls Without Evidence')], '', 232, mainY + 168, 192, 112, z++, { valueSize: 22, labelSize: 10, valueColor: '#D83B01', accentColor: '#D83B01', bare: true }));
    visuals.push(textbox(`${pageId}:security-evidence-note`, 'Observed controls are decision-qualified. Missing evidence remains explicitly blocked from conclusions.', 44, mainY + 296, 376, 72, z++, { size: '11px', color: '#52606D' }));
    visuals.push(image(`${pageId}:security-evidence-icon`, 'smartworkplace-kpi-quality-20260913.svg', 'Security evidence qualification icon', 392, mainY + 12, 24, 24, 912));
  }
  const heroX = isRiskPage ? 616 : isWideDetailExplorer ? 1480 : 1400;
  const heroWidth = isRiskPage ? 760 : isWideDetailExplorer ? 408 : 488;
  const detailX = isRiskPage ? 1400 : isSecurityPage ? 456 : isWideDetailExplorer ? 376 : 616;
  const detailWidth = isRiskPage ? 488 : isSecurityPage ? 920 : isWideDetailExplorer ? 1080 : 760;
  const hasIdentitySecuritySummary = Array.isArray(page.identitySecurityKpis);
  const hasDeviceSecuritySummary = Array.isArray(page.deviceSecurityKpis);
  const heroHeight = hasDeviceSecuritySummary ? 200 : isSecurityPage || hasIdentitySecuritySummary ? 168 : mainHeight;
  visuals.push((kind === 'line' ? line : bar)(`${pageId}:hero`, category, value, title, heroX, mainY, heroWidth, heroHeight, z++, page.heroOptions));
  visuals.push(image(`${pageId}:hero-icon`, pageIcon, `${title} icon`, heroX + heroWidth - 40, mainY + 12, 24, 24, 913));
  const detailHeight = page.manufacturerShare ? 200 : mainHeight;
  visuals.push(tableVisual(`${pageId}:detail`, page.detail, `${page.name} detail`, detailX, mainY, detailWidth, detailHeight, z++));
  visuals.push(image(`${pageId}:detail-icon`, 'smartworkplace-section-findings-20260913.svg', `${page.name} detail icon`, detailX + detailWidth - 40, mainY + 12, 24, 24, 914));
  if (page.manufacturerShare) {
    const compositionWidth = page.windowsVersionShare ? 368 : detailWidth;
    visuals.push(bar(`${pageId}:manufacturer-share`, page.manufacturerShare[0], page.manufacturerShare[1], page.manufacturerShare[2], detailX, mainY + 216, compositionWidth, 184, z++));
    visuals.push(image(`${pageId}:manufacturer-share-icon`, 'smartworkplace-kpi-devices-20260913.svg', `${page.manufacturerShare[2]} icon`, detailX + compositionWidth - 40, mainY + 228, 24, 24, 918));
    if (page.windowsVersionShare) {
      const windowsVersionX = detailX + compositionWidth + 24;
      visuals.push(bar(`${pageId}:windows-version-share`, page.windowsVersionShare[0], page.windowsVersionShare[1], page.windowsVersionShare[2], windowsVersionX, mainY + 216, compositionWidth, 184, z++, { categoryFontSize: 9 }));
      visuals.push(image(`${pageId}:windows-version-share-icon`, 'smartworkplace-lifecycle-20260913.svg', `${page.windowsVersionShare[2]} icon`, windowsVersionX + compositionWidth - 40, mainY + 228, 24, 24, 918));
    }
  }
  if (isSecurityPage && page.secondaryKpis) {
    const secondaryY = mainY + 184;
    visuals.push(shapeContainer(`${pageId}:secondary-security-frame`, heroX, secondaryY, heroWidth, 216, z++));
    visuals.push(textbox(`${pageId}:secondary-security-title`, 'Secondary security indicators', heroX + 12, secondaryY + 8, 360, 24, z++, { size: '13px', color: '#172B4D', bold: true }));
    visuals.push(textbox(`${pageId}:secondary-security-endpoint-label`, 'Endpoint protection', heroX + 12, secondaryY + 34, 220, 20, z++, { size: '10px', color: '#52606D', bold: true }));
    visuals.push(card(`${pageId}:secondary-security-endpoint`, page.secondaryKpis[0], '', heroX + 8, secondaryY + 54, heroWidth - 16, 62, z++, { valueSize: 16, labelSize: 10, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(textbox(`${pageId}:secondary-security-identity-label`, 'Identity protection', heroX + 12, secondaryY + 122, 220, 20, z++, { size: '10px', color: '#52606D', bold: true }));
    visuals.push(card(`${pageId}:secondary-security-identity`, page.secondaryKpis[1], '', heroX + 8, secondaryY + 142, heroWidth - 16, 62, z++, { valueSize: 17, labelSize: 10, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(image(`${pageId}:secondary-security-icon`, 'smartworkplace-health-20260913.svg', 'Secondary security indicators icon', heroX + heroWidth - 40, secondaryY + 8, 24, 24, 918));
  }
  if (hasIdentitySecuritySummary) {
    const identitySecurityY = mainY + 184;
    visuals.push(shapeContainer(`${pageId}:identity-security-frame`, heroX, identitySecurityY, heroWidth, 216, z++));
    visuals.push(textbox(`${pageId}:identity-security-title`, page.identitySummaryTitle || 'Identity security posture', heroX + 12, identitySecurityY + 8, 320, 24, z++, { size: '13px', color: '#172B4D', bold: true }));
    visuals.push(textbox(`${pageId}:identity-security-coverage-label`, page.identitySummaryGroup1 || 'Coverage and privilege', heroX + 12, identitySecurityY + 34, 240, 20, z++, { size: '10px', color: '#52606D', bold: true }));
    visuals.push(card(`${pageId}:identity-security-coverage`, page.identitySecurityKpis[0], '', heroX + 8, identitySecurityY + 54, heroWidth - 16, 62, z++, { valueSize: 17, labelSize: 10, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(textbox(`${pageId}:identity-security-authentication-label`, page.identitySummaryGroup2 || 'Authentication readiness', heroX + 12, identitySecurityY + 122, 260, 20, z++, { size: '10px', color: '#52606D', bold: true }));
    visuals.push(card(`${pageId}:identity-security-authentication`, page.identitySecurityKpis[1], '', heroX + 8, identitySecurityY + 142, heroWidth - 16, 62, z++, { valueSize: 17, labelSize: 10, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(image(`${pageId}:identity-security-icon`, 'smartworkplace-health-20260913.svg', 'Identity security posture icon', heroX + heroWidth - 40, identitySecurityY + 8, 24, 24, 919));
  }
  if (hasDeviceSecuritySummary) {
    const deviceSecurityY = mainY + 216;
    visuals.push(shapeContainer(`${pageId}:device-security-frame`, heroX, deviceSecurityY, heroWidth, 184, z++));
    visuals.push(textbox(`${pageId}:device-security-title`, 'Device security posture', heroX + 12, deviceSecurityY + 8, 320, 24, z++, { size: '13px', color: '#172B4D', bold: true }));
    visuals.push(textbox(`${pageId}:device-security-coverage-label`, 'Endpoint protection coverage', heroX + 12, deviceSecurityY + 34, 240, 18, z++, { size: '10px', color: '#52606D', bold: true }));
    visuals.push(card(`${pageId}:device-security-coverage`, page.deviceSecurityKpis[0], '', heroX + 8, deviceSecurityY + 52, heroWidth - 16, 50, z++, { valueSize: 16, labelSize: 9, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(textbox(`${pageId}:device-security-health-label`, 'Protection health', heroX + 12, deviceSecurityY + 106, 220, 18, z++, { size: '10px', color: '#52606D', bold: true }));
    visuals.push(card(`${pageId}:device-security-health`, page.deviceSecurityKpis[1], '', heroX + 8, deviceSecurityY + 124, heroWidth - 16, 52, z++, { valueSize: 16, labelSize: 9, valueColor: '#2F6B9A', accentColor: '#2F6B9A', bare: true }));
    visuals.push(image(`${pageId}:device-security-icon`, 'smartworkplace-health-20260913.svg', 'Device security posture icon', heroX + heroWidth - 40, deviceSecurityY + 8, 24, 24, 920));
  }

  const bottomY = 680;
  const bottomFields = page.bottom || [
    C('Findings', 'Severity Label'),
    C('Finding Type', 'Signal Type'),
    C('Findings', 'Finding Category'),
    C('Finding Type', 'Potential Impact'),
    M('Findings', '# Affected Entities'),
    C('Findings', 'Recommended Action')
  ];
  const signalFilters = page.operationalSignals && page.bottomSource && page.bottomSource !== 'Findings'
    ? [categoricalFilter(`${pageId}:signals:operational-context`, page.bottomSource, 'Signal', page.operationalSignals)]
    : (!page.bottomSource || page.bottomSource === 'Findings') && page.signalFilter
      ? [categoricalFilter(`${pageId}:signals:context`, 'Findings', page.signalFilter[0], page.signalFilter[1])]
      : [];
  const hasTrend = Boolean(page.trend);
  const hasBottomChart = isDeviceExplorer && Boolean(page.windowsVersionShare);
  const compactSignalsPages = new Set();
  const useCompactSignals = compactSignalsPages.has(page.name);
  const signalsX = isDeviceExplorer ? detailX : 32;
  const signalsWidth = isDeviceExplorer ? detailWidth : hasTrend ? (isWideDetailExplorer ? 1424 : 1344) : hasBottomChart ? 1424 : 1856;
  if (isDeviceExplorer) {
    visuals.push(matrix(`${pageId}:attention-context`, [C('Device Inventory Evidence', 'Disk Space State')], [], [M('Device Inventory Evidence', 'Managed Devices', '# Devices')], 'Device attention context', 32, bottomY, contextWidth, 328, z++));
    visuals.push(image(`${pageId}:attention-context-icon`, 'smartworkplace-health-20260913.svg', 'Device attention context icon', 32 + contextWidth - 40, bottomY + 12, 24, 24, 915));
  }
  if (useCompactSignals) {
    visuals.push(card(`${pageId}:signals-summary`, [M('Findings', 'Risk and Incident Summary')], 'Risk and incident status', signalsX, bottomY, signalsWidth, 104, z++, { valueSize: 18, labelSize: 10, valueColor: '#2F6B9A', accentColor: '#2F6B9A', showLabel: false, filters: signalFilters, altText: `Compact evidence-backed signal status for ${page.name}.` }));
    visuals.push(image(`${pageId}:signals-icon`, 'smartworkplace-health-20260913.svg', 'No open evidence-backed signals icon', signalsX + signalsWidth - 40, bottomY + 12, 24, 24, 915));
  } else {
    visuals.push(tableVisual(`${pageId}:signals`, bottomFields, page.signalTitle || 'Risk and incident signals', signalsX, bottomY, signalsWidth, 328, z++, {
      filters: signalFilters,
      ...(page.operationalSignals && page.bottomSource && page.bottomSource !== 'Findings' ? { sortBy: C(page.bottomSource, 'Signal') } : {})
    }));
    visuals.push(image(`${pageId}:signals-icon`, 'smartworkplace-section-findings-20260913.svg', 'Risk and incident signals icon', signalsX + signalsWidth - 40, bottomY + 12, 24, 24, 915));
  }
  if (hasTrend) {
    const trend = page.trend;
    const trendX = isWideDetailExplorer ? heroX : 1400;
    const trendWidth = isWideDetailExplorer ? heroWidth : 488;
    const trendChangeWidth = isSecurityPage ? 160 : 132;
    const trendLineWidth = isWideDetailExplorer ? 240 : isSecurityPage ? 288 : 320;
    visuals.push(shapeContainer(`${pageId}:trend-frame`, trendX, bottomY, trendWidth, 328, z++));
    visuals.push(textbox(`${pageId}:trend-title`, trend.title, trendX + 12, bottomY + 8, trendLineWidth, 24, z++, { size: '13px', color: '#172B4D', bold: true }));
    visuals.push(textbox(`${pageId}:trend-series`, `${trend.seriesLabel} · observed history when available`, trendX + 12, bottomY + 32, trendLineWidth, 18, z++, { size: '10px', color: '#52606D' }));
    visuals.push(sparkline(`${pageId}:trend-line`, trend.category || C('Date', 'Date'), trend.value, trend.tooltips || [trend.delta, M('Executive KPI Trends', 'Executive Trend Evidence Status')], trendX + 12, bottomY + 52, trendLineWidth, 238, z++, trend.color, { showCategoryAxis: true, axisFontSize: 9, axisType: 'Scalar' }));
    visuals.push(card(`${pageId}:trend-change`, [trend.change], 'Period change', trendX + trendWidth - trendChangeWidth - 12, bottomY + 48, trendChangeWidth, 76, z++, { valueSize: isSecurityPage ? 11 : 14, labelSize: 8, valueColor: trend.color, showLabel: false, titleSize: 9, bare: true, altText: `Period change for ${trend.seriesLabel}.` }));
    visuals.push(image(`${pageId}:trend-icon`, 'smartworkplace-kpi-analytics-20260913.svg', `${trend.title} icon`, 1848, bottomY + 12, 24, 24, 917));
  }
  if (hasBottomChart) {
    visuals.push(bar(`${pageId}:windows-version-share`, page.windowsVersionShare[0], page.windowsVersionShare[1], page.windowsVersionShare[2], heroX, bottomY, heroWidth, 328, z++, { categoryFontSize: 9 }));
    visuals.push(image(`${pageId}:windows-version-share-icon`, 'smartworkplace-lifecycle-20260913.svg', `${page.windowsVersionShare[2]} icon`, heroX + heroWidth - 40, bottomY + 12, 24, 24, 916));
  }
  visuals.push(textbox(`${pageId}:footer`, 'Product version · Data as of · Evidence status · © 2026 WorkplaceCloudHub — https://workplacecloudhub.com/', 32, 1032, 1856, 28, z++, { size: '11px', color: '#6B7280', align: 'center' }));

  const pageJson = {
    $schema: schemas.page,
    name: pageId,
    displayName: page.name,
    displayOption: 'FitToPage',
    height: 1080,
    width: 1920,
    objects: {
      background: [{ properties: { color: color('#F4F7FB'), transparency: literal(0) } }],
      outspace: [{ properties: { color: color('#E8EEF5'), transparency: literal(0) } }]
    }
  };
  if (page.drillthrough) {
    const filterName = `Filter${stableId(`${pageId}:drill`).padEnd(24, '0').slice(0, 24)}`;
    pageJson.filterConfig = { filters: [{ name: filterName, field: columnField(page.drillthrough.table, page.drillthrough.name).field, type: 'Categorical', howCreated: 'Drillthrough' }] };
    pageJson.pageBinding = { name: 'Pod', type: 'Drillthrough', parameters: [{ name: `Param_${filterName}`, boundFilter: filterName, fieldExpr: columnField(page.drillthrough.table, page.drillthrough.name).field }] };
  }
  const pageDir = path.join(pagesRoot, pageId);
  writeJson(path.join(pageDir, 'page.json'), pageJson);
  for (const visual of visuals) writeJson(path.join(pageDir, 'visuals', visual.name, 'visual.json'), visual);
  return { pageId, visualCount: visuals.length };
}

ensureInside(definitionRoot);
if (fs.existsSync(definitionRoot)) fs.rmSync(definitionRoot, { recursive: true, force: true });
fs.mkdirSync(pagesRoot, { recursive: true });
const approvedPageOrder = [
  'Executive Overview',
  'Licensing & Cost',
  'Risk & Incident Signals',
  'Security Posture',
  'Control & Evidence Explorer',
  'Workforce & Identity',
  'User Explorer',
  'Devices & Compliance',
  'Device Explorer',
  'Windows Lifecycle',
  'Endpoint Experience',
  'Applications & Standardization',
  'Application Explorer',
  'Collaboration Adoption',
  'Service & Content Explorer',
  'Content & Storage',
  'Messaging & Hybrid',
  'Mailbox Explorer',
  'Backup & Resilience',
  'Data Trust'
];
const pageByName = new Map(pages.map(page => [page.name, page]));
if (approvedPageOrder.length !== pages.length || approvedPageOrder.some(name => !pageByName.has(name))) {
  throw new Error('Approved page order does not match the configured report pages.');
}
const buildResults = approvedPageOrder.map(name => buildPage(pageByName.get(name)));
const pageOrder = buildResults.map(result => result.pageId);

writeJson(path.join(definitionRoot, 'version.json'), { $schema: schemas.version, version: '2.0.0' });
writeJson(path.join(pagesRoot, 'pages.json'), { $schema: schemas.pages, pageOrder, activePageName: pageOrder[0] });

const themeFile = 'SmartWorkplaceIntelligence-CorporateCool-8f4e1c.json';
writeJson(path.join(resourcesRoot, themeFile), {
  name: themeFile,
  dataColors: ['#0078D4', '#005A9E', '#2B88D8', '#00B7C3', '#107C10', '#FFB900', '#D83B01', '#A4262C'],
  background: '#F4F7FB',
  foreground: '#1F2937',
  tableAccent: '#005A9E',
  good: '#107C10',
  neutral: '#FFB900',
  bad: '#D83B01',
  maximum: '#005A9E',
  center: '#DCEAF7',
  minimum: '#A4262C',
  textClasses: {
    callout: { fontFace: 'Segoe UI Semibold', fontSize: 24, color: '#005A9E' },
    title: { fontFace: 'Segoe UI Semibold', fontSize: 13, color: '#1F2937' },
    label: { fontFace: 'Segoe UI', fontSize: 10, color: '#52606D' }
  },
  visualStyles: {
    '*': { '*': { visualHeader: [{ show: false }], border: [{ show: true, color: { solid: { color: '#CBD8E5' } }, radius: 12 }], padding: [{ top: 8, bottom: 8, left: 10, right: 10 }] } },
    tableEx: { '*': { columnHeaders: [{ autoSizeColumnWidth: true, columnAdjustment: 'growToFit' }] } },
    pivotTable: { '*': { columnHeaders: [{ autoSizeColumnWidth: true, columnAdjustment: 'growToFit' }] } }
  }
});

writeJson(path.join(definitionRoot, 'report.json'), {
  $schema: schemas.report,
  themeCollection: {
    baseTheme: { name: 'CY24SU06', reportVersionAtImport: { visual: '1.8.92', report: '2.0.92', page: '1.3.92' }, type: 'SharedResources' },
    customTheme: { name: themeFile, reportVersionAtImport: { visual: '2.6.0', report: '3.1.0', page: '2.3.0' }, type: 'RegisteredResources' }
  },
  resourcePackages: [
    { name: 'SharedResources', type: 'SharedResources', items: [{ name: 'CY24SU06', path: 'BaseThemes/CY24SU06.json', type: 'BaseTheme' }] },
    { name: 'RegisteredResources', type: 'RegisteredResources', items: [
      { name: themeFile, path: themeFile, type: 'CustomTheme' },
      ...iconFiles.map(name => ({ name, path: name, type: 'Image' }))
    ] }
  ],
  settings: { useStylableVisualContainerHeader: true, defaultDrillFilterOtherVisuals: true, allowChangeFilterTypes: true, useEnhancedTooltips: true, useDefaultAggregateDisplayName: true, pagesPosition: 'Bottom' },
  filterConfig: { filterSortOrder: 'Custom' },
  annotations: [{ name: 'defaultPage', value: pageOrder[0] }]
});

writeJson(path.join(reportRoot, '.platform'), {
  $schema: 'https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json',
  metadata: {
    type: 'Report',
    displayName: 'SmartWorkplaceIntelligence BETA 1.0.0-beta.1',
    description: 'Source-backed Smart Workplace Intelligence BETA report built from current private evidence and observed histories where available.'
  },
  config: {
    version: '2.0',
    logicalId: '8f4e1c3a-7b21-4b7d-9a44-02d7fb0ee201'
  }
});

process.stdout.write(JSON.stringify({ reportRoot, pages: pages.length, visuals: buildResults.reduce((sum, result) => sum + result.visualCount, 0), pageOrder }, null, 2));
