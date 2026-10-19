// Builds only the approved migration page. Does not open or refresh Desktop.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const root = path.resolve(process.argv[2] || path.join(__dirname, '../pbip'));
const report = path.join(root, 'SmartWorkplaceIntelligence.Report/definition');
const pageId = 'migration-to-sharepoint-online';
const read = p => JSON.parse(fs.readFileSync(p, 'utf8').replace(/^\uFEFF/, ''));
const write = (p, value) => { fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, JSON.stringify(value, null, 2) + '\n'); };
const id = name => crypto.createHash('sha1').update('spo-migration/' + name).digest('hex').slice(0, 20);
const lit = value => ({ expr: { Literal: { Value: typeof value === 'string' ? "'" + value.replaceAll("'", "''") + "'" : typeof value === 'boolean' ? String(value) : `${value}D` } } });
const color = value => ({ solid: { color: lit(value) } });
const entry = properties => [{ properties }];
const selected = (properties, selector = 'default') => [{ selector: { id: selector }, properties }];
const column = (table, name, label = name) => ({ field: { Column: { Expression: { SourceRef: { Entity: table } }, Property: name } }, queryRef: `${table}.${name}`, nativeQueryRef: label, displayName: label });
const measureTables = {};
// TOM definitions are neutral; this reads object metadata, never tenant data.
const child = require('child_process').spawnSync(process.env.SMARTWORKPLACE_PYTHON || 'python', ['-B', path.join(__dirname, 'migration_model_definitions.py')], { encoding: 'utf8' });
if (child.status !== 0) throw new Error(child.stderr || 'Model definition helper failed');
for (const m of JSON.parse(child.stdout).measures) measureTables[m.name] = m.tableName;
const measure = (name, label = name.replace(/^SPO /, '')) => {
  if (!measureTables[name]) throw new Error('Unknown migration measure: ' + name);
  return { field: { Measure: { Expression: { SourceRef: { Entity: measureTables[name] } }, Property: name } }, queryRef: `${measureTables[name]}.${name}`, nativeQueryRef: label, displayName: label };
};
const sample = path.join(report, 'pages/daily-directory-mailbox-trends');
const examples = fs.readdirSync(path.join(sample, 'visuals')).map(n => read(path.join(sample, 'visuals', n, 'visual.json')));
const visualSchema = examples.find(v => v.visual?.visualType === 'cardVisual').$schema;
const page = { ...read(path.join(sample, 'page.json')), name: pageId, displayName: 'Migration to SharePoint Online' };
page.objects.background = entry({ color: color('#FFFFFF'), transparency: lit(0) });
const visuals = [];
let order = 10;
function chrome(title, background = '#FFFFFF') {
  return {
    title: entry({ show: lit(Boolean(title)), text: lit(title || ''), fontSize: lit(14), fontColor: color('#183153') }),
    background: entry({ show: lit(true), color: color(background), transparency: lit(0) }),
    border: entry({ show: lit(true), color: color('#CCDCEC'), radius: lit(8), width: lit(1) }),
    padding: entry({ top: lit(8), bottom: lit(8), left: lit(8), right: lit(8) }),
    visualHeader: entry({ show: lit(true) }),
    general: entry({ altText: lit(title || 'Migration context') })
  };
}
function add(name, type, x, y, width, height, parent, query, objects = {}, title = '') {
  const value = { $schema: visualSchema, name: id(name), position: { x, y, width, height, z: order, tabOrder: order++ },
    ...(parent ? { parentGroupName: parent } : {}),
    visual: { visualType: type, ...(query ? { query: { queryState: query } } : {}), objects, visualContainerObjects: chrome(title), drillFilterOtherVisuals: true } };
  visuals.push(value);
  return value;
}
function group(name, parent, hidden) {
  const result = { $schema: visualSchema, name: id('group-' + name), position: { x: 0, y: 0, width: 1920, height: 1080, z: 1, tabOrder: 1 },
    visualGroup: { displayName: name, groupMode: 'ScaleMode' }, isHidden: hidden, ...(parent ? { parentGroupName: parent } : {}) };
  visuals.push(result);
  return result.name;
}
function card(name, metric, title, x, y, w, h, parent, accent = '#005A9E', tint = '#E6F4FF', font = 30) {
  const v = add(name, 'cardVisual', x, y, w, h, parent, { Data: { projections: [measure(metric)] } }, {
    value: selected({ fontSize: lit(font), fontColor: color(accent), textWrap: lit(true), show: lit(true), showBlankAs: lit('n/a') }),
    label: selected({ show: lit(false) }), outline: selected({ show: lit(false) }),
    fillCustom: selected({ show: lit(true), fillColor: color(tint), transparency: lit(0) })
  }, title);
  v.visual.visualContainerObjects.background = entry({ show: lit(true), color: color(tint), transparency: lit(0) });
  return v;
}
function textbox(name, text, x, y, w, h, parent, size = 14) {
  const v = add(name, 'textbox', x, y, w, h, parent, null, { general: entry({ paragraphs: [{ textRuns: [{ value: text, textStyle: { fontFamily: 'Segoe UI', fontSize: `${size}pt`, color: '#183153' } }] }] }) });
  v.visual.visualContainerObjects = { title: entry({ show: lit(false) }), background: entry({ show: lit(false) }), padding: entry({ top: lit(0), bottom: lit(0), left: lit(0), right: lit(0) }), general: entry({ altText: lit(text) }) };
  return v;
}
function slicer(name, table, col, title, x, y, w) {
  return add(name, 'slicer', x, y, w, 80, null, { Values: { projections: [column(table, col)] } }, {
    data: entry({ mode: lit('Dropdown') }), header: entry({ show: lit(true), text: lit(title), textSize: lit(12), fontColor: color('#183153') }),
    selection: entry({ selectAllCheckboxEnabled: lit(true), strictSingleSelect: lit(false) }), general: entry({ selfFilterEnabled: lit(true) })
  });
}
function chart(name, type, category, metrics, title, x, y, w, h, parent, series) {
  const query = { Category: { projections: category }, Y: { projections: metrics.map(m => measure(m)) } };
  if (series) query.Series = { projections: [series] };
  const colors = ['#005A9E', '#107C10', '#605E5C'];
  const objects = { categoryAxis: entry({ labelColor: color('#334155'), fontSize: lit(12) }), valueAxis: entry({ start: lit(0), labelColor: color('#334155'), fontSize: lit(12) }),
    legend: entry({ show: lit(Boolean(series) || metrics.length > 1) }),
    dataPoint: metrics.length > 1 ? metrics.map((m, i) => ({ selector: { metadata: measure(m).queryRef }, properties: { fill: color(colors[i % colors.length]) } })) : entry({ defaultColor: color('#005A9E') }) };
  return add(name, type, x, y, w, h, parent, query, objects, title);
}
function recordFilter(v, metric) {
  const m = measure(metric), field = m.field.Measure;
  v.filterConfig = { filters: [{ name: id(v.name + '-visible'), field: m.field, type: 'Advanced', filter: { Version: 2,
    From: [{ Name: 'm', Entity: field.Expression.SourceRef.Entity, Type: 0 }], Where: [{ Condition: { Comparison: { ComparisonKind: 2,
      Left: { Measure: { Expression: { SourceRef: { Source: 'm' } }, Property: field.Property } }, Right: { Literal: { Value: '0L' } } } } }] }, howCreated: 'User' }] };
}
function table(name, fields, title, x, y, w, h, parent, metricFilter) {
  const v = add(name, 'tableEx', x, y, w, h, parent, { Values: { projections: fields } }, {
    columnHeaders: entry({ columnAdjustment: lit('growToFit'), autoSizeColumnWidth: lit(true), wordWrap: lit(true), fontSize: lit(12), fontColor: color('#FFFFFF'), backColor: color('#005A9E'), bold: lit(true) }),
    values: entry({ fontSize: lit(14), wordWrap: lit(false), backColorPrimary: color('#FFFFFF'), backColorSecondary: color('#F3F7FA'), fontColorPrimary: color('#183153'), fontColorSecondary: color('#183153') }),
    total: entry({ totals: lit(false) })
  }, title);
  v.visual.visualContainerObjects.stylePreset = entry({ name: lit('None') });
  if (metricFilter) recordFilter(v, metricFilter);
  return v;
}
const names = ['Progress', 'Schedule', 'Reconciliation', 'Quality'];
const groups = names.map((n, i) => group(n, null, i !== 0));
textbox('title', 'Migration to SharePoint Online', 32, 24, 880, 42, null, 24);
const context = card('context', 'SPO Page Context', '', 32, 72, 880, 40, null, '#334155', '#FFFFFF', 11);
context.visual.visualContainerObjects.border = entry({ show: lit(false) });
context.visual.visualContainerObjects.padding = entry({ top: lit(0), bottom: lit(0), left: lit(0), right: lit(0) });
for (const [i, label] of ['Source Type', 'Project', 'Business Unit', 'Phase'].entries()) {
  slicer('filter-' + label, 'SPO Migration ' + label, label === 'Project' ? 'Project Key' : label, label, 944 + i * 240, 24, 224);
}
const bookmarksIndex = read(path.join(report, 'bookmarks/bookmarks.json'));
const bookmarkSchema = read(path.join(report, 'bookmarks/bf6e5721270ee694da00.bookmark.json')).$schema;
const bookmarks = [];
function bookmarksFor(key, labels, groupIds) {
  const groupName = id('bookmark-group-' + key);
  const markIds = labels.map(label => id('bookmark-' + key + '-' + label));
  bookmarksIndex.items = bookmarksIndex.items.filter(x => x.name !== groupName);
  bookmarksIndex.items.push({ name: groupName, displayName: 'Migration ' + key, children: markIds });
  labels.forEach((label, i) => bookmarks.push({ $schema: bookmarkSchema, name: markIds[i], displayName: label,
    options: { applyOnlyToTargetVisuals: true, targetVisualNames: groupIds, suppressActiveSection: true, suppressData: true },
    explorationState: { version: '1.0', activeSection: pageId, sections: { [pageId]: { visualContainers: {},
      visualContainerGroups: Object.fromEntries(groupIds.map((g, j) => [g, { isHidden: j !== i }])) } } } }));
  return { groupName, markIds };
}
function navigator(key, labels, groupIds, x, y, w, h, parent) {
  const marks = bookmarksFor(key, labels, groupIds);
  const template = structuredClone(examples.find(v => v.visual?.visualType === 'bookmarkNavigator'));
  template.name = id('navigator-' + key); template.position = { x, y, width: w, height: h, z: 100, tabOrder: 100 };
  if (parent) template.parentGroupName = parent;
  template.visual.objects.bookmarks = entry({ bookmarkGroup: lit(marks.groupName), selectedBookmark: lit(marks.markIds[0]), allowDeselectionBookmark: lit(false) });
  template.visual.objects.layout[0].properties.columnCount = lit(labels.length);
  template.visual.visualContainerObjects.general = entry({ altText: lit('Switch migration view; preserve all filter selections.') });
  visuals.push(template);
}
navigator('views', names, groups, 32, 128, 1640, 48, null);
const clear = add('clear-slicers', 'actionButton', 1696, 128, 192, 48, null, null, { text: selected({ show: lit(true), text: lit('Clear slicers'), fontSize: lit(14), fontColor: color('#183153') }) });
clear.visual.visualContainerObjects.visualLink = entry({ show: lit(true), type: lit('ClearAllSlicers'), tooltip: lit('Clear slicer selections while keeping the active view. Advanced filters have their own clear control.') });
const P = 'SPO Migration Plan', W = 'SPO Migration Web', F = 'SPO Migration Findings', Q = 'SPO Migration Source Summary';
const p = n => column(P, n);
const kpiSets = [
  [['SPO Reported Progress', 'Reported migration completion'], ['SPO Validated Progress', 'Validated migration completion'], ['SPO Remaining to Execute', 'Remaining to execute']],
  [['SPO Overdue', 'Overdue migrations'], ['SPO Due Within 14 Days', 'Due today through next 14 days'], ['SPO Done Missing Actual Date', 'DONE without actual date']],
  [['SPO Source Identity Review', 'On-prem source identity review'], ['SPO Target Review', 'Target mapping review'], ['SPO Webs Without Plan', 'Observed webs without a plan (global)']],
  [['SPO Operations With Findings', 'Operations with findings'], ['SPO Business Review Findings', 'Operations needing business review'], ['SPO Source Qualification Context', 'Source qualification']]
];
for (const [tab, set] of kpiSets.entries()) set.forEach(([m, title], i) => card('kpi-' + tab + '-' + i, m, title, 32 + i * 632, 208, 592, 128, groups[tab], ['#005A9E', '#107C10', '#605E5C'][i], ['#E6F4FF', '#E7F2E7', '#F3F2F1'][i], m.endsWith('Context') ? 18 : m.endsWith('Progress') ? 24 : 30));
chart('progress-execution', 'barChart', [column('SPO Migration Business Unit', 'Business Unit'), column('SPO Migration Phase', 'Phase')], ['SPO Migrate Operations'], 'MIGRATE execution by business unit / phase', 32, 360, 1200, 304, groups[0], p('Execution Status'));
chart('progress-decisions', 'barChart', [p('Disposition')], ['SPO Operations'], 'Disposition of all operations', 1256, 360, 632, 304, groups[0]);
table('progress-records', ['Migration ID', 'Source Type', 'Source Name', 'Source Reference', 'Disposition', 'Execution Status', 'Validation Status', 'Next Action'].map(p), 'Operations and next action — source URL or shared-folder path', 32, 696, 1856, 352, groups[0], 'SPO Operations');
const modes = [group('Deadlines', groups[1], false), group('Observed history', groups[1], true)];
navigator('schedule modes', ['Deadlines', 'Observed history'], modes, 32, 360, 1200, 40, groups[1]);
chart('schedule-deadlines', 'columnChart', [column('Date', 'Week Start')], ['SPO Planned Calendar Operations'], 'Declared planned dates by week (all MIGRATE operations)', 32, 408, 1200, 256, modes[0]);
const history = chart('schedule-observed', 'lineChart', [column('SPO Migration Snapshot Date', 'Date')], ['SPO Historical Reported Done', 'SPO Historical Validated Done'], 'Observed daily snapshots — markers only; gaps are unmeasured', 32, 408, 1200, 256, modes[1]);
history.visual.objects.categoryAxis[0].properties.axisType = lit('Scalar');
history.visual.objects.lineStyles = entry({ strokeShow: lit(false), showMarker: lit(true), markerShape: lit('circle'), markerSize: lit(7) });
history.visual.query.queryState.Tooltips = { projections: ['SPO Historical Migrate Operations', 'SPO Historical Completion', 'SPO Historical Removed Operations'].map(m => measure(m)) };
chart('schedule-cumulative', 'lineChart', [column('Date', 'Date')], ['SPO Cumulative Declared Completions'], 'Cumulative declared ActualDate completions', 1256, 360, 632, 304, groups[1]);
textbox('schedule-caveat', 'Observed history begins with the first prepared snapshot. ActualDate is a declared completion calendar, not reconstructed history.', 32, 672, 1856, 22, groups[1], 11);
table('schedule-records', ['Migration ID', 'Source Type', 'Source Name', 'Planned Date', 'Actual Date', 'Schedule Variance Days', 'Execution Status', 'Business Contact', 'Next Action'].map(p), 'Schedule and ownership — variance against the latest known planned date', 32, 696, 1856, 352, groups[1], 'SPO Operations');
const matrix = add('source-matrix', 'pivotTable', 32, 360, 1200, 304, groups[2], {
  Rows: { projections: [column('SPO Migration Source Type', 'Source Type')] }, Columns: { projections: [p('Source Qualification')] },
  Values: { projections: [measure('SPO Operations')] }
}, { columnHeaders: entry({ columnAdjustment: lit('growToFit'), autoSizeColumnWidth: lit(true), fontSize: lit(12), fontColor: color('#FFFFFF'), backColor: color('#005A9E') }),
  rowHeaders: entry({ fontSize: lit(14), fontColor: color('#183153'), stepped: lit(false) }), values: entry({ fontSize: lit(14) })
}, 'Operation source qualification');
matrix.visual.visualContainerObjects.stylePreset = entry({ name: lit('None') });
const coverage = chart('web-coverage', 'barChart', [column('SharePoint On-prem Web Applications', 'Web Application')], ['SPO Observed Webs'], 'On-prem observed web coverage', 1256, 360, 632, 304, groups[2], column(W, 'Coverage Kind'));
coverage.visual.query.queryState.Tooltips = { projections: ['SPO Current Collection Storage GiB', 'SPO Fully Qualified Migrating GiB', 'SPO Volume Context', 'SPO On-prem Coverage Context'].map(m => measure(m)) };
textbox('reconciliation-caveat', 'Usage-site observations do not verify migrated contents. Shared-folder existence, volume and effective ACLs are not measured.', 32, 672, 1856, 22, groups[2], 11);
table('reconciliation-records', ['Migration ID', 'Source Type', 'Source Reference', 'Source Qualification', 'Resolved Target URL', 'Target Mapping Level', 'Target Approval Evidence', 'Target Observation', 'Permission Strategy', 'Next Action'].map(n => n === 'Resolved Target URL' ? column(P, n, 'Target URL') : p(n)), 'Source / target evidence — approvals, observation and next action', 32, 696, 1856, 352, groups[2], 'SPO Operations');
chart('finding-rules', 'barChart', [column(F, 'Finding Rule')], ['SPO Finding Count'], 'Operation findings by rule — categories overlap', 32, 360, 1200, 304, groups[3]);
table('source-freshness', ['Source', 'Captured UTC', 'Age Hours', 'Qualification', 'Coverage State', 'Review State'].map(n => column(Q, n)), 'Source freshness / coverage / review', 1256, 360, 632, 304, groups[3]);
const global = card('global-findings', 'SPO Global Quality Context', '', 32, 672, 1856, 22, groups[3], '#334155', '#FFFFFF', 11);
global.visual.visualContainerObjects.border = entry({ show: lit(false) });
global.visual.visualContainerObjects.padding = entry({ top: lit(0), bottom: lit(0), left: lit(0), right: lit(0) });
table('quality-records', ['Severity', 'Finding Rule', 'Source Type', 'Migration ID', 'Source Name', 'Source Reference', 'Observed Value', 'Expected Action'].map(n => column(F, n)), 'Actionable findings — global findings remain visible', 32, 696, 1856, 352, groups[3], 'SPO Visible Finding');
const advanced = [
  ...['Disposition', 'Execution Status', 'Validation Status', 'Source Object Type', 'Source Qualification', 'Target Observation', 'Target Approval Evidence', 'Source Reference'].map(n => column(P, n)),
  column('SharePoint On-prem Farms', 'Farm ID'), column('SharePoint On-prem Web Applications', 'Web Application'),
  column('SPO Migration Collection', 'Source URL'), column('SPO Migration Snapshot Date', 'Date')
];
page.filterConfig = { filters: advanced.map(f => ({ name: id('advanced-' + f.queryRef), field: f.field, type: 'Categorical', howCreated: 'User' })) };
const pageRoot = path.join(report, 'pages', pageId);
const orderPath = path.join(report, 'pages/pages.json');
const pages = read(orderPath);
pages.pageOrder = pages.pageOrder.filter(p => p !== pageId);
const anchor = pages.pageOrder.find(p => read(path.join(report, 'pages', p, 'page.json')).displayName === 'On-prem Content & Storage');
if (!anchor) throw new Error('Approved insertion anchor not found');
pages.pageOrder.splice(pages.pageOrder.indexOf(anchor) + 1, 0, pageId);
if (visuals.length !== new Set(visuals.map(v => v.name)).size) throw new Error('Duplicate visual identity');
write(path.join(pageRoot, 'page.json'), page);
for (const v of visuals) write(path.join(pageRoot, 'visuals', v.name, 'visual.json'), v);
for (const b of bookmarks) write(path.join(report, 'bookmarks', b.name + '.bookmark.json'), b);
write(path.join(report, 'bookmarks/bookmarks.json'), bookmarksIndex);
write(orderPath, pages);
console.log(JSON.stringify({ page: pageId, pages: pages.pageOrder.length, visuals: visuals.length, bookmarks: bookmarks.length, refreshed: false }));
