import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const pageDir = path.join(
  path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..'),
  'pbip', 'SmartWorkplaceIntelligence.Report',
  'definition', 'pages', '74096d393d05fa994d48', 'visuals'
);

const ids = {
  workforceChart: 'fab0058c078b4f12ccc2',
  deviceChart: 'cf5c06152355cce791eb',
  windowsChart: 'f9b63098ff94d5610ef7',
  workforceChange: '31e1000ebd39fefba25c',
  deviceChange: '8a6e8f6b30bbb0431c60',
  windowsChange: '447e1a02d029d61ba1c1'
};

function visualFile(id) { return path.join(pageDir, id, 'visual.json'); }
function read(id) {
  const file = visualFile(id);
  const document = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (document.name !== id) throw new Error(`Unexpected visual at ${file}`);
  return {file, document};
}
function write(item) { fs.writeFileSync(item.file, `${JSON.stringify(item.document, null, 2)}\n`); }
function field(table, property, kind = 'Column') {
  return {[kind]: {Expression: {SourceRef: {Entity: table}}, Property: property}};
}
function projection(table, property, kind = 'Measure', active = false) {
  return {
    field: field(table, property, kind),
    queryRef: `${table}.${property}`,
    nativeQueryRef: property,
    ...(active ? {active: true} : {})
  };
}
function dateFilter(table, suffix) {
  const alias = 't';
  return {
    name: `Filter30Days${suffix}`,
    field: field(table, 'Snapshot Date'),
    type: 'RelativeDate',
    filter: {
      Version: 2,
      From: [{Name: alias, Entity: table, Type: 0}],
      Where: [{Condition: {Between: {
        Expression: {Column: {Expression: {SourceRef: {Source: alias}}, Property: 'Snapshot Date'}},
        LowerBound: {DateSpan: {Expression: {DateAdd: {Expression: {Now: {}}, Amount: -30, TimeUnit: 0}}, TimeUnit: 0}},
        UpperBound: {DateSpan: {Expression: {Now: {}}, TimeUnit: 0}}
      }}}]
    },
    howCreated: 'User'
  };
}
function setLast30Days(item, table, suffix) {
  const existing = item.document.filterConfig?.filters ?? [];
  const withoutDate = existing.filter(f => f.field?.Column?.Property !== 'Snapshot Date');
  item.document.filterConfig = {
    ...item.document.filterConfig,
    filters: [...withoutDate, dateFilter(table, suffix)]
  };
}
function assertBinding(item, role, table, property) {
  const actual = item.document.visual.query.queryState[role].projections[0].queryRef;
  if (actual !== `${table}.${property}`) throw new Error(`Unexpected ${role} binding in ${item.file}: ${actual}`);
}
function setTitle(item, title, alt) {
  item.document.visual.visualContainerObjects.title[0].properties.text.expr.Literal.Value = `'${title}'`;
  item.document.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value = `'${alt}'`;
}

const v = Object.fromEntries(Object.entries(ids).map(([key, id]) => [key, read(id)]));
assertBinding(v.workforceChart, 'Y', 'Workforce Trend Evidence', 'Trend Active Workforce');
assertBinding(v.deviceChart, 'Y', 'Endpoint Experience Trend Evidence', 'Trend Endpoint Analytics Score');
assertBinding(v.windowsChart, 'Y', 'Windows Lifecycle Trend Evidence', 'Trend Windows 11 Adoption');
assertBinding(v.workforceChange, 'Data', 'Workforce Trend Evidence', 'Active Workforce Period Change');
assertBinding(v.deviceChange, 'Data', 'Endpoint Experience Trend Evidence', 'Endpoint Analytics Score Period Change');
assertBinding(v.windowsChange, 'Data', 'Windows Lifecycle Trend Evidence', 'Windows 11 Adoption Period Change');

const lifecycle = 'Windows Lifecycle Trend Evidence';
setLast30Days(v.workforceChart, 'Workforce Trend Evidence', 'WorkforceChart');
setLast30Days(v.workforceChange, 'Workforce Trend Evidence', 'WorkforceChange');
setLast30Days(v.deviceChart, lifecycle, 'DeviceChart');
setLast30Days(v.deviceChange, lifecycle, 'DeviceChange');
setLast30Days(v.windowsChart, lifecycle, 'WindowsChart');
setLast30Days(v.windowsChange, lifecycle, 'WindowsChange');

const deviceQuery = v.deviceChart.document.visual.query;
deviceQuery.queryState.Category.projections = [projection(lifecycle, 'Snapshot Date', 'Column', true)];
deviceQuery.queryState.Tooltips.projections = [
  projection(lifecycle, 'Trend Windows 10 Devices'),
  projection(lifecycle, 'Trend Windows 11 Devices')
];
deviceQuery.queryState.Y.projections = [projection(lifecycle, 'Trend Managed Devices')];
deviceQuery.sortDefinition.sort[0].field = field(lifecycle, 'Snapshot Date');
const changeQuery = v.deviceChange.document.visual.query.queryState.Data;
changeQuery.projections = [projection(lifecycle, 'Managed Devices Period Change')];

setTitle(v.workforceChart, 'Active workforce trend · last 30 days', 'Active workforce by observed snapshot date, limited to the last 30 calendar days.');
setTitle(v.deviceChart, 'Managed devices trend · last 30 days', 'Managed device count by observed snapshot date, limited to the last 30 calendar days.');
setTitle(v.windowsChart, 'Windows 11 adoption trend · last 30 days', 'Windows 11 adoption by observed snapshot date, limited to the last 30 calendar days.');
setTitle(v.deviceChange, 'Period change', 'Managed device count change between the first and last displayed snapshots.');

for (const item of Object.values(v)) write(item);
process.stdout.write(`Updated ${Object.keys(v).length} Executive Overview visuals.\n`);
