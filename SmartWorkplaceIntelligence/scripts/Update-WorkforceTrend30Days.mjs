import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pages = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');
const executive = path.join(pages, '74096d393d05fa994d48', 'visuals');
const workforce = path.join(pages, '0443c39bb6eba584da07', 'visuals');

function read(base, id) {
  const file = path.join(base, id, 'visual.json');
  const json = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (json.name !== id) throw new Error(`Unexpected visual ID in ${file}`);
  return {file, json};
}
function save(item) { fs.writeFileSync(item.file, `${JSON.stringify(item.json, null, 2)}\n`); }
function copyDateFilter(source, target, name) {
  const filter = source.json.filterConfig?.filters?.find(f =>
    f.type === 'RelativeDate' && f.field?.Column?.Property === 'Snapshot Date');
  if (!filter) throw new Error(`Missing 30-day reference filter in ${source.file}`);
  const copy = structuredClone(filter);
  copy.name = name;
  const existing = target.json.filterConfig?.filters ?? [];
  target.json.filterConfig = {
    ...target.json.filterConfig,
    filters: [...existing.filter(f => f.field?.Column?.Property !== 'Snapshot Date'), copy]
  };
}

const referenceChart = read(executive, 'fab0058c078b4f12ccc2');
const referenceChange = read(executive, '31e1000ebd39fefba25c');
const chart = read(workforce, '5565e85fb1fe7df8eb89');
const change = read(workforce, '66a3b9cad96ac677b1b0');
const title = read(workforce, '02cfd1477315674f24a7');

const expected = 'Workforce Trend Evidence.Trend Active Workforce';
if (chart.json.visual.query.queryState.Y.projections[0].queryRef !== expected) {
  throw new Error(`Unexpected workforce chart binding in ${chart.file}`);
}
if (change.json.visual.query.queryState.Data.projections[0].queryRef !==
    'Workforce Trend Evidence.Active Workforce Period Change') {
  throw new Error(`Unexpected workforce change binding in ${change.file}`);
}

copyDateFilter(referenceChart, chart, 'Filter30DaysWorkforcePageChart');
copyDateFilter(referenceChange, change, 'Filter30DaysWorkforcePageChange');
chart.json.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
  "'Active workforce by observed snapshot date, limited to the last 30 calendar days.'";
change.json.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
  "'Active workforce change between the first and last observed snapshots in the last 30 calendar days.'";
title.json.visual.objects.general[0].properties.paragraphs[0].textRuns[0].value =
  'Active workforce trend · last 30 days';

for (const item of [chart, change, title]) save(item);
process.stdout.write('Updated Workforce & Identity trend and change to the last 30 days.\n');
