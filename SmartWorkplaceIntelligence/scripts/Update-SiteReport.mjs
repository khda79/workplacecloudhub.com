// Narrow PBIR update for the approved three-page lot. No other pages are touched.
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');
const pages = {
  executive: '74096d393d05fa994d48',
  personas: '9ae7c13f821d4a608e52',
  user: '3a0fb4ff8eb05c8835b3',
};
const literal = value => ({ expr: { Literal: { Value: value } } });
function file(page, id) { return path.join(root, page, 'visuals', id, 'visual.json'); }
function read(page, id) { return JSON.parse(fs.readFileSync(file(page, id), 'utf8')); }
function write(page, id, visual) {
  const target = file(page, id);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, JSON.stringify(visual, null, 2) + '\n', 'utf8');
}
function setMeasure(visual, table, measure, label) {
  const projection = visual.visual.query.queryState.Data.projections[0];
  projection.field.Measure.Expression.SourceRef.Entity = table;
  projection.field.Measure.Property = measure;
  projection.queryRef = `${table}.${measure}`;
  projection.nativeQueryRef = label;
}
function measureProjection(table, measure, label) {
  return {
    field: { Measure: { Expression: { SourceRef: { Entity: table } }, Property: measure } },
    queryRef: `${table}.${measure}`,
    nativeQueryRef: label,
  };
}
function setCardTitle(visual, text) {
  const title = visual.visual.visualContainerObjects.title[0].properties;
  title.show = literal('true');
  title.text = literal(`'${text}'`);
}
function setSlicer(visual, name, table, field, label, x, width, z) {
  visual.name = name;
  visual.position = { ...visual.position, x, width, z, tabOrder: z };
  const projection = visual.visual.query.queryState.Values.projections[0];
  projection.field.Column.Expression.SourceRef.Entity = table;
  projection.field.Column.Property = field;
  projection.queryRef = `${table}.${field}`;
  projection.nativeQueryRef = field;
  visual.visual.objects.header[0].properties.text = literal(`'${label}'`);
  visual.visual.visualContainerObjects.general[0].properties.altText = literal(`'Filter by ${label.toLowerCase()}.'`);
  return visual;
}

for (const [id, title] of [
  ['fab0058c078b4f12ccc2', 'Active workforce trend'],
  ['cf5c06152355cce791eb', 'Managed devices trend'],
  ['f9b63098ff94d5610ef7', 'Windows 11 adoption trend'],
]) {
  const visual = read(pages.executive, id);
  visual.visual.objects.categoryAxis[0].properties.show = literal('true');
  visual.visual.objects.categoryAxis[0].properties.axisType = literal("'Scalar'");
  visual.visual.objects.categoryAxis[0].properties.fontSize = literal('8D');
  setCardTitle(visual, title);
  visual.visual.visualContainerObjects.title[0].properties.fontSize = literal('9D');
  visual.visual.visualContainerObjects.general[0].properties.altText = literal(`'${title} by snapshot date; date scale is shown on the horizontal axis.'`);
  write(pages.executive, id, visual);
}

for (const [id, valueSize] of [
  ['6b0b054d4011dc8f3732', '24D'],
  ['6854dad122544a331a47', '21D'],
  ['fc4687a2f85a9a7a7ab2', '24D'],
  ['e4d06fe45cb5c0cf9efe', '22D'],
  ['1f5377ed66dd2acf9e82', '21D'],
  ['415ba06e1fecc76819c7', '24D'],
  ['0385ca0d1c75d8020694', '30D'],
]) {
  const visual = read(pages.executive, id);
  visual.visual.objects.value[0].properties.fontSize = literal(valueSize);
  write(pages.executive, id, visual);
}
const encrypted = read(pages.executive, '12365a8234c915ca9682');
encrypted.position.height = 62;
setMeasure(encrypted, 'Device Inventory Evidence', 'Encrypted Devices Among Known Intune State (%)', 'Encrypted devices (%)');
encrypted.visual.query.queryState.Tooltips = { projections: [
  measureProjection('Device Inventory Evidence', 'Known Intune Encryption State Coverage (%)', 'Known-state coverage (%)'),
  measureProjection('Device Inventory Evidence', 'Devices Encrypted According to Intune', 'Encrypted devices'),
] };
encrypted.visual.visualContainerObjects.visualTooltip = [{ properties: { show: literal('true') } }];
if (encrypted.filterConfig?.filters?.[0]?.field?.Measure) {
  encrypted.filterConfig.filters[0].field.Measure.Property = 'Encrypted Devices Among Known Intune State (%)';
}
setCardTitle(encrypted, 'Encrypted devices (Intune)');
encrypted.visual.objects.value[0].properties.fontSize = literal('22D');
encrypted.visual.visualContainerObjects.general[0].properties.altText = literal("'Percentage of Intune devices marked encrypted among those with known encryption state. This is not proof of BitLocker key escrow; the card below shows known-state coverage.'");
write(pages.executive, '12365a8234c915ca9682', encrypted);
const coverage = structuredClone(encrypted);
coverage.name = '6ff6305cd8a44c2280cf';
coverage.position = { ...coverage.position, y: 654, height: 62, z: 127, tabOrder: 127 };
setMeasure(coverage, 'Device Inventory Evidence', 'Known Intune Encryption State Coverage (%)', 'Known-state coverage (%)');
coverage.visual.query.queryState.Tooltips = { projections: [
  measureProjection('Device Inventory Evidence', 'Devices Encrypted According to Intune', 'Encrypted devices'),
] };
if (coverage.filterConfig?.filters?.[0]?.field?.Measure) {
  coverage.filterConfig.filters[0].name = 'f9630d36ab864ac181de';
  coverage.filterConfig.filters[0].field.Measure.Property = 'Known Intune Encryption State Coverage (%)';
}
setCardTitle(coverage, 'Encryption state coverage');
coverage.visual.objects.value[0].properties.fontSize = literal('18D');
coverage.visual.visualContainerObjects.general[0].properties.altText = literal("'Share of managed Intune devices whose encryption state is known. Unknown states are excluded from the encrypted-device percentage.'");
write(pages.executive, coverage.name, coverage);

const personaSlicers = [
  ['521a0ef7c32945cc1c17', 816, 160],
  ['d3f8c14571af77ec3ce8', 992, 160],
  ['e8b3a6398343fd57e94f', 1168, 128],
  ['2cddfae8432a26339970', 1312, 128],
  ['5577d3f3e4e1ab0de03c', 1456, 112],
];
for (const [id, x, width] of personaSlicers) {
  const visual = read(pages.personas, id);
  visual.position.x = x;
  visual.position.width = width;
  write(pages.personas, id, visual);
}
const personaSite = setSlicer(structuredClone(read(pages.personas, '5577d3f3e4e1ab0de03c')),
  '75b76bd72caa4784af5e', 'User Inventory Evidence', 'Site Type', 'Site type', 1584, 112, 106);
write(pages.personas, personaSite.name, personaSite);

for (const [id, x, width] of [
  ['ce8fff0f0d93da4249c8', 1128, 128],
  ['a01af72d75320d14ca69', 1272, 128],
  ['51dc75f2eb71a770500f', 1416, 128],
]) {
  const visual = read(pages.user, id);
  visual.position.x = x;
  visual.position.width = width;
  write(pages.user, id, visual);
}
const userSite = setSlicer(structuredClone(read(pages.user, '51dc75f2eb71a770500f')),
  'c44dabba931c4fb49b67', 'User Inventory Evidence', 'Site Type', 'Site type', 1560, 160, 106);
write(pages.user, userSite.name, userSite);

process.stdout.write('Executive Overview, Workforce Personas and User Explorer visuals updated.\n');
