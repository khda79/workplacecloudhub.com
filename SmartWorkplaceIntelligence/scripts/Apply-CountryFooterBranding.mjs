import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

const projectRoot = path.resolve(import.meta.dirname, '..');
const reportRoot = path.join(projectRoot, 'pbip', 'SmartWorkplaceIntelligence.Report');
const pagesRoot = path.join(reportRoot, 'definition', 'pages');
const resourceRoot = path.join(reportRoot, 'StaticResources', 'RegisteredResources');
const sourceLogo = path.resolve(projectRoot, '..', 'WorkplaceCloudHub-lockup-WPF.png');
const resourceName = 'WorkplaceCloudHub-lockup-WPF-20260924.png';
const executiveId = '74096d393d05fa994d48';
const oldDeviceCountryId = '2dc37f365ba781ab25b4';
const workforceCountryId = '0d29a264e6543f511857';

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function writeJson(file, value) {
  fs.writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function literal(value) {
  return { expr: { Literal: { Value: value } } };
}

function within(parent, target) {
  const relative = path.relative(parent, target);
  return relative !== '' && !relative.startsWith('..') && !path.isAbsolute(relative);
}

if (!fs.existsSync(sourceLogo)) throw new Error(`Logo not found: ${sourceLogo}`);
const pageDirs = fs.readdirSync(pagesRoot, { withFileTypes: true }).filter((entry) => entry.isDirectory());
if (pageDirs.length !== 21) throw new Error(`Expected 21 report pages, got ${pageDirs.length}`);

const reportFile = path.join(reportRoot, 'definition', 'report.json');
const report = readJson(reportFile);
const resources = report.resourcePackages.find((item) => item.name === 'RegisteredResources');
if (!resources) throw new Error('RegisteredResources package is missing');
if (!resources.items.some((item) => item.name === resourceName)) {
  resources.items.push({ name: resourceName, path: resourceName, type: 'Image' });
  writeJson(reportFile, report);
}
const targetLogo = path.join(resourceRoot, resourceName);
if (!fs.existsSync(targetLogo)) fs.copyFileSync(sourceLogo, targetLogo);
if (!fs.readFileSync(targetLogo).equals(fs.readFileSync(sourceLogo))) {
  throw new Error('Registered logo differs from the supplied logo');
}

let changedFooters = 0;
let addedLogos = 0;
for (const entry of pageDirs) {
  const pageDir = path.join(pagesRoot, entry.name);
  const page = readJson(path.join(pageDir, 'page.json'));
  const visualsRoot = path.join(pageDir, 'visuals');
  const visuals = fs.readdirSync(visualsRoot, { withFileTypes: true })
    .filter((visualDir) => visualDir.isDirectory())
    .map((visualDir) => ({ id: visualDir.name, file: path.join(visualsRoot, visualDir.name, 'visual.json') }))
    .map(({ id, file }) => ({ id, file, definition: readJson(file) }));
  const footerMatches = visuals.filter(({ definition }) =>
    definition.visual?.visualType === 'textbox' &&
    definition.position?.y >= 1000 &&
    definition.visual.objects?.general?.[0]?.properties?.paragraphs?.[0]?.textRuns?.[0]?.value
  );
  if (footerMatches.length !== 1) throw new Error(`${page.displayName}: expected exactly one footer, found ${footerMatches.length}`);
  const footer = footerMatches[0];
  const paragraph = footer.definition.visual.objects.general[0].properties.paragraphs[0];
  const isLongMethodology = page.displayName === 'Licensing & Cost';
  footer.definition.position.x = isLongMethodology ? 160 : 640;
  footer.definition.position.y = 1032;
  footer.definition.position.width = isLongMethodology ? 1728 : 1248;
  footer.definition.position.height = 28;
  paragraph.horizontalTextAlignment = 'right';
  writeJson(footer.file, footer.definition);
  changedFooters++;

  const existingLogo = visuals.find(({ definition }) =>
    definition.visual?.visualType === 'image' &&
    definition.visual.objects?.general?.[0]?.properties?.imageUrl?.expr?.ResourcePackageItem?.ItemName === resourceName
  );
  if (existingLogo) continue;
  const id = crypto.randomBytes(10).toString('hex');
  const visualDir = path.join(visualsRoot, id);
  if (!within(visualsRoot, visualDir)) throw new Error(`Unsafe logo visual directory: ${visualDir}`);
  const maxOrder = Math.max(...visuals.map(({ definition }) => Number(definition.position?.z ?? 0)));
  const image = {
    $schema: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json',
    name: id,
    position: { x: 32, y: 1008, z: maxOrder + 1, width: 72, height: 60, tabOrder: maxOrder + 1 },
    visual: {
      visualType: 'image',
      objects: {
        general: [{ properties: { imageUrl: { expr: { ResourcePackageItem: {
          PackageName: 'RegisteredResources', PackageType: 1, ItemName: resourceName
        } } } } }]
      },
      visualContainerObjects: {
        general: [{ properties: { altText: literal("'WorkplaceCloudHub logo; replace with the customer logo when supplied.'") } }],
        background: [{ properties: { show: literal('true'), color: { solid: { color: literal("'#0B2345'") } }, transparency: literal('0D') } }],
        border: [{ properties: { show: literal('false'), radius: literal('6D') } }],
        visualHeader: [{ properties: { show: literal('false') } }],
        padding: [{ properties: { top: literal('0D'), bottom: literal('0D'), left: literal('0D'), right: literal('0D') } }]
      },
      drillFilterOtherVisuals: true
    }
  };
  fs.mkdirSync(visualDir);
  writeJson(path.join(visualDir, 'visual.json'), image);
  addedLogos++;
}

const executiveVisuals = path.join(pagesRoot, executiveId, 'visuals');
const unifiedFile = path.join(executiveVisuals, workforceCountryId, 'visual.json');
const country = readJson(unifiedFile);
if (country.visual?.visualType !== 'slicer') throw new Error('Expected Workforce country slicer');
const projection = country.visual.query.queryState.Values.projections[0];
const oldEntity = projection.field.Column.Expression.SourceRef.Entity;
if (!['User Inventory Evidence', 'Country Selection'].includes(oldEntity)) throw new Error(`Unexpected country source: ${oldEntity}`);
projection.field.Column.Expression.SourceRef.Entity = 'Country Selection';
projection.field.Column.Property = 'Country';
projection.queryRef = 'Country Selection.Country';
projection.nativeQueryRef = 'Country';
country.position.width = 352;
country.visual.objects.header[0].properties.text = literal("'Country'");
country.visual.visualContainerObjects.general[0].properties.altText = literal("'Filter workforce and devices by country.'");
writeJson(unifiedFile, country);

const oldDeviceCountryDir = path.join(executiveVisuals, oldDeviceCountryId);
if (fs.existsSync(oldDeviceCountryDir)) {
  if (!within(executiveVisuals, oldDeviceCountryDir)) throw new Error('Unsafe old slicer path');
  const files = fs.readdirSync(oldDeviceCountryDir);
  if (files.length !== 1 || files[0] !== 'visual.json') throw new Error('Old device country visual has unexpected files');
  const oldSlicer = readJson(path.join(oldDeviceCountryDir, 'visual.json'));
  if (oldSlicer.visual?.visualType !== 'slicer' ||
      oldSlicer.visual.query.queryState.Values.projections[0].queryRef !== 'Device Inventory Evidence.Country') {
    throw new Error('Old Device country slicer does not match the expected definition');
  }
  fs.rmSync(oldDeviceCountryDir, { recursive: true });
}

console.log(JSON.stringify({ pages: pageDirs.length, footersRightAligned: changedFooters, logosAdded: addedLogos, resourceName, unifiedCountrySlicer: workforceCountryId, removedDeviceCountrySlicer: oldDeviceCountryId }));
