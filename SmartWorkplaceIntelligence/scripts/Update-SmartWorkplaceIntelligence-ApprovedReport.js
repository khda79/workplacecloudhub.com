const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(process.argv[2] || path.join(__dirname, '..'));
const pagesRoot = path.join(projectRoot, 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');

const pages = {
  executive: '74096d393d05fa994d48', licensing: 'e0a71730533d84d5c995', workforce: '0443c39bb6eba584da07',
  devices: 'c8ce269962e9453ea038', deviceExplorer: 'a431090c3f25c4a369d0', messaging: '71a60a59e0eb52a7a7d0',
  mailbox: '1b19e6ba3cf47f2e8398', backup: '7e703070ffe685a52305'
};
const licensingFootnote = 'Reclaimable: disabled or no AD/M365 activity >90 days. E3 to F3: enabled E3, no Office PC use in 180 days, mailbox <2 GB, archive disabled. Inactive accounts may appear in both; prioritize reclaim. Validate business needs.';

function visualPath(page, id) { return path.join(pagesRoot, page, 'visuals', id, 'visual.json'); }
function readJson(file) { return JSON.parse(fs.readFileSync(file, 'utf8')); }
function writeJson(file, value) { fs.writeFileSync(file, JSON.stringify(value, null, 2) + '\n', 'utf8'); }
function walk(value, fn) {
  if (!value || typeof value !== 'object') return;
  for (const [key, child] of Object.entries(value)) { fn(value, key, child); walk(child, fn); }
}
function field(entity, property, kind = 'Measure', label = property, active = false) {
  const out = { field: { [kind]: { Expression: { SourceRef: { Entity: entity } }, Property: property } }, queryRef: `${entity}.${property}`, nativeQueryRef: label };
  if (active) out.active = true;
  return out;
}
function setCard(page, id, specs) {
  const file = visualPath(page, id), json = readJson(file);
  json.visual.query.queryState.Data.projections = specs.map(s => field(...s));
  writeJson(file, json);
}
function setTitle(json, title) {
  const objects = json.visual.visualContainerObjects ||= {};
  const titleObject = objects.title ||= [{ properties: {} }];
  titleObject[0].properties ||= {};
  titleObject[0].properties.text = { expr: { Literal: { Value: `'${title}'` } } };
}
function setMap(page, id) {
  const file = visualPath(page, id), json = readJson(file);
  json.visual.query.queryState = {
    Category: { projections: [field('User Inventory Evidence', 'Country', 'Column', 'Workforce Country', true)] },
    Size: { projections: [field('User Inventory Evidence', 'Enabled Workforce Accounts', 'Measure', 'Enabled Workforce Accounts', true)] },
    Tooltips: { projections: [
      field('User Inventory Evidence', 'Active Workforce', 'Measure', 'Active Workforce', true),
      field('User Inventory Evidence', 'Confirmed Human Users', 'Measure', 'Confirmed Human Users', true),
      field('User Inventory Evidence', 'Activity Evidence Coverage', 'Measure', 'Activity Evidence Coverage', true)
    ] }
  };
  setTitle(json, 'Workforce footprint by country');
  writeJson(file, json);
}
function setTableSort(page, id, entity, property) {
  const file = visualPath(page, id), json = readJson(file);
  json.visual.query.sortDefinition = { sort: [{ field: { Column: { Expression: { SourceRef: { Entity: entity } }, Property: property } }, direction: 'Descending' }] };
  writeJson(file, json);
}
function setText(page, id, text) {
  const file = visualPath(page, id), json = readJson(file);
  const run = json.visual.objects.general[0].properties.paragraphs[0].textRuns[0];
  run.value = text;
  run.textStyle ||= {};
  run.textStyle.fontSize = '10px';
  json.visual.objects.general[0].properties.paragraphs[0].horizontalTextAlignment = 'left';
  writeJson(file, json);
}
if (process.argv.includes('--licensing-footnote-only')) {
  setText(pages.licensing, '3c204b2481b9774eb5d5', licensingFootnote);
  console.log('Licensing footnote updated.');
  process.exit(0);
}
function renameProjection(page, id, queryRef, label) {
  const file = visualPath(page, id), json = readJson(file);
  const projections = json.visual?.query?.queryState?.Data?.projections || [];
  const projection = projections.find(p => p.queryRef === queryRef);
  if (!projection) throw new Error(`Projection not found: ${queryRef} in ${file}`);
  projection.nativeQueryRef = label;
  writeJson(file, json);
}
function addMailboxTypeChart(page, sourceId, newId) {
  const source = visualPath(page, sourceId), targetDir = path.dirname(visualPath(page, newId)), target = path.join(targetDir, 'visual.json');
  const json = readJson(source);
  json.name = newId;
  json.position.y = 464; json.position.height = 192; json.position.z += 40; json.position.tabOrder += 40;
  json.visual.query.queryState = {
    Category: { projections: [field('Mailbox Evidence', 'Mailbox Type', 'Column', 'Mailbox Type', true)] },
    Y: { projections: [field('Mailbox Evidence', 'Mailboxes', 'Measure', 'Mailboxes')] }
  };
  json.visual.query.sortDefinition = { sort: [{ field: { Measure: { Expression: { SourceRef: { Entity: 'Mailbox Evidence' } }, Property: 'Mailboxes' } }, direction: 'Descending' }] };
  setTitle(json, 'Mailboxes by type');
  fs.mkdirSync(targetDir, { recursive: true });
  writeJson(target, json);
  const original = readJson(source); original.position.height = 192; setTitle(original, 'Mailboxes by Exchange version'); writeJson(source, original);
}
function enableSearchAndFullNumbers() {
  const files = [];
  function collect(dir) { for (const e of fs.readdirSync(dir, { withFileTypes: true })) { const p = path.join(dir, e.name); e.isDirectory() ? collect(p) : e.name === 'visual.json' && files.push(p); } }
  collect(pagesRoot);
  const searchable = new Set(['User Principal Name','Primary User UPN','Device Name','Mailbox Key','Primary SMTP Address','Application Name','Normalized Product','Display Name']);
  for (const file of files) {
    const json = readJson(file); let changed = false, hasSearchField = false;
    walk(json, (parent, key, value) => {
      if (['labelDisplayUnits','displayUnits','dataLabelDisplayUnits'].includes(key)) {
        const literal = value?.expr?.Literal;
        if (literal && literal.Value !== '1D') { literal.Value = '1D'; changed = true; }
      }
      if (key === 'Property' && searchable.has(value)) hasSearchField = true;
      if (typeof value === 'string' && value.includes('Enabled Human Users')) { parent[key] = value.replaceAll('Enabled Human Users', 'Enabled Workforce Accounts'); changed = true; }
    });
    for (const axis of json.visual?.objects?.valueAxis || []) {
      axis.properties ||= {};
      const units = axis.properties.labelDisplayUnits ||= { expr: { Literal: { Value: '1D' } } };
      units.expr ||= {};
      units.expr.Literal ||= {};
      if (units.expr.Literal.Value !== '1D') {
        units.expr.Literal.Value = '1D';
      }
      changed = true;
    }
    if (json.visual?.visualType === 'slicer' && hasSearchField) {
      json.visual.objects ||= {};
      // Native dropdown slicers already expose search when opened; PBIR 2.12
      // does not persist a searchBox.show property.
      changed = true;
    }
    if (changed) writeJson(file, json);
  }
}

enableSearchAndFullNumbers();

setCard(pages.executive, '6b0b054d4011dc8f3732', [
  ['User Inventory Evidence','Active Workforce','Measure','Active Workforce'],
  ['User Inventory Evidence','Enabled Workforce Accounts','Measure','Enabled Workforce Accounts'],
  ['User Inventory Evidence','Confirmed Human Users','Measure','Confirmed Human Users']
]);
setCard(pages.executive, '6854dad122544a331a47', [
  ['Device Inventory Evidence','Managed Devices','Measure','Managed Devices'],
  ['Device Directory Summary','Enabled AD PCs Managed by Intune (%)','Measure','Enabled AD PCs Managed by Intune (%)'],
  ['Device Inventory Evidence','Device Compliance Rate','Measure','Device Compliance Rate'],
  ['Device Inventory Evidence','Current Endpoint Analytics Score','Measure','Endpoint Analytics Score']
]);
setCard(pages.executive, 'fc4687a2f85a9a7a7ab2', [
  ['License Evidence','Microsoft 365 Utilization (%)','Measure','M365 Utilization (%)'],
  ['License Evidence','F3 Reclaimable Assignments','Measure','F3 Reclaimable'],
  ['License Evidence','E3 Reclaimable Assignments','Measure','E3 Reclaimable'],
  ['License Evidence','E5 Reclaimable Assignments','Measure','E5 Reclaimable']
]);
setCard(pages.executive, '1f5377ed66dd2acf9e82', [
  ['Mailbox Evidence','Mailboxes','Measure','Mailboxes'],
  ['Mailbox Evidence','Observed Exchange Online Adoption (%)','Measure','Exchange Online Adoption'],
  ['Mailbox Evidence','Total Mailbox Storage (TB)','Measure','Mailbox Storage (TB)'],
  ['Content Storage Evidence','Messaging and Collaboration Storage (TB)','Measure','Messaging + Collaboration (TB)'],
  ['Content Storage Evidence','Messaging and Collaboration Storage per Workforce Account (GB)','Measure','Storage / Workforce (GB)']
]);
setMap(pages.executive, '77c1c3612450d2b500ab');

setCard(pages.licensing, '8d1389a2777fd48e426c', [
  ['License Evidence','Microsoft 365 Utilization (%)','Measure','M365 Utilization (%)'],
  ['License Evidence','F3 Reclaimable Assignments','Measure','F3 Reclaimable'],
  ['License Evidence','E3 Reclaimable Assignments','Measure','E3 Reclaimable'],
  ['License Evidence','E5 Reclaimable Assignments','Measure','E5 Reclaimable']
]);
setCard(pages.licensing, 'ae0d4b5e2f909ed97187', [
  ['License Evidence','Microsoft 365 Enabled Units','Measure','M365 Enabled Units'],
  ['License Evidence','Microsoft 365 Consumed Units','Measure','M365 Consumed Units'],
  ['License Evidence','Immediate Reclaimable Assignments','Measure','Immediate Reclaim'],
  ['License Evidence','Dormant 31-90D Review Assignments','Measure','Dormant 31-90D'],
  ['License Evidence','E3 to F3 Review Candidates','Measure','E3 to F3 Review']
]);
setTableSort(pages.licensing, 'b75b600343998374287c', 'License Evidence', 'Enabled Units');
setText(pages.licensing, '3c204b2481b9774eb5d5', licensingFootnote);

setCard(pages.workforce, 'ae41d658575702b0194a', [
  ['User Inventory Evidence','Enabled Workforce Accounts','Measure','Enabled Workforce Accounts'],
  ['User Inventory Evidence','Confirmed Human Users','Measure','Confirmed Human Users'],
  ['Identity Reconciliation Evidence','AD to Entra Coverage','Measure','AD to Entra Coverage'],
  ['Identity Reconciliation Evidence','Identity Conflicts','Measure','Classification Conflicts'],
  ['User Inventory Evidence','Missing Activity Evidence','Measure','Missing Activity Evidence'],
  ['User Inventory Evidence','Job Title Completeness','Measure','Job Title Completeness']
]);
setCard(pages.workforce, '87a4fa109d980eb55bc1', [
  ['User Inventory Evidence','Users With Missing Manager','Measure','Users Missing Manager'],
  ['User Inventory Evidence','Department Completeness','Measure','Department Completeness'],
  ['User Inventory Evidence','Job Title Completeness','Measure','Job Title Completeness']
]);

setCard(pages.devices, '98ef56cfde48bf2da6e2', [
  ['Device Directory Summary','Enabled AD Windows PCs','Measure','Enabled AD Windows PCs'],
  ['Device Directory Summary','Enabled AD PCs Managed by Intune (%)','Measure','Managed by Intune (%)'],
  ['Device Directory Summary','Enabled AD Windows PCs Unmanaged','Measure','Not Managed by Intune'],
  ['Device Directory Summary','Disabled AD Computer Objects','Measure','Disabled AD Computers'],
  ['Device Directory Summary','Disabled Entra Devices','Measure','Disabled Entra Devices'],
  ['Device Inventory Evidence','Devices Low on Disk','Measure','Devices Low on Disk']
]);
{
  const f=visualPath(pages.devices,'221b9a0eced24ec67e3f'), j=readJson(f); setTitle(j,'Managed devices by primary-user country'); writeJson(f,j);
}
setCard(pages.deviceExplorer, '04bfaf628fbcc4aa32ed', [
  ['Device Inventory Evidence','Managed Devices','Measure','Managed Devices'],
  ['Device Inventory Evidence','Device Compliance Rate','Measure','Device Compliance Rate'],
  ['Device Inventory Evidence','Noncompliant Devices','Measure','Noncompliant Devices']
]);
setCard(pages.deviceExplorer, 'e5eb78b1fef7b65e7c00', [
  ['Device Inventory Evidence','Intune Devices Not Synced 30+ Days','Measure','Intune Not Synced 30+ Days'],
  ['Device Inventory Evidence','Devices Low on Disk','Measure','Devices Low on Disk'],
  ['Device Directory Summary','Enabled AD PCs Managed by Intune (%)','Measure','Managed by Intune (%)'],
  ['Device Directory Summary','Disabled AD Computer Objects','Measure','Disabled AD Computers'],
  ['Device Directory Summary','Disabled Entra Devices','Measure','Disabled Entra Devices'],
  ['Device Inventory Evidence','Average Observed Tenure (Years)','Measure','Average Tenure (Years)']
]);

setCard(pages.messaging, 'f7a71d5ed220862a263b', [
  ['Mailbox Evidence','Mailboxes','Measure','Mailboxes'],
  ['Mailbox Evidence','Exchange Online Mailboxes','Measure','Exchange Online Mailboxes'],
  ['Mailbox Evidence','Total Mailbox Storage (TB)','Measure','Total Mailbox Storage (TB)']
]);
setCard(pages.messaging, '6eeaca8411d85c007b26', [
  ['Mailbox Evidence','On-premises Mailboxes','Measure','On-premises Mailboxes'],
  ['Mailbox Evidence','Observed Exchange Online Adoption (%)','Measure','Exchange Online Adoption'],
  ['Mailbox Evidence','Mailboxes with Archive','Measure','Mailboxes with Archive'],
  ['Mailbox Evidence','Mailboxes with Delegations','Measure','Mailboxes with Delegations'],
  ['Mailbox Evidence','Mailboxes With Extensive Delegation','Measure','Extensive Delegation'],
  ['Mailbox Evidence','Large Mailboxes Without Archive','Measure','Mailboxes >=45 GB Without Archive']
]);
addMailboxTypeChart(pages.messaging, 'c9b0677d10bcd794a6ec', 'a8b4e78f2b9c41d8a013');

setCard(pages.mailbox, '1aa6829ab71ed56b0754', [
  ['Mailbox Evidence','Mailboxes','Measure','Mailboxes'],
  ['Mailbox Evidence','Exchange Online Mailboxes','Measure','Exchange Online Mailboxes'],
  ['Mailbox Evidence','Large Mailboxes Without Archive','Measure','Mailboxes >=45 GB Without Archive']
]);
setCard(pages.mailbox, '9406d0701160f20466ea', [
  ['Mailbox Evidence','On-premises Mailboxes','Measure','On-premises Mailboxes'],
  ['Mailbox Evidence','Observed Exchange Online Adoption (%)','Measure','Exchange Online Adoption'],
  ['Mailbox Evidence','Mailboxes with Archive','Measure','Mailboxes with Archive'],
  ['Mailbox Evidence','Mailboxes with Delegations','Measure','Mailboxes with Delegations'],
  ['Mailbox Evidence','Total Mailbox Storage (TB)','Measure','Total Mailbox Storage (TB)'],
  ['Mailbox Evidence','Mailboxes With Extensive Delegation','Measure','Extensive Delegation']
]);
addMailboxTypeChart(pages.mailbox, 'e50ffcbab12026b6540c', 'bd7e91a6830d4b79a18e');

renameProjection(pages.backup, '1d3ee8c5c7240973b5b5', 'Backup Mailbox Evidence.Protected Mailboxes', 'All Protected Mailboxes (incl. outside scope)');

console.log('Approved report plan applied successfully.');
