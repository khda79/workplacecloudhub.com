// Apply the approved local PBIR changes. No Desktop refresh or publication is performed here.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

const report = path.resolve(import.meta.dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition');
const pages = path.join(report, 'pages');
const id = {
  executive: '74096d393d05fa994d48', licensing: 'e0a71730533d84d5c995',
  security: 'c823af8cf0f8d76eb252', workforce: '0443c39bb6eba584da07',
  user: '3a0fb4ff8eb05c8835b3', lifecycle: '0793507e95ae84a32c4a',
  personas: '9ae7c13f821d4a608e52',
};

function read(file) { return JSON.parse(fs.readFileSync(file, 'utf8')); }
function write(file, value) { fs.writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8'); }
function visualFile(page, visual) { return path.join(pages, page, 'visuals', visual, 'visual.json'); }
function edit(page, visual, fn) { const file = visualFile(page, visual); const v = read(file); fn(v); write(file, v); }
function measure(table, name, display = name) {
  return { field: { Measure: { Expression: { SourceRef: { Entity: table } }, Property: name } },
    queryRef: `${table}.${name}`, nativeQueryRef: display, displayName: display };
}
function column(table, name, display = name, active = undefined) {
  const p = { field: { Column: { Expression: { SourceRef: { Entity: table } }, Property: name } },
    queryRef: `${table}.${name}`, nativeQueryRef: display };
  if (active !== undefined) p.active = active;
  return p;
}
function setCard(v, measures) {
  if (v.visual.visualType !== 'cardVisual') throw new Error(`Not a card: ${v.name}`);
  v.visual.query.queryState.Data.projections = measures.map(m => measure(...m));
}
function title(v, value) {
  const entry = v.visual.visualContainerObjects?.title?.[0];
  if (!entry) throw new Error(`No title object: ${v.name}`);
  entry.properties.text.expr.Literal.Value = `'${value.replaceAll("'", "''")}'`;
}
function text(v, value) {
  v.visual.objects.general[0].properties.paragraphs[0].textRuns[0].value = value;
}
function setBar(v, field, label) {
  if (v.visual.visualType !== 'barChart') throw new Error(`Not a bar chart: ${v.name}`);
  v.visual.query.queryState.Category.projections = [column('User Inventory Evidence', field, field, true)];
  v.visual.query.queryState.Y.projections = [measure('User Inventory Evidence', 'Enabled Workforce Accounts', 'Users')];
  v.visual.query.sortDefinition = { sort: [{ field: measure('User Inventory Evidence', 'Enabled Workforce Accounts').field, direction: 'Descending' }], isDefaultSort: true };
  title(v, label);
  if (v.visual.objects?.valueAxis?.[0]?.properties) {
    v.visual.objects.valueAxis[0].properties.labelDisplayUnits = { expr: { Literal: { Value: '1D' } } };
    v.visual.objects.valueAxis[0].properties.labelPrecision = { expr: { Literal: { Value: '0L' } } };
  }
}
function setSlicer(v, field, label) {
  v.visual.query.queryState.Values.projections = [column('User Inventory Evidence', field, field)];
  v.visual.objects.header[0].properties.text.expr.Literal.Value = `'${label}'`;
  delete v.visual.objects.general;
}

function updatePersonaStatus() {
  // A copied pivot retained a small pre-refresh sample in Desktop even though
  // its measure evaluated against the full workforce.
  const statusVisualId = copied('928e023a66d403298322');
  const statusPosition = read(visualFile(id.personas, statusVisualId)).position;
  const statusChart = read(visualFile(id.personas, copied('93cfd1ba5a2539e431e3')));
  statusChart.name = statusVisualId;
  statusChart.position = { ...statusPosition };
  setBar(statusChart, 'Persona Match State', 'Classification status');
  write(visualFile(id.personas, statusVisualId), statusChart);
}

function updatePersonaInsights() {
  const matrixTemplate = read(visualFile(id.user, '928e023a66d403298322'));
  const licenseId = copied('a65c81cadf407124bcf7');
  const licenseMatrix = structuredClone(matrixTemplate);
  licenseMatrix.name = licenseId;
  licenseMatrix.position = { x: 32, y: 256, width: 1424, height: 232, z: 180, tabOrder: 180 };
  licenseMatrix.visual.query.queryState.Rows.projections = [column('User Inventory Evidence', 'Workforce Persona', 'Persona', true)];
  licenseMatrix.visual.query.queryState.Values.projections = [
    measure('User Inventory Evidence', 'M365 Assignments Without Observed Use (30D)', 'No use · review'),
    measure('User Inventory Evidence', 'E3 to F3 Review Candidates by Persona', 'E3 to F3 review'),
    measure('User Inventory Evidence', 'M365 F3 Assigned', 'F3 assigned'),
    measure('User Inventory Evidence', 'M365 F3 Persona Share of All Assigned (%)', 'F3 share %'),
    measure('User Inventory Evidence', 'M365 E3 Assigned', 'E3 assigned'),
    measure('User Inventory Evidence', 'M365 E3 Persona Share of All Assigned (%)', 'E3 share %'),
    measure('User Inventory Evidence', 'M365 E5 Assigned', 'E5 assigned'),
    measure('User Inventory Evidence', 'M365 E5 Persona Share of All Assigned (%)', 'E5 share %'),
  ];
  licenseMatrix.visual.objects.columnHeaders[0].properties.columnAdjustment.expr.Literal.Value = "'growToFit'";
  licenseMatrix.visual.objects.columnHeaders[0].properties.autoSizeColumnWidth.expr.Literal.Value = 'true';
  title(licenseMatrix, 'Microsoft 365 F3, E3 and E5 assignments by persona');
  if (licenseMatrix.visual.visualContainerObjects?.general?.[0]?.properties?.altText)
    licenseMatrix.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
      "'No use review counts F3, E3 and E5 assignments without observed Exchange, Teams, SharePoint or OneDrive activity and with explicit no Office Windows/Mac app usage over 30 days. Missing Apps usage rows are excluded. E3 to F3 also requires explicit no desktop app usage over 180 days; no E5 downgrade is proposed.'";
  write(visualFile(id.personas, licenseId), licenseMatrix);

  const serviceId = '887a323f886143298c9e';
  const serviceMatrix = structuredClone(matrixTemplate);
  serviceMatrix.name = serviceId;
  serviceMatrix.position = { x: 32, y: 504, width: 1424, height: 232, z: 181, tabOrder: 181 };
  serviceMatrix.visual.query.queryState.Rows.projections = [column('User Inventory Evidence', 'Workforce Persona', 'Persona', true)];
  serviceMatrix.visual.query.queryState.Values.projections = [
    measure('User Inventory Evidence', 'Any M365 Service Usage (30D %)', 'Any service %'),
    measure('User Inventory Evidence', 'Office Desktop Usage (30D %)', 'Office PC use %'),
    measure('User Inventory Evidence', 'Primary PC Users', 'Primary PC users'),
    measure('User Inventory Evidence', 'Primary PC Users (%)', 'Primary PC %'),
    ...[['Exchange', 'Exchange'], ['Teams', 'Teams'], ['SharePoint', 'SharePoint'], ['OneDrive', 'OneDrive']].map(([service, label]) =>
      measure('User Inventory Evidence', `${service} Usage (30D %)`, `${label} %`)),
  ];
  serviceMatrix.visual.objects.columnHeaders[0].properties.columnAdjustment.expr.Literal.Value = "'growToFit'";
  serviceMatrix.visual.objects.columnHeaders[0].properties.autoSizeColumnWidth.expr.Literal.Value = 'true';
  title(serviceMatrix, 'Service use by persona · last 30 days');
  if (serviceMatrix.visual.visualContainerObjects?.general?.[0]?.properties?.altText)
    serviceMatrix.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
      "'Thirty-day activity among covered accounts. Office PC use is intentional Outlook, Word, Excel, PowerPoint or OneNote activity on Windows or Mac. Missing Apps usage evidence is excluded. Primary PC is an observed Intune association.'";
  fs.mkdirSync(path.dirname(visualFile(id.personas, serviceId)), { recursive: true });
  write(visualFile(id.personas, serviceId), serviceMatrix);

  edit(id.personas, copied('c1a62ad26686472ff656'), v => text(v,
    'Workforce Personas · F3/E3/E5 share: of all assigned · no use: review only · Office: PC app use · 30-day covered accounts'));
  edit(id.personas, copied('62b771092d7f62b037b0'), v => {
    const p = v.visual.query.queryState.Data.projections;
    if (!p.some(x => x.field.Measure.Property === 'Workforce Accounts Without M365 Usage Report')) {
      p.push(measure('User Inventory Evidence', 'Workforce Accounts Without M365 Usage Report', 'Usage evidence missing'));
    }
  });
  // User-level detail is available on User Explorer; the two persona matrices
  // occupy its former panel. Keep the classification signals below.
}

function updatePersonaLayout() {
  const place = (visualId, x, y, width, height) => edit(id.personas, visualId, v => {
    Object.assign(v.position, { x, y, width, height });
  });
  place('503efc5a5818c64bd5a6', 1480, 256, 408, 192); // Classification status
  place('70d3a4c47fd52ed4c7d4', 1480, 464, 408, 192); // Users by persona
  place('b3128e872840587b94cf', 32, 256, 1424, 232); // License assignments
  place('887a323f886143298c9e', 32, 504, 1424, 232); // Service use
  place('d3d4837563c0ef1ac2bb', 32, 760, 1424, 248); // Classification evidence gaps
  place('294a3a0c0feac5ffadcd', 1848, 268, 24, 24); // Classification icon
  place('8a089199f26a88ef9a82', 1848, 476, 24, 24); // Persona icon
}

function updatePersonaPrimaryPc() {
  // Keep the user picker on the same workforce evidence table as the persona
  // matrices, so one UPN selection filters all measures consistently.
  const userSlicerId = crypto.createHash('sha1').update('workforce-personas-user-filter').digest('hex').slice(0, 20);
  const userSlicer = read(visualFile(id.user, 'ce8fff0f0d93da4249c8'));
  userSlicer.name = userSlicerId;
  userSlicer.position = { x: 1024, y: 8, z: 105, width: 192, height: 80, tabOrder: 105 };
  setSlicer(userSlicer, 'User Principal Name', 'User');
  userSlicer.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
    "'Search and filter by user principal name.'";
  fs.mkdirSync(path.dirname(visualFile(id.personas, userSlicerId)), { recursive: true });
  write(visualFile(id.personas, userSlicerId), userSlicer);

  for (const [visualId, x, width] of [
    ['e8b3a6398343fd57e94f', 1232, 144], // Persona
    ['2cddfae8432a26339970', 1392, 144], // Job title
    ['5577d3f3e4e1ab0de03c', 1552, 144], // Country
    ['4359dd81c989f0a6787f', 1712, 176], // Evidence & freshness
  ]) edit(id.personas, visualId, v => Object.assign(v.position, { x, width }));
  for (const visualId of ['0607d15f5c3e3b0eac4a', '9dead798b5367a10653f']) {
    edit(id.personas, visualId, v => { v.position.width = 912; });
  }

  edit(id.personas, '887a323f886143298c9e', v => {
    const projections = v.visual.query.queryState.Values.projections;
    const primaryMeasures = [
      ['Any M365 Service Usage (30D %)', 'Any service %'],
      ['Office Desktop Usage (30D %)', 'Office PC use %'],
      ['Primary PC Users', 'Primary PC users'],
      ['Primary PC Users (%)', 'Primary PC %'],
    ];
    v.visual.query.queryState.Values.projections = [
      ...primaryMeasures.map(([name, label]) =>
        projections.find(p => p.field?.Measure?.Property === name) || measure('User Inventory Evidence', name, label)),
      ...projections.filter(p => !primaryMeasures.some(([name]) => p.field?.Measure?.Property === name)),
    ];
    title(v, 'Service use by persona · last 30 days');
    if (v.visual.visualContainerObjects?.general?.[0]?.properties?.altText) {
      v.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
        "'Activity over 30 days among covered accounts. Office PC use reflects intentional Outlook, Word, Excel, PowerPoint or OneNote activity on Windows or Mac. Missing Apps usage evidence is excluded. Primary PC is an Intune association.'";
    }
  });
}

if (process.argv.includes('--persona-primary-pc-only')) {
  updatePersonaPrimaryPc();
  process.stdout.write('Workforce Personas user filter and primary PC measures updated locally.\n');
  process.exit(0);
}

if (process.argv.includes('--persona-layout-only')) {
  updatePersonaLayout();
  process.stdout.write('Workforce Personas layout updated locally.\n');
  process.exit(0);
}

if (process.argv.includes('--persona-status-only')) {
  updatePersonaStatus();
  process.stdout.write('Workforce Personas classification visual updated locally.\n');
  process.exit(0);
}

if (process.argv.includes('--persona-insights-only')) {
  updatePersonaInsights();
  process.stdout.write('Workforce Personas license and usage matrices updated locally.\n');
  process.exit(0);
}

if (process.argv.includes('--storage-labels-only')) {
  edit(id.executive, '95b252a8ee982c487d3c', v => {
    v.position.y = 384;
    v.position.height = 88;
    title(v, 'Messaging & collaboration storage');
    v.visual.objects.label[0].properties.show.expr.Literal.Value = 'true';
    v.visual.objects.label[0].properties.fontSize.expr.Literal.Value = '10D';
  });
  process.stdout.write('Executive storage labels updated locally.\n');
  process.exit(0);
}

// Rename old capability wording in all existing visual bindings. Registration and CA enforcement stay distinct.
for (const page of fs.readdirSync(pages, { withFileTypes: true }).filter(x => x.isDirectory())) {
  const folder = path.join(pages, page.name, 'visuals');
  if (!fs.existsSync(folder)) continue;
  for (const child of fs.readdirSync(folder, { withFileTypes: true }).filter(x => x.isDirectory())) {
    const file = path.join(folder, child.name, 'visual.json');
    if (!fs.existsSync(file)) continue;
    const old = fs.readFileSync(file, 'utf8');
    const next = old.replaceAll('MFA Capability (%)', 'Workforce MFA Registered (%)')
      .replaceAll('Users Not MFA Capable (#)', 'Workforce Without MFA Registration (#)');
    if (next !== old) fs.writeFileSync(file, next, 'utf8');
  }
}

// Executive: keep separate incompatible and undetermined Windows 10 states.
edit(id.executive, 'e4d06fe45cb5c0cf9efe', v => {
  const values = v.visual.query.queryState.Data.projections;
  values.splice(3, values.length - 3, measure('Device Directory Summary', 'Windows 10 Eligibility Undetermined', 'Win10 eligibility unknown'));
  v.visual.objects.value[0].properties.fontSize.expr.Literal.Value = '15D';
});
edit(id.executive, '1f5377ed66dd2acf9e82', v => {
  v.visual.query.queryState.Data.projections = v.visual.query.queryState.Data.projections.slice(0, 2);
});
edit(id.executive, '95b252a8ee982c487d3c', v => {
  setCard(v, [
    ['Content Storage Evidence', 'Messaging and Collaboration Storage (TB)', 'Messaging + collaboration (TB)'],
    ['Content Storage Evidence', 'Messaging and Collaboration Storage per Workforce Account (GB)', 'Storage / workforce (GB)'],
    ['Mailbox Evidence', 'Total Mailbox Storage (TB)', 'Mailbox storage (TB)'],
  ]);
  v.position.y = 384; v.position.height = 88; v.position.width = 744;
  title(v, 'Messaging & collaboration storage');
  v.visual.objects.value[0].properties.fontSize.expr.Literal.Value = '15D';
  v.visual.objects.label[0].properties.show.expr.Literal.Value = 'true';
  v.visual.objects.label[0].properties.fontSize.expr.Literal.Value = '10D';
});
edit(id.executive, '0385ca0d1c75d8020694', v => {
  setCard(v, [['License Evidence', 'E3 to F3 Review Candidates', 'E3 → F3 candidates']]);
  title(v, 'E3 to F3 review candidates');
  v.visual.objects.value[0].properties.fontSize.expr.Literal.Value = '18D';
});
edit(id.executive, '12365a8234c915ca9682', v => {
  setCard(v, [['Device Inventory Evidence', 'Devices Encrypted According to Intune', 'Encrypted devices']]);
  title(v, 'Encrypted devices (Intune)');
  v.visual.objects.value[0].properties.fontSize.expr.Literal.Value = '18D';
});

// Financial review candidates move into the primary licensing KPIs.
edit(id.licensing, '8d1389a2777fd48e426c', v => {
  const projections = v.visual.query.queryState.Data.projections;
  if (!projections.some(p => p.field.Measure.Property === 'E3 to F3 Review Candidates')) {
    projections.push(measure('License Evidence', 'E3 to F3 Review Candidates', 'E3 → F3 review candidates'));
  }
});
edit(id.licensing, 'ae0d4b5e2f909ed97187', v => {
  v.visual.query.queryState.Data.projections = v.visual.query.queryState.Data.projections.filter(p => p.field.Measure.Property !== 'E3 to F3 Review Candidates');
});

// Security: registration is measured among covered workforce accounts; show actual Intune encryption count.
edit(id.security, '90dd02b57d7360645ca4', v => {
  const p = v.visual.query.queryState.Data.projections;
  const n = p.findIndex(x => x.field.Measure.Property === 'Disk Encryption Healthy (%)');
  if (n >= 0) p[n] = measure('Device Inventory Evidence', 'Devices Encrypted According to Intune', 'Encrypted (Intune)');
  else if (!p.some(x => x.field.Measure.Property === 'Devices Encrypted According to Intune')) throw new Error('Security supporting KPI not found');
});
edit(id.security, 'fd6fd2b2ac737b57de06', v => {
  const p = v.visual.query.queryState.Data.projections;
  if (!p.some(x => x.field.Measure.Property === 'Workforce MFA Registration Coverage (%)')) {
    p.push(measure('User Inventory Evidence', 'Workforce MFA Registration Coverage (%)', 'MFA evidence coverage'));
  }
});

// Workforce: the broader enabled workforce belongs in primary KPIs.
edit(id.workforce, 'cbe2399603bb2121bd65', v => {
  const p = v.visual.query.queryState.Data.projections;
  if (!p.some(x => x.field.Measure.Property === 'Enabled Workforce Accounts')) p.unshift(measure('User Inventory Evidence', 'Enabled Workforce Accounts'));
});
edit(id.workforce, 'ae41d658575702b0194a', v => {
  v.visual.query.queryState.Data.projections = v.visual.query.queryState.Data.projections.filter(p => p.field.Measure.Property !== 'Enabled Workforce Accounts');
});

// Windows Lifecycle uses the joined AD/Intune unknown population as a separate KPI.
edit(id.lifecycle, '19cc17a69e73ee1c293d', v => {
  const p = v.visual.query.queryState.Data.projections;
  if (!p.some(x => x.field.Measure.Property === 'Windows 10 Eligibility Undetermined')) {
    p.push(measure('Device Directory Summary', 'Windows 10 Eligibility Undetermined', 'Win10 eligibility unknown'));
  }
});

// User Explorer: two compact sorted views, one by persona and one by job title.
edit(id.user, '93cfd1ba5a2539e431e3', v => {
  setBar(v, 'Workforce Persona', 'Users by persona');
  v.position.height = 192;
});
const jobChartId = '0b327acf93e04d91865a';
const jobChartFile = visualFile(id.user, jobChartId);
if (!fs.existsSync(jobChartFile)) {
  const chart = read(visualFile(id.user, '93cfd1ba5a2539e431e3'));
  chart.name = jobChartId;
  chart.position = { ...chart.position, x: 1480, y: 464, height: 192, tabOrder: 250, z: 250 };
  setBar(chart, 'Job Title', 'Users by job title');
  fs.mkdirSync(path.dirname(jobChartFile), { recursive: true });
  write(jobChartFile, chart);
}

// New Workforce Personas page, placed immediately before Workforce & Identity.
const sourcePage = path.join(pages, id.user);
const targetPage = path.join(pages, id.personas);
if (!fs.existsSync(targetPage)) {
  fs.mkdirSync(path.join(targetPage, 'visuals'), { recursive: true });
  const metadata = read(path.join(sourcePage, 'page.json'));
  metadata.name = id.personas;
  metadata.displayName = 'Workforce Personas';
  delete metadata.filterConfig;
  delete metadata.pageBinding;
  write(path.join(targetPage, 'page.json'), metadata);
  const omit = new Set([
    'd7b514052000b9e26eb2', 'b6930250059a16161dad', 'be12c9f82530345feeaf',
    '81d79d551eac12698e9f', '2e1e3494e2b4e1c35456', 'f5eb42f34008c4a584fc',
    'bcbe3916207958d30bfa', jobChartId,
  ]);
  for (const child of fs.readdirSync(path.join(sourcePage, 'visuals'), { withFileTypes: true }).filter(x => x.isDirectory())) {
    if (omit.has(child.name)) continue;
    const file = path.join(sourcePage, 'visuals', child.name, 'visual.json');
    if (!fs.existsSync(file)) continue;
    const v = read(file);
    const newId = crypto.createHash('sha1').update(`${id.personas}:${child.name}`).digest('hex').slice(0, 20);
    v.name = newId;
    fs.mkdirSync(path.join(targetPage, 'visuals', newId), { recursive: true });
    write(visualFile(id.personas, newId), v);
  }
}
function copied(originalId) { return crypto.createHash('sha1').update(`${id.personas}:${originalId}`).digest('hex').slice(0, 20); }
edit(id.personas, copied('ab6f6be841ab2f9ae2cf'), v => text(v, 'How is the enabled workforce classified by persona and job title?'));
edit(id.personas, copied('c1a62ad26686472ff656'), v => text(v, 'Workforce Personas · governed classification rules · unmatched and ambiguous titles shown separately'));
edit(id.personas, copied('ce8fff0f0d93da4249c8'), v => setSlicer(v, 'Workforce Persona', 'Persona'));
edit(id.personas, copied('a01af72d75320d14ca69'), v => setSlicer(v, 'Job Title', 'Job title'));
edit(id.personas, copied('51dc75f2eb71a770500f'), v => setSlicer(v, 'Country', 'Country'));
edit(id.personas, copied('81b6e6123ee562415f83'), v => setCard(v, [
  ['User Inventory Evidence', 'Enabled Workforce Accounts'],
  ['User Inventory Evidence', 'Classified Workforce Accounts'],
  ['User Inventory Evidence', 'Persona Classification Coverage (%)'],
]));
edit(id.personas, copied('62b771092d7f62b037b0'), v => setCard(v, [
  ['User Inventory Evidence', 'Unclassified Workforce Accounts'],
  ['User Inventory Evidence', 'Workforce Accounts Requiring Persona Review'],
  ['User Inventory Evidence', 'Job Title Completeness'],
]));
updatePersonaStatus();
edit(id.personas, copied('a65c81cadf407124bcf7'), v => {
  v.visual.query.queryState.Values.projections = [
    column('User Inventory Evidence', 'User Principal Name', 'User'),
    column('User Inventory Evidence', 'Display Name'),
    column('User Inventory Evidence', 'Job Title'),
    column('User Inventory Evidence', 'Workforce Persona', 'Persona'),
    column('User Inventory Evidence', 'Persona Match State', 'Match state'),
    column('User Inventory Evidence', 'Persona Matched Keywords', 'Matched keyword(s)'),
    column('User Inventory Evidence', 'Department'),
  ];
  title(v, 'Workforce persona detail');
});
edit(id.personas, copied('93cfd1ba5a2539e431e3'), v => {
  setBar(v, 'Workforce Persona', 'Users by persona');
  v.position.height = 400;
});
edit(id.personas, copied('ea363c448e1db86764bb'), v => title(v, 'Identity and classification evidence gaps'));
const personaJobChart = read(visualFile(id.personas, copied('93cfd1ba5a2539e431e3')));
personaJobChart.name = 'd91fb0e36ab74c2598e1';
personaJobChart.position = { ...personaJobChart.position, x: 1480, y: 680, width: 408, height: 328, tabOrder: 251, z: 251 };
setBar(personaJobChart, 'Job Title', 'Users by job title');
fs.mkdirSync(path.dirname(visualFile(id.personas, personaJobChart.name)), { recursive: true });
write(visualFile(id.personas, personaJobChart.name), personaJobChart);

const metadata = path.join(pages, 'pages.json');
const order = read(metadata);
if (!order.pageOrder.includes(id.personas)) {
  const index = order.pageOrder.indexOf(id.workforce);
  if (index < 0) throw new Error('Workforce page missing from page order');
  order.pageOrder.splice(index, 0, id.personas);
  write(metadata, order);
}

updatePersonaInsights();
updatePersonaLayout();
updatePersonaPrimaryPc();
process.stdout.write('PBIR report edits applied locally; Desktop refresh and publication not performed.\n');
