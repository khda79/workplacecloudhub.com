const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(process.argv[2] || path.join(__dirname, '..'));
const definitionRoot = path.join(projectRoot, 'pbip', 'SmartWorkplaceIntelligence.SemanticModel', 'definition');
const tableRoot = path.join(definitionRoot, 'tables');

function read(file) {
  return fs.readFileSync(file, 'utf8');
}

function write(file, text) {
  fs.writeFileSync(file, text, 'utf8');
}

function replaceOnce(text, from, to, label) {
  if (text.includes(to)) return text;
  const index = text.indexOf(from);
  if (index < 0) throw new Error(`Anchor not found for ${label}`);
  return text.slice(0, index) + to + text.slice(index + from.length);
}

function insertBefore(text, anchor, block, marker) {
  if (text.includes(marker)) return text;
  const index = text.indexOf(anchor);
  if (index < 0) throw new Error(`Insert anchor not found for ${marker}`);
  return text.slice(0, index) + block + text.slice(index);
}

function updateTable(name, updater) {
  const file = path.join(tableRoot, `${name}.tmdl`);
  const original = read(file);
  const eol = original.includes('\r\n') ? '\r\n' : '\n';
  const normalized = original.replace(/\r\n/g, '\n');
  const updated = updater(normalized);
  if (updated !== normalized) write(file, updated.replace(/\n/g, eol));
  return updated !== normalized;
}

function updateUserInventory(text) {
  const replacements = [
    ["measure 'Enabled Human Users' = DISTINCTCOUNT ( 'User Inventory Evidence'[User Source ID] )", "measure 'Enabled Workforce Accounts' = DISTINCTCOUNT ( 'User Inventory Evidence'[User Source ID] )", 'workforce measure'],
    ['Columns=26', 'Columns=31', 'user CSV width'],
    ['This evidence table contains only confirmed Human rows.', 'Rows include governed Human accounts and AD-linked review classifications; explicit non-human and cloud-only accounts are excluded.', 'population documentation'],
  ];
  for (const [from, to, label] of replacements) text = replaceOnce(text, from, to, label);
  text = text.replaceAll('[Enabled Human Users]', '[Enabled Workforce Accounts]');

  text = insertBefore(text, "\tcolumn 'User Source ID'", `\t/// Accounts still classified explicitly as Human by the governed AD rules.\n\tmeasure 'Confirmed Human Users' = CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Account Population] = "Human" )\n\t\tformatString: #,0\n\t\tdisplayFolder: Workforce | Coverage\n\n\tmeasure 'Job Title Completeness' = DIVIDE ( CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Job Title Evidence State] = "Observed" ), [Enabled Workforce Accounts] )\n\t\tformatString: 0.0%\n\t\tdisplayFolder: Workforce | Coverage\n\n`, "measure 'Confirmed Human Users'");

  text = insertBefore(text, "\tcolumn Manager", `\tcolumn 'Job Title'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Job Title\n\n\tcolumn 'Job Title Evidence State'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Job Title Evidence State\n\n`, "column 'Job Title'");

  text = insertBefore(text, "\tcolumn 'Last Activity Workload'", `\tcolumn 'AD Last Activity Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: AD Last Activity Date\n\n\tcolumn 'M365 Last Activity Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: M365 Last Activity Date\n\n\tcolumn 'Activity Source'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Activity Source\n\n`, "column 'AD Last Activity Date'");

  text = replaceOnce(text,
    `{"Department Evidence State", type text},\n\t\t\t\t        {"Manager", type text},`,
    `{"Department Evidence State", type text},\n\t\t\t\t        {"Job Title", type text},\n\t\t\t\t        {"Job Title Evidence State", type text},\n\t\t\t\t        {"Manager", type text},`,
    'user job-title query types');
  text = replaceOnce(text,
    `{"Last Activity Date", type date},\n\t\t\t\t        {"Last Activity Workload", type text},`,
    `{"Last Activity Date", type date},\n\t\t\t\t        {"AD Last Activity Date", type date},\n\t\t\t\t        {"M365 Last Activity Date", type date},\n\t\t\t\t        {"Activity Source", type text},\n\t\t\t\t        {"Last Activity Workload", type text},`,
    'user activity query types');
  return text;
}

function updateUserHistory(text) {
  text = text.replaceAll('Historical Enabled Human Users', 'Historical Enabled Workforce Accounts');
  text = text.replaceAll('Days Since Last M365 Activity', 'Days Since Last Workforce Activity');
  text = replaceOnce(text, 'Columns=8', 'Columns=11', 'history CSV width');
  text = insertBefore(text, "\tpartition 'User Activity History Evidence Import'", `\tcolumn 'AD Last Activity Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: AD Last Activity Date\n\n\tcolumn 'M365 Last Activity Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: M365 Last Activity Date\n\n\tcolumn 'Activity Source'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Activity Source\n\n`, "column 'AD Last Activity Date'");
  text = replaceOnce(text,
    `{"Last Activity Date", type date}\n`,
    `{"Last Activity Date", type date},\n\t\t\t\t        {"AD Last Activity Date", type date},\n\t\t\t\t        {"M365 Last Activity Date", type date},\n\t\t\t\t        {"Activity Source", type text}\n`,
    'history activity query types');
  return text;
}

function updateWorkforceTrend(text) {
  text = text.replaceAll('Trend Enabled Human Users', 'Trend Enabled Workforce Accounts');
  text = text.replaceAll("[Enabled Human Users]", "[Enabled Workforce Accounts]");
  text = text.replaceAll("column 'Enabled Human Users'", "column 'Enabled Workforce Accounts'");
  text = text.replaceAll('sourceColumn: Enabled Human Users', 'sourceColumn: Enabled Workforce Accounts');
  text = replaceOnce(text, 'Columns=9', 'Columns=10', 'workforce trend CSV width');
  text = insertBefore(text, "\tcolumn 'Activity Covered Users'", `\tcolumn 'Confirmed Human Users'\n\t\tdataType: int64\n\t\tformatString: #,0\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Confirmed Human Users\n\n`, "column 'Confirmed Human Users'");
  text = insertBefore(text, "\tcolumn 'Snapshot Date'", `\tmeasure 'Trend Confirmed Human Users' = SUM ( 'Workforce Trend Evidence'[Confirmed Human Users] )\n\t\tformatString: #,0\n\t\tdisplayFolder: Workforce | Trend\n\n`, "measure 'Trend Confirmed Human Users'");
  text = replaceOnce(text,
    `{"Enabled Human Users", Int64.Type},\n`,
    `{"Enabled Workforce Accounts", Int64.Type},\n\t\t\t\t        {"Confirmed Human Users", Int64.Type},\n`,
    'workforce trend query types');
  return text;
}

function updateLicenseEvidence(text) {
  text = insertBefore(text, "\tcolumn 'Capacity Class'", `\tcolumn 'Microsoft 365 Plan'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Microsoft 365 Plan\n\n`, "column 'Microsoft 365 Plan'");
  text = insertBefore(text, "\tcolumn 'Dormant 31-90D Assignments'", `\tcolumn 'Disabled Reclaimable Assignments'\n\t\tdataType: int64\n\t\tformatString: #,0\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Disabled Reclaimable Assignments\n\n`, "column 'Disabled Reclaimable Assignments'");
  text = insertBefore(text, "\tcolumn 'Risk Priority'", `\tcolumn 'E3 to F3 Review Candidate Assignments'\n\t\tdataType: int64\n\t\tformatString: #,0\n\t\tsummarizeBy: sum\n\t\tsourceColumn: E3 to F3 Review Candidates\n\n`, "column 'E3 to F3 Review Candidate Assignments'");
  text = insertBefore(text, "\tcolumn 'License Product'", `\tmeasure 'Microsoft 365 Enabled Units' = CALCULATE ( SUM ( 'License Evidence'[Enabled Units] ), 'License Evidence'[Microsoft 365 Plan] IN { "F3", "E3", "E5" } )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Microsoft 365\n\n\tmeasure 'Microsoft 365 Consumed Units' = CALCULATE ( SUM ( 'License Evidence'[Consumed Units] ), 'License Evidence'[Microsoft 365 Plan] IN { "F3", "E3", "E5" } )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Microsoft 365\n\n\tmeasure 'Microsoft 365 Utilization (%)' = DIVIDE ( [Microsoft 365 Consumed Units], [Microsoft 365 Enabled Units] )\n\t\tformatString: 0.0%\n\t\tdisplayFolder: Licensing | Microsoft 365\n\n\tmeasure 'F3 Reclaimable Assignments' = CALCULATE ( SUM ( 'License Evidence'[Potentially Reclaimable Assignments] ), 'License Evidence'[Microsoft 365 Plan] = "F3" )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Microsoft 365\n\n\tmeasure 'E3 Reclaimable Assignments' = CALCULATE ( SUM ( 'License Evidence'[Potentially Reclaimable Assignments] ), 'License Evidence'[Microsoft 365 Plan] = "E3" )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Microsoft 365\n\n\tmeasure 'E5 Reclaimable Assignments' = CALCULATE ( SUM ( 'License Evidence'[Potentially Reclaimable Assignments] ), 'License Evidence'[Microsoft 365 Plan] = "E5" )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Microsoft 365\n\n\tmeasure 'E3 to F3 Review Candidates' = SUM ( 'License Evidence'[E3 to F3 Review Candidate Assignments] )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Optimization\n\n`, "measure 'Microsoft 365 Enabled Units'");
  text = insertBefore(text, "\tcolumn 'License Product'", `\tmeasure 'Immediate Reclaimable Assignments' = SUM ( 'License Evidence'[Potentially Reclaimable Assignments] )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Optimization\n\n\tmeasure 'Dormant 31-90D Review Assignments' = SUM ( 'License Evidence'[Dormant 31-90D Assignments] )\n\t\tformatString: #,0\n\t\tdisplayFolder: Licensing | Optimization\n\n`, "measure 'Immediate Reclaimable Assignments'");
  text = replaceOnce(text,
    `{"License Product", type text}, {"SKU Part Number", type text}, {"Capacity Class", type text},`,
    `{"License Product", type text}, {"SKU Part Number", type text}, {"Microsoft 365 Plan", type text}, {"Capacity Class", type text},`,
    'license plan query type');
  text = replaceOnce(text,
    `{"Potentially Reclaimable Assignments", Int64.Type}, {"Dormant 31-90D Assignments", Int64.Type},`,
    `{"Potentially Reclaimable Assignments", Int64.Type}, {"Disabled Reclaimable Assignments", Int64.Type}, {"Dormant 31-90D Assignments", Int64.Type}, {"E3 to F3 Review Candidates", Int64.Type},`,
    'license optimization query types');
  for (const plan of ['F3', 'E3', 'E5']) {
    const plain = `measure '${plan} Reclaimable Assignments' = CALCULATE ( SUM ( 'License Evidence'[Potentially Reclaimable Assignments] ), 'License Evidence'[Microsoft 365 Plan] = "${plan}" )`;
    const coalesced = `measure '${plan} Reclaimable Assignments' = COALESCE ( CALCULATE ( SUM ( 'License Evidence'[Potentially Reclaimable Assignments] ), 'License Evidence'[Microsoft 365 Plan] = "${plan}" ), 0 )`;
    text = text.replace(plain, coalesced);
  }
  return text;
}

function updateDeviceEvidence(text) {
  text = insertBefore(text, "\tcolumn 'Enrollment DateTime'", `\tcolumn 'Entra Account State'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Entra Account State\n\n`, "column 'Entra Account State'");
  text = insertBefore(text, "\tcolumn 'Device Source ID'", `\tmeasure 'Intune Devices Not Synced 30+ Days' = [Devices Not Synced 30+ Days]\n\t\tformatString: #,0\n\t\tdisplayFolder: Device | Risk\n\n\tmeasure 'Disabled Entra Managed Devices' = CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Entra Account State] = "Disabled" )\n\t\tformatString: #,0\n\t\tdisplayFolder: Device | Risk\n\n`, "measure 'Intune Devices Not Synced 30+ Days'");
  text = replaceOnce(text, 'Columns = 52', 'Columns = 53', 'device CSV width');
  text = replaceOnce(text,
    `{"Entra Match", type text},\n`,
    `{"Entra Match", type text},\n\t\t\t\t        {"Entra Account State", type text},\n`,
    'device Entra state query type');
  return text;
}

function updateMailboxEvidence(text) {
  text = insertBefore(text, "\tcolumn 'Delegation Count'", `\tcolumn 'Total Storage GB'\n\t\tdataType: double\n\t\tformatString: #,0.0\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Total Storage GB\n\n`, "column 'Total Storage GB'");
  text = insertBefore(text, "\tcolumn 'Mailbox Key'", `\t/// Primary mailbox plus archive storage.\n\tmeasure 'Total Mailbox Storage (TB)' = DIVIDE ( SUM ( 'Mailbox Evidence'[Total Storage GB] ), 1024 )\n\t\tformatString: #,0.0\n\t\tdisplayFolder: Messaging | Capacity\n\n\tmeasure 'Mailboxes >=45 GB Without Archive' = CALCULATE ( [Mailboxes], KEEPFILTERS ( 'Mailbox Evidence'[Mailbox Size GB] >= 45 ), KEEPFILTERS ( 'Mailbox Evidence'[Archive State] <> "Enabled" ) )\n\t\tformatString: #,0\n\t\tdisplayFolder: Messaging | Risk\n\n`, "measure 'Total Mailbox Storage (TB)'");
  text = replaceOnce(text,
    `{"Archive Size GB", type number}, {"Delegation Count", Int64.Type},`,
    `{"Archive Size GB", type number}, {"Total Storage GB", type number}, {"Delegation Count", Int64.Type},`,
    'mailbox storage query type');
  return text;
}

function updateContentStorage(text) {
  return insertBefore(text, "\tcolumn 'Content Object Key'", `\tmeasure 'Messaging and Collaboration Storage (TB)' = [Total Mailbox Storage (TB)] + [Content Storage Used (TB)]\n\t\tformatString: #,0.0\n\t\tdisplayFolder: Content | Executive\n\n\tmeasure 'Messaging and Collaboration Storage per Workforce Account (GB)' = DIVIDE ( [Total Mailbox Storage (TB)] * 1024 + [Content Storage Used (GB)], [Enabled Workforce Accounts] )\n\t\tformatString: #,0.0\n\t\tdisplayFolder: Content | Executive\n\n`, "measure 'Messaging and Collaboration Storage (TB)'");
}

const licenseOptimization = `table 'License Optimization Evidence'\n\tmeasure 'License Optimization Candidates' = COUNTROWS ( 'License Optimization Evidence' )\n\t\tformatString: #,0\n\n\tmeasure 'Immediate Reclaim Review Candidates' = CALCULATE ( [License Optimization Candidates], 'License Optimization Evidence'[Optimization Opportunity] = "Immediate reclaim review" )\n\t\tformatString: #,0\n\n\tcolumn 'User Principal Name'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: User Principal Name\n\n\tcolumn 'SKU Part Number'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: SKU Part Number\n\n\tcolumn 'License Product'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: License Product\n\n\tcolumn 'Microsoft 365 Plan'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Microsoft 365 Plan\n\n\tcolumn 'Account State'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Account State\n\n\tcolumn 'Activity State'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Activity State\n\n\tcolumn 'Last Activity Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: Last Activity Date\n\n\tcolumn 'Last Office Desktop Activation Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: Last Office Desktop Activation Date\n\n\tcolumn 'Mailbox Size GB'\n\t\tdataType: double\n\t\tformatString: #,0.0\n\t\tsummarizeBy: none\n\t\tsourceColumn: Mailbox Size GB\n\n\tcolumn 'Archive State'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Archive State\n\n\tcolumn 'Optimization Opportunity'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Optimization Opportunity\n\n\tcolumn 'Optimization Reason'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Optimization Reason\n\n\tcolumn 'Recommended Action'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Recommended Action\n\n\tcolumn 'Evidence Status'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Evidence Status\n\n\tcolumn 'Snapshot Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: Snapshot Date\n\n\tpartition 'License Optimization Evidence Import' = m\n\t\tmode: import\n\t\tsource =\n\t\t\t\tlet\n\t\t\t\t    Source = Csv.Document(File.Contents("C:\\Users\\Khaled\\OneDrive\\Documents\\SmartIntune\\SmartWorkplaceIntelligence\\_private\\LicenseOptimizationEvidence.csv"), [Delimiter=",", Columns=15, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),\n\t\t\t\t    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),\n\t\t\t\t    Typed = Table.TransformColumnTypes(Promoted, {{"User Principal Name", type text}, {"SKU Part Number", type text}, {"License Product", type text}, {"Microsoft 365 Plan", type text}, {"Account State", type text}, {"Activity State", type text}, {"Last Activity Date", type date}, {"Last Office Desktop Activation Date", type date}, {"Mailbox Size GB", type number}, {"Archive State", type text}, {"Optimization Opportunity", type text}, {"Optimization Reason", type text}, {"Recommended Action", type text}, {"Evidence Status", type text}, {"Snapshot Date", type date}}, "en-US")\n\t\t\t\tin\n\t\t\t\t    Typed\n`;

const deviceDirectorySummary = `table 'Device Directory Summary'\n\tmeasure 'Enabled AD Windows PCs' = MAX ( 'Device Directory Summary'[Enabled AD Windows PCs Value] )\n\t\tformatString: #,0\n\n\tmeasure 'Enabled AD Windows PCs Managed by Intune' = MAX ( 'Device Directory Summary'[Enabled AD Windows PCs Managed by Intune Value] )\n\t\tformatString: #,0\n\n\tmeasure 'Enabled AD Windows PCs Unmanaged' = MAX ( 'Device Directory Summary'[Enabled AD Windows PCs Unmanaged Value] )\n\t\tformatString: #,0\n\n\tmeasure 'Enabled AD PCs Managed by Intune (%)' = MAX ( 'Device Directory Summary'[Enabled AD PCs Managed by Intune Value] )\n\t\tformatString: 0.0%\n\n\tmeasure 'Disabled AD Computer Objects' = MAX ( 'Device Directory Summary'[Disabled AD Computer Objects Value] )\n\t\tformatString: #,0\n\n\tmeasure 'Disabled Entra Devices' = MAX ( 'Device Directory Summary'[Disabled Entra Devices Value] )\n\t\tformatString: #,0\n\n\tcolumn 'Snapshot Date'\n\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tsummarizeBy: none\n\t\tsourceColumn: Snapshot Date\n\n\tcolumn 'Enabled AD Windows PCs Value'\n\t\tdataType: int64\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Enabled AD Windows PCs\n\n\tcolumn 'Enabled AD Windows PCs Managed by Intune Value'\n\t\tdataType: int64\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Enabled AD Windows PCs Managed by Intune\n\n\tcolumn 'Enabled AD Windows PCs Unmanaged Value'\n\t\tdataType: int64\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Enabled AD Windows PCs Unmanaged\n\n\tcolumn 'Enabled AD PCs Managed by Intune Value'\n\t\tdataType: double\n\t\tformatString: 0.0%\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Enabled AD PCs Managed by Intune (%)\n\n\tcolumn 'Disabled AD Computer Objects Value'\n\t\tdataType: int64\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Disabled AD Computer Objects\n\n\tcolumn 'Disabled Entra Devices Value'\n\t\tdataType: int64\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Disabled Entra Devices\n\n\tcolumn 'Country Definition'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Country Definition\n\n\tcolumn 'Evidence Status'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: Evidence Status\n\n\tpartition 'Device Directory Summary Import' = m\n\t\tmode: import\n\t\tsource =\n\t\t\t\tlet\n\t\t\t\t    Source = Csv.Document(File.Contents("C:\\Users\\Khaled\\OneDrive\\Documents\\SmartIntune\\SmartWorkplaceIntelligence\\_private\\DeviceDirectorySummary.csv"), [Delimiter=",", Columns=9, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),\n\t\t\t\t    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),\n\t\t\t\t    Typed = Table.TransformColumnTypes(Promoted, {{"Snapshot Date", type date}, {"Enabled AD Windows PCs", Int64.Type}, {"Enabled AD Windows PCs Managed by Intune", Int64.Type}, {"Enabled AD Windows PCs Unmanaged", Int64.Type}, {"Enabled AD PCs Managed by Intune (%)", type number}, {"Disabled AD Computer Objects", Int64.Type}, {"Disabled Entra Devices", Int64.Type}, {"Country Definition", type text}, {"Evidence Status", type text}}, "en-US")\n\t\t\t\tin\n\t\t\t\t    Typed\n`;

const changed = [];
if (updateTable('User Inventory Evidence', updateUserInventory)) changed.push('User Inventory Evidence');
if (updateTable('User Activity History Evidence', updateUserHistory)) changed.push('User Activity History Evidence');
if (updateTable('Workforce Trend Evidence', updateWorkforceTrend)) changed.push('Workforce Trend Evidence');
if (updateTable('License Evidence', updateLicenseEvidence)) changed.push('License Evidence');
if (updateTable('Device Inventory Evidence', updateDeviceEvidence)) changed.push('Device Inventory Evidence');
if (updateTable('Mailbox Evidence', updateMailboxEvidence)) changed.push('Mailbox Evidence');
if (updateTable('Content Storage Evidence', updateContentStorage)) changed.push('Content Storage Evidence');

const licenseOptimizationPath = path.join(tableRoot, 'License Optimization Evidence.tmdl');
if (!fs.existsSync(licenseOptimizationPath)) {
  write(licenseOptimizationPath, licenseOptimization);
  changed.push('License Optimization Evidence');
}
const deviceDirectorySummaryPath = path.join(tableRoot, 'Device Directory Summary.tmdl');
if (!fs.existsSync(deviceDirectorySummaryPath)) {
  write(deviceDirectorySummaryPath, deviceDirectorySummary);
  changed.push('Device Directory Summary');
}

const modelPath = path.join(definitionRoot, 'model.tmdl');
let model = read(modelPath);
for (const tableName of ['License Optimization Evidence', 'Device Directory Summary']) {
  const refLine = `ref table '${tableName}'`;
  if (!model.includes(refLine)) model += `\n${refLine}\n`;
}
write(modelPath, model);

// User-facing percentage values use one decimal throughout the report.
for (const entry of fs.readdirSync(tableRoot, { withFileTypes: true })) {
  if (!entry.isFile() || !entry.name.endsWith('.tmdl')) continue;
  const file = path.join(tableRoot, entry.name);
  const original = read(file);
  let normalized = original.replace(/^([\t ]*formatString:\s*)0\.00%\s*$/gm, (_, prefix) => `${prefix}0.0%`);
  normalized = normalized.replaceAll('[Enabled Human Users]', '[Enabled Workforce Accounts]');
  if (normalized !== original) {
    write(file, normalized);
    changed.push(`${entry.name}: percentage format`);
  }
}

console.log(JSON.stringify({ changed }, null, 2));
