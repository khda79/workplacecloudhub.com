// Create a generic, shareable persona-classification workbook. Never import tenant data here.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const modules = process.env.CODEX_ARTIFACT_TOOL_NODE_MODULES ||
  path.join(os.homedir(), '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'node', 'node_modules');
const packagePath = createRequire(path.join(modules, '_anchor.cjs')).resolve('@oai/artifact-tool');
const { SpreadsheetFile, Workbook } = await import(pathToFileURL(packagePath).href);

const root = path.resolve(import.meta.dirname, '..');
const outputPath = process.argv[2] || path.join(root, 'config', 'SmartWorkplaceIntelligence-PersonaClassification-template.xlsx');
const workbook = Workbook.create();

function addSheet(name, headers, widths, tableName, rows) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const cells = sheet.getRangeByIndexes(0, 0, rows.length + 1, headers.length);
  cells.values = [headers, ...rows];
  sheet.getRangeByIndexes(0, 0, 1, headers.length).format = {
    fill: '#17365D', font: { name: 'Arial', bold: true, color: '#FFFFFF', size: 10 },
  };
  sheet.getRangeByIndexes(1, 0, rows.length, headers.length).format.font = { name: 'Arial', color: '#172B4D', size: 10 };
  widths.forEach((width, index) => { sheet.getRangeByIndexes(0, index, rows.length + 1, 1).format.columnWidth = width; });
  cells.format.rowHeight = 22;
  const table = sheet.tables.add(cells.address, true, tableName);
  table.style = 'TableStyleMedium2';
  sheet.freezePanes.freezeRows(1);
  return sheet;
}

addSheet('Personas',
  ['Persona ID', 'Persona name', 'Description', 'Display order', 'Enabled'],
  [24, 38, 58, 18, 15], 'PersonasCatalog', [
    ['HEADQUARTERS', 'Headquarters staff', 'Accounts linked to a governed HQ site; site match takes priority', 0, 'Yes'],
    ['LEADERSHIP', 'Leadership', 'Executives and senior business leaders', 1, 'Yes'],
    ['IT', 'IT and Digital', 'Technology, security and digital delivery roles', 2, 'Yes'],
    ['FINANCE', 'Finance', 'Accounting, controlling and procurement roles', 3, 'Yes'],
    ['HR', 'Human Resources', 'People, talent and payroll roles', 4, 'Yes'],
    ['SALES', 'Sales and Marketing', 'Commercial, account and marketing roles', 5, 'Yes'],
    ['OPERATIONS', 'Operations', 'Operational delivery and field roles', 6, 'Yes'],
    ['SUPPORT', 'Administrative Support', 'Administrative and customer support roles', 7, 'Yes'],
  ]);
addSheet('Rules',
  ['Rule ID', 'Keyword', 'Country', 'Persona ID', 'Match type', 'Priority', 'Enabled', 'Source category'],
  [16, 38, 15, 24, 22, 15, 15, 40], 'PersonaRules', [
    ['R001', 'chief executive', 'ALL', 'LEADERSHIP', 'Contains phrase', 100, 'Yes', 'Generic example'],
    ['R002', 'director', 'ALL', 'LEADERSHIP', 'Contains phrase', 40, 'Yes', 'Generic example'],
    ['R003', 'information technology', 'ALL', 'IT', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R004', 'cybersecurity', 'ALL', 'IT', 'Contains phrase', 70, 'Yes', 'Generic example'],
    ['R005', 'accountant', 'ALL', 'FINANCE', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R006', 'procurement', 'ALL', 'FINANCE', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R007', 'human resources', 'ALL', 'HR', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R008', 'recruitment', 'ALL', 'HR', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R009', 'sales', 'ALL', 'SALES', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R010', 'marketing', 'ALL', 'SALES', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R011', 'operations', 'ALL', 'OPERATIONS', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R012', 'field technician', 'ALL', 'OPERATIONS', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R013', 'administrative assistant', 'ALL', 'SUPPORT', 'Contains phrase', 60, 'Yes', 'Generic example'],
    ['R014', 'customer support', 'ALL', 'SUPPORT', 'Contains phrase', 60, 'Yes', 'Generic example'],
  ]);
addSheet('Exclusions',
  ['Rule ID', 'Keyword', 'Country', 'Reason', 'Result', 'Enabled'],
  [16, 38, 15, 58, 44, 15], 'PersonaExclusions', [
    ['X001', 'scan to mail', 'ALL', 'Likely a shared or technical account', 'Review required', 'Yes'],
  ]);
addSheet('TestCases',
  ['Job title', 'Country', 'Expected result', 'Purpose'],
  [42, 15, 46, 56], 'PersonaTestCases', [
    ['Chief Executive Officer', 'ALL', 'Leadership', 'Positive match'],
    ['Cybersecurity Analyst', 'ALL', 'IT and Digital', 'Positive match'],
    ['Unspecified role', 'ALL', 'Unclassified', 'No forced classification'],
    ['Scan to mail account', 'ALL', 'Review required', 'Non-human safeguard'],
  ]);

const guide = workbook.worksheets.add('Guide');
guide.showGridLines = false;
const guidance = [
  ['Topic', 'Guidance'],
  ['Purpose', 'Adapt the generic example personas and job-title rules to your organization. No customer data is included.'],
  ['Setup', 'Save the completed workbook as SmartWorkplaceIntelligence-PersonaClassification.xlsx.'],
  ['Location', 'Place it in a private data folder approved by your organization, outside the source-code repository.'],
  ['Personas', 'Add a stable unique Persona ID, name, description, order and Enabled = Yes.'],
  ['Rules', 'Use a unique Rule ID, keyword, Country, existing Persona ID, Match type, Priority and Enabled = Yes.'],
  ['Country', 'Use one ISO country code per rule or ALL for a global rule.'],
  ['Match type', 'The classifier currently supports Contains phrase.'],
  ['Priority', 'Use an integer. If equal-priority rules match different personas, the account requires review.'],
  ['Matching order', 'Explicit non-human exclusions first; then Headquarters site; then unique job title; then AD Description if title has no unique match.'],
  ['Exclusions', 'List terms that should not assign a human persona. This sheet may be empty.'],
  ['Test cases', 'Document representative titles and expected classifications before using new rules.'],
  ['Source category', 'Optional provenance label for your own rules.'],
  ['Privacy', 'Keep the completed mapping and any customer job titles outside source control.'],
];
const guideRange = guide.getRangeByIndexes(0, 0, guidance.length, 2);
guideRange.values = guidance;
guide.getRange('A1:B1').format = {
  fill: '#17365D', font: { name: 'Arial', bold: true, color: '#FFFFFF', size: 10 },
};
guide.getRangeByIndexes(1, 0, guidance.length - 1, 2).format.font = { name: 'Arial', size: 10, color: '#172B4D' };
guide.getRangeByIndexes(0, 0, guidance.length, 1).format.columnWidth = 24;
guide.getRangeByIndexes(0, 1, guidance.length, 1).format.columnWidth = 112;
guideRange.format.rowHeight = 23;
const guideTable = guide.tables.add(guideRange.address, true, 'PersonaGuide');
guideTable.style = 'TableStyleMedium2';
guide.freezePanes.freezeRows(1);

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const result = await SpreadsheetFile.exportXlsx(workbook);
await result.save(outputPath);
process.stdout.write(JSON.stringify({ outputPath, sheets: ['Personas', 'Rules', 'Exclusions', 'TestCases', 'Guide'] }) + '\n');
