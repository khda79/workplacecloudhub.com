// Read-only offline checks against BIM snapshots exported by the modeling MCP.
// Never opens Desktop, evaluates DAX, reads tenant CSVs, or refreshes a model.
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import {fileURLToPath} from 'node:url';
import {isDeepStrictEqual} from 'node:util';

export const retiredTables = [
  'User', 'Device', 'Application', 'License Product', 'Service', 'Group', 'Team',
  'Site', 'Mailbox', 'Data Source', 'Finding Type', 'User Snapshots', 'User Activities',
  'Identity Matches', 'User License Assignments', 'License Capacity', 'Device Snapshots',
  'Device Compliance Snapshots', 'Device Matches', 'Windows Lifecycle Assessments',
  'Endpoint Experience Measurements', 'Application Installations', 'Collaboration Usage',
  'Content Usage', 'Mailbox Snapshots', 'Mailbox Matches', 'Backup Coverage',
  'Security Posture Observations', 'Source Runs', 'Source Coverage', 'Data Quality Results',
  'Findings', 'User Device Relationships', 'User Group Memberships', 'Team Memberships',
  'Mailbox Permissions', 'Service Relationships', 'User Directory Observations',
  'Identity Security Status', 'Conditional Access Coverage', 'Endpoint Security Status',
  'LocalDateTable_975b6c35-8bb1-4f3d-a39b-7f73fb68c00e',
  'LocalDateTable_ab3547e2-cc3e-4859-b585-ddc0b959c470',
  'LocalDateTable_0e067407-fa48-4653-bed4-bd560ca242d8',
  'LocalDateTable_224137e9-4fb1-4b15-8c63-ab2635f24d56',
  'LocalDateTable_db4727be-32c4-4cef-85a9-8b3b83f7f111',
  'LocalDateTable_0a080a09-8c9b-4ecd-b1d6-928b95ea4abf',
  'LocalDateTable_a7b04572-0e01-4ac5-b43f-70dfcc411998',
  'LocalDateTable_c9c293c5-cf73-4433-897e-aa085f83eb42',
];
const scope = new Set(retiredTables);
const text = value => Array.isArray(value) ? value.join('\n') : value;
const description = 'Business-facing semantic model backed by governed SmartM365 evidence and observed histories, including native CMDB identities and relationships. Synthetic demonstration tables are retired. Offline validation does not imply production or Desktop qualification.';

export function checkDependencies(model) {
  const ownedMeasures = new Set(model.tables.filter(t => scope.has(t.name)).flatMap(t => (t.measures || []).map(m => m.name)));
  const tableNames = new Set(model.tables.map(t => t.name));
  let checked = 0;
  function dax(value, where, localColumns) {
    const expression = text(value);
    if (typeof expression !== 'string') return;
    checked++;
    const tokens = [...expression.matchAll(/"(?:[^"]|"")*"|'(?:[^']|'')*'|\[(?:[^\]]|\]\])*\]|\/\/[^\n]*|\/\*[\s\S]*?\*\/|[A-Za-z_][A-Za-z_0-9]*/g)]
      .filter(t => !t[0].startsWith('"') && !t[0].startsWith('//') && !t[0].startsWith('/*'));
    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i], value = token[0];
      if (value.startsWith("'")) {
        assert(!scope.has(value.slice(1, -1).replace(/''/g, "'")), `Retired table dependency: ${where}`);
      } else if (value.startsWith('[')) {
        const name = value.slice(1, -1).replace(/]]/g, ']'), previous = tokens[i - 1];
        const qualified = previous && expression.slice(previous.index + previous[0].length, token.index).trim() === '' &&
          (previous[0].startsWith("'") || tableNames.has(previous[0]));
        assert(qualified || localColumns.has(name) || !ownedMeasures.has(name), `Retired measure dependency: ${where}`);
      } else {
        assert(!scope.has(value), `Retired bare table dependency: ${where}`);
      }
    }
  }
  for (const table of model.tables.filter(t => !scope.has(t.name))) {
    const columns = new Set((table.columns || []).map(c => c.name));
    for (const item of [...(table.measures || []), ...(table.columns || [])]) {
      if (item.expression) dax(item.expression, `${table.name}[${item.name}]`, columns);
    }
    for (const partition of table.partitions || []) {
      if (partition.source.type === 'calculated') dax(partition.source.expression, table.name, columns);
      if (partition.source.type === 'm') {
        for (const match of text(partition.source.expression).matchAll(/#"((?:[^"]|"")*)"/g)) {
          assert(!scope.has(match[1]), `Retired M query dependency: ${table.name}`);
        }
      }
    }
  }
  for (const relationship of model.relationships || []) {
    assert(scope.has(relationship.fromTable) === scope.has(relationship.toTable), `Cross-boundary relationship: ${relationship.name}`);
  }
  // A stale Q&A binding is also a dependency, even without a report consumer.
  function bindings(value) {
    if (!value || typeof value !== 'object') return;
    for (const [key, item] of Object.entries(value)) {
      if (['ConceptualEntity', 'Entity', 'entity'].includes(key)) assert(!scope.has(item), `Retired semantic/report binding: ${item}`);
      if (item && typeof item === 'object') bindings(item);
    }
  }
  bindings(model.cultures);
  assert(!(model.roles || []).length && !(model.perspectives || []).length && !(model.functions || []).length,
    'This retirement audit requires review if roles, perspectives or model functions are added');
  return {checked, bindings};
}

export function verifyRetirement(before, after) {
  const model = before.model;
  assert.equal(retiredTables.length, 49);
  assert.equal(model.tables.filter(t => scope.has(t.name)).length, 49, 'Baseline scope changed');
  const dependency = checkDependencies(model);
  const expected = structuredClone(before);
  expected.model.tables = expected.model.tables.filter(t => !scope.has(t.name));
  expected.model.relationships = expected.model.relationships.filter(r => !scope.has(r.fromTable) && !scope.has(r.toTable));
  expected.model.description = description;
  const order = expected.model.annotations.find(a => a.name === 'PBI_QueryOrder');
  order.value = JSON.stringify(JSON.parse(order.value).filter(name => !scope.has(name)));
  const actual = structuredClone(after);
  for (const db of [expected, actual]) db.model.annotations.sort((a, b) => a.name.localeCompare(b.name));
  // Do not dump a private BIM (including source parameters) into assertion logs.
  assert(isDeepStrictEqual(actual, expected), 'Change outside approved retirement closure');
  assert.equal(after.model.tables.length, 154);
  assert.equal(after.model.tables.reduce((n, t) => n + (t.measures || []).length, 0), 871);
  assert.equal(after.model.relationships.length, 91);
  return dependency;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [beforePath, afterPath, reportRoot] = process.argv.slice(2);
  assert(beforePath && afterPath && reportRoot, 'Usage: node Test-SyntheticTableRetirement.mjs before.bim after.bim ReportRoot');
  const read = file => JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));
  const before = read(beforePath), after = read(afterPath), result = verifyRetirement(before, after);
  let reportFiles = 0;
  function visit(dir) {
    for (const entry of fs.readdirSync(dir, {withFileTypes: true})) {
      const file = path.join(dir, entry.name);
      if (entry.isDirectory()) visit(file);
      else if (file.endsWith('.json') || file.endsWith('.pbir')) { result.bindings(read(file)); reportFiles++; }
    }
  }
  visit(reportRoot);
  // Fault injection proves that accidental retained changes or new dependencies fail closed.
  const changed = structuredClone(after);
  changed.model.tables[0].description = 'Unexpected mutation';
  assert.throws(() => verifyRetirement(before, changed), /outside approved/);
  const broken = structuredClone(before);
  const probeTable = broken.model.tables.find(t => !scope.has(t.name));
  (probeTable.measures ||= []).push({name: 'Probe', expression: "COUNTROWS('User')"});
  assert.throws(() => checkDependencies(broken.model), /Retired table dependency/);
  assert.throws(() => result.bindings({Column: {Expression: {SourceRef: {Entity: 'User'}}}}), /Retired semantic\/report binding/);
  console.log(`Synthetic retirement passed: 49 tables removed; 154 tables, 871 measures, 91 relationships retained; ${result.checked} DAX expressions and ${reportFiles} report JSON/binding files checked; 3 negative controls passed. Offline only.`);
}
