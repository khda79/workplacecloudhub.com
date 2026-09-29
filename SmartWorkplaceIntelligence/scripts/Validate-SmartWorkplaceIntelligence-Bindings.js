#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const reportPages = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');
const modelTables = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.SemanticModel', 'definition', 'tables');

function walkJson(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    const target = path.join(directory, entry.name);
    if (entry.isDirectory()) return walkJson(target);
    return entry.name === 'visual.json' ? [target] : [];
  });
}

function unquoteTmdl(identifier) {
  const value = identifier.trim();
  if (value.startsWith("'") && value.endsWith("'")) return value.slice(1, -1).replaceAll("''", "'");
  return value;
}

const model = new Map();
for (const file of fs.readdirSync(modelTables).filter(name => name.endsWith('.tmdl'))) {
  const table = path.basename(file, '.tmdl');
  const objects = { column: new Set(), measure: new Set() };
  for (const line of fs.readFileSync(path.join(modelTables, file), 'utf8').split(/\r?\n/)) {
    const match = line.match(/^\t(column|measure) (.+?)(?: =|$)/);
    if (match) objects[match[1]].add(unquoteTmdl(match[2]));
  }
  model.set(table, objects);
}

const bindings = [];
function collect(node, file) {
  if (!node || typeof node !== 'object') return;
  for (const kind of ['Column', 'Measure']) {
    const binding = node[kind];
    const table = binding?.Expression?.SourceRef?.Entity;
    const property = binding?.Property;
    if (table && property) bindings.push({ file, kind: kind.toLowerCase(), table, property });
  }
  for (const value of Object.values(node)) collect(value, file);
}

for (const file of walkJson(reportPages)) collect(JSON.parse(fs.readFileSync(file, 'utf8')), file);

const unique = new Map();
for (const binding of bindings) unique.set(`${binding.kind}|${binding.table}|${binding.property}`, binding);
const missing = [...unique.values()].filter(binding => !model.get(binding.table)?.[binding.kind]?.has(binding.property));

const result = {
  modelTables: model.size,
  reportVisualFiles: walkJson(reportPages).length,
  uniqueBindings: unique.size,
  missingBindings: missing
};
process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
if (missing.length) process.exitCode = 1;
