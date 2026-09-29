import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pages = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');
const replacements = new Map([
  ['Enterprise Operational Signals.Enterprise Signal Affected', 'Enterprise Operational Signals.Affected Signal Observations'],
  ['Enterprise Signal Affected', 'Affected Signal Observations'],
  ['Collaboration Adoption Evidence.Collaboration Active User Rate (%)', 'Collaboration Adoption Evidence.Service-weighted Activity Rate (%)'],
  ['Collaboration Active User Rate (%)', 'Service-weighted Activity Rate (%)'],
  ['Collaboration Adoption Evidence.Collaboration Active Users 30D', 'Collaboration Adoption Evidence.Active User-Service Observations 30D'],
  ['Collaboration Active Users 30D', 'Active User-Service Observations 30D']
]);

function walk(dir) {
  return fs.readdirSync(dir, {withFileTypes: true}).flatMap(entry => {
    const target = path.join(dir, entry.name);
    return entry.isDirectory() ? walk(target) : entry.name === 'visual.json' ? [target] : [];
  });
}
function remap(value) {
  if (typeof value === 'string') return replacements.get(value) ?? value;
  if (Array.isArray(value)) return value.map(remap);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, remap(item)]));
  }
  return value;
}

const changes = [];
for (const file of walk(pages)) {
  const before = fs.readFileSync(file, 'utf8');
  const parsed = JSON.parse(before);
  const mapped = remap(parsed);
  if (JSON.stringify(mapped) !== JSON.stringify(parsed)) {
    changes.push({file, content: JSON.stringify(mapped, null, 2) + '\n'});
  }
}
if (changes.length !== 14) throw new Error(`Expected exactly 14 renamed visual files, found ${changes.length}`);
for (const change of changes) fs.writeFileSync(change.file, change.content);
process.stdout.write(`Updated metric bindings in ${changes.length} visual files.\n`);
