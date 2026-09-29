import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const visuals = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report',
  'definition', 'pages', 'a431090c3f25c4a369d0', 'visuals');

const changes = [
  // Diagnostic context above the observed-tenure chart in the right rail.
  ['0715cb44e8590e99f674', {x: 32, y: 256, width: 320, height: 400},
    {x: 1480, y: 256, width: 408, height: 150}],
  ['8b59891c937fbb65a40d', {x: 312, y: 268}, {x: 1848, y: 268}],
  // Expand the device table across the full left content area.
  ['5a9585bad37575c109cd', {x: 376, y: 256, width: 1080, height: 400},
    {x: 32, y: 256, width: 1424}],
  // Preserve the tenure chart but shorten it to fit below the diagnostic.
  ['661f82352c7b55e9f0b9', {x: 1480, y: 256, width: 408, height: 400},
    {y: 422, height: 234}],
  ['9ad57c2675b6735b37d7', {x: 1848, y: 268}, {y: 434}],
];

const output = [];
for (const [id, expected, target] of changes) {
  const file = path.join(visuals, id, 'visual.json');
  const json = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (json.name !== id) throw new Error(`Unexpected visual ID: ${file}`);
  for (const [property, value] of Object.entries(expected)) {
    if (json.position[property] !== value) {
      throw new Error(`Unexpected ${property} for ${id}: ${json.position[property]}`);
    }
  }
  Object.assign(json.position, target);
  output.push({file, json});
}
for (const {file, json} of output) {
  fs.writeFileSync(file, `${JSON.stringify(json, null, 2)}\n`);
}
process.stdout.write('Updated Device Explorer diagnostic, detail, and tenure layout.\n');
