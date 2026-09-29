import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pages = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');
const page = path.join(pages, '3a0fb4ff8eb05c8835b3', 'visuals');
const reference = path.join(pages, '0443c39bb6eba584da07', 'visuals',
  '5565e85fb1fe7df8eb89', 'visual.json');

function read(id) {
  const file = path.join(page, id, 'visual.json');
  const json = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (json.name !== id) throw new Error(`Unexpected visual ID: ${file}`);
  return {file, json};
}

const items = new Map();
function update(id, original, target) {
  const item = read(id);
  for (const [key, value] of Object.entries(original)) {
    if (item.json.position[key] !== value) {
      throw new Error(`Unexpected ${key} for ${id}: ${item.json.position[key]}`);
    }
  }
  Object.assign(item.json.position, target);
  items.set(id, item);
  return item.json;
}

// The left-hand diagnostic block moves to the top of the right rail.
update('928e023a66d403298322', {x: 32, y: 256, width: 320, height: 400},
  {x: 1480, y: 256, width: 408, height: 136});
update('59df53a8b96b6f4c50c6', {x: 312, y: 268}, {x: 1848, y: 268});

// The detail table gets the full left content width.
update('a65c81cadf407124bcf7', {x: 376, y: 256, width: 1080, height: 400},
  {x: 32, y: 256, width: 1424});

// Preserve chart height while stacking persona and job title under the diagnostic.
update('93cfd1ba5a2539e431e3', {x: 1480, y: 256, height: 192},
  {y: 408, height: 184});
update('6fc01f33d0d09c8961bd', {x: 1848, y: 268}, {y: 420});
update('0b327acf93e04d91865a', {x: 1480, y: 464, height: 192},
  {y: 608, height: 184});

// Compact the final panel to fit below the other right-rail visuals.
update('d7b514052000b9e26eb2', {x: 1480, y: 680, height: 328},
  {y: 808, height: 200});
const heading = update('b6930250059a16161dad', {x: 1492, y: 688}, {y: 816});
heading.visual.objects.general[0].properties.paragraphs[0].textRuns[0].value =
  'Selected-scope activity · last 30 days';
update('2e1e3494e2b4e1c35456', {x: 1492, y: 712}, {y: 840});
update('81d79d551eac12698e9f', {x: 1848, y: 692}, {y: 820});
const trend = update('bcbe3916207958d30bfa', {x: 1492, y: 732, height: 238},
  {y: 860, height: 124});
const change = update('f5eb42f34008c4a584fc', {x: 1744, y: 728}, {y: 856});

const relative = JSON.parse(fs.readFileSync(reference, 'utf8'))
  .filterConfig?.filters?.find(f => f.type === 'RelativeDate' &&
    f.field?.Column?.Property === 'Snapshot Date');
if (!relative || relative.filter?.Where?.[0]?.Condition?.Between?.LowerBound
  ?.DateSpan?.Expression?.DateAdd?.Amount !== -30) {
  throw new Error('Missing last-30-days reference filter');
}
function applyDateFilter(visual, name) {
  if (visual.filterConfig?.filters?.some(f =>
    f.field?.Column?.Property === 'Snapshot Date')) {
    throw new Error(`Date filter already exists on ${visual.name}`);
  }
  const filter = structuredClone(relative);
  filter.name = name;
  filter.field.Column.Expression.SourceRef.Entity = 'User Activity History Evidence';
  filter.filter.From[0].Entity = 'User Activity History Evidence';
  visual.filterConfig = {filters: [...(visual.filterConfig?.filters ?? []), filter]};
}
applyDateFilter(trend, 'Filter30DaysUserExplorerTrend');
applyDateFilter(change, 'Filter30DaysUserExplorerChange');
trend.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
  "'Selected-scope active workforce by observed snapshot date in the last 30 days.'";

for (const {file, json} of items.values()) {
  fs.writeFileSync(file, `${JSON.stringify(json, null, 2)}\n`);
}
process.stdout.write('Updated User Explorer layout and 30-day trend.\n');
