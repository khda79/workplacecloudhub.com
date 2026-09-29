import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const file = path.join(root, 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages',
  '3a0fb4ff8eb05c8835b3', 'visuals', '928e023a66d403298322', 'visual.json');
const visual = JSON.parse(fs.readFileSync(file, 'utf8'));
const measure = 'User Inventory Evidence.Enabled Workforce Accounts';
if (visual.visual.visualType !== 'pivotTable' ||
    visual.visual.query.queryState.Values.projections[0].queryRef !== measure) {
  throw new Error('Unexpected User Explorer diagnostic visual or measure binding.');
}
visual.visual.objects.columnFormatting = [{
  properties: {
    labelDisplayUnits: {expr: {Literal: {Value: '1D'}}},
    labelPrecision: {expr: {Literal: {Value: '0L'}}}
  },
  selector: {metadata: measure}
}];
fs.writeFileSync(file, `${JSON.stringify(visual, null, 2)}\n`);
process.stdout.write('User Explorer diagnostic counts set to full integer display.\n');
