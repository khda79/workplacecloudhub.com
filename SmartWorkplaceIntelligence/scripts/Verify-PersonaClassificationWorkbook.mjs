// Render each governance sheet for visual review without modifying the workbook.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const modules = process.env.CODEX_ARTIFACT_TOOL_NODE_MODULES ||
  path.join(os.homedir(), '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'node', 'node_modules');
const packagePath = createRequire(path.join(modules, '_anchor.cjs')).resolve('@oai/artifact-tool');
const { FileBlob, SpreadsheetFile } = await import(pathToFileURL(packagePath).href);

const root = path.resolve(import.meta.dirname, '..');
const input = process.argv[2] || path.join(root, 'config', 'SmartWorkplaceIntelligence-PersonaClassification-template.xlsx');
const output = process.argv[3] || path.join(root, '_private', 'persona-workbook-review');
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const ranges = new Map([
  ['Personas', 'A1:E5'],
  ['Rules', 'A1:H12'],
  ['Exclusions', 'A1:F3'],
  ['TestCases', 'A1:D6'],
  ['Guide', 'A1:B13'],
]);
await fs.mkdir(output, { recursive: true });
for (const [sheetName, range] of ranges) {
  const image = await workbook.render({ sheetName, range, scale: 1, format: 'png' });
  await fs.writeFile(path.join(output, `${sheetName}.png`), new Uint8Array(await image.arrayBuffer()));
}
process.stdout.write(JSON.stringify({ input, output, sheets: [...ranges.keys()] }) + '\n');
