// Small local QA previews of the public templates; previews stay in the ignored review folder.
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
for (const [base, name, sheetName, previewName] of [
  ['config', 'SmartWorkplaceIntelligence-PersonaClassification-template.xlsx', 'Personas', 'persona-template.png'],
  ['config', 'SmartWorkplaceIntelligence-SiteClassification-template.xlsx', 'Sites', 'site-template.png'],
  ['_private', 'SmartWorkplaceIntelligence-PersonaClassification.preserved.xlsx', 'Personas', 'private-personas.png'],
]) {
  const book = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root, base, name)));
  const preview = await book.render({ sheetName, autoCrop: 'all', scale: 1, format: 'png' });
  const target = path.join(root, '_review', previewName);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await preview.arrayBuffer()));
  process.stdout.write(`${target}\n`);
}
