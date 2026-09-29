// Apply the approved, customer-neutral review filter to the local PBIP.
// Activation dates remain diagnostic and are never used to infer Office app usage.
import fs from 'node:fs';
import path from 'node:path';

const project = path.resolve(import.meta.dirname, '..', 'pbip');
const model = path.join(project, 'SmartWorkplaceIntelligence.SemanticModel', 'definition');
const report = path.join(project, 'SmartWorkplaceIntelligence.Report', 'definition');
const read = file => fs.readFileSync(file, 'utf8');
const write = (file, value) => fs.writeFileSync(file, value, 'utf8');
const writeJson = (file, value) => write(file, `${JSON.stringify(value, null, 2)}\n`);

const bridgeFile = path.join(model, 'tables', 'Workforce License Review Selection.tmdl');
const bridge = `/// User-level bridge for the two Workforce Personas review filters. A user may belong to both groups; neither group authorizes an automatic license change.
table 'Workforce License Review Selection'

\t/// Stable Entra user identifier for filtering the governed workforce population.
\tcolumn 'User Source ID'
\t\tdataType: string
\t\tisHidden
\t\tsummarizeBy: none
\t\tsourceColumn: [User Source ID]

\t/// Review category based only on covered usage evidence and the governed E3-to-F3 rule.
\tcolumn 'Review Type'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: [Review Type]

\tpartition 'Workforce License Review Selection' = calculated
\t\tmode: import
\t\tsource =
\t\t\t\tVAR _e3Users =
\t\t\t\t    CALCULATETABLE (
\t\t\t\t        VALUES ( 'License Optimization Evidence'[User Principal Name] ),
\t\t\t\t        'License Optimization Evidence'[Optimization Opportunity] = "E3 to F3 review candidate"
\t\t\t\t    )
\t\t\t\tVAR _e3Review =
\t\t\t\t    SELECTCOLUMNS (
\t\t\t\t        FILTER (
\t\t\t\t            'User Inventory Evidence',
\t\t\t\t            'User Inventory Evidence'[User Principal Name] IN _e3Users
\t\t\t\t        ),
\t\t\t\t        "User Source ID", 'User Inventory Evidence'[User Source ID],
\t\t\t\t        "Review Type", "E3 to F3 review"
\t\t\t\t    )
\t\t\t\tVAR _noUseReview =
\t\t\t\t    SELECTCOLUMNS (
\t\t\t\t        FILTER (
\t\t\t\t            'User Inventory Evidence',
\t\t\t\t            ( 'User Inventory Evidence'[Has Microsoft 365 F3] = "Yes"
\t\t\t\t                || 'User Inventory Evidence'[Has Microsoft 365 E3] = "Yes"
\t\t\t\t                || 'User Inventory Evidence'[Has Microsoft 365 E5] = "Yes" )
\t\t\t\t                && 'User Inventory Evidence'[Exchange Usage State (30D)] <> "Active in 30D"
\t\t\t\t                && 'User Inventory Evidence'[Teams Usage State (30D)] <> "Active in 30D"
\t\t\t\t                && 'User Inventory Evidence'[SharePoint Usage State (30D)] <> "Active in 30D"
\t\t\t\t                && 'User Inventory Evidence'[OneDrive Usage State (30D)] <> "Active in 30D"
\t\t\t\t                && 'User Inventory Evidence'[Office Desktop Usage State (30D)] = "No PC app use in 30D"
\t\t\t\t                && ( 'User Inventory Evidence'[Exchange Usage State (30D)] = "No activity in 30D"
\t\t\t\t                    || 'User Inventory Evidence'[Teams Usage State (30D)] = "No activity in 30D"
\t\t\t\t                    || 'User Inventory Evidence'[SharePoint Usage State (30D)] = "No activity in 30D"
\t\t\t\t                    || 'User Inventory Evidence'[OneDrive Usage State (30D)] = "No activity in 30D" )
\t\t\t\t        ),
\t\t\t\t        "User Source ID", 'User Inventory Evidence'[User Source ID],
\t\t\t\t        "Review Type", "No use - review"
\t\t\t\t    )
\t\t\t\tRETURN DISTINCT ( UNION ( _e3Review, _noUseReview ) )
`;
if (!fs.existsSync(bridgeFile)) write(bridgeFile, bridge);
else if (!read(bridgeFile).includes('RETURN DISTINCT ( UNION ( _e3Review, _noUseReview ) )'))
  throw new Error('Review bridge already exists with different content; preserve the existing edit.');

const modelFile = path.join(model, 'model.tmdl');
let tmdl = read(modelFile);
if (!tmdl.includes("ref table 'Workforce License Review Selection'")) {
  const anchor = "ref table 'License Optimization Evidence'";
  if (!tmdl.includes(anchor)) throw new Error('Model table anchor missing.');
  tmdl = tmdl.replace(anchor, `${anchor}\nref table 'Workforce License Review Selection'`);
  write(modelFile, tmdl);
}

const relationshipFile = path.join(model, 'relationships.tmdl');
tmdl = read(relationshipFile);
if (!tmdl.includes('relationship rel_workforce_license_review_to_inventory')) {
  tmdl += `\nrelationship rel_workforce_license_review_to_inventory
\tcrossFilteringBehavior: bothDirections
\tfromColumn: 'Workforce License Review Selection'.'User Source ID'
\ttoColumn: 'User Inventory Evidence'.'User Source ID'
`;
  write(relationshipFile, tmdl);
}

const userModelFile = path.join(model, 'tables', 'User Inventory Evidence.tmdl');
tmdl = read(userModelFile);
const metricStart = tmdl.indexOf("measure 'Office Desktop Activation (30D %)'");
if (metricStart < 0) throw new Error('Activation diagnostic measure not found.');
const metricEnd = tmdl.indexOf('\n\t\tlineageTag:', metricStart);
if (metricEnd < 0) throw new Error('Activation diagnostic measure boundary not found.');
let metric = tmdl.slice(metricStart, metricEnd);
if (!metric.includes('\n\t\tisHidden')) {
  const oldFolder = '\n\t\tdisplayFolder: Workforce | Personas | Service use';
  if (!metric.includes(oldFolder)) throw new Error('Unexpected activation metric formatting.');
  metric = metric.replace(oldFolder, '\n\t\tisHidden\n\t\tdisplayFolder: Activation diagnostics');
  tmdl = tmdl.slice(0, metricStart) + metric + tmdl.slice(metricEnd);
  write(userModelFile, tmdl);
}
const activationColumn = "\tcolumn 'Office Desktop Activation State (30D)'\n\t\tdataType: string";
if (tmdl.includes(activationColumn)) {
  tmdl = tmdl.replace(activationColumn, `${activationColumn}\n\t\tisHidden`);
  write(userModelFile, tmdl);
}
const optimizationFile = path.join(model, 'tables', 'License Optimization Evidence.tmdl');
tmdl = read(optimizationFile);
const activationDateColumn = "\tcolumn 'Last Office Desktop Activation Date'\n\t\tdataType: dateTime";
if (tmdl.includes(activationDateColumn)) {
  tmdl = tmdl.replace(activationDateColumn, `${activationDateColumn}\n\t\tisHidden`);
  write(optimizationFile, tmdl);
}

const page = '9ae7c13f821d4a608e52';
const visuals = path.join(report, 'pages', page, 'visuals');
const visualFile = id => path.join(visuals, id, 'visual.json');
const reviewId = '521a0ef7c32945cc1c17';
const reviewFile = visualFile(reviewId);
const reviewVisual = JSON.parse(read(visualFile('d3f8c14571af77ec3ce8')));
reviewVisual.name = reviewId;
reviewVisual.position = { x: 816, y: 8, z: 104, width: 192, height: 80, tabOrder: 104 };
const binding = reviewVisual.visual.query.queryState.Values.projections[0];
binding.field.Column.Expression.SourceRef.Entity = 'Workforce License Review Selection';
binding.field.Column.Property = 'Review Type';
binding.queryRef = 'Workforce License Review Selection.Review Type';
binding.nativeQueryRef = 'Review Type';
reviewVisual.visual.objects.header[0].properties.text.expr.Literal.Value = "'License review'";
if (reviewVisual.visual.visualContainerObjects?.general?.[0]?.properties?.altText)
  reviewVisual.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value =
    "'Filter workforce users by E3 to F3 review or No use - review. Accounts in both categories appear under either selection.'";
fs.mkdirSync(path.dirname(reviewFile), { recursive: true });
writeJson(reviewFile, reviewVisual);

for (const id of ['0607d15f5c3e3b0eac4a', '9dead798b5367a10653f']) {
  const file = visualFile(id);
  const visual = JSON.parse(read(file));
  visual.position.width = 704;
  if (id === '0607d15f5c3e3b0eac4a')
    visual.visual.objects.general[0].properties.paragraphs[0].textRuns[0].value =
      'How is the workforce classified and using Microsoft 365?';
  writeJson(file, visual);
}

const licensingNote = path.join(report, 'pages', 'e0a71730533d84d5c995', 'visuals', '3c204b2481b9774eb5d5', 'visual.json');
const note = JSON.parse(read(licensingNote));
const run = note.visual.objects.general[0].properties.paragraphs[0].textRuns[0];
run.value = 'Reclaimable: disabled account or no AD/M365 activity for more than 90 days. E3 to F3: active review candidate, explicitly no Office Windows/Mac app use in the 180-day report, mailbox <= 2 GB and no archive; other workload requirements require business validation.';
writeJson(licensingNote, note);

process.stdout.write('Workforce review filter, activation separation and licensing note updated locally.\n');
