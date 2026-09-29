// Deterministic local TMDL edit for the approved Workforce Personas measures.
// Source values and customer-specific classification rules are never embedded here.
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve(import.meta.dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.SemanticModel',
  'definition', 'tables', 'User Inventory Evidence.tmdl');
let source = fs.readFileSync(file, 'utf8');
const newline = source.includes('\r\n') ? '\r\n' : '\n';
const normalize = value => value.replaceAll('\n', newline);
function once(before, after) {
  if (!source.includes(before)) throw new Error(`Expected TMDL anchor missing: ${before.slice(0, 80)}`);
  source = source.replace(before, normalize(after));
}

if (!source.includes("measure 'Any M365 Service Usage (30D %)'") ) {
  once("\tcolumn 'User Source ID'", `
\t/// Share of covered workforce accounts active in at least one of Exchange, Teams, SharePoint or OneDrive during the last 30 days. Accounts without a covered workload are excluded.
\tmeasure 'Any M365 Service Usage (30D %)' = VAR _covered = FILTER ( 'User Inventory Evidence', 'User Inventory Evidence'[Exchange Usage State (30D)] IN { "Active in 30D", "No activity in 30D" } || 'User Inventory Evidence'[Teams Usage State (30D)] IN { "Active in 30D", "No activity in 30D" } || 'User Inventory Evidence'[SharePoint Usage State (30D)] IN { "Active in 30D", "No activity in 30D" } || 'User Inventory Evidence'[OneDrive Usage State (30D)] IN { "Active in 30D", "No activity in 30D" } ) RETURN DIVIDE ( COUNTROWS ( FILTER ( _covered, 'User Inventory Evidence'[Exchange Usage State (30D)] = "Active in 30D" || 'User Inventory Evidence'[Teams Usage State (30D)] = "Active in 30D" || 'User Inventory Evidence'[SharePoint Usage State (30D)] = "Active in 30D" || 'User Inventory Evidence'[OneDrive Usage State (30D)] = "Active in 30D" ) ), COUNTROWS ( _covered ) )
\t\tformatString: 0.0%
\t\tdisplayFolder: Workforce | Personas | Service use
\t\tlineageTag: 131e1bfd-c54c-4d61-9054-637d0a4642ba

\t/// Recent Microsoft 365 Apps for enterprise Windows or Mac activation among workforce accounts with an Apps activation report row. Activation is not proof of app usage.
\tmeasure 'Office Desktop Activation (30D %)' = DIVIDE ( CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Office Desktop Activation State (30D)] = "Activated on PC in 30D" ), CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Office Desktop Activation State (30D)] IN { "Activated on PC in 30D", "No PC activation in 30D" } ) )
\t\tformatString: 0.0%
\t\tisHidden
\t\tdisplayFolder: Activation diagnostics
\t\tlineageTag: fa8f8904-d70d-4c24-82db-a401aa947be2

\t/// F3/E3/E5 assignments with no active four-service signal, at least one explicit no-activity workload, and explicit no Office PC app use in the covered 30-day report. Review candidates only.
\tmeasure 'M365 Assignments Without Observed Use (30D)' = SUMX ( FILTER ( 'User Inventory Evidence', 'User Inventory Evidence'[Exchange Usage State (30D)] <> "Active in 30D" && 'User Inventory Evidence'[Teams Usage State (30D)] <> "Active in 30D" && 'User Inventory Evidence'[SharePoint Usage State (30D)] <> "Active in 30D" && 'User Inventory Evidence'[OneDrive Usage State (30D)] <> "Active in 30D" && 'User Inventory Evidence'[Office Desktop Usage State (30D)] = "No PC app use in 30D" && ( 'User Inventory Evidence'[Exchange Usage State (30D)] = "No activity in 30D" || 'User Inventory Evidence'[Teams Usage State (30D)] = "No activity in 30D" || 'User Inventory Evidence'[SharePoint Usage State (30D)] = "No activity in 30D" || 'User Inventory Evidence'[OneDrive Usage State (30D)] = "No activity in 30D" ) ), IF ( 'User Inventory Evidence'[Has Microsoft 365 F3] = "Yes", 1, 0 ) + IF ( 'User Inventory Evidence'[Has Microsoft 365 E3] = "Yes", 1, 0 ) + IF ( 'User Inventory Evidence'[Has Microsoft 365 E5] = "Yes", 1, 0 ) )
\t\tformatString: #,0
\t\tdisplayFolder: Workforce | Personas | Licenses
\t\tlineageTag: 4fc5756f-1bd3-44ea-9dce-ff35d933f6af

\t/// Existing E3-to-F3 review candidate assignments attributed to the current workforce persona by UPN. No E5 downgrade rule is applied.
\tmeasure 'E3 to F3 Review Candidates by Persona' = CALCULATE ( COUNTROWS ( 'License Optimization Evidence' ), 'License Optimization Evidence'[Optimization Opportunity] = "E3 to F3 review candidate", TREATAS ( VALUES ( 'User Inventory Evidence'[User Principal Name] ), 'License Optimization Evidence'[User Principal Name] ) )
\t\tformatString: #,0
\t\tdisplayFolder: Workforce | Personas | Licenses
\t\tlineageTag: 5ec09db8-3b26-4255-9281-20abb0f1b757

\tcolumn 'User Source ID'`);
}

if (!source.includes("column 'Office Desktop Activation State (30D)'")) {
  once("\tcolumn 'M365 Usage Report Date'", `\tcolumn 'Office Desktop Activation State (30D)'
\t\tdataType: string
\t\tisHidden
\t\tlineageTag: 0d0d5979-641e-4799-8f90-48a2db2bd445
\t\tsummarizeBy: none
\t\tsourceColumn: Office Desktop Activation State (30D)

\tcolumn 'M365 Usage Report Date'`);
}
if (source.includes('Columns=43, Encoding=65001'))
  once('Columns=43, Encoding=65001', 'Columns=44, Encoding=65001');
if (!source.includes('{"Office Desktop Activation State (30D)", type text}'))
  once('{"OneDrive Usage State (30D)", type text},', '{"OneDrive Usage State (30D)", type text},\n\t\t\t\t        {"Office Desktop Activation State (30D)", type text},');
if (!source.includes("column 'Office Desktop Usage State (30D)'")) {
  once("\tcolumn 'Office Desktop Activation State (30D)'", `\tcolumn 'Office Desktop Usage State (30D)'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Office Desktop Usage State (30D)

\tcolumn 'Office Desktop Usage Report Date'
\t\tdataType: dateTime
\t\tformatString: Short Date
\t\tsummarizeBy: none
\t\tsourceColumn: Office Desktop Usage Report Date

\tcolumn 'Office Desktop Activation State (30D)'`);
}
if (source.includes('Columns=44, Encoding=65001'))
  once('Columns=44, Encoding=65001', 'Columns=46, Encoding=65001');
else if (!source.includes('Columns=46, Encoding=65001'))
  throw new Error('Unexpected User Inventory Evidence CSV column count');
if (!source.includes('{"Office Desktop Usage State (30D)", type text}'))
  once('{"Office Desktop Activation State (30D)", type text},', '{"Office Desktop Usage State (30D)", type text},\n\t\t\t\t        {"Office Desktop Usage Report Date", type date},\n\t\t\t\t        {"Office Desktop Activation State (30D)", type text},');
source = source.replaceAll('recent PC activation in 30 days', 'explicit no Office Windows/Mac app use in the 30-day report');
source = source.replaceAll('F3/E3/E5 assignments on covered accounts with no observed four-service activity or explicit no Office Windows/Mac app use in the 30-day report. Review candidates, not confirmed reclaimable licenses.', 'F3/E3/E5 assignments with no active four-service signal, at least one explicit no-activity workload, and explicit no Office PC app use in the covered 30-day report. Review candidates only.');
source = source.replaceAll('F3/E3/E5 assignments with no observed four-service activity and explicit no Office Windows/Mac app use in the 30-day report. Unknown app evidence is excluded; review only.', 'F3/E3/E5 assignments with no active four-service signal, at least one explicit no-activity workload, and explicit no Office PC app use in the covered 30-day report. Review candidates only.');
source = source.replaceAll('/// Graph D180 report refresh date used as reference for 30-day service activity.\n\tcolumn \'Office Desktop Usage State (30D)\'', '/// Graph D30 report state for intentional Office app usage on Windows or Mac.\n\tcolumn \'Office Desktop Usage State (30D)\'');
source = source.replaceAll("'User Inventory Evidence'[Office Desktop Activation State (30D)] <> \"Activated on PC in 30D\"", "'User Inventory Evidence'[Office Desktop Usage State (30D)] = \"No PC app use in 30D\"");
if (!source.includes("measure 'Office Desktop Usage (30D %)'")) {
  once("\t/// Existing E3-to-F3 review candidate assignments", `\t/// Share of covered workforce accounts with intentional Outlook, Word, Excel, PowerPoint or OneNote activity on Windows or Mac in the 30-day Apps report.
\tmeasure 'Office Desktop Usage (30D %)' = DIVIDE ( CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Office Desktop Usage State (30D)] = "Used on PC in 30D" ), CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Office Desktop Usage State (30D)] IN { "Used on PC in 30D", "No PC app use in 30D" } ) )
\t\tformatString: 0.0%
\t\tdisplayFolder: Workforce | Personas | Service use

\t/// Existing E3-to-F3 review candidate assignments`);
}
fs.writeFileSync(file, source, 'utf8');
process.stdout.write('Persona semantic model measures and Office app usage source columns updated locally.\n');
