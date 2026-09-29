// Deterministic PBIP/TMDL synchronization for the approved workforce and device evidence columns.
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.SemanticModel', 'definition', 'tables');

function update(name, changes) {
  const file = path.join(root, `${name}.tmdl`);
  let value = fs.readFileSync(file, 'utf8');
  const original = value;
  if (name === 'User Inventory Evidence') {
    value = value.replace("measure 'Workforce MFA Registered (%)'", "measure 'Workforce MFA Registration Rate (%)'");
    const marker = "\tmeasure 'Classified Workforce Accounts'";
    const first = value.indexOf(marker);
    const second = first < 0 ? -1 : value.indexOf(marker, first + marker.length);
    if (second >= 0) {
      const column = value.indexOf("\tcolumn 'User Source ID'", second);
      if (column < 0) throw new Error('Cannot locate workforce user column after duplicate measure block');
      value = value.slice(0, second) + value.slice(column);
    }
  }
  for (const [before, after] of changes) {
    if (name === 'User Inventory Evidence' && after.startsWith("\tmeasure 'Classified Workforce Accounts'") && value.includes("\tmeasure 'Classified Workforce Accounts'")) continue;
    if (after.endsWith(before)) {
      const insertion = after.slice(0, -before.length);
      const declarations = [...insertion.matchAll(/\t(measure|column) '([^']+)'/g)];
      if (declarations.length) {
        const present = declarations.map(([, kind, label]) => value.includes(`\t${kind} '${label}'`));
        if (present.every(Boolean)) continue;
        if (present.some(Boolean)) throw new Error(`${name}: partial declaration insertion before ${before.slice(0, 60)}`);
      }
    }
    if (after.startsWith(before) && after !== before) {
      const added = after.slice(before.length);
      const typedNames = [...added.matchAll(/\{"([^"]+)", type text\},/g)].map(([, label]) => label);
      if (typedNames.length && typedNames.every(label => value.includes(`{"${label}", type text},`))) continue;
    }
    if (after.endsWith(before)) {
      const insertion = after.slice(0, -before.length);
      while (value.includes(insertion + insertion + before)) {
        value = value.replace(insertion + insertion + before, insertion + before);
      }
    }
    if (value.includes(after)) continue;
    const count = value.split(before).length - 1;
    if (count === 1) value = value.replace(before, after);
    else throw new Error(`${name}: expected one match, found ${count}: ${before.slice(0, 90)}`);
  }
  const typedFields = name === 'User Inventory Evidence'
    ? ['Workforce Persona', 'Persona Match State', 'Persona Matched Keywords', 'MFA Registration State']
    : name === 'Device Inventory Evidence' ? ['Intune Encryption State'] : [];
  for (const field of typedFields) {
    const typedEntry = `{"${field}", type text},`;
    let seen = false;
    value = value.split('\n').filter(line => {
      if (!line.includes(typedEntry)) return true;
      if (seen) return false;
      seen = true;
      return true;
    }).join('\n');
  }
  if (value !== original) fs.writeFileSync(file, value, 'utf8');
  process.stdout.write(`${name}: ${value === original ? 'already current' : 'updated'}\n`);
}

const deviceSummaryMeasures = `	measure 'Enabled AD Windows 10 PCs' = MAX ( 'Device Directory Summary'[Enabled AD Windows 10 PCs Value] )
		formatString: #,0

	measure 'Enabled AD Windows 10 PCs Without Intune' = MAX ( 'Device Directory Summary'[Enabled AD Windows 10 PCs Without Intune Value] )
		formatString: #,0

	measure 'AD Windows 10 Ambiguous Intune Matches' = MAX ( 'Device Directory Summary'[AD Windows 10 Ambiguous Intune Matches Value] )
		formatString: #,0

	measure 'Intune Windows 10 Eligibility Undetermined' = MAX ( 'Device Directory Summary'[Intune Windows 10 Eligibility Undetermined Value] )
		formatString: #,0

	/// Intune Windows 10 devices without a known readiness result, plus enabled AD Windows 10 PCs without an Intune match. Ambiguous matches are kept separate.
	measure 'Windows 10 Eligibility Undetermined' = MAX ( 'Device Directory Summary'[Windows 10 Eligibility Undetermined Value] )
		formatString: #,0

`;
const deviceSummaryColumns = [
  'Enabled AD Windows 10 PCs',
  'Enabled AD Windows 10 PCs Without Intune',
  'AD Windows 10 Ambiguous Intune Matches',
  'Intune Windows 10 Eligibility Undetermined',
  'Windows 10 Eligibility Undetermined',
].map(name => `	column '${name} Value'\n\t\tdataType: int64\n\t\tsummarizeBy: sum\n\t\tsourceColumn: ${name}\n\n`).join('');
const deviceSummaryTyped = [
  'Enabled AD Windows 10 PCs',
  'Enabled AD Windows 10 PCs Without Intune',
  'AD Windows 10 Ambiguous Intune Matches',
  'Intune Windows 10 Eligibility Undetermined',
  'Windows 10 Eligibility Undetermined',
].map(name => `{"${name}", Int64.Type}, `).join('');
update('Device Directory Summary', [
  ["\tcolumn 'Snapshot Date'", `${deviceSummaryMeasures}\tcolumn 'Snapshot Date'`],
  ["\tcolumn 'Country Definition'", `${deviceSummaryColumns}\tcolumn 'Country Definition'`],
  ['Columns=9, Encoding=65001', 'Columns=14, Encoding=65001'],
  ['{"Disabled Entra Devices", Int64.Type}, {"Country Definition"', `{"Disabled Entra Devices", Int64.Type}, ${deviceSummaryTyped}{"Country Definition"`],
]);

const encryptionMeasure = `	/// Devices with Encrypted=True in the current Intune managed-device export. This does not prove BitLocker recovery-key escrow.
	measure 'Devices Encrypted According to Intune' = CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Intune Encryption State] = "Encrypted" )
		formatString: #,0
		displayFolder: Device | Security

`;
const encryptionColumn = `	column 'Intune Encryption State'
		dataType: string
		summarizeBy: none
		sourceColumn: Intune Encryption State

`;
update('Device Inventory Evidence', [
  ["\tcolumn 'Device Source ID'", `${encryptionMeasure}\tcolumn 'Device Source ID'`],
  ["\tcolumn 'Entra Account State'", `${encryptionColumn}\tcolumn 'Entra Account State'`],
  ['Columns = 53, Encoding = 65001', 'Columns = 54, Encoding = 65001'],
  ['{"Entra Account State", type text},', '{"Entra Account State", type text},\n\t\t\t\t        {"Intune Encryption State", type text},'],
]);

const workforceMeasures = `	measure 'Classified Workforce Accounts' = CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Persona Match State] = "Classified" )
		formatString: #,0
		displayFolder: Workforce | Personas

	measure 'Persona Classification Coverage (%)' = DIVIDE ( [Classified Workforce Accounts], [Enabled Workforce Accounts] )
		formatString: 0.0%
		displayFolder: Workforce | Personas

	measure 'Workforce Accounts Requiring Persona Review' = CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Persona Match State] IN { "Ambiguous", "Non-human keyword" } )
		formatString: #,0
		displayFolder: Workforce | Personas

	measure 'Workforce MFA Registration Covered' = CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[MFA Registration State] IN { "Registered", "Not registered" } )
		formatString: #,0
		displayFolder: Workforce | MFA

	measure 'Workforce MFA Registration Rate (%)' = DIVIDE ( CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[MFA Registration State] = "Registered" ), [Workforce MFA Registration Covered] )
		formatString: 0.0%
		displayFolder: Workforce | MFA

	measure 'Workforce MFA Registration Coverage (%)' = DIVIDE ( [Workforce MFA Registration Covered], [Enabled Workforce Accounts] )
		formatString: 0.0%
		displayFolder: Workforce | MFA

`;
const workforceColumns = ['Workforce Persona', 'Persona Match State', 'Persona Matched Keywords', 'MFA Registration State']
  .map(name => `	column '${name}'\n\t\tdataType: string\n\t\tsummarizeBy: none\n\t\tsourceColumn: ${name}\n\n`).join('');
const workforceTyped = ['Workforce Persona', 'Persona Match State', 'Persona Matched Keywords', 'MFA Registration State']
  .map(name => `\n\t\t\t\t        {"${name}", type text},`).join('');
update('User Inventory Evidence', [
  ["\tcolumn 'User Source ID'", `${workforceMeasures}\tcolumn 'User Source ID'`],
  ["\tcolumn 'AD Last Activity Date'", `${workforceColumns}\tcolumn 'AD Last Activity Date'`],
  ['Columns=31, Encoding=65001', 'Columns=35, Encoding=65001'],
  ['{"Job Title Evidence State", type text},', `{"Job Title Evidence State", type text},${workforceTyped}`],
]);

update('User Inventory Evidence', [
  ['/// Private current workforce evidence built from enabled human Entra accounts, M365 activity, AD identity reconciliation, manager, department, and license evidence.', '/// Private current workforce evidence built from enabled AD-linked Entra member accounts excluding governed non-human types and cloud-only accounts; AD and M365 activity are combined.'],
  ['/// Enabled human users with observed M365 activity in the last 30 days.', '/// Enabled workforce accounts with observed AD or Microsoft 365 activity in the last 30 days.'],
  ['/// Activity-covered enabled human users without observed M365 activity for more than 30 days, including never-used accounts.', '/// Activity-covered enabled workforce accounts without observed AD or Microsoft 365 activity for more than 30 days, including never-used accounts.'],
  ['/// Enabled human users covered by the M365 activity export.', '/// Enabled workforce accounts with observed AD or Microsoft 365 activity evidence.'],
  ['/// Share of the enabled human workforce covered by the M365 activity export.', '/// Share of enabled workforce accounts covered by AD or Microsoft 365 activity evidence.'],
]);

if (!fs.readFileSync(path.join(root, 'User Inventory Evidence.tmdl'), 'utf8').includes("measure 'Unclassified Workforce Accounts'")) {
  update('User Inventory Evidence', [
    ["\tmeasure 'Workforce MFA Registration Covered'", `\tmeasure 'Unclassified Workforce Accounts' = CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Persona Match State] IN { "Missing job title", "Unmatched" } )\n\t\tformatString: #,0\n\t\tdisplayFolder: Workforce | Personas\n\n\tmeasure 'Workforce MFA Registration Covered'`],
  ]);
}

update('Security Control Evidence', [
  ['/// Share of observed users reported as MFA capable by Microsoft Graph authentication-method registration details.', '/// Share of covered enabled workforce accounts with an MFA method registered in Entra. Conditional Access enforcement is a separate control.'],
  ["measure 'MFA Capability (%)' = CALCULATE(MAX('Security Control Evidence'[Health Rate]), KEEPFILTERS('Security Control Evidence'[Control Name] = \"MFA capability\"))", "measure 'Workforce MFA Registered (%)' = CALCULATE(MAX('Security Control Evidence'[Health Rate]), KEEPFILTERS('Security Control Evidence'[Control Name] = \"MFA registered among workforce\"))"],
  ['/// Observed users not reported as MFA capable.', '/// Covered enabled workforce accounts without a registered MFA method.'],
  ["measure 'Users Not MFA Capable (#)' = CALCULATE(MAX('Security Control Evidence'[Affected Entities]), KEEPFILTERS('Security Control Evidence'[Control Name] = \"MFA capability\"))", "measure 'Workforce Without MFA Registration (#)' = CALCULATE(MAX('Security Control Evidence'[Affected Entities]), KEEPFILTERS('Security Control Evidence'[Control Name] = \"MFA registered among workforce\"))"],
]);

update('Enterprise Operational Signals', [
  ['14, [Users Not MFA Capable (#)]', '14, [Workforce Without MFA Registration (#)]'],
  ['"Users not MFA capable", "Identity protection", "Observed", "Complete MFA registration and remediate users that are not MFA capable."',
    '"Workforce without MFA registration", "Identity protection", "Observed", "Complete MFA method registration for covered workforce accounts; assess Conditional Access enforcement separately."'],
]);
