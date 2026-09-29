// Narrow, idempotent TMDL migration for the approved site/persona and encryption measures.
// The report source files are OneDrive-backed; do not rewrite unrelated model objects.
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.SemanticModel', 'definition', 'tables');
function edit(name, changes) {
  const file = path.join(root, name);
  let source = fs.readFileSync(file, 'utf8');
  const nl = source.includes('\r\n') ? '\r\n' : '\n';
  for (const [anchor, replacement, already] of changes) {
    if (already && source.includes(already)) continue;
    if (!source.includes(anchor)) throw new Error(`Missing TMDL anchor in ${name}: ${anchor.slice(0, 65)}`);
    source = source.replace(anchor, replacement.replaceAll('\n', nl));
  }
  fs.writeFileSync(file, source, 'utf8');
}

edit('User Inventory Evidence.tmdl', [
  ["\tcolumn 'User Source ID'", `\t/// Share of enabled workforce accounts with a governed AD site code matched to the private site mapping.
\tmeasure 'Workforce Site Classification Coverage (%)' = DIVIDE ( CALCULATE ( [Enabled Workforce Accounts], 'User Inventory Evidence'[Site Match State] = "Mapped" ), [Enabled Workforce Accounts] )
\t\tformatString: 0.0%
\t\tdisplayFolder: Workforce | Personas | Coverage

\tcolumn 'User Source ID'`, "measure 'Workforce Site Classification Coverage (%)'"],
  ["\tcolumn 'AD Last Activity Date'", `\t/// Site code from the configured AD user attribute, preserved as text including leading zeroes.
\tcolumn 'Directory Site Code'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Directory Site Code

\t/// Mutually exclusive site category from the private governed site mapping.
\tcolumn 'Site Type'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Site Type

\tcolumn 'Site Match State'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Site Match State

\tcolumn 'AD Last Activity Date'`, "column 'Directory Site Code'"],
  ["\tcolumn 'MFA Registration State'", `\t/// Provenance of the final persona decision: site, job title, or AD description fallback.
\tcolumn 'Persona Classification Source'
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: Persona Classification Source

\tcolumn 'MFA Registration State'`, "column 'Persona Classification Source'"],
  ['Columns=46, Encoding=65001', 'Columns=50, Encoding=65001', 'Columns=50, Encoding=65001'],
  ['{"Workforce Persona", type text},', `{"Directory Site Code", type text},
\t\t\t\t        {"Site Type", type text},
\t\t\t\t        {"Site Match State", type text},
\t\t\t\t        {"Workforce Persona", type text},`, '{"Site Type", type text}'],
  ['{"Persona Matched Keywords", type text},', `{"Persona Matched Keywords", type text},
\t\t\t\t        {"Persona Classification Source", type text},`, '{"Persona Classification Source", type text}'],
]);

edit('Device Inventory Evidence.tmdl', [
  ["\tcolumn 'Device Source ID'", `\t/// Share of managed Intune devices with a known encryption state that are marked Encrypted. This does not prove BitLocker key escrow.
\tmeasure 'Encrypted Devices Among Known Intune State (%)' = DIVIDE ( [Devices Encrypted According to Intune], CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Intune Encryption State] IN { "Encrypted", "Not encrypted" } ) )
\t\tformatString: 0.0%
\t\tdisplayFolder: Device | Security

\t/// Share of managed Intune devices for which the collected encryption state is known.
\tmeasure 'Known Intune Encryption State Coverage (%)' = DIVIDE ( CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Intune Encryption State] IN { "Encrypted", "Not encrypted" } ), [Managed Devices] )
\t\tformatString: 0.0%
\t\tdisplayFolder: Device | Security

\tcolumn 'Device Source ID'`, "measure 'Encrypted Devices Among Known Intune State (%)'"],
]);
process.stdout.write('Site/persona columns and Intune encryption measures updated.\n');
