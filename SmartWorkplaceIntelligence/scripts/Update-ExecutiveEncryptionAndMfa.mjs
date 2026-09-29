// One-time, guarded local PBIP migration. No private data or publication.
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..', 'pbip');
const model = path.join(root, 'SmartWorkplaceIntelligence.SemanticModel', 'definition', 'tables', 'Device Inventory Evidence.tmdl');
const page = path.join(root, 'SmartWorkplaceIntelligence.Report', 'definition', 'pages', '74096d393d05fa994d48', 'visuals');
const file = id => path.join(page, id, 'visual.json');
const data = new Map();
const load = id => { const v = JSON.parse(fs.readFileSync(file(id), 'utf8')); data.set(id, v); return v; };
const projection = (measure, label) => ({
  field: { Measure: { Expression: { SourceRef: { Entity: 'Device Inventory Evidence' } }, Property: measure } },
  queryRef: `Device Inventory Evidence.${measure}`,
  nativeQueryRef: label,
});
const setFilter = (v, measure) => { if (v.filterConfig?.filters?.[0]?.field?.Measure) v.filterConfig.filters[0].field.Measure.Property = measure; };
const setAlt = (v, description) => { v.visual.visualContainerObjects.general[0].properties.altText.expr.Literal.Value = `'${description}'`; };

const tmdl = fs.readFileSync(model, 'utf8');
const marker = "\tmeasure 'Known Intune Encryption State Coverage (%)' = DIVIDE ( CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Intune Encryption State] IN { \"Encrypted\", \"Not encrypted\" } ), [Managed Devices] )";
if (!tmdl.includes(marker) || tmdl.includes("measure 'Recent Intune Encryption Evidence Coverage (%)'")) throw new Error('TMDL precondition failed');
const nextTmdl = tmdl.replace(marker, `	/// Intune devices with a known encryption state and a last sync within the trailing 30 calendar days.
	measure 'Devices With Recent Intune Encryption Evidence' = CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Intune Encryption State] IN { "Encrypted", "Not encrypted" }, 'Device Inventory Evidence'[Last Sync DateTime] >= TODAY() - 30, 'Device Inventory Evidence'[Last Sync DateTime] < TODAY() + 1 )
		formatString: #,0
		displayFolder: Device | Security

	/// Recently synced Intune devices marked Encrypted; not proof of BitLocker recovery-key escrow.
	measure 'Recently Synced Devices Encrypted According to Intune' = CALCULATE ( [Managed Devices], 'Device Inventory Evidence'[Intune Encryption State] = "Encrypted", 'Device Inventory Evidence'[Last Sync DateTime] >= TODAY() - 30, 'Device Inventory Evidence'[Last Sync DateTime] < TODAY() + 1 )
		formatString: #,0
		displayFolder: Device | Security

	/// Share of all Intune devices with a known encryption state updated within the trailing 30 calendar days.
	measure 'Recent Intune Encryption Evidence Coverage (%)' = DIVIDE ( [Devices With Recent Intune Encryption Evidence], [Managed Devices] )
		formatString: 0.0%
		displayFolder: Device | Security

	/// Share encrypted among Intune devices with a known encryption state updated within the trailing 30 calendar days.
	measure 'Encrypted Recently Synced Intune Devices (%)' = DIVIDE ( [Recently Synced Devices Encrypted According to Intune], [Devices With Recent Intune Encryption Evidence] )
		formatString: 0.0%
		displayFolder: Device | Security

${marker}`);

const security = load('415ba06e1fecc76819c7');
const securityMetrics = security.visual.query.queryState.Data.projections;
if (securityMetrics.length !== 3 || securityMetrics[1].field.Measure.Property !== 'Workforce MFA Registered (%)') throw new Error('MFA visual precondition failed');
securityMetrics.splice(1, 1);
setAlt(security, 'Executive security posture: Microsoft Secure Score and active Defender protection.');

const encrypted = load('12365a8234c915ca9682');
if (encrypted.visual.query.queryState.Data.projections[0].field.Measure.Property !== 'Encrypted Devices Among Known Intune State (%)') throw new Error('Encryption visual precondition failed');
encrypted.visual.query.queryState.Data.projections = [projection('Encrypted Recently Synced Intune Devices (%)', 'Encrypted devices (%)')];
encrypted.visual.query.queryState.Tooltips.projections = [
  projection('Known Intune Encryption State Coverage (%)', 'Known-state coverage (%)'),
  projection('Devices With Recent Intune Encryption Evidence', 'Recently synced devices with known state'),
  projection('Recently Synced Devices Encrypted According to Intune', 'Recently synced encrypted devices'),
];
setFilter(encrypted, 'Encrypted Recently Synced Intune Devices (%)');
setAlt(encrypted, 'Share marked encrypted among Intune devices with a known state and sync within 30 days. Excludes stale records and does not prove BitLocker key escrow.');

const coverage = load('6ff6305cd8a44c2280cf');
if (coverage.visual.query.queryState.Data.projections[0].field.Measure.Property !== 'Known Intune Encryption State Coverage (%)') throw new Error('Coverage visual precondition failed');
coverage.visual.query.queryState.Data.projections = [projection('Recent Intune Encryption Evidence Coverage (%)', 'Recent encryption evidence (%)')];
coverage.visual.query.queryState.Tooltips.projections = [
  projection('Known Intune Encryption State Coverage (%)', 'Technical known-state coverage (%)'),
  projection('Devices With Recent Intune Encryption Evidence', 'Recently synced devices with known state'),
];
setFilter(coverage, 'Recent Intune Encryption Evidence Coverage (%)');
coverage.visual.visualContainerObjects.title[0].properties.text.expr.Literal.Value = "'Encryption evidence updated within 30 days'";
setAlt(coverage, 'Share of all managed Intune devices with a known encryption state and sync within 30 days. Technical known-state completeness remains in the tooltip.');

// All preconditions are checked before the first write.
fs.writeFileSync(model, nextTmdl, 'utf8');
for (const [id, v] of data) fs.writeFileSync(file(id), JSON.stringify(v, null, 2) + '\n', 'utf8');
console.log('Updated Executive Overview encryption evidence and removed Workforce MFA card.');
