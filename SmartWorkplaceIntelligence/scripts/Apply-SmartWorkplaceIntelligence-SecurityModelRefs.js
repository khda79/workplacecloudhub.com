#!/usr/bin/env node
'use strict';

// Idempotently registers the synthetic security tables and their conformed
// dimension relationships in the canonical TMDL definition.

const fs = require('fs');
const path = require('path');

const definitionRoot = path.resolve(__dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.SemanticModel', 'definition');
const modelPath = path.join(definitionRoot, 'model.tmdl');
const relationshipsPath = path.join(definitionRoot, 'relationships.tmdl');
const tableNames = ['Identity Security Status', 'Conditional Access Coverage', 'Endpoint Security Status'];

let model = fs.readFileSync(modelPath, 'utf8');
if (!model.includes('"Identity Security Status"')) {
  model = model.replace(
    '"User Directory Observations"]',
    '"User Directory Observations","Identity Security Status","Conditional Access Coverage","Endpoint Security Status"]'
  );
}
for (const tableName of tableNames) {
  const declaration = `ref table '${tableName}'`;
  if (!model.includes(declaration)) model = `${model.trimEnd()}\n${declaration}\n`;
}
fs.writeFileSync(modelPath, model, 'utf8');

const relationshipBlock = `
relationship rel_Date_Identity_Security_Status_Date_Key
\tfromColumn: 'Identity Security Status'.'Date Key'
\ttoColumn: Date.'Date Key'

relationship rel_Date_Conditional_Access_Coverage_Date_Key
\tfromColumn: 'Conditional Access Coverage'.'Date Key'
\ttoColumn: Date.'Date Key'

relationship rel_Date_Endpoint_Security_Status_Date_Key
\tfromColumn: 'Endpoint Security Status'.'Date Key'
\ttoColumn: Date.'Date Key'

relationship rel_Tenant_Identity_Security_Status_Tenant_Key
\tfromColumn: 'Identity Security Status'.'Tenant Key'
\ttoColumn: Tenant.'Tenant Key'

relationship rel_Tenant_Conditional_Access_Coverage_Tenant_Key
\tfromColumn: 'Conditional Access Coverage'.'Tenant Key'
\ttoColumn: Tenant.'Tenant Key'

relationship rel_Tenant_Endpoint_Security_Status_Tenant_Key
\tfromColumn: 'Endpoint Security Status'.'Tenant Key'
\ttoColumn: Tenant.'Tenant Key'

relationship rel_Environment_Identity_Security_Status_Environment_Key
\tfromColumn: 'Identity Security Status'.'Environment Key'
\ttoColumn: Environment.'Environment Key'

relationship rel_Environment_Conditional_Access_Coverage_Environment_Key
\tfromColumn: 'Conditional Access Coverage'.'Environment Key'
\ttoColumn: Environment.'Environment Key'

relationship rel_Environment_Endpoint_Security_Status_Environment_Key
\tfromColumn: 'Endpoint Security Status'.'Environment Key'
\ttoColumn: Environment.'Environment Key'

relationship rel_User_Identity_Security_Status_User_Key
\tfromColumn: 'Identity Security Status'.'User Key'
\ttoColumn: User.'User Key'

relationship rel_Device_Endpoint_Security_Status_Device_Key
\tfromColumn: 'Endpoint Security Status'.'Device Key'
\ttoColumn: Device.'Device Key'
`;

let relationships = fs.readFileSync(relationshipsPath, 'utf8');
if (!relationships.includes('relationship rel_Date_Identity_Security_Status_Date_Key')) {
  relationships = `${relationships.trimEnd()}\n\n${relationshipBlock.trim()}\n`;
  fs.writeFileSync(relationshipsPath, relationships, 'utf8');
}

process.stdout.write(JSON.stringify({ status: 'PASS', tablesRegistered: tableNames.length, relationshipsRegistered: 11 }));
