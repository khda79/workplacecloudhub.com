'use strict';

const fs = require('fs');
const path = require('path');

const reportRoot = path.resolve(__dirname, '..', 'pbip', 'SmartWorkplaceIntelligence.Report', 'definition', 'pages');
const riskPageId = 'c801a6e5a6a2683bc55e';

let updatedFiles = 0;

function visit(directory) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      visit(fullPath);
      continue;
    }
    if (entry.name !== 'visual.json') continue;

    const original = fs.readFileSync(fullPath, 'utf8');
    let updated = original
      .replaceAll('Supporting indicators', 'Supporting KPIs')
      .replaceAll('Supporting Indicators', 'Supporting KPIs');

    if (fullPath.includes(`${path.sep}${riskPageId}${path.sep}`)) {
      updated = updated
        .replaceAll('Affected entities by domain', 'Affected signal observations by domain')
        .replaceAll('"nativeQueryRef": "Enterprise Signal Affected"', '"nativeQueryRef": "Affected signal observations"')
        .replaceAll('"nativeQueryRef": "Affected entities"', '"nativeQueryRef": "Affected signal observations"');
    }

    if (updated.includes("'Mailboxes by Exchange version'")) {
      updated = updated
        .replaceAll('"Entity": "Mailbox Evidence"', '"Entity": "Mailbox Exchange Version Catalog"')
        .replaceAll('"Property": "Mailboxes"', '"Property": "Mailboxes by Exchange Version"')
        .replaceAll('"queryRef": "Mailbox Evidence.Exchange Version"', '"queryRef": "Mailbox Exchange Version Catalog.Exchange Version"')
        .replaceAll('"queryRef": "Mailbox Evidence.Mailboxes"', '"queryRef": "Mailbox Exchange Version Catalog.Mailboxes by Exchange Version"')
        .replaceAll('"nativeQueryRef": "Mailboxes"', '"nativeQueryRef": "Mailboxes by Exchange Version"');
    }

    if (updated !== original) {
      fs.writeFileSync(fullPath, updated, 'utf8');
      updatedFiles += 1;
    }
  }
}

visit(reportRoot);
process.stdout.write(`Updated ${updatedFiles} visual definition files.\n`);
