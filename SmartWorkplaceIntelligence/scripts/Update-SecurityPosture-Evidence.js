#!/usr/bin/env node
'use strict';

// Targeted PBIR migration for the Security Posture page. This intentionally
// updates only the visuals that consume Security Control Evidence so Desktop
// layout adjustments on other pages remain untouched.

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const visualRoot = path.join(
  root,
  'pbip',
  'SmartWorkplaceIntelligence.Report',
  'definition',
  'pages',
  'c823af8cf0f8d76eb252',
  'visuals'
);

const executiveVisualRoot = path.join(
  root,
  'pbip',
  'SmartWorkplaceIntelligence.Report',
  'definition',
  'pages',
  '74096d393d05fa994d48',
  'visuals'
);

function visualPath(id) {
  return path.join(visualRoot, id, 'visual.json');
}

function readVisual(id) {
  const target = visualPath(id);
  return { target, value: JSON.parse(fs.readFileSync(target, 'utf8')) };
}

function writeVisual(target, value) {
  fs.writeFileSync(target, `${JSON.stringify(value, null, 2)}\r\n`, 'utf8');
}

function measure(name, label = name) {
  return {
    field: {
      Measure: {
        Expression: { SourceRef: { Entity: 'Security Control Evidence' } },
        Property: name
      }
    },
    queryRef: `Security Control Evidence.${name}`,
    nativeQueryRef: label
  };
}

function column(name, label = name) {
  return {
    field: {
      Column: {
        Expression: { SourceRef: { Entity: 'Security Control Evidence' } },
        Property: name
      }
    },
    queryRef: `Security Control Evidence.${name}`,
    nativeQueryRef: label
  };
}

function setTitle(visual, title) {
  const titleObject = visual.visual.visualContainerObjects?.title?.[0];
  if (!titleObject) throw new Error(`Missing title object on visual ${visual.name}`);
  titleObject.properties.text.expr.Literal.Value = `'${title}'`;
}

function setAltText(visual, altText) {
  const general = visual.visual.visualContainerObjects?.general?.[0];
  if (general?.properties?.altText?.expr?.Literal) {
    general.properties.altText.expr.Literal.Value = `'${altText}'`;
  }
}

function updateCard(id, measures, altText) {
  const { target, value } = readVisual(id);
  if (value.visual.visualType !== 'cardVisual') throw new Error(`${id} is not a cardVisual`);
  value.visual.query.queryState.Data.projections = measures.map((name) => measure(name));
  setAltText(value, altText);
  writeVisual(target, value);
}

function updateHero() {
  const { target, value } = readVisual('6ead047915962471b5dc');
  if (value.visual.visualType !== 'barChart') throw new Error('Security control hero is not a barChart');
  value.visual.query.queryState.Category.projections = [column('Control Name')];
  value.visual.query.queryState.Y.projections = [measure('Observed Security Health (%)')];
  setTitle(value, 'Control health by security control');
  setAltText(value, 'Observed health percentage by security control.');
  writeVisual(target, value);
}

function updateDetail() {
  const { target, value } = readVisual('ac67320ec18a5ed67221');
  if (value.visual.visualType !== 'tableEx') throw new Error('Security Posture detail is not a tableEx');
  value.visual.query.queryState.Values.projections = [
    column('Control Name'),
    column('Security Domain'),
    column('Severity Label'),
    column('Evidence Status'),
    column('Covered Entities'),
    column('Affected Entities'),
    column('Metric Unit'),
    column('Health Rate'),
    column('Gap State'),
    column('Evidence Source'),
    column('Recommended Action')
  ];
  setAltText(value, 'Security Posture detail with observed coverage, unit, health rate, evidence source, and recommended action.');
  writeVisual(target, value);
}

function updateSubtitle() {
  const { target, value } = readVisual('179371460dfd1cee8be8');
  const textRun = value.visual.objects?.general?.[0]?.properties?.paragraphs?.[0]?.textRuns?.[0];
  if (!textRun) throw new Error('Security Posture subtitle text run was not found');
  textRun.value = 'Security Posture · BETA 1.0.0-beta.1 · current private device, identity, access, endpoint protection, and Secure Score evidence';
  writeVisual(target, value);
}

function updateTextBox(id, text) {
  const { target, value } = readVisual(id);
  const textRun = value.visual.objects?.general?.[0]?.properties?.paragraphs?.[0]?.textRuns?.[0];
  if (!textRun) throw new Error(`Text run was not found on ${id}`);
  textRun.value = text;
  writeVisual(target, value);
}

function updateSecondaryAttentionCards() {
  updateCard(
    '57137a0d5e39c627549c',
    ['Noncompliant Devices (#)', 'Devices Without Active Defender Protection (#)', 'Devices Without Enabled Firewall (#)'],
    'Endpoint attention counts for compliance, Defender protection, and firewall status.'
  );
  updateCard(
    'fd6fd2b2ac737b57de06',
    ['Users Not MFA Capable (#)', 'Conditional Access Policies Not Enforced (#)'],
    'Identity and access attention counts for MFA capability and Conditional Access enforcement.'
  );
  updateTextBox('53a9ddee934dbc706952', 'Endpoint attention');
  updateTextBox('48c392087553bbbaffe9', 'Identity and access attention');
}

function updateExecutiveSecurityCard() {
  const target = path.join(executiveVisualRoot, '415ba06e1fecc76819c7', 'visual.json');
  const value = JSON.parse(fs.readFileSync(target, 'utf8'));
  if (value.visual.visualType !== 'cardVisual') throw new Error('Executive security visual is not a cardVisual');
  value.visual.query.queryState.Data.projections = [
    measure('Microsoft Secure Score (%)'),
    measure('MFA Capability (%)'),
    measure('Defender Protection Active (%)')
  ];
  setTitle(value, 'Security posture');
  setAltText(value, 'Executive security posture: Microsoft Secure Score, MFA capability, and Defender protection active.');
  writeVisual(target, value);
}

function updateSecurityOperationalSignals() {
  const { target, value } = readVisual('6e290719f85de88595dc');
  const filter = value.filterConfig?.filters?.find(item => item.field?.Column?.Property === 'Signal');
  const inCondition = filter?.filter?.Where?.[0]?.Condition?.In;
  if (!inCondition) throw new Error('Security operational signal filter was not found');
  const signals = [
    'Secure Boot gaps',
    'Disk encryption gaps',
    'Code integrity gaps',
    'Active Directory health warnings or critical checks',
    'Entra Connect synchronization health issues',
    'Users not MFA capable',
    'Devices without active Defender protection',
    'Devices without enabled firewall',
    'Conditional Access policies not enforced'
  ];
  inCondition.Values = signals.map(signal => [{ Literal: { Value: `'${signal}'` } }]);
  setAltText(value, 'Observed security control gaps and recommended remediation actions.');
  writeVisual(target, value);
}

updateCard(
  '97237f763e831ad410ab',
  ['Microsoft Secure Score (%)', 'MFA Capability (%)', 'Defender Protection Active (%)'],
  'Three priority Security Posture KPIs: Secure Score, MFA capability, and Defender protection.'
);
updateCard(
  '90dd02b57d7360645ca4',
  ['Conditional Access Enforcement (%)', 'Firewall Enabled (%)', 'Device Compliance Healthy (%)', 'Secure Boot Healthy (%)', 'Disk Encryption Healthy (%)'],
  'Supporting Security Posture indicators for access, endpoint, compliance, Secure Boot, and disk encryption.'
);
updateHero();
updateDetail();
updateSubtitle();
updateSecondaryAttentionCards();
updateExecutiveSecurityCard();
updateSecurityOperationalSignals();

console.log('Security Posture evidence visuals updated.');
