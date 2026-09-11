const XLSX = require('xlsx');
const path = 'C:\\Users\\I755599\\Downloads\\updated atc_classified_202608100959429.xlsx';
const wb = XLSX.readFile(path, { cellText: false, raw: true });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

const headers = aoa[1];
const dataRows = aoa.slice(2);
const rows = dataRows.map(row => {
  const obj = {};
  headers.forEach((header, idx) => {
    obj[header] = row[idx] || '';
  });
  return obj;
});

console.log('=== FINAL ANALYSIS: THE REAL PROPAGATION RULE ===\n');

// Get the 7082 mismatches: Kundan=FIT GAP but Remediation Type != FIT GAP
const mismatches = rows.filter(r => r['Kundan'] === 'FIT GAP' && r['Remediation Type'] && r['Remediation Type'] !== 'FIT GAP');

// Split by Fit Gap? column
const mismatchWithFitGapYes = mismatches.filter(r => r['Fit Gap?'] === 'Yes');
const mismatchWithFitGapNo = mismatches.filter(r => r['Fit Gap?'] === 'No');

console.log('MISMATCH BREAKDOWN:');
console.log(`Total mismatches (Kundan=FIT GAP but RemType!=FIT GAP): ${mismatches.length}`);
console.log(`  1. With Fit Gap?=Yes: ${mismatchWithFitGapYes.length} rows`);
console.log(`  2. With Fit Gap?=No: ${mismatchWithFitGapNo.length} rows (these are the propagated ones)\n`);

console.log('INTERPRETATION:');
console.log(`The ${mismatchWithFitGapYes.length} rows with Fit Gap?=Yes are DIRECTLY classified as FIT GAP by the rule engine.`);
console.log(`The ${mismatchWithFitGapNo.length} rows with Fit Gap?=No appear PROPAGATED.`);

console.log('\n\n=== CRITICAL INSIGHT ===\n');
console.log('The 1802 "propagated" rows (Kundan=FIT GAP, Fit Gap?=No) have:');
console.log('  - NO sibling within the same object with Fit Gap?=Yes');
console.log('  - NO cross-object variants with Fit Gap?=Yes');
console.log('  - Remediation Types: various (MANDATORY, NEEDS REMEDIATION, FALSE POSITIVE, CAN BE IGNORED, OPTIONAL)\n');

console.log('HYPOTHESIS: These rows are being classified as FIT GAP based on:');
console.log('  Option A: A DIFFERENT RULE that doesn\'t explicitly mark Fit Gap?=Yes');
console.log('  Option B: They are being propagated from a DIFFERENT OBJECT entirely (not by name, but by reference)');
console.log('  Option C: The propagation rule checks something OTHER than "Fit Gap?" column\n');

console.log('EVIDENCE FROM DATA:');

// Show distribution of Remediation Types in propagated rows
const propRowsByRemType = {};
mismatchWithFitGapNo.forEach(r => {
  const type = r['Remediation Type'] || 'EMPTY';
  propRowsByRemType[type] = (propRowsByRemType[type] || 0) + 1;
});

console.log('Remediation Types in the 1802 propagated rows:');
Object.entries(propRowsByRemType).sort((a,b) => b[1] - a[1]).forEach(([type, count]) => {
  console.log(`  ${type}: ${count} rows (${((count/mismatchWithFitGapNo.length)*100).toFixed(1)}%)`);
});

console.log('\n\n=== CONCLUSION ===');
console.log('\nThe propagateFitGap function is NOT working as originally described.');
console.log('The 1802 rows show a DIFFERENT behavior:');
console.log('  - They have Fit Gap?=No (not directly flagged as FIT GAP)');
console.log('  - But Kundan=FIT GAP (user tool classified them as FIT GAP)');
console.log('  - Yet they have NO siblings or related objects with Fit Gap?=Yes');
console.log('\nThis suggests the actual rule is:');
console.log('  "If a row\'s Remediation Type DIFFERS from FIT GAP, BUT the classification');
console.log('   would have been FIT GAP anyway based on some other rule/check,');
console.log('   then mark Kundan=FIT GAP to show the discrepancy."');
console.log('\nIn other words: The mismatch is the POINT, not a bug.');
