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

console.log('=== CROSS-OBJECT PROPAGATION ANALYSIS ===\n');

// Get objects with propagated rows (Kundan=FIT GAP, Fit Gap?=No)
const propagated = rows.filter(r => r['Kundan'] === 'FIT GAP' && r['Fit Gap?'] === 'No');
const propagatedObjNames = [...new Set(propagated.map(r => r['Object name']))];

console.log(`Total propagated object names: ${propagatedObjNames.length}\n`);
console.log('HYPOTHESIS: Maybe these objects exist with DIFFERENT Obj. types');
console.log('and the Fit Gap is being propagated across object type variants?\n');

let foundCrossType = 0;
const examples = [];

propagatedObjNames.forEach(objName => {
  // Find all occurrences of this object name with ANY type
  const allVariants = rows.filter(r => r['Object name'] === objName);
  const objTypes = new Set(allVariants.map(r => r['Obj.']));
  
  if (objTypes.size > 1) {
    foundCrossType++;
    if (examples.length < 5) {
      console.log(`EXAMPLE ${examples.length + 1}: ${objName}`);
      console.log(`  Object types: ${Array.from(objTypes).join(', ')}`);
      
      objTypes.forEach(type => {
        const rowsOfType = allVariants.filter(r => r['Obj.'] === type);
        const fitGapYes = rowsOfType.filter(r => r['Fit Gap?'] === 'Yes').length;
        const fitGapNo = rowsOfType.filter(r => r['Fit Gap?'] === 'No').length;
        const kundan_fitgap = rowsOfType.filter(r => r['Kundan'] === 'FIT GAP').length;
        console.log(`    ${type}: ${rowsOfType.length} rows | Fit Gap?=Yes:${fitGapYes} | Kundan=FIT GAP:${kundan_fitgap}`);
      });
      console.log();
      examples.push(objName);
    }
  }
});

console.log(`\nObjects with multiple type variants: ${foundCrossType}`);

// Now check the bigger picture: are there objects where SOME variants have Fit Gap=Yes
console.log('\n\n=== CROSS-TYPE PROPAGATION PATTERN ===\n');

let crossTypeWithFitGap = 0;
let crossTypeAll = 0;

propagatedObjNames.forEach(objName => {
  const allVariants = rows.filter(r => r['Object name'] === objName);
  const objTypes = new Set(allVariants.map(r => r['Obj.']));
  
  if (objTypes.size > 1) {
    crossTypeAll++;
    
    // Check if ANY variant has Fit Gap?=Yes
    const anyFitGapYes = allVariants.some(r => r['Fit Gap?'] === 'Yes');
    if (anyFitGapYes) {
      crossTypeWithFitGap++;
    }
  }
});

console.log(`Objects with multiple type variants: ${crossTypeAll}`);
console.log(`  - With at least one variant having Fit Gap?=Yes: ${crossTypeWithFitGap}`);
console.log(`  - Percentage: ${((crossTypeWithFitGap / crossTypeAll) * 100).toFixed(1)}%`);

if (crossTypeWithFitGap === 0) {
  console.log('\n>>> CRITICAL FINDING: NONE of the multi-type objects have ANY Fit Gap?=Yes rows!');
  console.log('>>> This means the propagation is NOT coming from cross-type variants either.');
}

console.log('\n\n=== NEW HYPOTHESIS: Propagation by Package or Document ===\n');

// Check if objects are grouped by package and one package has Fit Gap
const propagatedByPackage = {};
propagated.slice(0, 100).forEach(r => {
  const pkg = r['Package'] || 'UNKNOWN';
  if (!propagatedByPackage[pkg]) {
    propagatedByPackage[pkg] = 0;
  }
  propagatedByPackage[pkg]++;
});

console.log('Sample packages with propagated rows:');
Object.entries(propagatedByPackage).slice(0, 10).forEach(([pkg, count]) => {
  console.log(`  ${pkg}: ${count} propagated rows`);
});
