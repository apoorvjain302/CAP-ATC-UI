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

console.log('=== DEEP DIVE: UNDERSTANDING THE PROPAGATION MECHANISM ===\n');

// Get propagated rows
const propagated = rows.filter(r => r['Kundan'] === 'FIT GAP' && r['Fit Gap?'] === 'No');

// Check what is common about these objects
console.log('1. PATTERN ANALYSIS: What makes an object get FIT GAP propagation?\n');

// Look at first 10 propagated objects
const uniqueObjs = new Set();
propagated.forEach(r => uniqueObjs.add(r['Object name'] + ':::' + r['Obj.']));
const objList = Array.from(uniqueObjs).slice(0, 10);

objList.forEach((key, idx) => {
  const [objName, objType] = key.split(':::');
  const allRowsForObj = rows.filter(r => r['Object name'] === objName && r['Obj.'] === objType);
  
  const fitGapYesRows = allRowsForObj.filter(r => r['Fit Gap?'] === 'Yes');
  const fitGapNoRows = allRowsForObj.filter(r => r['Fit Gap?'] === 'No');
  const kundan_fitgap = allRowsForObj.filter(r => r['Kundan'] === 'FIT GAP');
  const kundan_fitgap_no = allRowsForObj.filter(r => r['Kundan'] === 'FIT GAP' && r['Fit Gap?'] === 'No');
  
  console.log(`\n${idx + 1}. ${objName} (${objType})`);
  console.log(`   Total rows: ${allRowsForObj.length}`);
  console.log(`   Fit Gap?=Yes: ${fitGapYesRows.length} rows`);
  console.log(`   Fit Gap?=No: ${fitGapNoRows.length} rows`);
  console.log(`   Kundan=FIT GAP: ${kundan_fitgap.length} rows`);
  console.log(`   Kundan=FIT GAP + Fit Gap?=No (PROPAGATED): ${kundan_fitgap_no.length} rows`);
  
  // Check if any row has Fit Gap?=Yes but Kundan!=FIT GAP
  const directFitGap = allRowsForObj.filter(r => r['Fit Gap?'] === 'Yes');
  const directButNotKundan = directFitGap.filter(r => r['Kundan'] !== 'FIT GAP');
  
  if (directButNotKundan.length > 0) {
    console.log(`   >>> FOUND: ${directButNotKundan.length} rows with Fit Gap?=Yes but Kundan!=FIT GAP`);
    directButNotKundan.slice(0, 2).forEach(r => {
      console.log(`       - Kundan=${r['Kundan']}, RemType=${r['Remediation Type']}`);
    });
  }
});

console.log('\n\n2. HYPOTHESIS: The propagation rule might be different');
console.log('   Maybe propagation is based on "Fit Gap?" column, not "Kundan"?');
console.log('   Or maybe there\'s a cross-object relationship (different Obj. types)?\n');

// Check if objects with mixed Obj. types
console.log('3. CHECKING OBJECT TYPE MIX:\n');
propagated.slice(0, 100).forEach((row, idx) => {
  if (idx % 20 === 0) {
    const objName = row['Object name'];
    const allWithSameName = rows.filter(r => r['Object name'] === objName);
    const objTypes = new Set(allWithSameName.map(r => r['Obj.']));
    if (objTypes.size > 1) {
      console.log(`   ${objName}: MIXED OBJECT TYPES: ${Array.from(objTypes).join(', ')}`);
    }
  }
});

console.log('\n\n4. CRITICAL QUESTION: Are the propagated rows in DOCUMENTS with multiple Obj. types?');
const propagatedObjNames = [...new Set(propagated.map(r => r['Object name']))];
let multiTypeCount = 0;
propagatedObjNames.slice(0, 50).forEach(objName => {
  const allRows = rows.filter(r => r['Object name'] === objName);
  const objTypes = new Set(allRows.map(r => r['Obj.']));
  if (objTypes.size > 1) {
    multiTypeCount++;
    console.log(`   ${objName}: ${objTypes.size} object types - ${Array.from(objTypes).join(', ')}`);
  }
});
console.log(`\n   Result: ${multiTypeCount} out of 50 objects have multiple Obj. types`);
