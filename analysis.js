const XLSX = require('xlsx');
const path = 'C:\\Users\\I755599\\Downloads\\updated atc_classified_202608100959429.xlsx';
const wb = XLSX.readFile(path, { cellText: false, raw: true });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

// Headers are in row 1 (index 1), data starts from row 2 (index 2)
const headers = aoa[1];
const dataRows = aoa.slice(2);

// Convert to objects
const rows = dataRows.map(row => {
  const obj = {};
  headers.forEach((header, idx) => {
    obj[header] = row[idx] || '';
  });
  return obj;
});

console.log('Total data rows:', rows.length);
console.log('\n=== FIRST 3 DATA ROWS ===');
console.log(JSON.stringify(rows[0], null, 2));

// Check for mismatches: Kundan='FIT GAP' but Remediation Type != 'FIT GAP'
const mismatches = rows.filter(r => r['Kundan'] === 'FIT GAP' && r['Remediation Type'] && r['Remediation Type'] !== 'FIT GAP');
console.log('\n\n=== MISMATCH OVERVIEW ===');
console.log('Total FIT GAP mismatches:', mismatches.length);

// Analysis 2: Count mismatches with Fit Gap? = Yes vs No
const mismatchYes = mismatches.filter(r => r['Fit Gap?'] === 'Yes').length;
const mismatchNo = mismatches.filter(r => r['Fit Gap?'] === 'No').length;
console.log('Mismatches with Fit Gap?=Yes:', mismatchYes);
console.log('Mismatches with Fit Gap?=No:', mismatchNo);

// Get 10 object names with mismatches
const mismatchObjectNames = [...new Set(mismatches.map(r => r['Object name']))].slice(0, 10);
console.log('\n=== ANALYSIS 1: 10 Sample Object Names ===');
mismatchObjectNames.forEach((name, idx) => console.log(`${idx + 1}. ${name}`));

// For each of the 10 objects, get ALL rows
console.log('\n=== ANALYSIS 1 DETAILED: ALL ROWS FOR 10 SAMPLE OBJECTS ===');
mismatchObjectNames.forEach(objName => {
  const objRows = rows.filter(r => r['Object name'] === objName);
  console.log(`\n[${objName}] - Total rows: ${objRows.length}`);
  objRows.forEach(r => {
    console.log(`  Obj=${r['Obj.']} | Check=${r['Check Title']?.substring(0,40)} | HCA/S4H=${r['HCA/S4H?']} | Kundan=${r['Kundan']} | RemType=${r['Remediation Type']} | FitGap?=${r['Fit Gap?']}`);
  });
});

// Analysis 3 & 4: Rows where Kundan='FIT GAP' AND Fit Gap?='No' (propagated candidates)
const propagated = rows.filter(r => r['Kundan'] === 'FIT GAP' && r['Fit Gap?'] === 'No');
console.log('\n\n=== ANALYSIS 3 & 4: PROPAGATED ROWS (Kundan=FIT GAP AND Fit Gap?=No) ===');
console.log('Count of propagated rows:', propagated.length);

// Get 5 example objects from propagated rows
const propagatedObjectNames = [...new Set(propagated.map(r => r['Object name']))].slice(0, 5);
console.log('\n5 example propagated objects:');
propagatedObjectNames.forEach((name, idx) => console.log(`${idx + 1}. ${name}`));

console.log('\n=== 5 EXAMPLE OBJECTS WITH ALL ROWS (propagated + original) ===');
propagatedObjectNames.forEach(objName => {
  const objRows = rows.filter(r => r['Object name'] === objName);
  console.log(`\n[${objName}] - Total rows: ${objRows.length}`);
  objRows.forEach(r => {
    const isProp = r['Kundan'] === 'FIT GAP' && r['Fit Gap?'] === 'No' ? '[PROPAGATED]' : '';
    console.log(`  ${isProp} Obj=${r['Obj.']} | Check=${r['Check Title']?.substring(0,40)} | Kundan=${r['Kundan']} | RemType=${r['Remediation Type']} | FitGap?=${r['Fit Gap?']}`);
  });
});

// Analysis 4: Sibling detection
console.log('\n\n=== ANALYSIS 4: SIBLING DETECTION ===');
let siblingCount = 0;
const uniqueObjs = new Set();

propagated.forEach(r => {
  const key = r['Object name'] + ':::' + r['Obj.'];
  uniqueObjs.add(key);
  
  const siblings = rows.filter(sibling => 
    sibling['Object name'] === r['Object name'] && 
    sibling['Obj.'] === r['Obj.'] &&
    sibling['Fit Gap?'] === 'Yes'
  );
  
  if (siblings.length > 0) {
    siblingCount++;
  }
});

const uniqueObjList = Array.from(uniqueObjs);
let withSiblings = 0;
let withoutSiblings = 0;
const anomalies = [];

uniqueObjList.forEach(key => {
  const [objName, obj] = key.split(':::');
  const siblings = rows.filter(sibling => 
    sibling['Object name'] === objName && 
    sibling['Obj.'] === obj &&
    sibling['Fit Gap?'] === 'Yes'
  );
  
  if (siblings.length > 0) {
    withSiblings++;
  } else {
    withoutSiblings++;
    if (anomalies.length < 5) {
      anomalies.push({ key, objRows: rows.filter(r => r['Object name'] === objName && r['Obj.'] === obj) });
    }
  }
});

console.log(`\nTotal propagated rows: ${propagated.length}`);
console.log(`Rows with a sibling (Fit Gap?=Yes): ${siblingCount}`);
console.log(`Percentage: ${((siblingCount / propagated.length) * 100).toFixed(2)}%`);

console.log(`\nUnique Object+Obj combinations with propagated rows: ${uniqueObjList.length}`);
console.log(`  - With sibling (Fit Gap?=Yes): ${withSiblings}`);
console.log(`  - Without sibling (anomaly): ${withoutSiblings}`);

if (anomalies.length > 0) {
  console.log('\n=== ANOMALY EXAMPLES: Objects with propagated rows but NO sibling with Fit Gap?=Yes ===');
  anomalies.forEach(({ key, objRows }, idx) => {
    const [objName, obj] = key.split(':::');
    console.log(`\n${idx + 1}. ${key} (${objRows.length} total rows)`);
    objRows.forEach(r => {
      console.log(`   Kundan=${r['Kundan']} | RemType=${r['Remediation Type']} | FitGap?=${r['Fit Gap?']} | Check=${r['Check Title']?.substring(0,35)}`);
    });
  });
}
