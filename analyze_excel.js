const XLSX = require('xlsx');

const filePath = 'C:/Users/I755599/Downloads/updated atc_classified_202608100959429.xlsx';
const wb = XLSX.readFile(filePath);
const ws = wb.Sheets[wb.SheetNames[0]];

const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

console.log("=== SHEET STRUCTURE ===");
console.log("Total rows:", rows.length);
console.log("\nRow 0:", rows[0].slice(0, 16));
console.log("\nRow 1:", rows[1].slice(0, 16));
console.log("\nRow 2:", rows[2].slice(0, 16));
console.log("\nRow 3:", rows[3].slice(0, 16));

const headerRow = rows[1];
console.log("\n=== HEADERS (from Row 1) ===");
headerRow.forEach((h, i) => {
  if (h) console.log(`  ${i}: ${h}`);
});

const dataRows = rows.slice(3).map(row => {
  const obj = {};
  headerRow.forEach((header, idx) => {
    if (header) obj[header] = row[idx];
  });
  return obj;
});

console.log("\n=== DATA INFO ===");
console.log(`Total data rows: ${dataRows.length}`);
console.log("\nFirst data row:");
console.log(JSON.stringify(dataRows[0], null, 2));

console.log("\n=== LOOKING FOR KUNDAN AND REMEDIATION TYPE COLUMNS ===");
const kundalIndex = headerRow.findIndex(h => h && h.toString().toLowerCase() === 'kundan');
const remediationIndex = headerRow.findIndex(h => h && h.toString().toLowerCase().includes('remediation type'));

console.log(`Found "Kundan" at index: ${kundalIndex}`);
console.log(`Found "Remediation Type" at index: ${remediationIndex}`);

console.log("\n=== ALL HEADERS MATCHING SEARCH TERMS ===");
headerRow.forEach((h, i) => {
  const lower = h ? h.toString().toLowerCase() : '';
  if (lower.includes('kundan') || lower.includes('remediation') || lower.includes('remediation type')) {
    console.log(`  Index ${i}: "${h}"`);
  }
});

if (kundalIndex >= 0 && remediationIndex >= 0) {
  console.log("\n=== FINDING MISMATCHES ===");
  const mismatchRows = [];
  const mismatchCombinations = {};

  const kundalHeader = headerRow[kundalIndex];
  const remediationHeader = headerRow[remediationIndex];

  dataRows.forEach((row, idx) => {
    const kundan = (row[kundalHeader] || '').toString().trim();
    const remediation = (row[remediationHeader] || '').toString().trim();
    
    if (kundan && remediation && kundan !== remediation) {
      mismatchRows.push(row);
      const combo = `Kundan="${kundan}" vs Remediation Type="${remediation}"`;
      mismatchCombinations[combo] = (mismatchCombinations[combo] || 0) + 1;
    }
  });

  console.log(`Total mismatches found: ${mismatchRows.length} out of ${dataRows.length} rows`);

  console.log("\n=== MISMATCH COMBINATIONS (Top 20) ===");
  Object.entries(mismatchCombinations)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 20)
    .forEach(([combo, count]) => {
      console.log(`  ${combo}: ${count} rows`);
    });

  console.log(`\n=== SAMPLE ROWS WITH MISMATCHES (first 50 of ${mismatchRows.length}) ===`);
  const requiredColumns = [
    'Object name', 
    'Obj.',
    'Check Title', 
    'HCA/S4H?', 
    'Clone?', 
    'Syntax Error?', 
    'Fit Gap?', 
    kundalHeader,
    remediationHeader
  ];

  console.log(JSON.stringify(
    mismatchRows.slice(0, 50).map(row => {
      const result = {};
      requiredColumns.forEach(col => {
        if (col) result[col] = row[col] || '';
      });
      return result;
    }), 
    null, 2
  ));

  console.log(`\n=== STATISTICS ===`);
  console.log(`Total rows in sheet: ${dataRows.length}`);
  console.log(`Rows with mismatches: ${mismatchRows.length}`);
  console.log(`Mismatch percentage: ${((mismatchRows.length / dataRows.length) * 100).toFixed(2)}%`);
} else {
  console.log("\n=== ERROR: Could not find both columns ===");
}
