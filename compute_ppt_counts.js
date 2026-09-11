const XLSX = require('./node_modules/xlsx');

const wb = XLSX.readFile('C:/Users/I755599/Downloads/test_Ankush_ATC_Results (5)/atc_classified_202608170523060.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

const hcas4h  = r => (r['HCA/S4H?'] || '').trim();
const remType = r => (r['Remediation Type'] || '').trim().toUpperCase();
const syn     = r => (r['Syntax Error?'] || '').trim() === 'Yes';
const fg      = r => (r['Fit Gap?'] || '').trim() === 'Yes';
const cln     = r => (r['Clone?'] || '').trim() === 'Yes';
const tp      = r => (r['Object name'] || '').startsWith('/');
const isHCA   = r => hcas4h(r) === 'HCA';
const isS4H   = r => hcas4h(r) === 'S/4H';

console.log('=== BASIC COUNTS ===');
console.log('1. Total rows:', rows.length);

console.log('\n2. HCA/S4H? unique values and counts:');
const hcas4hCounts = {};
rows.forEach(r => { const v = hcas4h(r); hcas4hCounts[v] = (hcas4hCounts[v]||0)+1; });
console.log(JSON.stringify(hcas4hCounts, null, 2));

const hcaCount = rows.filter(isHCA).length;
const s4Count  = rows.filter(isS4H).length;
console.log('3. hcaCount:', hcaCount);
console.log('4. s4Count:', s4Count);

const syntaxByRemType = rows.filter(r => isHCA(r) && remType(r) === 'SYNTAX ERROR').length;
const syntaxBySynFlag = rows.filter(r => isHCA(r) && syn(r)).length;
const syntaxImpacts   = rows.filter(r => isHCA(r) && (remType(r) === 'SYNTAX ERROR' || syn(r))).length;
console.log('5. syntaxImpacts:', syntaxImpacts);
console.log('   (sub) HCA AND RemType=SYNTAX ERROR:', syntaxByRemType);
console.log('   (sub) HCA AND Syntax Error?=Yes:', syntaxBySynFlag);

const tpImpacts = rows.filter(r => tp(r)).length;
console.log('6. tpImpacts (Object name starts with /):', tpImpacts);

console.log('\n=== SLIDE 6 BARS ===');
console.log('7. Slide6 bar2 (HCA Impact) = hcaCount - syntaxImpacts:', hcaCount - syntaxImpacts);
console.log('8. Slide6 bar3 (S/4HANA Impact):', s4Count);
console.log('9. Slide6 bar4 (3rd Party):', tpImpacts);

const hcaMandatory = rows.filter(r => isHCA(r) && !syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r) === 'MANDATORY').length;
console.log('\n10. hcaMandatory:', hcaMandatory);

const s4NeedsRem = rows.filter(r => isS4H(r) && !syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r) === 'NEEDS REMEDIATION').length;
console.log('11. s4NeedsRem:', s4NeedsRem);

console.log('\n--- RemType unique values (HCA rows) ---');
const hcaRemTypes = {};
rows.filter(isHCA).forEach(r => { const v = remType(r); hcaRemTypes[v] = (hcaRemTypes[v]||0)+1; });
console.log(JSON.stringify(hcaRemTypes, null, 2));

console.log('\n--- RemType unique values (S4H rows) ---');
const s4RemTypes = {};
rows.filter(isS4H).forEach(r => { const v = remType(r); s4RemTypes[v] = (s4RemTypes[v]||0)+1; });
console.log(JSON.stringify(s4RemTypes, null, 2));

console.log('\n=== SLIDE 7 HCA WATERFALL (chart7Values) ===');

const hcaPreexistCloneFg     = rows.filter(r => isHCA(r) && syn(r) && cln(r) && fg(r)).length;
const hcaTp                  = rows.filter(r => isHCA(r) && tp(r)).length;
const hcaFitgap              = rows.filter(r => isHCA(r) && fg(r) && !syn(r) && !cln(r) && !tp(r)).length;
// hcaMandatory already computed
const hcaClone               = rows.filter(r => isHCA(r) && cln(r) && !syn(r) && !fg(r) && !tp(r) && (remType(r)==='MANDATORY'||remType(r)==='OPTIONAL')).length;
const hcaCloneAndFg          = rows.filter(r => isHCA(r) && cln(r) && fg(r)).length;
const hcaPreexistClone       = rows.filter(r => isHCA(r) && syn(r) && cln(r)).length;
const hcaPreExistingWaterfall= rows.filter(r => isHCA(r) && syn(r) && !cln(r) && !tp(r)).length;
const hcaOptional            = rows.filter(r => isHCA(r) && !syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r)==='OPTIONAL').length;

console.log('hcaPreexistCloneFg (syn+cln+fg):', hcaPreexistCloneFg);
console.log('hcaTp (HCA+tp):', hcaTp);
console.log('hcaFitgap (fg, no syn/cln/tp):', hcaFitgap);
console.log('hcaMandatory (no syn/fg/cln/tp, MANDATORY):', hcaMandatory);
console.log('hcaClone (cln, no syn/fg/tp, MANDATORY or OPTIONAL):', hcaClone);
console.log('hcaCloneAndFg (cln+fg):', hcaCloneAndFg);
console.log('hcaPreexistClone (syn+cln):', hcaPreexistClone);
console.log('hcaPreExistingWaterfall (syn, no cln/tp):', hcaPreExistingWaterfall);
console.log('hcaOptional (no syn/fg/cln/tp, OPTIONAL):', hcaOptional);

const chart7Values = { hcaPreexistCloneFg, hcaTp, hcaFitgap, hcaMandatory, hcaClone, hcaCloneAndFg, hcaPreexistClone, hcaPreExistingWaterfall, hcaOptional };
const chart7Sum = Object.values(chart7Values).reduce((a,b)=>a+b,0);
console.log('\nSum of chart7Values:', chart7Sum, '(hcaCount =', hcaCount, ')');

if (chart7Sum !== hcaCount) {
  console.log('DIFFERENCE:', chart7Sum - hcaCount);
  const uncatHCA = rows.filter(r => {
    if (!isHCA(r)) return false;
    let cat = false;
    if (syn(r) && cln(r) && fg(r)) cat=true;
    if (tp(r)) cat=true;
    if (fg(r) && !syn(r) && !cln(r) && !tp(r)) cat=true;
    if (!syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r)==='MANDATORY') cat=true;
    if (cln(r) && !syn(r) && !fg(r) && !tp(r) && (remType(r)==='MANDATORY'||remType(r)==='OPTIONAL')) cat=true;
    if (cln(r) && fg(r)) cat=true;
    if (syn(r) && cln(r)) cat=true;
    if (syn(r) && !cln(r) && !tp(r)) cat=true;
    if (!syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r)==='OPTIONAL') cat=true;
    return !cat;
  });
  console.log('Uncategorized HCA rows count:', uncatHCA.length);
  uncatHCA.slice(0,5).forEach(s => {
    console.log('  UNCATEGORIZED HCA:', JSON.stringify({syn:s['Syntax Error?'],fg:s['Fit Gap?'],cln:s['Clone?'],tp:s['Object name'].startsWith('/'),rem:s['Remediation Type']}));
  });
}

console.log('\n=== SLIDE 7 S4H WATERFALL (chart8Values) ===');

const s4PreExistingWaterfall = rows.filter(r => isS4H(r) && syn(r) && !cln(r) && !tp(r)).length;
const s4Tp                   = rows.filter(r => isS4H(r) && tp(r) && !syn(r) && !cln(r) && remType(r) !== 'CAN BE IGNORED').length;
const s4Clone                = rows.filter(r => isS4H(r) && cln(r) && !syn(r) && !fg(r) && !tp(r)).length;
const s4CloneAndFg           = rows.filter(r => isS4H(r) && cln(r) && fg(r)).length;
// s4NeedsRem already computed
const s4FitgapWaterfall      = rows.filter(r => {
  if (!isS4H(r)) return false;
  const cond1 = fg(r) && !syn(r) && !cln(r) && !tp(r);
  const cond2 = remType(r).startsWith('NEED TO CHECK') && !syn(r) && !fg(r) && !cln(r) && !tp(r);
  return cond1 || cond2;
}).length;
const s4PreexistClone        = rows.filter(r => isS4H(r) && syn(r) && cln(r)).length;
const s4FalsePos             = rows.filter(r => isS4H(r) && !syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r)==='FALSE POSITIVE').length;
const s4CanIgnore            = rows.filter(r => isS4H(r) && remType(r)==='CAN BE IGNORED' && !syn(r) && !fg(r) && !cln(r)).length;

console.log('s4PreExistingWaterfall (syn, no cln/tp):', s4PreExistingWaterfall);
console.log('s4Tp (tp, no syn/cln, not CAN BE IGNORED):', s4Tp);
console.log('s4Clone (cln, no syn/fg/tp):', s4Clone);
console.log('s4CloneAndFg (cln+fg):', s4CloneAndFg);
console.log('s4NeedsRem (no syn/fg/cln/tp, NEEDS REMEDIATION):', s4NeedsRem);
console.log('s4FitgapWaterfall (fg no syn/cln/tp OR NEED TO CHECK no syn/fg/cln/tp):', s4FitgapWaterfall);
console.log('s4PreexistClone (syn+cln):', s4PreexistClone);
console.log('s4FalsePos (no syn/fg/cln/tp, FALSE POSITIVE):', s4FalsePos);
console.log('s4CanIgnore (CAN BE IGNORED, no syn/fg/cln):', s4CanIgnore);

const chart8Values = { s4PreExistingWaterfall, s4Tp, s4Clone, s4CloneAndFg, s4NeedsRem, s4FitgapWaterfall, s4PreexistClone, s4FalsePos, s4CanIgnore };
const chart8Sum = Object.values(chart8Values).reduce((a,b)=>a+b,0);
console.log('\nSum of chart8Values:', chart8Sum, '(s4Count =', s4Count, ')');

if (chart8Sum !== s4Count) {
  console.log('DIFFERENCE:', chart8Sum - s4Count);
  const uncatS4 = rows.filter(r => {
    if (!isS4H(r)) return false;
    let cat = false;
    if (syn(r) && !cln(r) && !tp(r)) cat=true;
    if (tp(r) && !syn(r) && !cln(r) && remType(r)!=='CAN BE IGNORED') cat=true;
    if (cln(r) && !syn(r) && !fg(r) && !tp(r)) cat=true;
    if (cln(r) && fg(r)) cat=true;
    if (!syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r)==='NEEDS REMEDIATION') cat=true;
    const c1 = fg(r) && !syn(r) && !cln(r) && !tp(r);
    const c2 = remType(r).startsWith('NEED TO CHECK') && !syn(r) && !fg(r) && !cln(r) && !tp(r);
    if (c1||c2) cat=true;
    if (syn(r) && cln(r)) cat=true;
    if (!syn(r) && !fg(r) && !cln(r) && !tp(r) && remType(r)==='FALSE POSITIVE') cat=true;
    if (remType(r)==='CAN BE IGNORED' && !syn(r) && !fg(r) && !cln(r)) cat=true;
    return !cat;
  });
  console.log('Uncategorized S4H rows count:', uncatS4.length);
  uncatS4.slice(0,10).forEach(s => {
    console.log('  UNCATEGORIZED S4H:', JSON.stringify({syn:s['Syntax Error?'],fg:s['Fit Gap?'],cln:s['Clone?'],tp:s['Object name'].startsWith('/'),rem:s['Remediation Type'],obj:s['Object name'].substring(0,30)}));
  });
}
