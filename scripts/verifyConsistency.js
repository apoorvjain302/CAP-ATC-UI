"use strict";
const XLSX        = require("xlsx");
const atcClassify = require("../lib/atcClassify");
const xlsData     = require("../lib/xlsData");

const wb = XLSX.readFile("C:/Users/I755599/Downloads/ATC_Analysis (3).xlsx", { cellText: true, raw: false });
const ws = wb.Sheets[wb.SheetNames[0]];
const bibleRows = XLSX.utils.sheet_to_json(ws, { defval: "" });

const COL_CANONICAL = {
  "object name":"Object name","check title":"Check Title","check message":"Check Message",
  "priority":"Priority","obj.":"Obj.","object type":"Obj.","note":"Note",
  "sap note number":"Note","referenced object":"Referenced Object",
  "ref. object name":"Referenced Object","ref. object type":"Ref. Object Type",
};
const rows = bibleRows.map(r => {
  const out = {};
  for (const k of Object.keys(r)) {
    const c = COL_CANONICAL[k.trim().toLowerCase()] || k.trim();
    out[c] = String(r[k] === null || r[k] === undefined ? "" : r[k]).trim();
  }
  return out;
});

atcClassify.classifyRows(rows, null, "conversion");
const result = xlsData.readXlsDataFromRows(rows);

// ── Bible counts ────────────────────────────────────────────────────────────
const bTotal    = bibleRows.length;
const bHcaRows  = bibleRows.filter(r => String(r["Impact Category"]||"").toUpperCase().includes("HCA")).length;
const bS4Rows   = bibleRows.filter(r => String(r["Impact Category"]||"").toUpperCase().includes("S/4")).length;
const bClone    = bibleRows.filter(r => String(r["Clone?"]||"").toUpperCase() === "YES").length;
const bTP       = bibleRows.filter(r => String(r["3rdPartyObj?"]||"").toUpperCase() === "YES").length;
const bInScope  = bibleRows.filter(r => String(r["InScope?"]||"").toUpperCase() === "YES").length;
const bOutScope = bibleRows.filter(r => String(r["InScope?"]||"").toUpperCase() === "NO").length;
const bFitGap   = bibleRows.filter(r => String(r["FitGap?"]||"").toUpperCase() === "YES").length;
const bSynErr   = bibleRows.filter(r => String(r["SyntaxError?"]||"").toUpperCase() === "YES").length;
const bMandatory= bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "MANDATORY").length;
const bNR       = bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "NEEDS REMEDIATION").length;
const bFP       = bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "FALSE POSITIVE").length;
const bOpt      = bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "OPTIONAL").length;
const bCBI      = bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "CAN BE IGNORED").length;
const bFG       = bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "FIT GAP").length;
const bSyn      = bibleRows.filter(r => String(r["Rem category"]||"").toUpperCase() === "SYNTAX ERROR").length;
const bTBCS     = bibleRows.filter(r => {
  const c = String(r["Rem category"]||"").toUpperCase();
  return c === "TO BE CHECKED IN SYSTEM" || c === "NEED TO CHECK USES OF SEASONAL FIELD INTO SYSTEM A" || c === "NEED TO CHECK INTO SYSTEM AS PER NOTES" || c === "MANUALLY CHECK IN SYSTEM";
}).length;

const bUniqueAll = new Set(bibleRows.map(r => String(r["Concat"]||"").trim()).filter(Boolean));
const bUniqueHCA = new Set(bibleRows.filter(r => String(r["Impact Category"]||"").toUpperCase().includes("HCA")).map(r => String(r["Concat"]||"").trim()).filter(Boolean));
const bUniqueS4H = new Set(bibleRows.filter(r => String(r["Impact Category"]||"").toUpperCase().includes("S/4")).map(r => String(r["Concat"]||"").trim()).filter(Boolean));

// ── App counts ──────────────────────────────────────────────────────────────
const aHcaRows    = rows.filter(r => r["HCA/S4H?"] === "HCA").length;
const aS4Rows     = rows.filter(r => r["HCA/S4H?"] === "S/4H").length;
const aClone      = rows.filter(r => r["Clone?"] === "Yes").length;
const aFGRows     = rows.filter(r => r["Fit Gap?"] === "Yes").length;
const aSynRows    = rows.filter(r => r["Syntax Error?"] === "Yes").length;
const aFGDelta    = rows.filter(r => r["Fit Gap Delta?"] === "Yes").length;
const aMand       = rows.filter(r => r["Remediation Type"] === "Mandatory").length;
const aNR         = rows.filter(r => r["Remediation Type"] === "Needs Remediation").length;
const aFP         = rows.filter(r => r["Remediation Type"] === "False Positive").length;
const aOpt        = rows.filter(r => r["Remediation Type"] === "Optional").length;
const aCBI        = rows.filter(r => r["Remediation Type"] === "Can be ignored").length;

// ── 3rd Party (app: object name starts with "/") ────────────────────────────
const aTP = rows.filter(r => (r["Object name"]||"").trim().startsWith("/")).length;

const chk = (b, a) => b === a ? "✓ MATCH" : `✗ DIFF  (bible=${b} app=${a})`;
const pad = (s, n=6) => String(s).padStart(n);

console.log("================================================================");
console.log("         BIBLE vs APP LOGIC — FULL CONSISTENCY CHECK");
console.log("================================================================\n");

console.log("── ROW TOTALS ──────────────────────────────────────────────────");
console.log(`Total rows              ${pad(bTotal)}  ${pad(bTotal)}  ${chk(bTotal, rows.length)}`);
console.log(`HCA rows                ${pad(bHcaRows)}  ${pad(aHcaRows)}  ${chk(bHcaRows, aHcaRows)}`);
console.log(`S/4H rows               ${pad(bS4Rows)}  ${pad(aS4Rows)}  ${chk(bS4Rows, aS4Rows)}`);
console.log(`Syntax Error rows       ${pad(bSynErr)}  ${pad(aSynRows)}  ${chk(bSynErr, aSynRows)}`);
console.log(`Fit Gap rows (total)    ${pad(bFitGap)}  ${pad(aFGRows)}  (bible FG only; app incl TBCS→FG)`);
console.log(`  Pure Fit Gap          ${pad(bFG)}  ${pad(aFGRows-aFGDelta)}  ${chk(bFG, aFGRows-aFGDelta)}`);
console.log(`  TBCS→FitGapDelta      ${pad(bTBCS)}  ${pad(aFGDelta)}  ${chk(bTBCS, aFGDelta)}`);
console.log(`  Combined FG+TBCS      ${pad(bFG+bTBCS)}  ${pad(aFGRows)}  ${chk(bFG+bTBCS, aFGRows)}`);
console.log("");

console.log("── REM CATEGORY ROW COUNTS ─────────────────────────────────────");
console.log(`                          BIBLE    APP`);
console.log(`Mandatory               ${pad(bMandatory)}  ${pad(aMand)}  ${chk(bMandatory, aMand)}`);
console.log(`Needs Remediation       ${pad(bNR)}  ${pad(aNR)}  ${chk(bNR, aNR)}`);
console.log(`False Positive          ${pad(bFP)}  ${pad(aFP)}  ${chk(bFP, aFP)}`);
console.log(`Optional                ${pad(bOpt)}  ${pad(aOpt)}  ${chk(bOpt, aOpt)}`);
console.log(`Can be Ignored          ${pad(bCBI)}  ${pad(aCBI)}  ${chk(bCBI, aCBI)}`);
console.log(`Fit Gap (pure)          ${pad(bFG)}  ${pad(aFGRows-aFGDelta)}  ${chk(bFG, aFGRows-aFGDelta)}`);
console.log(`Syntax Error            ${pad(bSyn)}  ${pad(aSynRows)}  ${chk(bSyn, aSynRows)}`);
console.log(`To Be Checked (→FGD)    ${pad(bTBCS)}  ${pad(aFGDelta)}  ${chk(bTBCS, aFGDelta)}`);
console.log("");

console.log("── CLONE / 3RD PARTY / SCOPE ───────────────────────────────────");
console.log(`Clone rows              ${pad(bClone)}  ${pad(aClone)}  (0 expected — no clone file uploaded in test)`);
console.log(`3rd Party rows (bible)  ${pad(bTP)}`);
console.log(`3rd Party rows (app)    ${pad(aTP)}  (derived: obj name starts with "/")`);
console.log(`InScope=Yes (bible)     ${pad(bInScope)}`);
console.log(`InScope=No  (bible)     ${pad(bOutScope)}`);
console.log("  NOTE: App does not use an InScope? column — scope is derived");
console.log("  dynamically from rem type, clone, syntax error, 3rd party flags.");
console.log("");

console.log("── UNIQUE OBJECT COUNTS ────────────────────────────────────────");
console.log(`All unique objects      ${pad(bUniqueAll.size)}  (from bible Concat column)`);
console.log(`Unique HCA objects      ${pad(bUniqueHCA.size)}`);
console.log(`Unique S/4H objects     ${pad(bUniqueS4H.size)}`);
console.log("");

// result keys: hcaCount (adjusted = hca-syntax), s4Count, totalCount, fitGapDeltaCount, hcaMandatory, s4TechRemediable
const rHca    = result.hcaCount;
const rS4     = result.s4Count;
const rTotal  = result.totalCount;
const rFGD    = result.fitGapDeltaCount;
const rHcaMan = result.hcaMandatory;
const rS4Rem  = result.s4TechRemediable;

// bible: hcaCount adjusted = HCA rows minus syntax-error HCA rows (matches slide 6)
const bSynHCA = bibleRows.filter(r => String(r["Impact Category"]||"").toUpperCase().includes("HCA") &&
  String(r["SyntaxError?"]||"").toUpperCase() === "YES").length;
const bHcaAdj = bHcaRows - bSynHCA;

console.log("── xlsData PIPELINE OUTPUT (used in Summary tiles + PPT) ───────");
console.log(`hcaCount (HCA - syntax)  ${pad(rHca)}  bible adj=${bHcaAdj}  ${chk(rHca, bHcaAdj)}`);
console.log(`s4Count  (row count)     ${pad(rS4)}   bible=${bS4Rows}  ${chk(rS4, bS4Rows)}`);
console.log(`totalCount               ${pad(rTotal)}  ${chk(rTotal, bTotal)}`);
console.log(`fitGapDeltaCount         ${pad(rFGD)}    ${chk(rFGD, aFGDelta)}`);
console.log(`hcaMandatory (effort C4) ${pad(rHcaMan)}`);
console.log(`s4TechRemediable (C5)    ${pad(rS4Rem)}`);
console.log("");

console.log("── TBCS STRINGS IN OUTPUT ──────────────────────────────────────");
const tbcsInOut = rows.filter(r => {
  const rem = (r["Remediation Type"]||"").toLowerCase();
  return rem.includes("to be checked") || rem.includes("manually check");
});
console.log(`TBCS string in Remediation Type: ${tbcsInOut.length === 0 ? "NONE ✓" : "FOUND " + tbcsInOut.length + " ✗"}`);
console.log("");

console.log("── ENUM VALIDITY ───────────────────────────────────────────────");
const VALID = new Set(["Mandatory","False Positive","Can be ignored","Needs Remediation","Optional","Fit Gap","Syntax Error"]);
const invalid = rows.filter(r => !VALID.has(r["Remediation Type"]));
console.log(`Invalid Remediation Type values: ${invalid.length === 0 ? "NONE ✓" : "FOUND " + invalid.length + " ✗"}`);
const allVals = [...new Set(rows.map(r => r["Remediation Type"]))].sort();
console.log(`All output values: ${allVals.join(", ")}`);
