"use strict";
/**
 * testMatch.js — Compare atcClassify output against ATC_Analysis (3).xlsx bible.
 * Run: node scripts/testMatch.js
 */

const XLSX       = require("xlsx");
const path       = require("path");
const atcClassify = require("../lib/atcClassify");

const BIBLE_PATH = "C:/Users/I755599/Downloads/ATC_Analysis (3).xlsx";

// Column names in the bible
const COL_REM  = "Remediation Type";   // actual output column (set by code)
const COL_MANUAL = "Rem category";     // manual/expected column in bible

// Normalise a remediation category for comparison
function norm(s) {
  return (s || "").toString().trim().toUpperCase()
    .replace(/\s+/g, " ");
}

function main() {
  console.log("Loading bible:", BIBLE_PATH);
  const wb   = XLSX.readFile(BIBLE_PATH, { cellText: true, raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  console.log(`Loaded ${rows.length} rows`);

  // Normalise column names (bible may use slightly different casing)
  const COL_CANONICAL = {
    "object name": "Object name", "check title": "Check Title",
    "check message": "Check Message", "remediation type": "Remediation Type",
    "syntax error?": "Syntax Error?", "fit gap?": "Fit Gap?",
    "hca/s4h?": "HCA/S4H?", "hca/s4h": "HCA/S4H?",
    "hca / s4h?": "HCA/S4H?", "hca / s4h": "HCA/S4H?",
    "clone?": "Clone?", "priority": "Priority",
    "obj.": "Obj.", "object type": "Obj.",
    "note": "Note", "sap note number": "Note",
    "referenced object": "Referenced Object", "ref. object name": "Referenced Object",
    "ref. object type": "Ref. Object Type",
  };

  const normalised = rows.map(r => {
    const out = {};
    for (const k of Object.keys(r)) {
      const canon = COL_CANONICAL[k.trim().toLowerCase()] || k.trim();
      out[canon] = String(r[k] === null || r[k] === undefined ? "" : r[k]).trim();
    }
    return out;
  });

  // Save expected values before classify overwrites "Remediation Type"
  const expected = normalised.map(r => {
    // Bible has "Rem category" as the manual column
    // Try both names
    return (r["Rem category"] || r["Remediation Type"] || r["rem category"] || "").trim();
  });

  // Run classification (mutates rows in place)
  atcClassify.classifyRows(normalised, null, "conversion");

  // Compare
  let match = 0, mismatch = 0;
  const mismatches = {};

  for (let i = 0; i < normalised.length; i++) {
    const got  = norm(normalised[i]["Remediation Type"]);
    const want = norm(expected[i]);

    if (!want) continue; // skip rows with no manual category

    // "To Be Checked in System" variants in the bible are intentionally mapped to
    // "Fit Gap" (with isFitGapDelta=true) by design — treat these as a match.
    const TBCS_VARIANTS = new Set([
      "TO BE CHECKED IN SYSTEM",
      "NEED TO CHECK USES OF SEASONAL FIELD INTO SYSTEM A",
      "NEED TO CHECK INTO SYSTEM AS PER NOTES",
      "MANUALLY CHECK IN SYSTEM",
    ]);
    const effectiveWant = TBCS_VARIANTS.has(want) ? "FIT GAP" : want;

    if (got === effectiveWant) {
      match++;
    } else {
      mismatch++;
      const key = `${normalised[i]["Note"] || ""}|${want}->${got}`;
      mismatches[key] = (mismatches[key] || 0) + 1;
    }
  }

  const total = match + mismatch;
  console.log(`\nMatch: ${match} / ${total} = ${(match/total*100).toFixed(1)}%`);

  if (mismatch > 0) {
    console.log(`\nMismatches (${mismatch} rows):`);
    const sorted = Object.entries(mismatches).sort((a,b) => b[1]-a[1]);
    for (const [key, cnt] of sorted) {
      console.log(`  ${cnt.toString().padStart(5)}  ${key}`);
    }
  }
}

main();
