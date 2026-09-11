"use strict";
/**
 * compareClassify.js
 *
 * Runs cap-atc-ui classification on Kundan_test.xlsx, then compares
 * the resulting 'Remediation Type' column against the user's reference file
 * 'updated atc_classified_202608100959429.xlsx'.
 *
 * Matches rows by: Obj. + Object name + Check Title + Check Message
 *
 * Usage: node scripts/compareClassify.js
 */

const path    = require("path");
const XLSX    = require("xlsx");
const atcClassify = require("../lib/atcClassify");

// CLI: node compareClassify.js [reference_file] [input_file]
// If only one arg given, it is used as BOTH reference AND input (self-comparison of a pre-classified file).
// Defaults kept for backwards compatibility.
const arg1 = process.argv[2];
const arg2 = process.argv[3];
const USER_FILE  = arg1 || "C:/Users/I755599/Downloads/updated atc_classified_202608100959429.xlsx";
const INPUT_FILE = arg2 || arg1 || "C:/Users/I755599/Downloads/Kundan_test.xlsx";

const _COL_CANONICAL = {
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
  "syntex error?": "Syntax Error?",
};

function readRows(filePath) {
  const wb   = XLSX.readFile(filePath, { cellText: true, raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  // Try default JSON parse first; if first row has no real headers, use row-array mode
  const raw  = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
  // Find the header row: first row where at least 3 cells are non-empty strings
  let headerIdx = 0;
  for (let i = 0; i < Math.min(raw.length, 5); i++) {
    const nonEmpty = (raw[i] || []).filter(v => v && String(v).trim().length > 1).length;
    if (nonEmpty >= 3) { headerIdx = i; break; }
  }
  const headers = (raw[headerIdx] || []).map(h => String(h || "").trim());
  const dataRows = raw.slice(headerIdx + 1);
  return dataRows.map(r => {
    const out = {};
    headers.forEach((h, i) => {
      if (!h) return;
      const canonical = _COL_CANONICAL[h.toLowerCase()] || h;
      out[canonical] = String(r[i] === null || r[i] === undefined ? "" : r[i]).trim();
    });
    return out;
  }).filter(r => Object.values(r).some(v => v));
}

function rowKey(r) {
  return [
    (r["Obj."] || "").toUpperCase(),
    (r["Object name"] || "").toUpperCase(),
    (r["Check Title"] || "").toUpperCase(),
    (r["Check Message"] || "").toUpperCase(),
  ].join("|||");
}

function main() {
  console.log("=== cap-atc-ui Classification Comparison ===\n");

  // 1. Read user reference file
  console.log("Reading user reference:", USER_FILE);
  const userRows = readRows(USER_FILE);
  console.log(`  ${userRows.length} rows\n`);

  // 2. Read input file and classify
  console.log("Reading input:", INPUT_FILE);
  let ourRows = readRows(INPUT_FILE);
  console.log(`  ${ourRows.length} rows`);

  console.log("\nRunning atcClassify.classifyRows()...");
  atcClassify.classifyRows(ourRows, null, "conversion");
  console.log("Done.\n");

  // 3. Build lookup map from user rows (key → Remediation Type)
  const userMap = new Map();
  for (const r of userRows) {
    const k = rowKey(r);
    if (k && !userMap.has(k)) {
      userMap.set(k, r["Remediation Type"] || "");
    }
  }

  // 4. Compare
  let matched = 0, mismatched = 0, notFound = 0;
  const mismatchDetails = [];
  const mismatchBuckets = {}; // "userVal → ourVal" → count

  for (const r of ourRows) {
    const k = rowKey(r);
    if (!userMap.has(k)) {
      notFound++;
      continue;
    }
    const userVal = (userMap.get(k) || "").toUpperCase();
    const ourVal  = (r["NP Remediation Type"] || "").toUpperCase();
    if (userVal === ourVal) {
      matched++;
    } else {
      mismatched++;
      const bucket = `"${userVal}" → "${ourVal}"`;
      mismatchBuckets[bucket] = (mismatchBuckets[bucket] || 0) + 1;
      if (mismatchDetails.length < 30) {
        mismatchDetails.push({
          obj:      r["Obj."],
          name:     r["Object name"],
          check:    r["Check Title"],
          user:     userVal,
          ours:     r["NP Remediation Type"] || "",
          refType:  r["Ref. Object Type"] || "",
          refObj:   r["Referenced Object"] || "",
          msg:      (r["Check Message"] || "").slice(0, 120),
        });
      }
    }
  }

  // 5. Print results
  console.log("=== RESULTS ===");
  console.log(`  Matched    : ${matched}`);
  console.log(`  Mismatched : ${mismatched}`);
  console.log(`  Not found  : ${notFound} (key not in user file)`);
  console.log(`  Total      : ${ourRows.length}\n`);

  console.log("--- Mismatch buckets (userValue → ourValue) ---");
  const sorted = Object.entries(mismatchBuckets).sort((a, b) => b[1] - a[1]);
  for (const [bucket, count] of sorted) {
    console.log(`  ${count.toString().padStart(5)}x  ${bucket}`);
  }

  console.log("\n--- First 30 mismatches ---");
  for (const d of mismatchDetails) {
    console.log(`  [${d.obj}] ${d.name}`);
    console.log(`    Check  : ${d.check}`);
    console.log(`    RefType: ${d.refType}  RefObj: ${d.refObj}`);
    console.log(`    Msg    : ${d.msg}`);
    console.log(`    User   : ${d.user}`);
    console.log(`    Ours   : ${d.ours}`);
    console.log();
  }
}

main();
