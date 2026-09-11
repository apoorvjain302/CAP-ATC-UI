"use strict";
/**
 * deepAnalysis.js
 *
 * Deep verification of cap-atc-ui vs ZATCASSESSMENT_KK reference:
 * 1. Classifies Kundan_test.xlsx (conversion mode) and compares vs reference
 * 2. Classifies upgrade mode and checks counts differ appropriately
 * 3. Verifies PPT/estimation counts (atcData fields)
 * 4. Saves all mismatches with exact reasons to Excel
 *
 * Usage: node scripts/deepAnalysis.js
 */

const path  = require("path");
const XLSX  = require("xlsx");
const atcClassify = require("../lib/atcClassify");
const xlsData     = require("../lib/xlsData");

const USER_FILE  = "C:/Users/I755599/Downloads/updated atc_classified_202608100959429.xlsx";
const INPUT_FILE = "C:/Users/I755599/Downloads/Kundan_test.xlsx";
const OUT_FILE   = "C:/Users/I755599/Downloads/deep_analysis_" + Date.now() + ".xlsx";

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
  const wb  = XLSX.readFile(filePath, { cellText: true, raw: false });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
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

function categorizeMismatch(userVal, ourVal, r) {
  const u = userVal.toUpperCase();
  const o = ourVal.toUpperCase();
  const note = (r["Note"] || "").trim();
  const msg  = (r["Check Message"] || "").toUpperCase();
  const title = (r["Check Title"] || "").toUpperCase();
  const refObj = (r["Referenced Object"] || "").toUpperCase();
  const refType = (r["Ref. Object Type"] || "").toUpperCase();

  // Syntax Error in user but we give FP/FG/NR — ABAP propagates SE to all rows of object
  if (u === "SYNTAX ERROR" && o !== "SYNTAX ERROR") {
    return "Kundan_SE_Propagation: User set Syntax Error on all rows of SE object; we keep per-row value";
  }

  // User has "TO BE CHECKED IN SYSTEM" — ABAP note 2368913 / MM flag
  if (u === "TO BE CHECKED IN SYSTEM") {
    return `User_TBCS: User has 'To Be Checked In System' (note ${note}); ABAP uses MM flag for 2368913; we give ${o}`;
  }

  // NR→FG: we say NR, user says FG
  if (u === "FIT GAP" && o === "NEEDS REMEDIATION") {
    return `NR_vs_FG: We classify NR, user says FG. Note=${note} RefObj=${refObj}(${refType})`;
  }

  // FP→FG
  if (u === "FIT GAP" && o === "FALSE POSITIVE") {
    return `FP_vs_FG: We classify FP, user says FG. Note=${note} RefObj=${refObj}(${refType})`;
  }

  // NR→FP
  if (u === "FALSE POSITIVE" && o === "NEEDS REMEDIATION") {
    return `NR_vs_FP: We classify NR, user says FP. Note=${note} RefObj=${refObj}(${refType})`;
  }

  // FP→NR
  if (u === "NEEDS REMEDIATION" && o === "FALSE POSITIVE") {
    return `FP_vs_NR: We classify FP, user says NR. Note=${note} RefObj=${refObj}(${refType})`;
  }

  // FG→FP
  if (u === "FALSE POSITIVE" && o === "FIT GAP") {
    return `FG_vs_FP: We classify FG, user says FP. Note=${note} RefObj=${refObj}(${refType})`;
  }

  // FG→NR
  if (u === "NEEDS REMEDIATION" && o === "FIT GAP") {
    return `FG_vs_NR: We classify FG, user says NR. Note=${note} RefObj=${refObj}(${refType})`;
  }

  // FP→CBI
  if (u === "CAN BE IGNORED" && o === "FALSE POSITIVE") {
    return `FP_vs_CBI: We classify FP, user says Can Be Ignored. Note=${note}`;
  }

  // Other special
  if (u.includes("SEASONAL")) {
    return `Special_Seasonal: User has seasonal note override; we give ${o}. Note=${note}`;
  }

  return `Other: user=${u}, ours=${o}. Note=${note} Title=${title.slice(0,80)}`;
}

function main() {
  console.log("=== Deep Analysis: cap-atc-ui vs ZATCASSESSMENT_KK ===\n");

  // ── Read reference ─────────────────────────────────────────────────────────
  console.log("Reading reference:", USER_FILE);
  const userRows = readRows(USER_FILE);
  console.log(`  ${userRows.length} rows`);

  const userMap = new Map();
  for (const r of userRows) {
    const k = rowKey(r);
    if (k && !userMap.has(k)) userMap.set(k, r);
  }

  // ── Conversion classification ──────────────────────────────────────────────
  console.log("\nReading input:", INPUT_FILE);
  let convRows = readRows(INPUT_FILE);
  console.log(`  ${convRows.length} rows`);

  console.log("Classifying (conversion)...");
  atcClassify.classifyRows(convRows, null, "conversion");
  console.log("Done.");

  // ── Upgrade classification (for regression check) ─────────────────────────
  let upgrRows = readRows(INPUT_FILE);
  console.log("\nClassifying (upgrade)...");
  atcClassify.classifyRows(upgrRows, null, "upgrade");
  console.log("Done.");

  // ── PPT / Effort Estimation counts ────────────────────────────────────────
  console.log("\n=== PPT / Effort Estimation Counts (CONVERSION) ===");
  const convData = xlsData.readXlsDataFromRows(convRows);
  const atcC = convData.atcData;
  if (atcC) {
    const fields = [
      ["uniqueObjCount",        atcC.uniqueObjCount],
      ["totalCount (rows)",     atcC.totalCount],
      ["hcaCount (net)",        atcC.hcaCount],
      ["hcaWaterfallCount",     atcC.hcaWaterfallCount],
      ["s4Count",               atcC.s4Count],
      ["preExistingErrorCount", atcC.preExistingErrorCount],
      ["hcaFitgap",             atcC.hcaFitgap],
      ["hcaMandatory (C4)",     atcC.hcaMandatory],
      ["hcaOptional",           atcC.hcaOptional],
      ["hcaClone",              atcC.hcaClone],
      ["hcaPreExisting",        atcC.hcaPreExisting],
    ];
    for (const [k, v] of fields) console.log(`  ${k.padEnd(30)} = ${v}`);
    console.log(`  ${"d4AutoCount (D4)".padEnd(30)} = ${convData.d4AutoCount}`);
    console.log(`  ${"d5AutoCount (D5)".padEnd(30)} = ${convData.d5AutoCount}`);
    console.log(`  ${"d9AutoCount".padEnd(30)} = ${convData.d9AutoCount}`);
    console.log(`  ${"hcaMandatory (effort C4)".padEnd(30)} = ${convData.hcaMandatory}`);
    console.log(`  ${"s4TechRemediable (C5)".padEnd(30)} = ${convData.s4TechRemediable}`);
    console.log(`  ${"fitGapDeltaCount".padEnd(30)} = ${convData.fitGapDeltaCount}`);
  }

  console.log("\n=== PPT / Effort Estimation Counts (UPGRADE) ===");
  const upgrData = xlsData.readXlsDataFromRows(upgrRows);
  const atcU = upgrData.atcData;
  if (atcU) {
    const fields = [
      ["uniqueObjCount",        atcU.uniqueObjCount],
      ["totalCount (rows)",     atcU.totalCount],
      ["hcaCount (net)",        atcU.hcaCount],
      ["hcaWaterfallCount",     atcU.hcaWaterfallCount],
      ["s4Count",               atcU.s4Count],
      ["preExistingErrorCount", atcU.preExistingErrorCount],
      ["hcaFitgap",             atcU.hcaFitgap],
      ["hcaMandatory (C4)",     atcU.hcaMandatory],
    ];
    for (const [k, v] of fields) console.log(`  ${k.padEnd(30)} = ${v}`);
    console.log(`  ${"hcaMandatory (effort C4)".padEnd(30)} = ${upgrData.hcaMandatory}`);
    console.log(`  ${"s4TechRemediable (C5)".padEnd(30)} = ${upgrData.s4TechRemediable}`);
  }

  // ── Mismatch comparison ────────────────────────────────────────────────────
  let matched = 0, mismatched = 0, notFound = 0;
  const mismatchRows = [];
  const buckets = {};

  for (const r of convRows) {
    const k = rowKey(r);
    if (!userMap.has(k)) { notFound++; continue; }
    const ref  = userMap.get(k);
    const userVal = (ref["Remediation Type"] || "").trim().toUpperCase();
    const ourVal  = (r["Remediation Type"] || "").trim().toUpperCase();

    if (userVal === ourVal) {
      matched++;
    } else {
      mismatched++;
      const bucket = `"${userVal}" → "${ourVal}"`;
      buckets[bucket] = (buckets[bucket] || 0) + 1;

      const reason = categorizeMismatch(userVal, ourVal, r);
      mismatchRows.push({
        "Obj.":              r["Obj."] || "",
        "Object name":       r["Object name"] || "",
        "Check Title":       r["Check Title"] || "",
        "Check Message":     (r["Check Message"] || "").slice(0, 150),
        "Note":              r["Note"] || "",
        "Priority":          r["Priority"] || "",
        "HCA/S4H?":          r["HCA/S4H?"] || "",
        "Ref. Object Type":  r["Ref. Object Type"] || "",
        "Referenced Object": r["Referenced Object"] || "",
        "User_Value":        ref["Remediation Type"] || "",
        "Our_Value":         r["Remediation Type"] || "",
        "Our_FitGap?":       r["Fit Gap?"] || "",
        "Our_SyntaxError?":  r["Syntax Error?"] || "",
        "Reason":            reason,
      });
    }
  }

  console.log("\n=== COMPARISON RESULTS (CONVERSION) ===");
  console.log(`  Matched    : ${matched}`);
  console.log(`  Mismatched : ${mismatched}`);
  console.log(`  Not found  : ${notFound}`);
  console.log(`  Total      : ${convRows.length}`);
  console.log(`  Accuracy   : ${((matched / (matched + mismatched)) * 100).toFixed(2)}%\n`);

  console.log("--- Mismatch buckets (user → ours) ---");
  const sorted = Object.entries(buckets).sort((a, b) => b[1] - a[1]);
  for (const [b, c] of sorted) console.log(`  ${String(c).padStart(5)}x  ${b}`);

  // ── Regression check: upgrade vs conversion ───────────────────────────────
  console.log("\n=== REGRESSION CHECK: Upgrade vs Conversion diff ===");
  const convByCat = {};
  const upgrByCat = {};
  for (const r of convRows) {
    const rt = r["Remediation Type"] || "?";
    convByCat[rt] = (convByCat[rt] || 0) + 1;
  }
  for (const r of upgrRows) {
    const rt = r["Remediation Type"] || "?";
    upgrByCat[rt] = (upgrByCat[rt] || 0) + 1;
  }
  const allCats = new Set([...Object.keys(convByCat), ...Object.keys(upgrByCat)]);
  for (const c of [...allCats].sort()) {
    const cv = convByCat[c] || 0;
    const uv = upgrByCat[c] || 0;
    const diff = uv - cv;
    console.log(`  ${c.padEnd(25)} conv=${String(cv).padStart(6)}  upgr=${String(uv).padStart(6)}  diff=${diff >= 0 ? "+" : ""}${diff}`);
  }

  // Check key expected differences (upgrade should have more FP, less NR for obsolete objects)
  const convFP = convByCat["False Positive"] || 0;
  const upgrFP = upgrByCat["False Positive"] || 0;
  const convNR = convByCat["Needs Remediation"] || 0;
  const upgrNR = upgrByCat["Needs Remediation"] || 0;
  console.log(`\n  Expected: upgrade FP >= conv FP  → ${upgrFP >= convFP ? "✓ OK" : "✗ REGRESSION"} (conv=${convFP}, upgr=${upgrFP})`);
  console.log(`  Expected: upgrade NR <= conv NR  → ${upgrNR <= convNR ? "✓ OK" : "✗ REGRESSION"} (conv=${convNR}, upgr=${upgrNR})`);

  // ── Build summary sheet ────────────────────────────────────────────────────
  const summaryRows = [
    ["=== COMPARISON RESULTS (CONVERSION) ==="],
    ["Matched",           matched],
    ["Mismatched",        mismatched],
    ["Not found",         notFound],
    ["Total",             convRows.length],
    ["Accuracy %",        Number(((matched / (matched + mismatched)) * 100).toFixed(2))],
    [],
    ["=== PPT / EFFORT COUNTS (CONVERSION) ==="],
    ["Field",             "Value"],
    ["uniqueObjCount",    atcC && atcC.uniqueObjCount],
    ["totalCount (rows)", atcC && atcC.totalCount],
    ["hcaCount (net)",    atcC && atcC.hcaCount],
    ["hcaWaterfallCount", atcC && atcC.hcaWaterfallCount],
    ["s4Count",           atcC && atcC.s4Count],
    ["preExistingErrors", atcC && atcC.preExistingErrorCount],
    ["hcaFitgap",         atcC && atcC.hcaFitgap],
    ["hcaMandatory (C4)", atcC && atcC.hcaMandatory],
    ["hcaOptional",       atcC && atcC.hcaOptional],
    ["d4AutoCount (D4)",  convData.d4AutoCount],
    ["d5AutoCount (D5)",  convData.d5AutoCount],
    ["d9AutoCount",       convData.d9AutoCount],
    ["s4TechRemediable (C5)", convData.s4TechRemediable],
    ["fitGapDeltaCount",  convData.fitGapDeltaCount],
    ["hcaMandatory (effort C4)", convData.hcaMandatory],
    [],
    ["=== PPT / EFFORT COUNTS (UPGRADE) ==="],
    ["Field",             "Value"],
    ["uniqueObjCount",    atcU && atcU.uniqueObjCount],
    ["totalCount (rows)", atcU && atcU.totalCount],
    ["hcaCount (net)",    atcU && atcU.hcaCount],
    ["hcaWaterfallCount", atcU && atcU.hcaWaterfallCount],
    ["s4Count",           atcU && atcU.s4Count],
    ["preExistingErrors", atcU && atcU.preExistingErrorCount],
    ["hcaFitgap",         atcU && atcU.hcaFitgap],
    ["hcaMandatory (C4)", atcU && atcU.hcaMandatory],
    ["d4AutoCount (D4)",  upgrData.d4AutoCount],
    ["d5AutoCount (D5)",  upgrData.d5AutoCount],
    ["hcaMandatory (effort C4)", upgrData.hcaMandatory],
    ["s4TechRemediable (C5)", upgrData.s4TechRemediable],
    [],
    ["=== REGRESSION CHECK (upgrade vs conversion) ==="],
    ["Category",          "Conversion",  "Upgrade",    "Diff"],
  ];
  for (const c of [...allCats].sort()) {
    summaryRows.push([c, convByCat[c] || 0, upgrByCat[c] || 0, (upgrByCat[c] || 0) - (convByCat[c] || 0)]);
  }
  summaryRows.push([]);
  summaryRows.push(["FP upgrade >= conv", upgrFP >= convFP ? "OK" : "REGRESSION"]);
  summaryRows.push(["NR upgrade <= conv", upgrNR <= convNR ? "OK" : "REGRESSION"]);
  summaryRows.push([]);
  summaryRows.push(["=== MISMATCH BUCKETS ===", "", "user → ours", "count"]);
  for (const [b, c] of sorted) summaryRows.push(["", "", b, c]);

  // ── Build bucket sheet ─────────────────────────────────────────────────────
  const bucketData = [["User Value", "Our Value", "Count"]];
  for (const [b, c] of sorted) {
    const parts = b.replace(/"/g, "").split(" → ");
    bucketData.push([parts[0] || "", parts[1] || "", c]);
  }

  // ── Build per-category mismatch sheets ────────────────────────────────────
  // Group by category prefix
  const categoryGroups = {};
  for (const r of mismatchRows) {
    const cat = r["Reason"].split(":")[0];
    if (!categoryGroups[cat]) categoryGroups[cat] = [];
    categoryGroups[cat].push(r);
  }

  // ── Write Excel ────────────────────────────────────────────────────────────
  const wb = XLSX.utils.book_new();

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

  const wsBuckets = XLSX.utils.aoa_to_sheet(bucketData);
  XLSX.utils.book_append_sheet(wb, wsBuckets, "Mismatch_Buckets");

  // All mismatches
  if (mismatchRows.length) {
    const wsAll = XLSX.utils.json_to_sheet(mismatchRows);
    XLSX.utils.book_append_sheet(wb, wsAll, "All_Mismatches");
  }

  // Per-category sheets
  for (const [cat, rows] of Object.entries(categoryGroups)) {
    const sheetName = cat.replace(/[\/\\?*\[\]]/g, "_").slice(0, 31);
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  XLSX.writeFile(wb, OUT_FILE);
  console.log(`\n✓ Excel saved: ${OUT_FILE}`);
}

main();
