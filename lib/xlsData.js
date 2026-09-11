"use strict";
/**
 * xlsData.js — Node.js port of tool_sap.py::read_xls_data + compute_atc_classification
 *
 * Uses the 'xlsx' library (SheetJS) to parse both binary .xls and XML spreadsheet formats.
 * Returns the same atcData structure consumed by pptxGen.js.
 */

const XLSX = require("xlsx");

const OBJ_LABEL_MAP = {
  PROG: "Program (PROG)",
  FUGR: "Function Group (FUGR)",
  CLAS: "Class (CLAS)",
  FUGS: "Function Group with Customer Include: SAP Part (FUGS)",
  FUGX: "Function Group with Customer Include: Customer Part (FUGX)",
  TABL: "Table (TABL)",
  IEXT: "Interface Extension (IEXT)",
  INTF: "Interface (ABAP Objects) (INTF)",
  ENHO: "Enhancement Implementation (ENHO)",
  TTYP: "Table Type (TTYP)",
  DTEL: "Data Element (DTEL)",
  SFPF: "Adobe Form (SFPF)",
  VIEW: "View (VIEW)",
  SSFO: "SAP Smart Form (SSFO)",
  WDYN: "Web Dynpro Component (WDYN)",
  SHLP: "Search Help (SHLP)",
  LDBA: "Logical Database (LDBA)",
  DOMA: "Domain (DOMA)",
};

const PREREQ_TITLES = new Set([
  "PREREQUISITES FOR THE TEST",
  "TEST EXISTENCE OF A PROGRAM",
]);

function _isSyntaxError(row) {
  return (
    (row["Remediation Type"] || "").trim().toUpperCase() === "SYNTAX ERROR" ||
    (row["Syntax Error?"] || "").trim() === "Yes"
  );
}

// Pre-existing Error Impact = rows whose Check Title is "Prerequisites for the test".
// These are shown as the "Pre-existing Error Impact" bar on slide 6 and are excluded
// from the HCA waterfall on slide 7 (they are not part of the HCA remediation scope).
function _isPrereq(row) {
  return (row["Check Title"] || "").toLowerCase().includes("prerequisites for the test");
}

// Canonical column name map: normalise known casing variants from different ATC export formats.
// SATC uses "Object name"; Fiori/SAPUI5 uses "Object Name". ABAP preserves input casing in output.
const _COL_CANONICAL = {
  "object name": "Object name",       // SATC input: "Object name"; ABAP result: "Object Name"
  "check title": "Check Title",
  "check message": "Check Message",
  "remediation type": "Remediation Type",
  "syntax error?": "Syntax Error?",
  "fit gap?": "Fit Gap?",
  "hca/s4h?": "HCA/S4H?",
  "hca/s4h":  "HCA/S4H?",           // ABAP result may drop the trailing '?'
  "hca / s4h?": "HCA/S4H?",         // ABAP result may add spaces around '/'
  "hca / s4h":  "HCA/S4H?",
  "clone?": "Clone?",
  "priority": "Priority",
  "obj.": "Obj.",                      // input file column name
  "object type": "Obj.",               // ABAP result column name for the same field
  "note": "Note",                      // input file column name
  "sap note number": "Note",           // ABAP result column name for the same field
  "referenced object": "Referenced Object",    // input file column name
  "ref. object name": "Referenced Object",     // ABAP result column name for the same field
  "robt type": "Ref. Object Type",             // ATC Extract input column name
  "ref. object type": "Ref. Object Type",      // alternate / ABAP result column name
  "refobjecttype": "Ref. Object Type",         // compact variant
  "ref object type": "Ref. Object Type",       // no-dot variant
  "referenced object type": "Ref. Object Type",// full-name variant from production export
  "contact person": "Contact",                 // production export variant
  "first found on": "1st Found",               // production export variant
  "exemption": "Exemption State",              // short variant
  "short description": "Short text",           // alternate label
};

function _normaliseRow(r) {
  const out = {};
  for (const k of Object.keys(r)) {
    const canonical = _COL_CANONICAL[k.trim().toLowerCase()] || k.trim();
    out[canonical] = String(r[k] === null || r[k] === undefined ? "" : r[k]).trim();
  }
  return out;
}

function _readXlsRows(xlsBuf) {
  const wb = XLSX.read(xlsBuf, { type: "buffer", cellText: true, raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (rows.length > 0) {
    console.log("[readXlsRows] Raw column names:", Object.keys(rows[0]).join(" | "));
    const normalised = rows.map(_normaliseRow);
    console.log("[readXlsRows] Normalised column names:", Object.keys(normalised[0]).join(" | "));
    const hcaRows = normalised.filter(r => (r["HCA/S4H?"] || "").trim() === "HCA").length;
    const s4Rows  = normalised.filter(r => (r["HCA/S4H?"] || "").trim() === "S/4H").length;
    console.log(`[readXlsRows] totalRows=${normalised.length} hcaRows=${hcaRows} s4Rows=${s4Rows}`);
    return normalised;
  }
  return [];
}

function computeAtcClassification(rows) {
  const thirdPartyObjs  = new Set();
  const syntaxErrorObjs = new Set();
  const inconsistentObjs= new Set();
  const cloneObjs       = new Set();
  const fitGapObjs      = new Set();

  for (const row of rows) {
    const objName    = (row["Object name"]  || "").trim();
    const checkTitle = (row["Check Title"]  || "").trim().toUpperCase();
    const priority   = (row["Priority"]     || "").trim();

    if (!objName) continue;
    if (objName.startsWith("/"))                          thirdPartyObjs.add(objName);
    if (_isSyntaxError(row))                              syntaxErrorObjs.add(objName);
    if (PREREQ_TITLES.has(checkTitle) && (priority === "2" || priority === "3")) inconsistentObjs.add(objName);
    if ((row["Clone?"]    || "").trim() === "Yes")        cloneObjs.add(objName);
    if ((row["Fit Gap?"]  || "").trim() === "Yes")        fitGapObjs.add(objName);
  }

  // Count rows (not unique objects) where Fit Gap Delta? = Yes
  const fitGapDeltaCount = rows.filter(r => (r["Fit Gap Delta?"] || "").trim() === "Yes").length;

  const remScopeTypes = new Set(["MANDATORY", "NEEDS REMEDIATION", "FALSE POSITIVE"]);
  const noRemTypes    = new Set(["OPTIONAL", "CAN BE IGNORED"]);
  const oosForScope   = new Set([...cloneObjs, ...thirdPartyObjs]);

  const remediationScopeObjs = new Set();
  const noRemediationObjs    = new Set();

  // Debug: sample what values Remediation Type and Fit Gap? actually contain
  const _remSample   = new Set();
  const _fgSample    = new Set();
  let   _skippedOos  = 0, _skippedSyn = 0, _skippedFg = 0, _skippedRemType = 0;
  for (const row of rows) {
    const objName = (row["Object name"] || "").trim();
    if (!objName || oosForScope.has(objName)) { if (objName) _skippedOos++; continue; }
    const rem = (row["Remediation Type"] || "").trim().toUpperCase();
    const fg  = (row["Fit Gap?"]         || "").trim();
    _remSample.add(rem || "(empty)");
    _fgSample.add(fg  || "(empty)");
    if (_isSyntaxError(row)) { _skippedSyn++; continue; }
    if (fg === "Yes") { _skippedFg++; continue; }
    if (remScopeTypes.has(rem))       remediationScopeObjs.add(objName);
    else if (noRemTypes.has(rem))     noRemediationObjs.add(objName);
    else                              _skippedRemType++;
  }
  console.log(`[computeAtcClassification] remScope loop: oosSkipped=${_skippedOos} synSkipped=${_skippedSyn} fgSkipped=${_skippedFg} unknownRemType=${_skippedRemType}`);
  console.log(`[computeAtcClassification] Remediation Type values seen: ${[..._remSample].join(" | ")}`);
  console.log(`[computeAtcClassification] Fit Gap? values seen: ${[..._fgSample].join(" | ")}`);
  console.log(`[computeAtcClassification] After loop: remediationScopeObjs=${remediationScopeObjs.size} noRemediationObjs=${noRemediationObjs.size} inconsistentObjs=${inconsistentObjs.size}`);

  // Inconsistent objs → mandatory findings in remediation scope
  // Skip objects already classified as syntax errors — they belong in the
  // "Syntax Error" bucket on slide 5, not in "Remediation Scope".
  for (const o of inconsistentObjs) {
    if (!oosForScope.has(o) && !syntaxErrorObjs.has(o)) remediationScopeObjs.add(o);
  }
  console.log(`[computeAtcClassification] After inconsistent merge: remediationScopeObjs=${remediationScopeObjs.size}`);
  for (const o of remediationScopeObjs) noRemediationObjs.delete(o);

  const outOfScopeObjCount = new Set([
    ...thirdPartyObjs, ...syntaxErrorObjs, ...cloneObjs, ...fitGapObjs,
  ]).size;

  // ── Slide 6 row counts ────────────────────────────────────────────────────
  const hcaCount     = rows.filter(r => (r["HCA/S4H?"] || "").trim() === "HCA").length;
  const s4Count      = rows.filter(r => (r["HCA/S4H?"] || "").trim() === "S/4H").length;
  // syntaxImpacts = "Prerequisites for the test" check title rows in the HCA bucket.
  // These are the true pre-existing errors shown on slide 6.
  // Other Syntax Error?=Yes rows (e.g. SELECT/ORDER BY) remain in HCA scope and appear
  // in the HCA waterfall as "Pre-existing Issue Impact" (hcaPreExistingWaterfall).
  const syntaxImpacts = rows.filter(r =>
    _isPrereq(r) && (r["HCA/S4H?"] || "").trim() === "HCA"
  ).length;
  console.log(`[computeAtcClassification] totalRows=${rows.length} hcaRows=${hcaCount} s4Rows=${s4Count} syntaxImpacts=${syntaxImpacts}`);
  const tpImpacts    = rows.filter(r => (r["Object name"] || "").trim().startsWith("/")).length;

  // ── HCA row helpers ───────────────────────────────────────────────────────
  const hcaRows = rows.filter(r => (r["HCA/S4H?"] || "").trim() === "HCA");
  const s4hRows = rows.filter(r => (r["HCA/S4H?"] || "").trim() === "S/4H");

  const _syn  = r => _isSyntaxError(r);
  const _fg   = r => (r["Fit Gap?"] || "").trim() === "Yes";
  const _cln  = r => (r["Clone?"]   || "").trim() === "Yes";
  const _tp   = r => (r["Object name"] || "").trim().startsWith("/");
  const _allNo= r => !_syn(r) && !_fg(r) && !_cln(r) && !_tp(r);
  const _rem  = r => (r["Remediation Type"] || "").trim().toUpperCase();

  const hcaFitgap       = hcaRows.filter(r =>  _fg(r) && !_syn(r) && !_cln(r) && !_tp(r)).length;
  // hcaPreExisting = all HCA syntax-error rows (syn takes priority over fg/cln).
  // Matches col P RemType="Syntax Error" count. Sub-cases (syn+fg, syn+cln) are separately
  // tracked in waterfall buckets (hcaPreexistCloneFg, hcaPreexistClone) but counted here too
  // so the slide-8 summary table and estimation match the Excel output.
  const hcaPreExisting  = hcaRows.filter(r =>  _syn(r) && !_tp(r)).length;
  const hcaMandatory    = hcaRows.filter(r =>  _allNo(r) && _rem(r) === "MANDATORY").length;
  const hcaOptional     = hcaRows.filter(r =>  _allNo(r) && _rem(r) === "OPTIONAL").length;
  const hcaClone        = hcaRows.filter(r =>  _cln(r) && !_syn(r) && !_fg(r) && !_tp(r) && ["MANDATORY","OPTIONAL"].includes(_rem(r))).length;

  const s4Fitgap        = s4hRows.filter(r =>  _fg(r) && !_syn(r) && !_cln(r) && !_tp(r)).length;
  // s4PreExisting = all S4H syntax-error rows (syn takes priority over fg/cln).
  // Matches col P RemType="Syntax Error" count. Sub-cases tracked in waterfall separately.
  const s4PreExisting   = s4hRows.filter(r =>  _syn(r) && !_tp(r)).length;
  const s4NeedsRem      = s4hRows.filter(r =>  _allNo(r) && _rem(r) === "NEEDS REMEDIATION").length;
  const s4FalsePos      = s4hRows.filter(r =>  _allNo(r) && _rem(r) === "FALSE POSITIVE").length;
  // Include TP rows in CAN BE IGNORED so the PPT number matches the sheet filter.
  // TP rows (object name starts with /) are classified as "Can be ignored" in the sheet;
  // excluding them via _allNo (which has !_tp) would create a visible mismatch.
  const s4CanIgnore     = s4hRows.filter(r =>  _rem(r) === "CAN BE IGNORED" && !_syn(r) && !_fg(r) && !_cln(r)).length;
  const s4Clone         = s4hRows.filter(r =>  _cln(r) && !_syn(r) && !_fg(r) && !_tp(r)).length;

  return {
    thirdPartyObjs, syntaxErrorObjs, inconsistentObjs, cloneObjs, fitGapObjs,
    outOfScopeObjCount,
    remediationScopeObjs, noRemediationObjs,
    hcaCount, s4Count, syntaxImpacts, tpImpacts,
    fitGapDeltaCount,
    hcaFitgap, hcaPreExisting, hcaMandatory, hcaOptional, hcaClone,
    s4Fitgap, s4PreExisting, s4NeedsRem, s4FalsePos, s4CanIgnore, s4Clone,
  };
}

// ── Shared computation kernel ─────────────────────────────────────────────────
// Single source of truth for all slide/chart/estimation logic.
// Both readXlsData() and readXlsDataFromRows() delegate here.
function _computeFromRows(rows) {
  const cls = computeAtcClassification(rows);
  const {
    thirdPartyObjs, syntaxErrorObjs, inconsistentObjs, cloneObjs, fitGapObjs,
    outOfScopeObjCount, remediationScopeObjs, noRemediationObjs,
    hcaCount, s4Count, syntaxImpacts, tpImpacts, fitGapDeltaCount,
    hcaFitgap, hcaPreExisting, hcaMandatory, hcaOptional, hcaClone,
    s4Fitgap, s4PreExisting, s4NeedsRem, s4FalsePos, s4CanIgnore, s4Clone,
  } = cls;

  // ── Slide 4 ───────────────────────────────────────────────────────────────
  const objTypeMap = {};
  for (const row of rows) {
    const ot = (row["Obj."] || "").trim();
    const on = (row["Object name"] || "").trim();
    if (ot && on) {
      if (!objTypeMap[ot]) objTypeMap[ot] = new Set();
      objTypeMap[ot].add(on);
    }
  }
  const objEntries    = Object.entries(objTypeMap).sort((a, b) => b[1].size - a[1].size);
  const uniqueObjTypes  = objEntries.map(([k]) => OBJ_LABEL_MAP[k] || k);
  const uniqueObjCounts = objEntries.map(([, v]) => v.size);

  // ── Slide 5 ───────────────────────────────────────────────────────────────
  const tpSet      = thirdPartyObjs;
  const synObjs    = new Set([...syntaxErrorObjs].filter(o => !tpSet.has(o)));
  const clnObjs    = new Set([...cloneObjs].filter(o => !tpSet.has(o)));
  const fgObjs     = new Set([...fitGapObjs].filter(o => !tpSet.has(o)));

  const _and    = (a, b)       => new Set([...a].filter(o => b.has(o)));
  const _andNot = (a, ...ex)   => new Set([...a].filter(o => ex.every(e => !e.has(o))));

  const synCloneFg = _and(_and(synObjs, clnObjs), fgObjs);
  const synClone   = _andNot(_and(synObjs, clnObjs), fgObjs);
  const synFg      = _andNot(_and(synObjs, fgObjs),  clnObjs);
  const synOnly    = _andNot(synObjs, clnObjs, fgObjs);
  const cloneFg    = _andNot(_and(clnObjs, fgObjs), synObjs);
  const cloneOnly  = _andNot(clnObjs, synObjs, fgObjs);
  const fitGapOnly = _andNot(fgObjs, synObjs, clnObjs);

  const slide5Categories = [
    "Syntax Error, Clone & Fit Gap", "Syntax Error & Clone", "Syntax Error & Fit Gap",
    "Syntax Error", "Remediation Scope", "No Remediation Required",
    "Fit Gap", "Clone & Fit Gap", "Clone", "3rd Party Obj",
  ];
  // Each unique object is counted in exactly one bucket (exclusive partition).
  // remediationScopeObjs no longer contains syntax-error objects (fixed upstream
  // in computeAtcClassification), so no extra subtraction is needed here.
  const slide5Values = [
    synCloneFg.size, synClone.size, synFg.size, synOnly.size,
    remediationScopeObjs.size, noRemediationObjs.size,
    fitGapOnly.size, cloneFg.size, cloneOnly.size, thirdPartyObjs.size,
  ];

  // ── Slide 6 ───────────────────────────────────────────────────────────────
  const slide6Categories = [
    "01. Pre-existing Issue Impact", "02. HCA Impact",
    "03. S/4HANA Impact",           "04. Impact in 3rd Party Obj",
  ];
  const slide6Values = [syntaxImpacts, hcaCount - syntaxImpacts, s4Count, tpImpacts];

  // ── Slide 7 HCA waterfall ─────────────────────────────────────────────────
  // "Pre-existing Issue, Clone & Fit Gap Impact": all THREE flags required (syn+cln+fg).
  // syn+fg rows without clone fall into hcaPreExistingWaterfall instead (slot 7).
  const hcaPreexistCloneFg = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="HCA" && _isSyntaxError(r) &&
    (r["Clone?"]||"").trim()==="Yes" && (r["Fit Gap?"]||"").trim()==="Yes"
  ).length;
  const hcaTp = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="HCA" && (r["Object name"]||"").trim().startsWith("/")
  ).length;
  const hcaCloneAndFg = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="HCA" && (r["Clone?"]||"").trim()==="Yes" && (r["Fit Gap?"]||"").trim()==="Yes"
  ).length;
  const hcaPreexistClone = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="HCA" && _isSyntaxError(r) && (r["Clone?"]||"").trim()==="Yes"
  ).length;
  // Waterfall "Pre-existing Issue Impact" slot: Syntax Error?=Yes rows that are NOT prereq,
  // NOT clone, NOT third-party. Prereq rows are already accounted for in the slide 6
  // Pre-existing Error Impact bar and must NOT appear in the HCA waterfall breakdown.
  const hcaPreExistingWaterfall = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="HCA" && _isSyntaxError(r) && !_isPrereq(r) &&
    (r["Clone?"]||"").trim()!=="Yes" &&
    !(r["Object name"]||"").trim().startsWith("/")
  ).length;

  const chart7Categories = [
    "Pre-existing Issue, Clone & Fit Gap Impact", "3rd Party Obj Impact",
    "Fit Gap Impact", "MANDATORY", "Clone Impact", "Clone & Fit Gap Impact",
    "Pre-existing Issue & Clone Impact", "Pre-existing Issue Impact", "OPTIONAL",
  ];
  const chart7Values = [
    hcaPreexistCloneFg, hcaTp, hcaFitgap, hcaMandatory, hcaClone,
    hcaCloneAndFg, hcaPreexistClone, hcaPreExistingWaterfall, hcaOptional,
  ];

  // ── Slide 7 S/4H waterfall ────────────────────────────────────────────────
  // s4PreExisting covers all Syntax Error rows in S/4H after propagateSyntaxError.
  // s4Tp: TP rows in S/4H that do NOT fall into other buckets.
  // TP rows with rem=CAN BE IGNORED are already counted in s4CanIgnore.
  // TP rows with syn flag go to s4PreExisting / s4PreexistClone.
  // TP rows with clone flag go to s4Clone / s4CloneAndFg.
  // The remaining TP rows (with NR/FP/FG rem) are captured here so the waterfall sums to s4Count.
  const s4Tp = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim() === "S/4H" &&
    (r["Object name"]||"").trim().startsWith("/") &&
    !_isSyntaxError(r) && (r["Clone?"]||"").trim() !== "Yes" &&
    (r["Remediation Type"]||"").trim().toUpperCase() !== "CAN BE IGNORED"
  ).length;
  const s4CloneAndFg = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="S/4H" && (r["Clone?"]||"").trim()==="Yes" && (r["Fit Gap?"]||"").trim()==="Yes"
  ).length;
  const s4PreexistClone = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="S/4H" && _isSyntaxError(r) && (r["Clone?"]||"").trim()==="Yes"
  ).length;
  // Waterfall "Pre-existing Issue Impact": all syn rows (including syn+fg) excluding clone/tp
  const s4PreExistingWaterfall = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="S/4H" && _isSyntaxError(r) &&
    (r["Clone?"]||"").trim()!=="Yes" &&
    !(r["Object name"]||"").trim().startsWith("/")
  ).length;
  // Waterfall "Fit Gap Impact": fg rows (no syn/cln/tp) + FitGapDelta rows (RemType="Fit Gap",
  // Fit Gap?=No, Fit Gap Delta?=Yes — these don't propagate so Fit Gap? stays No)
  const _isFitGapDelta = r => (r["Fit Gap Delta?"] || "").trim() === "Yes" && (r["Fit Gap?"] || "").trim() !== "Yes";
  const s4FitgapWaterfall = rows.filter(r =>
    (r["HCA/S4H?"]||"").trim()==="S/4H" && (
      ((r["Fit Gap?"]||"").trim()==="Yes" && !_isSyntaxError(r) && (r["Clone?"]||"").trim()!=="Yes" && !(r["Object name"]||"").trim().startsWith("/")) ||
      (_isFitGapDelta(r) && !_isSyntaxError(r) && (r["Clone?"]||"").trim()!=="Yes" && !(r["Object name"]||"").trim().startsWith("/"))
    )
  ).length;

  const chart8Categories = [
    "Pre-existing Issue Impact", "3rd Party Obj Impact",
    "Clone Impact", "Clone & Fit Gap Impact", "NEEDS REMEDIATION",
    "Fit Gap Impact", "Pre-existing Issue & Clone Impact",
    "FALSE POSITIVE", "CAN BE IGNORED",
  ];
  const chart8Values = [
    s4PreExistingWaterfall, s4Tp, s4Clone, s4CloneAndFg, s4NeedsRem,
    s4FitgapWaterfall, s4PreexistClone, s4FalsePos, s4CanIgnore,
  ];

  // ── Effort Estimation: D4 auto count (HCA) ───────────────────────────────
  // Rows where HCA/S4H?=HCA, _allNo (no syn/fg/clone/tp), Remediation Type=MANDATORY,
  // Obj. in PROG/FUGR/CLAS/FUGS/FUGX, Check Title contains 'SELECT' and 'ORDER BY'.
  // Population matches hcaMandatory so D4 <= C4 is always guaranteed.
  const _HCA_AUTO_OBJ_TYPES = new Set(["PROG", "FUGR", "CLAS", "FUGS", "FUGX"]);
  const d4AutoCount = rows.filter(r => {
    const obj   = (r["Obj."]              || "").trim().toUpperCase();
    const rem   = (r["Remediation Type"]  || "").trim().toUpperCase();
    const title = (r["Check Title"]       || "").trim().toUpperCase();
    const hcas4 = (r["HCA/S4H?"]         || "").trim();
    const isSyn = _isSyntaxError(r);
    const isFg  = (r["Fit Gap?"]          || "").trim() === "Yes";
    const isCln = (r["Clone?"]            || "").trim() === "Yes";
    const isTp  = (r["Object name"]       || "").trim().startsWith("/");
    return hcas4 === "HCA" &&
           !isSyn && !isFg && !isCln && !isTp &&   // same as _allNo — ensures D4 ≤ C4
           _HCA_AUTO_OBJ_TYPES.has(obj) &&
           rem === "MANDATORY" &&
           title.includes("SELECT") && title.includes("ORDER BY");
  }).length;

  // ── Effort Estimation: D5 auto count (S/4HANA) ───────────────────────────
  // Only NEEDS REMEDIATION rows for S/4H (subset of s4TechRemediable = s4NeedsRem + s4FalsePos),
  // _allNo (no syn/fg/clone/tp), matching any of 6 note/ref-object filters.
  // Population is a subset of s4NeedsRem ⊆ s4TechRemediable so D5 <= C5 is always guaranteed.
  const _FLE_MSG_KEYWORDS = ["WRITE", "IMPORT", "EXPORT", "TRANSFER", "READ-DATASET", "READ DATASET", "REPLACE"];
  const _FLE_NOTES        = new Set(["2215424", "2610650"]);
  const _SD_REF_OBJS      = new Set(["VBFA", "VBUK", "VBUP", "VBTYP", "LIKPUK", "RV_ORDER_FLOW_INFORMATION"]);
  const _BILL_REF_OBJS    = new Set(["VBRK", "VBRP", "VBELN", "DRAFT"]);
  const _KONV_REF_OBJS    = new Set(["KONV", "KOLNR"]);

  const d5AutoCount = rows.filter(r => {
    const hcas4  = (r["HCA/S4H?"]          || "").trim();
    const rem    = (r["Remediation Type"]   || "").trim().toUpperCase();
    const obj    = (r["Obj."]              || "").trim().toUpperCase();
    if (hcas4 !== "S/4H" || rem !== "NEEDS REMEDIATION") return false;
    if (!_HCA_AUTO_OBJ_TYPES.has(obj)) return false;
    // _allNo guard — ensures D5 ≤ C5 (s4TechRemediable)
    if (_isSyntaxError(r))                                 return false;
    if ((r["Fit Gap?"]     || "").trim() === "Yes")        return false;
    if ((r["Clone?"]       || "").trim() === "Yes")        return false;
    if ((r["Object name"]  || "").trim().startsWith("/"))  return false;

    const note   = (r["Note"]              || "").trim();
    const refObj = (r["Referenced Object"] || "").trim().toUpperCase();
    const title  = (r["Check Title"]       || "").trim().toUpperCase();
    const msg    = (r["Check Message"]     || "").trim().toUpperCase();

    // Filters 1 & 2: field length extension
    if (_FLE_NOTES.has(note) && title.includes("FIELD LENGTH") &&
        _FLE_MSG_KEYWORDS.some(k => msg.includes(k))) return true;
    // Filter 3: SD data model changes
    if (note === "2198647" && _SD_REF_OBJS.has(refObj))   return true;
    // Filter 4: SD billing document draft
    if (note === "2768887" && _BILL_REF_OBJS.has(refObj)) return true;
    // Filter 5: pricing/condition technique
    if (note === "2220005" && _KONV_REF_OBJS.has(refObj)) return true;
    // Filter 6: general ledger changes
    if (note === "2431747" && refObj === "BSEG")           return true;

    return false;
  }).length;

  const atcData = {
    uniqueObjCount:    uniqueObjCounts.reduce((a, b) => a + b, 0),
    uniqueObjTypes,
    uniqueObjCounts,
    slide5Categories,
    slide5Values,
    slide6Categories,
    slide6Values,
    chart7Categories,
    chart7Values,
    chart8Categories,
    chart8Values,
    hcaCount:              hcaCount - syntaxImpacts,
    hcaWaterfallCount:     hcaCount - syntaxImpacts,
    hcaMandatoryCount:     hcaMandatory,
    s4TechRemediableCount: s4NeedsRem,
    s4Count,
    totalCount:            rows.length,
    preExistingErrorCount: syntaxImpacts,
    thirdPartyCount:       tpImpacts,
    cloneObjCount:         cloneObjs.size,
    syntaxErrorObjCount:   syntaxErrorObjs.size,
    inconsistentObjCount:  inconsistentObjs.size,
    fitGapObjCount:        fitGapObjs.size,
    outOfScopeObjCount,
    remediationScopeCount: remediationScopeObjs.size,
    noRemediationCount:    noRemediationObjs.size,
    // Top Fit Gap Check Titles for slide 10 bullet points (up to 2)
    fitGapCategories: (() => {
      const titleCounts = {};
      for (const row of rows) {
        if ((row["Fit Gap?"] || "").trim() === "Yes") {
          const t = (row["Check Title"] || "").trim();
          if (t) titleCounts[t] = (titleCounts[t] || 0) + 1;
        }
      }
      return Object.entries(titleCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 2)
        .map(([t]) => t);
    })(),
    cloneImpactCount: rows.filter(r => (r["Clone?"] || "").trim() === "Yes").length,
  };

  const d9AutoCount = computeCcmAgentEligible(rows);

  return {
    atcData,
    hcaCount:        hcaCount - syntaxImpacts,
    s4Count,
    totalCount:      rows.length,
    d4AutoCount,
    d5AutoCount,
    d9AutoCount,
    fitGapDeltaCount,
    // Effort estimation inputs: C4 = mandatory HCA findings only;
    // C5 = NEEDS REMEDIATION S/4H only (FALSE POSITIVE excluded)
    hcaMandatory,
    s4TechRemediable: s4NeedsRem,
    error:           null,
  };
}

function readXlsData(xlsBuf) {
  try {
    const rows = _readXlsRows(xlsBuf);
    if (!rows.length) return { atcData: null, error: "No data rows in XLS" };
    return _computeFromRows(rows);
  } catch (err) {
    return { atcData: null, error: err.message };
  }
}

/**
 * Split an ATC XLSX buffer into an array of chunk XLSX buffers.
 * Each chunk contains the original header row + up to chunkSize data rows.
 * If the file has <= chunkSize data rows, returns [xlsBuf] unchanged — no parse at all.
 *
 * Uses JSZip to read raw sheet XML directly — never builds a full JS row array for the
 * input file, which would OOM for 100k+ row files in a 512MB process.
 * Returns a Promise<Buffer[]>.
 */
async function splitXlsxRows(xlsBuf, chunkSize) {
  const JSZip = require("jszip");

  // ── Open the XLSX ZIP and extract the first sheet's raw XML ──────────────
  let zip;
  try {
    zip = await JSZip.loadAsync(xlsBuf);
  } catch (err) {
    console.error("[splitXlsxRows] ZIP open failed:", err.message);
    return [xlsBuf];
  }

  // Find the first sheet file (xl/worksheets/sheet1.xml or similar)
  const sheetEntry = zip.file("xl/worksheets/sheet1.xml")
                  || zip.file(/xl\/worksheets\/sheet\d+\.xml/)[0];
  if (!sheetEntry) {
    console.error("[splitXlsxRows] sheet1.xml not found in XLSX");
    return [xlsBuf];
  }

  const xmlStr = await sheetEntry.async("string");

  // ── Parse <row> elements from the raw XML ────────────────────────────────
  // Match every <row ...>...</row> block (including self-closing cells inside)
  const rowRegex = /<row\b[^>]*>[\s\S]*?<\/row>/g;
  const rowMatches = xmlStr.match(rowRegex);
  if (!rowMatches || rowMatches.length === 0) return [xlsBuf];

  // Row 0 = header row, rows 1..N = data rows
  const headerXml = rowMatches[0];
  const dataXmls  = rowMatches.slice(1);

  // No split needed — return original buffer untouched, no rewrite
  if (dataXmls.length <= chunkSize) return [xlsBuf];

  // ── Extract the sheet XML wrapper (everything before first <row> and after last </row>) ──
  const firstRowIdx = xmlStr.indexOf(headerXml);
  const lastRowEnd  = xmlStr.lastIndexOf("</row>") + "</row>".length;
  const xmlPrefix   = xmlStr.slice(0, firstRowIdx);       // <worksheet><sheetData>
  const xmlSuffix   = xmlStr.slice(lastRowEnd);            // </sheetData></worksheet>

  // Also need shared strings + styles + workbook XML from the original ZIP for valid XLSX
  // Simplest: rebuild each chunk as a complete XLSX by reusing all ZIP entries except sheet1
  const baseEntries = {};
  zip.forEach((relPath, file) => {
    if (relPath !== "xl/worksheets/sheet1.xml") {
      baseEntries[relPath] = file;
    }
  });

  // ── Renumber a row's r= attributes so rows are sequential from 1 ──────────
  // ABAP's cl_fdt_xl_spreadsheet uses the r= row index directly.
  // Chunk 2 rows have r="30002" etc — must be renumbered to r="2", r="3"...
  // Also renumber cell references: <c r="A30002"> → <c r="A2">
  const COL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  function _colFromIndex(colIdx) {  // 0-based column index → letter(s)
    let s = "";
    colIdx++;
    while (colIdx > 0) {
      colIdx--;
      s = COL_LETTERS[colIdx % 26] + s;
      colIdx = Math.floor(colIdx / 26);
    }
    return s;
  }

  function _renumberRow(rowXml, newRowNum) {
    // Replace row r= attribute (single, no /g needed — only one <row> tag)
    let out = rowXml.replace(/(<row\b[^>]*\br=")[^"]*(")/,
                             (_, pre, suf) => `${pre}${newRowNum}${suf}`);
    // Replace ALL cell r= attributes: letter(s) + old number → letter(s) + newRowNum
    // Must use /g flag — each row has multiple <c> elements
    out = out.replace(/(<c\b[^>]*\br=")([A-Z]+)\d+(")/g,
                      (_, pre, col, suf) => `${pre}${col}${newRowNum}${suf}`);
    return out;
  }

  // ── Build chunk XLSX buffers ──────────────────────────────────────────────
  const chunks = [];
  for (let i = 0; i < dataXmls.length; i += chunkSize) {
    const slice    = dataXmls.slice(i, i + chunkSize);

    // Renumber: header stays at row 1, data rows start at 2
    const renumberedHeader = _renumberRow(headerXml, 1);
    const renumberedSlice  = slice.map((rowXml, idx) => _renumberRow(rowXml, idx + 2));

    // Fix the <dimension ref="A1:R28517"/> row count to reflect actual chunk row count.
    // Extract the last column letter(s) from the original dimension (e.g. "R" from "A1:R28517")
    // so we don't accidentally narrow the dimension to a single column.
    const dimMatch  = xmlPrefix.match(/<dimension\s+ref="[A-Z]+\d+:([A-Z]+)\d+"/);
    const lastCol   = dimMatch ? dimMatch[1] : "ZZ";   // fallback: ZZ covers any width
    const chunkRowCount = slice.length + 1;  // +1 for header
    const fixedPrefix   = xmlPrefix.replace(/(<dimension\s+ref=")[^"]*(")/,
                            (_, pre, suf) => `${pre}A1:${lastCol}${chunkRowCount}${suf}`);

    const chunkXml = fixedPrefix + renumberedHeader + renumberedSlice.join("") + xmlSuffix;

    const chunkZip = new JSZip();
    // Copy all base entries (shared strings, styles, workbook, rels, etc.)
    for (const [relPath, file] of Object.entries(baseEntries)) {
      chunkZip.file(relPath, await file.async("nodebuffer"));
    }
    // Replace sheet1.xml with the chunk's rows
    chunkZip.file("xl/worksheets/sheet1.xml", chunkXml);

    const chunkBuf = await chunkZip.generateAsync({
      type:               "nodebuffer",
      compression:        "DEFLATE",
      compressionOptions: { level: 1 },   // fast compression, smaller chunks
    });
    chunks.push(chunkBuf);
  }

  console.log(`[splitXlsxRows] Split ${dataXmls.length} rows into ${chunks.length} chunks of up to ${chunkSize}`);
  return chunks;
}

/**
 * Merge an array of XLS/XLSX result buffers (one per chunk, returned by ABAP)
 * into a single XLSX buffer whose rows are the concatenation of all chunks.
 * Each buffer must have a header row; only the first chunk's header is kept.
 * If only one buffer is supplied, it is returned unchanged — no rewrite.
 *
 * Handles two formats:
 *   • OOXML (.xlsx ZIP) — uses JSZip to merge raw sheet XML (row-level splice)
 *   • SpreadsheetML XML — ABAP's cl_fdt_xl_spreadsheet output (not a ZIP); detected
 *     by the absence of the PK ZIP magic bytes. Merged by raw XML string splicing —
 *     no SheetJS serialization, so heap stays proportional to one chunk at a time.
 */
async function mergeXlsResults(xlsBuffers) {
  if (xlsBuffers.length === 1) return xlsBuffers[0];

  // ── Detect format: ZIP starts with 0x50 0x4B ('PK') ─────────────────────
  const isZip = xlsBuffers[0].length >= 2 &&
                xlsBuffers[0][0] === 0x50 &&
                xlsBuffers[0][1] === 0x4B;

  if (!isZip) {
    // ── SpreadsheetML XML path (ABAP result format) ───────────────────────
    // ABAP's cl_fdt_xl_spreadsheet produces SpreadsheetML XML with <Row> elements
    // (capital R) inside <Table>. We merge purely at the XML string level:
    //   • capture prefix (everything before first <Row>) and suffix (after last </Row>)
    //     from chunk 0
    //   • append data <Row> elements from all chunks (skip header row for chunks 1+)
    //   • renumber ss:Index attributes so rows are sequential
    // This avoids building any SheetJS cell map, keeping peak heap ~ one chunk.
    console.log("[mergeXlsResults] SpreadsheetML format detected — using XML string merge");

    // SpreadsheetML row regex: <Row ...>...</Row> (capital R, ss: namespace attrs)
    const smlRowRe = /<Row\b[^>]*>[\s\S]*?<\/Row>/g;

    let xmlPrefix    = null;  // everything before first <Row> in chunk 0
    let xmlSuffix    = null;  // everything after last </Row> in chunk 0
    let headerRowXml = null;  // the header <Row> from chunk 0
    const dataRowParts = [];  // collected data <Row> XML strings, all chunks
    let totalDataRows = 0;

    for (let ci = 0; ci < xlsBuffers.length; ci++) {
      const xmlStr   = xlsBuffers[ci].toString("utf8");
      // Release the raw buffer immediately after string conversion to free ~6.5 MB
      xlsBuffers[ci] = null;
      const rowMatches = xmlStr.match(smlRowRe);
      if (!rowMatches || rowMatches.length === 0) continue;

      if (ci === 0) {
        headerRowXml = rowMatches[0];
        const firstRowStart = xmlStr.indexOf(headerRowXml);
        const lastRowEnd    = xmlStr.lastIndexOf("</Row>") + "</Row>".length;
        xmlPrefix = xmlStr.slice(0, firstRowStart);
        xmlSuffix = xmlStr.slice(lastRowEnd);
        // Data rows from first chunk (keep as-is, already numbered from 2)
        for (let i = 1; i < rowMatches.length; i++) {
          dataRowParts.push(rowMatches[i]);
          totalDataRows++;
        }
      } else {
        // Subsequent chunks: skip header row (index 0), renumber remaining rows
        for (let i = 1; i < rowMatches.length; i++) {
          const newIdx = totalDataRows + 2; // 1-based, header = row 1
          // Update ss:Index attribute if present; if absent rows are auto-sequential
          const renumbered = rowMatches[i].replace(/(\bss:Index=")[^"]*(")/,
                                                   (_, pre, suf) => `${pre}${newIdx}${suf}`);
          dataRowParts.push(renumbered);
          totalDataRows++;
        }
      }
    }

    if (!headerRowXml || !xmlPrefix) {
      console.log("[mergeXlsResults] SpreadsheetML merge failed — falling back to first chunk");
      return xlsBuffers[0];
    }

    // Update ExpandedRowCount in the <Table> tag so Excel knows the true row count
    const fixedPrefix = xmlPrefix.replace(/(\bExpandedRowCount=")[^"]*(")/,
                          (_, pre, suf) => `${pre}${totalDataRows + 1}${suf}`);

    // Join parts then immediately clear the array so GC can reclaim ~30 MB before Buffer.from
    const mergedXml = fixedPrefix + headerRowXml + dataRowParts.join("") + xmlSuffix;
    dataRowParts.length = 0;
    console.log(`[mergeXlsResults] SpreadsheetML merged: ${totalDataRows} data rows`);
    return Buffer.from(mergedXml, "utf8");
  }

  // ── OOXML ZIP path (standard .xlsx format) ───────────────────────────────
  const JSZip = require("jszip");

  const rowRegex = /<row\b[^>]*>[\s\S]*?<\/row>/g;

  let headerXml   = null;   // <row> XML of header from first chunk
  let xmlPrefix   = null;   // everything before first <row> in first chunk
  let xmlSuffix   = null;   // everything after last </row> in first chunk
  let baseEntries = null;   // non-sheet ZIP entries from first chunk
  const allDataRowXmls = []; // raw data <row> XML strings from all chunks, renumbered
  let totalDataRows = 0;

  function _renumberRow(rowXml, newRowNum) {
    let out = rowXml.replace(/(<row\b[^>]*\br=")[^"]*(")/,
                             (_, pre, suf) => `${pre}${newRowNum}${suf}`);
    out = out.replace(/(<c\b[^>]*\br=")([A-Z]+)\d+(")/g,
                      (_, pre, col, suf) => `${pre}${col}${newRowNum}${suf}`);
    return out;
  }

  for (let ci = 0; ci < xlsBuffers.length; ci++) {
    let zip;
    try { zip = await JSZip.loadAsync(xlsBuffers[ci]); } catch (_) { continue; }

    const sheetEntry = zip.file("xl/worksheets/sheet1.xml")
                    || zip.file(/xl\/worksheets\/sheet\d+\.xml/)[0];
    if (!sheetEntry) continue;

    const xmlStr   = await sheetEntry.async("string");
    const rowMatches = xmlStr.match(rowRegex);
    if (!rowMatches || rowMatches.length === 0) continue;

    if (ci === 0) {
      // First chunk: capture structural metadata
      headerXml = rowMatches[0];
      const firstRowIdx = xmlStr.indexOf(headerXml);
      const lastRowEnd  = xmlStr.lastIndexOf("</row>") + "</row>".length;
      xmlPrefix = xmlStr.slice(0, firstRowIdx);
      xmlSuffix = xmlStr.slice(lastRowEnd);

      baseEntries = {};
      zip.forEach((relPath, file) => {
        if (relPath !== "xl/worksheets/sheet1.xml") baseEntries[relPath] = file;
      });

      // Data rows from first chunk — keep original row numbers (start at 2)
      for (let i = 1; i < rowMatches.length; i++) {
        allDataRowXmls.push(rowMatches[i]);
      }
      totalDataRows += rowMatches.length - 1;
    } else {
      // Subsequent chunks: skip header (row 0), renumber data rows to continue sequence
      for (let i = 1; i < rowMatches.length; i++) {
        const newRowNum = totalDataRows + 2; // +2: 1-based + header occupies row 1
        allDataRowXmls.push(_renumberRow(rowMatches[i], newRowNum));
        totalDataRows++;
      }
    }
  }

  if (!headerXml || !xmlPrefix) return xlsBuffers[0];  // fallback

  // Fix the <dimension> to reflect the total row count
  const dimMatch  = xmlPrefix.match(/<dimension\s+ref="[A-Z]+\d+:([A-Z]+)\d+"/);
  const lastCol   = dimMatch ? dimMatch[1] : "ZZ";
  const totalRows = totalDataRows + 1;  // +1 for header
  const fixedPrefix = xmlPrefix.replace(/(<dimension\s+ref=")[^"]*(")/,
                        (_, pre, suf) => `${pre}A1:${lastCol}${totalRows}${suf}`);

  // Renumber header to row 1
  const fixedHeader = _renumberRow(headerXml, 1);

  const mergedXml = fixedPrefix + fixedHeader + allDataRowXmls.join("") + xmlSuffix;

  const mergedZip = new JSZip();
  for (const [relPath, file] of Object.entries(baseEntries)) {
    mergedZip.file(relPath, await file.async("nodebuffer"));
  }
  mergedZip.file("xl/worksheets/sheet1.xml", mergedXml);

  return mergedZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 1 } });
}

/**
 * Multi-chunk result path: parse each ABAP result buffer into rows (one XLSX.read per chunk),
 * skip the header row of all but the first chunk, and return the concatenated row array.
 *
 * This avoids ever building a merged in-memory XLSX, keeping peak heap usage proportional
 * to one chunk at a time rather than the sum of all chunks.
 */
function mergeXlsResultBufsToRows(xlsBuffers) {
  const allRows = [];
  for (let ci = 0; ci < xlsBuffers.length; ci++) {
    const buf  = xlsBuffers[ci];
    const wb   = XLSX.read(buf, { type: "buffer", cellText: true, raw: false });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    // Log column names from first chunk to diagnose mapping issues
    if (ci === 0 && rows.length > 0) {
      console.log("[mergeXlsResultBufsToRows] Column names from chunk 0:", Object.keys(rows[0]).join(" | "));
    }
    const normalised = rows.map(_normaliseRow);
    // Log first normalised row to verify canonical mapping
    if (ci === 0 && normalised.length > 0) {
      console.log("[mergeXlsResultBufsToRows] First normalised row keys:", Object.keys(normalised[0]).join(" | "));
      console.log("[mergeXlsResultBufsToRows] First normalised row 'Object name':", normalised[0]["Object name"]);
    }
    allRows.push(...normalised);
  }
  console.log(`[mergeXlsResultBufsToRows] Total rows after merge: ${allRows.length}`);
  return allRows;
}

/**
 * Combined multi-chunk processing: in a SINGLE pass over xlsBuffers, collect
 * all normalised rows for analysis AND build the merged SpreadsheetML/OOXML
 * download artifact.  Each source buffer is nulled immediately after being
 * consumed so peak heap is bounded to roughly one chunk at a time.
 *
 * Returns { dataResult, xlsBuf } — equivalent to calling
 *   mergeXlsResultBufsToRows + readXlsDataFromRows + mergeXlsResults
 * but with ~65 MB less peak heap because the raw buffers are freed before
 * the merged output is assembled.
 */
async function processMultiChunkResults(xlsBuffers, { applyClassify = false, cloneBuf = null, migType = "conversion", tadirSet = null, tfdirSet = null } = {}) {
  // Read rows from all chunks for analysis; null each buffer immediately after use.
  const allRows = [];

  for (let ci = 0; ci < xlsBuffers.length; ci++) {
    const buf = xlsBuffers[ci];
    if (!buf) continue;

    const wb   = XLSX.read(buf, { type: "buffer", cellText: true, raw: false });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (rows.length > 0) {
      console.log(`[processMultiChunkResults] Chunk ${ci} raw column names:`, Object.keys(rows[0]).join(" | "));
    }
    const normalised = rows.map(_normaliseRow);
    if (normalised.length > 0) {
      const chunkHca = normalised.filter(r => (r["HCA/S4H?"] || "").trim() === "HCA").length;
      const chunkS4  = normalised.filter(r => (r["HCA/S4H?"] || "").trim() === "S/4H").length;
      console.log(`[processMultiChunkResults] Chunk ${ci}: rows=${normalised.length} hca=${chunkHca} s4=${chunkS4}`);
    }
    allRows.push(...normalised);

    // Free the raw buffer immediately
    xlsBuffers[ci] = null;
  }

  console.log(`[processMultiChunkResults] Total rows across all chunks: ${allRows.length}`);
  if (allRows.length > 0) {
    console.log("[processMultiChunkResults] First normalised row 'Object name':", allRows[0]["Object name"]);
  }

  // Build classified XLSX from all rows before freeing them
  let xlsBuf = null;
  let dataResult = { atcData: null, error: "No data rows" };
  if (allRows.length > 0) {
    try {
      // Cross-chunk Fit Gap propagation: reclassify all rows of objects that have any Fit Gap row.
      // Cross-chunk Syntax Error propagation runs first so it takes priority over Fit Gap.
      // Must run on the merged row set — per-chunk propagation in classifyRows cannot cross chunk boundaries.
      // IMPORTANT: propagation must run BEFORE _computeFromRows so that hcaMandatory /
      // s4TechRemediable counts exclude rows whose objects were flagged via cross-chunk propagation.
      const atcClassify = require("./atcClassify");
      atcClassify.propagateSyntaxError(allRows);
      atcClassify.propagateFitGap(allRows);
      dataResult = _computeFromRows(allRows);
      xlsBuf = await buildClassifiedXlsx(allRows);
      console.log(`[processMultiChunkResults] Built classified XLSX: ${xlsBuf.length} bytes`);
    } catch (e) {
      console.error("[processMultiChunkResults] Failed to build classified XLSX:", e.message);
    }
  }

  // Free allRows after use
  allRows.length = 0;

  return { dataResult, xlsBuf };
}

/**
 * Run the full ATC classification pipeline on a pre-parsed rows array.
 * Identical result to readXlsData() — both delegate to _computeFromRows().
 * Used by the multi-chunk path where rows are already available.
 */
function readXlsDataFromRows(rows) {
  try {
    if (!rows.length) return { atcData: null, error: "No data rows in XLS" };
    return _computeFromRows(rows);
  } catch (err) {
    return { atcData: null, error: err.message };
  }
}

/**
 * Remove duplicate columns from an XLSX buffer before sending to ABAP.
 *
 * ABAP's create_dynamic_table crashes (Message E 0K 000) when two columns share
 * the same name because it appends duplicate fieldnames to the field catalogue.
 *
 * Strategy: parse the header row to find which column indices are duplicates
 * (keep the first occurrence, drop all later ones), then rewrite every row in
 * the sheet XML to remove the duplicate cell elements.
 *
 * Operates at the XML level (via JSZip) so it works for any file size without
 * loading the full sheet into a JS row array.
 *
 * Returns a Promise<Buffer> — the deduplicated XLSX buffer. If the file has no
 * duplicates or cannot be processed, returns the original buffer unchanged.
 */
async function deduplicateXlsxColumns(xlsBuf) {
  const JSZip = require("jszip");

  let zip;
  try {
    zip = await JSZip.loadAsync(xlsBuf);
  } catch (_) { return xlsBuf; }

  const sheetEntry = zip.file("xl/worksheets/sheet1.xml")
                  || zip.file(/xl\/worksheets\/sheet\d+\.xml/)[0];
  if (!sheetEntry) return xlsBuf;

  const xmlStr = await sheetEntry.async("string");

  // ── Extract shared strings so we can resolve cell text for header detection ─
  const ssEntry = zip.file("xl/sharedStrings.xml");
  const ssXml   = ssEntry ? await ssEntry.async("string") : "";
  const ssValues = [];
  if (ssXml) {
    const siRegex = /<si>([\s\S]*?)<\/si>/g;
    let m;
    while ((m = siRegex.exec(ssXml)) !== null) {
      // Extract all <t> text inside each <si>, concatenate
      const tMatches = m[1].match(/<t[^>]*>([^<]*)<\/t>/g) || [];
      ssValues.push(tMatches.map(t => t.replace(/<[^>]+>/g, "")).join(""));
    }
  }

  // ── Parse row elements ────────────────────────────────────────────────────
  const rowRegex = /<row\b[^>]*>[\s\S]*?<\/row>/g;
  const rowMatches = xmlStr.match(rowRegex);
  if (!rowMatches || rowMatches.length === 0) return xlsBuf;

  const headerXml = rowMatches[0];

  // ── Resolve a cell's text value (shared string or inline string or numeric) ─
  function _cellText(cellXml) {
    const tAttr = (cellXml.match(/\bt="([^"]*)"/) || [])[1] || "";
    const vMatch = cellXml.match(/<v>([^<]*)<\/v>/);
    const val    = vMatch ? vMatch[1] : "";
    if (tAttr === "s") return ssValues[parseInt(val, 10)] || "";
    if (tAttr === "inlineStr") {
      const it = cellXml.match(/<t[^>]*>([^<]*)<\/t>/);
      return it ? it[1] : "";
    }
    return val;
  }

  // ── Parse header cells ────────────────────────────────────────────────────
  const cellRegex = /<c\b[^>]*>[\s\S]*?<\/c>/g;
  const headerCells = headerXml.match(cellRegex) || [];

  // Build set of seen column names; record which 0-based indices are duplicates
  const seen = new Set();
  const dropIndices = new Set();   // 0-based column position to drop

  for (let i = 0; i < headerCells.length; i++) {
    const name = _cellText(headerCells[i]).trim().toUpperCase();
    if (!name) continue;  // empty header cell — ABAP stops at first empty anyway
    if (seen.has(name)) {
      dropIndices.add(i);
      console.log(`[deduplicateXlsxColumns] Dropping duplicate column at index ${i}: "${name}"`);
    } else {
      seen.add(name);
    }
  }

  if (dropIndices.size === 0) return xlsBuf;  // no duplicates — return unchanged

  // ── Rewrite every row, removing cells at dropIndices ─────────────────────
  // We match cells in order and drop those at duplicate positions.
  function _dropCellsFromRow(rowXml) {
    const cells = rowXml.match(cellRegex) || [];
    const kept  = cells.filter((_, i) => !dropIndices.has(i));
    // Replace the entire cell block (everything between <row ...> and </row>)
    return rowXml.replace(/(<row\b[^>]*>)([\s\S]*)(<\/row>)/,
                          (_, open, _body, close) => open + kept.join("") + close);
  }

  const newRows = rowMatches.map(_dropCellsFromRow);

  // ── Rebuild sheet XML using position-tracked replacement ─────────────────
  // Re-run the regex to get exact start positions for each row match, then
  // build the new XML by splicing — avoids indexOf ambiguity when two rows
  // share identical XML.
  const rowPositions = [];
  const rowRe = /<row\b[^>]*>[\s\S]*?<\/row>/g;
  let rm;
  while ((rm = rowRe.exec(xmlStr)) !== null) {
    rowPositions.push({ start: rm.index, end: rm.index + rm[0].length });
  }

  // Build newXml by concatenating unchanged segments and replacement rows
  const parts = [];
  let cursor = 0;
  for (let i = 0; i < rowPositions.length; i++) {
    const { start, end } = rowPositions[i];
    parts.push(xmlStr.slice(cursor, start));  // text before this row
    parts.push(newRows[i]);                   // replacement row
    cursor = end;
  }
  parts.push(xmlStr.slice(cursor));           // text after last row
  const newXml = parts.join("");

  // ── Rebuild ZIP — copy all base entries as raw buffers, then replace sheet ─
  const newZip = new JSZip();
  const baseEntries = {};
  zip.forEach((relPath, file) => {
    if (relPath !== "xl/worksheets/sheet1.xml") baseEntries[relPath] = file;
  });
  for (const [relPath, file] of Object.entries(baseEntries)) {
    newZip.file(relPath, await file.async("nodebuffer"));
  }
  newZip.file("xl/worksheets/sheet1.xml", newXml);

  return newZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 1 } });
}

// ── CCM Agent Eligible Count (D9) ─────────────────────────────────────────────
// Counts S/4H ATC findings eligible for CCM Agent code proposals.
// Population: S/4H remediable scope rows (same base as C5/D5) that pass CCM
// title/message/note rules AND are NOT already counted in D5.
//
// Remediable scope guard (mirrors D5 base filter exactly):
//   HCA/S4H?=S/4H, Remediation Type=NEEDS REMEDIATION,
//   no syntax error, no Fit Gap?=Yes, no Clone?=Yes, no third-party obj (starts with /)
//
// CCM rules (ALL must pass):
//   1. Check Title is one of 3 applicable titles
//   2. For "S/4HANA Search for Usage of Simplified objects": Check Message must be
//      one of the 4 supported messages; the other 2 titles have no message restriction
//   3. SAP Note must NOT be in the 40-note exclusion list
//
// D5 exclusion: rows already matched by D5 note/refObj filters are subtracted
// so D9 + D5 never double-count the same finding.
function computeCcmAgentEligible(rows) {
  const _APPLICABLE_TITLES = new Set([
    "S/4HANA FIELD LENGTH EXTENSIONS",
    "S/4HANA: FIELD LENGTH EXTENSIONS",
    "S/4HANA SEARCH FOR DATABASE OPERATIONS",
    "S/4HANA: SEARCH FOR DATABASE OPERATIONS",
    "S/4HANA SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
  ]);

  const _SIMPLIFIED_OBJS_TITLES = new Set([
    "S/4HANA SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
  ]);

  const _SUPPORTED_MESSAGES = new Set([
    "FUNCTIONALITY DEPRECATED (ALTERNATIVE EXISTS)",
    "SYNTACTICALLY INCOMPATIBLE CHANGE OF EXISTING FUNCTIONALITY",
    "NON-STRATEGIC-FUNCTION (ALTERNATIVE EXISTS)",
    "FUNCTIONALITY UNAVAILABLE (ALTERNATIVE EXISTS)",
  ]);

  const _EXCLUDED_NOTES = new Set([
    "2217124", "2227014", "2221717", "1976487", "2270335", "2270387", "2602107", "2480067",
    "3256877", "2223144", "3159740", "2469353", "2368913", "2370131", "2371539", "2270199",
    "2371631", "2267292", "2365665", "2368835", "2206980", "2337368", "2224144", "2368747",
    "2227532", "2227579", "2267246", "2268085", "3211383", "2224778", "2250183", "2669781",
    "2669857", "2670006", "2268063", "2267333", "2209696", "2214585",
  ]);

  // D5 note/refObj filters — replicated exactly from d5AutoCount so we can exclude
  // any row that was already counted there.
  const _D5_FLE_MSG_KEYWORDS = ["WRITE", "IMPORT", "EXPORT", "TRANSFER", "READ-DATASET", "READ DATASET", "REPLACE"];
  const _D5_FLE_NOTES        = new Set(["2215424", "2610650"]);
  const _D5_SD_REF_OBJS      = new Set(["VBFA", "VBUK", "VBUP", "VBTYP", "LIKPUK", "RV_ORDER_FLOW_INFORMATION"]);
  const _D5_BILL_REF_OBJS    = new Set(["VBRK", "VBRP", "VBELN", "DRAFT"]);
  const _D5_KONV_REF_OBJS    = new Set(["KONV", "KOLNR"]);
  const _D5_AUTO_OBJ_TYPES   = new Set(["PROG", "FUGR", "CLAS", "FUGS", "FUGX"]);

  function _isD5Row(r) {
    const obj    = (r["Obj."]              || "").trim().toUpperCase();
    const note   = (r["Note"]              || "").trim();
    const refObj = (r["Referenced Object"] || "").trim().toUpperCase();
    const title  = (r["Check Title"]       || "").trim().toUpperCase();
    const msg    = (r["Check Message"]     || "").trim().toUpperCase();
    if (!_D5_AUTO_OBJ_TYPES.has(obj)) return false;
    if (_D5_FLE_NOTES.has(note) && title.includes("FIELD LENGTH") &&
        _D5_FLE_MSG_KEYWORDS.some(k => msg.includes(k))) return true;
    if (note === "2198647" && _D5_SD_REF_OBJS.has(refObj))   return true;
    if (note === "2768887" && _D5_BILL_REF_OBJS.has(refObj)) return true;
    if (note === "2220005" && _D5_KONV_REF_OBJS.has(refObj)) return true;
    if (note === "2431747" && refObj === "BSEG")              return true;
    return false;
  }

  let count = 0;
  for (const row of rows) {
    // ── Remediable scope base guard (mirrors D5 base filter exactly) ─────────
    if ((row["HCA/S4H?"] || "").trim() !== "S/4H")                       continue;
    if ((row["Remediation Type"] || "").trim().toUpperCase() !== "NEEDS REMEDIATION") continue;
    if (_isSyntaxError(row))                                               continue;
    if ((row["Fit Gap?"]    || "").trim() === "Yes")                       continue;
    if ((row["Clone?"]      || "").trim() === "Yes")                       continue;
    if ((row["Object name"] || "").trim().startsWith("/"))                 continue;

    // ── D5 exclusion: skip rows already counted in D5 ────────────────────────
    if (_isD5Row(row)) continue;

    const title = (row["Check Title"]   || "").trim().toUpperCase();
    const msg   = (row["Check Message"] || "").trim().toUpperCase();
    const note  = (row["Note"]          || "").trim();

    // ── CCM Rule 1: applicable Check Title ────────────────────────────────────
    if (!_APPLICABLE_TITLES.has(title)) continue;

    // ── CCM Rule 2: message restriction for "Simplified objects" title ────────
    if (_SIMPLIFIED_OBJS_TITLES.has(title) && !_SUPPORTED_MESSAGES.has(msg)) continue;

    // ── CCM Rule 3: SAP Note not in exclusion list ────────────────────────────
    if (_EXCLUDED_NOTES.has(note)) continue;

    count++;
  }
  return count;
}

// ── Automation Fix? row classifier ────────────────────────────────────────────
// Returns true if a row qualifies for D4 (HCA SELECT/ORDER BY), D5 (S4H note/refObj),
// or D9 (CCM Agent eligible) automation.  Mirrors the filter logic in _computeFromRows
// and computeCcmAgentEligible exactly so column values agree with estimation counts.
function _isAutoFixRow(r) {
  const hcas4  = (r["HCA/S4H?"]          || "").trim();
  const rem    = (r["Remediation Type"]   || "").trim().toUpperCase();
  const obj    = (r["Obj."]              || "").trim().toUpperCase();
  const isSyn  = _isSyntaxError(r);
  const isFg   = (r["Fit Gap?"]           || "").trim() === "Yes";
  const isCln  = (r["Clone?"]            || "").trim() === "Yes";
  const isTp   = (r["Object name"]        || "").trim().startsWith("/");

  const _OBJ_TYPES = new Set(["PROG", "FUGR", "CLAS", "FUGS", "FUGX"]);

  // ── D4: HCA SELECT/ORDER BY ──────────────────────────────────────────────────
  if (hcas4 === "HCA" && !isSyn && !isFg && !isCln && !isTp &&
      _OBJ_TYPES.has(obj) && rem === "MANDATORY") {
    const title = (r["Check Title"] || "").trim().toUpperCase();
    if (title.includes("SELECT") && title.includes("ORDER BY")) return true;
  }

  // ── D5 & D9: S/4H base guard ─────────────────────────────────────────────────
  if (hcas4 !== "S/4H" || rem !== "NEEDS REMEDIATION") return false;
  if (!_OBJ_TYPES.has(obj)) return false;
  if (isSyn || isFg || isCln || isTp) return false;

  const note   = (r["Note"]              || "").trim();
  const refObj = (r["Referenced Object"] || "").trim().toUpperCase();
  const title  = (r["Check Title"]       || "").trim().toUpperCase();
  const msg    = (r["Check Message"]     || "").trim().toUpperCase();

  // ── D5 filters ───────────────────────────────────────────────────────────────
  const _FLE_NOTES     = new Set(["2215424", "2610650"]);
  const _FLE_KEYWORDS  = ["WRITE", "IMPORT", "EXPORT", "TRANSFER", "READ-DATASET", "READ DATASET", "REPLACE"];
  const _SD_REF_OBJS   = new Set(["VBFA", "VBUK", "VBUP", "VBTYP", "LIKPUK", "RV_ORDER_FLOW_INFORMATION"]);
  const _BILL_REF_OBJS = new Set(["VBRK", "VBRP", "VBELN", "DRAFT"]);
  const _KONV_REF_OBJS = new Set(["KONV", "KOLNR"]);

  if (_FLE_NOTES.has(note) && title.includes("FIELD LENGTH") && _FLE_KEYWORDS.some(k => msg.includes(k))) return true;
  if (note === "2198647" && _SD_REF_OBJS.has(refObj))   return true;
  if (note === "2768887" && _BILL_REF_OBJS.has(refObj)) return true;
  if (note === "2220005" && _KONV_REF_OBJS.has(refObj)) return true;
  if (note === "2431747" && refObj === "BSEG")           return true;

  // ── D9 (CCM Agent): not already D5, applicable title, message, note checks ───
  const _CCM_TITLES = new Set([
    "S/4HANA FIELD LENGTH EXTENSIONS",
    "S/4HANA: FIELD LENGTH EXTENSIONS",
    "S/4HANA SEARCH FOR DATABASE OPERATIONS",
    "S/4HANA: SEARCH FOR DATABASE OPERATIONS",
    "S/4HANA SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
  ]);
  const _SIMPLIFIED_TITLES = new Set([
    "S/4HANA SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
  ]);
  const _SUPPORTED_MSGS = new Set([
    "FUNCTIONALITY DEPRECATED (ALTERNATIVE EXISTS)",
    "SYNTACTICALLY INCOMPATIBLE CHANGE OF EXISTING FUNCTIONALITY",
    "NON-STRATEGIC-FUNCTION (ALTERNATIVE EXISTS)",
    "FUNCTIONALITY UNAVAILABLE (ALTERNATIVE EXISTS)",
  ]);
  const _EXCLUDED_NOTES = new Set([
    "2217124", "2227014", "2221717", "1976487", "2270335", "2270387", "2602107", "2480067",
    "3256877", "2223144", "3159740", "2469353", "2368913", "2370131", "2371539", "2270199",
    "2371631", "2267292", "2365665", "2368835", "2206980", "2337368", "2224144", "2368747",
    "2227532", "2227579", "2267246", "2268085", "3211383", "2224778", "2250183", "2669781",
    "2669857", "2670006", "2268063", "2267333", "2209696", "2214585",
  ]);

  if (!_CCM_TITLES.has(title)) return false;
  if (_SIMPLIFIED_TITLES.has(title) && !_SUPPORTED_MSGS.has(msg)) return false;
  if (_EXCLUDED_NOTES.has(note)) return false;

  return true;
}

// ── CCM Agent Fix? row classifier ────────────────────────────────────────────
// Returns true only for D9 (CCM Agent eligible) findings.
// Mirrors the D9 branch of _isAutoFixRow() and computeCcmAgentEligible() exactly.
// D5 rows are explicitly excluded — Automation Fix? and CCM Agent Fix? never both = Yes for S/4H rows.
function _isCcmAgentFixRow(r) {
  const hcas4  = (r["HCA/S4H?"]        || "").trim();
  const rem    = (r["Remediation Type"] || "").trim().toUpperCase();
  const obj    = (r["Obj."]            || "").trim().toUpperCase();
  const _OBJ_TYPES = new Set(["PROG", "FUGR", "CLAS", "FUGS", "FUGX"]);

  if (hcas4 !== "S/4H" || rem !== "NEEDS REMEDIATION") return false;
  if (!_OBJ_TYPES.has(obj))                             return false;
  if (_isSyntaxError(r))                                return false;
  if ((r["Fit Gap?"]    || "").trim() === "Yes")        return false;
  if ((r["Clone?"]      || "").trim() === "Yes")        return false;
  if ((r["Object name"] || "").trim().startsWith("/"))  return false;

  const note   = (r["Note"]              || "").trim();
  const title  = (r["Check Title"]       || "").trim().toUpperCase();
  const msg    = (r["Check Message"]     || "").trim().toUpperCase();
  const refObj = (r["Referenced Object"] || "").trim().toUpperCase();

  // Exclude D5 rows — must not overlap with Automation Fix?
  const _FLE_NOTES     = new Set(["2215424", "2610650"]);
  const _FLE_KEYWORDS  = ["WRITE", "IMPORT", "EXPORT", "TRANSFER", "READ-DATASET", "READ DATASET", "REPLACE"];
  const _SD_REF_OBJS   = new Set(["VBFA", "VBUK", "VBUP", "VBTYP", "LIKPUK", "RV_ORDER_FLOW_INFORMATION"]);
  const _BILL_REF_OBJS = new Set(["VBRK", "VBRP", "VBELN", "DRAFT"]);
  const _KONV_REF_OBJS = new Set(["KONV", "KOLNR"]);
  if (_FLE_NOTES.has(note) && title.includes("FIELD LENGTH") && _FLE_KEYWORDS.some(k => msg.includes(k))) return false;
  if (note === "2198647" && _SD_REF_OBJS.has(refObj))   return false;
  if (note === "2768887" && _BILL_REF_OBJS.has(refObj)) return false;
  if (note === "2220005" && _KONV_REF_OBJS.has(refObj)) return false;
  if (note === "2431747" && refObj === "BSEG")           return false;

  const _CCM_TITLES = new Set([
    "S/4HANA FIELD LENGTH EXTENSIONS",        "S/4HANA: FIELD LENGTH EXTENSIONS",
    "S/4HANA SEARCH FOR DATABASE OPERATIONS", "S/4HANA: SEARCH FOR DATABASE OPERATIONS",
    "S/4HANA SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
  ]);
  const _SIMPLIFIED_TITLES = new Set([
    "S/4HANA SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGE OF SIMPLIFIED OBJECTS",
    "S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS",
  ]);
  const _SUPPORTED_MSGS = new Set([
    "FUNCTIONALITY DEPRECATED (ALTERNATIVE EXISTS)",
    "SYNTACTICALLY INCOMPATIBLE CHANGE OF EXISTING FUNCTIONALITY",
    "NON-STRATEGIC-FUNCTION (ALTERNATIVE EXISTS)",
    "FUNCTIONALITY UNAVAILABLE (ALTERNATIVE EXISTS)",
  ]);
  const _EXCLUDED_NOTES = new Set([
    "2217124", "2227014", "2221717", "1976487", "2270335", "2270387", "2602107", "2480067",
    "3256877", "2223144", "3159740", "2469353", "2368913", "2370131", "2371539", "2270199",
    "2371631", "2267292", "2365665", "2368835", "2206980", "2337368", "2224144", "2368747",
    "2227532", "2227579", "2267246", "2268085", "3211383", "2224778", "2250183", "2669781",
    "2669857", "2670006", "2268063", "2267333", "2209696", "2214585",
  ]);

  if (!_CCM_TITLES.has(title))                                     return false;
  if (_SIMPLIFIED_TITLES.has(title) && !_SUPPORTED_MSGS.has(msg)) return false;
  if (_EXCLUDED_NOTES.has(note))                                   return false;
  return true;
}

/**
 * Build a classified results XLSX buffer from an array of (already-classified) rows.
 * Columns: Original 14 ATC Extract columns + 6 classification columns.
 *
 * Uses direct sheet XML construction via JSZip instead of SheetJS json_to_sheet(),
 * which builds a giant cell-map object in RAM.  Writing raw XML is ~3-5x more
 * memory-efficient and has no upper row-count limit.
 *
 * Returns a Buffer (XLSX format).
 */
async function buildClassifiedXlsx(rows) {
  const ORIG_COLS = [
    "Priority", "Check Title", "Check Message", "Object name", "Obj.",
    "Exemption State", "Contact", "Package", "1st Found", "Note",
    "Short text", "Component ID", "Ref. Object Type", "Referenced Object",
  ];
  const CLASS_COLS = [
    "HCA/S4H?", "Remediation Type", "NP Remediation Type", "Syntax Error?", "Fit Gap?", "Clone?",
    "Manual Check/Delta Fit gap", "Automation Fix?", "CCM Agent Fix?",
  ];
  const COLS = [...ORIG_COLS, ...CLASS_COLS];
  const numCols = COLS.length;

  // ── Column letter helper (0-based index → "A", "B", …, "Z", "AA", …) ──────
  function _colLetter(idx) {
    let s = "";
    idx++;
    while (idx > 0) {
      idx--;
      s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[idx % 26] + s;
      idx = Math.floor(idx / 26);
    }
    return s;
  }

  // Precompute column letters once
  const colLetters = COLS.map((_, i) => _colLetter(i));
  const lastCol    = colLetters[numCols - 1];

  // ── XML escape helper ────────────────────────────────────────────────────────
  function _esc(v) {
    return String(v)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      // Strip control characters that are illegal in XML 1.0
      // (keep tab \x09, newline \x0A, carriage return \x0D)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  }

  // ── Hidden column indices (0-based) ─────────────────────────────────────────
  const HIDDEN_COLS = new Set(
    COLS.map((name, i) => name === "NP Remediation Type" ? i : -1).filter(i => i >= 0)
  );

  // ── Build sheet XML row by row ───────────────────────────────────────────────
  const totalRows = rows.length + 1; // +1 for header

  // Build <cols> element to hide "NP Remediation Type" column in Excel
  // Use hidden="1" alone (without width="0") — the correct OOXML way to hide columns.
  // customWidth="1" is required whenever any col attribute is set.
  const colsXml = HIDDEN_COLS.size > 0
    ? `<cols>${[...HIDDEN_COLS].map(i => {
        const colNum = i + 1; // 1-based
        return `<col min="${colNum}" max="${colNum}" hidden="1" customWidth="1" width="9.14"/>`;
      }).join("")}</cols>`
    : "";

  const xmlParts  = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
    `<sheetPr/>`,
    `<dimension ref="A1:${lastCol}${totalRows}"/>`,
    `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`,
    `<sheetFormatPr defaultColWidth="9.14" defaultRowHeight="15"/>`,
    colsXml,
    `<sheetData>`,
  ];

  // Header row
  const headerCells = COLS.map((name, ci) =>
    `<c r="${colLetters[ci]}1" t="inlineStr"><is><t>${_esc(name)}</t></is></c>`
  ).join("");
  xmlParts.push(`<row r="1">${headerCells}</row>`);

  // Data rows
  for (let ri = 0; ri < rows.length; ri++) {
    const r       = rows[ri];
    const rowNum  = ri + 2;
    const cells   = COLS.map((col, ci) => {
      let val = r[col] !== undefined ? r[col] : "";
      if (col === "Manual Check/Delta Fit gap") val = (r["Fit Gap Delta?"] === "Yes") ? "Yes" : "No";
      if (col === "Fit Gap?") val = (r["Fit Gap?"] === "Yes" || r["Remediation Type"] === "Fit Gap") ? "Yes" : "No";
      if (col === "Automation Fix?") val = _isAutoFixRow(r) ? "Yes" : "No";
      if (col === "CCM Agent Fix?")  val = _isCcmAgentFixRow(r) ? "Yes" : "No";
      return `<c r="${colLetters[ci]}${rowNum}" t="inlineStr"><is><t>${_esc(val)}</t></is></c>`;
    }).join("");
    xmlParts.push(`<row r="${rowNum}">${cells}</row>`);
  }

  xmlParts.push(`</sheetData></worksheet>`);
  const sheetXml = xmlParts.join("");

  // ── Assemble minimal XLSX ZIP ────────────────────────────────────────────────
  const JSZip = require("jszip");
  const zip   = new JSZip();

  zip.file("[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
    `</Types>`
  );

  zip.file("_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `</Relationships>`
  );

  zip.file("xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<sheets><sheet name="ATC Results" sheetId="1" r:id="rId1"/></sheets>` +
    `</workbook>`
  );

  zip.file("xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
    `</Relationships>`
  );

  zip.file("xl/worksheets/sheet1.xml", sheetXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 1 } });
}

module.exports = { readXlsData, computeAtcClassification, computeCcmAgentEligible, splitXlsxRows, mergeXlsResults, mergeXlsResultBufsToRows, readXlsDataFromRows, deduplicateXlsxColumns, processMultiChunkResults, buildClassifiedXlsx };
