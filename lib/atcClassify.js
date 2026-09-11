"use strict";
/**
 * atcClassify.js — Node.js port of ZCL_ZATC_ASSESSMENT_HANDLER classification logic.
 *
 * Adds these columns to every row (in-place):
 *   "HCA/S4H?"        — "HCA" or "S/4H"
 *   "Remediation Type"— Mandatory | Needs Remediation | False Positive |
 *                       Fit Gap | Can be ignored | Optional | Syntax Error
 *   "Syntax Error?"   — "Yes" / "No"
 *   "Fit Gap?"        — "Yes" / "No"
 *   "Clone?"          — "Yes" / "No"  (requires cloneRows to be passed)
 *
 * Input rows must already be normalised via _normaliseRow (xlsData._COL_CANONICAL).
 * Clone rows are the rows from the clone XLSX (OBJ_TYPE + OBJ_NAME columns).
 */

const XLSX = require("xlsx");

// ── Embedded TADIR/TFDIR reference data (offline orphan detection) ────────────
// Lazily loaded on first use — avoids OOM during app startup in memory-constrained environments.
// Uses sorted array + binary search instead of Set to reduce memory by ~4x.
// The gz files contain JSON arrays that are already sorted alphabetically.
function _bsearch(arr, key) {
  let lo = 0, hi = arr.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (arr[mid] === key) return true;
    if (arr[mid] < key) lo = mid + 1; else hi = mid - 1;
  }
  return false;
}

// Thin wrapper so callers can use .has() and .size like a Set
class SortedArray {
  constructor(arr) { this._a = arr; this.size = arr.length; }
  has(key) { return _bsearch(this._a, key); }
}

function _loadRefSet(baseName) {
  const path = require("path");
  const fs   = require("fs");
  const gz   = path.join(__dirname, baseName + ".gz");
  const json = path.join(__dirname, baseName + ".json");
  let arr;
  if (fs.existsSync(gz)) {
    const zlib = require("zlib");
    arr = JSON.parse(zlib.gunzipSync(fs.readFileSync(gz)).toString("utf8"));
  } else {
    arr = JSON.parse(fs.readFileSync(json, "utf8"));
  }
  // Ensure sorted (data from dump script is sorted, but guard anyway)
  if (arr.length > 1 && arr[0] > arr[arr.length - 1]) arr.sort();
  return new SortedArray(arr);
}
let _TADIR_SET = null;
let _TFDIR_SET = null;
function _getTadirSet() { if (!_TADIR_SET) _TADIR_SET = _loadRefSet("tadirData"); return _TADIR_SET; }
function _getTfdirSet() { if (!_TFDIR_SET) _TFDIR_SET = _loadRefSet("tfdirData"); return _TFDIR_SET; }

// ── Clone file type-name → SAP short code mapping ────────────────────────────
const _CLONE_TYPE_MAP = {
  "REPORT SOURCE CODE":           "PROG",
  "PROGRAM":                      "PROG",
  "INCLUDE PROGRAM":              "PROG",
  "INCLUDE":                      "PROG",
  "SAP SMART FORM":               "SSFO",
  "FORM":                         "SSFO",
  "FUNCTION MODULE":              "FUGR",
  "FUNCTION GROUP":               "FUGR",
  "CLASS (ABAP OBJECTS)":         "CLAS",
  "CLASS":                        "CLAS",
  "INTERFACE (ABAP OBJECTS)":     "INTF",
  "INTERFACE":                    "INTF",
  "TABLE":                        "TABL",
  "DATABASE TABLE":               "TABL",
  "DATA ELEMENT":                 "DTEL",
  "DOMAIN":                       "DOMA",
  "MESSAGE CLASS":                "MSAG",
  "TRANSACTION":                  "TRAN",
  "TYPE GROUP":                   "TYPE",
  "ADOBE FORM":                   "SFPF",
  "WEB DYNPRO COMPONENT":         "WDYN",
  "ENHANCEMENT IMPLEMENTATION":   "ENHO",
};

function buildCloneSet(cloneRows) {
  const set = new Set();
  if (!cloneRows || !cloneRows.length) return set;

  // Detect which key names hold type/name in the first row (case-insensitive header match)
  // Falls back to first/second column by position if no known header is found.
  const firstRow = cloneRows[0];
  const keys = Object.keys(firstRow);

  const _TYPE_KEYS = ["obj_type", "objtype", "object type", "obj.", "type"];
  const _NAME_KEYS = ["obj_name", "objname", "object name", "object_name", "name"];

  let typeKey = null;
  let nameKey = null;
  for (const k of keys) {
    const kl = k.trim().toLowerCase();
    if (!typeKey && _TYPE_KEYS.includes(kl)) typeKey = k;
    if (!nameKey && _NAME_KEYS.includes(kl)) nameKey = k;
  }
  // Positional fallback: column 0 = type, column 1 = name
  if (!typeKey && keys.length >= 1) typeKey = keys[0];
  if (!nameKey && keys.length >= 2) nameKey = keys[1];

  if (typeKey) {
    console.log(`[buildCloneSet] Using typeKey="${typeKey}" nameKey="${nameKey}" from ${cloneRows.length} clone rows`);
  }

  for (const row of cloneRows) {
    const rawType = typeKey ? (row[typeKey] || "").toString().trim().toUpperCase() : "";
    const rawName = nameKey ? (row[nameKey] || "").toString().trim().toUpperCase() : "";
    if (!rawType || !rawName) continue;
    const mappedType = _CLONE_TYPE_MAP[rawType] || rawType;
    set.add(`${mappedType},${rawName}`);
  }
  return set;
}

function readCloneRows(cloneBuf) {
  if (!cloneBuf) return [];
  try {
    const wb   = XLSX.read(cloneBuf, { type: "buffer", cellText: true, raw: false });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
  } catch (_) {
    return [];
  }
}

// ── Hard-conflict check message detector (mirrors lv_check='X' in ZATCASSESSMENT_KK) ──────────
// Returns true when the check message is a hard type/length conflict that requires code change
// even when the referenced object IS available in S/4HANA (TADIR/TFDIR found).
// Maps to the WHEN block that sets lv_check = 'X' (lines 3276-3285 of ZATCASSESSMENT_KK).
function _isHardConflict(checkMsg) {
  const m = checkMsg.toUpperCase();
  return (
    m.includes("CONCATENATE LENGTH CONFLICT") ||
    m.includes("TYPE-CONFLICT") ||
    m.includes("TYPE CONFLICT") ||
    m.includes("MOVE LENGTH CONFLICT") ||
    m.includes("STRUCTURE-COMPONENT TYPE CONFLICT") ||
    m.includes("STRUCTURE COMPONENT TYPE CONFLICT") ||
    m.includes("SYNTACTICALLY INCOMPATIBLE CHANGE") ||
    m.includes("ARITHMETIC TYPE CONFLICT") ||
    m.includes("ARIHMETIC TYPE CONFLICT") ||
    m.includes("WRITE TO TYPE CONFLICT") ||
    m.includes("USAGE OF NO LONGER AVAILABLE DEVELOPMENT OBJECT") ||
    m.includes("TRANSFER ISSUE") ||
    m.includes("COMPARE TYPE CONFLICT") ||
    m.includes("READ DATASET ISSUE") ||
    m.includes("REPLACE ISSUE") ||
    m.includes("WRITE ISSUE") ||
    m.includes("COMPARE LENGTH CONFLICT") ||
    m.includes("STRUCTURE-COMPONENT LENGTH CONFLICT") ||
    m.includes("STRUCTURE COMPONENT LENGTH CONFLICT") ||
    m.includes("MOVE TYPE CONFLICT") ||
    m.includes("OFFSET/LENGTH-ACCESS") ||
    m.includes("OFFSET LENGTH ACCESS") ||
    m.includes("MESSAGE-WITH LENGTH CONFLICT") ||
    m.includes("MESSAGE WITH LENGTH CONFLICT") ||
    m.includes("NON-STRATEGIC-FUNCTION: NO FUNCTIONAL EQUIVALENT") ||
    m.includes("NON STRATEGIC FUNCTION NO FUNCTIONAL EQUIVALENT") ||
    m.includes("WRITE-LENGTH ISSUE") ||
    m.includes("WRITE LENGTH ISSUE") ||
    m.includes("SELECTION-SCREEN LENGTH CONFLICT") ||
    m.includes("SELECTION SCREEN LENGTH CONFLICT") ||
    m.includes("SELECT TYPE CONFLICT") ||
    m.includes("IS-INITIAL-CHECK FOR COMPONENT") ||
    m.includes("IS INITIAL CHECK FOR COMPONENT") ||
    // FUNCTIONALITY UNAVAILABLE: ATC is explicitly reporting the function/object does not
    // work in S/4HANA even if it physically exists in TFDIR/TADIR. Must be remediated.
    m === "FUNCTIONALITY UNAVAILABLE"
  );
}

// Returns true when the check message maps to lv_check = 'Y' in ZATCASSESSMENT_KK loop 2
// (lines 3289-3293): OLD / RFC / GENERIC / CASTING / DYNAMIC ASSIGN / DYNAMIC DB-ACCESS /
// EXPORT ISSUE / GENERIC CODE ISSUE / CONCATENATION DETECTED / CONSTANT COMPARE CONFLICT /
// DESCRIBE FIELD ISSUE / IMPORT ISSUE / MOVE LENGTH EXTEND / SET PARAMETER ISSUE.
// For TFDIR-found FMs: lv_check='Y' → "Can be ignored". Blank lv_check → default NR.
function _isLCheckY(checkMsg) {
  const m = checkMsg.toUpperCase();
  return (
    m.includes("OLD") ||
    m.includes("RFC") ||
    m.includes("GENERIC") ||
    m.includes("CASTING FROM/TO") ||
    m.includes("DYNAMIC ASSIGN") ||
    m.includes("DYNAMIC DB-ACCESS") ||
    m.includes("EXPORT ISSUE") ||
    m.includes("GENERIC CODE ISSUE") ||
    m.includes("CONCATENATION DETECTED") ||
    m.includes("CONSTANT COMPARE CONFLICT") ||
    m.includes("DESCRIBE FIELD ISSUE") ||
    m.includes("IMPORT ISSUE") ||
    m.includes("MOVE LENGTH EXTEND") ||
    m.includes("SET PARAMETER ISSUE")
  );
}

// ── Main classification engine ────────────────────────────────────────────────
function _classifyRow(row, migType, tadirSet, tfdirSet) {
  const priority    = (row["Priority"]          || "").toString().trim();
  const checkMsg    = (row["Check Message"]     || row["Check message"] || "").toString().trim().toUpperCase();
  const note        = (row["Note"]              || row["SAP Note"] || row["SAP Note Number"] || "").toString().trim();
  const refObjType  = (row["Ref. Object Type"]  || row["Ref Object Type"] || row["RefObjectType"] || "").toString().trim().toUpperCase();
  const refObj    = (row["Referenced Object"] || row["Ref. Object Name"] || "").toString().trim().toUpperCase();
  const objType   = (row["Obj."]              || row["Object Type"] || "").toString().trim().toUpperCase();
  const objName   = (row["Object name"]       || row["Object Name"] || "").toString().trim().toUpperCase();
  const title     = (row["Check Title"]       || row["Check title"] || "").toString().trim().toUpperCase();

  const _allData = [checkMsg, note, refObj, objType, objName, title];
  const _cs  = (s) => _allData.some(d => d.includes(s.toUpperCase()));
  const _eq  = (s) => _allData.some(d => d === s.toUpperCase());
  const _eqN = (s) => note.trim() === s;  // exact note match

  // S/4H detection: mirrors ABAP logic — any cell in the row contains 'S/4HANA:'
  // ABAP scans all columns for CS 'S/4HANA:' (with colon) to set gv_s4h='X'.
  // Checking _cs covers Check Title, Check Message, and all other fields identically.
  const isS4H = _cs("S/4HANA:");

  const prio12 = priority === "1" || priority === "2";
  const prio3  = priority === "3";
  const isUpgrade = migType === "upgrade";

  let remType = "";
  let remark  = "";
  let isFitGapDelta = false;

  // ── SYNTAX / PRE-EXISTING ERROR ───────────────────────────────────────────
  // ABAP: CS 'SYMBOL ... NOT FOUND' (literal ellipsis in ATC message text)
  // AQQU/AQSG (SAP Query objects) and IDOC objects are Fit Gap even on INTERNAL ERROR / DOES NOT EXIST
  const _isQueryObj = objType === "AQQU" || objType === "AQSG" || objType === "AQAD" || objType === "IDOC";
  if (!_isQueryObj && (
      _cs("INTERNAL ERROR") || _cs("SYNTAX ERROR") || _cs("SYMBOL ... NOT FOUND") ||
      _cs("DOES NOT EXIST") || _cs("SCAN ERROR")   || _cs("UNKNOWN ERROR CODE") ||
      _cs("MISSING INCLUDE"))) {
    remType = "Syntax Error";
    remark  = "Syntax Error Internal Error in Programs are Out Of Scope";
    return { isS4H, remType, remark, isSyntaxError: true, isFitGap: false, isFitGapDelta: false };
  }

  // ── DB OPERATION flags ───────────────────────────────────────────────────
  // ABAP DML flag ('Y'): INSERT, UPDATE, MODIFY, DELETE, SELECT FOR UPDATE
  const isDml   = _cs("DB OPERATION INSERT") || _cs("DB OPERATION UPDATE") ||
                  _cs("DB OPERATION MODIFY") || _cs("DB OPERATION DELETE") ||
                  _cs("DB OPERATION SELECT FOR UPDATE FOUND");
  // ABAP SELECT flag ('Y1'): 'DB OPERATION SELECT FOUND' (exact suffix — excludes DML SELECT FOR UPDATE)
  const isQuery = _cs("DB OPERATION SELECT FOUND") || _cs("DB OPERATION QUERY");
  // ABAP cursor flag ('4'): OPEN/CLOSE/FETCH NEXT CURSOR FOUND
  const isCursor= _cs("DB OPERATION OPEN CURSOR FOUND") || _cs("DB OPERATION CLOSE CURSOR FOUND") ||
                  _cs("DB OPERATION FETCH NEXT CURSOR FOUND");
  const isLiteral = _cs("TRANSACTIONS IN LITERALS") || _cs("TRANSACTION IN LITERALS");

  if (isQuery) {
    if (_eq("MLHD") || _eq("MLCR") || _eq("MLIT") ||
        _eq("BSEG") || _eq("KONV") || _eq("T881")) {
      remType = "Fit Gap";
      remark  = "The new Material Ledger Document tables MLDOC and MLDOCCCS replace most former periodic tables.";
    } else if (_eq("BPEG") || _eq("GLPCP") || _eq("CKMLCR")) {
      remType = "Fit Gap";
      remark  = "Changes Required in Select Statement with help of Functional Team.";
    } else if (_eq("GLPCA") || _eq("GLPCT")) {
      remType = "False Positive";
      remark  = "";
    } else if (_eq("SKA1") || _eq("SKB1") || _eq("T004")) {
      remType = "False Positive";
      remark  = "No Changes Required in Select Statement.";
    } else if (_eq("VBFA")) {
      remType = "False Positive";
      remark  = "Primary Key changed in VBFA.";
    }
  }

  if (isCursor) {
    if (_cs("VBUK") || _cs("VBUP")) {
      remType = "Needs Remediation";
      remark  = "Changes Required in Select Statement.";
    } else if (_eq("VBFA")) {
      remType = "False Positive";
      remark  = "Changes Required in Select Statement.";
    }
  }

  if (_cs("S/4HANA: IDOC CHECK") || title.includes("IDOC CHECK")) {
    remType = "Fit Gap";
    remark  = "No Reference provided.";
  }

  // ── ORDER BY / LOOP / READ TABLE patterns ────────────────────────────────
  // Note: _cs() searches all fields including the check title.
  // "SEARCH PROBLEMATIC STATEMENTS FOR RESULT OF SELECT/OPEN CURSOR WITHOUT ORDER BY"
  // always contains "ORDER BY" and "SELECT" in the title, so all rows of this check
  // fire this block. LOOP AT EMPTY rows are the exception — they are handled by the
  // LOOP AT + EMPTY block below and should produce Optional (prio3), not Mandatory.
  if (_cs("ORDER BY") && (_cs("WRITE") || _cs("EXIT") || _cs("LEAVE") || _cs("RETURN") ||
      _cs("SELECT") || _cs("ENDSELECT"))) {
    if (!remType && !checkMsg.includes("LOOP AT EMPTY")) {
      remType = "Mandatory";
      remark  = "ORDER BY Missing...Changes Required in select statement.";
    }
  }

  if (_cs("LOOP AT") && (_cs("ORDER BY") || _cs("EQUALITY TEST") || _cs("AT") || _cs("ON") ||
      _cs("MODIFY") || _cs("EMPTY"))) {
    if (!remType) {
      if (prio12) { remType = "Mandatory"; remark = "ORDER BY Missing...Changes Required in this statement."; }
      else if (prio3) { remType = "Optional"; remark = "Changes Not Required in this statement."; }
    }
  }

  if (_cs("LOOP AT") && (_cs("WRITE") || _cs("EXIT") || _cs("RETURN") || _cs("LEAVE"))) {
    // KK: flag='K' via CONTINUE — last-write-wins. This overrides ORDER BY (flag 'L3') which may
    // have fired earlier because the Check Title contains "ORDER BY" + "SELECT". KK processes each
    // cell separately (DO loop with CONTINUE), so 'K' always wins over 'L3' for this pattern.
    if (prio12) { remType = "Mandatory"; remark = "Sort statement required before LOOP"; }
    else if (prio3) { remType = "Optional"; remark = "No change required in LOOP"; }
  }

  if (_cs("WRITE IN LOOP FOR")) {
    // KK: flag='K' via CONTINUE → same K-block handling as LOOP AT EXIT/RETURN/LEAVE.
    if (prio12) { remType = "Mandatory"; remark = "Sort statement required before LOOP"; }
    else if (prio3) { remType = "Optional"; remark = "No change required in LOOP"; }
  }

  if (_cs("ALV CALL AT")) {
    // KK: flag='A' → Optional (any priority), overrides ORDER BY fired on title text.
    remType = "Optional"; remark = "ALV Statement is correct. No Change Required.";
  }

  if (_cs("READ TABLE") && _cs("INDEX") && !_cs("INDEX 1")) {
    if (!remType) {
      if (prio12) { remType = "Mandatory"; remark = "Sort statement required"; }
      else if (prio3) { remType = "Mandatory"; remark = "Sort statement required with READ"; }
    }
  }
  if (_cs("READ TABLE") && _cs("INDEX 1")) {
    // KK: flag='7Z' via CONTINUE — overrides ORDER BY (flag 'L3') which fires on title text.
    remType = "Optional"; remark = "Sort statement can be placed with READ";
  }

  // ── FIELD LENGTH / TYPE CONFLICTS ────────────────────────────────────────
  if (_cs("CONCATENATE IN LOOP")) {
    remType = "Mandatory";
    remark  = "Changes required in Loop statement.";
  } else if (_cs("TYPE CONFLICT") || _cs("LENGTH CONFLICT") || _cs("COMPARE") && _cs("LENGTH") ||
             _cs("WRITE TO") || _cs("TRANSFER") || _cs("REPLACE") || _cs("OFFSET") ||
             _cs("STRUCTURE COMPONENT") || _cs("READ DATASET")) {
    if (!remType) {
      remType = "Needs Remediation";
      remark  = "Changes required in statement.";
    }
  }

  // ── NOTE-BASED RULES ─────────────────────────────────────────────────────

  // Note 2993220 — Classical PCA compatibility
  if (_cs("2993220")) {
    remType = "False Positive";
    remark  = "Classical Profit Center Accounting (PCA) is part of the SAP S/4HANA compatibility scope.";
  }

  // Note 2602107 — Compatibility Views (guard: not DML/Query)
  if (_eqN("2602107") && !isDml && !isQuery) {
    const compViews = new Set(["BSAD","BSAK","BSAS","BSID","BSIK","BSIS","COEP","COSP",
      "GLT0","FAGLFLEXT","FMGLFLEXT","PSGLFLEXT","COSS","COVP","JVGLFLEXT","FAGLBSIS","FAGLBSAS"]);
    if ([...compViews].some(t => _eq(t))) {
      remType = "False Positive";
      remark  = "Select Statement will work with CDS Views.";
    }
  }

  // Note 2438131 — Material Number Field Length Extension (flag '8')
  // Upgrade: FM/BAPI already adapted in S/4HANA — no code change required → FP
  // Conversion: code must be adapted to use _LONG parameters → NR
  if (_cs("2438131")) {
    if (isUpgrade) {
      remType = "False Positive";
      remark  = "Already on S/4HANA — BAPI/RFC extended fields (_LONG) are already in use. No code change required.";
    } else {
      remType = "Needs Remediation";
      remark  = "BAPI or RFC contains Extended Field with addition of <_LONG> like MATNR_LONG.";
    }
  }

  // Note 2610650 — Amount Field Length Extension (flag '9')
  if (_cs("2610650") && !_cs("USED BY RFC")) {
    if (prio3) { remType = "Can be ignored"; remark = "Amount Field Length Extension: Code Adaptations might require."; }
    else       { remType = "Needs Remediation"; remark = "Amount Field Length Extension: Code Adaptations might require."; }
  }

  // Note 2768887 — SD Billing Document Draft
  if (_cs("2768887") && !isDml && !isQuery) {
    remType = "Needs Remediation";
    remark  = "Code will require remediations.";
  }

  // Note 2270387 — Asset Accounting
  if (_cs("2270387") && !isDml) {
    remType = "Needs Remediation";
    remark  = "Code will require remediations.";
  }

  // Note 2431747 — GL Incompatible Changes
  // KK: runs even when lv_flag='Y1' (SELECT) — SELECT+2431747+BSEG → YH → Needs Remediation.
  // Guard only DML (flag Y / UD) — not SELECT (flag Y1).
  if (_cs("2431747") && !isDml) {
    const needsRem = new Set(["BSEG","CSKA","TKA02","T881","FAGL_ACTIVEC","T882G","CSKB","FAGL_LEDGER_SCEN"]);
    const falsePos = new Set(["SKA1","SKB1","T001","RPL_S130_CHANGE_IN_UPDATE_TASK"]);
    const fitGap   = new Set(["AQQU","AQSG"]);
    if ([...needsRem].some(t => _eq(t))) { remType = "Needs Remediation"; remark = "Code will require remediations."; }
    else if ([...falsePos].some(t => _eq(t))) { remType = "False Positive"; remark = "Code will not require remediations."; }
    else if ([...fitGap].some(t => _eq(t)))   { remType = "Fit Gap"; remark = "ABAP SQL Query."; }
  }

  // Notes 2365665 / 2535093 — Seasonal Fields
  // ABAP flag YJ: lv_data='DTEL' → NR; lv_data='TABL' → TBCS (need to check)
  if (_cs("2365665") || _eqN("2535093")) {
    if (refObjType === "DTEL" || _eq("DTEL")) {
      remType = "Needs Remediation"; remark = "Replace data element with new or custom.";
    }
    else if (refObjType === "TABL" || _eq("TABL")) {
      remType = "Fit Gap"; isFitGapDelta = true; remark = "The season fields are no longer maintained in table MARA. But maintained in table FSH_SEASONS_MAT (TBCS).";
    }
  }

  // Note 2220005 — Pricing Condition Technique
  if (_cs("2220005") && !isDml) {
    if      (_eq("DTEL")) { remType = "Needs Remediation"; remark = "Program Change would require as Data Element length increased."; }
    else if (_eq("TABL")) { remType = "Needs Remediation"; remark = "Changes in pricing structure tables."; }
    else if (_eq("KONV")) { remType = "Needs Remediation"; remark = "Changes in pricing structure tables would require change."; }
  }

  // Note 2215424 — Material Number Field Length Extension
  if (_eqN("2215424")) {
    if (prio12) { remType = "Needs Remediation"; remark = "Use of Extended Field length."; }
    else        { remType = "Can be ignored";     remark = "Use of Extended Field length."; }
  }

  // Note 2389136 — Cost Element Changes
  // ABAP (flag HH): CSKA/CSKB/CSKU → NR; SKA1/SKB1/SKAT → FP; T004 → FP; TKSKA → FG; TKA3A → NR
  if (_eqN("2389136") && !isQuery) {
    if      (_eq("CSKA") || _eq("CSKB") || _eq("CSKU")) { remType = "Needs Remediation"; remark = "Do not write to these tables. CSKB-AUFNR and KOSTL are no longer valid. Check table TKA3A"; }
    else if (_eq("SKA1") || _eq("SKB1") || _eq("SKAT")) { remType = "False Positive";    remark = "These tables now contain also secondary cost elements. Do not write to these tables"; }
    else if (_eq("T004"))  { remType = "False Positive";    remark = "T004: The field INTEG_CO is obsolete. All cost elements are G/L accounts as well from now on."; }
    else if (_eq("TKSKA")) { remType = "Fit Gap";            remark = "TABL TKSKA: The table is obsolete and is cleared during the data migration."; }
    else if (_eq("TKA3A")) { remType = "Needs Remediation"; remark = "TABL TKA3A: The table may also contain cost elements and replaces the according fields in CSKB."; }
  }

  // Note 2226131 — Public Sector Business Partner
  if (_eqN("2226131") && !isLiteral) {
    const fitGapTrans = new Set(["XD01","XD02","XK01","XK02","FK01","FK02"]);
    const falsePos    = new Set(["KNA1","LFA1","SHLP"]);
    if ([...fitGapTrans].some(t => _eq(t))) { remType = "Fit Gap";         remark = "TCODE No longer executable and diverted to transaction BP for business partner creation/maintenance."; }
    else if (_eq("TABL"))                   { remType = "Needs Remediation"; remark = "May be False Positive if upgrading HANA lower version to higher version."; }
    else if ([...falsePos].some(t => _eq(t)))  { remType = "False Positive"; remark = "No Change required in object."; }
  }

  // Note 1803189 — Obsolete PO APIs
  // KK: FUNC/BAPI_PO_GETDETAIL → NR; TRAN + ME21/22/23/24/25/27/28/51/52/53/54/59 → FG; ME23/ME26 → NR
  if (_eqN("1803189")) {
    if (_eq("FUNC") || _cs("BAPI_PO_GETDETAIL")) {
      remType = "Needs Remediation"; remark = "The corresponding IDoc/BAPI replacements are available like BAPI_PO_CREATE -> BAPI_PO_CREATE1.";
    } else if ((refObjType === "TRAN" || _eq("TRAN")) && (_eq("ME53") || _eq("ME23") || _eq("ME26"))) {
      remType = "Needs Remediation"; remark = "Transaction no longer available.";
    } else if (refObjType === "TRAN" || _eq("TRAN")) {
      remType = "Fit Gap"; remark = "The mentioned obsolete transactions have been replaced by more user-friendly new transactions.";
    }
  }

  // Note 2265093 — Obsolete Transactions
  if (_eqN("2265093") && !isDml) {
    const fitGapT  = new Set(["FD06","FK06","MK06","MK12","MK18","MK19","VD06","XD06","V+21","V+22","V+23","LFA1","LFB1","KNA1","KNB1",
      "FD01","FD02","FK01","FK02","MAP1","MAP2","MK01","XK07","MK02","V-03","XK02",
      "V-04","V-05","V-06","V-07","V-08","V-09","V-10","V-11","VAP1","VAP2","VAP3","VD01","VD02","XD01","XD02","XK01","XK06","XK07"]);
    const needsRem = new Set(["FD03","FK03","MAP3","MK03","VAP3","VD03","XD03","XK06","XK03"]);
    const fitGapDB = new Set(["LFA1","LFB1","KNA1","KNB1"]);
    if ([...needsRem].some(t => _eq(t)))   { remType = "Needs Remediation"; remark = "Transaction is working."; }
    else if ([...fitGapDB].some(t => _eq(t))) { remType = "Fit Gap"; remark = "DB Operation INSERT/UPDATE/MODIFY might lead to inconsistency."; }
    else if ([...fitGapT].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Transactions are obsolete or get redirected to transaction BP."; }
  }

  // Note 2689873 — Deprecated Dictionary Objects
  // KK: uses lv_data (cell value in DO loop) to match object names.
  // Here _eq() searches all fields. Use refObj/objType where possible to avoid false matches.
  if (_eqN("2689873")) {
    if (refObjType === "FUNC") { remType = "Needs Remediation"; remark = "DEPRECATED Functional Module in use."; }
    else if (_eq("FUNC")) { remType = "Needs Remediation"; remark = "DEPRECATED Functional Module in use."; }
    else if (refObjType === "DTEL" && (refObj === "CHAR02" || refObj === "CHAR05")) { remType = "False Positive"; remark = "No Remediation Needed."; }
    else if (refObjType === "TRAN" && refObj === "SICH") { remType = "False Positive"; remark = "No Remediation Needed."; }
    else if (refObj === "SICH" || refObj === "CHAR02") { remType = "False Positive"; remark = "No Remediation Needed."; }
    else if (refObjType === "DTEL" || _eq("DTEL") || _eq("RN1DATUM") || _eq("ISH_T_RNG_XFELD")) {
      if (isUpgrade) {
        remType = "False Positive";
        remark  = "Already on S/4HANA — deprecated DTEL object is no longer flagged in upgrade scenario.";
      } else {
        remType = "Needs Remediation"; remark = "Object has been marked for either DELETION or DEPRECATED";
      }
    }
    else if (_eq("DOMA") || _eq("TTYP") || _eq("MSAG") || refObjType === "TTYP" || refObjType === "DOMA" || refObjType === "MSAG") { remType = "Needs Remediation"; remark = "DB Objects either has been deleted or DEPRECATED"; }
    else if (_eq("PARA") || _eq("TRAN") || _eq("TABL") || _eq("NC10") || _eq("V_NMARC")) { remType = "Fit Gap"; remark = "DEPRECATED Objects found. Change Required."; }
    else if (_eq("NLEI")) { remType = "Fit Gap"; remark = "SAP does not plan a conversion from SAP Patient Management (IS-H) to S/4HANA."; }
    else if (_eq("ISH_T_RNG_DATE") || _eq("TY_TLINE")) { remType = "Needs Remediation"; remark = "Table type is available but belongs to Patient Management component."; }
    // N1C003, N2CLIM_LINE and other unlisted objects with DTEL ref type → NR per manual
    else if (refObjType === "TTYP") { remType = "Needs Remediation"; remark = "Object has been marked for either DELETION or DEPRECATED"; }
  }

  // Note 2198647 — Sales Document Status
  if (_eqN("2198647") && !isDml && !isQuery) {
    if (_eq("VBFA")) { remType = "False Positive"; remark = "Remediation would not require with VBFA."; }
    else if (_eq("VBUK") || _eq("VBUP") || _eq("VBAKUK") || _eq("LIKPUK") || _eq("LIPSUP") || _eq("RVVBTYP")) {
      remType = "Needs Remediation"; remark = "Remediation would require for VBUK, VBUP & VBTYP.";
    } else if (_eq("VBAP_VAPMA") || _eq("VAKPA_REF")) {
      remType = "Fit Gap"; remark = "This view is not available in S4HANA.";
    }
  }

  // Note 2203518 — WBS Element Primary Key
  if (_eqN("2203518")) {
    if (_cs("WBRF")) { remType = "False Positive"; remark = "Primary key changed but fields are there."; }
    else { remType = "Needs Remediation"; remark = "Primary key changed in Table WBRFN, WBRP and few fields splited and replaced."; }
  }

  // Note 1804812 — MB Transaction Changes
  if (_eqN("1804812")) {
    if      (_eq("MB1C") || _eq("MCHA") || _eq("MCHB") || _eq("MCH1") || _eq("MB1B") || _eq("MBST")) {
      remType = "Fit Gap"; remark = "MB Transaction can be replaced with MIGO";
    } else if (_eq("MB03")) {
      remType = "Needs Remediation"; remark = "MB Transaction can be used";
    } else if (_eq("MB01") || _eq("MB02") || _eq("MBUS") || _eq("MB04") || _eq("MB05")) {
      remType = "Fit Gap"; remark = "MB transactions like MB01, MB02, MBUS, MB04, MB05, etc have been replaced by the single-screen transaction MIGO";
    } else {
      remType = "Fit Gap"; remark = "MB transactions have been replaced by the single-screen transaction MIGO";
    }
  }

  // Note 2227568 — Delivery Goods Receipt
  if (_eqN("2227568")) {
    if (_cs("MM_LAST_DAY_OF_MONTHS") || _eq("LIKP")) { remType = "False Positive"; remark = "FM/Table Available in S/4HANA"; }
  }

  // Note 2268085 — MRP Live Feature
  // ABAP (flag ZZ): MD04/MD4C/MDBS → FP; MD_STOCK_REQUIREMENTS_LIST_API → FP; else → FG.
  // MDKP is removed in S/4H → FG (functional gap).
  if (_eqN("2268085") && !isLiteral) {
    if (_eq("MD04") || _eq("MD4C") || _eq("MDBS")) {
      remType = "False Positive"; remark = "No change required in FM/Table or Interface";
    } else if (_cs("MD_STOCK_REQUIREMENTS_LIST_API") || _cs("AUFBAUEN_MDPSX_ANZEIGEN")) {
      remType = "Needs Remediation"; remark = "Function module requires adaptation for SAP S/4HANA MRP Live.";
    } else if (refObj === "MDKP" || _eq("MDKP")) {
      remType = "Fit Gap"; remark = "Table MDKP has been removed in SAP S/4HANA MRP Live. Functional assessment required.";
    } else {
      remType = "Fit Gap"; remark = "SAP S/4HANA features MRP Live (MD01N); a new MRP run optimized for SAP HANA.";
    }
  }

  // Notes 2217124/2227014 — Pricing Master Data / Credit Management Changes in FI
  if (_eqN("2217124") || _eqN("2227014")) {
    // S066/S067/VKM3/VKM5 → always FG per manual (not FP even when isS4H)
    const alwaysFitGap = new Set(["S066","S067","VKM3","VKM5","VKM2","VKMI","VAKCR"]);
    const fitGapT2  = new Set(["OB02","S_ER9_11000074","RFDKLI40","RFDKLI41","RFDKLI20","RFDKLIAB"]);
    const conditionalFP = new Set(["V_T024B","KNKK","KNKA","T024B","T691B"]);
    if ([...alwaysFitGap].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Valid for variables declaration but Update, Modify, Insert would require Remediation."; }
    else if ([...conditionalFP].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Valid for variables declaration but Update, Modify, Insert would require Remediation."; }
    else if ([...fitGapT2].some(t => _eq(t))) { remType = "Fit Gap"; remark = "You must eliminate the indicated usages from your customer objects."; }
    // 2227014 — S/4HANA: Credit Management Changes in FI → Fit Gap
    // Also flag orphaned objects if obj type is TADIR/TFDIR
    if (_eqN("2227014")) {
      if (_eq("TADIR") || _eq("TFDIR") || objType === "TADIR" || objType === "TFDIR") {
        remType = "False Positive";
        remark  = "Orphaned object check in TADIR/TFDIR — no action required.";
      } else {
        remType = "Fit Gap";
        remark  = "2227014 - S/4HANA: Credit Management Changes in FI — replaced by new Credit Management in SAP S/4HANA.";
      }
    }
  }

  // Note 2270335 — General Ledger Reporting
  if (_eqN("2270335") && !isDml && !isLiteral) {
    const falsePos = new Set(["FORM","CPUB","GLPA","FORW","FAGLGA48","GP12N","FD03","J1ST","SAPF011","RFKOPO00","RFDKLIAB","RFKQSU20","SAPF100"]);
    const fitGapT  = new Set(["F.35","F.31","FD32","FD33","FCV3","FSS1","FSS2","FS01","FS02","KB21","FO36","FOIO","FS03","KA01","KE21",
      "FOBJ","RAHAFA01","RFDKLI40","RFDKLI41","KA02","KA03","KA06","KP06","KP07","KP65","KP66","CJ41","CJ42","CK74","CJ43","CJR2","CJR1","KKCS"]);
    const tbcsObjs = new Set(["RKKEKLVE"]);
    if ([...falsePos].some(t => _eq(t))) { remType = "False Positive"; remark = "No Change Required. Object available"; }
    else if ([...tbcsObjs].some(t => _eq(t))) { remType = "Fit Gap"; isFitGapDelta = true; remark = "Check usage of this object in system."; }
    else if ([...fitGapT].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Help needed from business or functional team. Transaction Code have been replaced with new functionality."; }
    else { remType = "Fit Gap"; remark = "Help needed from business or functional team. Transaction Code have been replaced with new functionality."; }
  }

  // Note 2469385 — ISM/SAP Media
  if (_eqN("2469385") && !isLiteral) {
    if (_cs("ISM_CUSTOMER_PI") || _cs("RJGBUSISM007") || _eq("A386") || _eq("JPTTITLERELATTR") || refObjType === "TABL") {
      remType = "Fit Gap"; remark = "SAP Media Not Supported in S/4HANA";
    } else if (_eq("J1ST") || _eq("JSE4") || _eq("JSE2") || _eq("JHC2") || _eq("JKK1") || _eq("JSB1") || _eq("JSB2") || _eq("TBDLS")) {
      remType = "False Positive"; remark = "No Change required.";
    } else if (_cs("ISM_DAY_NAME_OF_DATE_GET") || _cs("ISP_SELECT_SINGLE_T151T")) {
      remType = "Fit Gap"; remark = "SAP Media function module not available in S/4HANA.";
    } else {
      remType = "Needs Remediation"; remark = "Changes required in objects like DTEL,DOMA,Range,FM or Message Class.";
    }
  }

  // Note 2223144 — Foreign Trade SD/MM
  // ABAP (flag ZG): ALL rows → Fit Gap unconditionally.
  if (_eqN("2223144")) {
    remType = "Fit Gap";
    remark  = "Foreign Trade in SD/MM. This functionality is no longer supported in SAP S/4HANA.";
  }

  // Note 2340247 — Warehouse Management
  if (_eqN("2340247")) {
    const wqTrans = new Set(["WC01","WQ01","WQ02","WQ03","WQ04","WQ05","WQ11","WQ12","WQ13","WQ14","WQ15","WQ21","WVAL"]);
    if ([...wqTrans].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Need to check manually."; }
    else { remType = "False Positive"; remark = "No change required."; }
  }

  // Note 2217202 — LE Replication Server
  // ABAP (flag ZI): ALL rows → Fit Gap unconditionally.
  if (_eqN("2217202") && !isDml) {
    remType = "Fit Gap"; remark = "Functional help required. This function is no longer supported in SAP S/4HANA.";
  }

  // Notes 2217205/2217206 — Occupational Health & Environmental Compliance
  if (_eqN("2217205") || _eqN("2217206")) {
    remType = "Fit Gap";
    remark  = "Occupational Health and Environmental Compliance (EC) are no longer supported in SAP S/4HANA.";
  }

  // Note 2224144 — /BEV* Objects
  // ABAP (flag ZK): /BEV CS match → FG; then CASE for specific objects → FP
  if (_eqN("2224144")) {
    if (_cs("/BEV")) {
      remType = "Fit Gap"; remark = "Objects starting with /BEV* no more available. Replacement needed.";
    }
    const falsePos = new Set(["CMM_MTM_ANTCP_INV_RES","CMM_MTM_ANTCP_INV_RES_DATA","CMM_S_INV_ANTCP_DATA",
      "CMM_S_TEST_MOCK_DO_PRICING","KOMG","KOMP","KOMP_INCL_D","KOMPAZ","KOMPCZ","KONA",
      "MCDOKOB","MCVBRPUSR","KONP","MSEG","VBAP"]);
    if ([...falsePos].some(t => _eq(t))) { remType = "False Positive"; remark = "No change required with KONP table."; }
  }

  // Note 2354101 — WB Transaction Phases
  if (_eqN("2354101")) {
    remType = "Fit Gap";
    remark  = "The transactions WB01, WB02, WB03 are no longer provide a means of accessing customer/vendor data directly.";
  }

  // Note 2438006 — RFC/BAPI Long Fields
  // Upgrade: already on S/4HANA — long fields already in place → FP
  // Conversion: adaptation required → NR
  if (_eqN("2438006")) {
    if (isUpgrade) {
      remType = "False Positive";
      remark  = "Already on S/4HANA — RFC/BAPI long field parameters are already in use. No code change required.";
    } else {
      remType = "Needs Remediation";
      remark  = "RFC/BAPI in which parameter adjusted for compatbility by introducing new, long fields like MATNR_LONG.";
    }
  }

  // Note 2468869 — Deleted DB Objects
  // ABAP: mv_s4h is never set → always Needs Remediation
  if (_eqN("2468869")) {
    remType = "Needs Remediation";
    remark  = "Multiple Database objects has been deleted in S4HANA. Cross verify with attached file in note";
  }

  // Note 2468834 — Prepayment/External Interfacing Obsolete
  // ABAP: mv_s4h is never set → always Needs Remediation
  if (_eqN("2468834")) {
    remType = "Needs Remediation";
    remark  = "Prepayment, External Interfacing are obsolete and objects related to those functionalities have been deleted.";
  }

  // Note 2209696 — VBBS/VBBE Structures
  if (_eqN("2209696")) {
    remType = "Fit Gap";
    remark  = "If the VBBS is used in customer code, the code has to be adapted. A solution is to create a view on the VBBE.";
  }

  // Note 2225107 — Internet Sales
  // ABAP (flag ZQ): specific objects → FG; no FP fallback per ABAP code
  if (_eqN("2225107")) {
    const fgObjs = new Set(["ISALES_SEL_PARAM","USER_FIELD","TTYP","FUNC","ERP_WEC_USER_ADMIN",
      "IF_EX_ERP_ISA_GEN_DOC_SEL","CL_ERP_WEC_USER_MAIL_CTX_UTIL","CL_SW_DETERMINE_WEC_MILESTONES",
      "CL_ISA_GEN_DOC_SEL_HELP","CL_SW_BUFFER_WEC_STATUS","REGIONS","COUNTRIES","CURRENCYCODES",
      "ISA_PRICING_HEADER_ATTRIBUTE","ISA_PRICING_ITEM_ATTRIBUTE","ISA_SHIPPING_COND","ISA_PARTNER_BUFFER"]);
    if ([...fgObjs].some(t => _eq(t)) ||
        _cs("ISA_") || _cs("ERPS_WEC") || _cs("ERP_WEC") || _cs("ISALES_") || _cs("TDS_") ||
        refObjType === "FUNC") {
      remType = "Fit Gap"; remark = "Custom code does not comply with the scope and data structure of SAP S/4HANA.";
    }
  }

  // Note 2358921 — DPR Portfolio management (CL_DPR_API_CO available in S/4HANA → FP)
  if (_eqN("2358921")) {
    remType = "False Positive"; remark = "Object available in SAP S/4HANA compatibility scope — no code change required.";
  }

  // Notes 2217112 / 2212593 — PLM/WUI switch check (/PLMB/CL_PLMWUI_SWITCH_CHECK available → FP)
  if (_eqN("2217112") || _eqN("2212593")) {
    remType = "False Positive"; remark = "Object available in SAP S/4HANA — no code change required.";
  }

  // Note 2228056 — Sales document table types (TT_VBAK/TT_VBAP/TT_VBUK available in S/4H → FP)
  if (_eqN("2228056")) {
    remType = "False Positive"; remark = "Table type available in SAP S/4HANA — no code change required.";
  }

  // Note 2371631 — Tax Functionality
  if (_eqN("2371631")) {
    remType = "False Positive";
    remark  = "No special actions for additionals functionality is needed when converting from SAP ERP to SAP S/4HANA.";
  }

  // Note 2206980 — ERP Product Availability
  // ABAP (flag ZS): ALL rows → Needs Remediation unconditionally.
  if (_eqN("2206980")) {
    remType = "Needs Remediation"; remark = "SAP ERP 6.0 product do still exist in S/4HANA as DDIC definition as well as database object.";
  }

  // Note 2214585 — Sales/Order Management
  if (_eqN("2214585") && !isDml) {
    const falsePos = new Set([
      "RV_ADDRESS_WINDOW_DISPLAY","SDPARTNERLIST","E1BP_SENDER","TVC2T","BAPI_SENDER","BAPI_RECEIVER","TVC1T",
      "SDCAS_SALES_ACTIVITY_READ_MANY","SD_AUTHORITY_SALES_ACTIVITY","VBKAVB",
      "BAPI_RANGESSPART","BAPI_RANGESKUNNR","BAPI_RANGESAUDAT","BAPI_RANGESVKORG","BAPI_RANGESBSTKD",
      "BAPI_RANGESVKGRP","BAPI_RANGESVKBUR","BAPI_RANGESVTWEG","SDCAS_KUNNR","BAPI_SDVTBER","TVKB",
      "BAPI_BUS1037_VBKAKOM_UPD","BAPI_BUS1037_VBKAKOM_UPDX","BAPI_VBKA_BOID","BAPI_VBKA_BOIDREF",
      "BAPI_VBKA_TLINEKOM","BAPI_VBKA_TLINEKOMX","BAPI_VBKA_VBPA2KOM","LIPS_VLPMA","MAPPING_TO_TLINEKOM",
      "MAPPING_TO_VBKA_TLINEKOM","RV_REPORT_VARIANTE_SEARCH","TLINEKOM","VBKA_KONTAKTE",
      "BAPI_BPCONTACT_CHANGE","BAPI_BPCONTACT_CREATEFROMDATA","BAPI_BPCONTACT_DELETE","BAPI_BPCONTACT_GETDETAIL",
      "BAPI_BPCONTACT_SAVEFROMMEMORY","BAPI_BUS1037_VBKAKOM_CR","BAPI_BUS1037_VBKAKOM_CRX",
    ]);
    const fitGapT3 = new Set(["VBKA","RVDEBCF0","RV76A001","SADLSTWU","CAS01","CAS03","SADLSTADM","SADLSTDIR1",
      "SADLSTRECP","TVKK","TVKKT","TVC3T","TVC1","TVC2","TVC3","TVC4","TVC7","TVCA","TVCB","TVCC","TVCD",
      "TVCE","TVCF","TVCG","TVCH","TVCI","TVCJ","VC/2","VC01","VC03","VC02","VC05","VBRP_VRPMA","VBAP_VAPMA"]);
    if ([...falsePos].some(t => _eq(t))) { remType = "False Positive"; remark = "Objects Available. No Change Required."; }
    else if ([...fitGapT3].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Sales support not available."; }
  }

  // Note 2296016 — Remove References to Deprecated Objects
  if (_eqN("2296016") && !isLiteral) {
    const fp1 = new Set([
      "BSAK","BSAD","BSAS","BSID","BSIK","BSIS","CKIS",
      "SUBST_START_REPORT_IN_BATCH","GLT0","COSS","COSP","RSEG","CL_SALV_GUI_OBSERVER_MANAGER",
      "IF_OIJ05_INV_COCKPIT_CONSTANTS","BAPIITEMILO","LIPS","MARA","MARDH","MARC","MARM","EKBE",
      "PPDS_ON_ERP_CRHD_STR","AUFM","EKBZ","EKET","RKWA","ADRC","FAGLFLEXT","LFC3",
      "CX_RDSV_MAP","CL_RDSV_MAP_UTIL","CL_RDSV_MAP_UTIL_SD","CX_RDSV_QRCODE","IF_SIPT_MAP_QRC_UTIL",
      "IF_RDSV_QRCODE_GET","CL_RDSV_QRCODE","CL_SIPT_SERIES_REP_FI","CL_SIPT_MAP_QRC_UTIL_SD",
      "CL_SIPT_MAP_QRC_UTIL","CL_SIPT_QRCODE_PT","IF_RDSV_MAP_GET","CL_SIPT_SERIES_REP_SD",
      "IF_EDOC_MAP_GR","CL_SIPT_PR_UTIL","CX_SIPT_SERIES","IF_SIPT_SERIES_UTIL_REP","RDSV_QRC_XX",
      "BAPI1076_SPPICLASS","MARD","MARI","MCHB",
      "BAPI_JBD_DTE_CRUSER","BAPI_JBD_DTE_LAND_GP","CL_OO_SOURCE","EKES","J_1BEFDS_C_3_C101","KNKK","LFA1","VER_CURR15",
      "AM_SHOW_POST_DEPR","FIAA_UKV_ACTIV",
      "T156","CMM_MTM_ANTCP_INV_RES","CMM_MTM_ANTCP_INV_RES_DATA","CMM_S_INV_ANTCP_DATA",
      "CMM_S_TEST_MOCK_DO_PRICING","ISUWA_EQUI_DATA","ISUWA_WASTE_DATA","KOMG","KOMP","KOMP_INCL_D",
      "MAT_LUBSCDATARPLCTNREQ_INPUT","MAT_WRHSMGMTVWRPLCTNREQ_INPUT","MBEW","MT61D","QALS","QCERT_TS_CERT",
      "T005","TVAK","TVAKZ","TVKO"
    ]);
    const nr1 = new Set([
      "WEEK_DAY","CH_SPLIT_FILENAME","DD06L","DD16S","DD06T","KNC1","KNC2","KNC3","LFC1",
      "DTRESR","FILE","NOORYES","QUEUE","DATE_TO_DAY","CHAR20K","NEXT_WEEK","DCHAR10","ENQUEUE_ESINDX","VBUP",
      "EKPO","MARA_UEB","KEKO","MCHBH","MARM_UEB","TIME_CONVERSION","MSAG","VEPO","LTAP","HUINV_ITEM",
      "SHLP","/BEV1/TSMMTT","/BEV2/EDMSE","/BEV3/CHBUKRSF4","/KYK/OPS_SD",
      "GET_WEEK_INFO_BASED_ON_DA","BUKRS_D","CONVERT_BDCMSGCOLL_TO_BAPIRET2","/SAPDII/SPP05_CONVERT_DATE",
      "WRETAIL_BAPIEINE_TTY","EAMPRT_TT_STEUS",
      // RPRAPA00: not in ABAP W-flag CASE → blank lv_rem_ty → default NR for S/4H prio1/2
      // LAST_DAY_OF_MONTHS: ABAP says FP but user override → NR
      // MSEG: ABAP says Fit Gap but user override → NR
      "RPRAPA00","LAST_DAY_OF_MONTHS","MSEG"
    ]);
    const fg1 = new Set([
      "IBSP","V_BAM_T001","GBAPPS_APV_USER_FORWARDING","GBAPPS_ATTACHMENTS","T156","NLEI",
      "BAPIEKPOV_CONF","CL_EHHSS_PCO_INC_ACTION","SUBST_CHECK_EXIST","SUBST_OBJECT_PACKAGES_DELIVERY",
      "DTEL","TTYP","DOMA","IF_EHHSS_IAC_C","CATT_GET_TABLE_ENTRIES","IF_EHHSS_PIH_C",
      "/SAPDII/DWB_GET_TABLE_FIELDS","IF_EHHSS_PIL_C","IF_EHHSS_PIL_C","IF_EHHSS_PIR_C",
      "SER02","SER01","ISEG","RASIMU02","SERBL","KOTP002","RAABST02",
      "MOB_BP_TEL_REC","PARTY_ID_STRUC","CHAR_ARRAY","FSX_BD_MSG_H_PARTY","FSX_BUSINESS_SCOPE",
      "CN_MESSAGE_TEXT_BUILD","BUPA_BAPIMTCS_TO_SET","BUPA_SET_TO_BAPIMTCS","MAORD","CATDH","SMET02",
      "RFINDEX","KNVT","SBCOPT","RAPOST2000","V_BAM_KONV","RSTABLESIZE",
      "GBAPPS_APV_PO_IT_X_HNI","GBAPPS_APV_PO_IT_X_IANPA","GBAPPS_PO_ITEM_DETAILS_INCL",
      "SRA001S_PC_HEADER_DETAILS","GBAPPS_PO_HEADER_DETAILS_INCL","SRA001S_PC_ITEM_DETAILS",
      "GBAPPS_CONTACT_DETAILS","GBAPPS_DELIVERY_ADDRESS","GBAPPS_NOTES","GBAPPS_NOTES_INCL",
      "GBAPPS_PO_ACC_DETAILS","GBAPPS_PO_ACC_DETAILS_INCL","GBAPPS_PO_PRODUCT_DETAILS",
      "GBAPPS_PO_SERVICE_LINE","GBAPPS_PO_SERVICE_LINE_INCL","GBAPPS_PR_ACCOUNTING_DET",
      "GBAPPS_PR_ACCOUNTING_DET_INCL","GBAPPS_PR_HEADER_ACC_DETAILS","GBAPPS_PR_HEADER_DETAILS",
      "GBAPPS_PR_ITEM_DETAILS","GBAPPS_PR_ITEM_DETAILS_INCL","GBAPPS_PR_LIMITS",
      "GBAPPS_PR_PRODUCT_DETAILS","GBAPPS_PR_SERVICE_LINE",
      "CL_SD_BILL_VBTYP_EXT",
    ]);
    if      ([...fp1].some(t => _cs(t) || _eq(t))) { remType = "False Positive";    remark = "FM/Class OR Tables CDS is available. No change required."; }
    else if ([...nr1].some(t => _cs(t) || _eq(t))) {
      if (isUpgrade) {
        remType = "False Positive";
        remark  = "Already on S/4HANA — deprecated object is available or no longer flagged in upgrade scenario.";
      } else {
        remType = "Needs Remediation"; remark = "Remove its reference from code and create custom if needed.";
      }
    }
    else if ([...fg1].some(t => _cs(t) || _eq(t))) { remType = "Fit Gap";           remark = "Need functional or business help to resolve. Few fields/Tables are missing."; }
    // Unknown objects — let TADIR/TFDIR block classify (customer FMs not in any list → FG via TFDIR)
    // Do not set default here; TADIR block will handle via !remType gate.
  }

  // Transactions in Literals (LM flag)
  // Note 2223144 sets blanket FG — do not override with isLiteral FP rules
  if (isLiteral && !_eqN("2223144")) {
    const fp = new Set(["FORM","UPDA","J1ST","CPUB","SICH","GP12N","FORW","UPGRADE","WAST","KNC1","LFC1","SDIN"]);
    const fitGapLiteral = new Set([
      "CCIHS_IPIOT","CCIHS_PRAPI","CCIHS_BW_IAL_TD","CCIHS_IALHIOT","ANGB","PA1403","PA1404",
      "CBIH_IA03_TC1_IOTAB_CLOSE","EHS_CALC_YEARS_BETWEEN_DATES","EHS-IHS","XX-SER-REL"
    ]);
    if ([...fp].some(t => _eq(t))) { remType = "False Positive"; remark = "No Change Required."; }
    else if (_eq("TRAN")) { remType = "Fit Gap"; remark = "Check the use of reference objects."; }
    else if ([...fitGapLiteral].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Reference object not Available IN S/4HANA."; }
    else if (!remType) { remType = "Fit Gap"; remark = "Check the use of reference objects."; }
  }

  // ── Simple note → fixed outcome rules (lines 2300–3122 of ABAP) ────────────

  // Note 2226380 — Obsolete SAP Objects
  if (_eqN("2226380") && !isLiteral) {
    remType = "Fit Gap";
    remark  = "All usages of SAP objects in customer objects will no longer work.";
  }

  // Note 2271233 — APIs
  if (_eqN("2271233") && !isLiteral) {
    remType = "Fit Gap";
    remark  = "SAP offers APIs which are available with SAP S/4HANA on-premise 2021.";
  }

  // Note 3067314 — Business Partner Data Model
  if (_eqN("3067314") && !isQuery) {
    remType = "False Positive";
    remark  = "Enhanced support and consumption of the SAP Business Partner data model available.";
  }

  // Note 2268070
  if (_eqN("2268070")) { remType = "False Positive"; remark = "Functionality available in SAP S/4HANA but not considered as future technology"; }
  // Note 2270235
  if (_eqN("2270235")) { remType = "False Positive"; remark = "Process Operator Cockpit is still available in S/4HANA, but not considered as future technology"; }
  // Note 2877717
  if (_eqN("2877717")) { remType = "Needs Remediation"; remark = "Fields from tables J_1IMOCUST and J_1IMOVEND are also migrated into tables KNA1 and LFA1 in S/4 HANA."; }
  // Note 2480067
  if (_eqN("2480067")) { remType = "Fit Gap"; remark = "Name of the new legal report delivered with SAP Document and Reporting."; }
  // Note 2470721
  if (_eqN("2470721") && !isQuery) { remType = "False Positive"; remark = "Custom Code related to table VBFA may be subject to changes.."; }
  // Note 2228241
  if (_eqN("2228241") && !isLiteral) { remType = "Needs Remediation"; remark = "As mentioned in SAP Note 735529, you have been advised to add the code in customer specific programs."; }

  // Notes 2437332 / 2445654 — Data element changes
  if ((_eqN("2437332") || _eqN("2445654")) && !isLiteral) {
    if (refObjType === "DTEL" && objType === "ENHO") { remType = "False Positive"; remark = "Enhancement implementation using data element — no code change needed."; }
    else if (_eq("DTEL")) { remType = "Needs Remediation"; remark = "Data element can be changed."; }
  }

  // Note 2354768 — Material Ledger data model
  if (_eqN("2354768")) { remType = "Fit Gap"; remark = "With S/4HANA the data model for material ledger data has been changed significantly."; }

  // Notes 2227532 / 2227579 — Subcontracting planning segments
  // ABAP (flag YG): ALL → Fit Gap.
  // Exception: MD_STOCK_REQUIREMENTS_LIST_API / AUFBAUEN_MDPSX_ANZEIGEN → NR (Kundan override).
  if (_eqN("2227532") || _eqN("2227579")) {
    if (_eq("MD_STOCK_REQUIREMENTS_LIST_API") || _eq("AUFBAUEN_MDPSX_ANZEIGEN")) {
      remType = "Needs Remediation"; remark = "Function module available in S/4HANA — check if parameter changes apply.";
    } else {
      remType = "Fit Gap"; remark = "Evaluations of the planning segments for subcontracting will not work any longer.";
    }
  }

  // Note 2348023 — Manufacturer Part Number
  if (_cs("2348023")) { remType = "Fit Gap"; remark = "The functionality for the Manufacturer Part Number (MPN) is available in SAP S/4HANA 1610 onwards."; }

  // Note 2229126
  if (_eqN("2229126")) { remType = "Fit Gap"; remark = "Need Business or functional help."; }

  // Note 2371539 — Replenishment
  if (_eqN("2371539")) {
    const fp = new Set(["RPL_S130_CHANGE_IN_UPDATE_TASK","RPL_ONORDER_CHG_IN_UPDATE_TASK","RPL_INV_CHANGE_IN_UPDATE_TASK","WRPL","RPL_MATERIAL_TEXT_GET","MASTERIDOC_CREATE_RPLMAS"]);
    if ([...fp].some(t => _cs(t))) { remType = "False Positive"; remark = "No Changes Required."; }
    else if (_cs("WRPT") || _eq("RWVMIPAV")) { remType = "Fit Gap"; remark = "Changes Required with functional or business help."; }
  }

  // Note 2195701 — Brazil country version
  if (_eqN("2195701")) { remType = "Fit Gap"; remark = "Obsolete transaction codes and programs in Brazil country version."; }

  // Note 2225170 — SD Revenue Recognition
  if (_eqN("2225170")) { remType = "Fit Gap"; remark = "SD-Revenue Recognition functionality is replaced by new application SAP Revenue Accounting and Reporting."; }

  // Note 2378796 — MARC fields STAWN/EXPME
  // DML operations on MARC → FG per manual (field moved); otherwise NR
  if (_eqN("2378796")) {
    if (isDml && (refObj === "MARC" || _eq("MARC"))) { remType = "Fit Gap"; remark = "MARC fields STAWN/EXPME have been moved — use new tables."; }
    else { remType = "Needs Remediation"; remark = "Database table \"MARC\" shall not be used to read fields STAWN and EXPME."; }
  }

  // Note 2267298 — Enjoy T-code
  if (_eqN("2267298")) { remType = "Needs Remediation"; remark = "Enjoy T-code available in SAP S/4HANA on-premise edition."; }

  // Note 2200691 — Table no longer used
  if (_eqN("2200691")) { remType = "Fit Gap"; remark = "This table is no longer used."; }

  // Note 2337368 — Inventory Valuation
  if (_eqN("2337368")) { remType = "Needs Remediation"; remark = "New relevant database objects in the S/4HANA Inventory Valuation data model Available."; }

  // Note 2522971 — Segment field length
  if (_eqN("2522971") && !_cs("USED BY RFC")) {
    if (prio12) { remType = "Needs Remediation"; remark = "Segment Field length extended."; }
    else        { remType = "Can be ignored";     remark = "No change required. Can be ignored."; }
  }

  // Note 2628699 — Extended fields _LONG
  if (_eqN("2628699")) { remType = "Needs Remediation"; remark = "Extended fields are in use. Need to change with _LONG."; }

  // Note 3211383 — MRP / Subcontracting (same FMs as 2268085)
  // Manual: MD_STOCK_REQUIREMENTS_LIST_API + AUFBAUEN_MDPSX_ANZEIGEN → NR; T438A → FP
  if (_eqN("3211383")) {
    if (_cs("MD_STOCK_REQUIREMENTS_LIST_API") || _cs("AUFBAUEN_MDPSX_ANZEIGEN")) {
      remType = "Needs Remediation"; remark = "Function module available in S/4HANA — check if parameter changes apply.";
    } else {
      remType = "False Positive"; remark = "No Change required.";
    }
  }

  // Note 2227059 — PPH_DBVM table
  if (_eqN("2227059")) { remType = "Needs Remediation"; remark = "New database table PPH_DBVM available."; }

  // Note 2228244 — Replace FM
  if (_eqN("2228244")) {
    if (_eq("/SAPDII/SPP05_CONVERT_DATE") || _eq("/SAPHT/SW_LTSPS")) {
      remType = "Needs Remediation"; remark = "Use another FM.";
    }
  }

  // Note 2268131 — SAP S/4HANA compatibility scope
  if (_eqN("2268131")) { remType = "False Positive"; remark = "Currently its part of the SAP S/4HANA compatibility scope."; }

  // Note 2268063 — Custom reports table removal
  if (_eqN("2268063")) { remType = "Fit Gap"; remark = "Custom reports reading one of the following tables will not work."; }

  // Note 2380568 — Capacity requirements
  if (_eqN("2380568")) { remType = "Fit Gap"; remark = "Check in your system whether capacity requirements of rate or rough-cut scheduling despite SOP exist."; }

  // Note 2332591 — Material Ledger ACDOCA
  if (_eqN("2332591")) { remType = "Fit Gap"; remark = "The content of most of the former Material Ledger database tables is now stored in table ACDOCA."; }

  // Note 2469353 — DTEL/DOMA/MSAG changes
  if (_eqN("2469353")) {
    if (_eq("DTEL") || _eq("DOMA") || _eq("MSAG")) {
      remType = "Needs Remediation"; remark = "Chanes required with Data Element, Domain OR Message.";
    }
  }

  // Note 3159740 — JIT Inbound (blanket FG per ABAP UH block)
  if (_eqN("3159740")) {
    remType = "Fit Gap"; remark = "JIT Inbound as part of the automotive industry solution is not the target solution in SAP S/4HANA.";
  }

  // Note 2442292 — PDCE solution
  if (_eqN("2442292")) { remType = "Fit Gap"; remark = "The PDCE solution is not available in SAP S/4HANA."; }

  // Note 2226431 — PDN transaction
  if (_eqN("2226431") && !isLiteral) { remType = "Fit Gap"; remark = "Transaction PDN is no longer supported in S/4 HANA."; }

  // Notes 2267292 / 2368835 — Online store
  if (_eqN("2267292") || _eqN("2368835")) {
    const fitGapT = new Set(["WW10","WW20","WW30","WW31","FUNC","TABL","TRAN"]);
    const fp = new Set(["BAPILAYHLP","BAPILAYLT"]);
    if ([...fitGapT].some(t => _eq(t))) { remType = "Fit Gap"; remark = "In SAP S/4HANA, online store functionality is not available anymore."; }
    else if ([...fp].some(t => _eq(t))) { remType = "False Positive"; remark = "Structure/Table available in S4H."; }
    else if (_eq("DTEL")) { remType = "False Positive"; remark = "No Changes required in code."; }
  }

  // Note 2371157 — Retail Demand Management
  if (_eqN("2371157")) { remType = "False Positive"; remark = "No special actions for Retail Demand Management integration is needed when converting from SAP ERP to SAP S/4HANA."; }

  // Note 2628706 — Currency amount field changes
  if (_eqN("2628706")) { remType = "Needs Remediation"; remark = "Currency amount field Changes happened."; }

  // Note 2382203 — DIMP not available
  // ABAP (flag UM): EKKO/EKET/EKPO/VBAK/VBAP/VBEP OR CHAR65/CHAR02/CHAR20K → FP; else no classification
  if (_eqN("2382203")) {
    const fp = new Set(["EKKO","EKET","EKPO","VBAK","VBAP","VBEP","CHAR65","CHAR02","CHAR20K"]);
    if ([...fp].some(t => _eq(t) || refObj === t)) {
      remType = "False Positive"; remark = "Discrete Industries and Mill Products (DIMP) is not available for SAP S/4HANA.";
    }
  }

  // Note 2270388 — New Asset Accounting
  if (_eqN("2270388")) { remType = "Fit Gap"; remark = "New Asset Accounting in SAP S/4HANA."; }

  // Note 2999249 — Consolidation
  if (_eqN("2999249")) {
    const fp = new Set(["FC_POPUP_ERR_WARN_MESSAGE","CXCR","CX50","CXIN","CX25","FICDELRUDOC00"]);
    if ([...fp].some(t => _eq(t))) { remType = "False Positive"; remark = "Consolidation are part of the SAP S/4HANA compatibility scope."; }
  }

  // Note 2368803 — Retail Ledger
  if (_eqN("2368803")) { remType = "Fit Gap"; remark = "In SAP S/4HANA, Retail Ledger / Profit Center Analytics is not available anymore."; }

  // Note 2879911 — Transceiver
  if (_eqN("2879911")) { remType = "Fit Gap"; remark = "Transceiver functionality is discontinued."; }

  // TXT25 data element
  if (_eq("TXT25")) { remType = "Needs Remediation"; remark = "Replace Data Element."; }

  // Note 2370183 — Retail buying iViews
  if (_eqN("2370183")) { remType = "Fit Gap"; remark = "In SAP S/4HANA, retail buying iViews are not available anymore."; }

  // Note 2370133 — Retail Store
  if (_eqN("2370133")) { remType = "Fit Gap"; remark = "In SAP S/4HANA, SAP Retail Store and SAP In-Store MIM Mobile (Java) are not available anymore."; }

  // Note 2383997 — Eliminate usages
  // DTEL refObjType → NR (code needs adaptation to replace deprecated data element)
  if (_eqN("2383997")) {
    if (refObjType === "DTEL") { remType = "Needs Remediation"; remark = "Replace the indicated data element in your customer objects."; }
    else { remType = "Fit Gap"; remark = "Eliminate the indicated usages from your customer objects."; }
  }

  // Note 2552076 — TBAC_DCS_LOF
  if (_eqN("2552076")) { remType = "False Positive"; remark = "Five new fields has been added in table TBAC_DCS_LOF. No impacts on ABAP coding"; }

  // IDOCTYPE_COLLECT_ATC_INFOS not found
  if (_cs("IDOCTYPE_COLLECT_ATC_INFOS NOT FOUND IN SYSTEM")) { remType = "Fit Gap"; remark = "Not found anyhting in ATC. Unable to analyze issue"; }

  // Note 2226048 — FBS functionality
  if (_eqN("2226048")) { remType = "Fit Gap"; remark = "The FBS functionality is no longer supported in S/4 HANA. Supported via the Budget Control System (BCS) of PSM."; }

  // Note 2383423 — Eliminate usages
  if (_eqN("2383423")) { remType = "Fit Gap"; remark = "Eliminate the indicated usages from your customer objects."; }

  // /SAPHT/SW_LTSPS data element
  if (_eq("/SAPHT/SW_LTSPS")) { remType = "Needs Remediation"; remark = "Replace Data Element with new or custom."; }

  // Note 2226072 — German Local Authorities
  if (_eqN("2226072")) { remType = "Fit Gap"; remark = "The German Local Authorities functionality is no longer supported in S/4 HANA."; }

  // Note 2228218 — Range table changes
  // Manual: ISI_MATERIAL_R → FG, LIKP → FG (manual overrides KK FP for LIKP)
  if (_eqN("2228218")) {
    const nr = new Set(["ISI_EBELN_RA","TTYP","ISI_PLANT_R"]);
    const fg = new Set(["IS-A-SWP","ISI_MATERIAL_R","LIKP","LIPS"]);
    if ([...fg].some(t => _eq(t))) { remType = "Fit Gap"; remark = "Change required with functional help."; }
    else if ([...nr].some(t => _eq(t))) { remType = "Needs Remediation"; remark = "Change the range table or create custom."; }
  }

  // Note 2371559 — Bonus buy
  if (_eqN("2371559")) { remType = "False Positive"; remark = "No special actions for bonus buy is needed when converting from SAP ERP to SAP S/4HANA."; }

  // Note 2862992 — Message class
  if (_eqN("2862992")) { remType = "Needs Remediation"; remark = "Change message class."; }

  // Note 2226134 — ACDOCA
  if (_eqN("2226134")) { remType = "Needs Remediation"; remark = "Table has been replaced with ACDOCA."; }

  // Note 2226966 — Outdated PP interfaces
  // TTYP refObjType → NR per manual; else FG
  if (_eqN("2226966")) {
    if (refObjType === "TTYP") { remType = "Needs Remediation"; remark = "Table type used in PP interface requires remediation."; }
    else { remType = "Fit Gap"; remark = "In SAP S/4 HANA outdated interfaces to external systems in production planning are no longer supported."; }
  }

  // Note 2228098 — Use new object
  if (_eqN("2228098")) { remType = "Needs Remediation"; remark = "Use new object as mentioned in notes."; }

  // Note 2226803 — CET Plant Data Collection (PDC) — all rows → TBCS per manual
  if (_eqN("2226803")) {
    remType = "Fit Gap"; isFitGapDelta = true; remark = "Check if CET Plant Data Collection tables/functions are used and whether migration is needed.";
  }

  // Note 2224778 — cFolders / SAP Content Management
  // TTYP/DTEL/CLAS refObjType → FP; TABL/FUNC/MSAG refObjType → TBCS
  if (_eqN("2224778")) {
    const fpRefTypes = new Set(["TTYP","DTEL","CLAS","INTF"]);
    const tbcsRefTypes = new Set(["TABL","FUNC","MSAG","FUGR","PROG"]);
    if (fpRefTypes.has(refObjType)) { remType = "False Positive"; remark = "No Change required."; }
    else if (tbcsRefTypes.has(refObjType)) { remType = "Fit Gap"; isFitGapDelta = true; remark = "Check usage of cFolders/SAP Content Management object in system."; }
  }

  // Note 2371616 — Workforce Scheduling (WFCS) — TABL/FUNC → TBCS per manual
  if (_eqN("2371616")) {
    remType = "Fit Gap"; isFitGapDelta = true; remark = "Check usage of Workforce Scheduling object in system.";
  }

  // Note 2214585 — Customer Master Extensions
  // Specific TABL refs (KNA1EXIT, BAPI_RANGESBSTNK) → TBCS; BAPI_RANGES* → FP
  if (_eqN("2214585")) {
    const tbcsRefs = new Set(["KNA1EXIT","BAPI_RANGESBSTNK"]);
    if (tbcsRefs.has(refObj)) { remType = "Fit Gap"; isFitGapDelta = true; remark = "Check usage of customer master extension object in system."; }
    else if (refObj.startsWith("BAPI_RANGES") || _cs("BAPI_RANGES")) { remType = "False Positive"; remark = "No Change required."; }
  }

  // Note 2383466 — DTEL/MSAG/TRAN changes
  if (_eqN("2383466")) {
    if (_eq("DTEL") || _eq("MSAG")) { remType = "Needs Remediation"; remark = "Change with new or custom objects."; }
    else if (_eq("TRAN")) { remType = "Fit Gap"; remark = "Transactions not Available with S4HANA."; }
  }

  // Note 3018718 — MSAG / RLB_DB_VST
  if (_eqN("3018718")) {
    if (_eq("MSAG")) { remType = "Needs Remediation"; remark = "Replace message class or create custom."; }
    else if (_eq("RLB_DB_VST")) { remType = "Fit Gap"; remark = "PACKAGE SRLB IS MISSING."; }
  }

  // Dynamic DB-Access — only override if note hasn't set a specific classification
  if (_cs("DYNAMIC DB-ACCESS") && !remType) { remType = "Fit Gap"; remark = "Depricated tables should not be in use."; }

  // Note 2198035 — Primary key changes
  if (_eqN("2198035")) { remType = "Needs Remediation"; remark = "Primary key of few table has been changed."; }

  // Note 1912445 — Cluster/Pool table / order-by / result-set check
  // prio3 rows are Optional only for specific "low severity" checkMsg patterns:
  //   LOOP AT ITAB. EXIT/RETURN/LEAVE ... FOR RESULT OF STATEMENT (not SELECT STATEMENT)
  //   READ TABLE ... INDEX 1 FOR RESULT OF STATEMENT
  //   WRITE IN LOOP FOR RESULT OF STATEMENT
  //   ALV CALL AT ... LINE ...
  //   OPEN CURSOR ... FOR (FORMER) CLUSTER/POOL TABLE ... WITHOUT ORDER BY
  // All other 1912445 rows → Mandatory
  // NOTE: Must check checkMsg directly (not _cs) because the Check Title contains "OPEN CURSOR"
  //       and would falsely match the OPEN CURSOR condition for all prio3 rows.
  // Some rows have empty Note — detect by the unique Check Title text instead.
  const _is1912445 = _eqN("1912445") || title.includes("SEARCH PROBLEMATIC STATEMENTS FOR RESULT OF SELECT");
  if (_is1912445) {
    const _hasResultOf = checkMsg.includes("FOR RESULT OF STATEMENT") && !checkMsg.includes("FOR RESULT OF SELECT STATEMENT");
    const _isOptional1912 = prio3 && (
      _hasResultOf ||
      checkMsg.includes("ALV CALL") ||
      (checkMsg.includes("OPEN CURSOR") && (checkMsg.includes("CLUSTER TABLE") || checkMsg.includes("POOL TABLE")))
    );
    if (_isOptional1912) {
      remType = "Optional"; remark = "Non-critical result-set usage — review recommended.";
    } else {
      remType = "Mandatory"; remark = "Use of cluster/pool table or non-deterministic order by requires remediation.";
    }
  }

  // Note 2371602 — Order optimizing
  // DTEL refObjType → NR per manual; else FP
  if (_eqN("2371602")) {
    if (refObjType === "DTEL") { remType = "Needs Remediation"; remark = "Data element used in order optimizing requires replacement."; }
    else { remType = "False Positive"; remark = "No special actions for order optimizing via load build, and investment buy is needed."; }
  }

  // Note 2343388 — Home building
  // refObjType=DTEL → NR; refObjType=TABL → FP; refObjType=TRAN → FG; else existing logic
  if (_eqN("2343388")) {
    if (refObjType === "DTEL") { remType = "Needs Remediation"; remark = "Replace deprecated Home Building data element."; }
    else if (refObjType === "TRAN") { remType = "Fit Gap"; remark = "Home Building solution will not be available in SAP S/4HANA."; }
    else if (refObjType === "TABL" || _eq("HBSCREASON") || _eq("CHBSNCM") || _eq("HBSSPH") || _eq("VBAK")) { remType = "False Positive"; remark = "No Change required."; }
    else if (_eq("SI_DI_HOME_BUILDG_SL")) { remType = "Fit Gap"; remark = "Home Building solution will not be available in SAP S/4HANA."; }
    else if (_eq("IS-HT-SW-LIC") || _eq("IS-ADEC-HBS") || _eq("IS-M")) { remType = "False Positive"; remark = "No Change required."; }
  }

  // Note 2370151 — Prepack allocation planning
  if (_eqN("2370151")) { remType = "Fit Gap"; remark = "In SAP S/4HANA, prepack allocation planning is not available anymore."; }

  // Note 2371149 — IDoc WP_PLU
  if (_eqN("2371149")) { remType = "False Positive"; remark = "In SAP S/4HANA, No special action needed FOR IDoc WP_PLU."; }

  // Note 2358356 — Select Statement ORIGITEM/WORKITEM
  if (_eqN("2358356")) { remType = "Needs Remediation"; remark = "Change Select Statement logic as per note for ORIGITEM or WORKITEM."; }

  // Note 2223800 — Deprecated SAP objects
  if (_eqN("2223800")) { remType = "Fit Gap"; remark = "SAP objects used by the customer objects are deprecated and shall not be used any more."; }

  // Note 2368680 — Obsolete Functionality
  if (_eqN("2368680")) { remType = "Fit Gap"; remark = "Obsolete Functionality."; }

  // Note 2368741 — Merchandise/assortment planning
  if (_eqN("2368741")) { remType = "Fit Gap"; remark = "In SAP S/4HANA, merchandise and assortment planning/category management is not available anymore."; }

  // Note 2842844 — Obsolete SAP_BASIS tables
  if (_eqN("2842844")) { remType = "Fit Gap"; remark = "Obsolete SAP_BASIS tables have been removed."; }

  // Note 2226129 — CDS view V_FMBDT
  if (_eqN("2226129")) { remType = "Needs Remediation"; remark = "A compatibility view is attached as proxy object to the totals table (CDS view V_FMBDT)."; }

  // Note 3046934 — RLM compatibility scope
  if (_eqN("3046934")) { remType = "False Positive"; remark = "(RLM) is part of the SAP S/4HANA compatibility scope."; }

  // Note 3036174 — Remove references before migration
  if (_eqN("3036174")) { remType = "Fit Gap"; remark = "Remove any references to these objects in your custom code before the migration to SAP S/4HANA 1610 or higher."; }

  // Note 2267246 — MRP fields in Material Master
  if (_eqN("2267246")) { remType = "Needs Remediation"; remark = "Simplification: MRP fields in Material Master in MM01/02/03."; }

  // Note 2460847 — CPET structures
  if (_eqN("2460847")) {
    if (_eq("CPET_FE_CALLER_DATA_I") || _eq("CPET_FE_CALLER_DATA")) {
      remType = "False Positive"; remark = "Structure Available.";
    }
  }

  // Note 3224217 — Change CDS
  if (_eqN("3224217")) { remType = "Needs Remediation"; remark = "Change CDS according to Note."; }

  // Note 3404390 — P_BUSINESSPLACE
  // ABAP (flag RD): ALL rows → Needs Remediation unconditionally.
  if (_eqN("3404390")) {
    remType = "Needs Remediation"; remark = "Use CDS view P_BUSINESSPLACE instead of J_1BBRANCH in your custom program";
  }

  // Note 3224317 — Enterprise Compensation Management
  if (_eqN("3224317")) { remType = "Fit Gap"; remark = "Enterprise Compensation Management (PA-EC) is not available for RISE with SAP S/4HANA Cloud"; }

  // Note 2852082 — DNO_OW / DNOS_RANGES
  if (_eqN("2852082")) {
    if (_eq("DNOS_RANGES")) { remType = "Fit Gap"; remark = "Can be ignored"; }
    else if (_eq("DNO_OW_EXTERN_SEND_NOTIF_2_BC")) { remType = "Fit Gap"; remark = "DNO_OW_EXTERN_SEND_NOTIF_2_BC Not available S/4HANA"; }
  }

  // Note 3224353 — RISE with SAP S/4HANA Cloud
  if (_eqN("3224353")) { remType = "Fit Gap"; remark = "Additional information for customers using RISE with SAP S/4HANA Cloud, private edition"; }

  // Note 2349182 — Different behavior
  if (_eqN("2349182")) { remType = "Needs Remediation"; remark = "Use Provided note. Different behavior compared to that in the SAP Business Suite."; }

  // Note 2559018 — New CDS in note attachment
  if (_eqN("2559018")) { remType = "Needs Remediation"; remark = "New CDS is provided in Note attachment."; }

  // Note 2468070 — New table /SCMTMS/C_SFIR_T
  if (_eqN("2468070")) { remType = "Needs Remediation"; remark = "New table \"/SCMTMS/C_SFIR_T\" is provided in Note attachment."; }

  // Note 2217299 — /SCMTMS/C_SFIR_T (dual: Needs Rem + False Pos — ABAP sets both in sequence; last wins = False Positive)
  if (_eqN("2217299")) { remType = "False Positive"; remark = "Table available. No changes needed here."; }

  // Note 2326769 — Process Batch
  if (_eqN("2326769")) { remType = "Fit Gap"; remark = "Process Batch functionality is not available in SAP S/4HANA."; }

  // Note 3261647 — Deleted ABAP artifacts
  if (_eqN("3261647")) { remType = "Fit Gap"; remark = "Various ABAP artifacts have been deleted."; }

  // Note 3224320 — Training & Events
  if (_eqN("3224320")) {
    const fp = new Set(["BAPI_BUS_EVENT_SCHEDULE","BAPI_BOOK_ATTENDANCE","BAPI_BUS_EVENTGROUP_LIST","BAPI_BUS_EVENTTYPE_INFO",
      "BAPI_DELETE_ATTENDANCE","RH_GET_EVENT_DATA","RH_GET_PARTICIPANTS","RH_MESSAGE_OUTPUT",
      "RH_PARTICIPATION_DELETE","RH_PARTICIPATION_INSERT","RH_REFDOC_INSERT","RH_REFDOC_NUMBER_GET_NEXT"]);
    if ([...fp].some(t => _eq(t))) { remType = "False Positive"; remark = "No Change Needed."; }
    else if (_eq("PV04")) { remType = "Fit Gap"; remark = "Need Functional Support."; }
  }

  // Note 2227963 — Use new transactions/programs
  if (_eqN("2227963")) { remType = "Fit Gap"; remark = "Use new transactions and programs.."; }

  // Note 3211826 — flag RS: Non-strategic functions (KK: False Positive)
  if (_eqN("3211826")) { remType = "False Positive"; remark = "FM Available"; }

  // Note 2267918 — flag RT: Transaction not available in S/4HANA 1511 (KK: Fit Gap)
  if (_eqN("2267918")) { remType = "Fit Gap"; remark = "Transaction not available in SAP S/4HANA on-premise edition 1511"; }

  // Note 2370131 — flag ZF: RIS functionality not available
  // DTEL refObjType → FP (MCW_* objects available); else FG per KK
  if (_eqN("2370131")) {
    if (refObjType === "DTEL") { remType = "False Positive"; remark = "Data element available in S/4HANA — no action needed."; }
    else { remType = "Fit Gap"; remark = "RIS functionality not available anymore. Object might referring to obsolete T-Code or Tables."; }
  }

  // ── Missing simple flags (hardcoded outcomes in CASE block) ─────────────────

  // Note 2406571 — flag OO
  if (_eqN("2406571")) { remType = "False Positive"; remark = "If already on HANA. Relevant for Ariba Interface if planning SAP ERP 6.x to SAP S/4HANA"; }

  // Note 2270199 — flag PP
  if (_eqN("2270199")) { remType = "False Positive"; remark = "The functionality is available in SAP S/4HANA but not considered as future technology."; }

  // Note 2628704 — flag QQ
  // Conversion: relevant when converting from ECC60 → NR.
  // Upgrade: already on S/4HANA 1809+ → finding is no longer applicable → FP.
  if (_eqN("2628704")) {
    if (isUpgrade) {
      remType = "False Positive";
      remark  = "Already on SAP S/4HANA On-Premise 1809 or higher — this finding is no longer applicable.";
    } else {
      remType = "Needs Remediation";
      remark  = "This is relevant, if you are converting from SAP ERP ECC60 or upgrading to SAP S/4HANA On-Premise 1809 or higher.";
    }
  }

  // Note 2438110 — flag RR: FUNC refObjType → NR per manual; else FP
  if (_eqN("2438110")) {
    if (refObjType === "FUNC" || _eq("FUNC")) { remType = "Needs Remediation"; remark = "Syntactically incompatible change of existing functionality — code adaptation required."; }
    else { remType = "False Positive"; remark = "For local calls, no changes to the custom code are necessary"; }
  }

  // Note 2694441 — ISU Payment Functions (IS-U industry-specific FMs)
  // FUNC refObjType: ISU_* FMs require manual verification → NR with check-in-system guidance.
  // DTEL/other refObjType: object exists in S/4HANA → FP.
  if (_eqN("2694441")) {
    if (refObjType === "FUNC" || (!refObjType && (objType === "FUNC" || objType === "FUGR"))) {
      remType = "Needs Remediation";
      remark  = "ISU/IS-U function module — verify availability and parameter compatibility in target S/4HANA system before go-live.";
    } else {
      remType = "False Positive";
      remark  = "IS-U object available in S/4HANA — no code change required.";
    }
  }

  // Note 2359662 — flag TT
  if (_eqN("2359662")) { remType = "Fit Gap"; remark = "Global Settings have been changed in SAP Portfolio and Project Management for SAP S/4HANA."; }

  // Note 2198031 — flag VV
  if (_eqN("2198031")) { remType = "False Positive"; remark = "DB Tabels have new or changed KEY Fields. No Impacts on code"; }

  // Note 1976487 — flag XX
  if (_eqN("1976487") && !isDml) { remType = "Needs Remediation"; remark = "Tables have been replaced by views of the same name but no fields are missing."; }

  // Note 2669857 — flag ZB
  // Conversion: BAPI Object Number parameter may need adaptation → NR.
  // Upgrade: already on S/4HANA — BAPI is available → FP.
  if (_eqN("2669857")) {
    if (isUpgrade) {
      remType = "False Positive";
      remark  = "Already on S/4HANA — BAPI is available. Object Number parameter check not required.";
    } else {
      remType = "Needs Remediation";
      remark  = "Remediation would require if BAPI parameters Object Number being filled.";
    }
  }

  // Note 2368913 — flag MM (mapped to Fit Gap + isFitGapDelta; ABAP used 'Need to check' but we normalize to standard type)
  // WRF_* objects exist in S/4H → Kundan overrides to FP; keep FP for those.
  if (_eqN("2368913")) {
    const tadirKey = refObjType && refObj ? refObjType + "|" + refObj : null;
    const foundInTadir = tadirKey && tadirSet && tadirSet.has(tadirKey);
    if ((refObj && refObj.startsWith("WRF_")) || (foundInTadir && !(refObj === "EKET" || refObj === "EKKO"))) {
      remType = "False Positive"; remark = "Object available in S/4HANA.";
    } else {
      remType = "Fit Gap"; isFitGapDelta = true; remark = "In SAP S/4HANA, seasonal procurement is not available anymore (TBCS).";
    }
  }

  // Note 2368747 — flag N
  if (_eqN("2368747")) { remType = "False Positive"; remark = "In SAP S/4HANA, fresh item procurement is not available anymore."; }

  // Notes 2250183 / 2270080 — flag ZD
  if (_eqN("2250183") || _eqN("2270080")) { remType = "False Positive"; remark = "Mobile Asset Management solution is not supported in S/4HANA."; }

  // Search in ITAB for result
  if (_cs("SEARCH ... IN ITAB FOR RESULT")) { remType = "Mandatory"; remark = "Sort Statement needed."; }

  // ── S/4H: FLE — FIELD LENGTH EXTENSIONS classification ─────────────────────
  // Mirrors ZCL_ZATC_ASSESSMENT_HANDLER lines 3755-3767:
  //   prio 3           → Can be ignored
  //   prio 1/2 + RFC   → False Positive (informational)
  //   prio 1/2 + upgr  → False Positive (already on S/4HANA — FLE not applicable)
  //   prio 1/2 + conv  → Needs Remediation
  if (title.includes("FIELD LENGTH EXTENSIONS") ||
      _cs("S/4HANA: FIELD LENGTH EXTENSIONS") || _cs("S/4HANA: FLE")) {
    if (prio3) {
      remType = "Can be ignored";
      remark  = "FLE prio 3 — informational only.";
    } else if (prio12) {
      if (_cs("RFC-FUNCTION PARAMETER")) {
        remType = "False Positive";
        remark  = "FLE: RFC-Function parameter type reference is informational — no code change required.";
      } else if (isUpgrade) {
        remType = "False Positive";
        remark  = "Already on S/4HANA — FLE finding not applicable in upgrade scenario.";
      } else {
        remType = "Needs Remediation";
        remark  = "S/4HANA field length extension — review field usage and adapt data types if needed.";
      }
    }
  }

  // ── S/4H: Readiness Check for SAP Queries → Fit Gap ────────────────────────
  if (title.includes("READINESS CHECK FOR SAP QUERIES") ||
      _cs("S/4HANA: READINESS CHECK FOR SAP QUERIES")) {
    remType = "Fit Gap";
    remark  = "SAP Queries need to be reviewed by functional team — SAP Query tool behaviour changed in S/4HANA.";
  }

  // ── S/4H: SEARCH FOR SIMPLIFIED TRANSACTIONS IN LITERALS ───────────────────
  // Only applies to S/4H rows (isS4H=true) — avoids conflict with the generic
  // isLiteral block above which handles the HCA "TRANSACTIONS IN LITERALS" check.
  // ABAP class (lines 3805-3818):
  //   False Positive: UPDA/SICH/FORM/MCHA/MCHB/MCH1/MCHP/FOUR/MR01
  //   Needs Remediation: XD03/XK03/FD03/FK03/VD03/VK03/MB03/ME23/ME53/ME26
  //   Fit Gap: everything else
  if (isS4H && (title.includes("SEARCH FOR SIMPLIFIED TRANSACTIONS IN LITERALS") ||
      _cs("S/4HANA: SEARCH FOR SIMPLIFIED TRANSACTIONS IN LITERALS"))) {
    const fpLit = new Set(["UPDA","SICH","FORM","MCHA","MCHB","MCH1","MCHP","FOUR","MR01"]);
    const nrLit = new Set(["XD03","XK03","FD03","FK03","VD03","VK03","MB03","ME23","ME53","ME26"]);
    if ([...fpLit].some(t => _eq(t))) {
      remType = "False Positive";
      remark  = "No change required for this transaction reference.";
    } else if ([...nrLit].some(t => _eq(t))) {
      remType = "Needs Remediation";
      remark  = "Transaction reference must be updated — transaction no longer available or replaced in S/4HANA.";
    } else {
      remType = "Fit Gap";
      remark  = "Simplified transaction reference — functional assessment required.";
    }
  }

  // ── S/4H: Search for Usages of Simplified Objects ──────────────────────────
  // Mirrors the ABAP TADIR/TFDIR lookup in ZCL_ZATC_ASSESSMENT_HANDLER:
  //   - Extract referenced object type + name from the row
  //   - If object type is FUNC/FUGR: look up in tfdirSet (TFDIR)
  //   - Otherwise: look up in tadirSet (TADIR) as "OBJECT|OBJ_NAME"
  //   - FOUND → object exists in S/4HANA system
  //   - NOT FOUND → orphaned object
  if (title.includes("SEARCH FOR USAGES OF SIMPLIFIED OBJECTS") ||
      _cs("S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS")) {

    // 2227014 Credit Management always → Fit Gap regardless of TADIR
    // Use note-exact match only — _cs("CREDIT MANAGEMENT") was too broad and
    // caught unrelated rows where "credit management" appeared in the message text.
    if (_eqN("2227014") || _cs("2227014")) {
      remType = "Fit Gap";
      remark  = "2227014 - S/4HANA: Credit Management Changes in FI — replaced by new Credit Management in SAP S/4HANA.";
    // 2224778 — handled by note-specific block earlier (sets TBCS or FP); skip here
    } else if (_eqN("2224778") || _cs("2224778")) {
      // Classification already set by note 2224778 block above — do not override.
      // (TBCS for TABL/FUNC/MSAG, FP for TTYP/DTEL/CLAS)
    // Only do TADIR/TFDIR lookup when no note-based rule has already set remType.
    // KK loop 2 (TADIR) only meaningfully overrides when the check message pattern
    // matches specific lv_check values. For rows already classified by a note rule,
    // the note outcome takes priority (matches manual sheet as ground truth).
    } else if (!remType && ((tadirSet && tadirSet.size > 0) || (tfdirSet && tfdirSet.size > 0))) {
      // TADIR/TFDIR data available — do orphan check.
      // KK loop 2 (lines 3264-3453) uses REF_15 (ref object type) to determine lookup path:
      //   WHEN 'FUNC' → TFDIR lookup on ref object name
      //   WHEN 'DTEL'/'DOMA'/'TTYP'/'INTF'/'CLAS' → TADIR lookup on ref object type + name
      // Mirror this: use refObjType when available, fall back to objType for backward compatibility.
      const lookupType  = refObjType || objType;
      const isFuncType  = lookupType === "FUNC";
      // Bug 5 fix: include TABL so table references are looked up in TADIR.
      // Not-found TABL → Fit Gap (table removed/simplified in S/4H, needs functional help).
      const isTadirType = ["DTEL","DOMA","TTYP","INTF","CLAS","TABL"].includes(lookupType);

      if (isFuncType) {
        // Function module: look up funcname in TFDIR
        // tfdirSet now contains the actual FMs found in real TFDIR for the referenced objects
        // in this ATC batch. Found = exists in S/4H, not found = removed/simplified.
        const fnKey = refObj || objName;
        if (tfdirSet && tfdirSet.has(fnKey)) {
          // FOUND in TFDIR — FM exists in S/4H
          if (isUpgrade) {
            remType = "False Positive";
            remark  = "Already on S/4HANA. Object Available as Referenced FM in program.";
          } else {
            // ABAP loop 2 TFDIR-found branch (lines 3376-3420):
            //   lv_check='X' (hard conflict) → NR
            //   lv_check='Y' (OLD/GENERIC/DYNAMIC etc) → Can be ignored
            //   lv_check=blank (everything else) → no assignment → default NR for prio1/2
            // NON-STRATEGIC-FUNCTION: set FP by loop 1, loop 2 does not overwrite.
            if (_isHardConflict(checkMsg)) {
              remType = "Needs Remediation";
              remark  = "Extended fields has been provided in FM/BAPI — check if parameter changes apply.";
            } else if (checkMsg === "NON-STRATEGIC-FUNCTION") {
              remType = "False Positive";
              remark  = "FM Available.";
            } else if (_cs("RFC")) {
              remType = "False Positive";
              remark  = "FM available in S/4HANA — RFC parameter reference, no code change required.";
            } else if (_isLCheckY(checkMsg)) {
              remType = "Can be ignored";
              remark  = "FM available in S/4HANA. No action required.";
            } else {
              // Blank lv_check, prio1/2 → ABAP default: Needs Remediation
              remType = "Needs Remediation";
              remark  = "FM available in S/4HANA but functionality may not work as expected — check and remediate.";
            }
          }
        } else if (tfdirSet && tfdirSet.size > 0) {
          // NOT FOUND in tfdirSet — FM not present in S/4H TFDIR (removed/simplified)
          // Applies to both SAP-standard and customer FMs equally.
          remType = "Fit Gap";
          remark  = "FM Not Available in S/4H. Need functional help.";
        }
      } else if (isTadirType) {
        // DDIC ref type: look up in TADIR as "REFTYPE|REFNAME" (mirrors KK REF_15-based lookup)
        const tadirKey = `${lookupType}|${refObj || objName}`;
        if (tadirSet && tadirSet.has(tadirKey)) {
          // FOUND in TADIR — object exists in S/4H
          if (isUpgrade) {
            const isRfcRef = _cs("RFC");
            remType = isRfcRef ? "False Positive" : (prio3 ? "Can be ignored" : "Needs Remediation");
            remark  = isRfcRef ? "RFC-Function parameter — object available in S/4HANA." : "Object available in S/4HANA — check if code adaptation required.";
          } else {
            // Bug 4 fix: only NR on hard-conflict check messages (lv_check='X' in ABAP).
            // Non-hard-conflict (blank/Y/Z) → Can be ignored; RFC → False Positive.
            if (_isHardConflict(checkMsg)) {
              if (lookupType === "INTF" || lookupType === "CLAS") {
                remType = "Needs Remediation";
                remark  = "Interface/Class Objects Available. Check uses of reference objects.";
              } else {
                remType = "Needs Remediation";
                remark  = "Check uses of DDIC objects as it is available in S/4H.";
              }
            } else if (_cs("RFC")) {
              remType = "False Positive";
              remark  = "Object available in S/4HANA — RFC parameter reference, no code change required.";
            } else {
              if (lookupType === "INTF" || lookupType === "CLAS") {
                remType = "Needs Remediation";
                remark  = "Class/Interface available in S/4HANA but check if usage requires adaptation.";
              } else {
                remType = "Needs Remediation";
                remark  = "Data Element, Domain or Table Type available in S/4HANA but check if usage requires adaptation.";
              }
            }
          }
        } else if (tadirSet && tadirSet.size > 0) {
          // NOT FOUND in TADIR — orphaned object
          if (lookupType === "INTF" || lookupType === "CLAS") {
            remType = "Fit Gap";
            remark  = "Functional or Business help needed.";
          } else if (lookupType === "TABL") {
            // Bug 5 fix: table not found in TADIR = removed/simplified in S/4H → Fit Gap.
            // Matches ZATCASSESSMENT_KK first-LOOP flag logic for removed inventory/DB tables.
            remType = "Fit Gap";
            remark  = "Table no longer available in S/4HANA. Functional/business assessment required.";
          } else {
            remType = "Needs Remediation";
            remark  = "Create Custom Data Element, Domain or Table Types.";
          }
        }
        // If tadirSet is empty (no SAP connection), leave remType unset → falls to default
      } else {
        // Neither FUNC nor explicit DDIC type — fall back to original objType-based logic
        const objIsFuncType = objType === "FUNC" || objType === "FUGR";
        if (objIsFuncType) {
          const fnKey = refObj || objName;
          if (tfdirSet && tfdirSet.has(fnKey)) {
            if (isUpgrade) {
              remType = "False Positive";
              remark  = "Already on S/4HANA. Object Available as Referenced FM in program.";
            } else {
              remType = "Needs Remediation";
              remark  = "Extended fields has been provided in FM/BAPI — check if parameter changes apply.";
            }
          } else if (tfdirSet && tfdirSet.size > 0) {
            const isCustomerFm2 = /^[YZ]/i.test(fnKey) || fnKey.startsWith("/");
            if (!isCustomerFm2) {
              if (isUpgrade) {
                remType = "False Positive";
                remark  = "Already on S/4HANA. Object Available as Referenced FM in program.";
              } else {
                remType = "Needs Remediation";
                remark  = "Extended fields has been provided in FM/BAPI — check if parameter changes apply.";
              }
            } else {
              const knownOrphans = ["DR_GET_COUNTRY_NAME","CH_SPLIT_FILENAME","DATE_TO_DAY","FORMAT_DATE_4_OUTPUT",
                "ISP_CONVERT_FIRSTCHARS_TOUPPER","ISP_GET_MONTH_NAME","WSAF_GET_LAST_DAY_OF_MONTH",
                "DATEFORMAT","NEXT_WEEK","ISP_GET_WEEKDAY_NAME","ISH_GET_DAY_OF_WEEK","/SAPDII/SPP05_CONVERT_DATE"];
              if (knownOrphans.some(fn => fnKey.includes(fn))) {
                remType = "Needs Remediation";
                remark  = "FM Not Available in S/4H, create custom.";
              } else {
                remType = "Fit Gap";
                remark  = "FM Not Available in S/4H. Need functional help.";
              }
            }
          }
        } else {
          // DDIC/non-func object: look up in TADIR as "OBJTYPE|OBJ_NAME"
          const tadirKey = `${objType}|${refObj || objName}`;
          if (tadirSet && tadirSet.has(tadirKey)) {
            if (isUpgrade) {
              const isRfcRef = _cs("RFC");
              remType = isRfcRef ? "False Positive" : (prio3 ? "Can be ignored" : "Needs Remediation");
              remark  = isRfcRef ? "RFC-Function parameter — object available in S/4HANA." : "Object available in S/4HANA — check if code adaptation required.";
            } else {
              // Bug 4 fix: only NR on hard-conflict check messages; others → Can be ignored
              if (_isHardConflict(checkMsg)) {
                if (objType === "INTF" || objType === "CLAS") {
                  remType = "Needs Remediation";
                  remark  = "Interface/Class Objects Available. Check uses of reference objects.";
                } else {
                  remType = "Needs Remediation";
                  remark  = "Check uses of DDIC objects as it is available in S/4H.";
                }
              } else if (_cs("RFC")) {
                remType = "False Positive";
                remark  = "Object available in S/4HANA — RFC parameter reference, no code change required.";
              } else {
                if (objType === "INTF" || objType === "CLAS") {
                  remType = "Needs Remediation";
                  remark  = "Class/Interface available in S/4HANA but check if usage requires adaptation.";
                } else {
                  remType = "Needs Remediation";
                  remark  = "Data Element, Domain or Table Type available in S/4HANA but check if usage requires adaptation.";
                }
              }
            }
          } else if (tadirSet && tadirSet.size > 0) {
            if (objType === "INTF" || objType === "CLAS") {
              remType = "Fit Gap";
              remark  = "Functional or Business help needed.";
            } else if (objType === "TABL") {
              // Bug 5 fix: table not found in TADIR → removed in S/4H → Fit Gap
              remType = "Fit Gap";
              remark  = "Table no longer available in S/4HANA. Functional/business assessment required.";
            } else {
              remType = "Needs Remediation";
              remark  = "Create Custom Data Element, Domain or Table Types.";
            }
          }
        }
      }
    }
  }

  // ── S/4H: SEARCH FOR DATABASE OPERATIONS — DML (INSERT/UPDATE/MODIFY/DELETE) → Fit Gap ────────
  // Mirrors ABAP WHEN 'Y' → 'Fit Gap' hardcoded in add_analysis: DML flag always wins,
  // even if a note rule set remType earlier (ABAP ignores lv_rem_ty for flag 'Y').
  // SELECT/query rows are handled earlier and are NOT overridden here (isDml is false for them).
  if (isS4H && isDml &&
      (title.includes("SEARCH FOR DATABASE OPERATIONS") ||
       title.includes("S/4HANA: SEARCH FOR DATABASE OPERATIONS"))) {
    remType = "Fit Gap";
    remark  = "Need Functional help before making changes to DML statements like Update, Modify, Insert or Delete.";
  }

  // ── HCA P3: Check Title contains "problematic" + Check Message contains "select" → Mandatory
  if (!isS4H && prio3) {
    if (!remType && title.includes("PROBLEMATIC")) {
      if (checkMsg.includes("SELECT")) {
        remType = "Mandatory";
        remark  = "P3 finding with problematic title and SELECT in message — mandatory review required.";
      }
    }
    // Check Title contains "ADBC" (e.g. "USE OF ADBC INTERFACE") → Mandatory
    if (!remType && title.includes("ADBC")) {
      remType = "Mandatory";
      remark  = "P3 ADBC finding — mandatory remediation required.";
    }
  }

  // ── Default fallback ─────────────────────────────────────────────────────
  if (!remType) {
    if (isS4H) {
      if (prio3) { remType = "Can be ignored"; remark = ""; }
      else       { remType = "Needs Remediation"; remark = ""; }
    } else {
      // HCA path — Upgrade mode: objects already exist in S/4H environment, relax mandatory to optional.
      // Exception: "CRITICAL STATEMENTS" check (e.g. Use of Database Hint) is always Mandatory
      // regardless of upgrade/conversion — these are code-quality issues independent of migration mode.
      if (isUpgrade) {
        if (prio12 && title.includes("CRITICAL STATEMENTS")) {
          remType = "Mandatory";
          remark  = "Critical statement (e.g. Database Hint) — mandatory remediation required regardless of upgrade mode.";
        } else if (prio12) { remType = "Optional"; remark = "Upgrade mode: HCA finding — object exists in S/4HANA environment, review if still relevant."; }
        else if (prio3) { remType = "Optional"; remark = "Upgrade mode: HCA prio3 finding — review if still relevant."; }
        else { remType = "Can be ignored"; remark = ""; }
      } else {
        if (prio12) { remType = "Mandatory"; remark = ""; }
        else if (prio3) { remType = "Optional"; remark = ""; }
        else { remType = "Needs Remediation"; remark = ""; }
      }
    }
  }

  // isFitGap drives Fit Gap?=Yes and propagation to sibling rows.
  // FitGapDelta rows use remType="Fit Gap" for display but must NOT trigger
  // propagation (their sibling rows may have independent classifications like
  // False Positive). Keep isFitGap=false for delta rows.
  const isFitGap      = remType === "Fit Gap" && !isFitGapDelta;
  const isSyntaxError = remType === "Syntax Error";

  return { isS4H, remType, remark, isSyntaxError, isFitGap, isFitGapDelta };
}

// ── Public API ────────────────────────────────────────────────────────────────
function classifyRows(rows, cloneBuf, migType, tadirSet, tfdirSet) {
  const _mig      = migType === "upgrade" ? "upgrade" : "conversion";
  const _tadir    = (tadirSet instanceof Set && tadirSet.size > 0) ? tadirSet : _getTadirSet();
  const _tfdir    = (tfdirSet instanceof Set && tfdirSet.size > 0) ? tfdirSet : _getTfdirSet();
  const cloneRows = cloneBuf ? readCloneRows(cloneBuf) : [];
  const cloneSet  = buildCloneSet(cloneRows);
  const hasClone  = cloneSet.size > 0;

  let classified = 0, s4hCount = 0, hcaCount = 0;

  for (const row of rows) {
    const { isS4H, remType, isSyntaxError, isFitGap, isFitGapDelta } = _classifyRow(row, _mig, _tadir, _tfdir);

    row["HCA/S4H?"]             = isS4H ? "S/4H" : "HCA";
    row["Remediation Type"]     = remType;
    row["NP Remediation Type"]  = remType;   // pre-propagation value; overwritten by propagation later
    row["Syntax Error?"]        = isSyntaxError  ? "Yes" : "No";
    row["Fit Gap?"]             = isFitGap       ? "Yes" : "No";
    row["Fit Gap Delta?"]       = isFitGapDelta  ? "Yes" : "No";

    if (hasClone) {
      const objType = (row["Obj."] || "").toString().trim().toUpperCase();
      const objName = (row["Object name"] || "").toString().trim().toUpperCase();
      row["Clone?"] = cloneSet.has(`${objType},${objName}`) ? "Yes" : "No";
    } else {
      row["Clone?"] = "No";
    }

    classified++;
    if (isS4H) s4hCount++; else hcaCount++;
  }

  // ── Post-processing: propagate Syntax Error and Fit Gap FLAGS + Remediation Type ──
  // Mirrors ZATCASSESSMENT_KK lines 3461-3535: for any object that has a Syntax Error
  // or Fit Gap row, ALL rows of that object get the corresponding flag set to 'Yes'
  // AND the Remediation Type is updated so col P matches the propagated classification.
  // propagateSyntaxError runs first (sets RemType="Syntax Error"); propagateFitGap then
  // runs but skips syn rows so Syntax Error always takes priority over Fit Gap.
  // Original pre-propagation RemType is preserved in NP Remediation Type for tracking.
  // NOTE: For multi-chunk processing, propagation is done AFTER merging all chunks
  // via propagateSyntaxError() + propagateFitGap(). Calling both here for single-chunk runs.
  rows = propagateSyntaxError(rows);
  rows = propagateFitGap(rows);

  // Recount (no propagation changes HCA/S4H assignment, just flag columns)
  s4hCount = 0; hcaCount = 0;
  for (const row of rows) {
    if (row["HCA/S4H?"] === "S/4H") s4hCount++; else hcaCount++;
  }

  console.log(`[atcClassify] Classified ${classified} rows: hca=${hcaCount} s4h=${s4hCount} cloneKeys=${cloneSet.size} migType=${_mig}`);
  return rows;
}

/**
 * Post-processing pass: if ANY row for an object is Fit Gap,
 * set Fit Gap? = 'Yes' on ALL rows of that same object AND update
 * Remediation Type to "Fit Gap" (preserving original in NP Remediation Type).
 * Syntax Error rows are skipped — Syntax Error takes priority over Fit Gap.
 * Mirrors ZATCASSESSMENT_KK lines 3499-3530 (FITGAP column propagated).
 * Must be called AFTER all chunks are merged (i.e. on the full row set).
 * Safe to call multiple times — idempotent.
 */
function propagateFitGap(rows) {
  // Build set of object keys that have at least one Fit Gap row.
  const fitGapObjKeys = new Set(); // "OBJTYPE||OBJNAME"
  for (const row of rows) {
    if (row["Fit Gap?"] === "Yes") {
      const objKey = `${(row["Obj."] || "").toString().trim().toUpperCase()}||${(row["Object name"] || "").toString().trim().toUpperCase()}`;
      fitGapObjKeys.add(objKey);
    }
  }

  if (fitGapObjKeys.size > 0) {
    let propagated = 0;
    for (const row of rows) {
      if (row["Fit Gap?"] === "Yes") continue; // already set

      const objKey = `${(row["Obj."] || "").toString().trim().toUpperCase()}||${(row["Object name"] || "").toString().trim().toUpperCase()}`;
      if (fitGapObjKeys.has(objKey)) {
        row["Fit Gap?"] = "Yes";
        // Only update Remediation Type if row is not already a Syntax Error
        // (Syntax Error takes priority — syn+fg rows keep RemType="Syntax Error")
        if (row["Syntax Error?"] !== "Yes") {
          row["NP Remediation Type"] = row["Remediation Type"]; // preserve original for tracking
          row["Remediation Type"] = "Fit Gap";                  // update to propagated type
        }
        propagated++;
      }
    }
    if (propagated > 0) {
      console.log(`[atcClassify] Fit Gap flag propagation: ${fitGapObjKeys.size} objects, ${propagated} rows flagged`);
    }
  }

  return rows;
}

/**
 * Post-processing pass: if ANY row for an object is Syntax Error,
 * set Syntax Error? = 'Yes' on ALL rows of that same object AND update
 * Remediation Type to "Syntax Error" (preserving original in NP Remediation Type).
 * Mirrors ZATCASSESSMENT_KK lines 3467-3498 (SYN_ERROR column propagated).
 * Must be called AFTER all chunks are merged (i.e. on the full row set).
 * Must be called BEFORE propagateFitGap so that SYN_ERROR takes priority.
 * Safe to call multiple times — idempotent.
 */
function propagateSyntaxError(rows) {
  // Build set of object keys that have at least one Syntax Error row.
  const syntaxObjKeys = new Set(); // "OBJTYPE||OBJNAME"
  for (const row of rows) {
    if (row["Syntax Error?"] === "Yes") {
      const objKey = `${(row["Obj."] || "").toString().trim().toUpperCase()}||${(row["Object name"] || "").toString().trim().toUpperCase()}`;
      syntaxObjKeys.add(objKey);
    }
  }

  if (syntaxObjKeys.size > 0) {
    let propagated = 0;
    for (const row of rows) {
      if (row["Syntax Error?"] === "Yes") continue; // already set

      const objKey = `${(row["Obj."] || "").toString().trim().toUpperCase()}||${(row["Object name"] || "").toString().trim().toUpperCase()}`;
      if (!syntaxObjKeys.has(objKey)) continue;

      // Set the flag and update Remediation Type; preserve original in NP Remediation Type
      row["NP Remediation Type"] = row["Remediation Type"]; // preserve original for tracking
      row["Remediation Type"] = "Syntax Error";              // update to propagated type
      row["Syntax Error?"] = "Yes";
      propagated++;
    }
    if (propagated > 0) {
      console.log(`[atcClassify] Syntax Error flag propagation: ${syntaxObjKeys.size} objects, ${propagated} rows flagged`);
    }
  }

  return rows;
}

/**
 * Pre-scan ATC rows to extract all referenced objects (type + name + funcname).
 * Mirrors the ABAP pre-scan pass in ZCL_ZATC_ASSESSMENT_HANDLER that builds it_tadirt.
 * Returns an array of { object, obj_name, funcname } entries for TADIR/TFDIR lookup.
 */
function extractObjectEntries(rows) {
  // Object types that appear in TADIR (PGMID='R3TR')
  const TADIR_TYPES = new Set([
    "DTEL","DOMA","TTYP","FUNC","INTF","CLAS","TABL","DDLS","VIEW","ENHO",
    "PROG","ENHS","FUGR","MSAG","XSLT","SUSO","SMTG","AUTH","IWSV","CHAR",
  ]);

  const seen   = new Set();
  const result = [];

  for (const row of rows) {
    const objType = (row["Obj."] || row["Object Type"] || "").toString().trim().toUpperCase();
    const objName = (row["Object name"] || row["Object Name"] || "").toString().trim().toUpperCase();
    const refObj  = (row["Referenced Object"] || row["Ref. Object Name"] || "").toString().trim().toUpperCase();
    const title   = (row["Check Title"] || row["Check title"] || "").toString().trim().toUpperCase();

    // Only collect entries for "Search for Usages of Simplified Objects" check
    if (!title.includes("SEARCH FOR USAGES OF SIMPLIFIED OBJECTS") &&
        !title.includes("S/4HANA: SEARCH FOR USAGES OF SIMPLIFIED OBJECTS")) continue;

    const name     = refObj || objName;
    const isFuncFM = objType === "FUNC" || objType === "FUGR";

    if (!name) continue;

    if (isFuncFM) {
      const key = `FUNC|${name}`;
      if (!seen.has(key)) { seen.add(key); result.push({ object: "FUNC", obj_name: name, funcname: name }); }
    } else if (TADIR_TYPES.has(objType)) {
      const key = `${objType}|${name}`;
      if (!seen.has(key)) { seen.add(key); result.push({ object: objType, obj_name: name, funcname: "" }); }
    }
  }

  return result;
}

module.exports = { classifyRows, buildCloneSet, readCloneRows, extractObjectEntries, propagateFitGap, propagateSyntaxError,
  preloadRefSets() { _getTadirSet(); _getTfdirSet(); } };
