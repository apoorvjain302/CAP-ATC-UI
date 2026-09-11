"use strict";
/**
 * excelGen.js — Node.js port of:
 *   tool_excel.py::write_estimation
 *   tool_excel.py::write_tua
 *   agent.py::_compute_spdd_spau
 *   agent.py::_read_excel_rows
 *   agent.py::_build_master_map
 *
 * Uses SheetJS (xlsx) for all Excel manipulation.
 * For TUA: we use SheetJS to write the data rows, then JSZip to restore
 * any embedded formulas that SheetJS would strip.
 */

const XLSX  = require("xlsx");
const JSZip = require("jszip");

// ── Read Excel rows from a Buffer ──────────────────────────────────────────────
function readExcelRows(buf, numCols) {
  try {
    const wb  = XLSX.read(buf, { type: "buffer", cellText: true, raw: false });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!raw.length) return [];

    // SAP list exports have a title row at index 0, then empty rows, then the real
    // column header row (which has null at col 0 and column names from col 1 onwards),
    // then an empty row, then data (also with null at col 0).
    // Find the actual header row: the first row where col 1 is a non-empty string
    // that looks like a field name (all-caps or matches known SAP field patterns).
    // Then detect whether col 0 is always null in data rows (SAP line-number column)
    // and strip it so that indices start at the real first field.
    let headerRowIdx = 0;
    for (let i = 0; i < Math.min(10, raw.length); i++) {
      const row = raw[i] || [];
      // Header row: col 0 = null, col 1 = first field name (non-numeric string)
      const c0null = (row[0] === null || row[0] === undefined || String(row[0] || "").trim() === "");
      const c1name = row[1] && typeof row[1] === "string" && /^[A-Z_]{2,}/.test(String(row[1]).trim());
      if (c0null && c1name) { headerRowIdx = i; break; }
    }

    // Detect col-0 offset: if every non-empty row has null/empty at col 0, strip it
    // (SAP exports place a row-sequence number or nothing in col A)
    let colOffset = 0;
    const sampleRows = raw.slice(headerRowIdx + 1, headerRowIdx + 10).filter(r => r && r.some(v => v !== null && v !== ""));
    if (sampleRows.length > 0 && sampleRows.every(r => (r[0] === null || r[0] === undefined || r[0] === ""))) {
      colOffset = 1;
    }

    const rows = [];
    for (let i = headerRowIdx + 1; i < raw.length; i++) {
      const row = raw[i] || [];
      const slice = row.slice(colOffset, colOffset + numCols);
      while (slice.length < numCols) slice.push(null);
      if (slice.some(v => v !== null && v !== "")) rows.push(slice);
    }
    return rows;
  } catch (err) {
    console.error("[excelGen] readExcelRows error:", err.message);
    return [];
  }
}

// ── Build master namespace owner map ──────────────────────────────────────────
function buildMasterMap(rows) {
  const map = {};
  rows.forEach((row, i) => {
    if (row.length < 2 || !row[0]) return;
    const ns  = String(row[0]).trim();
    if (!ns.startsWith("/") && i === 0) return;  // skip header
    const own = row[1] ? String(row[1]).trim() : "";
    if (own) map[ns] = own;
  });
  return map;
}

// ── Compute namespace findings (unmaintained + no-owner) ──────────────────────
// "no owner" is derived from TRNSPACET col[9] directly — NS Owner file is optional.
// masterMap (from NS Owner file) supplements missing owners but is not required.
// SAP rule: if SSCRFLAG (col 3), SAPFLAG (col 4), or GEN_ONLY (col 6) = "X",
// the namespace is SAP-owned → treat owner as "SAP SE" (never flagged as missing).
function computeNsFindings(smodiRows, trnsRows, masterMap) {
  // Extract unique namespace prefixes from SMODILOG object names (col 1)
  const smodiNs = new Set();
  for (const row of smodiRows) {
    const obj = String(row[1] || "").trim();
    if (obj.startsWith("/")) {
      const parts = obj.split("/");
      if (parts.length >= 3) smodiNs.add("/" + parts[1] + "/");
    }
  }
  // Namespaces declared in TRNSPACET (col 0 = namespace)
  const trnsNs = new Set(trnsRows.map(r => String(r[0] || "").trim()).filter(Boolean));
  // Unmaintained = in SMODILOG but not in TRNSPACET
  const unmaintained = [...smodiNs].filter(ns => !trnsNs.has(ns)).sort();
  // No owner = in TRNSPACET where resolved owner is blank
  // TRNSPACET column layout: [0]=NAMESPACE [3]=SSCRFLAG [4]=SAPFLAG [5]=GEN_ONLY
  //                          [9]=DESCRIPTN [10]=OWNER
  // SAP rule: SSCRFLAG (col 3) | SAPFLAG (col 4) | GEN_ONLY (col 5) = "X" → owner = "SAP SE"
  const noOwner = trnsRows
    .filter(r => {
      const ns  = String(r[0] || "").trim();
      if (!ns) return false;
      const owner   = r.length > 10 ? String(r[10] || "").trim() : "";
      const sscr    = r.length > 3 ? String(r[3] || "").trim().toUpperCase() : "";
      const sapflag = r.length > 4 ? String(r[4] || "").trim().toUpperCase() : "";
      const genOnly = r.length > 5 ? String(r[5] || "").trim().toUpperCase() : "";
      const isSap   = sscr === "X" || sapflag === "X" || genOnly === "X";
      const mapOwner = masterMap ? (masterMap[ns] || "") : "";
      return !owner && !isSap && !mapOwner;
    })
    .map(r => String(r[0]).trim())
    .sort();
  return { unmaintained, noOwner };
}

// ── Compute SPDD/SPAU counts ───────────────────────────────────────────────────
function computeSpddSpau(templateBuf, smodilogRows, namespaceRows, masterMap) {
  // Read Classification sheet from template
  const wb       = XLSX.read(templateBuf, { type: "buffer" });
  const classWs  = wb.Sheets["Classification"];
  if (!classWs) throw new Error("Classification sheet not found in TUA template");

  const classRaw = XLSX.utils.sheet_to_json(classWs, { header: 1, defval: null });
  const classMap = {};
  for (let i = 1; i < classRaw.length; i++) {
    const r = classRaw[i];
    const objType  = r[1] ? String(r[1]).trim() : "";
    const category = r[3] ? String(r[3]).trim() : "";
    const relevant = r[4] ? String(r[4]).trim() : "";
    if (objType) classMap[objType] = { category, relevant };
  }

  // Build namespace → owner map from namespace rows
  // TRNSPACET column layout: [0]=NAMESPACE [3]=SSCRFLAG [4]=SAPFLAG [5]=GEN_ONLY
  //                          [9]=DESCRIPTN [10]=OWNER
  // SAP rule: SSCRFLAG (col 3) | SAPFLAG (col 4) | GEN_ONLY (col 5) = "X" → owner = "SAP SE"
  const nsOwnerMap  = {};
  const nsSapflagMap= {};
  for (const row of namespaceRows) {
    if (!row[0]) continue;
    const ns       = String(row[0]).trim();
    const trnOwner = (row.length > 10 && row[10]) ? String(row[10]).trim() : "";
    const sscr     = (row.length > 3 && row[3]) ? String(row[3]).trim().toUpperCase() : "";
    const sap      = (row.length > 4 && row[4]) ? String(row[4]).trim().toUpperCase() : "";
    const genOnly  = (row.length > 5 && row[5]) ? String(row[5]).trim().toUpperCase() : "";
    const isSap    = sscr === "X" || sap === "X" || genOnly === "X";
    nsSapflagMap[ns] = isSap;
    nsOwnerMap[ns]   = trnOwner || (isSap ? "SAP SE" : (masterMap[ns] || ""));
  }

  function _extractNamespace(objName) {
    if (!objName) return null;
    if (objName.startsWith("/")) {
      const pos2 = objName.indexOf("/", 1);
      return pos2 !== -1 ? objName.slice(0, pos2 + 1) : null;
    }
    // Non-slash objects: Z/Y prefix = customer-owned, anything else = SAP-owned
    return /^[ZY]/i.test(objName) ? "__CUSTOMER__" : "__SAP__";
  }

  const spddConcats = new Set();
  const spauConcats = new Set();

  for (const row of smodilogRows) {
    if (row.length < 10) continue;
    const objName  = row[1] ? String(row[1]).trim() : "";
    const operation= row[6] ? String(row[6]).trim() : "";
    const subType  = row[2] ? String(row[2]).trim() : "";
    const subName  = row[3] ? String(row[3]).trim() : "";
    const modName  = row[9] ? String(row[9]).trim() : "";

    let finalObjType, finalObjName, sapOwned;

    if (operation === "NOTE") {
      finalObjType = "NOTE";
      try {
        finalObjName = String(Math.floor(parseFloat(modName) / 10000));
      } catch (_) {
        finalObjName = modName;
      }
      sapOwned = true;
    } else {
      finalObjType = subType;
      finalObjName = subName;
      const nsKey  = _extractNamespace(objName);
      if (nsKey === "__CUSTOMER__") {
        sapOwned = false;
      } else if (nsKey === "__SAP__") {
        sapOwned = true;
      } else if (nsKey === null) {
        sapOwned = false;
      } else {
        // Slash-prefix namespace: look up in namespace owner map
        const owner  = nsOwnerMap[nsKey] || "";
        if (owner.toUpperCase().startsWith("SAP"))  sapOwned = true;
        else if (!owner && nsSapflagMap[nsKey])      sapOwned = true;
        else                                         sapOwned = false;
      }
    }

    if (!sapOwned) continue;

    const concat = `${finalObjType}|${finalObjName}`;
    const cls    = classMap[finalObjType];
    if (!cls || cls.relevant !== "Yes") continue;

    if (cls.category === "SPDD") spddConcats.add(concat);
    else if (cls.category === "SPAU") spauConcats.add(concat);
  }

  return { spddCount: spddConcats.size, spauCount: spauConcats.size };
}

// ── Write Estimation Excel ─────────────────────────────────────────────────────
// Uses JSZip to patch cells directly into the template ZIP so all formatting,
// styles, formulas, and print settings are preserved.
async function writeEstimation(templateBuf, hcaCount, s4Count, spddCount, spauCount, d4AutoCount, d5AutoCount, d9AutoCount) {
  const zip = await JSZip.loadAsync(templateBuf);

  // Find the first worksheet file
  const wsEntry = zip.file("xl/worksheets/sheet1.xml");
  if (!wsEntry) throw new Error("xl/worksheets/sheet1.xml not found in Estimation Template");

  let xml = await wsEntry.async("string");

  // Patch or replace a cell with a numeric value.
  // Handles: full <c ...><f>...</f><v>...</v></c>, full without <f>, self-closing <c .../>,
  // and cells that don't exist at all (inserted before </row>).
  const _setCell = (x, ref, value) => {
    const num    = Number(value);
    const rowNum = parseInt(ref.replace(/[A-Z]+/, ""), 10);

    // 1. Full cell with <v> (with or without <f>, including shared formulas):
    //    <c r="REF"...><f...>...</f><v>OLD</v></c>  or  <c r="REF"...><f.../><v>OLD</v></c>
    const fullRe = new RegExp(`(<c r="${ref}"[^>]*>(?:<f[^>]*>[^<]*</f>|<f[^/]*/?>)?<v>)[^<]*(</v>)`);
    if (fullRe.test(x)) {
      return x.replace(fullRe, `$1${num}$2`);
    }

    // 2. Self-closing cell: <c r="REF" .../> — replace with full cell preserving style
    const selfRe = new RegExp(`<c r="${ref}"([^>]*)/>`);
    const selfM  = x.match(selfRe);
    if (selfM) {
      const styleAttr = selfM[1];
      return x.replace(selfRe, `<c r="${ref}"${styleAttr}><v>${num}</v></c>`);
    }

    // 3. Cell doesn't exist: insert before </row> of the matching row
    const rowRe = new RegExp(`(<row r="${rowNum}"[^>]*>[\\s\\S]*?)(</row>)`);
    if (rowRe.test(x)) {
      return x.replace(rowRe, `$1<c r="${ref}"><v>${num}</v></c>$2`);
    }

    return x;  // row not found — no-op
  };

  // Clear a cell — remove <v> content so the cell is blank, preserving style.
  // Handles: full cell with <v>, self-closing cell (already blank — no-op).
  const _clearCell = (x, ref) => {
    // Full cell with optional formula and value: strip <f> and <v> entirely, keep <c> tag with style
    const fullRe = new RegExp(`(<c r="${ref}"[^>]*>)(?:<f[^>]*>[^<]*<\\/f>|<f[^/]*\\/>)?<v>[^<]*<\\/v>(<\\/c>)`);
    if (fullRe.test(x)) {
      return x.replace(fullRe, `$1$2`);
    }
    // Self-closing cell is already blank — no-op
    return x;
  };

  // ── Input cells ───────────────────────────────────────────────────────────
  xml = _setCell(xml, "C2", spddCount);
  xml = _setCell(xml, "C3", spauCount);
  xml = _setCell(xml, "C4", hcaCount);
  xml = _setCell(xml, "C5", s4Count);

  // ── Pre-compute ALL formula cells ─────────────────────────────────────────
  // E2 has formula C2-D2 (D2 is self-closing/blank in template).
  // E2 = spddCount, E3 = spauCount (D2/D3 are blank so E2=C2, E3=C3).
  // G2 = E2/F2, G3 = E3/F3 (F2=50, F3=50).
  const f2 = 50, f3 = 50, f4 = 50, f5 = 35;
  const g2 = spddCount / f2;
  const g3 = spauCount / f3;
  const h2 = g2 * 8;
  const h3 = g3 * 8;

  const d4 = (d4AutoCount !== undefined ? d4AutoCount : hcaCount * 0.5);
  const d5 = (d5AutoCount !== undefined ? d5AutoCount : s4Count  * 0.1);
  const e4 = hcaCount - d4;
  const e5 = s4Count  - d5;
  const g4 = f4 ? e4 / f4 : 0;
  const g5 = f5 ? e5 / f5 : 0;
  const h4 = g4 * 8;
  const h5 = g5 * 8;

  // Row 6
  // C6 = SUM(C2:C5) = spddCount + spauCount + hcaCount + s4Count
  const c6 = spddCount + spauCount + hcaCount + s4Count;
  // D6 = SUM(D2:D5) = 0 + 0 + D4 + D5
  const d6 = d4 + d5;
  // E6 = SUM(E2:E5) = E2 + E3 + E4 + E5 = spddCount + spauCount + e4 + e5
  const e6 = spddCount + spauCount + e4 + e5;
  const g6 = g2 + g3 + g4 + g5;        // SUM(G2:G5)
  const h6 = h2 + h3 + h4 + h5;        // SUM(H2:H5)

  // Row 7: Lead Effort = 30% of g6 (C7 is blank per user requirement, G7 = 30% of subtotal)
  // Row 8: Testing Effort = 35% of g6
  const g7 = g6 * 0.3;
  const g8 = g6 * 0.35;
  const h7 = g7 * 8;
  const h8 = g8 * 8;

  // Rows 9-11: G9/G10 cleared (blank); H9 = IF(G9="","",G9*8); G11=0, H11=0
  const g10 = 0;
  const g11 = 0, h11 = 0;

  // Row 12: G12=SUM(G6:G11), H12=SUM(H6:H11)
  const c12 = g6 + g7 + g8 + 0 + g10 + g11;
  const h12 = h6 + h7 + h8 + 0 + 0   + h11;

  // ── Apply all computed values ─────────────────────────────────────────────
  // Row 2: E2 = spddCount (formula C2-D2, D2 blank so E2=C2); G2/H2 have formulas
  xml = _setCell(xml, "E2", spddCount);
  xml = _setCell(xml, "G2", g2);
  xml = _setCell(xml, "H2", h2);
  // Row 3: E3 = spauCount; G3/H3 have formulas
  xml = _setCell(xml, "E3", spauCount);
  xml = _setCell(xml, "G3", g3);
  xml = _setCell(xml, "H3", h3);
  // Row 4: D4/E4/G4/H4 all have formulas, patch <v> then strip <f> from D4 (auto count)
  xml = _setCell(xml, "D4", d4);
  xml = _setCell(xml, "E4", e4);
  xml = _setCell(xml, "G4", g4);
  xml = _setCell(xml, "H4", h4);
  xml = xml.replace(/(<c r="D4"[^>]*>)<f>[^<]*<\/f>(<v>)/, "$1$2");
  // Row 5: D5/E5/G5/H5 all have formulas, patch <v> then strip <f> from D5 (auto count)
  xml = _setCell(xml, "D5", d5);
  xml = _setCell(xml, "E5", e5);
  xml = _setCell(xml, "G5", g5);
  xml = _setCell(xml, "H5", h5);
  xml = xml.replace(/(<c r="D5"[^>]*>)<f>[^<]*<\/f>(<v>)/, "$1$2");

  // ── D9: CCM Agent eligible count ─────────────────────────────────────────────
  xml = _setCell(xml, "D9", d9AutoCount);
  xml = xml.replace(/(<c r="D9"[^>]*>)<f>[^<]*<\/f>(<v>)/, "$1$2");

  // Row 6: C6 = SUM(C2:C5) — replace formula in C6 with static value and update <v>
  // Template has C6 formula SUM(D2:D5); rewrite to SUM(C2:C5) and patch <v>
  xml = xml.replace(/(<c r="C6"[^>]*>)<f>[^<]*<\/f>(<v>)[^<]*(<\/v>)/,
    `$1<f>SUM(C2:C5)</f>$2${c6}$3`);
  // If C6 is self-closing or missing <f>, fall back to _setCell
  if (!xml.includes('<f>SUM(C2:C5)</f>')) xml = _setCell(xml, "C6", c6);
  xml = _setCell(xml, "D6", d6);   // D6 SUM(D2:D5)
  xml = _setCell(xml, "E6", e6);
  xml = _setCell(xml, "G6", g6);
  xml = _setCell(xml, "H6", h6);
  // Row 7: G7 = 30% of subtotal; H7 = G7*8
  xml = _setCell(xml, "G7", g7);
  xml = _setCell(xml, "H7", h7);
  // Row 8: G8 = 35% of subtotal; H8 = G8*8
  xml = _setCell(xml, "G8", g8);
  xml = _setCell(xml, "H8", h8);
  // Row 9: G9 cleared (blank); H9 formula IF(G9="","",G9*8) — handled below
  // Row 10: G10 = 0
  xml = _setCell(xml, "G10", g10);
  xml = _setCell(xml, "H10", g10 * 8);  // H10 = G10*8 = 0
  xml = _setCell(xml, "G11", g11);
  xml = _setCell(xml, "H11", h11);
  // Row 12: G12/H12 have formulas SUM(G6:G11)/SUM(H6:H11); C12 is self-closing (label)
  xml = _setCell(xml, "G12", c12);
  xml = _setCell(xml, "H12", h12);

  // ── Clear cells that must be blank in the output ──────────────────────────
  // C7, C8: user requirement — blank input cells
  // G9: static template value cleared (user fills manually)
  // G10: already set to 0 above; H10 already set to 0
  // C14,C20,D15,D20,E16,E20,F16,F17,F20,G17,G18,H18,H20: template pre-fills cleared
  const _cellsToClear = [
    "C7", "C8",    // blank per user requirement
    "G9",          // user fills manually
    "C14", "C20",
    "D15", "D20",
    "E16", "E20",
    "F16", "F17", "F20",
    "G17", "G18", "G20",
    "H18", "H20",
  ];
  for (const ref of _cellsToClear) xml = _clearCell(xml, ref);

  // ── Restore H9 formula IF(G9="","",G9*8) ─────────────────────────────────
  // Template H9 is a static value. Replace with conditional formula so H9
  // stays blank when G9 is blank (user fills G9 manually).
  xml = xml.replace(
    /(<c r="H9"[^>]*>)(?:<f[^>]*>[^<]*<\/f>)?<v>[^<]*<\/v>(<\/c>)/,
    `$1<f>IF(G9="","",G9*8)</f>$2`
  );
  // If H9 is self-closing (no <v>), insert full cell with formula
  if (!xml.includes('<c r="H9"')) {
    xml = xml.replace(/(<row[^>]*\br="9"[^>]*>[\s\S]*?)(<\/row>)/,
      `$1<c r="H9"><f>IF(G9="","",G9*8)</f></c>$2`);
  }

  zip.file("xl/worksheets/sheet1.xml", xml);

  // Remove stale calcChain so Excel rebuilds it cleanly on first open.
  // Without this, Excel shows "Removed Records: Formula from /xl/calcChain.xml" repair dialog
  // because our formula edits (stripping <f> tags, inserting static <v>) leave calcChain stale.
  zip.remove("xl/calcChain.xml");

  // Set fullCalcOnLoad so desktop Excel recalculates any remaining dependent formulas
  const wbEntry = zip.file("xl/workbook.xml");
  if (wbEntry) {
    let wbXml = await wbEntry.async("string");
    wbXml = wbXml.replace(/calcId="[^"]*"/g, "");
    if (wbXml.includes("fullCalcOnLoad=")) {
      wbXml = wbXml.replace(/fullCalcOnLoad="[^"]*"/, `fullCalcOnLoad="1"`);
    } else {
      wbXml = wbXml.replace(/(<calcPr)/, `$1 fullCalcOnLoad="1"`);
    }
    zip.file("xl/workbook.xml", wbXml);
  }

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ── Write TUA Analysis Excel ───────────────────────────────────────────────────
/**
 * Writes SMODILOG + Namespace rows into the TUA template.
 * Uses SheetJS to paste data rows, then uses JSZip to restore any
 * Power Pivot / custom XML entries that SheetJS would strip (same as Python).
 * Returns a Buffer.
 */
async function writeTua(templateBuf, smodilogRows, namespaceRows, masterMap, spddCount, spauCount) {
  const wb      = XLSX.read(templateBuf, { type: "buffer" });
  const wsModi  = wb.Sheets["SMODILOG"];
  const wsNs    = wb.Sheets["Namespace"];

  if (!wsModi) throw new Error("SMODILOG sheet not found in TUA template");
  if (!wsNs)   throw new Error("Namespace sheet not found in TUA template");

  // Paste SMODILOG data (starting row 2, 24 cols)
  _pasteRows(wsModi, smodilogRows, 24);

  // Remap TRNSPACET rows (11 cols) to TUA Namespace sheet layout (10 cols).
  // TRNSPACET: [0..8]=same, [9]=DESCRIPTN, [10]=OWNER
  // TUA Namespace sheet: [0..8]=same, [9]=Owner
  // → drop col 9 (DESCRIPTN), move col 10 (OWNER) to position 9.
  const nsRowsRemapped = namespaceRows.map(r => {
    const remapped = r.slice(0, 9);          // cols 0-8
    remapped.push(r.length > 10 ? (r[10] || null) : null);  // col 10 → position 9
    return remapped;
  });

  // Paste Namespace data (10 cols matching TUA template: col 9 = Owner)
  _pasteRows(wsNs, nsRowsRemapped, 10);

  // Fill blank owners from master map (col index 9 = col J in Namespace sheet)
  _fillOwnerFromMaster(wsNs, masterMap);

  // Update sheet references
  _updateSheetRange(wsModi, smodilogRows.length + 1);
  _updateSheetRange(wsNs,   namespaceRows.length + 1);

  // Write to buffer (SheetJS will strip power pivot — we'll patch with JSZip below)
  const outBuf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  // Restore Power Pivot entries + patch pivot cells in a single JSZip pass
  return _restorePowerPivot(templateBuf, outBuf, spddCount, spauCount);
}

function _pasteRows(ws, rows, numCols) {
  if (!ws["!ref"]) return;

  // Start from row 2 (data rows; row 1 is header)
  rows.forEach((row, ri) => {
    const excelRow = ri + 2;  // 1-indexed, skip header
    for (let c = 0; c < numCols; c++) {
      const cellRef = XLSX.utils.encode_cell({ r: excelRow - 1, c });
      const val     = row[c];
      if (val === null || val === undefined || val === "") {
        delete ws[cellRef];
      } else {
        const isNum = typeof val === "number" || (typeof val === "string" && !isNaN(val) && val.trim() !== "");
        ws[cellRef] = isNum ? { t: "n", v: Number(val) } : { t: "s", v: String(val) };
      }
    }
  });

  // Update sheet !ref
  const maxRow = rows.length + 1;
  const maxCol = XLSX.utils.decode_range(ws["!ref"]).e.c;
  ws["!ref"] = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: maxRow - 1, c: maxCol });
}

function _fillOwnerFromMaster(wsNs, masterMap) {
  const range = XLSX.utils.decode_range(wsNs["!ref"] || "A1");
  for (let row = 1; row <= range.e.r; row++) {  // 0-indexed; row 0 = header
    const nsCell    = wsNs[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const ownerCell = wsNs[XLSX.utils.encode_cell({ r: row, c: 9 })];
    if (!nsCell || !nsCell.v) continue;
    const ns    = String(nsCell.v).trim();
    const owner = ownerCell ? String(ownerCell.v || "").trim() : "";
    if (owner) continue;
    // SAP rule: SSCRFLAG (col 3) | SAPFLAG (col 4) | GEN_ONLY (col 5) = "X" → "SAP SE"
    const sscrCell    = wsNs[XLSX.utils.encode_cell({ r: row, c: 3 })];
    const sapflagCell = wsNs[XLSX.utils.encode_cell({ r: row, c: 4 })];
    const genOnlyCell = wsNs[XLSX.utils.encode_cell({ r: row, c: 5 })];
    const sscr    = sscrCell    ? String(sscrCell.v    || "").trim().toUpperCase() : "";
    const sapflag = sapflagCell ? String(sapflagCell.v || "").trim().toUpperCase() : "";
    const genOnly = genOnlyCell ? String(genOnlyCell.v || "").trim().toUpperCase() : "";
    const resolvedOwner = (sscr === "X" || sapflag === "X" || genOnly === "X")
      ? "SAP SE"
      : (masterMap[ns] || "");
    if (resolvedOwner) {
      const ref = XLSX.utils.encode_cell({ r: row, c: 9 });
      wsNs[ref] = { t: "s", v: resolvedOwner };
    }
  }
}

function _updateSheetRange(ws, lastRow) {
  if (!ws["!ref"]) return;
  const range  = XLSX.utils.decode_range(ws["!ref"]);
  ws["!ref"]   = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: lastRow - 1, c: range.e.c });
}

// Merged: restores Power Pivot entries AND patches pivot cells in a single JSZip pass,
// avoiding the extra loadAsync + generateAsync cycle that _patchPivotCells previously added.
async function _restorePowerPivot(templateBuf, outBuf, spdd, spau) {
  const templateZip = await JSZip.loadAsync(templateBuf);
  const outZip      = await JSZip.loadAsync(outBuf);

  // Copy xl/model/* and customXml/* from template into output
  const restoreKeys = Object.keys(templateZip.files).filter(
    k => k.startsWith("xl/model/") || k.startsWith("customXml/")
  );
  for (const key of restoreKeys) {
    const data = await templateZip.file(key).async("nodebuffer");
    outZip.file(key, data);
  }

  // Restore xl/_rels/workbook.xml.rels from template
  const wbRels = templateZip.file("xl/_rels/workbook.xml.rels");
  if (wbRels) {
    const data = await wbRels.async("nodebuffer");
    outZip.file("xl/_rels/workbook.xml.rels", data);
  }

  // Patch pivot cells (B5=spdd, B6=spau, B7=total) in sheet2.xml — same pass, no extra load
  const sheet2 = outZip.file("xl/worksheets/sheet2.xml");
  if (sheet2) {
    let xml = await sheet2.async("string");
    xml = xml.replace(/(<c r="B5"[^>]*><v>)[^<]*(<\/v><\/c>)/, `$1${spdd}$2`);
    xml = xml.replace(/(<c r="B6"[^>]*><v>)[^<]*(<\/v><\/c>)/, `$1${spau}$2`);
    xml = xml.replace(/(<c r="B7"[^>]*><v>)[^<]*(<\/v><\/c>)/, `$1${spdd + spau}$2`);
    outZip.file("xl/worksheets/sheet2.xml", xml);
  }

  return outZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = {
  readExcelRows,
  buildMasterMap,
  computeNsFindings,
  computeSpddSpau,
  writeEstimation,
  writeTua,
};
