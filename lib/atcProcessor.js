"use strict";
/**
 * atcProcessor.js
 * Runs outside a CDS request context (called via setImmediate).
 * Must use cds.db.run() with explicit query objects — NOT the global SELECT/INSERT/UPDATE.
 */

const cds      = require("@sap/cds");
const path     = require("path");
const fs       = require("fs");

const sapApi      = require("./sapApi");
const xlsData     = require("./xlsData");
const pptxGen     = require("./pptxGen");
const excelGen    = require("./excelGen");
const atcClassify = require("./atcClassify");
const { toBuffer, toHex } = require("./bufferUtil");

// Maximum rows per ABAP call. Files with <= CHUNK_SIZE rows are sent as-is (no split).
// Files exceeding this are split into sequential chunks; each chunk gets the full clone file.
//
// 3 000 rows per chunk keeps each ABAP result small enough (~3-5 MB SpreadsheetML)
// for Node.js to parse without OOM inside the 512 MB CF memory limit, and also stays
// well within ABAP's internal processing time / row limits.
const CHUNK_SIZE = 1500;

async function processJob(jobId) {
  return _runJob(jobId);
}

async function _runJob(jobId) {

  // Wait until cds.db is available (set by CAP after bootstrap)
  let db;
  for (let i = 0; i < 20; i++) {
    if (cds.db) { db = cds.db; break; }
    await new Promise(r => setTimeout(r, 500));
  }
  if (!db) throw new Error("cds.db not available after 10s");

  const _update = async (msg, extra = {}) => {
    await db.run(UPDATE("atc.Jobs").set({ statusMsg: msg, ...extra }).where({ id: jobId }));
    console.log(`[ATC][${jobId}] ${msg}`);
  };

  try {
    // ── Load job ────────────────────────────────────────────────────────────
    const [job] = await db.run(SELECT.from("atc.Jobs").where({ id: jobId }));
    if (!job) throw new Error(`Job ${jobId} not found`);

    const useAppLogic = job.useAppLogic === true || job.useAppLogic === 1;
    const migType     = job.migType === "upgrade" ? "upgrade" : "conversion";

    // Derive analysis mode flags from stored analysisMode field
    const analysisMode = job.analysisMode || "atc";
    const tuaOnly = analysisMode === "tua_only";
    const useTua  = tuaOnly || analysisMode === "atc_tua" || analysisMode === "atc_tua_clone";

    await _update("Loading input files...");

    // ── Load files ──────────────────────────────────────────────────────────
    let atcFile = null;
    if (!tuaOnly) {
      [atcFile] = await db.run(`SELECT id, filename, mimeType, size, content FROM atc_Files WHERE id = ?`, [job.atcFileId]);
      if (!atcFile) throw new Error("ATC input file not found in database");
    }
    const _loadFile = async (id) => {
      if (!id) return null;
      const [f] = await db.run(`SELECT id, filename, mimeType, size, content FROM atc_Files WHERE id = ?`, [id]);
      return f || null;
    };

    const cloneFile   = await _loadFile(job.cloneFileId);
    const smodiFile   = await _loadFile(job.smodilogId);
    const trnFile     = await _loadFile(job.trnspacetId);
    const nsOwnerFile = await _loadFile(job.nsOwnerId);
    const hasTuaFiles = !!(smodiFile && trnFile);

    let dataResult = null;
    let atcData = null, totalCount = 0, hcaCount = 0, s4Count = 0;
    let d4AutoCount = 0, d5AutoCount = 0, d9AutoCount = 0;
    let hcaMandatory = 0, s4TechRemediable = 0, fitGapDeltaCount = 0;
    let classifiedXlsBuf = null;

    if (!tuaOnly) {
    // ── Step 1: Classify / enrich ATC data ─────────────────────────────────
    // Path A (useAppLogic=true):  classify rows in-process using atcClassify.js —
    //                             no I10 API call, no Cloud Connector needed.
    // Path B (useAppLogic=false): send chunks to SAP ATC API on I10 (existing path).
    // Both paths produce xlsResultBufs[] consumed identically by Step 2 onwards.

    // ── Buffer loading (shared by both paths) ──────────────────────────────
    const atcBuf   = await toBuffer(atcFile.content);
    if (!atcBuf) throw new Error("ATC file content is empty");
    const cloneBuf = cloneFile ? await toBuffer(cloneFile.content) : null;

    // Remove duplicate columns (ABAP's create_dynamic_table crashes on duplicates)
    const dedupAtcBuf   = await xlsData.deduplicateXlsxColumns(atcBuf);
    const dedupCloneBuf = cloneBuf ? await xlsData.deduplicateXlsxColumns(cloneBuf) : null;

    const xlsResultBufs = [];

    if (useAppLogic) {
      await _update("Classifying ATC data using app logic...", { status: "running" });
      console.log(`[ATC][${jobId}] Using app-logic classification (no I10 API call)`);

      // Direct row accumulation path — no XLSX round-trip per chunk.
      // Each chunk is: read buffer → parse rows → classify in-place → accumulate rows → free buffer.
      // After all chunks: propagate → compute → buildClassifiedXlsx once.
      // This avoids holding 20-50+ serialized XLSX buffers in memory simultaneously.
      const chunkBufs2 = await xlsData.splitXlsxRows(dedupAtcBuf, CHUNK_SIZE);
      if (chunkBufs2.length > 1) {
        await _update(`Classifying ${chunkBufs2.length} chunks using app logic...`);
      }

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
      };

      const allClassifiedRows = [];

      for (let i = 0; i < chunkBufs2.length; i++) {
        if (chunkBufs2.length > 1) {
          await _update(`Classifying chunk ${i + 1}/${chunkBufs2.length}...`);
        }

        // Parse chunk to rows
        const XLSX = require("xlsx");
        const wb   = XLSX.read(chunkBufs2[i], { type: "buffer", cellText: true, raw: false });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        chunkBufs2[i] = null;  // free chunk buffer immediately after parse

        // Normalise and classify in-place
        const rows = rawRows.map(r => {
          const out = {};
          for (const k of Object.keys(r)) {
            const canonical = _COL_CANONICAL[k.trim().toLowerCase()] || k.trim();
            out[canonical] = String(r[k] === null || r[k] === undefined ? "" : r[k]).trim();
          }
          return out;
        });

        atcClassify.classifyRows(rows, dedupCloneBuf, migType);
        allClassifiedRows.push(...rows);
      }

      // Cross-chunk propagation on merged row set
      await _update("Computing analysis results...");
      atcClassify.propagateSyntaxError(allClassifiedRows);
      atcClassify.propagateFitGap(allClassifiedRows);

      const dr = xlsData.readXlsDataFromRows(allClassifiedRows);
      if (dr.error) throw new Error(`XLS parse error: ${dr.error}`);

      classifiedXlsBuf = await xlsData.buildClassifiedXlsx(allClassifiedRows);
      allClassifiedRows.length = 0; // free rows after XLSX is built

      ({ atcData, totalCount, hcaCount, s4Count, d4AutoCount, d5AutoCount, d9AutoCount,
         hcaMandatory, s4TechRemediable, fitGapDeltaCount } = dr);

      if (classifiedXlsBuf) {
        await db.run(
          `INSERT INTO atc_Artifacts (id, jobId, role, filename, mimeType, content, size) VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [cds.utils.uuid(), jobId, "atc_result", `atc_classified_${_ts()}.xlsx`,
           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
           toHex(classifiedXlsBuf), classifiedXlsBuf.length]
        );
        classifiedXlsBuf = null; // free immediately — large buffer no longer needed
      }

    } else {
      // ── Path B: existing I10 API path (unchanged) ───────────────────────
      const chunkBufs = await xlsData.splitXlsxRows(dedupAtcBuf, CHUNK_SIZE);
      if (chunkBufs.length > 1) {
        await _update(`File split into ${chunkBufs.length} chunks of up to ${CHUNK_SIZE} rows each...`);
      }
      for (let i = 0; i < chunkBufs.length; i++) {
        const chunkLabel = chunkBufs.length > 1 ? ` (chunk ${i + 1}/${chunkBufs.length})` : "";
        await _update(`Calling SAP ATC API${chunkLabel}...`, { status: "running" });
        const apiResult = await sapApi.callAtcApi(
          chunkBufs[i], atcFile.filename,
          dedupCloneBuf, cloneFile ? cloneFile.filename : null,
        );
        chunkBufs[i] = null;
        if (apiResult.error) throw new Error(`SAP API error${chunkLabel}: ${apiResult.error}`);

        // Store each API chunk result as its own artifact (for individual download)
        if (chunkBufs.length > 1 && apiResult.xlsBuffer) {
          const chunkBuf = apiResult.xlsBuffer;
          await db.run(
            `INSERT INTO atc_Artifacts (id, jobId, role, filename, mimeType, content, size) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [cds.utils.uuid(), jobId, `atc_result_chunk_${i + 1}`,
             `atc_api_chunk_${i + 1}_of_${chunkBufs.length}_${_ts()}.xls`,
             "application/vnd.ms-excel", toHex(chunkBuf), chunkBuf.length]
          );
        }

        xlsResultBufs.push(apiResult.xlsBuffer);
      }
    }

    if (!useAppLogic) {
    // ── Step 2: Parse XLS result (Path B — I10 API only) ───────────────────
    // useAppLogic path already extracted atcData and stored the artifact above.
    await _update("Parsing ATC results...");

    let dataResult;
    let classifiedXlsBuf = null;

    if (xlsResultBufs.length === 1) {
      // Single-chunk path — classify rows first, then compute dataResult from classified rows
      const rawBuf = xlsResultBufs[0];
      try {
        const XLSX = require("xlsx");
        const wb   = XLSX.read(rawBuf, { type: "buffer", cellText: true, raw: false });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
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
        };
        const normalised = rows.map(r => {
          const out = {};
          for (const k of Object.keys(r)) {
            const canonical = _COL_CANONICAL[k.trim().toLowerCase()] || k.trim();
            out[canonical] = String(r[k] === null || r[k] === undefined ? "" : r[k]).trim();
          }
          return out;
        });
        atcClassify.propagateSyntaxError(normalised);
        atcClassify.propagateFitGap(normalised);
        dataResult = xlsData.readXlsDataFromRows(normalised);
        classifiedXlsBuf = await xlsData.buildClassifiedXlsx(normalised);
      } catch (e) {
        console.warn(`[ATC][${jobId}] Classification failed, falling back to raw parse:`, e.message);
        dataResult = xlsData.readXlsData(rawBuf);
        classifiedXlsBuf = rawBuf;
      }
    } else {
      // Multi-chunk path
      const { dataResult: dr, xlsBuf: xb } = await xlsData.processMultiChunkResults(xlsResultBufs, {
        applyClassify: false,
        cloneBuf:      dedupCloneBuf,
        migType,
        tadirSet:      new Set(),
        tfdirSet:      new Set(),
      });
      dataResult       = dr;
      classifiedXlsBuf = xb;
      xlsResultBufs.length = 0;
    }

    if (dataResult.error) throw new Error(`XLS parse error: ${dataResult.error}`);

    ({ atcData, totalCount, hcaCount, s4Count, d4AutoCount, d5AutoCount, d9AutoCount,
       hcaMandatory, s4TechRemediable, fitGapDeltaCount } = dataResult);

    if (classifiedXlsBuf) {
      await db.run(
        `INSERT INTO atc_Artifacts (id, jobId, role, filename, mimeType, content, size) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [cds.utils.uuid(), jobId, "atc_result", `atc_classified_${_ts()}.xlsx`,
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
         toHex(classifiedXlsBuf), classifiedXlsBuf.length]
      );
    }
    } // end if (!useAppLogic) — Step 2

    } // end if (!tuaOnly)

    let spddCount = 0, spauCount = 0;
    let nsFindings = { unmaintained: [], noOwner: [] };

    // ── Step 3: TUA ─────────────────────────────────────────────────────────
    if (useTua && hasTuaFiles) {
      await _update(tuaOnly ? "Running TUA analysis..." : "Generating TUA analysis...");
      const smodiRows = excelGen.readExcelRows(await toBuffer(smodiFile.content), 24);
      const trnsRows  = excelGen.readExcelRows(await toBuffer(trnFile.content),   11);
      const masterMap = nsOwnerFile
        ? excelGen.buildMasterMap(excelGen.readExcelRows(await toBuffer(nsOwnerFile.content), 2))
        : {};

      // ── SMODILOG validation: check for Z* custom objects ───────────────────
      // SMODILOG should only contain namespace-based objects (e.g. /NAMESPACE/OBJ).
      // If Z* or Y* objects are present, the file is likely an incorrect export
      // (e.g. SE16 export of SMODILOG table with non-namespace objects included).
      const zStarRows = smodiRows.filter(r => {
        const objName = String(r[1] || "").trim();
        return /^[ZzYy]/.test(objName);
      });
      if (zStarRows.length > 0) {
        const zSamples = [...new Set(zStarRows.slice(0, 5).map(r => String(r[1] || "").trim()))].join(", ");
        console.warn(`[ATC][${jobId}] SMODILOG WARNING: ${zStarRows.length} row(s) with Z*/Y* object names found (e.g. ${zSamples}). File may be incorrect.`);
        await _update(
          `WARNING: SMODILOG file contains ${zStarRows.length} Z*/Y* custom object(s) (e.g. ${zSamples.slice(0, 120)}). ` +
          `SMODILOG should only contain namespace-based objects (/NAMESPACE/...). ` +
          `Please verify the correct SMODILOG export was uploaded. Continuing analysis with available data...`
        );
      }

      // Compute namespace findings (unmaintained + no-owner)
      nsFindings = excelGen.computeNsFindings(smodiRows, trnsRows, masterMap);

      const tuaTemplate = _loadTemplate("TUA_Analysis.xlsx");
      if (!tuaTemplate) throw new Error("TUA_Analysis.xlsx template not found in db/templates");

      const counts = excelGen.computeSpddSpau(tuaTemplate, smodiRows, trnsRows, masterMap);
      spddCount = counts.spddCount;
      spauCount = counts.spauCount;

      const tuaBuf = await excelGen.writeTua(tuaTemplate, smodiRows, trnsRows, masterMap, spddCount, spauCount);
      await db.run(
        `INSERT INTO atc_Artifacts (id, jobId, role, filename, mimeType, content, size) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [cds.utils.uuid(), jobId, "tua", `TUA_Analysis_${_ts()}.xlsx`,
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
         toHex(tuaBuf), tuaBuf.length]
      );
    }

    // ── Step 4: PowerPoint ──────────────────────────────────────────────────
    await _update("Generating PowerPoint presentation...");
    const pptxTemplate = _loadTemplate("template.pptx");
    if (!pptxTemplate) throw new Error("template.pptx not found in db/templates");

    // For TUA-only mode, atcData is null — pptxGen will skip ATC-specific chart slides
    const pptxBuf = await pptxGen.writePptx(pptxTemplate, atcData, spddCount, spauCount, job.customer);
    const pptxArtId = cds.utils.uuid();
    await db.run(
      `INSERT INTO atc_Artifacts (id, jobId, role, filename, mimeType, content, size) VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [pptxArtId, jobId, "pptx", `atc_presentation_${_ts()}.pptx`,
       "application/vnd.openxmlformats-officedocument.presentationml.presentation",
       toHex(pptxBuf), pptxBuf.length]
    );

    // ── Step 5: Estimation Excel (skip for TUA-only) ──────────────
    if (!tuaOnly) {
      await _update("Generating effort estimation...");
      const estTemplate = _loadTemplate("Estimation Template.xlsx") || _loadTemplate("Estimation Template .xlsx");
      if (!estTemplate) throw new Error("Estimation Template.xlsx not found in db/templates");

      const estBuf = await excelGen.writeEstimation(estTemplate, hcaMandatory, s4TechRemediable, spddCount, spauCount, d4AutoCount, d5AutoCount, d9AutoCount);
      await db.run(
        `INSERT INTO atc_Artifacts (id, jobId, role, filename, mimeType, content, size) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [cds.utils.uuid(), jobId, "estimation", `effort_estimation_${_ts()}.xlsx`,
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
         toHex(estBuf), estBuf.length]
      );
    }

    // ── Done ─────────────────────────────────────────────────────────────────
    await db.run(UPDATE("atc.Jobs").set({
      status: "done",
      statusMsg: "Analysis complete — all artifacts ready",
      totalCount, hcaCount, s4Count, spddCount, spauCount,
      fitGapDeltaCount: fitGapDeltaCount || 0,
      chartData: atcData ? JSON.stringify(atcData) : null,
      nsFindings: JSON.stringify(nsFindings),
    }).where({ id: jobId }));

    console.log(`[ATC][${jobId}] Done. mode=${analysisMode} total=${totalCount} hca=${hcaCount} s4=${s4Count} spdd=${spddCount} spau=${spauCount} fitGapDelta=${fitGapDeltaCount || 0} ns.unmaintained=${nsFindings.unmaintained.length} ns.noOwner=${nsFindings.noOwner.length}`);

  } catch (err) {
    console.error(`[ATC][${jobId}] FAILED:`, err);
    try {
      await db.run(UPDATE("atc.Jobs").set({
        status: "failed",
        statusMsg: err.message || "Unknown error",
      }).where({ id: jobId }));
    } catch (_) {}
  }
}

function _ts() {
  return new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 15);
}

function _loadTemplate(filename) {
  const templateDir = path.join(__dirname, "..", "db", "templates");
  const filePath    = path.join(templateDir, filename);
  if (!fs.existsSync(filePath)) {
    try {
      const files = fs.readdirSync(templateDir);
      const match = files.find(f => f.toLowerCase() === filename.toLowerCase());
      if (match) return fs.readFileSync(path.join(templateDir, match));
    } catch (_) {}
    return null;
  }
  return fs.readFileSync(filePath);
}

module.exports = { processJob };
