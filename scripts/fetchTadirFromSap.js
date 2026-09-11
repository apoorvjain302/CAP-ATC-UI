"use strict";
/**
 * fetchTadirFromSap.js
 *
 * Builds a growing TADIR/TFDIR cache database by querying only the objects
 * actually referenced in the current ATC Excel file — then merges results
 * into the persistent cache files so they grow richer over time.
 *
 * Strategy:
 *   1. Read the ATC Excel (auto-detect latest in Downloads, or use CLI arg).
 *   2. Extract unique (refObjType, refObjName) pairs for TADIR types
 *      and unique funcnames for TFDIR.
 *   3. Query SAP in batches of 100 using IN(...) clauses.
 *   4. Write results into lib/tadirData.json and lib/tfdirData.json,
 *      MERGING with any existing entries (so cache grows across runs).
 *
 * The output files are true positive sets — entries that EXIST in S/4H.
 * The absence of an entry means "not found in this or any prior run".
 *
 * Run: node scripts/fetchTadirFromSap.js [path-to-atc.xlsx]
 */

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const fs   = require("fs");
const path = require("path");
const http  = require("http");
const https = require("https");
const XLSX  = require("xlsx");

// ── SAP connection ─────────────────────────────────────────────────────────────
const SAP_URL  = "http://10.237.101.128:50000";
const SAP_USER = "I755599";
const SAP_PASS = "Lion@9810821543";
const AUTH     = "Basic " + Buffer.from(`${SAP_USER}:${SAP_PASS}`).toString("base64");

// Object types to look up from TADIR
const TADIR_OBJECT_TYPES = new Set(["DTEL", "DOMA", "TTYP", "INTF", "CLAS", "TABL"]);

// Batch size for IN() queries
// ADT wraps the SQL into ABAP: adds ~60 chars overhead. ABAP text literal limit = 255.
// Safe max SQL length ≈ 195 chars. Dynamically split by SQL length rather than item count.
const MAX_SQL_CHARS = 190;

// ── HTTP helper ────────────────────────────────────────────────────────────────
function httpRequest(options, body) {
  return new Promise((resolve, reject) => {
    const baseUrl = new URL(SAP_URL);
    const lib     = baseUrl.protocol === "https:" ? https : http;
    const req = lib.request({
      hostname:           baseUrl.hostname,
      port:               parseInt(baseUrl.port) || (baseUrl.protocol === "https:" ? 443 : 80),
      path:               options.path,
      method:             options.method || "GET",
      headers:            options.headers || {},
      rejectUnauthorized: false,
    }, res => {
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => resolve({
        status:  res.statusCode,
        headers: res.headers,
        body:    Buffer.concat(chunks).toString("utf8"),
      }));
    });
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getCsrfAndCookies() {
  const res = await httpRequest({
    path:    "/sap/bc/adt/datapreview/freestyle?rowNumber=1&dataPreviewId=CSRF_INIT",
    method:  "GET",
    headers: { Authorization: AUTH, "X-CSRF-Token": "Fetch" },
  });
  const csrf    = res.headers["x-csrf-token"] || "";
  const cookies = (res.headers["set-cookie"] || []).map(c => c.split(";")[0]).join("; ");
  if (!csrf) throw new Error("CSRF token fetch failed — check SAP connection");
  return { csrf, cookies };
}

function parseAdtXml(body) {
  // ADT XML is column-major: each <dataPreview:columns> block contains metadata + all values for that column.
  const columns = [];
  const colBlockRe = /<dataPreview:columns>([\s\S]*?)<\/dataPreview:columns>/g;
  let colBlock;
  while ((colBlock = colBlockRe.exec(body)) !== null) {
    const nameMatch = /dataPreview:name="([^"]+)"/.exec(colBlock[1]);
    if (!nameMatch) continue;
    const colName = nameMatch[1];
    // Use negative lookahead to exclude <dataPreview:dataSet> from matching
    const vals = [...colBlock[1].matchAll(/<dataPreview:data(?![A-Za-z])[^>]*>([\s\S]*?)<\/dataPreview:data>/g)].map(m => m[1].trim());
    columns.push({ name: colName, vals });
  }
  if (!columns.length) return [];
  const numRows = columns[0].vals.length;
  const rows = [];
  for (let r = 0; r < numRows; r++) {
    const row = {};
    for (const col of columns) row[col.name] = col.vals[r] || "";
    rows.push(row);
  }
  return rows;
}

async function runSql(sql, csrf, cookies, maxRows = BATCH_SIZE + 10) {
  const reqPath = `/sap/bc/adt/datapreview/freestyle?rowNumber=${maxRows}&dataPreviewId=FETCH`;
  const res = await httpRequest({
    path:    reqPath,
    method:  "POST",
    headers: {
      Authorization:  AUTH,
      "X-CSRF-Token": csrf,
      "Content-Type": "application/vnd.sap.adt.datapreview.query.v1+txt; charset=utf-8",
      Accept:         "application/vnd.sap.adt.datapreview.table.v1+xml",
      Cookie:         cookies,
    },
  }, sql);
  if (res.status < 200 || res.status >= 300) {
    throw new Error(`SQL failed [${res.status}]: ${res.body.slice(0, 300)}`);
  }
  return parseAdtXml(res.body);
}

// ── Read ATC Excel and extract referenced objects ───────────────────────────
function extractAtcObjects(atcPath) {
  const wb = XLSX.readFile(atcPath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });

  // Find header row (first row containing "Check Title")
  let headerIdx = -1;
  for (let i = 0; i < Math.min(rawRows.length, 5); i++) {
    if (rawRows[i].some(v => (v||"").toString().toLowerCase().includes("check title"))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) throw new Error("Could not find header row in ATC file");
  const headers = rawRows[headerIdx];
  const rows    = rawRows.slice(headerIdx + 1).map(r => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = r[i] || ""; });
    return obj;
  });

  // Find ref type / ref name column
  const refTypeKey  = headers.find(h => /ref.*object.*type|refobjecttype/i.test(h)) || "Ref. Object Type";
  const refNameKey  = headers.find(h => /referenced.*object|ref.*object.*name/i.test(h)) || "Referenced Object";
  const objTypeKey  = headers.find(h => /^obj\.$/i.test(h.trim()) || /^object.*type$/i.test(h.trim())) || "Obj.";
  const objNameKey  = headers.find(h => /^object.*name$/i.test(h.trim())) || "Object name";

  const tadirObjects = {}; // type → Set<name>
  const tfdirNames   = new Set();

  for (const row of rows) {
    const refType = (row[refTypeKey] || "").toString().trim().toUpperCase();
    const refName = (row[refNameKey] || "").toString().trim().toUpperCase();
    const objType = (row[objTypeKey] || "").toString().trim().toUpperCase();

    if (refType && refName) {
      if (refType === "FUNC") {
        tfdirNames.add(refName);
      } else if (TADIR_OBJECT_TYPES.has(refType)) {
        if (!tadirObjects[refType]) tadirObjects[refType] = new Set();
        tadirObjects[refType].add(refName);
      }
    }

    // Also check main object type for FUNC lookups
    if ((objType === "FUNC" || objType === "FUGR") && refName) {
      tfdirNames.add(refName);
    } else if (TADIR_OBJECT_TYPES.has(objType)) {
      const name = refName || (row[objNameKey] || "").toString().trim().toUpperCase();
      if (name) {
        if (!tadirObjects[objType]) tadirObjects[objType] = new Set();
        tadirObjects[objType].add(name);
      }
    }
  }

  return { tadirObjects, tfdirNames };
}

// ── Batch IN query ──────────────────────────────────────────────────────────
// SAP ADT has a ~255-char limit on the SQL text literal it wraps our query into.
// Split queries so each SQL stays under MAX_SQL_CHARS characters.
async function batchQuery(names, sqlPrefix, csrf, cookies) {
  const nameArr  = [...names];
  const found    = new Set();
  let processed  = 0;
  let batchIdx   = 0;

  while (batchIdx < nameArr.length) {
    // Build a batch that fits within MAX_SQL_CHARS
    const batch   = [];
    let inList    = "";
    while (batchIdx < nameArr.length) {
      const n        = nameArr[batchIdx];
      const escaped  = n.replace(/'/g, "''");
      const entry    = (inList ? ", " : "") + `'${escaped}'`;
      const testSql  = `${sqlPrefix} IN (${inList}${entry})`;
      if (testSql.length > MAX_SQL_CHARS && batch.length > 0) break; // would exceed limit — flush
      inList += entry;
      batch.push(n);
      batchIdx++;
    }
    if (!batch.length) { batchIdx++; continue; } // single name too long — skip it

    const sql = `${sqlPrefix} IN (${inList})`;
    try {
      const rows = await runSql(sql, csrf, cookies, batch.length + 5);
      for (const row of rows) {
        const val = row.OBJ_NAME || row.FUNCNAME || "";
        if (val) found.add(val.trim().toUpperCase());
      }
      processed += batch.length;
      process.stdout.write(`\r  ${processed}/${nameArr.length} queried, ${found.size} found`);
    } catch (e) {
      process.stdout.write(`\n  [ERR at ${processed}: ${e.message.slice(0, 80)}]`);
      processed += batch.length;
    }
  }
  process.stdout.write("\n");
  return found;
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  // Determine ATC file path
  let atcPath = process.argv[2];
  if (!atcPath) {
    // Auto-detect latest ATC file in Downloads
    const dl = path.join(process.env.USERPROFILE || process.env.HOME, "Downloads");
    const files = fs.readdirSync(dl)
      .filter(f => /atc.*\.xlsx$/i.test(f))
      .map(f => ({ f, t: fs.statSync(path.join(dl, f)).mtimeMs }))
      .sort((a, b) => b.t - a.t);
    if (!files.length) throw new Error("No ATC xlsx found in Downloads — pass path as argument");
    atcPath = path.join(dl, files[0].f);
  }
  console.log("=== TADIR/TFDIR cache build ===");
  console.log("ATC file:", atcPath);

  // 1. Extract referenced objects from ATC file
  console.log("\nReading ATC file...");
  const { tadirObjects, tfdirNames } = extractAtcObjects(atcPath);
  let tadirCount = 0;
  for (const [t, names] of Object.entries(tadirObjects)) {
    console.log(`  TADIR ${t}: ${names.size}`);
    tadirCount += names.size;
  }
  console.log(`  TFDIR FUNC: ${tfdirNames.size}`);
  console.log(`  Total lookups: ${tadirCount + tfdirNames.size}`);

  // 2. Load existing cache
  const libDir   = path.join(__dirname, "..", "lib");
  const tadirOut = path.join(libDir, "tadirData.json");
  const tfdirOut = path.join(libDir, "tfdirData.json");

  const existingTadir = new Set(fs.existsSync(tadirOut) ? JSON.parse(fs.readFileSync(tadirOut, "utf8")) : []);
  const existingTfdir = new Set(fs.existsSync(tfdirOut) ? JSON.parse(fs.readFileSync(tfdirOut, "utf8")) : []);
  console.log(`\nExisting cache: ${existingTadir.size} TADIR, ${existingTfdir.size} TFDIR entries`);

  // 3. CSRF + cookies
  console.log("\nFetching CSRF token...");
  const { csrf, cookies } = await getCsrfAndCookies();
  console.log("Token: ok");

  // 4. Quick connectivity test
  const testRows = await runSql("SELECT OBJ_NAME FROM TADIR WHERE OBJECT = 'DOMA' AND OBJ_NAME = 'CHAR1'", csrf, cookies, 5);
  if (!testRows.length) {
    console.warn("WARNING: test query returned 0 rows — SAP connectivity issue");
  } else {
    console.log("Connection OK — test row:", testRows[0]);
  }

  // 5. Query TADIR for each type
  const newTadir = new Set(existingTadir);
  const beforeTadir = existingTadir.size;

  console.log("\nQuerying TADIR...");
  for (const [objType, names] of Object.entries(tadirObjects)) {
    // Filter out names already in cache
    const needed = new Set([...names].filter(n => !existingTadir.has(`${objType}|${n}`)));
    if (!needed.size) {
      console.log(`  ${objType}: all ${names.size} already cached`);
      continue;
    }
    process.stdout.write(`  ${objType} (${needed.size} new to check):\n`);
    const found = await batchQuery(
      needed,
      `SELECT OBJ_NAME FROM TADIR WHERE OBJECT = '${objType}' AND OBJ_NAME`,
      csrf, cookies
    );
    for (const n of found) newTadir.add(`${objType}|${n}`);
    console.log(`  → ${found.size}/${needed.size} found`);
  }

  // 6. Query TFDIR
  const newTfdir = new Set(existingTfdir);
  const beforeTfdir = existingTfdir.size;

  const neededTfdir = new Set([...tfdirNames].filter(n => !existingTfdir.has(n)));
  if (!neededTfdir.size) {
    console.log(`\nTFDIR: all ${tfdirNames.size} already cached`);
  } else {
    console.log(`\nQuerying TFDIR (${neededTfdir.size} new to check)...`);
    const found = await batchQuery(
      neededTfdir,
      "SELECT FUNCNAME FROM TFDIR WHERE FUNCNAME",
      csrf, cookies
    );
    for (const n of found) newTfdir.add(n);
    console.log(`  → ${found.size}/${neededTfdir.size} found`);
  }

  // 7. Write updated cache
  const tadirArr = [...newTadir].sort();
  const tfdirArr = [...newTfdir].sort();

  fs.writeFileSync(tadirOut, JSON.stringify(tadirArr, null, 0), "utf8");
  fs.writeFileSync(tfdirOut, JSON.stringify(tfdirArr, null, 0), "utf8");

  console.log(`\nWritten: ${tadirOut}`);
  console.log(`  TADIR: ${tadirArr.length} entries (+${tadirArr.length - beforeTadir} new)`);
  console.log(`Written: ${tfdirOut}`);
  console.log(`  TFDIR: ${tfdirArr.length} entries (+${tfdirArr.length - beforeTfdir} new)`);
  console.log("\nDone.");
}

main().catch(e => { console.error("FATAL:", e.message); process.exit(1); });
