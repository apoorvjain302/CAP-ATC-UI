"use strict";
/**
 * dumpFullTadir.js
 *
 * Full TADIR/TFDIR dump from SAP backend — stores ALL objects of the types
 * used by the ABAP classification program (ZATCASSESSMENT_KK):
 *
 *   TADIR: OBJECT IN ('DTEL','DOMA','TTYP','INTF','CLAS','TABL')
 *   TFDIR: all function modules
 *
 * Strategy: paginate using OBJ_NAME > lastSeen ORDER BY OBJ_NAME
 * (keyset pagination — works reliably with ADT rowNumber limit).
 *
 * Output format (same as fetchTadirFromSap.js — no changes to atcClassify.js needed):
 *   tadirData.json — sorted array of "TYPE|NAME" strings
 *   tfdirData.json — sorted array of FUNCNAME strings
 *
 * Run: node scripts/dumpFullTadir.js
 * Resume: node scripts/dumpFullTadir.js --resume   (skips already-completed types)
 */

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const fs    = require("fs");
const path  = require("path");
const http  = require("http");
const https = require("https");
const zlib  = require("zlib");

// ── SAP connection ─────────────────────────────────────────────────────────────
const SAP_URL  = "http://10.237.101.128:50000";
const SAP_USER = "I755599";
const SAP_PASS = "Lion@9810821543";
const AUTH     = "Basic " + Buffer.from(`${SAP_USER}:${SAP_PASS}`).toString("base64");

// Types matching ABAP program CASE statement
const TADIR_TYPES = ["DTEL", "DOMA", "TTYP", "INTF", "CLAS", "TABL"];

// Page size — ADT supports up to 9999 rows per request
const PAGE_SIZE = 2000;

// Output files
const LIB_DIR    = path.join(__dirname, "..", "lib");
const TADIR_OUT  = path.join(LIB_DIR, "tadirData.json");
const TADIR_GZ   = path.join(LIB_DIR, "tadirData.gz");
const TFDIR_OUT  = path.join(LIB_DIR, "tfdirData.json");
const TFDIR_GZ   = path.join(LIB_DIR, "tfdirData.gz");

// Progress file — allows resume
const PROGRESS_FILE = path.join(__dirname, "dumpProgress.json");

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
    path:    "/sap/bc/adt/datapreview/freestyle?rowNumber=1",
    method:  "GET",
    headers: { Authorization: AUTH, "X-CSRF-Token": "Fetch" },
  });
  const csrf    = res.headers["x-csrf-token"] || "";
  const cookies = (res.headers["set-cookie"] || []).map(c => c.split(";")[0]).join("; ");
  if (!csrf) throw new Error("CSRF token fetch failed");
  return { csrf, cookies };
}

function parseAdtXml(body, colName) {
  // ADT XML is column-major. Extract all <dataPreview:data> values for the named column.
  const colBlockRe = /<dataPreview:columns>([\s\S]*?)<\/dataPreview:columns>/g;
  let colBlock;
  while ((colBlock = colBlockRe.exec(body)) !== null) {
    const nameMatch = /dataPreview:name="([^"]+)"/.exec(colBlock[1]);
    if (!nameMatch || nameMatch[1] !== colName) continue;
    const vals = [...colBlock[1].matchAll(/<dataPreview:data(?![A-Za-z])[^>]*>([\s\S]*?)<\/dataPreview:data>/g)]
      .map(m => m[1].trim()).filter(Boolean);
    return vals;
  }
  return [];
}

async function runSql(sql, csrf, cookies) {
  const res = await httpRequest({
    path:    `/sap/bc/adt/datapreview/freestyle?rowNumber=${PAGE_SIZE}`,
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
    throw new Error(`SQL [${res.status}]: ${res.body.slice(0, 200)}`);
  }
  return res.body;
}

// ── Fetch all names for one TADIR OBJECT type via keyset pagination ────────────
async function fetchAllForType(objType, csrf, cookies, existingNames) {
  const names = new Set(existingNames || []);
  let lastSeen = "";
  let page = 0;
  let totalFetched = 0;

  process.stdout.write(`  ${objType}: fetching`);

  while (true) {
    const where = lastSeen
      ? `WHERE OBJECT = '${objType}' AND OBJ_NAME > '${lastSeen.replace(/'/g, "''")}' ORDER BY OBJ_NAME`
      : `WHERE OBJECT = '${objType}' ORDER BY OBJ_NAME`;
    const sql = `SELECT OBJ_NAME FROM TADIR ${where}`;

    const xml  = await runSql(sql, csrf, cookies);
    const vals = parseAdtXml(xml, "OBJ_NAME");

    for (const v of vals) names.add(v);
    totalFetched += vals.length;
    page++;

    process.stdout.write(` ${totalFetched}`);

    if (vals.length < PAGE_SIZE) break; // last page
    lastSeen = vals[vals.length - 1];

    // Refresh CSRF token every 50 pages to avoid session expiry
    if (page % 50 === 0) {
      const refreshed = await getCsrfAndCookies();
      csrf    = refreshed.csrf;
      cookies = refreshed.cookies;
    }
  }

  process.stdout.write(`  → ${names.size} total\n`);
  return { names, csrf, cookies };
}

// ── Fetch all TFDIR FUNCNAME via keyset pagination ─────────────────────────────
async function fetchAllTfdir(csrf, cookies, existingNames) {
  const names = new Set(existingNames || []);
  let lastSeen = "";
  let page = 0;
  let totalFetched = 0;

  process.stdout.write(`  TFDIR: fetching`);

  while (true) {
    const where = lastSeen
      ? `WHERE FUNCNAME > '${lastSeen.replace(/'/g, "''")}' ORDER BY FUNCNAME`
      : `ORDER BY FUNCNAME`;
    const sql = `SELECT FUNCNAME FROM TFDIR ${where}`;

    const xml  = await runSql(sql, csrf, cookies);
    const vals = parseAdtXml(xml, "FUNCNAME");

    for (const v of vals) names.add(v);
    totalFetched += vals.length;
    page++;

    process.stdout.write(` ${totalFetched}`);

    if (vals.length < PAGE_SIZE) break;
    lastSeen = vals[vals.length - 1];

    if (page % 50 === 0) {
      const refreshed = await getCsrfAndCookies();
      csrf    = refreshed.csrf;
      cookies = refreshed.cookies;
    }
  }

  process.stdout.write(`  → ${names.size} total\n`);
  return { names, csrf, cookies };
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  const resume = process.argv.includes("--resume");

  console.log("=== Full TADIR/TFDIR dump ===");
  console.log(`SAP: ${SAP_URL}`);
  console.log(`Page size: ${PAGE_SIZE} rows/request`);
  console.log(`Mode: ${resume ? "RESUME (skipping completed types)" : "FULL"}\n`);

  // Load progress state
  const progress = resume && fs.existsSync(PROGRESS_FILE)
    ? JSON.parse(fs.readFileSync(PROGRESS_FILE, "utf8"))
    : { tadir: {}, tfdirDone: false };

  // Load existing cache (merge, don't overwrite)
  const tadirAll = {};  // type -> Set<name>
  for (const t of TADIR_TYPES) {
    tadirAll[t] = new Set(progress.tadir[t] || []);
  }
  const tfdirAll = new Set(progress.tfdirDone ? (JSON.parse(fs.readFileSync(TFDIR_OUT, "utf8"))) : []);

  // CSRF
  console.log("Fetching CSRF token...");
  let { csrf, cookies } = await getCsrfAndCookies();
  console.log("OK\n");

  // Dump TADIR per type
  console.log("Dumping TADIR...");
  for (const objType of TADIR_TYPES) {
    if (resume && progress.tadir[objType] && progress.tadir[objType].length > 0) {
      console.log(`  ${objType}: already done (${progress.tadir[objType].length} entries) — skipping`);
      continue;
    }
    const result = await fetchAllForType(objType, csrf, cookies, []);
    csrf    = result.csrf;
    cookies = result.cookies;
    tadirAll[objType] = result.names;

    // Save progress after each type
    progress.tadir[objType] = [...result.names].sort();
    fs.writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 2));

    // Write combined tadirData.json + tadirData.gz
    const combined = [];
    for (const t of TADIR_TYPES) {
      for (const n of (progress.tadir[t] || [])) combined.push(`${t}|${n}`);
    }
    combined.sort();
    const combinedJson = JSON.stringify(combined);
    fs.writeFileSync(TADIR_OUT, combinedJson);
    fs.writeFileSync(TADIR_GZ, zlib.gzipSync(combinedJson, { level: 9 }));
    console.log(`  → tadirData.json updated: ${combined.length} entries (gz: ${(fs.statSync(TADIR_GZ).size/1024/1024).toFixed(1)} MB)`);
  }

  // Dump TFDIR
  if (resume && progress.tfdirDone) {
    console.log(`\nTFDIR: already done — skipping`);
  } else {
    console.log("\nDumping TFDIR...");
    const result = await fetchAllTfdir(csrf, cookies, []);
    csrf    = result.csrf;
    cookies = result.cookies;

    const tfdirArr  = [...result.names].sort();
    const tfdirJson = JSON.stringify(tfdirArr);
    fs.writeFileSync(TFDIR_OUT, tfdirJson);
    fs.writeFileSync(TFDIR_GZ, zlib.gzipSync(tfdirJson, { level: 9 }));
    progress.tfdirDone = true;
    fs.writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 2));
    console.log(`  → tfdirData.json updated: ${tfdirArr.length} entries (gz: ${(fs.statSync(TFDIR_GZ).size/1024/1024).toFixed(1)} MB)`);
  }

  // Final summary
  const tadirFinal = JSON.parse(fs.readFileSync(TADIR_OUT, "utf8"));
  const tfdirFinal = JSON.parse(fs.readFileSync(TFDIR_OUT, "utf8"));
  console.log("\n=== Done ===");
  console.log(`tadirData.json: ${tadirFinal.length} entries | tadirData.gz: ${(fs.statSync(TADIR_GZ).size/1024/1024).toFixed(1)} MB`);
  console.log(`tfdirData.json: ${tfdirFinal.length} entries | tfdirData.gz: ${(fs.statSync(TFDIR_GZ).size/1024/1024).toFixed(1)} MB`);

  // Clean up progress file
  if (fs.existsSync(PROGRESS_FILE)) fs.unlinkSync(PROGRESS_FILE);
  console.log("Progress file cleaned up.");
}

main().catch(e => { console.error("FATAL:", e.message); process.exit(1); });
