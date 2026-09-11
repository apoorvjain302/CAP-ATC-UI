"use strict";
/**
 * buildDumps.js — one-off script to generate lib/tadirData.json and lib/tfdirData.json
 * Run: node scripts/buildDumps.js
 * Requires SAP_URL / SAP_USERNAME / SAP_PASSWORD in .env
 */
const path = require("path");
const fs   = require("fs");

try { require("dotenv").config({ path: path.join(__dirname, "..", ".env"), override: false }); } catch (_) {}

const https = require("https");
const fetch = (...args) => import("node-fetch").then(m => m.default(...args));
const agent = new https.Agent({ rejectUnauthorized: false });

const SAP_URL  = (process.env.SAP_URL  || "").replace(/\/$/, "");
const CLIENT   = process.env.SAP_CLIENT   || "100";
const USERNAME = process.env.SAP_USERNAME || "";
const PASSWORD = process.env.SAP_PASSWORD || "";
const AUTH     = "Basic " + Buffer.from(`${USERNAME}:${PASSWORD}`).toString("base64");

async function query(sql, rowNumber = 500) {
  const url = `${SAP_URL}/sap/bc/adt/datapreview/freestyle?sap-client=${CLIENT}&rowNumber=${rowNumber}`;
  const resp = await fetch(url, {
    method: "POST",
    headers: { Authorization: AUTH, "Content-Type": "text/plain", Accept: "application/json" },
    body: sql,
    agent,
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}: ${await resp.text().catch(()=>"")}`);
  const data = await resp.json();
  // Parse column-oriented ADT response
  const cols = data.columns || [];
  if (!cols.length) return data.rows || [];
  const count = (cols[0].keyValue || cols[0].values || []).length;
  const rows = [];
  for (let r = 0; r < count; r++) {
    const row = {};
    for (const col of cols) {
      const vals = col.keyValue || col.values || [];
      row[col.name || col.columnName || ""] = vals[r] || "";
    }
    rows.push(row);
  }
  return rows;
}

async function main() {
  console.log(`Connecting to ${SAP_URL} as ${USERNAME}`);

  // ── TFDIR ──────────────────────────────────────────────────────────────────
  console.log("Fetching TFDIR Z%...");
  const tfZ = await query("SELECT FUNCNAME FROM TFDIR WHERE FUNCNAME LIKE 'Z%'", 500);
  console.log(`TFDIR Z%: ${tfZ.length} rows`);

  console.log("Fetching TFDIR Y%...");
  const tfY = await query("SELECT FUNCNAME FROM TFDIR WHERE FUNCNAME LIKE 'Y%'", 500);
  console.log(`TFDIR Y%: ${tfY.length} rows`);

  const tfdirSet = new Set();
  for (const r of [...tfZ, ...tfY]) {
    const fn = (r.FUNCNAME || r.funcname || "").trim();
    if (fn) tfdirSet.add(fn);
  }
  const tfdirArr = [...tfdirSet].sort();
  fs.writeFileSync(path.join(__dirname, "..", "lib", "tfdirData.json"), JSON.stringify(tfdirArr, null, 0));
  console.log(`Written tfdirData.json: ${tfdirArr.length} entries`);

  // ── TADIR ──────────────────────────────────────────────────────────────────
  const OBJECT_TYPES = ["DTEL","DOMA","CLAS","TABL","DDLS","PROG","VIEW","TTYP","FUGR","INTF","MSAG","ENHO","ENHS","XSLT"];
  const tadirSet = new Set();

  for (const obj of OBJECT_TYPES) {
    for (const ns of ["Z%", "Y%"]) {
      console.log(`Fetching TADIR OBJECT=${obj} OBJ_NAME LIKE '${ns}'...`);
      try {
        const rows = await query(
          `SELECT OBJECT, OBJ_NAME FROM TADIR WHERE PGMID = 'R3TR' AND OBJECT = '${obj}' AND OBJ_NAME LIKE '${ns}'`,
          500,
        );
        console.log(`  ${rows.length} rows`);
        for (const r of rows) {
          const o = (r.OBJECT || r.object || "").trim();
          const n = (r.OBJ_NAME || r.obj_name || "").trim();
          if (o && n) tadirSet.add(`${o}|${n}`);
        }
      } catch (e) {
        console.warn(`  FAILED: ${e.message}`);
      }
    }
  }

  const tadirArr = [...tadirSet].sort();
  fs.writeFileSync(path.join(__dirname, "..", "lib", "tadirData.json"), JSON.stringify(tadirArr, null, 0));
  console.log(`Written tadirData.json: ${tadirArr.length} entries`);
  console.log("Done.");
}

main().catch(err => { console.error(err); process.exit(1); });
