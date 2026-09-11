"use strict";
/**
 * sapApi.js — SAP ATC API caller (Node.js port of tool_sap.py::call_sap_api)
 *
 * On BTP CF: reads SAP_URL/credentials from BTP Destination service via VCAP_SERVICES.
 * Locally: reads SAP_URL, SAP_USERNAME, SAP_PASSWORD from environment / .env.
 */

const https    = require("https");
const http     = require("http");
const url      = require("url");
const FormData = require("form-data");
const fetch    = (...args) => import("node-fetch").then(m => m.default(...args));

// Ignore self-signed certs (same as Python's ssl.CERT_NONE)
const _agent = new https.Agent({ rejectUnauthorized: false });

async function _getSapConnection() {
  const vcapRaw = process.env.VCAP_SERVICES;
  if (vcapRaw) {
    return _getCfDestination();
  }
  // Local fallback
  return {
    url:        (process.env.SAP_URL || "").replace(/\/$/, ""),
    client:     process.env.SAP_CLIENT || "100",
    username:   process.env.SAP_USERNAME || "",
    password:   process.env.SAP_PASSWORD || "",
    authHeader: null,
    proxyUrl:   null,
    proxyAuth:  null,
  };
}

async function _getCfDestination() {
  const vcap       = JSON.parse(process.env.VCAP_SERVICES);
  const destSvc    = (vcap.destination || [])[0];
  if (!destSvc) throw new Error("No 'destination' service bound in VCAP_SERVICES");

  const creds      = destSvc.credentials;
  const tokenUrl   = `${creds.url}/oauth/token`;
  const destApi    = creds.uri;
  const destName   = process.env.SAP_DESTINATION || "I10_ATC";

  // 1. Get OAuth token for Destination service
  const tokenResp  = await fetch(tokenUrl, {
    method:  "POST",
    agent:   _agent,
    headers: { "Content-Type": "application/x-www-form-urlencoded",
                "Authorization": "Basic " + Buffer.from(`${creds.clientid}:${creds.clientsecret}`).toString("base64") },
    body:    "grant_type=client_credentials",
  });
  if (!tokenResp.ok) throw new Error(`Destination OAuth failed: ${tokenResp.status}`);
  const { access_token } = await tokenResp.json();

  // 2. Fetch destination config
  const destResp = await fetch(
    `${destApi}/destination-configuration/v1/destinations/${destName}`,
    { headers: { Authorization: `Bearer ${access_token}` }, agent: _agent },
  );
  if (!destResp.ok) throw new Error(`Destination fetch failed: ${destResp.status}`);
  const destData = await destResp.json();

  const destUrl = destData.destinationConfiguration?.URL || "";
  const authTokens = destData.authTokens || [];
  let authHeader = null;
  let username = "", password = "";
  const sccLocationId = destData.destinationConfiguration?.CloudConnectorLocationId || "";

  if (authTokens.length && !authTokens[0].error) {
    authHeader = `${authTokens[0].type} ${authTokens[0].value}`;
  } else {
    const cfg = destData.destinationConfiguration || {};
    username  = cfg.User || cfg.user || "";
    password  = cfg.Password || cfg.password || "";
  }

  // 3. Connectivity proxy
  const connSvc = (vcap.connectivity || [])[0];
  let proxyUrl = null, proxyAuth = null;
  if (connSvc) {
    const c = connSvc.credentials;
    if (c.onpremise_proxy_host && c.onpremise_proxy_port) {
      proxyUrl = `http://${c.onpremise_proxy_host}:${c.onpremise_proxy_port}`;
      try {
        const pTokenResp = await fetch(`${c.url}/oauth/token`, {
          method:  "POST",
          agent:   _agent,
          headers: { "Content-Type": "application/x-www-form-urlencoded",
                      "Authorization": "Basic " + Buffer.from(`${c.clientid}:${c.clientsecret}`).toString("base64") },
          body:    "grant_type=client_credentials",
        });
        if (pTokenResp.ok) {
          const { access_token: pa } = await pTokenResp.json();
          proxyAuth = `Bearer ${pa}`;
        }
      } catch (_) {}
    }
  }

  return {
    url:          destUrl,
    client:       process.env.SAP_CLIENT || "100",
    username,
    password,
    authHeader,
    proxyUrl,
    proxyAuth,
    sccLocationId,
  };
}

/**
 * POST ATC file (+ optional clone file) to SAP ATC API.
 * Returns { xlsBuffer: Buffer, error: string|null }
 */
async function callAtcApi(atcBuffer, atcFilename, cloneBuffer, cloneFilename) {
  let conn;
  try {
    conn = await _getSapConnection();
  } catch (err) {
    return { xlsBuffer: null, error: err.message };
  }

  const { url: sapUrl, client, username, password, authHeader, proxyUrl, proxyAuth, sccLocationId } = conn;
  // Cloud Connector proxy requires HTTP (tunnel already secures the connection)
  const effectiveUrl = proxyUrl ? sapUrl.replace(/^https:/i, "http:") : sapUrl;
  const apiUrl = `${effectiveUrl}/sap/bc/atc_assessment?sap-client=${client}`;

  const boundary = `ATC_CAP_BOUNDARY_${Date.now()}`;

  // Build raw multipart body (same as Python)
  let body = Buffer.alloc(0);

  const part1Header = Buffer.from(
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="file"; filename="${atcFilename}"\r\n` +
    `Content-Type: ${_detectMime(atcBuffer)}\r\n\r\n`
  );
  body = Buffer.concat([body, part1Header, atcBuffer, Buffer.from("\r\n")]);

  if (cloneBuffer && cloneFilename) {
    const part2Header = Buffer.from(
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="clone_file"; filename="${cloneFilename}"\r\n` +
      `Content-Type: ${_detectMime(cloneBuffer)}\r\n\r\n`
    );
    body = Buffer.concat([body, part2Header, cloneBuffer, Buffer.from("\r\n")]);
  }

  body = Buffer.concat([body, Buffer.from(`--${boundary}--\r\n`)]);

  const authValue = authHeader || ("Basic " + Buffer.from(`${username}:${password}`).toString("base64"));

  const headers = {
    "Content-Type":  `multipart/form-data; boundary=${boundary}`,
    "Authorization": authValue,
    "Accept-Encoding": "gzip, deflate",
  };
  if (proxyAuth) headers["Proxy-Authorization"] = proxyAuth;
  if (sccLocationId) headers["SAP-Connectivity-SCC-Location_ID"] = sccLocationId;

  console.log(`[SAP-API] POST ${apiUrl}, file=${atcFilename} (${atcBuffer.length} bytes), clone=${cloneFilename || "none"}`);

  try {
    const fetchOptions = {
      method:  "POST",
      headers,
      body,
      agent:   _agent,
    };

    if (proxyUrl) {
      // Cloud Connector tunnel: keepAlive=false forces a fresh TCP connection every time,
      // preventing stale socket reuse after SCC tunnel idle timeout
      const { HttpProxyAgent } = require("http-proxy-agent");
      fetchOptions.agent = new HttpProxyAgent(proxyUrl, { keepAlive: false });
    }

    let resp, raw, lastErr;
    for (let attempt = 1; attempt <= 3; attempt++) {
      resp = undefined;
      raw  = undefined;

      // Fresh AbortController per attempt — a failed attempt aborts its own signal,
      // which would otherwise cancel subsequent attempts if shared
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 600_000);
      fetchOptions.signal = controller.signal;

      // Re-fetch fresh connection on retries (proxy token / tunnel may have expired)
      if (attempt > 1) {
        try {
          const freshConn = await _getSapConnection();
          if (freshConn.proxyAuth) fetchOptions.headers["Proxy-Authorization"] = freshConn.proxyAuth;
          if (freshConn.proxyUrl) {
            const { HttpProxyAgent } = require("http-proxy-agent");
            fetchOptions.agent = new HttpProxyAgent(freshConn.proxyUrl, { keepAlive: false });
          }
          if (freshConn.authHeader) fetchOptions.headers["Authorization"] = freshConn.authHeader;
          else if (freshConn.username) fetchOptions.headers["Authorization"] = "Basic " + Buffer.from(`${freshConn.username}:${freshConn.password}`).toString("base64");
        } catch (_) { /* use existing creds if refresh fails */ }
      }

      try {
        resp = await fetch(apiUrl, fetchOptions);
        raw  = await resp.text();
        clearTimeout(timer);
        break;
      } catch (err) {
        clearTimeout(timer);
        lastErr = err;
        if (attempt < 3) {
          console.log(`[SAP-API] Attempt ${attempt} failed (${err.message}), retrying in 3s...`);
          await new Promise(r => setTimeout(r, 3000));
        } else {
          console.log(`[SAP-API] Attempt ${attempt} failed (${err.message}), giving up.`);
        }
      }
    }
    if (!resp || raw === undefined) throw lastErr;

    console.log(`[SAP-API] HTTP ${resp.status}, len=${raw.length}`);

    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (_) {
      return { xlsBuffer: null, error: `Non-JSON SAP response: ${raw.slice(0, 500)}` };
    }

    if (parsed.status !== "success" || !parsed.fileContent) {
      return { xlsBuffer: null, error: `SAP API error: ${JSON.stringify(parsed).slice(0, 500)}` };
    }

    const xlsBuffer = Buffer.from(parsed.fileContent, "base64");
    return { xlsBuffer, totalRecords: parsed.totalRecords || 0, error: null };

  } catch (err) {
    return { xlsBuffer: null, error: `Network error: ${err.message}` };
  }
}

/**
 * Fetch TADIR and TFDIR entries for a given list of objects.
 * Used by in-app classification to determine if referenced objects exist in the target system.
 *
 * @param {Array<{object: string, obj_name: string, funcname?: string}>} objectEntries
 * @returns {{ tadirSet: Set<string>, tfdirSet: Set<string>, error: string|null }}
 *   tadirSet: Set of "OBJECT|OBJ_NAME" keys (e.g. "DTEL|ZMY_ELEMENT")
 *   tfdirSet: Set of function module names (FUNCNAME)
 */
async function fetchTadirTfdir(objectEntries) {
  if (!objectEntries || objectEntries.length === 0) {
    return { tadirSet: new Set(), tfdirSet: new Set(), error: null };
  }

  let conn;
  try {
    conn = await _getSapConnection();
  } catch (err) {
    console.warn(`[fetchTadirTfdir] Could not get SAP connection: ${err.message}`);
    return { tadirSet: new Set(), tfdirSet: new Set(), error: err.message };
  }

  const { url: sapUrl, client, username, password, authHeader, proxyUrl, proxyAuth, sccLocationId } = conn;
  const effectiveUrl = proxyUrl ? sapUrl.replace(/^https:/i, "http:") : sapUrl;

  const authValue = authHeader || ("Basic " + Buffer.from(`${username}:${password}`).toString("base64"));
  const baseHeaders = { "Authorization": authValue, "sap-client": client, "Accept": "application/json" };
  if (proxyAuth) baseHeaders["Proxy-Authorization"] = proxyAuth;
  if (sccLocationId) baseHeaders["SAP-Connectivity-SCC-Location_ID"] = sccLocationId;

  let fetchAgent = _agent;
  if (proxyUrl) {
    const { HttpProxyAgent } = require("http-proxy-agent");
    fetchAgent = new HttpProxyAgent(proxyUrl, { keepAlive: false });
  }

  const tadirSet = new Set();
  const tfdirSet = new Set();

  // Deduplicate and split into TADIR entries and TFDIR (funcname) entries
  const tadirEntries = [];
  const funcnames    = new Set();
  const seenTadir    = new Set();

  for (const e of objectEntries) {
    const obj  = (e.object  || "").toUpperCase().trim();
    const name = (e.obj_name || "").toUpperCase().trim();
    const fn   = (e.funcname || "").toUpperCase().trim();

    if (obj && name) {
      const key = `${obj}|${name}`;
      if (!seenTadir.has(key)) { seenTadir.add(key); tadirEntries.push({ object: obj, obj_name: name }); }
    }
    if (fn) funcnames.add(fn);
  }

  // ── TADIR lookup via ADT Data Preview (runs SQL on the SAP system) ──────────
  // We use the /sap/bc/adt/datapreview/freestyle endpoint which accepts a SQL query.
  // Chunk to avoid URL length limits.
  const CHUNK = 50;
  for (let i = 0; i < tadirEntries.length; i += CHUNK) {
    const chunk = tadirEntries.slice(i, i + CHUNK);
    // Build WHERE clause: (OBJECT='X' AND OBJ_NAME='Y') OR ...
    const where = chunk.map(e => `(OBJECT='${e.object.replace(/'/g,"''")}' AND OBJ_NAME='${e.obj_name.replace(/'/g,"''")}' AND PGMID='R3TR')`).join(" OR ");
    const sql   = `SELECT OBJECT,OBJ_NAME FROM TADIR WHERE ${where}`;
    try {
      const resp = await fetch(
        `${effectiveUrl}/sap/bc/adt/datapreview/freestyle?sap-client=${client}&rowNumber=500`,
        { method: "POST", headers: { ...baseHeaders, "Content-Type": "text/plain" }, body: sql, agent: fetchAgent },
      );
      if (resp.ok) {
        const data = await resp.json();
        const rows = data?.columns ? _parseAdtRows(data) : (data?.rows || []);
        for (const row of rows) {
          const obj  = (row.OBJECT  || row.object  || "").toUpperCase().trim();
          const name = (row.OBJ_NAME || row.obj_name || "").toUpperCase().trim();
          if (obj && name) tadirSet.add(`${obj}|${name}`);
        }
      } else {
        console.warn(`[fetchTadirTfdir] TADIR chunk ${i}: HTTP ${resp.status}`);
      }
    } catch (e) {
      console.warn(`[fetchTadirTfdir] TADIR chunk ${i} error: ${e.message}`);
    }
  }

  // ── TFDIR lookup ─────────────────────────────────────────────────────────────
  const fnArr = [...funcnames];
  for (let i = 0; i < fnArr.length; i += CHUNK) {
    const chunk = fnArr.slice(i, i + CHUNK);
    const where = chunk.map(fn => `FUNCNAME='${fn.replace(/'/g,"''")}'`).join(" OR ");
    const sql   = `SELECT FUNCNAME FROM TFDIR WHERE ${where}`;
    try {
      const resp = await fetch(
        `${effectiveUrl}/sap/bc/adt/datapreview/freestyle?sap-client=${client}&rowNumber=500`,
        { method: "POST", headers: { ...baseHeaders, "Content-Type": "text/plain" }, body: sql, agent: fetchAgent },
      );
      if (resp.ok) {
        const data = await resp.json();
        const rows = data?.columns ? _parseAdtRows(data) : (data?.rows || []);
        for (const row of rows) {
          const fn = (row.FUNCNAME || row.funcname || "").toUpperCase().trim();
          if (fn) tfdirSet.add(fn);
        }
      } else {
        console.warn(`[fetchTadirTfdir] TFDIR chunk ${i}: HTTP ${resp.status}`);
      }
    } catch (e) {
      console.warn(`[fetchTadirTfdir] TFDIR chunk ${i} error: ${e.message}`);
    }
  }

  console.log(`[fetchTadirTfdir] tadirSet=${tadirSet.size} tfdirSet=${tfdirSet.size}`);
  return { tadirSet, tfdirSet, error: null };
}

/** Parse ADT data preview column-oriented JSON into row objects */
function _parseAdtRows(data) {
  const cols = data.columns || [];
  if (!cols.length) return [];
  const count = cols[0].keyValue?.length || cols[0].values?.length || 0;
  const rows  = [];
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

function _detectMime(buf) {
  if (!buf || buf.length < 4) return "application/octet-stream";
  const magic = buf.slice(0, 4);
  if (magic[0] === 0x50 && magic[1] === 0x4B) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (magic[0] === 0xD0 && magic[1] === 0xCF) return "application/vnd.ms-excel";
  return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
}

module.exports = { callAtcApi, fetchTadirTfdir };
