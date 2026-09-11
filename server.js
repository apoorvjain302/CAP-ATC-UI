"use strict";
/**
 * server.js — Custom CAP server bootstrap
 * Registers custom REST endpoints BEFORE CDS OData routes.
 * All DB ops use cds.db.run() since bootstrap runs outside a request context.
 */

const cds      = require("@sap/cds");
const path     = require("path");
const express  = require("express");
const multer   = require("multer");
const archiver = require("archiver");
const passport = require("passport");
const xssec    = require("@sap/xssec");

try {
  require("dotenv").config({ path: path.join(__dirname, ".env"), override: false });
} catch (_) {}

// ── XSUAA JWT middleware for internal /atc/* routes ───────────────────────────
// Protects the UI-facing endpoints (upload, job create, download).
// Falls through silently if no XSUAA binding is present (local dev).
function _buildAtcAuthMiddleware() {
  try {
    const vcap  = JSON.parse(process.env.VCAP_SERVICES || "{}");
    const creds = vcap.xsuaa?.[0]?.credentials;
    if (!creds) return null;
    passport.use("JWT-internal", new xssec.JWTStrategy(creds));
    console.log("[server] XSUAA JWT auth enabled for /atc/* routes");
    return passport.authenticate("JWT-internal", { session: false });
  } catch (_) {
    return null;
  }
}
const _atcAuth = _buildAtcAuthMiddleware();

const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 60 * 1024 * 1024 },
});

// ── CF instance affinity helpers ──────────────────────────────────────────────
// Computed once at startup from CF env vars.
// X-CF-App-Instance format: "<app-guid>:<instance-index>"
// The browser sends this header on all fetch() calls to force the gorouter to
// route to the specific instance that holds the user's SQLite data.
// For <a> link clicks (downloads), we use the __VCAP_ID__ cookie instead —
// the /atc/pin endpoint plants it by responding from the correct instance.
let _cfAppInstance = null;
let _instanceGuid  = null;
(function _initCfInstance() {
  try {
    const vcapApp = JSON.parse(process.env.VCAP_APPLICATION || "{}");
    const idx     = process.env.CF_INSTANCE_INDEX;
    if (vcapApp.application_id && idx != null) {
      _cfAppInstance = `${vcapApp.application_id}:${idx}`;
    }
    _instanceGuid = process.env.CF_INSTANCE_GUID || null;
  } catch (_) {}
})();

// Helper: get db — waits for cds.db to be set after bootstrap
async function _db() {
  if (cds.db) return cds.db;
  return cds.connect.to("db");
}

// ── External REST API ────────────────────────────────────────────────────────
const externalApi = require("./srv/api");

cds.on("bootstrap", async (app) => {

  // ── Serve SAPUI5 frontend — must be first to beat CAP's own routing ──────
  const appBase = require("fs").existsSync(path.join(__dirname, "app"))
    ? __dirname
    : path.join(__dirname, "..", "..");
  const webappDir = path.join(appBase, "app", "atc-ui", "webapp");
  app.use(express.static(webappDir, { index: "index.html" }));
  app.use(express.json({ limit: "10mb" }));
  app.use(passport.initialize());

  // ── Register external REST API (/api/v1/*) ────────────────────────────────
  externalApi.register(app);

  // ── Pin browser session to this CF instance (no auth — called pre-login) ──
  // Called via fetch() with X-CF-App-Instance header AFTER the first upload
  // (so it is force-routed to the right instance). This response sets __VCAP_ID__
  // which the CF gorouter rewrites to its own affinity token for this instance.
  // The browser then sends that cookie on ALL subsequent requests, including
  // <a> link downloads that cannot carry custom headers.
  app.post("/atc/pin", (req, res) => {
    if (_instanceGuid) {
      // Set __VCAP_ID__ — gorouter will overwrite with its own value for this instance,
      // ensuring the browser is pinned here for all future requests including downloads.
      res.cookie("__VCAP_ID__", _instanceGuid, { path: "/", sameSite: "Lax" });
    }
    res.json({ ok: true, instance: _cfAppInstance });
  });

  // ── Upload a single file ───────────────────────────────────────────────────
  app.post("/atc/upload", ...(_atcAuth ? [_atcAuth] : []), upload.single("file"), async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: "No file received" });
      const { role } = req.body;
      if (!role) return res.status(400).json({ error: "role is required" });

      const db     = await _db();
      const fileId = cds.utils.uuid();

      await db.run(INSERT.into("atc.Files").entries({
        id:       fileId,
        jobId:    null,
        role,
        filename: req.file.originalname,
        mimeType: req.file.mimetype,
        content:  req.file.buffer,
        size:     req.file.size,
      }));

      res.json({
        fileId,
        filename:      req.file.originalname,
        size:          req.file.size,
        cfAppInstance: _cfAppInstance,  // frontend sends this as X-CF-App-Instance on all subsequent requests
      });
    } catch (err) {
      console.error("[upload]", err);
      res.status(500).json({ error: err.message });
    }
  });

  // ── Create a new job ───────────────────────────────────────────────────────
  app.post("/atc/jobs/create", ...(_atcAuth ? [_atcAuth] : []), async (req, res) => {
    try {
      const { customer, atcFileId, cloneFileId, smodilogId, trnspacetId, nsOwnerId, useAppLogic, migType, analysisMode } = req.body;
      if (!customer) return res.status(400).json({ error: "customer is required" });
      // atcFileId required unless mode is tua_only
      if (analysisMode !== "tua_only" && !atcFileId) return res.status(400).json({ error: "atcFileId is required" });

      const db    = await _db();
      const jobId = cds.utils.uuid();

      // ── Cleanup: delete completed/failed jobs older than 1 hour ──────────────
      // Only removes done/failed jobs — never touches uploaded/running jobs
      // (concurrent users mid-analysis are fully protected).
      // Cascades: delete their artifacts and uploaded files too.
      try {
        const cutoff = new Date(Date.now() - 60 * 60 * 1000).toISOString();
        const staleJobs = await db.run(
          SELECT.from("atc.Jobs").columns("id")
            .where({ status: { in: ["done", "failed"] } })
            .where("modifiedAt <=", cutoff)
        );
        for (const j of staleJobs) {
          await db.run(DELETE.from("atc.Artifacts").where({ jobId: j.id }));
          await db.run(DELETE.from("atc.Files").where({ jobId: j.id }));
          await db.run(DELETE.from("atc.Jobs").where({ id: j.id }));
          console.log(`[cleanup] Deleted stale job ${j.id}`);
        }
      } catch (cleanErr) {
        console.warn("[cleanup] Non-fatal cleanup error:", cleanErr.message);
      }

      const validModes = ["atc", "atc_clone", "atc_tua", "atc_tua_clone", "tua_only"];
      const mode = validModes.includes(analysisMode) ? analysisMode : "atc";

      await db.run(INSERT.into("atc.Jobs").entries({
        id:           jobId,
        customer,
        status:       "uploaded",
        statusMsg:    "Files uploaded — ready to run analysis",
        useAppLogic:  useAppLogic === true || useAppLogic === "true" ? true : false,
        migType:      migType === "upgrade" ? "upgrade" : "conversion",
        region:       "NA",
        analysisMode: mode,
        atcFileId:    atcFileId   || null,
        cloneFileId:  cloneFileId || null,
        smodilogId:   smodilogId  || null,
        trnspacetId:  trnspacetId || null,
        nsOwnerId:    nsOwnerId   || null,
      }));

      for (const fid of [atcFileId, cloneFileId, smodilogId, trnspacetId, nsOwnerId]) {
        if (fid) await db.run(UPDATE("atc.Files").set({ jobId }).where({ id: fid }));
      }

      res.json({ jobId });
    } catch (err) {
      console.error("[jobs/create]", err);
      res.status(500).json({ error: err.message });
    }
  });

  // ── Download a single artifact ─────────────────────────────────────────────
  app.get("/atc/download/:artifactId", ...(_atcAuth ? [_atcAuth] : []), async (req, res) => {
    try {
      const db  = await _db();
      const [art] = await db.run(SELECT.from("atc.Artifacts").columns("id","filename","mimeType","content").where({ id: req.params.artifactId }));
      if (!art) return res.status(404).json({ error: "Artifact not found" });

      const buf = await _toBuffer(art.content);
      res.setHeader("Content-Disposition", `attachment; filename="${art.filename}"`);
      res.setHeader("Content-Type", art.mimeType || "application/octet-stream");
      res.send(buf);
    } catch (err) {
      console.error("[download]", err);
      res.status(500).json({ error: err.message });
    }
  });

  // ── Download all artifacts as ZIP ──────────────────────────────────────────
  app.get("/atc/downloadAll/:jobId", ...(_atcAuth ? [_atcAuth] : []), async (req, res) => {
    try {
      const db   = await _db();
      const arts = await db.run(SELECT.from("atc.Artifacts").columns("id","filename","mimeType","content").where({ jobId: req.params.jobId }));
      if (!arts.length) return res.status(404).json({ error: "No artifacts found" });

      const [job]    = await db.run(SELECT.from("atc.Jobs").where({ id: req.params.jobId }));
      const customer = job?.customer ? job.customer.replace(/[^a-z0-9_\-]/gi, "_") : "ATC";

      res.setHeader("Content-Disposition", `attachment; filename="${customer}_ATC_Results.zip"`);
      res.setHeader("Content-Type", "application/zip");

      const archive = archiver("zip", { zlib: { level: 9 } });
      archive.on("error", err => res.status(500).json({ error: err.message }));
      archive.pipe(res);
      for (const art of arts) {
        archive.append(await _toBuffer(art.content), { name: art.filename });
      }
      await archive.finalize();
    } catch (err) {
      console.error("[downloadAll]", err);
      res.status(500).json({ error: err.message });
    }
  });
});

async function _toBuffer(val) {
  if (!val) return Buffer.alloc(0);
  if (Buffer.isBuffer(val)) return val;
  if (typeof val === "string") return Buffer.from(val, "hex");
  if (val instanceof Uint8Array) return Buffer.from(val);
  // CAP v8 returns LargeBinary as a Readable stream
  if (typeof val.pipe === "function" || typeof val.on === "function") {
    return new Promise((resolve, reject) => {
      const chunks = [];
      val.on("data", c => chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c)));
      val.on("end",  () => resolve(Buffer.concat(chunks)));
      val.on("error", reject);
    });
  }
  return Buffer.from(String(val), "hex");
}

module.exports = cds.server;

// Pre-warm TADIR/TFDIR sets after server is ready.
// Loading these during a job (on-demand) causes a memory spike that OOMs
// the 512MB CF container. Loading at startup gives Node time to GC before
// any job request arrives, keeping peak usage well under the limit.
cds.on("served", () => {
  setImmediate(() => {
    try {
      const atcClassify = require("./lib/atcClassify");
      if (typeof atcClassify.preloadRefSets === "function") {
        atcClassify.preloadRefSets();
        console.log("[server] TADIR/TFDIR pre-warm complete.");
      }
    } catch (e) {
      console.warn("[server] TADIR/TFDIR pre-warm failed:", e.message);
    }
  });
});
