"use strict";
/**
 * srv/api.js — External REST API for ATC Analysis
 *
 * Exposes a stateless REST interface that any application can call to run
 * the full ATC classification pipeline without touching the CAP OData layer.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * ENDPOINTS
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * POST /api/v1/analyze
 *   Trigger a new analysis job. Accepts multipart/form-data.
 *
 *   Fields:
 *     customer        (string, required)  — customer / project name
 *     migType         (string)            — "conversion" (default) | "upgrade"
 *     analysisMode    (string)            — "atc" (default) | "atc_clone" |
 *                                          "atc_tua" | "atc_tua_clone" | "tua_only"
 *     atcFile         (file)              — ATC extract XLSX (required unless tua_only)
 *     cloneFile       (file, optional)    — Clone extract XLSX
 *     smodilogFile    (file, optional)    — SMODILOG export XLSX (required for TUA)
 *     trnspacetFile   (file, optional)    — TRNSPACET export XLSX (required for TUA)
 *     nsOwnerFile     (file, optional)    — Namespace owner XLSX
 *     wait            (string)            — "true" (default) = wait for completion and
 *                                          return artifacts inline;
 *                                          "false" = return jobId immediately (async mode)
 *
 *   Response 200 (wait=true, default):
 *     {
 *       "jobId":    "uuid",
 *       "status":   "done" | "failed",
 *       "customer": "...",
 *       "counts": {
 *         "total": 1240, "hca": 85, "s4h": 320,
 *         "spdd": 12, "spau": 8, "fitGapDelta": 4
 *       },
 *       "artifacts": [
 *         {
 *           "role":     "atc_result",
 *           "filename": "atc_classified_20240901T120000.xlsx",
 *           "mimeType": "application/vnd.openxmlformats...",
 *           "size":     102400,
 *           "content":  "<base64>"
 *         },
 *         { "role": "pptx",       ... },
 *         { "role": "estimation", ... },
 *         { "role": "tua",        ... }
 *       ]
 *     }
 *
 *   Response 202 (wait=false):
 *     { "jobId": "uuid", "status": "running", "pollUrl": "/api/v1/analyze/uuid" }
 *
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * GET /api/v1/analyze/:jobId
 *   Poll a previously submitted job (use when wait=false).
 *
 *   Response 200:
 *     {
 *       "jobId":    "uuid",
 *       "status":   "running" | "done" | "failed",
 *       "statusMsg": "...",
 *       "customer": "...",
 *       "counts":   { ... },          // only when done
 *       "artifacts": [ ... ]          // only when done (with base64 content)
 *     }
 *
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * GET /api/v1/analyze/:jobId/download/:role
 *   Download a single artifact file by role (atc_result | pptx | estimation | tua).
 *   Returns the file as a binary download (no base64).
 *
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * Authentication (optional):
 *   Set API_KEY env var to require X-API-Key header on all /api/v1/* requests.
 *   If API_KEY is not set, the endpoints are open (suitable for internal apps).
 *
 * ─────────────────────────────────────────────────────────────────────────────
 */

const cds          = require("@sap/cds");
const multer       = require("multer");
const passport     = require("passport");
const xssec        = require("@sap/xssec");
const atcProcessor = require("../lib/atcProcessor");

const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 60 * 1024 * 1024 },  // 60 MB per file
});

// File fields accepted by POST /api/v1/analyze
const FILE_FIELDS = [
  { name: "atcFile",       maxCount: 1 },
  { name: "cloneFile",     maxCount: 1 },
  { name: "smodilogFile",  maxCount: 1 },
  { name: "trnspacetFile", maxCount: 1 },
  { name: "nsOwnerFile",   maxCount: 1 },
];

// Valid analysis modes
const VALID_MODES = new Set(["atc", "atc_clone", "atc_tua", "atc_tua_clone", "tua_only"]);

// ── DB helper ─────────────────────────────────────────────────────────────────
async function _db() {
  if (cds.db) return cds.db;
  return cds.connect.to("db");
}

// ── Buffer helper (handles streams returned by CAP v8 LargeBinary) ────────────
async function _toBuffer(val) {
  if (!val) return Buffer.alloc(0);
  if (Buffer.isBuffer(val)) return val;
  if (typeof val === "string") return Buffer.from(val, "hex");
  if (val instanceof Uint8Array) return Buffer.from(val);
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

// ── Auth middleware ────────────────────────────────────────────────────────────
// Priority: 1) XSUAA Bearer token  2) API key (legacy fallback)
//
// On CF the VCAP_SERVICES.xsuaa credentials are picked up automatically.
// Locally (no XSUAA binding) the API_KEY env var is the only guard.

function _buildXsuaaStrategy() {
  try {
    const vcap  = JSON.parse(process.env.VCAP_SERVICES || "{}");
    const creds = vcap.xsuaa?.[0]?.credentials;
    if (!creds) return null;
    // @sap/xssec v4 uses XssecPassportStrategy (JWTStrategy was removed in v4)
    return new xssec.XssecPassportStrategy(new xssec.XsuaaService(creds));
  } catch (e) {
    console.warn("[API] XSUAA strategy init failed:", e.message);
    return null;
  }
}

const _xsuaaStrategy = _buildXsuaaStrategy();
if (_xsuaaStrategy) {
  passport.use("JWT", _xsuaaStrategy);
  console.log("[API] XSUAA JWT auth enabled");
} else {
  console.log("[API] No XSUAA binding found — falling back to API key auth");
}

function _authGuard(req, res, next) {
  // ── Try XSUAA Bearer token first ──────────────────────────────────────────
  if (_xsuaaStrategy && req.headers.authorization?.startsWith("Bearer ")) {
    return passport.authenticate("JWT", { session: false }, (err, token) => {
      if (err || !token) {
        return res.status(401).json({ error: "Invalid or expired Bearer token" });
      }
      // token.scopes is an array of granted scopes (xssec v4)
      const scopes = token.scopes || [];
      const hasScope = scopes.some(s => s.endsWith(".Analyze") || s.endsWith(".User"));
      if (!hasScope) {
        return res.status(403).json({ error: "Insufficient scope — requires Analyze or User scope" });
      }
      req.user = token;
      return next();
    })(req, res, next);
  }

  // ── Fallback: API key ──────────────────────────────────────────────────────
  const expectedKey = process.env.API_KEY;
  if (expectedKey) {
    const provided = req.headers["x-api-key"];
    if (!provided || provided !== expectedKey) {
      return res.status(401).json({ error: "Missing or invalid X-API-Key header" });
    }
    return next();
  }

  // No auth configured at all — open (dev/internal use)
  next();
}

// ── Build artifact list from DB (with base64 content) ─────────────────────────
async function _loadArtifacts(db, jobId) {
  const arts = await db.run(
    SELECT.from("atc.Artifacts")
      .columns("id", "role", "filename", "mimeType", "size", "content")
      .where({ jobId })
  );
  const result = [];
  for (const art of arts) {
    const buf = await _toBuffer(art.content);
    result.push({
      role:     art.role,
      filename: art.filename,
      mimeType: art.mimeType,
      size:     art.size,
      content:  buf.toString("base64"),
    });
  }
  return result;
}

// ── Build counts object from job record ───────────────────────────────────────
function _counts(job) {
  return {
    total:       job.totalCount       || 0,
    hca:         job.hcaCount         || 0,
    s4h:         job.s4Count          || 0,
    spdd:        job.spddCount        || 0,
    spau:        job.spauCount        || 0,
    fitGapDelta: job.fitGapDeltaCount || 0,
  };
}

// ── Register routes on an Express app instance ────────────────────────────────
function register(app) {
  // Initialize passport (required even when using authenticate() directly)
  app.use(passport.initialize());

  // ── POST /api/v1/analyze ──────────────────────────────────────────────────
  app.post(
    "/api/v1/analyze",
    _authGuard,
    upload.fields(FILE_FIELDS),
    async (req, res) => {
      try {
        const { customer, migType, analysisMode, wait } = req.body || {};
        const files = req.files || {};

        // ── Validate required inputs ────────────────────────────────────────
        if (!customer) {
          return res.status(400).json({ error: "customer field is required" });
        }

        const mode = VALID_MODES.has(analysisMode) ? analysisMode : "atc";
        const tuaOnly = mode === "tua_only";

        if (!tuaOnly && !files.atcFile?.[0]) {
          return res.status(400).json({ error: "atcFile is required (unless analysisMode=tua_only)" });
        }

        if ((mode === "atc_tua" || mode === "atc_tua_clone" || tuaOnly) &&
            (!files.smodilogFile?.[0] || !files.trnspacetFile?.[0])) {
          return res.status(400).json({
            error: "smodilogFile and trnspacetFile are required for TUA analysis modes",
          });
        }

        const db    = await _db();
        const jobId = cds.utils.uuid();

        // ── Store uploaded files in DB ──────────────────────────────────────
        const _storeFile = async (fieldName, role) => {
          const f = files[fieldName]?.[0];
          if (!f) return null;
          const fileId = cds.utils.uuid();
          await db.run(INSERT.into("atc.Files").entries({
            id:       fileId,
            jobId,
            role,
            filename: f.originalname,
            mimeType: f.mimetype,
            content:  f.buffer,
            size:     f.size,
          }));
          return fileId;
        };

        const atcFileId    = await _storeFile("atcFile",       "atc");
        const cloneFileId  = await _storeFile("cloneFile",     "clone");
        const smodilogId   = await _storeFile("smodilogFile",  "smodilog");
        const trnspacetId  = await _storeFile("trnspacetFile", "trnspacet");
        const nsOwnerId    = await _storeFile("nsOwnerFile",   "ns_owner");

        // ── Create job record ──────────────────────────────────────────────
        await db.run(INSERT.into("atc.Jobs").entries({
          id:           jobId,
          customer,
          status:       "running",
          statusMsg:    "Analysis started via REST API",
          useAppLogic:  true,  // always use in-process classification (no I10 dependency)
          migType:      migType === "upgrade" ? "upgrade" : "conversion",
          region:       "NA",
          analysisMode: mode,
          atcFileId,
          cloneFileId,
          smodilogId,
          trnspacetId,
          nsOwnerId,
        }));

        const waitForResult = wait !== "false";  // default: wait=true

        if (!waitForResult) {
          // ── Async mode: fire and return jobId immediately ─────────────────
          setImmediate(() => atcProcessor.processJob(jobId).catch(console.error));
          return res.status(202).json({
            jobId,
            status:   "running",
            pollUrl:  `/api/v1/analyze/${jobId}`,
            message:  "Job started. Poll pollUrl for status and results.",
          });
        }

        // ── Sync mode (default): wait for completion ───────────────────────
        await atcProcessor.processJob(jobId);

        // Reload job for final counts
        const [job] = await db.run(SELECT.from("atc.Jobs").where({ id: jobId }));
        if (!job) return res.status(500).json({ error: "Job record not found after processing" });

        if (job.status === "failed") {
          return res.status(500).json({
            jobId,
            status:    "failed",
            error:     job.statusMsg || "Analysis failed",
            customer,
          });
        }

        const artifacts = await _loadArtifacts(db, jobId);

        return res.json({
          jobId,
          status:   "done",
          customer,
          migType:  job.migType,
          analysisMode: job.analysisMode,
          counts:   _counts(job),
          artifacts,
        });

      } catch (err) {
        console.error("[API /api/v1/analyze POST]", err);
        return res.status(500).json({ error: err.message });
      }
    }
  );

  // ── GET /api/v1/analyze/:jobId — poll job status ─────────────────────────
  app.get("/api/v1/analyze/:jobId", _authGuard, async (req, res) => {
    try {
      const db = await _db();
      const [job] = await db.run(SELECT.from("atc.Jobs").where({ id: req.params.jobId }));
      if (!job) return res.status(404).json({ error: "Job not found" });

      const response = {
        jobId:        job.id,
        status:       job.status,
        statusMsg:    job.statusMsg,
        customer:     job.customer,
        migType:      job.migType,
        analysisMode: job.analysisMode,
      };

      if (job.status === "done") {
        response.counts    = _counts(job);
        response.artifacts = await _loadArtifacts(db, job.id);
      }

      if (job.status === "failed") {
        response.error = job.statusMsg;
      }

      return res.json(response);
    } catch (err) {
      console.error("[API /api/v1/analyze GET]", err);
      return res.status(500).json({ error: err.message });
    }
  });

  // ── GET /api/v1/analyze/:jobId/download/:role — stream artifact file ──────
  app.get("/api/v1/analyze/:jobId/download/:role", _authGuard, async (req, res) => {
    try {
      const db = await _db();
      const { jobId, role } = req.params;

      // Verify job exists and is done
      const [job] = await db.run(SELECT.from("atc.Jobs").where({ id: jobId }));
      if (!job) return res.status(404).json({ error: "Job not found" });
      if (job.status !== "done") {
        return res.status(409).json({ error: `Job status is "${job.status}" — artifacts only available when done` });
      }

      // role may have chunk suffix (e.g. atc_result_chunk_1) — use prefix match
      const arts = await db.run(
        SELECT.from("atc.Artifacts")
          .columns("id", "filename", "mimeType", "content")
          .where({ jobId })
      );
      const art = arts.find(a => a.role === role || a.role.startsWith(role + "_chunk"));
      if (!art) {
        return res.status(404).json({
          error: `Artifact "${role}" not found`,
          availableRoles: arts.map(a => a.role),
        });
      }

      const buf = await _toBuffer(art.content);
      res.setHeader("Content-Disposition", `attachment; filename="${art.filename}"`);
      res.setHeader("Content-Type", art.mimeType || "application/octet-stream");
      res.setHeader("Content-Length", buf.length);
      return res.send(buf);
    } catch (err) {
      console.error("[API /api/v1/analyze/download]", err);
      return res.status(500).json({ error: err.message });
    }
  });

  // ── GET /api/v1/health — liveness check ──────────────────────────────────
  app.get("/api/v1/health", (req, res) => {
    res.json({ status: "ok", version: "1.0.0" });
  });

  console.log("[API] External REST API registered: POST/GET /api/v1/analyze, GET /api/v1/health");
}

module.exports = { register };
