using { atc } from '../db/schema';

// ── ATC Analysis Service ────────────────────────────────────────────────────
service ATCService @(path: '/api/atc') {

  // ── Jobs ──────────────────────────────────────────────────────────────────
  entity Jobs   as projection on atc.Jobs
    excluding { chartData }  // keep chartData out of OData; use custom action
    actions {
      // Trigger analysis after files are uploaded
      action runAnalysis(useTua: Boolean) returns Jobs;
      // Get full chart data for charts
      function getChartData() returns String;
    };

  // ── File upload/download (REST endpoints added in srv/atc-service.js) ────
  // POST  /api/atc/upload  → multipart upload of one file → returns fileId
  // GET   /api/atc/download/{artifactId} → stream file
  // GET   /api/atc/downloadAll/{jobId}   → ZIP of all artifacts

  // ── Artifacts listing (read-only) ────────────────────────────────────────
  entity Artifacts as projection on atc.Artifacts
    excluding { content };
}
