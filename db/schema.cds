// ────────────────────────────────────────────────────────────
//  SAP ATC Analysis — CAP Application (BTP CF)
// ────────────────────────────────────────────────────────────

using { managed } from '@sap/cds/common';

namespace atc;

// ── Analysis Job entity (tracks one analysis run) ─────────
entity Jobs : managed {
  key id          : UUID;
      customer    : String(200);
      status      : String(20)  default 'uploaded';  // uploaded|running|done|failed
      statusMsg   : String(2000);
      // Counts from ATC result
      totalCount  : Integer     default 0;
      hcaCount    : Integer     default 0;
      s4Count     : Integer     default 0;
      spddCount   : Integer     default 0;
      spauCount   : Integer     default 0;
      fitGapDeltaCount : Integer default 0;
      // Raw chart data stored as JSON string
      chartData   : LargeString;
      // Processing mode: true = app-logic classification (default), false = I10 API
      // Toggle removed from UI — app logic is always used; set to false to re-enable I10 path.
      useAppLogic : Boolean default true;
      // Migration type: 'conversion' (ECC→S/4HANA) or 'upgrade' (S/4HANA→S/4HANA)
      migType     : String(20) default 'conversion';
      // Region: 'NA' (North America) or 'APAC'
      region      : String(10) default 'NA';
      // Analysis mode: 'atc'|'atc_clone'|'atc_tua'|'atc_tua_clone'|'tua_only'
      analysisMode : String(20) default 'atc';
      // Namespace findings from TUA analysis (JSON: { unmaintained: [], noOwner: [] })
      nsFindings  : LargeString;
      // Input file references
      atcFileId   : UUID;
      cloneFileId : UUID;
      smodilogId  : UUID;
      trnspacetId : UUID;
      nsOwnerId   : UUID;
}

// ── Uploaded file blobs ────────────────────────────────────
entity Files {
  key id       : UUID;
      jobId    : UUID;
      role     : String(20);   // atc|clone|smodilog|trnspacet|ns_owner
      filename : String(500);
      mimeType : String(200);
      content  : LargeString;  // file bytes stored as hex string
      size     : Integer;
      createdAt: Timestamp     @cds.on.insert: $now;
}

// ── Output artifacts (generated files) ────────────────────
entity Artifacts {
  key id       : UUID;
      jobId    : UUID;
      role     : String(20);   // atc_result|tua|pptx|estimation
      filename : String(500);
      mimeType : String(200);
      content  : LargeString;  // file bytes stored as hex string
      size     : Integer;
      createdAt: Timestamp     @cds.on.insert: $now;
}
