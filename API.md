# ATC Analysis — External REST API Guide

This document explains how third-party applications can integrate with the ATC Analysis API
to upload SAP ATC extract files and receive classified results (Excel, PowerPoint, Effort
Estimation) without any local installation.

---

## Base URL

```
https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com
```

---

## Authentication

Every request to `/api/v1/analyze` must include the API key in the request header:

```
X-API-Key: <your-api-key>
```

The `/api/v1/health` endpoint does **not** require authentication.

---

## How the Data Flows

```
┌─────────────────────────────────────────────────────────────────────┐
│                      Third-Party Application                        │
│                                                                     │
│  1. Prepare input files (ATC extract XLSX + optional extras)        │
│  2. POST /api/v1/analyze  (multipart/form-data + X-API-Key header)  │
│  3. Receive JSON response with:                                     │
│       ├─ counts  (total, HCA, S/4H, SPDD, SPAU, FitGapDelta)       │
│       └─ artifacts[] (base64-encoded Excel, PPT, Estimation, TUA)  │
│  4. Decode base64 → save / display files in your app               │
└──────────────────────────┬──────────────────────────────────────────┘
                           │  HTTPS POST (multipart/form-data)
                           ▼
┌─────────────────────────────────────────────────────────────────────┐
│           ATC Analysis API  (SAP BTP Cloud Foundry EU10)            │
│                                                                     │
│  ① Receives uploaded files                                         │
│  ② Validates API key                                               │
│  ③ Runs classification engine:                                     │
│       - HCA / S/4HANA detection                                    │
│       - Remediation Type (Mandatory, Needs Rem., False Positive,   │
│         Fit Gap, Can be Ignored, Optional, Syntax Error)           │
│       - Syntax Error propagation                                   │
│       - Fit Gap / Fit Gap Delta propagation                        │
│       - Clone detection                                            │
│       - Automation Fix / CCM Agent Fix eligibility                 │
│  ④ Generates output files:                                         │
│       - Classified ATC Result  (.xlsx)                             │
│       - Executive Presentation (.pptx)                             │
│       - Effort Estimation      (.xlsx)                             │
│       - TUA Namespace Analysis (.xlsx)  [TUA modes only]           │
│  ⑤ Returns JSON with counts + base64-encoded file content          │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Endpoints

### `POST /api/v1/analyze` — Run Analysis

Accepts files and parameters, runs the full pipeline, returns results.

**Request:** `multipart/form-data`

| Field | Type | Required | Description |
|---|---|---|---|
| `customer` | string | **Yes** | Customer / project name (appears in output filenames) |
| `analysisMode` | string | No | See modes table below. Default: `atc` |
| `migType` | string | No | `conversion` (default) or `upgrade` |
| `atcFile` | file (.xlsx) | Yes* | ATC extract exported from SAP SATC transaction |
| `cloneFile` | file (.xlsx) | No | Clone analysis file |
| `smodilogFile` | file (.xlsx) | TUA only | SMODILOG export (namespace patch analysis) |
| `trnspacetFile` | file (.xlsx) | TUA only | TRNSPACET export (namespace patch analysis) |
| `nsOwnerFile` | file (.xlsx) | No | Namespace owner mapping file |
| `wait` | string | No | `true` (default) = wait and return artifacts; `false` = async (returns jobId immediately) |

*Not required when `analysisMode=tua_only`

---

### Analysis Modes

| `analysisMode` | Required Files | Output Artifacts |
|---|---|---|
| `atc` | `atcFile` | Classified Excel + PPT + Estimation |
| `atc_clone` | `atcFile` + `cloneFile` | Classified Excel (with Clone column) + PPT + Estimation |
| `atc_tua` | `atcFile` + `smodilogFile` + `trnspacetFile` | Classified Excel + PPT + Estimation + TUA Excel |
| `atc_tua_clone` | `atcFile` + `cloneFile` + `smodilogFile` + `trnspacetFile` | All of the above |
| `tua_only` | `smodilogFile` + `trnspacetFile` | TUA Excel + PPT |

### Migration Types

| `migType` | Description |
|---|---|
| `conversion` | ECC → S/4HANA system conversion (default) |
| `upgrade` | S/4HANA → S/4HANA in-place upgrade |

> `migType` affects classification rules — certain findings are marked False Positive in upgrade
> mode that would be Mandatory in conversion mode.

---

### Response: `200 OK` (synchronous — default)

```json
{
  "jobId": "f3c2a1b0-4e5d-4a2b-9c1d-123456789abc",
  "status": "done",
  "customer": "MyProject",
  "migType": "conversion",
  "analysisMode": "atc_tua",
  "counts": {
    "total": 1240,
    "hca": 85,
    "s4h": 320,
    "spdd": 12,
    "spau": 8,
    "fitGapDelta": 4
  },
  "artifacts": [
    {
      "role": "atc_result",
      "filename": "atc_classified_20260911T120000.xlsx",
      "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "size": 102400,
      "content": "<base64-encoded file bytes>"
    },
    {
      "role": "pptx",
      "filename": "atc_presentation_20260911T120000.pptx",
      "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "size": 512000,
      "content": "<base64-encoded file bytes>"
    },
    {
      "role": "estimation",
      "filename": "effort_estimation_20260911T120000.xlsx",
      "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "size": 48000,
      "content": "<base64-encoded file bytes>"
    },
    {
      "role": "tua",
      "filename": "TUA_Analysis_20260911T120000.xlsx",
      "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "size": 61000,
      "content": "<base64-encoded file bytes>"
    }
  ]
}
```

### Artifact Roles

| `role` | Description |
|---|---|
| `atc_result` | Classified ATC Excel — original columns + HCA/S4H?, Remediation Type, Syntax Error?, Fit Gap?, Clone?, Automation Fix?, CCM Agent Fix? |
| `pptx` | Executive PowerPoint presentation with analysis charts |
| `estimation` | Effort estimation Excel with remediation effort counts |
| `tua` | TUA namespace patch analysis Excel (TUA modes only) |

---

### `GET /api/v1/analyze/:jobId` — Poll Async Job

Use after submitting with `wait=false`.

```
GET /api/v1/analyze/f3c2a1b0-4e5d-4a2b-9c1d-123456789abc
X-API-Key: <your-api-key>
```

While running:
```json
{ "jobId": "...", "status": "running", "statusMsg": "Classifying chunk 2/4..." }
```

When done — same structure as synchronous response above (includes `counts` and `artifacts`).

---

### `GET /api/v1/analyze/:jobId/download/:role` — Download File Directly

Download a single artifact as a binary file (no base64 decoding needed in the client).

```
GET /api/v1/analyze/{jobId}/download/atc_result
GET /api/v1/analyze/{jobId}/download/pptx
GET /api/v1/analyze/{jobId}/download/estimation
GET /api/v1/analyze/{jobId}/download/tua
```

Returns the file with `Content-Disposition: attachment` and the correct MIME type.

---

### `GET /api/v1/health` — Health Check (no auth required)

```json
{ "status": "ok", "version": "1.0.0" }
```

---

## Implementation Examples

### cURL (command line)

**ATC Only — Conversion:**
```bash
curl -X POST https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze \
  -H "X-API-Key: <your-api-key>" \
  -F "customer=MyProject" \
  -F "analysisMode=atc" \
  -F "migType=conversion" \
  -F "atcFile=@/path/to/ATC_Extract.xlsx"
```

**ATC + Clone — Upgrade:**
```bash
curl -X POST https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze \
  -H "X-API-Key: <your-api-key>" \
  -F "customer=MyProject" \
  -F "analysisMode=atc_clone" \
  -F "migType=upgrade" \
  -F "atcFile=@/path/to/ATC_Extract.xlsx" \
  -F "cloneFile=@/path/to/Clone_Extract.xlsx"
```

**ATC + TUA + Clone — Conversion (full analysis):**
```bash
curl -X POST https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze \
  -H "X-API-Key: <your-api-key>" \
  -F "customer=MyProject" \
  -F "analysisMode=atc_tua_clone" \
  -F "migType=conversion" \
  -F "atcFile=@/path/to/ATC_Extract.xlsx" \
  -F "cloneFile=@/path/to/Clone_Extract.xlsx" \
  -F "smodilogFile=@/path/to/SMODILOG.xlsx" \
  -F "trnspacetFile=@/path/to/TRNSPACET.xlsx" \
  -F "nsOwnerFile=@/path/to/NS_Owner.xlsx"
```

**TUA Only:**
```bash
curl -X POST https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze \
  -H "X-API-Key: <your-api-key>" \
  -F "customer=MyProject" \
  -F "analysisMode=tua_only" \
  -F "smodilogFile=@/path/to/SMODILOG.xlsx" \
  -F "trnspacetFile=@/path/to/TRNSPACET.xlsx"
```

Save an artifact from the response:
```bash
# Linux / Mac
echo "<base64 content from response>" | base64 -d > atc_result.xlsx

# Windows PowerShell
[System.Convert]::FromBase64String("<base64>") | Set-Content atc_result.xlsx -Encoding Byte
```

---

### JavaScript / Node.js

```javascript
const FormData = require('form-data');
const fs       = require('fs');
const axios    = require('axios');

const API_BASE = 'https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com';
const API_KEY  = '<your-api-key>';

async function runAtcAnalysis({ customer, analysisMode = 'atc', migType = 'conversion',
                                 atcFile, cloneFile, smodilogFile, trnspacetFile, nsOwnerFile }) {
  const form = new FormData();
  form.append('customer', customer);
  form.append('analysisMode', analysisMode);
  form.append('migType', migType);

  if (atcFile)       form.append('atcFile',       fs.createReadStream(atcFile));
  if (cloneFile)     form.append('cloneFile',      fs.createReadStream(cloneFile));
  if (smodilogFile)  form.append('smodilogFile',   fs.createReadStream(smodilogFile));
  if (trnspacetFile) form.append('trnspacetFile',  fs.createReadStream(trnspacetFile));
  if (nsOwnerFile)   form.append('nsOwnerFile',    fs.createReadStream(nsOwnerFile));

  const response = await axios.post(`${API_BASE}/api/v1/analyze`, form, {
    headers: { ...form.getHeaders(), 'X-API-Key': API_KEY },
    timeout: 300000,  // 5 minutes
  });

  const { counts, artifacts } = response.data;
  console.log('Analysis complete. Counts:', counts);

  // Save each artifact to disk
  for (const artifact of artifacts) {
    const buffer = Buffer.from(artifact.content, 'base64');
    fs.writeFileSync(artifact.filename, buffer);
    console.log(`Saved: ${artifact.filename}  (${artifact.size} bytes)`);
  }

  return response.data;
}

// Example: ATC + TUA + Clone, upgrade mode
runAtcAnalysis({
  customer:      'CustomerABC',
  analysisMode:  'atc_tua_clone',
  migType:       'upgrade',
  atcFile:       './ATC_Extract.xlsx',
  cloneFile:     './Clone_Extract.xlsx',
  smodilogFile:  './SMODILOG.xlsx',
  trnspacetFile: './TRNSPACET.xlsx',
});
```

---

### Python

```python
import requests, base64, os

API_BASE = "https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com"
API_KEY  = "<your-api-key>"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def run_atc_analysis(customer, analysis_mode="atc", mig_type="conversion",
                     atc_file=None, clone_file=None,
                     smodilog_file=None, trnspacet_file=None, ns_owner_file=None):

    headers = {"X-API-Key": API_KEY}
    data    = {"customer": customer, "analysisMode": analysis_mode, "migType": mig_type}
    files   = {}

    def _add(field, path):
        if path:
            files[field] = (os.path.basename(path), open(path, "rb"), XLSX_MIME)

    _add("atcFile",       atc_file)
    _add("cloneFile",     clone_file)
    _add("smodilogFile",  smodilog_file)
    _add("trnspacetFile", trnspacet_file)
    _add("nsOwnerFile",   ns_owner_file)

    try:
        response = requests.post(
            f"{API_BASE}/api/v1/analyze",
            headers=headers, data=data, files=files, timeout=300
        )
        response.raise_for_status()
    finally:
        for _, (_, fh, _) in files.items():
            fh.close()

    result = response.json()
    print("Counts:", result["counts"])

    for artifact in result["artifacts"]:
        content = base64.b64decode(artifact["content"])
        with open(artifact["filename"], "wb") as f:
            f.write(content)
        print(f"Saved: {artifact['filename']}  ({artifact['size']} bytes)")

    return result

# Example: ATC + Clone, conversion mode
run_atc_analysis(
    customer      = "CustomerABC",
    analysis_mode = "atc_clone",
    mig_type      = "conversion",
    atc_file      = "ATC_Extract.xlsx",
    clone_file    = "Clone_Extract.xlsx",
)
```

---

### Async Flow (for large files)

Use `wait=false` when the ATC file exceeds ~5,000 rows to avoid HTTP timeouts.

```
Step 1 — Submit job
  POST /api/v1/analyze  (with wait=false)
  → 202 { "jobId": "uuid", "pollUrl": "/api/v1/analyze/uuid" }

Step 2 — Poll until done (every 5–10 seconds)
  GET /api/v1/analyze/{jobId}
  → { "status": "running", "statusMsg": "Classifying chunk 2/4..." }
  → { "status": "running", "statusMsg": "Generating PowerPoint..." }
  → { "status": "done",    "counts": {...}, "artifacts": [...] }  ← stop here

Step 3 — Use artifacts
  Option A: decode base64 content from the poll response
  Option B: GET /api/v1/analyze/{jobId}/download/{role}  → binary file download
```

Python async polling example:
```python
import time

# Step 1: submit
resp = requests.post(f"{API_BASE}/api/v1/analyze",
    headers={"X-API-Key": API_KEY},
    data={"customer": "ABC", "analysisMode": "atc", "migType": "conversion", "wait": "false"},
    files={"atcFile": ("ATC.xlsx", open("ATC.xlsx","rb"), XLSX_MIME)},
    timeout=30
)
job_id = resp.json()["jobId"]

# Step 2: poll
while True:
    poll = requests.get(f"{API_BASE}/api/v1/analyze/{job_id}",
                        headers={"X-API-Key": API_KEY}).json()
    print(poll["statusMsg"])
    if poll["status"] == "done":
        break
    if poll["status"] == "failed":
        raise Exception(poll["error"])
    time.sleep(8)

# Step 3: save artifacts
for artifact in poll["artifacts"]:
    with open(artifact["filename"], "wb") as f:
        f.write(base64.b64decode(artifact["content"]))
```

---

## Error Responses

| HTTP Status | Cause |
|---|---|
| `400 Bad Request` | Missing required field (`customer`, `atcFile`, or TUA files for TUA mode) |
| `401 Unauthorized` | Missing or incorrect `X-API-Key` |
| `404 Not Found` | `jobId` does not exist |
| `409 Conflict` | Download requested but job is not done yet |
| `500 Internal Server Error` | Analysis failed — `error` field contains the reason |

```json
{ "error": "smodilogFile and trnspacetFile are required for TUA analysis modes" }
```

---

## Limits

| Limit | Value |
|---|---|
| Max file size per upload | 60 MB |
| Supported file formats | `.xlsx` only |
| Sync mode timeout | 5 minutes (use `wait=false` for large files) |
| Max rows | No hard limit — large files are auto-chunked internally |
