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

> **Note:** If the server has no `API_KEY` environment variable configured (e.g. a private internal deployment), all endpoints are open and the header can be omitted.

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
| `wait` | string | No | `true` (default) = wait and return artifacts; `false` = async (returns jobId immediately). Must be the exact string `"false"` to trigger async — omitting it or any other value means synchronous. |

*Not required when `analysisMode=tua_only`

> **Note:** If an unrecognised `analysisMode` value is submitted, the server silently falls back to `atc`. No error is returned — check the `analysisMode` field in the response to confirm what was used.

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
{
  "jobId": "...",
  "status": "running",
  "statusMsg": "Classifying chunk 2/4...",
  "customer": "MyProject",
  "migType": "conversion",
  "analysisMode": "atc_tua"
}
```

If failed:
```json
{
  "jobId": "...",
  "status": "failed",
  "statusMsg": "Analysis failed — see error field",
  "error": "Analysis failed — see error field",
  "customer": "MyProject",
  "migType": "conversion",
  "analysisMode": "atc"
}
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

> **409 Conflict** is returned if the job status is not `done` — this includes both still-running and failed jobs.

> **404 Not Found** for an unknown role includes an `availableRoles` array to help discover what artifacts exist:
> ```json
> { "error": "Artifact \"xyz\" not found", "availableRoles": ["atc_result", "pptx", "estimation"] }
> ```

---

### `GET /api/v1/health` — Health Check (no auth required)

```json
{ "status": "ok", "version": "1.0.0" }
```

---

## How the Request Is Structured

The Excel file is **not** passed in the URL. It travels in the **request body** as a multipart form upload — the same mechanism as an HTML file-input form.

```
POST https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze
     ─────────────────────────────────────────────────────────────────────────
     ↑ URL only identifies the endpoint — no data here

Header:
  X-API-Key: <your-api-key>          ← authentication

Body (multipart/form-data):
  customer     = "Acme Corp"         ← plain text field
  analysisMode = "atc"               ← plain text field
  migType      = "conversion"        ← plain text field
  atcFile      = [binary .xlsx bytes] ← FILE — streamed in the body
```

URLs are limited to a few KB of text. A 10 MB Excel file can only travel inside the HTTP body. Every HTTP library handles this automatically when you use its multipart/form-data API — you just point it at the file path.

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
const FormData = require('form-data');  // npm install form-data axios
const fs       = require('fs');
const axios    = require('axios');

const API_BASE = 'https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com';
const API_KEY  = process.env.ATC_API_KEY;  // never hardcode — use env var

async function runAtcAnalysis({ customer, analysisMode = 'atc', migType = 'conversion',
                                 atcFile, cloneFile, smodilogFile, trnspacetFile, nsOwnerFile }) {
  const form = new FormData();

  // Text fields go in the body — NOT the URL
  form.append('customer',     customer);
  form.append('analysisMode', analysisMode);
  form.append('migType',      migType);

  // Files also go in the body as binary streams
  if (atcFile)       form.append('atcFile',      fs.createReadStream(atcFile));
  if (cloneFile)     form.append('cloneFile',     fs.createReadStream(cloneFile));
  if (smodilogFile)  form.append('smodilogFile',  fs.createReadStream(smodilogFile));
  if (trnspacetFile) form.append('trnspacetFile', fs.createReadStream(trnspacetFile));
  if (nsOwnerFile)   form.append('nsOwnerFile',   fs.createReadStream(nsOwnerFile));

  const response = await axios.post(`${API_BASE}/api/v1/analyze`, form, {
    headers: { ...form.getHeaders(), 'X-API-Key': API_KEY },
    timeout: 300_000,  // 5 minutes
  });

  const { counts, artifacts } = response.data;
  console.log('Counts:', counts);
  // counts: { total, hca, s4h, spdd, spau, fitGapDelta }

  // Each artifact has a base64 content field — decode and save
  for (const artifact of artifacts) {
    fs.writeFileSync(artifact.filename, Buffer.from(artifact.content, 'base64'));
    console.log(`Saved: ${artifact.filename}  (${artifact.size} bytes)`);
    // artifact.role: "atc_result" | "pptx" | "estimation" | "tua"
  }

  return response.data;
}

// ATC only
runAtcAnalysis({
  customer:  'Acme Corp',
  atcFile:   './ATC_Extract.xlsx',
});

// Full analysis — ATC + TUA + Clone, upgrade mode
runAtcAnalysis({
  customer:      'Acme Corp',
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
import os, requests, base64

API_BASE  = 'https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com'
API_KEY   = os.environ['ATC_API_KEY']  # never hardcode — use env var
XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

def run_atc_analysis(customer, analysis_mode='atc', mig_type='conversion',
                     atc_file=None, clone_file=None,
                     smodilog_file=None, trnspacet_file=None, ns_owner_file=None):

    # Text fields go in `data=` — NOT the URL
    data = {'customer': customer, 'analysisMode': analysis_mode, 'migType': mig_type}

    # Files go in `files=` — binary content streamed in the body
    files = {}
    def _add(field, path):
        if path:
            files[field] = (os.path.basename(path), open(path, 'rb'), XLSX_MIME)

    _add('atcFile',       atc_file)
    _add('cloneFile',     clone_file)
    _add('smodilogFile',  smodilog_file)
    _add('trnspacetFile', trnspacet_file)
    _add('nsOwnerFile',   ns_owner_file)

    try:
        r = requests.post(
            f'{API_BASE}/api/v1/analyze',
            headers={'X-API-Key': API_KEY},
            data=data,
            files=files,
            timeout=300,
        )
        r.raise_for_status()
    finally:
        for _, (_, fh, _) in files.items():
            fh.close()

    result = r.json()
    print('Counts:', result['counts'])
    # counts: { 'total', 'hca', 's4h', 'spdd', 'spau', 'fitGapDelta' }

    for artifact in result['artifacts']:
        with open(artifact['filename'], 'wb') as f:
            f.write(base64.b64decode(artifact['content']))
        print(f"Saved: {artifact['filename']}  ({artifact['size']} bytes)")
        # artifact['role']: "atc_result" | "pptx" | "estimation" | "tua"

    return result

# ATC only
run_atc_analysis('Acme Corp', atc_file='ATC_Extract.xlsx')

# ATC + Clone, conversion
run_atc_analysis('Acme Corp', analysis_mode='atc_clone',
                 atc_file='ATC_Extract.xlsx', clone_file='Clone_Extract.xlsx')
```

---

### Java (OkHttp)

```java
// Maven: com.squareup.okhttp3:okhttp:4.12.0  +  com.fasterxml.jackson.core:jackson-databind

import okhttp3.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.File;
import java.util.*;

public class AtcApiClient {

    private static final String API_BASE = "https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com";
    private static final MediaType XLSX   = MediaType.parse(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    private final OkHttpClient http = new OkHttpClient.Builder()
        .callTimeout(java.time.Duration.ofMinutes(5))
        .build();
    private final ObjectMapper json = new ObjectMapper();

    public Map<String, Object> runAnalysis(String customer, String analysisMode,
                                           String migType, File atcFile) throws Exception {

        // Text fields and files both go in the multipart body — NOT the URL
        MultipartBody.Builder body = new MultipartBody.Builder()
            .setType(MultipartBody.FORM)
            .addFormDataPart("customer",     customer)
            .addFormDataPart("analysisMode", analysisMode)
            .addFormDataPart("migType",      migType)
            .addFormDataPart("atcFile", atcFile.getName(),
                RequestBody.create(atcFile, XLSX));  // file streamed in body

        Request request = new Request.Builder()
            .url(API_BASE + "/api/v1/analyze")
            .addHeader("X-API-Key", System.getenv("ATC_API_KEY"))  // from env var
            .post(body.build())
            .build();

        try (Response response = http.newCall(request).execute()) {
            if (!response.isSuccessful())
                throw new RuntimeException("API error: " + response.code() + " " + response.body().string());

            Map result = json.readValue(response.body().string(), Map.class);

            // Save artifacts
            List<Map> artifacts = (List<Map>) result.get("artifacts");
            for (Map artifact : artifacts) {
                byte[] bytes = Base64.getDecoder().decode((String) artifact.get("content"));
                java.nio.file.Files.write(
                    java.nio.file.Path.of((String) artifact.get("filename")), bytes);
                System.out.println("Saved: " + artifact.get("filename"));
            }

            System.out.println("Counts: " + result.get("counts"));
            return result;
        }
    }

    public static void main(String[] args) throws Exception {
        new AtcApiClient().runAnalysis(
            "Acme Corp", "atc", "conversion", new File("ATC_Extract.xlsx"));
    }
}
```

---

### React / Browser App

> **Important:** Never call the ATC API directly from browser-side JavaScript — that would expose your API key to anyone who opens DevTools. Route all calls through your own backend.

```
Browser → Your backend server → ATC API
```

**Backend (Node/Express proxy):**

```javascript
const express  = require('express');
const multer   = require('multer');
const FormData = require('form-data');
const axios    = require('axios');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

app.post('/run-analysis', upload.single('atcFile'), async (req, res) => {
  const form = new FormData();
  form.append('customer',     req.body.customer);
  form.append('analysisMode', req.body.analysisMode || 'atc');
  form.append('migType',      req.body.migType      || 'conversion');
  // File buffer from multer re-attached to outgoing request body
  form.append('atcFile', req.file.buffer, { filename: req.file.originalname });

  const { data } = await axios.post(
    'https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze',
    form,
    { headers: { ...form.getHeaders(), 'X-API-Key': process.env.ATC_API_KEY } }
  );

  res.json(data);  // forward counts + artifacts to your frontend
});
```

**Frontend (React):**

```jsx
async function handleUpload(file, customer) {
  const form = new FormData();
  form.append('customer', customer);
  form.append('atcFile',  file);        // File object from <input type="file">

  const res    = await fetch('/run-analysis', { method: 'POST', body: form });
  const result = await res.json();

  console.log('Counts:', result.counts);

  // Download a file directly in the browser
  for (const artifact of result.artifacts) {
    const bytes = Uint8Array.from(atob(artifact.content), c => c.charCodeAt(0));
    const blob  = new Blob([bytes], { type: artifact.mimeType });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href      = url;
    a.download  = artifact.filename;
    a.click();
  }
}
```

---

### Async Flow (for large files)

Use `wait=false` when the ATC file exceeds ~5,000 rows to avoid HTTP timeouts.

```
Step 1 — Submit job
  POST /api/v1/analyze  (with wait=false)
  → 202 { "jobId": "uuid", "status": "running", "pollUrl": "/api/v1/analyze/uuid", "message": "..." }

Step 2 — Poll until done (every 5–10 seconds)
  GET /api/v1/analyze/{jobId}
  → { "status": "running", "statusMsg": "Classifying chunk 2/4...", "customer": "...", ... }
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
| `404 Not Found` | `jobId` does not exist, or artifact `role` not found (includes `availableRoles` in body) |
| `409 Conflict` | Download requested but job is not `done` (includes jobs that have `failed`) |
| `500 Internal Server Error` | Analysis failed — `error` field contains the reason |

Sync mode failure response (`wait=true`, job fails during processing):
```json
{ "jobId": "...", "status": "failed", "error": "reason", "customer": "MyProject" }
```

---

## Limits

| Limit | Value |
|---|---|
| Max file size per upload | 60 MB |
| Supported file formats | `.xlsx` only |
| Sync mode timeout | 5 minutes (use `wait=false` for large files) |
| Max rows | No hard limit — large files are auto-chunked internally |
