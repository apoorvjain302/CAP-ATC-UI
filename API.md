# ATC Analysis — External REST API Guide

This document explains how third-party applications can integrate with the ATC Analysis API to upload SAP ATC extract files and receive classified results (Excel, PowerPoint, Effort Estimation) without any local installation.

---

## Base URL

```
https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com
```

---

## How the Data Flows

```
┌─────────────────────────────────────────────────────────────────┐
│                     Third-Party Application                     │
│                                                                 │
│  1. Prepare input files (ATC extract XLSX, optional extras)     │
│  2. Send POST /api/v1/analyze  (multipart/form-data)            │
│  3. Receive JSON response with:                                 │
│       - counts  (HCA, S/4H, SPDD, SPAU, etc.)                  │
│       - artifacts[] (base64-encoded Excel, PPT, Estimation)     │
│  4. Decode base64 → save files locally / display in your UI     │
└──────────────────────────┬──────────────────────────────────────┘
                           │  HTTPS POST (multipart)
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│              ATC Analysis API  (SAP BTP Cloud Foundry)          │
│                                                                 │
│  ① Receives uploaded files, stores them in-memory              │
│  ② Runs ATC classification engine (HCA / S4H detection,        │
│     Remediation Type, Syntax Error, Fit Gap, Clone flags)       │
│  ③ Generates:                                                   │
│       - Classified ATC Result  (.xlsx)                          │
│       - Executive Presentation (.pptx)                          │
│       - Effort Estimation      (.xlsx)                          │
│       - TUA Namespace Analysis (.xlsx)  [if TUA mode]          │
│  ④ Returns everything as JSON (counts + base64 file content)    │
└─────────────────────────────────────────────────────────────────┘
```

---

## Authentication

Set the `X-API-Key` header on every request.  
Contact the API owner to obtain the key.

```
X-API-Key: <your-api-key>
```

> If no API key has been configured on the server, this header can be omitted.

---

## Endpoints

### 1. `POST /api/v1/analyze` — Run Analysis

Accepts input files and parameters, runs the full classification pipeline, returns results.

**Request: `multipart/form-data`**

| Field | Type | Required | Description |
|---|---|---|---|
| `customer` | string | Yes | Customer or project name (used in output filenames) |
| `atcFile` | file (.xlsx) | Yes* | ATC extract from SAP SATC transaction |
| `cloneFile` | file (.xlsx) | No | Clone analysis file |
| `smodilogFile` | file (.xlsx) | TUA only | SMODILOG export (for namespace patch analysis) |
| `trnspacetFile` | file (.xlsx) | TUA only | TRNSPACET export (for namespace patch analysis) |
| `nsOwnerFile` | file (.xlsx) | No | Namespace owner mapping file |
| `migType` | string | No | `conversion` (default) or `upgrade` |
| `analysisMode` | string | No | See modes table below. Default: `atc` |
| `wait` | string | No | `true` (default) = wait for result; `false` = async mode |

*Not required when `analysisMode=tua_only`

**Analysis Modes**

| `analysisMode` | Files needed | Output |
|---|---|---|
| `atc` | `atcFile` | Classified Excel + PPT + Estimation |
| `atc_clone` | `atcFile` + `cloneFile` | Same + Clone column populated |
| `atc_tua` | `atcFile` + `smodilogFile` + `trnspacetFile` | All above + TUA Excel |
| `atc_tua_clone` | All files | Full output |
| `tua_only` | `smodilogFile` + `trnspacetFile` | TUA Excel + PPT only |

---

**Response: `200 OK`** (when `wait=true`)

```json
{
  "jobId": "f3c2a1b0-4e5d-4a2b-9c1d-123456789abc",
  "status": "done",
  "customer": "MyProject",
  "migType": "conversion",
  "analysisMode": "atc",
  "counts": {
    "total": 1240,
    "hca": 85,
    "s4h": 320,
    "spdd": 0,
    "spau": 0,
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
    }
  ]
}
```

**Artifact roles**

| `role` | Description |
|---|---|
| `atc_result` | Classified ATC Excel — original columns + HCA/S4H?, Remediation Type, Syntax Error?, Fit Gap?, Clone?, Automation Fix?, CCM Agent Fix? |
| `pptx` | Executive PowerPoint presentation with charts |
| `estimation` | Effort estimation Excel with remediation counts |
| `tua` | TUA namespace patch analysis Excel (TUA modes only) |

---

### 2. `GET /api/v1/analyze/:jobId` — Poll Async Job

Use this after submitting with `wait=false` to check job status.

```
GET /api/v1/analyze/f3c2a1b0-4e5d-4a2b-9c1d-123456789abc
X-API-Key: <your-api-key>
```

Response while running:
```json
{
  "jobId": "f3c2a1b0-...",
  "status": "running",
  "statusMsg": "Classifying chunk 2/4..."
}
```

Response when done — same structure as synchronous response above (includes `counts` and `artifacts`).

---

### 3. `GET /api/v1/analyze/:jobId/download/:role` — Download File Directly

Download a single artifact as a binary file (no base64 decoding needed).

```
GET /api/v1/analyze/f3c2a1b0-.../download/atc_result
GET /api/v1/analyze/f3c2a1b0-.../download/pptx
GET /api/v1/analyze/f3c2a1b0-.../download/estimation
GET /api/v1/analyze/f3c2a1b0-.../download/tua
```

Returns the file as a binary download with `Content-Disposition: attachment`.

---

### 4. `GET /api/v1/health` — Health Check

```
GET /api/v1/health
```

```json
{ "status": "ok", "version": "1.0.0" }
```

---

## Implementation Examples

### JavaScript / Node.js

```javascript
const FormData = require('form-data');
const fs       = require('fs');
const axios    = require('axios');

async function runAtcAnalysis(atcFilePath, customer) {
  const form = new FormData();
  form.append('customer', customer);
  form.append('migType', 'conversion');
  form.append('analysisMode', 'atc');
  form.append('atcFile', fs.createReadStream(atcFilePath));

  const response = await axios.post(
    'https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze',
    form,
    {
      headers: {
        ...form.getHeaders(),
        'X-API-Key': process.env.ATC_API_KEY,
      },
      timeout: 300000, // 5 min — allow time for large files
    }
  );

  const { counts, artifacts } = response.data;
  console.log('Counts:', counts);

  // Save each artifact to disk
  for (const artifact of artifacts) {
    const buf = Buffer.from(artifact.content, 'base64');
    fs.writeFileSync(artifact.filename, buf);
    console.log(`Saved: ${artifact.filename} (${artifact.size} bytes)`);
  }
}

runAtcAnalysis('./ATC_Extract.xlsx', 'MyProject');
```

---

### Python

```python
import requests
import base64
import os

def run_atc_analysis(atc_file_path, customer):
    url = "https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze"
    headers = {"X-API-Key": os.environ.get("ATC_API_KEY", "")}

    with open(atc_file_path, "rb") as f:
        files = {"atcFile": (os.path.basename(atc_file_path), f,
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        data  = {"customer": customer, "migType": "conversion", "analysisMode": "atc"}
        response = requests.post(url, headers=headers, files=files, data=data, timeout=300)

    response.raise_for_status()
    result = response.json()

    print("Counts:", result["counts"])

    # Save each artifact to disk
    for artifact in result["artifacts"]:
        content = base64.b64decode(artifact["content"])
        with open(artifact["filename"], "wb") as out:
            out.write(content)
        print(f"Saved: {artifact['filename']} ({artifact['size']} bytes)")

run_atc_analysis("ATC_Extract.xlsx", "MyProject")
```

---

### Java (OkHttp)

```java
import okhttp3.*;
import java.io.*;
import java.util.Base64;

OkHttpClient client = new OkHttpClient.Builder()
    .callTimeout(300, TimeUnit.SECONDS)
    .build();

RequestBody requestBody = new MultipartBody.Builder()
    .setType(MultipartBody.FORM)
    .addFormDataPart("customer", "MyProject")
    .addFormDataPart("migType", "conversion")
    .addFormDataPart("analysisMode", "atc")
    .addFormDataPart("atcFile", "ATC_Extract.xlsx",
        RequestBody.create(new File("ATC_Extract.xlsx"),
            MediaType.parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")))
    .build();

Request request = new Request.Builder()
    .url("https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze")
    .addHeader("X-API-Key", System.getenv("ATC_API_KEY"))
    .post(requestBody)
    .build();

Response response = client.newCall(request).execute();
// Parse JSON response and decode base64 artifacts
```

---

### ABAP (SAP system — using CL_HTTP_CLIENT)

```abap
DATA: lo_client  TYPE REF TO if_http_client,
      lo_request TYPE REF TO if_http_request.

cl_http_client=>create_by_url(
  EXPORTING url    = 'https://cap-atc-ui-gdh.cfapps.eu10-005.hana.ondemand.com/api/v1/analyze'
  IMPORTING client = lo_client ).

lo_request = lo_client->request.
lo_request->set_method( 'POST' ).
lo_request->set_header_field( name = 'X-API-Key' value = '<your-api-key>' ).

" Build multipart body with atcFile + customer/migType params
" Decode base64 response artifacts using cl_http_utility=>decode_base64()
```

---

## Async Flow (for large files)

Use `wait=false` when the ATC file has more than ~5,000 rows to avoid HTTP timeouts.

```
1.  POST /api/v1/analyze?wait=false
        → 202 { jobId, pollUrl }

2.  Loop: GET /api/v1/analyze/{jobId}
        → { status: "running", statusMsg: "Classifying chunk 2/4..." }
        → { status: "running", statusMsg: "Generating PowerPoint..." }
        → { status: "done",    counts: {...}, artifacts: [...] }   ← stop polling

3.  Decode artifacts OR call GET /api/v1/analyze/{jobId}/download/{role}
```

Recommended polling interval: every 5–10 seconds.

---

## Error Responses

| HTTP Status | Meaning |
|---|---|
| `400 Bad Request` | Missing required field (`customer`, `atcFile`, or TUA files) |
| `401 Unauthorized` | Missing or wrong `X-API-Key` |
| `404 Not Found` | `jobId` does not exist |
| `409 Conflict` | Download requested but job not done yet |
| `500 Internal Server Error` | Analysis failed — `error` field contains the reason |

Error response body:
```json
{ "error": "atcFile is required (unless analysisMode=tua_only)" }
```

---

## Limits

| Limit | Value |
|---|---|
| Max file size per upload | 60 MB |
| Request timeout (sync mode) | 5 minutes |
| Supported file formats | `.xlsx` only |
| Max rows per file | No hard limit (large files auto-chunked internally) |
