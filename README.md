# Conversion Assessment Tool (CAP ATC UI)

SAP S/4HANA Code Conversion Assessment platform. Processes ATC (ABAP Test Cockpit) exports, optionally enriches them with Clone and TUA (Transport Usage Analysis) data, and generates a PowerPoint presentation, classified Excel, and effort estimation workbook.

---

## Table of Contents

1. [Overview](#overview)
2. [How It Works](#how-it-works)
3. [Analysis Modes](#analysis-modes)
4. [Input File Reference](#input-file-reference)
5. [Step-by-Step Usage](#step-by-step-usage)
6. [Output Files](#output-files)
7. [ATC File Format Requirements](#atc-file-format-requirements)
8. [Clone File Format Requirements](#clone-file-format-requirements)
9. [TUA File Format Requirements](#tua-file-format-requirements)
10. [Migration Types](#migration-types)
11. [Troubleshooting](#troubleshooting)

---

## Overview

The tool accepts one or more Excel files exported from an SAP system, classifies each ABAP object finding, and produces:

- A classified and enriched **ATC result Excel** with remediation types, HCA/S4H flags, and clone references
- A pre-filled **effort estimation workbook** with HCA and S/4HANA object counts
- A **PowerPoint presentation** (slides 4–7) with charts and key metrics
- A **TUA Analysis Excel** (when TUA files are provided) with SPDD/SPAU counts and namespace findings

---

## How It Works

```
Upload files → Select analysis mode → Click Run Analysis
       ↓
  Classify ATC rows (app logic or SAP ATC API)
       ↓
  [Optional] Cross-reference Clone list
       ↓
  [Optional] TUA analysis (SMODILOG + TRNSPACET)
       ↓
  Generate PowerPoint + Excel + Estimation outputs
       ↓
  Download individually or as ZIP
```

All processing happens server-side. Large files (> 1,500 rows) are automatically split into chunks and processed sequentially.

---

## Analysis Modes

Select the mode that matches the data you have available:

| Mode | ATC File | Clone File | TUA Files | Use When |
|------|:--------:|:----------:|:---------:|----------|
| **ATC Only** | Required | — | — | Standard ATC export, no clone or TUA data |
| **ATC + Clone** | Required | Required | — | ATC plus clone analysis for cross-object findings |
| **ATC + TUA** | Required | — | Required | ATC plus transport usage analysis |
| **ATC + TUA + Clone** | Required | Required | Required | Full analysis with all three data sources |
| **TUA Only** | — | — | Required | Only TUA analysis needed (no ATC export) |

---

## Input File Reference

### Summary

| File | Role | Format | When Required |
|------|------|--------|---------------|
| ATC Export | Main ATC findings | `.xlsx` / `.xls` | All modes except TUA Only |
| Clone List | Clone cross-reference | `.xlsx` / `.xls` | ATC + Clone, ATC + TUA + Clone |
| SMODILOG | Transport modification log | `.xlsx` / `.xls` | Any TUA mode |
| TRNSPACET | Transport request namespace | `.xlsx` / `.xls` | Any TUA mode |

All files must be **Excel format** (`.xlsx` preferred, `.xls` also accepted). Other formats (CSV, XML, PDF) are not supported.

---

## Step-by-Step Usage

### Step 1 — Enter Customer Name

Type the customer or project name in the **Customer Name** field. This is used to label the PowerPoint presentation and output filenames.

> Example: `Office Works`, `Contoso AG`, `Project Phoenix`

### Step 2 — Select Analysis Mode

Click one of the five options in the **Analysis Option** segmented button:

- `ATC Only` — default; use for a standard ATC export with no additional files
- `ATC + Clone` — enables the Clone file uploader
- `ATC + TUA` — enables the TUA file uploaders (SMODILOG + TRNSPACET)
- `ATC + TUA + Clone` — enables all three optional uploaders
- `TUA Only` — hides the ATC uploader; SMODILOG + TRNSPACET are required

### Step 3 — Upload Files

Upload the required files based on the selected mode. Each uploader shows its file status after selection.

- Drag and drop a file onto the upload area, or click to open the file browser
- Only `.xlsx` and `.xls` files are accepted
- Each slot accepts one file at a time

### Step 4 — Select Migration Type

Choose the migration scenario:

| Option | When to Use |
|--------|-------------|
| **Conversion (ECC → S/4HANA)** | Customer is migrating from SAP ECC to S/4HANA for the first time |
| **Upgrade (S/4HANA → S/4HANA)** | Customer is upgrading from one S/4HANA release to a newer one |

### Step 5 — Run Analysis

Click **Run Analysis**. The button is enabled only when all required files for the selected mode have been uploaded.

A progress indicator shows the current processing step:
- Loading input files
- Classifying ATC data
- Parsing ATC results
- Generating TUA analysis *(if applicable)*
- Generating PowerPoint presentation
- Generating effort estimation
- **Analysis complete — all artifacts ready**

Processing time depends on file size. A 5,000-row ATC file typically completes in 30–60 seconds.

### Step 6 — Download Results

When analysis completes, the Results screen shows:

- Summary KPI panel (total findings, HCA count, S/4HANA count)
- Per-file download buttons for each generated artifact
- **Download All (ZIP)** button to get everything in one archive

---

## Output Files

| File | Description | Format |
|------|-------------|--------|
| `atc_classified_<timestamp>.xlsx` | ATC findings with Remediation Type, HCA/S4H, Clone, Priority columns populated | `.xlsx` |
| `effort_estimation_<timestamp>.xlsx` | Pre-filled estimation template with HCA and S/4HANA object counts in cells C4/C5 | `.xlsx` |
| `atc_presentation_<timestamp>.pptx` | PowerPoint with ATC summary charts on slides 4–7 | `.pptx` |
| `TUA_Analysis_<timestamp>.xlsx` | TUA analysis with SPDD/SPAU counts and namespace findings *(TUA modes only)* | `.xlsx` |

---

## ATC File Format Requirements

The ATC file must be an **Excel export from transaction SATC** (or equivalent) in the customer SAP system.

### Required Columns

The following columns must be present (exact names or recognised variants):

| Column Name | Accepted Variants | Description |
|-------------|-------------------|-------------|
| `Object name` | `Object Name` | ABAP object name (program, class, function group, etc.) |
| `Check Title` | — | ATC check rule title |
| `Check Message` | — | Detailed check message text |
| `Remediation Type` | — | Remediation classification (populated by ABAP or by the tool) |
| `Obj.` | `Object type`, `Obj` | ABAP object type code (PROG, CLAS, FUGR, etc.) |
| `Priority` | — | ATC finding priority (1 = highest) |

### Optional Columns (enriched by the tool)

| Column | Description |
|--------|-------------|
| `Syntax Error?` | Set to `Yes` for syntax error rows; propagated to all objects with a syntax error |
| `Fit Gap?` | Set to `Yes` for fit/gap findings |
| `HCA/S4H?` | Populated automatically — `HCA` or `S4H` based on object classification |
| `Clone?` | Populated if a Clone file is provided and the object appears in the clone list |
| `Note` | SAP Note number reference |
| `Referenced Object` | Cross-referenced object name |


> The tool auto-deduplicates column names if the same column appears twice (which can happen with some ABAP export configurations).

### File Size

- Files up to ~50,000 rows are supported
- Files larger than 1,500 rows are automatically chunked for processing
- Very large files (> 20,000 rows) may take several minutes

---

## Clone File Format Requirements

The Clone file identifies custom ABAP objects that are clones of SAP standard objects.

### File Naming

The filename must **start with `clone`** (case-insensitive), e.g.:
- `clone_list.xlsx`
- `Clone_Analysis_2024.xlsx`
- `CLONE_results.xls`

### Required Columns

The clone file must have at least one column containing the ABAP object name. The tool looks for column names matching, this should be the final clone list.

- `Object name` / `Object Name`
- `Program` / `program`
- Any first column if none of the above are found

### Source

Clone analysis is typically run using the **SAP Clone Finder** tool or equivalent. Export the results to Excel and upload.

---

## TUA File Format Requirements

TUA (Transport Usage Analysis) requires two files exported from SAP: **SMODILOG** and **TRNSPACET**.

### SMODILOG File

**Source:** Export from SAP table `SMODILOG` (System Modification Log)

**How to export:**
1. Run transaction **SE16** or **SE16N**, table `SMODILOG`
2. Export to **Spreadsheet (.xlsx)**

**Expected structure:**
- The tool reads from **row 24 onwards** (rows 1–23 are treated as headers/metadata)
- Required data columns: modification namespace entries with transport request references

### TRNSPACET File

**Source:** Export from SAP table `TRNSPACET` (Transport Request Namespace)

**How to export:**
1. Run transaction **SE16** or **SE16N**, table `TRNSPACET`
2. Export to **Spreadsheet (.xlsx)**

**Expected structure:**
- The tool reads from **row 10 onwards** (rows 1–9 are treated as headers/metadata)
- Required data columns: transport request and namespace assignments

### Optional: Namespace Owner File

An optional third file can be uploaded to provide namespace owner mappings (for master map lookup). This is not exposed in the current UI but is supported internally.

---

## Migration Types

| Migration Type | Object Classification Logic |
|---------------|----------------------------|
| **Conversion (ECC → S/4HANA)** | Uses full HCA and S/4HANA classification rules for ECC-to-S/4 object scope |
| **Upgrade (S/4HANA → S/4HANA)** | Applies upgrade-specific classification; some object types are treated differently |

Choose **Conversion** for all new S/4HANA migration projects unless the customer is already on S/4HANA.

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|-------------|-----|
| Run Analysis button stays disabled | Required file not uploaded, or customer name empty | Check all required fields for the selected mode |
| `ATC file content is empty` error | File was uploaded but contained no data rows | Re-export the ATC file from SATC and ensure it has findings |
| `XLS parse error` | Column names in the ATC file don't match expected format | Ensure the export includes the required columns listed above |
| `SAP API error` | Cloud Connector or I10 destination unreachable | Check CF destination `i10_ATC` and connectivity service binding |
| TUA Analysis not generated | SMODILOG or TRNSPACET file not uploaded | Both TUA files are required for TUA output |
| `TUA_Analysis.xlsx template not found` | Template missing from `db/templates/` | Ensure `TUA_Analysis.xlsx`, `template.pptx`, and `Estimation Template.xlsx` are in `db/templates/` |
| Download shows empty file | Analysis still running | Wait for status to show "Analysis complete" |
| PowerPoint charts are blank | ATC file had zero classifiable findings | Verify the ATC export contains findings with valid object types |

---

## File Format Quick Reference

```
ATC Export      →  .xlsx / .xls   (from SATC transaction)
Clone List      →  .xlsx / .xls   (filename must start with "clone")
SMODILOG        →  .xlsx / .xls   (SE16 export of table SMODILOG)
TRNSPACET       →  .xlsx / .xls   (SE16 export of table TRNSPACET)
```

All uploads: **Excel only** — no CSV, no XML, no PDF.
