# SmartTech Points List Extractor

A local Python application for converting BAS/BMS points information from IFC drawing sets into structured Excel workbooks, with optional equipment-submittal enrichment and export to the standardized TATSU Points List format.

The SmartTech Points List Extractor was built to reduce the repetitive manual work involved in reading controls drawings, locating points-summary tables, recreating those tables in Excel, expanding repeated equipment points, and translating drawing information into a consistent integration points-list structure.

The application is implemented in Python with a Streamlit user interface and performs all document processing locally on the user's computer. Project PDFs are not sent to an external API or cloud service.

---

## Table of Contents

- [Overview](#overview)
- [Why This Tool Exists](#why-this-tool-exists)
- [Core Features](#core-features)
- [How the Workflow Works](#how-the-workflow-works)
- [System Architecture](#system-architecture)
- [Drawing Table Detection](#drawing-table-detection)
- [Extraction and Validation Logic](#extraction-and-validation-logic)
- [Submittal and BACnet Object Processing](#submittal-and-bacnet-object-processing)
- [Human-in-the-Loop Verification](#human-in-the-loop-verification)
- [TATSU Mapping Logic](#tatsu-mapping-logic)
- [Alarm and Trending Translation](#alarm-and-trending-translation)
- [Excel Output](#excel-output)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Running the Application](#running-the-application)
- [Step-by-Step Usage](#step-by-step-usage)
- [Input Requirements](#input-requirements)
- [Output Files](#output-files)
- [Troubleshooting](#troubleshooting)
- [Data Privacy and Security](#data-privacy-and-security)
- [Known Limitations](#known-limitations)
- [Engineering Verification](#engineering-verification)
- [Project Structure](#project-structure)
- [Dependencies](#dependencies)
- [Technical Design Notes](#technical-design-notes)
- [Potential Future Improvements](#potential-future-improvements)
- [Author](#author)
- [License](#license)

---

## Overview

Controls and building-management-system projects frequently require engineers to transform information contained in IFC drawing sets into structured points lists for integration, programming, checkout, and commissioning.

That process can become repetitive very quickly:

1. Open a large drawing set.
2. Find every equipment points-summary table.
3. Copy each row into Excel.
4. Recreate the equipment tabs.
5. Interpret I/O and integration columns.
6. Convert trending intervals.
7. Translate alarm priorities.
8. Review vendor submittals for BACnet object information.
9. Expand generic drawing points into multiple equipment instances.
10. Populate tags, object types, addresses, and units.
11. Reformat the entire workbook into the required project standard.

The SmartTech Points List Extractor automates a large portion of that workflow while deliberately keeping the engineer responsible for final verification.

At a high level, the application performs three major operations:

```text
IFC Drawing PDFs
        |
        v
Detect and Extract Points Tables
        |
        v
Optional Submittal / BACnet Enrichment
        |
        v
Engineer Review and Approval
        |
        v
TATSU or Raw Excel Workbook
```

The application is not intended to replace engineering judgment. It is intended to remove repetitive transcription, formatting, and deterministic mapping work so the engineer can spend more time validating the final deliverable.

---

## Why This Tool Exists

BAS points-list development is often a high-volume task where a significant amount of effort is spent moving information from one document format into another.

Manual workflows create several problems:

- Repetitive copy/paste work
- Inconsistent formatting between equipment tabs
- Risk of missing an equipment table
- Risk of transcription errors
- Manual conversion of poll intervals
- Manual interpretation of I/O types
- Manual alarm-priority mapping
- Repeated review of vendor BACnet object tables
- Difficulty tracking which rows were added from submittals
- Time spent rebuilding a standard workbook layout

This project approaches the problem as a deterministic document-processing pipeline.

Instead of asking a generative model to create a points list, the program attempts to extract the information that already exists in the engineering documents, apply defined mapping rules, identify uncertain cases, and expose those cases to the engineer for review.

This design makes the workflow easier to audit and keeps the original source documents central to the process.

---

## Core Features

### IFC Drawing Table Extraction

The application scans one or more IFC drawing PDFs and identifies points tables containing a `POINT DESCRIPTION` column.

It supports:

- Automatic table detection
- Specific-title/marker detection
- Multiple drawing PDFs in one session
- Multiple points tables on the same drawing sheet
- Side-by-side tables
- Vertically stacked tables
- Tables that continue across drawing pages
- Equipment-name extraction for Excel tab naming
- Editable extracted data before export
- Detection reporting
- Errors-and-omissions reporting

### Multiple Extraction Strategies

PDF engineering drawings are not always structured consistently. A table may visually look perfect while its internal PDF text geometry is difficult to parse.

The extractor therefore does not rely on a single table-detection strategy.

It can attempt:

- Standard ruled-table extraction
- Text-based vertical-boundary extraction
- Text-based vertical and horizontal extraction
- Per-region extraction when multiple tables appear on one sheet
- Raw drawing-line reconstruction as a fallback for difficult grid geometry

The goal is to recover structured table data while minimizing accidental extraction of unrelated drawing content.

### Table Quality Validation

Candidate tables are checked before being accepted.

The application verifies characteristics such as:

- Presence of a `POINT DESCRIPTION` header
- Minimum row and column counts
- Whether rows resemble actual points rows
- Whether enough rows pass the point-row quality test
- Whether the detected title indicates a legend rather than a points table

Low-quality candidates may be rejected and surfaced in the detection report instead of being silently included.

### Errors and Omissions Reporting

A major design goal of the application is that failures should be visible.

The interface reports cases such as:

- Marker/title not found
- `POINT DESCRIPTION` text not found
- Points text found but no readable grid detected
- Candidate table rejected as a likely frame, SOO text box, or misread grid
- Table extracted but title could not be read
- PDF could not be opened

When possible, the interface also gives the user a recommended action.

### Editable Data Before Export

Every extracted table is displayed in the Streamlit interface.

Before generating the workbook, the user can:

- Rename the Excel sheet
- Review rows and columns
- Edit individual cells
- Add or delete rows
- Correct extraction issues manually

This provides an additional verification layer before the final Excel file is created.

---

## How the Workflow Works

The application is organized into three main steps.

### Step 1 — Extract Points Tables

Upload one or more IFC drawing PDFs.

The recommended mode is:

```text
Auto-detect points tables
```

The application scans for pages containing a `POINT DESCRIPTION` column and attempts to isolate valid points tables.

After scanning, the interface displays:

- Detection results
- Source drawing filename
- Source drawing page
- Detected table title
- Rejected candidates
- Errors or omissions
- Extracted equipment tables

Each accepted table can be reviewed and edited directly in the application.

### Step 2 — Expand or Enrich from Submittals

This step is optional.

Upload equipment-submittal PDFs and map each submittal to the correct extracted equipment table.

For example:

```text
CRAH submittal -> CRAH points table
FWU submittal  -> FWU points table
```

The application then looks for information that can improve the extracted drawing points list.

Depending on the submittal structure, it may propose:

- Repeated-point expansion
- BACnet object tags
- BACnet object types
- Object instance/address values
- Units
- Multiple equipment instances

Importantly, the application does not automatically accept these proposed changes.

The engineer reviews the proposal before applying it.

### Step 3 — Export

Two output modes are available:

#### TATSU Points List

Produces the standardized TATSU workbook structure with SmartTech-style formatting and deterministic mapping from drawing fields into the TATSU columns.

#### As Extracted from Drawing

Produces a workbook that more closely preserves the original extracted table structure.

The output workbook is downloaded as:

```text
Points_List_Extracted.xlsx
```

---

## System Architecture

The application is a local Python/Streamlit workflow.

A simplified architecture is:

```text
+----------------------------+
|       Streamlit UI         |
+-------------+--------------+
              |
              v
+----------------------------+
|      PDF Input Layer       |
|        PyMuPDF / fitz      |
+-------------+--------------+
              |
              v
+----------------------------+
| Drawing Extraction Engine  |
| - text detection           |
| - table finding            |
| - region carving           |
| - line-grid reconstruction |
| - validation               |
+-------------+--------------+
              |
              v
+----------------------------+
| Structured Points Data     |
| Python lists / dictionaries|
| pandas DataFrames in UI    |
+-------------+--------------+
              |
       +------+------+
       |             |
       v             v
+-------------+ +-------------------+
| Raw Export  | | Submittal Engine  |
|             | | - BACnet parsing  |
|             | | - matching        |
|             | | - expansion       |
+------+------+ +---------+---------+
       |                  |
       +--------+---------+
                |
                v
+----------------------------+
|      TATSU Mapper          |
| - I/O type                 |
| - integration mode         |
| - trending                 |
| - alarms                   |
| - notes                    |
| - submittal fills          |
+-------------+--------------+
              |
              v
+----------------------------+
|      openpyxl Writer       |
|      .xlsx Workbook        |
+----------------------------+
```

No external web service is required for document processing.

---

## Drawing Table Detection

### Auto-Detect Mode

Auto-detect mode is the recommended default.

Instead of requiring every drawing to use one exact title such as:

```text
TYPICAL POINTS SUMMARY
```

the application looks for pages containing a `POINT DESCRIPTION` column.

This allows the extractor to work with multiple common BAS naming conventions, including variations such as:

- Points Summary
- Points List
- Points Schedule
- Points Matrix
- I/O Summary
- I/O List
- Hardwired Points
- Integration Points
- BACnet Points
- BAS Points
- BMS Points
- Control Points
- Alarm Points

Title keywords help identify and name tables, but auto-detection can still work when the exact title wording is different.

### Specific Title Mode

The user can also select:

```text
Find specific title
```

A title marker can then be supplied.

The default marker is:

```text
TYPICAL POINTS SUMMARY
```

Marker matching is designed to tolerate case and spacing differences.

This mode is useful when:

- A drawing set uses a known table title
- Auto-detection finds too much unrelated content
- The engineer wants to target a specific type of table

---

## Extraction and Validation Logic

Engineering PDFs can store apparently simple tables in surprisingly complicated ways.

For that reason, the extraction engine includes several layers of logic.

### 1. Candidate Page Detection

In auto mode, pages without `POINT DESCRIPTION` text are skipped.

This prevents unnecessary table extraction attempts on unrelated drawing sheets.

### 2. Table Discovery

PyMuPDF table detection is used to identify structured table candidates.

### 3. Header Validation

A candidate must contain a valid `POINT DESCRIPTION` header near the top of the extracted table.

### 4. Side-by-Side Table Splitting

Two adjacent points tables may be interpreted by the PDF parser as one very wide table.

The application checks for multiple `POINT DESCRIPTION` headers within the same row and can split the extracted columns into separate logical tables.

### 5. Vertically Stacked Table Splitting

Some sheets contain multiple equipment tables stacked vertically.

Interior title bands can be used to split a tall extracted grid into separate equipment blocks.

### 6. Region Carving

If a page appears to contain fused tables, title locations can be used to carve the drawing into smaller extraction regions.

Each region is then processed independently.

### 7. Alternative Text Strategies

If the default table extraction is weak or incomplete, the program can retry with text-driven boundary strategies.

### 8. Raw Rule-Line Reconstruction

For difficult dense tables, the extractor can inspect the actual horizontal and vertical drawing lines and rebuild the grid from those geometric boundaries.

Words are then assigned to reconstructed cells.

This serves as a fallback when the PDF library's normal cell segmentation produces scrambled results.

### 9. Quality Gate

A candidate is rejected when too few rows resemble actual points rows.

The current table-quality gate requires:

- At least two points-like rows
- At least 25% of candidate data rows to resemble points rows

Rejected candidates are reported rather than silently exported.

### 10. Continuation Handling

Tables with the same title appearing on later pages can be treated as continuation rows and appended to the same logical equipment table.

---

## Submittal and BACnet Object Processing

The optional submittal-processing engine adds another layer of automation.

### BACnet Object Table Detection

If a submittal includes BACnet object rows using recognizable object patterns, the application extracts structured records including:

- Object/tag information
- BACnet object type
- Object instance/address
- Description
- Units
- State text
- Source page

Supported BACnet-type mappings include:

| BACnet Object | Points-List Type |
|---|---|
| Analog Input | AI |
| Analog Output | AO |
| Analog Value | AV |
| Binary Input | DI |
| Binary Output | DO |
| Binary Value | BV |
| Multi-State Input | MSI |
| Multi-State Output | MSO |
| Multi-State Value | MSV |

### Description Normalization

Vendor terminology and drawing terminology are not always identical.

The matching engine normalizes several common concepts so equivalent descriptions have a better chance of matching.

Examples include concepts such as:

```text
Temperature <-> Temp
Setpoint    <-> SP
Command     <-> Cmd
Position    <-> Pos
Pressure    <-> Press
Differential Pressure / DP <-> Diff
Humidity / RH <-> Hum
Cooling / Chilled / CHW <-> Clg
```

### Meaning-Changing Word Protection

Matching is intentionally conservative.

Some extra words materially change what a point means.

For example:

```text
General Alarm
```

should not automatically match:

```text
General Alarm Delay 1
```

The matcher therefore rejects candidate descriptions that introduce certain meaning-changing concepts not present in the drawing point.

Examples include terms related to:

- delay
- setpoint
- rate
- runtime
- limits
- threshold
- restart
- minimum/maximum
- failure
- speed

### Numbered Equipment Families

A drawing may contain one generic point:

```text
Supply Fan Status
```

while the submittal contains:

```text
Supply Fan 1 Status
Supply Fan 2 Status
Supply Fan 3 Status
Supply Fan 4 Status
```

The application can identify these numbered families and propose expanding the single drawing row into multiple numbered rows.

Each new row can retain the corresponding submittal object's:

- Tag
- Type
- Address
- Units
- Source page

### Fallback Text Quantity Detection

If a structured BACnet object table is not found, the application includes a more general text-based expansion scan.

The fallback searches submittal text for point-description keywords and explicit quantity evidence such as:

```text
Qty: 4
Quantity 4
(4)
4 x ...
```

Only detected quantities of two or greater are proposed for expansion.

---

## Human-in-the-Loop Verification

The tool intentionally does not make submittal-driven changes automatically.

When possible, proposed changes show:

- Point description
- Proposed quantity
- Source submittal
- Source page
- Supporting text/evidence

The user can:

- Approve a proposal
- Reject a proposal
- Select all proposals
- Clear all proposals
- Adjust the detected count
- Apply only selected changes

There is also a manual expansion option for cases where the engineer already knows the required quantity.

This workflow is a core safety feature of the application.

The intended model is:

```text
Automation finds and prepares
        +
Engineer verifies and approves
        =
Faster but auditable workflow
```

---

## TATSU Mapping Logic

The TATSU export does more than copy text into a new workbook.

Drawing columns are resolved by header name, which makes the mapper less dependent on one fixed drawing-column order.

The application then deterministically translates drawing values into the 45-column TATSU structure.

Examples include:

### Hardwired I/O

Detected drawing I/O columns can map to:

- DI
- DO
- AI
- AO

and mark the point as hardwired.

### Integrated Points

Read/write indicators can map to:

```text
R
R/W
```

in the integrated field.

### MQTT

Drawing MQTT indicators can populate the MQTT field.

### Point-Type Inference

If a drawing does not explicitly provide a point type, the mapper can infer a likely type using:

- Units
- Read/write behavior
- Trending behavior
- Description keywords

Explicit drawing information and matched submittal object information take precedence over inferred defaults.

### Submittal Fills

When a verified submittal match exists, the TATSU row can receive:

- Point tag
- BACnet type
- Address/object instance
- Units

Blank source values are intentionally left blank when the source documents do not define them.

---

## Alarm and Trending Translation

### Poll Intervals

Numeric drawing trend intervals are interpreted as minutes and converted to seconds for the TATSU standard.

Example:

```text
Drawing trend interval: 5
TATSU interval:         300 seconds
```

`COS` and `COV` drawing values are mapped to COV trending.

### Alarm Priority Mapping

The application includes deterministic mapping logic between drawing alarm-priority conventions and the TATSU/Ignition notification structure.

The implemented correlation includes:

```text
1 - CRITICAL   -> Critical
2 - WARNING    -> High
3 - INFO       -> Low
4 - CONTROLLER -> Diagnostic / Maintenance-level mapping
```

It also supports drawing sets that use direct:

```text
CRITICAL
HIGH
MEDIUM
LOW
```

priority columns.

### Alarm Limits

The application can inspect setpoint/alarm bar information from the drawing sheet and attempt to replace generic low/high alarm placeholders with parsed alarm values when a suitable matching bar is found.

---

## Excel Output

### TATSU Points List Format

The standardized TATSU output includes:

- SmartTech-style navy/orange banner
- SmartTech logo when available
- Standardized TATSU column names
- Grouped headers
- Fixed column widths
- Rotated narrow-column labels
- Segoe UI based data styling
- Hidden Excel gridlines
- Freeze panes
- Source notes
- Expansion verification notes
- Submittal fill verification notes

Rows created through point expansion are visually differentiated using green text in the TATSU export.

### Raw / As-Extracted Format

The raw export retains a layout closer to the source drawing table.

It includes:

- Original extracted title
- Extracted header rows
- Extracted data rows
- Captured drawing notes
- Frozen panes
- Equipment-specific tabs
- Highlighted expanded rows
- Extraction Log worksheet

Expanded rows are highlighted in yellow in the raw workbook so they can be quickly reviewed.

### Extraction Log

The raw export includes an `Extraction Log` sheet that records information such as:

- Equipment sheet
- Source drawing file
- Source drawing pages
- Number of points
- Expansions applied

This provides traceability back to the drawing source.

---

## System Requirements

The documented deployment workflow is designed for Windows.

### Required

- Windows computer
- Python 3.10 or newer
- Modern web browser
- Access to the application files
- IFC drawing PDFs with selectable text
- Internet access during initial dependency installation, unless packages are already available internally

### Recommended

When installing Python on Windows, enable:

```text
Add python.exe to PATH
```

during installation.

---

## Installation

### 1. Clone or Download the Repository

Using Git:

```powershell
git clone <YOUR-REPOSITORY-URL>
cd <YOUR-REPOSITORY-FOLDER>
```

Or download the repository as a ZIP and extract it to a local folder.

For best performance, use a local folder rather than a synchronized/network location when practical.

Example:

```text
C:\Users\YourName\PointsTool
```

### 2. Confirm Python

Open PowerShell:

```powershell
py --version
```

or:

```powershell
python --version
```

Python 3.10 or newer is recommended by the current user guide.

### 3. Install Dependencies

From the project folder:

```powershell
py -m pip install -r requirements_v3.txt
```

If the `py` launcher is unavailable:

```powershell
python -m pip install -r requirements_v3.txt
```

### 4. Dependency Installation on Restricted Networks

If installation fails with an SSL or proxy-related error on a corporate network, package access may be blocked.

The current application depends on Python packages distributed through PyPI, so the environment must allow installation of the required packages or provide them through an approved internal package source.

---

## Running the Application

From PowerShell, navigate to the project folder:

```powershell
cd C:\Users\YourName\PointsTool
```

Start Streamlit:

```powershell
py -m streamlit run app_v3.py
```

Alternative commands:

```powershell
streamlit run app_v3.py
```

or:

```powershell
python -m streamlit run app_v3.py
```

The browser should open automatically.

The default local Streamlit address is:

```text
http://localhost:8501
```

Keep the PowerShell window open while the application is running.

To stop the application:

```text
Ctrl + C
```

in the PowerShell window.

---

## Step-by-Step Usage

### Step 1 — Upload IFC Drawings

1. Open the application.
2. Leave detection mode on **Auto-detect points tables** unless you have a specific reason to use marker mode.
3. Upload one or more IFC drawing PDFs.
4. Click **Scan drawings**.
5. Review the Detection Report.
6. Review Errors & Omissions.
7. Open each extracted equipment table.
8. Verify the source pages.
9. Rename the Excel tab if necessary.
10. Correct any extracted cells before continuing.

### Step 2 — Add Submittals

1. Upload the relevant equipment submittal PDFs.
2. Confirm that each submittal is mapped to the correct equipment table.
3. Click **Scan submittals for additional points**.
4. Review every proposed expansion or fill.
5. Check the source page and supporting evidence.
6. Adjust quantities if required.
7. Untick anything that is not correct.
8. Click **Apply ticked changes**.

### Step 3 — Export

Choose:

```text
TATSU Points List
```

or:

```text
As extracted from drawing
```

Then click:

```text
Download points list workbook
```

Open the resulting Excel file and perform final engineering verification.

---

## Input Requirements

### IFC Drawings

Best results are expected when:

- The PDF contains selectable text
- Points tables use drawn grid lines
- The table contains a recognizable `POINT DESCRIPTION` column
- Text is not rasterized into an image
- Equipment tables follow normal BAS drawing conventions

### Submittals

Submittal enrichment works best when:

- Text is selectable
- BACnet object lists are text-based
- Object descriptions are clearly structured
- Object instance numbers are present
- Equipment families use consistent numbering

---

## Output Files

The application produces an `.xlsx` workbook.

Default filename:

```text
Points_List_Extracted.xlsx
```

Depending on export mode, the workbook may contain:

- One worksheet per equipment type
- Structured TATSU columns
- Raw extracted drawing tables
- Source notes
- Expansion notes
- Submittal-derived fields
- Verification highlighting
- Extraction Log

---

## Troubleshooting

### `streamlit is not recognized`

Try:

```powershell
py -m streamlit run app_v3.py
```

### `py is not recognized`

Python may not have been added to the Windows PATH.

Options include:

- Reinstall Python and enable **Add python.exe to PATH**
- Use `python` instead of `py`
- Use the full path to `python.exe`

### Browser Does Not Open Automatically

Open:

```text
http://localhost:8501
```

manually in a browser.

### Browser Shows an Error or Blank Page

Check the PowerShell window.

The terminal usually contains the underlying Python or Streamlit error.

Stop the application with:

```text
Ctrl + C
```

and restart it after resolving the issue.

### A Table Did Not Extract

Check the **Errors & Omissions** section.

Possible causes include:

- The PDF is image-only
- The page does not contain selectable `POINT DESCRIPTION` text
- The table grid is unusually formatted
- The table was rejected by the quality gate
- The selected marker does not match the actual title

Try auto-detect mode if marker mode was being used.

### A Table Extracted Under the Wrong Name

If the title could not be read, the application may assign a placeholder such as:

```text
Points Table (pX.Y)
```

Rename the Excel tab in the table editor before export.

### A Submittal Found No Matches

Confirm that the submittal is mapped to the correct equipment table.

A CRAH submittal should not be mapped to an unrelated FWU table, for example.

Also confirm that the submittal contains readable text.

### Excel Font Looks Different

The TATSU formatting references fonts such as:

- Michroma
- Segoe UI
- Segoe UI Semibold

If those fonts are not installed, Excel may substitute another font.

The workbook data remains valid; only the appearance may differ.

---

## Data Privacy and Security

A central design decision of this application is local processing.

The current implementation makes no external document-processing calls.

Uploaded project files are processed by the locally running Streamlit application on the user's computer.

This is especially useful for engineering projects where drawing sets and vendor submittals may contain confidential or client-controlled information.

The local workflow means:

```text
Project PDF
   |
   v
Local Python Process
   |
   v
Local Excel Output
```

rather than:

```text
Project PDF
   |
   v
External document-processing API
```

Normal corporate security policies still apply to the computer, repository, source files, and generated output.

---

## Known Limitations

### Image-Only / Scanned PDFs

The current version does not include an OCR pipeline for scanned image-only drawings.

If a PDF does not contain readable/selectable text, the extractor may be unable to identify the points table or submittal content.

### Nonstandard Drawing Geometry

BAS drawing standards vary between design teams and projects.

The extractor includes several fallback strategies, but an unusual table layout may still require additional extraction logic.

### Submittal Format Variability

BACnet object-table processing is designed around recognizable structured object rows.

Vendor formats that differ substantially may not be parsed automatically.

### Matching Is Heuristic but Conservative

Description matching uses deterministic word normalization and similarity rules.

It is designed to reduce false matches, but no text-matching system should be treated as engineering authority.

Every submittal-driven result should be verified.

### No Universal Accuracy Percentage

The current project files do not define a formal benchmark dataset or validated universal extraction-accuracy percentage.

Accuracy can vary with:

- PDF generation method
- Table geometry
- Drawing standard
- Font/text encoding
- Submittal layout
- Vendor terminology

A formal validation study across multiple projects would be required before claiming a statistically meaningful accuracy rate.

### Current Deployment Requires Python

The documented workflow requires a local Python environment and Streamlit.

The application is not currently packaged as a standalone Windows executable in the supplied version.

---

## Engineering Verification

This tool is intended to accelerate engineering work, not eliminate engineering review.

Before issuing a final deliverable, the responsible engineer should verify:

- All expected equipment tables were extracted
- No equipment table is missing
- Point descriptions are correct
- I/O types are correct
- Integrated/read-write behavior is correct
- Trend types and intervals are correct
- Alarm priorities are correct
- Alarm limits are correct
- Expanded quantities are correct
- BACnet tags are correct
- BACnet object types are correct
- BACnet addresses are correct
- Units are correct
- Submittals were mapped to the correct equipment
- The final workbook matches project requirements

The application deliberately highlights or annotates changes that deserve additional attention.

---

## Project Structure

A minimal repository for the current version can be organized as:

```text
SmartTech-Points-List-Extractor/
|
|-- app_v3.py
|-- requirements_v3.txt
|-- Points_Tool_User_Guide.docx
|-- README.md
```

### `app_v3.py`

Main application.

Contains:

- Streamlit interface
- PDF extraction logic
- Table validation
- Drawing-title handling
- Setpoint/alarm parsing
- Submittal scanning
- BACnet object parsing
- Point matching
- Row expansion
- TATSU translation
- Excel workbook generation
- SmartTech styling

### `requirements_v3.txt`

Python dependency list.

### `Points_Tool_User_Guide.docx`

End-user installation and operating instructions.

### `README.md`

Repository-level project documentation.

---

## Dependencies

The supplied requirements file includes:

```text
streamlit>=1.32
pandas>=2.0
openpyxl>=3.1
pymupdf>=1.24
pypdf>=4.0
pillow>=10.0
```

### Streamlit

Provides the local browser-based user interface.

### pandas

Used for structured table display and editing within the application.

### PyMuPDF (`fitz`)

Primary PDF parsing and document-geometry engine.

Used for tasks including:

- Text extraction
- Word geometry
- Drawing geometry
- Table detection
- PDF page processing

### openpyxl

Creates the final Excel workbook and applies:

- Cell values
- Formatting
- Borders
- Fonts
- Column widths
- Freeze panes
- Logos
- Worksheet structure

### pypdf

Included as a PDF-related dependency in the supplied requirements.

### Pillow

Supports image handling used by workbook branding/logo insertion.

---

## Technical Design Notes

### Deterministic Processing

The current version is primarily deterministic automation rather than a generative-AI workflow.

This is intentional.

For engineering-document transformation, deterministic rules provide several advantages:

- Repeatable behavior
- Easier debugging
- Clear source-to-output relationships
- No prompt variability
- No external LLM dependency
- Local processing
- Easier engineering verification

### Session State

Streamlit session state is used to retain:

- Extracted entries
- Detection reports
- Submittal proposals

Rescanning drawings resets downstream proposal state so changes from a previous drawing scan are not accidentally carried into a new extraction.

### Sheet Name Sanitization

Excel worksheet names are sanitized to:

- Remove invalid characters
- Respect the Excel 31-character limit
- Avoid duplicate worksheet names

### Source Traceability

Equipment entries retain source metadata such as:

- Drawing filename
- Drawing pages
- Expansion source
- Submittal page

This information is used in the workbook and verification notes.

### Conservative Failure Behavior

When the program cannot confidently complete an operation, the preferred behavior is to:

- Skip the questionable item
- Report the problem
- Ask for engineer review

rather than silently manufacturing data.

---

## Potential Future Improvements

The following items are logical future-development opportunities. They are not claimed as features of the current supplied version.

### OCR Support

Add OCR for:

- Scanned drawings
- Image-only submittals
- Rasterized object schedules

### Formal Validation Suite

Create a test library of representative BAS drawing sets and measure:

- Table detection recall
- Table false-positive rate
- Cell-level extraction accuracy
- Submittal matching precision
- Expansion accuracy
- TATSU mapping accuracy

### Automated Regression Testing

Add unit and integration tests for:

- Title detection
- Table splitting
- Sheet sanitization
- Point-row validation
- BACnet parsing
- Description matching
- Alarm mapping
- Poll conversion
- Workbook generation

### Standalone Windows Packaging

Package the application so users do not need to manually install Python.

Potential approaches could include a managed installer or executable packaging workflow.

### Configurable Project Standards

Move hard-coded project mappings into configuration files so different clients or projects can define:

- Column standards
- Alarm priorities
- Trend rules
- Branding
- Export templates

### Broader Vendor Submittal Support

Create vendor-specific parsers for additional BACnet/object-table formats.

### Improved Logging

Add structured processing logs that can be saved with the output package.

### Modular Codebase

As the project grows, the single application file could be separated into modules such as:

```text
extractor/
    drawings.py
    tables.py
    validation.py

submittals/
    bacnet.py
    matching.py

exports/
    raw_excel.py
    tatsu.py

ui/
    streamlit_app.py
```

This would make testing and future maintenance easier.

---

## Design Philosophy

The project follows a simple principle:

> Automate repetitive document processing, but keep engineering judgment with the engineer.

The tool is most valuable when it performs deterministic, high-volume work while clearly identifying what still requires human review.

That means the application intentionally favors:

- Traceability over hidden automation
- Visible errors over silent failures
- Source evidence over unsupported inference
- Engineer approval over automatic acceptance
- Local processing over unnecessary external data transfer
- Repeatable rules over unpredictable output

---

## Author

**Vishan Vyas**

Originally developed as an engineering workflow-efficiency tool for BAS/BMS points-list creation and project-document processing.

---

## License

No software license is defined in the supplied project files.

If this repository will be made public, add an appropriate `LICENSE` file before allowing reuse, modification, or redistribution.

If the code or branding is company-owned or intended only for internal use, confirm the appropriate repository visibility and licensing terms before publishing.

---

## Disclaimer

This application is an engineering productivity tool.

Generated points lists should be reviewed against the governing project drawings, approved equipment submittals, controls requirements, sequences of operation, and applicable project standards before being used as a final engineering or commissioning deliverable.

The source documents remain the controlling reference.
