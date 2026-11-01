# MeMind

React + TypeScript + Vite application.

## Project Structure

- `src/` — React frontend components, services, types, styles
- `engine/` — Core Excel Engine (Python + openpyxl)
  - `excel_engine.py` — MIND-010 Core Excel Engine implementation
  - `file_enhancer.py` — MIND-020 File Enhancer (deterministic analysis + safe enhancement)
  - `formula_engine.py` — MIND-030 Formula Engine (deterministic formula detection, validation, generation, application)
  - `create_fixture.py` — Sample workbook generator
  - `__init__.py` — Package init
- `tests/` — Automated tests
  - `test_engine.py` — 23 tests covering the Core Excel Engine (MIND-010)
  - `test_file_enhancer.py` — 40 tests covering the File Enhancer (MIND-020)
  - `test_formula_engine.py` — 57 tests covering the Formula Engine (MIND-030)
- `public/` — Static assets (test.xlsx, icons, HTML)
- `fixtures/` — Test fixtures (sample.xlsx + analysis fixtures)

## Core Excel Engine (MIND-010)

The Core Excel Engine is implemented in Python using `openpyxl`.

### Capabilities

- **Read**: Load `.xlsx`, detect sheets, read cells, headers, dimensions
- **Modify**: Rename/create/delete sheets, add/delete rows/columns, update cells
- **Write**: Save as `.xlsx` file
- **Validate**: Re-open output, verify sheet count and structure match
- **Pipeline**: `load → modify → save → validate` in one call

### Running MIND-010 Tests

```bash
cd /data/data/com.termux/files/home/memind
python3 -m pytest tests/test_engine.py -v
```

### MIND-010 Test Count: 23 tests (all pass)

## File Enhancer (MIND-020)

The File Enhancer is a deterministic spreadsheet quality and safe enhancement layer
that sits on top of MIND-010.

### Capabilities

**DETECTION (detect-only):**

- Duplicate rows
- Missing values
- Type inconsistency
- Header quality issues (empty, whitespace-only, surrounding whitespace, duplicate)
- Formula presence

**SAFE AUTO-FIX (enhancement):**

- Text whitespace trimming on ordinary data cells only
- Headers are NOT trimmed (Section 13)
- Formulas are NOT modified (Section 14)
- Numeric, boolean, date values are NOT modified (Section 7.1)

**EXPLICIT INSTRUCTION (NOT automatic):**

- Duplicate deletion
- Missing-value filling
- Type conversion
- Header renaming / normalization
- Formula repair / generation
- Semantic data correction

### Running MIND-020 Tests

```bash
cd /data/data/com.termux/files/home/memind
python3 -m pytest tests/test_file_enhancer.py -v
```

### MIND-020 Test Count: 40 tests (all pass)

### Total Test Count: 120 tests (MIND-010 + MIND-020 + MIND-030)

### Architecture

```
User
↓
MeMind Application
↓
MIND-020 File Enhancer  (file_enhancer.py)
↓
MIND-010 Core Excel Engine  (excel_engine.py)
↓
.xlsx
```

### Contract

- **ANALYZE**: Reads workbook, detects issues, returns `AnalysisResult`. Does NOT modify input.
- **ENHANCE**: Loads workbook, analyzes, applies safe whitespace trimming only, saves, reopens, validates. Returns `EnhancementResult`.

### Issue Model

All detected issues are represented as `EnhancementIssue` with:

- `type` — issue category
- `severity` — INFO / WARNING / ERROR
- `sheet` — sheet name
- `location` — human-readable location
- `description` — issue description
- `current_value` — current cell value (if applicable)
- `suggested_action` — recommended manual action

### Real XLSX Validation

All tests use real `.xlsx` fixtures. Round-trip validation (load → analyze → enhance → save → reopen → validate) is verified.

## Formula Engine (MIND-030)

The Formula Engine is a deterministic Excel formula layer that sits on top of MIND-010.

### Capabilities

**DETECTION:**
- Scans all sheets for cells containing formulas
- Returns per-cell metadata: sheet, cell address, row, col, formula text, formula type, extracted references
- Deterministic — same input always produces identical results
- Excludes non-formula cells
- Does NOT evaluate formulas

**VALIDATION (structural only):**
- Checks formula starts with `=`
- Checks balanced parentheses
- Warns on unknown function names
- Detects dangling operators
- Does NOT calculate or verify computed results
- Does NOT claim full Excel calculation support

**GENERATION (structured, not natural language):**
- SUM, AVERAGE, MIN, MAX, COUNT, COUNTA — range-based
- ADD, SUBTRACT, MULTIPLY, DIVIDE — binary cell references
- IF — condition + true/false values
- Rejects unsupported operations
- Rejects missing required operands

**APPLICATION:**
- Explicit target cell + structured operation
- Existing formulas NOT overwritten by default
- Explicit overwrite policy required for replacement
- Protected sheets rejected
- Missing sheets rejected
- Invalid cell references rejected
- Non-target cells preserved
- Workbook structure preserved
- Formatting preserved

**PERSISTENCE:**
- Write → Save → Reopen → Verify
- Exact formula string verified after reopen
- Save/reopen validation via MIND-010

### Running MIND-030 Tests

```bash
cd /data/data/com.termux/files/home/memind
python3 -m pytest tests/test_formula_engine.py -v
```

### MIND-030 Test Count: 57 tests (all pass)

### Architecture

```
User / MeMind Application
↓
MIND-030 Formula Engine  (formula_engine.py)
↓  (public API only)
MIND-020 File Enhancer  (file_enhancer.py)
↓  (public API only)
MIND-010 Core Excel Engine  (excel_engine.py)
↓
.xlsx
```

### Data Model

- `FormulaInfo` — sheet, cell, row, col, formula, formula_type, references
- `FormulaValidationResult` — valid, sheet, cell, formula, errors, warnings
- `FormulaOperation` — target, operation, operands, overwrite_policy, sheet

### Known Limitations

- Does NOT implement a calculation engine — formula results are not computed or verified
- Does NOT implement a full Excel formula parser — only structural checks within declared scope
- Does NOT support AI/LLM formula generation
- Does NOT support natural-language formula interpretation
- Supported operations are deliberately limited (see Capabilities above)
- SheetJS (frontend) and openpyxl (backend) parse formulas independently — no bridge exists

### Safety Contract

- Existing formulas are NOT overwritten without explicit `overwrite_policy='explicit'`
- Protected sheets are NOT forcibly modified — rejected with error
- Formatting, merged cells, charts, hidden rows/columns are NOT modified
- Non-target cells are preserved
- Workbook structure is preserved
- If an operation is unsafe or unsupported: REJECT / REPORT

## Frontend

- React + TypeScript + Vite
- `xlsx` (SheetJS) library for browser-side spreadsheet operations
- `spreadsheetService.ts` handles all file parsing/export
- Components: Spreadsheet, FileDropzone, SpreadsheetToolbar, SpreadsheetGrid, SheetTabs, WorkbookStatus, CellEditor

### Running Dev Server

```bash
npm run dev
```

### Build

```bash
npm run build
```

## Known Limitations

### MIND-010

- **Supported**: `.xlsx` only
- **Preserved (expected)**: Formulas, merged cells, and formatting are preserved via openpyxl round-trip for unmodified features
- **Not guaranteed**: Advanced charts, external links, complex Excel-specific features, full workbook metadata

### MIND-020

- Protected sheet detection uses public `is_sheet_protected()` API (MIND-010). Protected sheets are detected and reported but not modified.
- Full Excel fidelity preservation is NOT claimed. Only verified behaviors are listed.
- SheetJS (frontend) and openpyxl (backend) parse independently — no bridge exists between them.

### MIND-030

- Does NOT implement a calculation engine — formula results are not computed or verified
- Does NOT implement a full Excel formula parser — only structural checks within declared scope
- Does NOT support AI/LLM formula generation
- Does NOT support natural-language formula interpretation
- Supported operations are deliberately limited (see Capabilities above)
- SheetJS (frontend) and openpyxl (backend) parse formulas independently — no bridge exists

## Data Region

Non-empty bounding box per worksheet.

Bedakan:

physical dimensions

dengan:

data region

Header Detection

Tiga status:

DETECTED
NOT_DETECTED
AMBIGUOUS

Rule-based dan deterministic.

Column Profiling

Frozen taxonomy:

EMPTY
TEXT
INTEGER
FLOAT
BOOLEAN
DATE/DATETIME
FORMULA
MIXED
UNKNOWN

Termasuk statistics/distribution yang memang tersedia di implementation.

Formula Profile

Berasal dari public MIND-030 API.

Tidak ada second formula parser.

Quality Aggregation

Berasal dari public MIND-020 analysis.

MIND-050 tidak mengambil alih authority quality detection.

Pattern Detection

Hanya:

SEQUENTIAL_NUMERIC
CONSTANT_VALUE
REPEATED_VALUE
DATE_SEQUENCE
FORMULA_RELATIVE_PATTERN
DOMINANT_VALUE

DATE_SEQUENCE:

DAILY
WEEKLY

saja.

Relationship Signals

MIND-050 hanya menyatakan:

CROSS_SHEET_REFERENCE

Jangan menyatakan cross-sheet reference sebagai key relationship secara otomatis.

Insights

Lima taxonomy:

STRUCTURAL
DATA_QUALITY
FORMULA
PATTERN
RELATIONSHIP

Dengan structured evidence/rule/severity.

Summary

Physical + data dimensions serta formula/quality/pattern/insight counters.

Serialization

"to_dict()" / JSON-compatible representation.

Safety

MIND-050 read-only.

Tidak mengubah workbook sumber.

Determinism

Input yang sama menghasilkan analysis yang sama.

Explicit Limitations

Dokumentasikan minimal:

- no multiple-table detection;
- no monthly/quarterly date semantics;
- no semantic interpretation;
- no formula calculation;
- no formula generation/application;
- no AI/LLM.

Jangan membuat capability claim yang tidak ada.

---

## Spreadsheet Intelligence (MIND-050)

MIND-050 adalah lapisan deterministic spreadsheet structural intelligence yang di atas MIND-010, MIND-020, dan MIND-030. MIND-050 melakukan analisis structured workbook tanpa memodifikasi input.

### Purpose

MIND-050 menghasilkan SpreadsheetAnalysis yang mengandung:

- Workbook structure (sheet names, counts)
- Per-sheet data region (non-empty bounding box)
- Header detection (DETECTED / NOT_DETECTED / AMBIGUOUS, rule-based)
- Column profiling (9-type taxonomy + statistics/distribution)
- Formula profile (from MIND-030 public API)
- Quality signals (from MIND-020 public API)
- Pattern detection (6 deterministic pattern types)
- Structural relationships (CROSS_SHEET_REFERENCE only, not key inference)
- Insights (5 taxonomy types, structured evidence/rule/severity)
- Summary (physical + data dimensions + counters)

Prinsip: "Understand more, infer less." Setiap insight harus traceable ke observable evidence dan deterministic rule.

### Input

- Format: `.xlsx` only
- Primary API: `analyze(input_path: Path) -> SpreadsheetAnalysis`

### Pipeline

```
LOAD
  → WORKBOOK STRUCTURE
  → DATA REGION (non-empty bounding box)
  → HEADER DETECTION (rule-based, 3 statuses)
  → COLUMN PROFILING (9-type taxonomy + stats)
  → FORMULA PROFILE (from MIND-030)
  → QUALITY AGGREGATION (from MIND-020)
  → PATTERN DETECTION (6 frozen types)
  → RELATIONSHIP SIGNALS (CROSS_SHEET_REFERENCE only)
  → INSIGHTS (5 taxonomy types)
  → SUMMARY (physical + data counters)
```

### Data Region

Per worksheet, MIND-050 mendeteksi rectangular bounding box dari seluruh non-empty cells (value bukan None dan bukan empty string).

Dibedakan dengan jelas antara:

- `physical_dimensions` — max_row × max_column dari worksheet (bisa termasuk empty rows/columns)
- `data_region` — min_row, max_row, min_col, max_col dari non-empty cells

### Header Detection

Tiga status deterministic:

- `DETECTED` — text density >= 0.5 AND uniqueness >= 0.5
- `NOT_DETECTED` — text density < 0.3
- `AMBIGUOUS` — intermediate (tidak memenuhi DETECTED atau NOT_DETECTED)

Header detection adalah rule-based, bukan semantic. MIND-050 tidak menebak apakah suatu row "merupakan" header.

### Column Profiling

Setiap kolom dalam data region diklasifikasikan ke dalam 9 tipe:

- `EMPTY` — no non-empty cells
- `TEXT` — all text values
- `INTEGER` — all integer values
- `FLOAT` — all float values
- `BOOLEAN` — all boolean values
- `DATE/DATETIME` — all date/datetime values
- `FORMULA` — all formula cells
- `MIXED` — multiple types present
- `UNKNOWN` — unrecognized type

Dilengkapi dengan:

- `non_empty_count` / `empty_count`
- `unique_count` / `duplicate_count`
- `formula_count`
- `type_distribution` (dict)
- `numeric_stats` (min/max/sum/mean/sample_count) — jika numeric
- `date_stats` (earliest/latest/sample_count) — jika date
- `value_distribution` (dominant_value, top_values) — jika text

### Formula Profile

Formula profile diambil dari public API MIND-030 (`FormulaEngine.scan_formulas()`). MIND-050 TIDAK membuat formula parser kedua.

Termasuk:

- `formula_count`
- `formula_cells` (cell, row, col, formula_type, references)
- `formula_distribution` (by formula_type)
- `has_cross_sheet_references` / `cross_sheet_references`
- `has_unknown_formulas`

### Quality Aggregation

Quality signals di-aggregate dari public API MIND-020 (`FileEnhancer.analyze()` → `AnalysisResult`). MIND-050 TIDAK membuat detector kualitas sendiri.

Signals yang di-mapped:

- `duplicate_rows`
- `missing_values`
- `type_inconsistency`
- `header_quality:*`
- `formula_detected`

### Pattern Detection

6 frozen pattern types, masing-masing dengan detection rule deterministic:

- `SEQUENTIAL_NUMERIC` — arithmetic progression dengan constant step
- `CONSTANT_VALUE` — semua non-empty cells sama
- `REPEATED_VALUE` — siklis dengan period >= 2
- `DATE_SEQUENCE` — hanya DAILY (step=1) atau WEEKLY (step=7)
- `FORMULA_RELATIVE_PATTERN` — formula dengan struktur relatif serupa
- `DOMINANT_VALUE` — satu value mendominasi > 50% non-empty cells

### Relationship Signals

MIND-050 hanya mendeteksi `CROSS_SHEET_REFERENCE` sebagai factual signal dari formula yang mengandung reference ke sheet lain (format `SheetName!Cell`).

MIND-050 TIDAK menyimpulkan:

- "Customers adalah master table"
- Foreign key relationships
- Kedalaman relasi antar-sheet

Jika formula `Sales!A1 = SUM(Customers!B2:B100)`, MIND-050 hanya melaporkan: "Sales!A1 references Customers!B2:B100" — bukan "Customers adalah master data untuk Sales".

### Insights

5 taxonomy insight types:

- `STRUCTURAL` — data region, header, trailing empty rows
- `DATA_QUALITY` — duplicate rows, missing values, type inconsistency, header quality
- `FORMULA` — formula count, cross-sheet references, unknown formulas
- `PATTERN` — detected patterns (sequential, constant, repeated, date sequence, formula pattern, dominant)
- `RELATIONSHIP` — cross-sheet references

Setiap insight memiliki:

- `type`, `severity` (INFO/WARNING/ERROR), `sheet`, `location`, `description`
- `evidence` (list), `rule` (string), `suggested_action`
- `source` (MIND-050 / MIND-020 / MIND-030)

### Summary

Workbook summary mencakup:

- `file_name`, `sheet_count`
- `total_physical_rows`, `total_physical_columns` (jumlah dari seluruh worksheet dimensions)
- `total_data_rows`, `total_data_columns` (jumlah dari data region non-empty)
- `total_formula_cells`, `total_quality_signals`, `pattern_count`, `insight_count`
- `sheet_summaries[]` — per-sheet breakdown
- `quality_summary` — aggregated by type dan severity

### Serialization

Semua domain models adalah dataclass dengan `to_dict()` method. `SpreadsheetAnalysis` juga memiliki `to_json_compatible()` untuk JSON serialization.

### Safety

MIND-050 adalah READ-ONLY. Implementation tidak:

- memodifikasi cell
- mengubah workbook structure
- menyimpan file output
- mengubah formula
- mengubah formatting

Source workbook hash (MD5) diverifikasi unchanged setelah analysis.

### Determinism

MIND-050 adalah deterministic. Untuk workbook yang sama, dengan aturan yang sama, dan implementasi yang sama, hasil analysis selalu identik.

Tidak ada:

- random values
- arbitrary confidence scores
- time-dependent results
- external API dependencies
- LLM output

### Limitations

MIND-050 TIDAK:

- mendeteksi multiple tables dalam satu sheet
- mendeteksi monthly/quarterly/annual date semantics (hanya DAILY/WEEKLY)
- melakukan semantic interpretation (bukan AI)
- menghitung formula (bukan calculation engine)
- membuat atau menerapkan formula baru (bukan formula generator)
- menggunakan AI/LLM
- melakukan data enhancement atau auto-fix

MIND-050 HANYA:

- mendeteksi struktur, pola, kualitas, dan hubungan secara deterministic
- melaporkan temuan dengan evidence dan rule yang dapat di-trace

### Running MIND-050 Tests

```bash
cd /data/data/com.termux/files/home/memind
python3 -m pytest tests/test_spreadsheet_intelligence.py -v
```

### MIND-050 Test Count: 63 tests (all pass)

### Combined Test Count: 240 tests (MIND-010 + MIND-020 + MIND-030 + MIND-040 + MIND-050)

---

## Auto Reporting (MIND-060)

MIND-060 is a deterministic structured report composer that transforms MIND-050's `SpreadsheetAnalysis` into a structured `Report`. It is NOT an intelligence engine, AI writer, or rendering engine. It composes and organizes evidence already discovered by previous milestones.

### Purpose

MIND-060 answers: «How can MeMind's structured spreadsheet intelligence be composed into a consistent, traceable, human-readable report?»

Pipeline:

```
XLSX
  ↓
MIND-050 Spreadsheet Intelligence
  ↓
SpreadsheetAnalysis
  ↓
MIND-060 Auto Reporting
  ↓
Report (structured, deterministic, serializable)
```

### Canonical API

```python
from engine.mind_060.auto_reporting import generate_report

report = generate_report(analysis)  # analysis: SpreadsheetAnalysis
```

- Input: `SpreadsheetAnalysis` from MIND-050 (already-analyzed, in-memory)
- Output: `Report` dataclass with 10 sections
- No file I/O, no workbook reopening, no analysis duplication

### Domain Model

```
Report
├── metadata (ReportMetadata)
├── executive_summary (ExecutiveSummary)
├── workbook_overview (WorkbookOverview)
├── sheet_reports[] (list[SheetReport])
├── quality_section (QualitySection)
├── formula_section (FormulaSection)
├── pattern_section (PatternSection)
├── relationship_section (RelationshipSection)
├── insight_section (InsightSection)
└── limitations (list[str])
```

### Report Sections

1. **Metadata**: report_type="auto_report", source_type="xlsx", source_identifier, analyzed_via, schema_version="1.0.0". No timestamps, no UUIDs (deterministic).

2. **Executive Summary**: sheet_count, sheets_with_data, sheets_with_formulas, sheets_with_quality_signals, total_formulas, total_quality_signals, total_patterns, total_insights, total_relationships. Only factual counters.

3. **Workbook Overview**: file_name, sheet_count, total_physical_rows, total_physical_columns, total_data_rows, total_data_columns. Directly from `WorkbookSummary`.

4. **Sheet Reports**: One per `SheetAnalysis`. Includes data_region, header_status, compact column summaries, formula info, quality signals, patterns, and filtered insights.

5. **Quality Section**: Aggregated from `analysis.quality_signals`. Total issues, by_severity (INFO/WARNING/ERROR), by_type, issue list with evidence preserved.

6. **Formula Section**: Aggregated from per-sheet `FormulaProfile`. Formula count, cells, distribution, cross-sheet references, unknown formula flag. Does NOT claim formula correctness.

7. **Pattern Section**: Aggregated from per-sheet `PatternDetection`. All 6 frozen types only: SEQUENTIAL_NUMERIC, CONSTANT_VALUE, REPEATED_VALUE, DATE_SEQUENCE, FORMULA_RELATIVE_PATTERN, DOMINANT_VALUE. No semantic reinterpretation.

8. **Relationship Section**: From `analysis.structural_relationships`. CROSS_SHEET_REFERENCE only — not promoted to foreign key/primary key/business relationship.

9. **Insight Section**: All MIND-050 insights preserved verbatim. Type, severity, sheet, location, description, evidence, rule, suggested_action, source, related_to.

10. **Limitations**: 7 verified statements about what MIND-060 does not do (no formula calculation, no semantic interpretation, no relationship inference, no multiple-table detection, DAILY/WEEKLY only date patterns, no independent analysis).

### Column Summary (Compact)

Per PO decision, only compact fields are included:

- index, column_letter, header, detected_type
- non_empty_count, empty_count, unique_count, duplicate_count, formula_count
- patterns (per-column)

Large statistical structures (numeric_stats, date_stats, value_distribution, type_distribution, notes) are NOT included.

### Serialization

- `Report.to_dict()` — deterministic dict representation
- `Report.to_json_compatible()` — JSON-safe dict (same as to_dict())

All serialization preserves evidence. No Path, datetime, Enum, or arbitrary Python objects in output.

### Determinism

Given the same `SpreadsheetAnalysis`, `generate_report()` produces equivalent `to_dict()` output on every call. No timestamps, no UUIDs, no random, no set iteration without sorting.

### Read-only Safety

MIND-060 is a pure transformation. It does not modify the input `SpreadsheetAnalysis` or any workbook. Verified by comparing `analysis.to_dict()` before and after `generate_report()`.

### Dependencies

```
MIND-060
  └── MIND-050 (SpreadsheetAnalysis)  [PRIMARY]
        ├── MIND-010  [transitive]
        ├── MIND-020  [transitive]
        └── MIND-030  [transitive]
```

MIND-040 is NOT a dependency. MIND-060 does NOT call MIND-020 or MIND-030 directly — it reads their aggregated output via MIND-050.

### Rendering Boundary

MIND-060 MVP produces ONLY structured `Report`. No Markdown, HTML, DOCX, PDF, PPTX, or UI rendering. Rendering is deferred to a future milestone.

### Limitations

1. Formula calculation is not performed.
2. Semantic business interpretation is not performed.
3. Unsupported Excel structures may not be fully analyzed.
4. Relationship signals do not imply confirmed database relationships.
5. Multiple tables within a single sheet are not detected.
6. Date sequence patterns limited to DAILY and WEEKLY.
7. This is a structured transformation, not independent analysis.

### Out of Scope

- LLM / AI-generated narrative
- Semantic business interpretation
- Formula calculation or verification
- Markdown / HTML / DOCX / PDF / PPTX rendering
- File I/O (reading XLSX, writing report files)
- `generate_report_from_path()` convenience wrapper
- Any analysis beyond reformatting MIND-050 output
- MIND-040 dependency

### Files

- `engine/mind_060/__init__.py` — package init
- `engine/mind_060/auto_reporting.py` — implementation
- `tests/test_auto_reporting.py` — 47 tests

### Running MIND-060 Tests

```bash
cd /data/data/com.termux/files/home/memind
python3 -m unittest tests.test_auto_reporting -v
```

Requires: Python environment with openpyxl importable and `memind/` on sys.path.

### MIND-060 Test Count: 47 tests (all pass)

### Combined Test Count: 287 tests (MIND-010 + MIND-020 + MIND-030 + MIND-040 + MIND-050 + MIND-060)

### Architecture

```
User / MeMind Application
  ↓
MIND-060 Auto Reporting  (engine/mind_060/auto_reporting.py)
  ↓  (consumes public models only)
MIND-050 Spreadsheet Intelligence  (engine/mind_050/spreadsheet_intelligence.py)
  ↓  (public API only)
MIND-040 Text Task Parser  (text_task_parser.py)
MIND-030 Formula Engine  (formula_engine.py)
MIND-020 File Enhancer  (file_enhancer.py)
MIND-010 Core Excel Engine  (excel_engine.py)
  ↓
.xlsx
```

### MIND-060 Development and Test Status

#### Test Results (latest)

```
MIND-010: 23/23 PASS
MIND-020: 40/40 PASS
MIND-030: 57/57 PASS
MIND-040: 57/57 PASS
MIND-050: 63/63 PASS
MIND-060: 47/47 PASS
Combined: 287/287 PASS
```

#### Contract Validation

- Input contract (SpreadsheetAnalysis only): PASS
- Output contract (Report with 10 sections): PASS
- Metadata (no timestamps, deterministic): PASS
- Executive summary (factual counters only): PASS
- Workbook overview (from WorkbookSummary): PASS
- Sheet reports (one per SheetAnalysis): PASS
- Compact column summaries: PASS
- Quality section (INFO/WARNING/ERROR preserved): PASS
- Formula section (aggregated, no correctness claims): PASS
- Pattern section (6 frozen types, no reinterpretation): PASS
- Relationship section (CROSS_SHEET_REFERENCE only): PASS
- Insight section (all fields preserved verbatim): PASS
- Limitations (7 verified statements): PASS
- Serialization (to_dict, to_json_compatible): PASS
- Determinism (identical on repeated calls): PASS
- Read-only (analysis unchanged): PASS
- No renderer (structured Report only): PASS
- No MIND-040 dependency: PASS
- No frozen milestone modification: PASS
- Private API usage: NONE
- Scope creep: NONE

#### Files

- `engine/mind_060/auto_reporting.py` — MIND-060 implementation
- `engine/mind_060/__init__.py` — package init
- `tests/test_auto_reporting.py` — 47 tests

---

## MIND-060 Development and Test Status

```
User / MeMind Application
  ↓
MIND-050 Spreadsheet Intelligence  (engine/mind_050/spreadsheet_intelligence.py)
  ↓  (public API only)
MIND-040 Text Task Parser  (text_task_parser.py)
MIND-030 Formula Engine  (formula_engine.py)
MIND-020 File Enhancer  (file_enhancer.py)
MIND-010 Core Excel Engine  (excel_engine.py)
  ↓
.xlsx
```

### MIND-050 Public API

```python
from engine.mind_050.spreadsheet_intelligence import SpreadsheetIntelligence

analyzer = SpreadsheetIntelligence()
analysis = analyzer.analyze(Path("workbook.xlsx"))
analyzer.close()

# Access results
print(analysis.summary.sheet_count)
print(analysis.summary.total_physical_rows)
print(analysis.summary.total_data_rows)
print(analysis.insights)  # list of Insight
print(analysis.structural_relationships)  # list of StructuralRelationship

# Serialization
d = analysis.to_dict()
j = analysis.to_json_compatible()

# Per-sheet access
for sheet in analysis.sheets:
    print(sheet.sheet_name)
    print(sheet.data_region)
    print(sheet.header)  # HeaderDetection
    print(sheet.columns)  # list of ColumnProfile
    print(sheet.formula_profile)  # FormulaProfile or None
    print(sheet.quality_signals)  # list of QualitySignal
    print(sheet.patterns)  # list of PatternDetection
```

### Domain Models

- `SpreadsheetAnalysis` — root analysis result
- `WorkbookSummary` — workbook-level counters
- `SheetSummary` — per-sheet counters
- `SheetAnalysis` — per-sheet full analysis
- `DataRegion` — non-empty bounding box (min_row, max_row, min_col, max_col, has_data)
- `HeaderDetection` — header status + headers + evidence
- `HeaderInfo` — per-column header detail
- `ColumnProfile` — column type + stats + distribution
- `FormulaProfile` — formula count + cells + distribution + cross-sheet refs
- `FormulaLocation` — per-formula location detail
- `CrossSheetRef` — cross-sheet reference detail
- `QualitySignal` — quality issue signal
- `PatternDetection` — detected pattern
- `PatternLocation` — per-pattern location detail
- `StructuralRelationship` — cross-sheet reference relationship
- `Insight` — structured insight with evidence/rule/severity

### Known Limitations

- Tidak mendukung multiple-table detection dalam satu sheet
- Date sequence pattern hanya mendeteksi DAILY dan WEEKLY, bukan MONTHLY/QUARTERLY/YEARLY
- Tidak ada semantic interpretation atau business logic inference
- Tidak ada formula calculation engine
- Tidak ada formula generation atau application
- Tidak ada AI/LLM integration
- Tidak ada auto data enhancement atau auto-fix

### Safety Contract

- Source workbook tidak dimodifikasi setelah analysis
- MD5 hash source file verified unchanged
- Tidak ada output file yang dicreate
- Protected sheets dapat dianalisis tanpa dimodifikasi
- Read-only adalah property implementation, bukan hanya test artifact

### Determinism Contract

- 5 consecutive runs pada workbook yang sama menghasilkan output yang identik
- Tidak ada random seed, timestamp, hash-based ordering, atau external dependencies
- Serialization output konsisten dan dapat di-bandcompare

---

## MIND-050 Development and Test Status

### Test Results (latest)

```
MIND-010: 23/23 PASS
MIND-020: 40/40 PASS
MIND-030: 57/57 PASS
MIND-040: 57/57 PASS
MIND-050: 63/63 PASS
Combined: 240/240 PASS
```

### Contract Validation

- Input contract (.xlsx only): PASS
- Output contract (SpreadsheetAnalysis): PASS
- Data region (non-empty bounding box): PASS
- Header detection (3 statuses, rule-based): PASS
- Column profiling (9-type taxonomy): PASS
- Formula profile (MIND-030 public API): PASS
- Quality aggregation (MIND-020 public API): PASS
- Pattern detection (6 frozen types): PASS
- Relationship signals (CROSS_SHEET_REFERENCE only): PASS
- Insight taxonomy (5 types, structured): PASS
- Summary (physical + data counters): PASS
- Serialization (to_dict/to_json_compatible): PASS
- Read-only (no mutation): PASS
- Determinism (5 identical runs): PASS
- Private API usage: NONE
- Scope creep: NONE

### Acceptance Audit Result

Product Owner Acceptance Audit: ACCEPTABLE WITH TECHNICAL DEBT

Technical Debt Retained (for future hardening):
- Input type annotation: `str` vs frozen contract `Path` — behavior compatible, annotation mismatch (FIXED in finalization)
- README documentation: not yet synced (FIXED in finalization)
- Weak tests: test_insight_types (no assert), test_insight_has_evidence (empty evidence allowed), test_quality_signal_structure (no meaningful value assert)
- Test blind spots: empty-string-only data region, all-duplicate headers, cross-sheet ref to non-existent target, pattern overlap, serialization round-trip

Blockers: NONE

### Files

- `engine/mind_050/spreadsheet_intelligence.py` — MIND-050 implementation (2036 lines)
- `tests/test_spreadsheet_intelligence.py` — MIND-050 tests (1200 lines, 63 tests)
- `README.md` — updated with MIND-050 documentation section

### Git Status

- Git identity: unchanged
- No Hermes commits
- No Hermes pushes
- Repository changes: MIND-050 implementation files (untracked) + minor pre-existing modifications

---

## Workflow Orchestrator (MIND-070)

MIND-070 is the Workflow Orchestrator (a coordination layer, **not** a spreadsheet engine). It consumes the components MIND-010..MIND-060 through their **public** APIs and drives a multi-step workflow to a **validated output plus persisted state, provenance, checkpoint and recovery**.

### Architecture

```text
User / MeMind Application
  ↓
MIND-070 Workflow Orchestrator  (engine/mind_070/executor.py)
  ├── Plan validation (engine/mind_070/plan_validator.py)
  ├── Operation registry (engine/mind_070/operation_registry.py)
  ├── Dependency-ordered step execution (topological sort)
  ├── Retry (max 1 automatic, transient-only)
  ├── Execution limits (global + per-workflow)
  ├── Cancellation (cooperative at step boundaries)
  ├── Output publication & permissions (engine/mind_070/publisher.py)
  ├── State / provenance / checkpoint / recovery
  │     (engine/mind_070/state, checkpoint, provenance, limits, cancellation)
```

### End-to-end workflow

1. Build a `WorkflowPlan` (unique step ids, every operation registered, no cyclic dependencies).
2. `run_workflow(workflow_id, plan, registry, limits, retry, cancellation_token, output_publication, overwrite_policy, explicit_overwrite, checkpoint_base_dir)`:
   - resolves ``$ref`` placeholders to prior outputs (chained steps bind to the live object),
   - executes steps in dependency order,
   - retryable `TRANSIENT` failures are re-run at most once per step,
   - permanent / validation / unsupported / publication / limit errors halt immediately with `FAILED`,
   - writes a crash-safe checkpoint after each SUCCEEDED step,
   - validates all required outputs before reporting `SUCCEEDED`,
   - publishes outputs with the configured overwrite policy.
3. Final outcome: `SUCCEEDED`, `FAILED`, `PARTIAL`, `CANCELLED`, or `NEEDS_RECOVERY`.

### Components

- `executor.py` — `run_workflow()` (entry point), `_execute_with_retry`, checkpoint/recovery wiring.
- `contracts.py` — `WorkflowPlan`, `Step`, `StepResult`, `WorkflowStatus`, `StepStatus`, `WorkflowErrorCategory`, `WorkflowError`.
- `operation_registry.py` — `OperationRegistry`, `OperationInfo`, `OperationType` (`TEXT_TASK_TO_FORMULA`, `FORMULA_APPLY`, `FILE_ENHANCE`, `SPREADSHEET_INTELLIGENCE`, `AUTO_REPORTING`, `COMPOSITE`).
- `state.py` — JSON workflow state at `.memind/state/<workflow_id>/state.json` + `checkpoint.json`. Crash-safe atomic writes (`tempfile` + `os.replace`). Schema `1.0.0`. Recovery classifications `RECOVERED`/`PARTIAL`/`NEEDS_RECOVERY`/`FAILED`/`NO_STATE`.
- `checkpoint.py` — safe-boundary checkpoint writes; `write_checkpoint`, `load_checkpoint`, `classify_recovery`, `recover_workflow`.
- `provenance.py` — step history (engine, status, artifacts, retry attempt, diagnostics).
- `limits.py` — `GlobalLimits` (defaults) + `WorkflowLimits` (overrides), `validate_limits_config`, `merge_limits`, `check_workflow_limits`.
- `cancellation.py` — `CancellationToken` (cooperative, checked **at step boundaries only**, never mid-operation).
- `publisher.py` — `OutputPublication`, `PublicationError`, `OVERWRITE_NEVER|ALWAYS|EXPLICIT`.

### Output publication & overwrite policy

- **New output is the DEFAULT.** Overwrite requires explicit permission.
- `validate_output_path` rejects path traversal (`..`), absolute paths, and input==output conflicts.
- `resolve` rejects outputs that escape the allowed outputs dir (symlink escape).
- `publish` writes to a staging temp file then `os.replace()` (atomic): an allowed overwrite does not delete the old file before the new one commits, so a failed transfer leaves the previous file intact.
- A publication failure is recorded and the workflow **MUST NOT** report `SUCCEEDED`.

### Retry, limits, cancellation

- **Retry:** at most one automatic retry per step; only transient/retryable errors retried; permanent errors never retried.
- **Limits:** global defaults + per-workflow overrides. A mis-configured limit raises (`WorkflowError`) and never silently disables enforcement.
- **Cancellation:** cooperative, checked at step boundaries only. A cancelled workflow is never `SUCCEEDED` and never re-executes an uncertain step (artifact missing → `NEEDS_RECOVERY`).

### State, provenance, checkpoint, recovery

- Location (default): `.memind/state/<workflow_id>/`.
- Provenance records track engine, status, artifact refs, retry attempt, executed_at, diagnostics.
- Recovery is safe: a state is classified `NEEDS_RECOVERY` when a completed step's artifact or a checkpoint reference is missing — recovery will **not** blind-repeat that step. Operator intervention is required.

### AIProvider (deferred, optional)

- `engine/mind_070/ai_provider.py` defines the minimal contract (`AIProvider`, `AIProviderConfig`, `AIError`, `DeterministicWorkflow`) and the deterministic-fallback flag.
- The orchestrator is deterministic by construction: **no AI provider is required**, no network is touched, and no AIProvider is a hard dependency. Any future AIProvider is a thin, optional overlay.
- A production AIProvider (model/provider selection, timeout, error handling, JSON validation, fallback) is tracked as a separate phase and is intentionally NOT part of this milestone.

### Known limitations

- No new `openpyxl` usage outside MIND-010; deterministic workflows run with zero network/AI dependencies.
- Limit values are safe defaults (provisional); AI response size is recorded but not measured in Phase D.
- `NEEDS_RECOVERY` workflows require operator intervention.
- `engine/mind_070` modules import only the standard library plus MIND-010..MIND-060 (no `openpyxl` outside MIND-010, no `requests`/`httpx`/AI SDKs).

### Running the tests

```bash
cd /data/data/com.termux/files/home/memind
/usr/bin/python3 -m pytest tests/ -q --tb=short
```

Full suite: **371 passed, 0 failed**. Interpreter: `/usr/bin/python3` (the only interpreter with `openpyxl` + `pytest`); the Hermes-agent python lacks both.

### MIND-070 Test Count

| Phase | Test file | Count |
|---|---|---|
| MIND-070A | contracts/registry/validator | 16 |
| MIND-070B | executor/adapters | 16 |
| MIND-070C | state/provenance/checkpoint/recovery | 17 |
| MIND-070D | limits/cancellation/publisher/integration | 17 |
| MIND-070E | integration + acceptance audit | 18 |
| **Total** | | **371** |
