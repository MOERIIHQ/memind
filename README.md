# MeMind

React + TypeScript + Vite application.

## Project Structure

- `src/` — React frontend components, services, types, styles
- `engine/` — Core Excel Engine (Python + openpyxl)
  - `excel_engine.py` — MIND-010 Core Excel Engine implementation
  - `create_fixture.py` — Sample workbook generator
  - `__init__.py` — Package init
- `tests/` — Automated tests
  - `test_engine.py` — 23 tests covering the Core Excel Engine
- `public/` — Static assets (test.xlsx, icons, HTML)
- `fixtures/` — Test fixtures (sample.xlsx)

## Core Excel Engine (MIND-010)

The Core Excel Engine is implemented in Python using `openpyxl`.

### Capabilities
- **Read**: Load `.xlsx`, detect sheets, read cells, headers, dimensions
- **Modify**: Rename/create/delete sheets, add/delete rows/columns, update cells
- **Write**: Save as `.xlsx` file
- **Validate**: Re-open output, verify sheet count and structure match
- **Pipeline**: `load → modify → save → validate` in one call

### Running Tests
```bash
cd /data/data/com.termux/files/home/memind
python3 -m pytest tests/test_engine.py -v
```

### All tests pass.

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

## Known Limitations (MIND-010)
- **Supported**: `.xlsx` only (`.xls`, `.csv` are in the frontend but Python engine handles `.xlsx`)
- **Preserved (expected)**: Formulas, merged cells, and formatting are preserved via openpyxl round-trip for unmodified features
- **Not guaranteed**: Advanced charts, external links, complex Excel-specific features, full workbook metadata
- **Not Yet Supported**: AI commands, Pivot, charts, File Enhancer, dashboard, DOCX/PDF/PPTX generation, authentication

## Future Work
- Backend API layer (FastAPI) to expose Core Excel Engine to frontend
- File upload/download REST endpoints
- AI natural-language commands (next milestone)
- SheetJS integration for read-only viewing in browser (already partially implemented)
- More complex operations (formulas, data validation, named ranges)

## Architecture
```
Frontend/UI (React)
    ↓
spreadsheetService.ts (SheetJS)
    ↓
Core Excel Engine (Python + openpyxl) [NEW]
    ↓
File System (.xlsx files)
```

The Python engine is decoupled from the UI and can be tested independently.
The `openpyxl` library preserves formulas, merged cells, and formatting
for features not being modified. Advanced Excel-specific features
(charts, external links) are not guaranteed.
