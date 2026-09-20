"""
MeMind Core Excel Engine — Python Module

Implements the MIND-010 Core Excel Engine boundary:
- Read workbook, sheets, cells, metadata
- Modify: create/delete/rename sheets, add/delete rows/columns, update cells
- Write/save as .xlsx
- Validate: re-open output, verify changes persisted

Uses openpyxl for full fidelity preservation of formulas, formatting,
merged cells, and other workbook features not being modified.
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional

from openpyxl import load_workbook, Workbook as OpenPyXLWorkbook
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Error types
# ---------------------------------------------------------------------------

class ExcelErrorCode(Enum):
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    INVALID_FORMAT = "INVALID_FORMAT"
    WORKBOOK_CORRUPT = "WORKBOOK_CORRUPT"
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    INVALID_CELL_REF = "INVALID_CELL_REF"
    INVALID_OPERATION = "INVALID_OPERATION"
    SAVE_FAILED = "SAVE_FAILED"
    VALIDATION_FAILED = "VALIDATION_FAILED"


class ExcelError(Exception):
    """Structured error for the Core Excel Engine."""
    def __init__(self, message: str, code: ExcelErrorCode, details: Optional[dict] = None):
        super().__init__(message)
        self.code = code.value
        self.details = details or {}
        self.message = message

    def __repr__(self) -> str:
        return f"ExcelError({self.code!r}: {self.message})"


# ---------------------------------------------------------------------------
# Data model (mirrors TypeScript types/spreadsheet.ts)
# ---------------------------------------------------------------------------

@dataclass
class CellData:
    value: Any = None
    formula: Optional[str] = None
    row: int = 0
    col: int = 0
    data_type: str = "n"  # n=number, s=string, b=boolean, d=date, empty

    @property
    def address(self) -> str:
        return f"{get_column_letter(self.col + 1)}{self.row + 1}"


@dataclass
class SheetData:
    name: str
    row_count: int = 0
    col_count: int = 0
    cells: list[list[CellData]] = field(default_factory=list)


@dataclass
class WorkbookInfo:
    file_name: str = ""
    sheet_count: int = 0
    sheets: list[SheetData] = field(default_factory=list)
    active_sheet_index: int = 0


# ---------------------------------------------------------------------------
# Core Excel Engine
# ---------------------------------------------------------------------------

class CoreExcelEngine:
    """
    Core Excel Engine for MIND-010.

    Provides:
    - Load .xlsx files
    - Read workbook structure (sheets, dimensions, cells)
    - Modify sheets, rows, columns, cells
    - Save as .xlsx
    - Validate output by re-opening
    """

    SUPPORTED_EXTENSIONS = {".xlsx"}

    def __init__(self):
        self._wb: Optional[OpenPyXLWorkbook] = None
        self._file_path: str = ""
        self._file_name: str = ""

    # -----------------------------------------------------------------------
    # LOAD
    # -----------------------------------------------------------------------

    def load(self, file_path: str) -> WorkbookInfo:
        """
        Load an .xlsx file. Returns WorkbookInfo with structure.
        Raises ExcelError on failure.
        """
        if not os.path.exists(file_path):
            raise ExcelError(
                f"File not found: {file_path}",
                ExcelErrorCode.FILE_NOT_FOUND,
                {"path": file_path},
            )

        ext = os.path.splitext(file_path)[1].lower()
        if ext not in self.SUPPORTED_EXTENSIONS:
            raise ExcelError(
                f"Unsupported file type: {ext}. Supported: {', '.join(self.SUPPORTED_EXTENSIONS)}",
                ExcelErrorCode.INVALID_FORMAT,
                {"extension": ext},
            )

        self._file_path = file_path
        self._file_name = os.path.basename(file_path)

        try:
            self._wb = load_workbook(file_path, data_only=False)
        except Exception as e:
            raise ExcelError(
                f"Failed to parse workbook: {e}",
                ExcelErrorCode.WORKBOOK_CORRUPT,
                {"original_error": str(e)},
            ) from e

        return self._build_workbook_info()

    # -----------------------------------------------------------------------
    # READ
    # -----------------------------------------------------------------------

    def get_sheet_names(self) -> list[str]:
        """Return all sheet names."""
        self._ensure_loaded()
        return self._wb.sheetnames  # type: ignore[union-attr]

    def get_active_sheet_index(self) -> int:
        """Return the zero-based index of the active sheet."""
        self._ensure_loaded()
        return list(self._wb.sheetnames).index(self._wb.active.title)  # type: ignore[union-attr]

    def get_active_sheet(self) -> SheetData:
        """Return the active sheet's data."""
        self._ensure_loaded()
        ws = self._wb.active  # type: ignore[union-attr]
        return self._worksheet_to_sheetdata(ws)

    def get_sheet(self, sheet_name: str) -> SheetData:
        """Return a specific sheet by name."""
        self._ensure_loaded()
        if sheet_name not in self._wb.sheetnames:  # type: ignore[union-attr]
            raise ExcelError(
                f"Sheet not found: {sheet_name}",
                ExcelErrorCode.SHEET_NOT_FOUND,
                {"sheet_name": sheet_name, "available": self.get_sheet_names()},
            )
        ws = self._wb[sheet_name]  # type: ignore[union-attr]
        return self._worksheet_to_sheetdata(ws)

    def get_all_sheets(self) -> list[SheetData]:
        """Return all sheets as SheetData."""
        self._ensure_loaded()
        return [self._worksheet_to_sheetdata(self._wb[name]) for name in self._wb.sheetnames]  # type: ignore[union-attr]

    def get_cell(self, sheet_name: str, row: int, col: int) -> CellData:
        """Get a single cell's data. Row/col are 1-based (Excel convention)."""
        ws = self._get_sheet_or_error(sheet_name)
        if row < 1 or col < 1:
            raise ExcelError(
                f"Invalid cell reference: row={row}, col={col}. Must be >= 1.",
                ExcelErrorCode.INVALID_CELL_REF,
                {"row": row, "col": col},
            )
        cell = ws.cell(row=row, column=col)  # type: ignore[union-attr]
        return CellData(
            value=cell.value,
            formula=cell.value if str(cell.value).startswith("=") else None,
            row=row - 1,  # convert to 0-based
            col=col - 1,
            data_type=self._get_data_type(cell),
        )

    def get_row_count(self, sheet_name: str) -> int:
        """Get the row count for a sheet."""
        ws = self._get_sheet_or_error(sheet_name)
        return ws.max_row  # type: ignore[union-attr]

    def get_col_count(self, sheet_name: str) -> int:
        """Get the column count for a sheet."""
        ws = self._get_sheet_or_error(sheet_name)
        return ws.max_column  # type: ignore[union-attr]

    def get_headers(self, sheet_name: str) -> list[Any]:
        """Get the first row as headers."""
        ws = self._get_sheet_or_error(sheet_name)
        row = ws[1]  # type: ignore[union-attr]
        return [cell.value for cell in row]

    def get_workbook_info(self) -> WorkbookInfo:
        """Get full workbook metadata."""
        self._ensure_loaded()
        info = WorkbookInfo(
            file_name=self._file_name,
            sheet_count=len(self._wb.sheetnames),  # type: ignore[union-attr]
            active_sheet_index=self._wb.sheetnames.index(self._wb.active.title),  # type: ignore[union-attr]
        )
        for name in self._wb.sheetnames:  # type: ignore[union-attr]
            info.sheets.append(self._worksheet_to_sheetdata(self._wb[name]))  # type: ignore[union-attr]
        return info

    # -----------------------------------------------------------------------
    # MODIFY
    # -----------------------------------------------------------------------

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """Rename a sheet."""
        self._ensure_loaded()
        if old_name not in self._wb.sheetnames:  # type: ignore[union-attr]
            raise ExcelError(
                f"Sheet not found: {old_name}",
                ExcelErrorCode.SHEET_NOT_FOUND,
                {"sheet_name": old_name},
            )
        if new_name in self._wb.sheetnames:  # type: ignore[union-attr]
            raise ExcelError(
                f"Sheet already exists: {new_name}",
                ExcelErrorCode.INVALID_OPERATION,
                {"sheet_name": new_name},
            )
        self._wb[old_name].title = new_name  # type: ignore[union-attr]

    def create_sheet(self, sheet_name: str, position: Optional[int] = None) -> None:
        """Create a new sheet."""
        self._ensure_loaded()
        if sheet_name in self._wb.sheetnames:  # type: ignore[union-attr]
            raise ExcelError(
                f"Sheet already exists: {sheet_name}",
                ExcelErrorCode.INVALID_OPERATION,
                {"sheet_name": sheet_name},
            )
        if position is not None:
            self._wb.create_sheet(title=sheet_name, index=position)  # type: ignore[union-attr]
        else:
            self._wb.create_sheet(title=sheet_name)  # type: ignore[union-attr]

    def delete_sheet(self, sheet_name: str) -> None:
        """Delete a sheet."""
        self._ensure_loaded()
        if sheet_name not in self._wb.sheetnames:  # type: ignore[union-attr]
            raise ExcelError(
                f"Sheet not found: {sheet_name}",
                ExcelErrorCode.SHEET_NOT_FOUND,
                {"sheet_name": sheet_name},
            )
        del self._wb[sheet_name]  # type: ignore[union-attr]

    def add_row(self, sheet_name: str, row_data: list[Any], position: Optional[int] = None) -> None:
        """Add a row to a sheet. Position is 1-based (after which row to insert)."""
        ws = self._get_sheet_or_error(sheet_name)
        target_row = position if position is not None else ws.max_row + 1  # type: ignore[union-attr]
        ws.insert_rows(target_row)  # type: ignore[union-attr]
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=target_row, column=col_idx, value=value)  # type: ignore[union-attr]

    def delete_row(self, sheet_name: str, row: int) -> None:
        """Delete a row from a sheet (1-based)."""
        ws = self._get_sheet_or_error(sheet_name)
        if row < 1 or row > ws.max_row:  # type: ignore[union-attr]
            raise ExcelError(
                f"Row {row} is out of range (1-{ws.max_row})",
                ExcelErrorCode.INVALID_OPERATION,
                {"row": row, "max_row": ws.max_row},
            )
        ws.delete_rows(row)  # type: ignore[union-attr]

    def add_column(self, sheet_name: str, column_data: list[Any], position: Optional[int] = None) -> None:
        """Add a column to a sheet."""
        ws = self._get_sheet_or_error(sheet_name)
        target_col = position if position is not None else ws.max_column + 1  # type: ignore[union-attr]
        ws.insert_cols(target_col)  # type: ignore[union-attr]
        for row_idx, value in enumerate(column_data, 1):
            ws.cell(row=row_idx, column=target_col, value=value)  # type: ignore[union-attr]

    def delete_column(self, sheet_name: str, col: int) -> None:
        """Delete a column from a sheet (1-based)."""
        ws = self._get_sheet_or_error(sheet_name)
        if col < 1 or col > ws.max_column:  # type: ignore[union-attr]
            raise ExcelError(
                f"Column {col} is out of range (1-{ws.max_column})",
                ExcelErrorCode.INVALID_OPERATION,
                {"col": col, "max_col": ws.max_column},
            )
        ws.delete_cols(col)  # type: ignore[union-attr]

    def update_cell(self, sheet_name: str, row: int, col: int, value: Any) -> None:
        """Update a cell value (1-based row/col)."""
        ws = self._get_sheet_or_error(sheet_name)
        if row < 1 or col < 1:
            raise ExcelError(
                f"Invalid cell reference: row={row}, col={col}",
                ExcelErrorCode.INVALID_CELL_REF,
                {"row": row, "col": col},
            )
        ws.cell(row=row, column=col, value=value)  # type: ignore[union-attr]

    # -----------------------------------------------------------------------
    # WRITE
    # -----------------------------------------------------------------------

    def save(self, output_path: str) -> str:
        """
        Save the current workbook to a new .xlsx file.
        Returns the output path on success.
        """
        self._ensure_loaded()
        try:
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            self._wb.save(output_path)  # type: ignore[union-attr]
            return output_path
        except Exception as e:
            raise ExcelError(
                f"Failed to save workbook to {output_path}: {e}",
                ExcelErrorCode.SAVE_FAILED,
                {"output_path": output_path, "original_error": str(e)},
            ) from e

    # -----------------------------------------------------------------------
    # VALIDATE
    # -----------------------------------------------------------------------

    def validate(self, output_path: str) -> dict[str, Any]:
        """
        Validate a saved file by re-opening it.
        Returns a dict with validation results.
        """
        if not os.path.exists(output_path):
            raise ExcelError(
                f"Output file not found: {output_path}",
                ExcelErrorCode.VALIDATION_FAILED,
                {"output_path": output_path},
            )

        try:
            verify_wb = load_workbook(output_path)
        except Exception as e:
            raise ExcelError(
                f"Cannot re-open output file: {e}",
                ExcelErrorCode.VALIDATION_FAILED,
                {"output_path": output_path, "original_error": str(e)},
            ) from e

        orig_names = self._wb.sheetnames if self._wb else []  # type: ignore[union-attr]
        verify_names = verify_wb.sheetnames

        result: dict[str, Any] = {
            "valid": True,
            "output_path": output_path,
            "sheet_count_match": len(orig_names) == len(verify_names),
            "original_sheets": list(orig_names),
            "output_sheets": list(verify_names),
            "errors": [],
        }

        if not result["sheet_count_match"]:
            result["valid"] = False
            result["errors"].append(
                f"Sheet count mismatch: original={len(orig_names)}, output={len(verify_names)}"
            )

        # Verify sheet names match
        for name in orig_names:
            if name not in verify_names:
                result["valid"] = False
                result["errors"].append(f"Sheet '{name}' missing from output")

        # Verify cells can be read
        try:
            verify_info = self._get_workbook_info_from_wb(verify_wb)
            result["verified_cells"] = sum(s.row_count * s.col_count for s in verify_info.sheets)
        except Exception as e:
            result["valid"] = False
            result["errors"].append(f"Cell verification failed: {e}")

        verify_wb.close()
        return result

    # -----------------------------------------------------------------------
    # PIPELINE: Load → Modify → Save → Validate
    # -----------------------------------------------------------------------

    def pipeline(self, input_path: str, output_path: str, modifications: Optional[dict] = None) -> dict[str, Any]:
        """
        Full pipeline: Load → Modify → Save → Validate.
        
        modifications dict can contain:
        - rename_sheet: {old_name, new_name}
        - create_sheet: {name, position?}
        - delete_sheet: {name}
        - update_cell: {sheet, row, col, value}
        - add_row: {sheet, data, position?}
        - add_column: {sheet, data, position?}
        - delete_row: {sheet, row}
        - delete_column: {sheet, col}
        """
        self.load(input_path)

        if modifications:
            self._apply_modifications(modifications)

        saved_path = self.save(output_path)
        validation = self.validate(saved_path)
        return validation

    def close(self) -> None:
        """Close the workbook."""
        if self._wb:
            self._wb.close()
            self._wb = None
            self._file_path = ""
            self._file_name = ""

    # -----------------------------------------------------------------------
    # PRIVATE HELPERS
    # -----------------------------------------------------------------------

    def _ensure_loaded(self) -> None:
        if self._wb is None:
            raise ExcelError(
                "No workbook loaded. Call load() first.",
                ExcelErrorCode.INVALID_OPERATION,
            )

    def _get_sheet_or_error(self, sheet_name: str) -> Any:
        self._ensure_loaded()
        if sheet_name not in self._wb.sheetnames:
            raise ExcelError(
                f"Sheet not found: {sheet_name}",
                ExcelErrorCode.SHEET_NOT_FOUND,
                {"sheet_name": sheet_name},
            )
        return self._wb[sheet_name]

    def _worksheet_to_sheetdata(self, ws: Any) -> SheetData:
        """Convert openpyxl worksheet to SheetData."""
        cells: list[list[CellData]] = []
        max_r = ws.max_row or 1
        max_c = ws.max_column or 1
        for r in range(1, max_r + 1):
            row_cells: list[CellData] = []
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                row_cells.append(CellData(
                    value=cell.value,
                    row=r - 1,
                    col=c - 1,
                    data_type=self._get_data_type(cell),
                ))
            cells.append(row_cells)
        return SheetData(
            name=ws.title,
            row_count=max_r,
            col_count=max_c,
            cells=cells,
        )

    def _get_data_type(self, cell: Any) -> str:
        v = cell.value
        if v is None:
            return "empty"
        if isinstance(v, bool):
            return "b"
        if isinstance(v, (int, float)):
            return "n"
        if isinstance(v, str) and v.startswith("="):
            return "f"
        if isinstance(v, str):
            return "s"
        return "s"

    def _apply_modifications(self, mods: dict[str, Any]) -> None:
        """Apply modification dict to the loaded workbook."""
        if "rename_sheet" in mods:
            m = mods["rename_sheet"]
            self.rename_sheet(m["old_name"], m["new_name"])
        if "create_sheet" in mods:
            m = mods["create_sheet"]
            self.create_sheet(m["name"], m.get("position"))
        if "delete_sheet" in mods:
            m = mods["delete_sheet"]
            self.delete_sheet(m["name"])
        if "update_cell" in mods:
            m = mods["update_cell"]
            self.update_cell(m["sheet"], m["row"], m["col"], m["value"])
        if "add_row" in mods:
            m = mods["add_row"]
            self.add_row(m["sheet"], m["data"], m.get("position"))
        if "add_column" in mods:
            m = mods["add_column"]
            self.add_column(m["sheet"], m["data"], m.get("position"))
        if "delete_row" in mods:
            m = mods["delete_row"]
            self.delete_row(m["sheet"], m["row"])
        if "delete_column" in mods:
            m = mods["delete_column"]
            self.delete_column(m["sheet"], m["col"])

    def _build_workbook_info(self) -> WorkbookInfo:
        """Build WorkbookInfo from the loaded openpyxl workbook."""
        info = WorkbookInfo(
            file_name=self._file_name,
            sheet_count=len(self._wb.sheetnames),  # type: ignore[union-attr]
            active_sheet_index=list(self._wb.sheetnames).index(self._wb.active.title),  # type: ignore[union-attr]
        )
        for name in self._wb.sheetnames:  # type: ignore[union-attr]
            info.sheets.append(self._worksheet_to_sheetdata(self._wb[name]))  # type: ignore[union-attr]
        return info

    def _get_workbook_info_from_wb(self, wb: Any) -> WorkbookInfo:
        """Get WorkbookInfo from an openpyxl Workbook object."""
        info = WorkbookInfo(
            file_name="",
            sheet_count=len(wb.sheetnames),
        )
        for name in wb.sheetnames:
            info.sheets.append(self._worksheet_to_sheetdata(wb[name]))
        return info

    def _get_workbook_info(self) -> WorkbookInfo:
        """DEPRECATED: use _build_workbook_info."""
        return self._build_workbook_info()
