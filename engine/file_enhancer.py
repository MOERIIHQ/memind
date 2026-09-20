"""
MeMind File Enhancer — MIND-020

Performs deterministic analysis on workbooks and applies safe
whitespace-trimming enhancement.

Architecture:
    CoreExcelEngine  →  FileEnhancer  →  Analysis + Enhancement  →  Save  →  Validate
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional

from engine.excel_engine import CoreExcelEngine, ExcelError, ExcelErrorCode, CellData


# ---------------------------------------------------------------------------
# Error types
# ---------------------------------------------------------------------------

class FileEnhancerError(Exception):
    """Structured error for the File Enhancer."""
    def __init__(self, message: str, code: str, details: Optional[dict] = None):
        super().__init__(message)
        self.code = code
        self.details = details or {}


class FileEnhancerErrorCode(Enum):
    ANALYSIS_FAILED = "ANALYSIS_FAILED"
    ENHANCEMENT_FAILED = "ENHANCEMENT_FAILED"
    INVALID_WORKBOOK = "INVALID_WORKBOOK"
    SAVE_FAILED = "SAVE_FAILED"
    VALIDATION_FAILED = "VALIDATION_FAILED"


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class IssueSeverity(Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


@dataclass
class EnhancementIssue:
    type: str
    severity: str
    sheet: str
    location: str
    description: str
    current_value: Optional[str] = None
    suggested_action: str = ""


@dataclass
class DuplicateRowIssue:
    sheet: str
    row_indices: list[int]  # 1-based Excel row numbers
    values: list[Any]


@dataclass
class MissingValueIssue:
    sheet: str
    column: str
    column_index: int  # 0-based
    missing_count: int


@dataclass
class TypeInconsistencyIssue:
    sheet: str
    column: str
    column_index: int  # 0-based
    types_found: list[str]


@dataclass
class HeaderQualityIssue:
    sheet: str
    column: str
    column_index: int  # 0-based
    issue_type: str  # "empty", "duplicate", "whitespace_only", "surrounding_whitespace"
    description: str
    current_value: Optional[str] = None
    suggested_action: str = ""
    location: str = ""


@dataclass
class FormulaPresence:
    sheet: str
    formula_count: int
    locations: list[str]  # cell addresses like "A1"


@dataclass
class ColumnTypeInfo:
    column: str
    column_index: int
    types_found: list[str]
    dominant_type: str


@dataclass
class AnalysisResult:
    sheet_count: int
    sheets_analyzed: list[str]
    duplicate_rows: list[DuplicateRowIssue] = field(default_factory=list)
    missing_values: list[MissingValueIssue] = field(default_factory=list)
    type_inconsistencies: list[TypeInconsistencyIssue] = field(default_factory=list)
    header_quality_issues: list[HeaderQualityIssue] = field(default_factory=list)
    formula_presence: list[FormulaPresence] = field(default_factory=list)
    column_types: list[ColumnTypeInfo] = field(default_factory=list)
    issues: list[EnhancementIssue] = field(default_factory=list)


@dataclass
class EnhancementResult:
    original_sheet_count: int
    enhanced_sheet_count: int
    issues_found: int
    enhancements_applied: int
    cells_trimmed: int
    analysis: AnalysisResult
    valid: bool = False
    errors: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# File Enhancer
# ---------------------------------------------------------------------------

class FileEnhancer:
    """
    MIND-020 File Enhancer.

    Analyzes workbooks for data quality issues and applies
    safe, deterministic enhancements.
    """

    def __init__(self):
        self._engine = CoreExcelEngine()

    # ------------------------------------------------------------------
    # ANALYSIS
    # ------------------------------------------------------------------

    def analyze(self, file_path: str) -> AnalysisResult:
        """
        Analyze a workbook for data quality issues.
        Returns structured AnalysisResult without modifying the file.
        """
        try:
            self._engine.load(file_path)
        except ExcelError as e:
            raise FileEnhancerError(
                f"Cannot analyze workbook: {e.message}",
                FileEnhancerErrorCode.INVALID_WORKBOOK.value,
                {"original_error": e.code},
            ) from e

        result = AnalysisResult(
            sheet_count=self._engine.get_workbook_info().sheet_count,
            sheets_analyzed=[],
        )

        sheet_names = self._engine.get_sheet_names()

        for sheet_name in sheet_names:
            result.sheets_analyzed.append(sheet_name)
            ws = self._engine._get_sheet_or_error(sheet_name)

            max_row = ws.max_row or 1  # type: ignore
            max_col = ws.max_column or 1  # type: ignore

            # Collect all cell data for this sheet
            rows_data: list[list[Any]] = []
            for r in range(1, max_row + 1):
                row_vals: list[Any] = []
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    row_vals.append(cell.value)
                rows_data.append(row_vals)

            # A. Duplicate rows (exact match, header-aware)
            self._detect_duplicate_rows(sheet_name, rows_data, result)

            # B. Missing values
            self._detect_missing_values(sheet_name, rows_data, max_col, result)

            # C. Column type distribution
            self._detect_type_inconsistency(sheet_name, rows_data, max_col, result)

            # D. Header quality
            self._detect_header_quality(sheet_name, rows_data, max_col, result)

            # E. Formula presence
            self._detect_formulas(sheet_name, ws, max_row, max_col, result)

        self._engine.close()
        return result

    def _detect_duplicate_rows(
        self,
        sheet_name: str,
        rows_data: list[list[Any]],
        result: AnalysisResult,
    ) -> None:
        """Detect exact duplicate rows (header-aware: skip row 0)."""
        if len(rows_data) < 2:
            return

        # Header row (index 0) is excluded from duplicate comparison
        data_rows = rows_data[1:]
        seen: dict[str, int] = {}  # serialized -> first row index (1-based)

        for i, row_vals in enumerate(data_rows):
            # Serialize for comparison, using None as sentinel
            key = self._serialize_row(row_vals)
            excel_row = i + 2  # 1-based, accounting for header

            if key in seen:
                first_row = seen[key]
                # Find or create duplicate entry
                dup = None
                for d in result.duplicate_rows:
                    if d.sheet == sheet_name and d.row_indices[0] == first_row:
                        dup = d
                        break
                if dup is None:
                    dup = DuplicateRowIssue(
                        sheet=sheet_name,
                        row_indices=[first_row],
                        values=rows_data[first_row - 1],
                    )
                    result.duplicate_rows.append(dup)
                dup.row_indices.append(excel_row)
                dup.values = rows_data[first_row - 1]
            else:
                seen[key] = excel_row

    def _detect_missing_values(
        self,
        sheet_name: str,
        rows_data: list[list[Any]],
        max_col: int,
        result: AnalysisResult,
    ) -> None:
        """Detect missing/empty cells per column."""
        if not rows_data:
            return

        header_row = rows_data[0] if rows_data else []

        for col_idx in range(max_col):
            missing_count = 0
            for row_vals in rows_data:
                if col_idx < len(row_vals):
                    val = row_vals[col_idx]
                    if self._is_missing(val):
                        missing_count += 1

            if missing_count > 0:
                col_name = header_row[col_idx] if col_idx < len(header_row) else f"Col{col_idx + 1}"
                result.missing_values.append(MissingValueIssue(
                    sheet=sheet_name,
                    column=str(col_name),
                    column_index=col_idx,
                    missing_count=missing_count,
                ))

    def _detect_type_inconsistency(
        self,
        sheet_name: str,
        rows_data: list[list[Any]],
        max_col: int,
        result: AnalysisResult,
    ) -> None:
        """Detect columns where values have mixed types."""
        if not rows_data:
            return

        header_row = rows_data[0] if rows_data else []

        for col_idx in range(max_col):
            types_found: set[str] = set()
            type_values: dict[str, int] = {}

            for row_vals in rows_data:
                if col_idx < len(row_vals):
                    val = row_vals[col_idx]
                    t = self._classify_type(val)
                    if t:
                        types_found.add(t)
                        type_values[t] = type_values.get(t, 0) + 1

            # Column is type-inconsistent if more than one non-empty type found
            non_empty_types = {t for t in types_found if t != "empty"}
            if len(non_empty_types) > 1:
                col_name = header_row[col_idx] if col_idx < len(header_row) else f"Col{col_idx + 1}"
                result.type_inconsistencies.append(TypeInconsistencyIssue(
                    sheet=sheet_name,
                    column=str(col_name),
                    column_index=col_idx,
                    types_found=sorted(non_empty_types),
                ))

            # Track column type info
            dominant = self._dominant_type(type_values)
            result.column_types.append(ColumnTypeInfo(
                column=str(header_row[col_idx]) if col_idx < len(header_row) else f"Col{col_idx + 1}",
                column_index=col_idx,
                types_found=sorted(types_found),
                dominant_type=dominant,
            ))

    def _detect_header_quality(
        self,
        sheet_name: str,
        rows_data: list[list[Any]],
        max_col: int,
        result: AnalysisResult,
    ) -> None:
        """Detect empty, duplicate, whitespace-only, or padded headers."""
        if not rows_data:
            return

        header_row = rows_data[0]
        seen_headers: dict[str, int] = {}  # normalized -> column index

        for col_idx in range(min(max_col, len(header_row))):
            val = header_row[col_idx]
            col_name = str(val) if val is not None else ""

            # Check empty header
            if val is None or col_name.strip() == "":
                result.header_quality_issues.append(HeaderQualityIssue(
                    sheet=sheet_name,
                    column=f"Col{col_idx + 1}",
                    column_index=col_idx,
                    issue_type="empty",
                    description=f"Header at column {col_idx + 1} is empty or whitespace-only",
                ))
                continue

            # Check whitespace-only
            if col_name.strip() == "" and col_name != "":
                result.header_quality_issues.append(HeaderQualityIssue(
                    sheet=sheet_name,
                    column=f"Col{col_idx + 1}",
                    column_index=col_idx,
                    issue_type="whitespace_only",
                    description=f"Header at column {col_idx + 1} contains only whitespace",
                ))
                continue

            # Check surrounding whitespace
            if col_name != col_name.strip():
                result.header_quality_issues.append(HeaderQualityIssue(
                    sheet=sheet_name,
                    column=f"Col{col_idx + 1}",
                    column_index=col_idx,
                    issue_type="surrounding_whitespace",
                    description=f"Header '{col_name}' has unnecessary surrounding whitespace",
                    current_value=col_name,
                    suggested_action=f"Trim to '{col_name.strip()}'",
                ))

            # Check duplicate headers (normalize for comparison)
            normalized = col_name.strip().lower()
            if normalized in seen_headers:
                prev_col = seen_headers[normalized]
                result.header_quality_issues.append(HeaderQualityIssue(
                    sheet=sheet_name,
                    column=f"Col{col_idx + 1}",
                    column_index=col_idx,
                    issue_type="duplicate",
                    description=f"Duplicate header '{col_name}' found at column {col_idx + 1} (duplicate of column {prev_col + 1})",
                    current_value=col_name,
                    suggested_action="Rename to unique header",
                ))
            else:
                seen_headers[normalized] = col_idx

    def _detect_formulas(
        self,
        sheet_name: str,
        ws: Any,
        max_row: int,
        max_col: int,
        result: AnalysisResult,
    ) -> None:
        """Detect cells containing formulas."""
        formula_count = 0
        locations: list[str] = []

        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    formula_count += 1
                    from openpyxl.utils import get_column_letter
                    addr = f"{get_column_letter(c)}{r}"
                    locations.append(addr)

        if formula_count > 0:
            result.formula_presence.append(FormulaPresence(
                sheet=sheet_name,
                formula_count=formula_count,
                locations=locations[:20],  # Cap at 20 for readability
            ))

    # ------------------------------------------------------------------
    # ENHANCEMENT
    # ------------------------------------------------------------------

    def enhance(self, input_path: str, output_path: str) -> EnhancementResult:
        """
        Analyze a workbook, apply safe whitespace trimming to text cells,
        save, and validate.

        Only safe auto-fix: text whitespace trimming.
        """
        try:
            self._engine.load(input_path)
        except ExcelError as e:
            raise FileEnhancerError(
                f"Cannot enhance workbook: {e.message}",
                FileEnhancerErrorCode.INVALID_WORKBOOK.value,
                {"original_error": e.code},
            ) from e

        # Run analysis first (uses engine, must not close it)
        analysis = self._analyze_loaded()
        cells_trimmed = 0

        # Apply whitespace trimming to all text cells
        sheet_names = self._engine.get_sheet_names()

        for sheet_name in sheet_names:
            ws = self._engine._get_sheet_or_error(sheet_name)
            max_row = ws.max_row or 1  # type: ignore
            max_col = ws.max_column or 1  # type: ignore

            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    if isinstance(cell.value, str) and not cell.value.startswith("="):
                        trimmed = cell.value.strip()
                        if trimmed != cell.value:
                            cell.value = trimmed
                            cells_trimmed += 1

        # Save
        try:
            self._engine.save(output_path)
        except ExcelError as e:
            raise FileEnhancerError(
                f"Failed to save enhanced workbook: {e.message}",
                FileEnhancerErrorCode.SAVE_FAILED.value,
                {"output_path": output_path, "original_error": e.code},
            ) from e

        # Validate
        try:
            validation = self._engine.validate(output_path)
        except ExcelError as e:
            raise FileEnhancerError(
                f"Validation failed: {e.message}",
                FileEnhancerErrorCode.VALIDATION_FAILED.value,
                {"output_path": output_path, "original_error": e.code},
            ) from e

        self._engine.close()

        result = EnhancementResult(
            original_sheet_count=analysis.sheet_count,
            enhanced_sheet_count=analysis.sheet_count if validation.get("valid", False) else 0,
            issues_found=len(analysis.issues),
            enhancements_applied=1 if cells_trimmed > 0 else 0,
            cells_trimmed=cells_trimmed,
            analysis=analysis,
            valid=validation.get("valid", False),
            errors=validation.get("errors", []),
        )

        return result

    def _analyze_loaded(self) -> AnalysisResult:
        """Analyze the already-loaded workbook (internal, doesn't close engine)."""
        result = AnalysisResult(
            sheet_count=self._engine.get_workbook_info().sheet_count,
            sheets_analyzed=[],
        )

        sheet_names = self._engine.get_sheet_names()

        for sheet_name in sheet_names:
            result.sheets_analyzed.append(sheet_name)
            ws = self._engine._get_sheet_or_error(sheet_name)

            max_row = ws.max_row or 1  # type: ignore
            max_col = ws.max_column or 1  # type: ignore

            rows_data: list[list[Any]] = []
            for r in range(1, max_row + 1):
                row_vals: list[Any] = []
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    row_vals.append(cell.value)
                rows_data.append(row_vals)

            self._detect_duplicate_rows(sheet_name, rows_data, result)
            self._detect_missing_values(sheet_name, rows_data, max_col, result)
            self._detect_type_inconsistency(sheet_name, rows_data, max_col, result)
            self._detect_header_quality(sheet_name, rows_data, max_col, result)
            self._detect_formulas(sheet_name, ws, max_row, max_col, result)

        return result

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _serialize_row(self, row_vals: list[Any]) -> str:
        """Serialize a row for duplicate comparison."""
        parts = []
        for v in row_vals:
            if v is None:
                parts.append("NULL")
            elif isinstance(v, str):
                parts.append(f"STR:{v}")
            elif isinstance(v, (int, float)):
                parts.append(f"NUM:{v}")
            elif isinstance(v, bool):
                parts.append(f"BOOL:{v}")
            elif isinstance(v, str) and v.startswith("="):
                parts.append(f"FORMULA:{v}")
            else:
                parts.append(f"RAW:{v}")
        return "|".join(parts)

    def _is_missing(self, val: Any) -> bool:
        """Check if a value represents missing/empty data."""
        if val is None:
            return True
        if isinstance(val, str) and val.strip() == "":
            return True
        return False

    def _classify_type(self, val: Any) -> str:
        """Classify a cell value's type for type distribution analysis."""
        if val is None:
            return "empty"
        if isinstance(val, bool):
            return "boolean"
        if isinstance(val, (int, float)):
            return "numeric"
        if isinstance(val, str):
            if val.startswith("="):
                return "formula"
            if val.strip() == "":
                return "empty"
            # Try to detect if it looks numeric
            try:
                float(val)
                return "numeric_string"
            except (ValueError, TypeError):
                return "text"
        return "other"

    def _dominant_type(self, type_counts: dict[str, int]) -> str:
        """Return the most common type from a type count dictionary."""
        if not type_counts:
            return "empty"
        return max(type_counts, key=type_counts.get)
