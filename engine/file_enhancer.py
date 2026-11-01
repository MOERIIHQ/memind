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

from engine.excel_engine import CoreExcelEngine, ExcelError, CellData, SheetData


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
    PROTECTED_SHEET = "PROTECTED_SHEET"


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

        result = self._analyze_workbook()
        self._engine.close()
        return result

    def _analyze_workbook(self) -> AnalysisResult:
        """
        Shared internal detection routine.
        Does not load or close the engine — caller manages lifecycle.
        Uses only public MIND-010 API.
        """
        info = self._engine.get_workbook_info()
        result = AnalysisResult(
            sheet_count=info.sheet_count,
            sheets_analyzed=[],
        )

        sheet_names = self._engine.get_sheet_names()

        for sheet_name in sheet_names:
            result.sheets_analyzed.append(sheet_name)
            sheet_data = self._engine.get_sheet(sheet_name)

            max_row = sheet_data.row_count or 1
            max_col = sheet_data.col_count or 1

            # Build rows_data from public SheetData.cells
            rows_data: list[list[Any]] = []
            for r in range(max_row):
                row_vals: list[Any] = []
                for c in range(max_col):
                    if r < len(sheet_data.cells) and c < len(sheet_data.cells[r]):
                        row_vals.append(sheet_data.cells[r][c].value)
                    else:
                        row_vals.append(None)
                rows_data.append(row_vals)

            # A. Duplicate rows (exact match, header-aware)
            self._detect_duplicate_rows(sheet_name, rows_data, result)

            # B. Missing values
            self._detect_missing_values(sheet_name, rows_data, max_col, result)

            # C. Column type distribution
            self._detect_type_inconsistency(sheet_name, rows_data, max_col, result)

            # D. Header quality
            self._detect_header_quality(sheet_name, rows_data, max_col, result)

            # E. Formula presence — uses SheetData cells (public API)
            self._detect_formulas_from_sheet(sheet_name, sheet_data, max_row, max_col, result)

        # Populate unified issues list from all detector results
        self._populate_issues(result)
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
        """Detect columns where values have mixed types.
        Excludes header row (index 0) from type inconsistency check."""
        if not rows_data:
            return

        header_row = rows_data[0] if rows_data else []
        data_rows = rows_data[1:]  # Exclude header from type analysis

        for col_idx in range(max_col):
            types_found: set[str] = set()
            type_values: dict[str, int] = {}

            for row_vals in data_rows:
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

            # Track column type info (includes all rows for complete picture)
            all_types: dict[str, int] = {}
            for row_vals in rows_data:
                if col_idx < len(row_vals):
                    val = row_vals[col_idx]
                    t = self._classify_type(val)
                    if t:
                        all_types[t] = all_types.get(t, 0) + 1
            dominant = self._dominant_type(all_types)
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

            # Check empty / whitespace-only header (combined check)
            if val is None or col_name.strip() == "":
                issue_type = "whitespace_only" if col_name != "" else "empty"
                result.header_quality_issues.append(HeaderQualityIssue(
                    sheet=sheet_name,
                    column=f"Col{col_idx + 1}",
                    column_index=col_idx,
                    issue_type=issue_type,
                    description=(
                        f"Header at column {col_idx + 1} is "
                        f"{'whitespace-only' if issue_type == 'whitespace_only' else 'empty'}"
                    ),
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

    def _detect_formulas_from_sheet(
        self,
        sheet_name: str,
        sheet_data: "SheetData",
        max_row: int,
        max_col: int,
        result: AnalysisResult,
    ) -> None:
        """Detect cells containing formulas using public SheetData.cells."""
        formula_count = 0
        locations: list[str] = []
        from openpyxl.utils import get_column_letter

        for r in range(max_row):
            for c in range(max_col):
                if r < len(sheet_data.cells) and c < len(sheet_data.cells[r]):
                    cell = sheet_data.cells[r][c]
                    val = cell.value
                    if isinstance(val, str) and val.startswith("="):
                        formula_count += 1
                        addr = f"{get_column_letter(c + 1)}{r + 1}"
                        locations.append(addr)

        if formula_count > 0:
            result.formula_presence.append(FormulaPresence(
                sheet=sheet_name,
                formula_count=formula_count,
                locations=locations[:20],
            ))

    def _populate_issues(self, result: AnalysisResult) -> None:
        """Populate AnalysisResult.issues from all detector results.
        Contract: issues must be consistent with detector result."""

        # From duplicate rows
        for dup in result.duplicate_rows:
            result.issues.append(EnhancementIssue(
                type="duplicate_rows",
                severity="warning",
                sheet=dup.sheet,
                location=f"rows {dup.row_indices[0]}-{dup.row_indices[-1]}",
                description=(
                    f"Row {dup.row_indices[0]} is duplicated in rows "
                    f"{', '.join(str(i) for i in dup.row_indices[1:])}"
                ),
                current_value=str(dup.values),
                suggested_action="Review duplicate rows manually",
            ))

        # From missing values
        for mv in result.missing_values:
            result.issues.append(EnhancementIssue(
                type="missing_values",
                severity="warning",
                sheet=mv.sheet,
                location=f"column {mv.column} (index {mv.column_index})",
                description=(
                    f"Column '{mv.column}' has {mv.missing_count} missing value(s)"
                ),
                current_value=None,
                suggested_action="Review missing values manually",
            ))

        # From type inconsistencies
        for ti in result.type_inconsistencies:
            result.issues.append(EnhancementIssue(
                type="type_inconsistency",
                severity="warning",
                sheet=ti.sheet,
                location=f"column {ti.column} (index {ti.column_index})",
                description=(
                    f"Column '{ti.column}' has mixed types: "
                    f"{', '.join(ti.types_found)}"
                ),
                current_value=str(ti.types_found),
                suggested_action="Review type inconsistency manually",
            ))

        # From header quality issues
        for hq in result.header_quality_issues:
            sev_map = {
                "empty": "error",
                "duplicate": "warning",
                "whitespace_only": "warning",
                "surrounding_whitespace": "info",
            }
            result.issues.append(EnhancementIssue(
                type=f"header_quality:{hq.issue_type}",
                severity=sev_map.get(hq.issue_type, "info"),
                sheet=hq.sheet,
                location=f"column {hq.column} (index {hq.column_index})",
                description=hq.description,
                current_value=hq.current_value,
                suggested_action=hq.suggested_action,
            ))

        # From formula presence — INFO level
        for fp in result.formula_presence:
            result.issues.append(EnhancementIssue(
                type="formula_detected",
                severity="info",
                sheet=fp.sheet,
                location=f"{fp.formula_count} cell(s)",
                description=(
                    f"Sheet '{fp.sheet}' contains {fp.formula_count} formula cell(s)"
                ),
                current_value=str(fp.locations[:5]),
                suggested_action="Formulas are protected from modification",
            ))

    # ------------------------------------------------------------------
    # ENHANCEMENT
    # ------------------------------------------------------------------

    def enhance(self, input_path: str, output_path: str) -> EnhancementResult:
        """
        Analyze a workbook, apply safe whitespace trimming to text cells,
        save, and validate.

        Only safe auto-fix: text whitespace trimming.
        Uses only public MIND-010 API for analysis and cell modification.
        Protected/hidden structures are detected and reported, not modified.
        """
        try:
            self._engine.load(input_path)
        except ExcelError as e:
            raise FileEnhancerError(
                f"Cannot enhance workbook: {e.message}",
                FileEnhancerErrorCode.INVALID_WORKBOOK.value,
                {"original_error": e.code},
            ) from e

        # Run analysis first — uses public API, does not close engine
        analysis = self._analyze_workbook()
        cells_trimmed = 0
        protected_sheets: list[str] = []
        skipped_cells: list[str] = []

        # Apply whitespace trimming to text cells using public API
        sheet_names = self._engine.get_sheet_names()

        for sheet_name in sheet_names:
            sheet_data = self._engine.get_sheet(sheet_name)
            max_row = sheet_data.row_count or 1
            max_col = sheet_data.col_count or 1

            # Check for protected sheet via public API
            if self._engine.is_sheet_protected(sheet_name):
                protected_sheets.append(sheet_name)
                continue  # Skip trimming on protected sheets

            for r in range(max_row):
                for c in range(max_col):
                    if r < len(sheet_data.cells) and c < len(sheet_data.cells[r]):
                        cell_data = sheet_data.cells[r][c]
                        val = cell_data.value
                        # Skip header row (r==0) from trimming per MIND-020 contract Section 13
                        if r == 0:
                            continue
                        if isinstance(val, str) and not val.startswith("="):
                            trimmed = val.strip()
                            if trimmed != val:
                                # Use public API to update cell
                                self._engine.update_cell(sheet_name, r + 1, c + 1, trimmed)
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

        # Count total issues across all categories
        issues_found = (
            len(analysis.duplicate_rows)
            + len(analysis.missing_values)
            + len(analysis.type_inconsistencies)
            + len(analysis.header_quality_issues)
            + len(analysis.formula_presence)
        )

        result = EnhancementResult(
            original_sheet_count=analysis.sheet_count,
            enhanced_sheet_count=analysis.sheet_count if validation.get("valid", False) else 0,
            issues_found=issues_found,
            enhancements_applied=cells_trimmed,
            cells_trimmed=cells_trimmed,
            analysis=analysis,
            valid=validation.get("valid", False),
            errors=validation.get("errors", []),
        )

        if protected_sheets:
            result.errors.append(
                f"Skipped trimming on protected sheet(s): {', '.join(protected_sheets)}"
            )

        return result

    def _analyze_loaded(self) -> AnalysisResult:
        """Analyze the already-loaded workbook (internal, doesn't close engine).
        Delegates to shared _analyze_workbook()."""
        return self._analyze_workbook()

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
        # Explicit loop to avoid type inference issues with max(key=...)
        best_type: str = "empty"
        best_count = 0
        for t, count in type_counts.items():
            if count > best_count:
                best_count = count
                best_type = t
        return best_type
