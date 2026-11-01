"""
MeMind Spreadsheet Intelligence — MIND-050

Deterministic spreadsheet analysis layer.
Produces structured SpreadsheetAnalysis without modifying input.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from pathlib import Path
from typing import Any, Optional

from engine.excel_engine import CoreExcelEngine, SheetData, CellData
from engine.file_enhancer import FileEnhancer, AnalysisResult
from engine.formula_engine import FormulaEngine, FormulaEngineError, FormulaInfo

# ---------------------------------------------------------------------------
# Error types
# ---------------------------------------------------------------------------

class SpreadsheetIntelligenceError(Exception):
    """Structured error for MIND-050."""

    def __init__(self, message: str, code: str, details: Optional[dict] = None):
        super().__init__(message)
        self.code = code
        self.details = details or {}


class SpreadsheetIntelligenceErrorCode(Enum):
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    INVALID_FORMAT = "INVALID_FORMAT"
    WORKBOOK_CORRUPT = "WORKBOOK_CORRUPT"
    ANALYSIS_FAILED = "ANALYSIS_FAILED"


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------

class HeaderStatus(Enum):
    DETECTED = "DETECTED"
    NOT_DETECTED = "NOT_DETECTED"
    AMBIGUOUS = "AMBIGUOUS"


class ColumnType(Enum):
    EMPTY = "EMPTY"
    TEXT = "TEXT"
    INTEGER = "INTEGER"
    FLOAT = "FLOAT"
    BOOLEAN = "BOOLEAN"
    DATE_DATETIME = "DATE/DATETIME"
    FORMULA = "FORMULA"
    MIXED = "MIXED"
    UNKNOWN = "UNKNOWN"


class InsightType(Enum):
    STRUCTURAL = "STRUCTURAL"
    DATA_QUALITY = "DATA_QUALITY"
    FORMULA = "FORMULA"
    PATTERN = "PATTERN"
    RELATIONSHIP = "RELATIONSHIP"


class InsightSeverity(Enum):
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


class PatternType(Enum):
    SEQUENTIAL_NUMERIC = "SEQUENTIAL_NUMERIC"
    CONSTANT_VALUE = "CONSTANT_VALUE"
    REPEATED_VALUE = "REPEATED_VALUE"
    DATE_SEQUENCE = "DATE_SEQUENCE"
    FORMULA_RELATIVE_PATTERN = "FORMULA_RELATIVE_PATTERN"
    DOMINANT_VALUE = "DOMINANT_VALUE"


class DateSequenceType(Enum):
    DAILY = "DAILY"
    WEEKLY = "WEEKLY"


class RelationshipType(Enum):
    CROSS_SHEET_REFERENCE = "CROSS_SHEET_REFERENCE"


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class DataRegion:
    min_row: int = 0
    max_row: int = 0
    min_col: int = 0
    max_col: int = 0
    has_data: bool = False

    def to_dict(self) -> dict:
        return {
            "min_row": self.min_row,
            "max_row": self.max_row,
            "min_col": self.min_col,
            "max_col": self.max_col,
            "has_data": self.has_data,
        }


@dataclass
class HeaderDetection:
    status: str = HeaderStatus.NOT_DETECTED.value
    row_index: int = 0
    column_count: int = 0
    headers: list = field(default_factory=list)
    evidence: list = field(default_factory=list)
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "status": self.status,
            "row_index": self.row_index,
            "column_count": self.column_count,
            "headers": self.headers,
            "evidence": self.evidence,
            "notes": self.notes,
        }


@dataclass
class HeaderInfo:
    column_index: int = 0
    column_letter: str = ""
    header_value: Any = None
    header_text: str = ""
    is_empty: bool = False
    is_whitespace_only: bool = False
    is_duplicate: bool = False
    duplicate_of: Optional[int] = None
    surrounding_whitespace: bool = False
    uniqueness_score: float = 0.0

    def to_dict(self) -> dict:
        return {
            "column_index": self.column_index,
            "column_letter": self.column_letter,
            "header_value": self.header_value,
            "header_text": self.header_text,
            "is_empty": self.is_empty,
            "is_whitespace_only": self.is_whitespace_only,
            "is_duplicate": self.is_duplicate,
            "duplicate_of": self.duplicate_of,
            "surrounding_whitespace": self.surrounding_whitespace,
            "uniqueness_score": self.uniqueness_score,
        }


@dataclass
class NumericStats:
    min: Optional[float] = None
    max: Optional[float] = None
    sum: Optional[float] = None
    mean: Optional[float] = None
    sample_count: int = 0

    def to_dict(self) -> dict:
        return {
            "min": self.min,
            "max": self.max,
            "sum": self.sum,
            "mean": self.mean,
            "sample_count": self.sample_count,
        }


@dataclass
class DateStats:
    earliest: Optional[date] = None
    latest: Optional[date] = None
    sample_count: int = 0

    def to_dict(self) -> dict:
        return {
            "earliest": self.earliest.isoformat() if self.earliest else None,
            "latest": self.latest.isoformat() if self.latest else None,
            "sample_count": self.sample_count,
        }


@dataclass
class ValueDistribution:
    dominant_value: Optional[str] = None
    dominant_value_count: int = 0
    top_values: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "dominant_value": self.dominant_value,
            "dominant_value_count": self.dominant_value_count,
            "top_values": self.top_values,
        }


@dataclass
class ColumnProfile:
    index: int = 0
    column_letter: str = ""
    header: Optional[str] = None
    header_raw: Any = None
    detected_type: str = ColumnType.EMPTY.value
    non_empty_count: int = 0
    empty_count: int = 0
    unique_count: int = 0
    duplicate_count: int = 0
    formula_count: int = 0
    type_distribution: dict = field(default_factory=dict)
    patterns: list = field(default_factory=list)
    numeric_stats: Optional[NumericStats] = None
    date_stats: Optional[DateStats] = None
    value_distribution: Optional[ValueDistribution] = None
    notes: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "column_letter": self.column_letter,
            "header": self.header,
            "header_raw": self.header_raw,
            "detected_type": self.detected_type,
            "non_empty_count": self.non_empty_count,
            "empty_count": self.empty_count,
            "unique_count": self.unique_count,
            "duplicate_count": self.duplicate_count,
            "formula_count": self.formula_count,
            "type_distribution": self.type_distribution,
            "patterns": self.patterns,
            "numeric_stats": self.numeric_stats.to_dict() if self.numeric_stats else None,
            "date_stats": self.date_stats.to_dict() if self.date_stats else None,
            "value_distribution": self.value_distribution.to_dict() if self.value_distribution else None,
            "notes": self.notes,
        }


@dataclass
class FormulaLocation:
    cell: str = ""
    row: int = 0
    col: int = 0
    formula_type: str = ""
    references: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "cell": self.cell,
            "row": self.row,
            "col": self.col,
            "formula_type": self.formula_type,
            "references": self.references,
        }


@dataclass
class CrossSheetRef:
    source_sheet: str = ""
    source_cell: str = ""
    target_sheet: str = ""
    target_cell: str = ""

    def to_dict(self) -> dict:
        return {
            "source_sheet": self.source_sheet,
            "source_cell": self.source_cell,
            "target_sheet": self.target_sheet,
            "target_cell": self.target_cell,
        }


@dataclass
class FormulaProfile:
    sheet_name: str = ""
    formula_count: int = 0
    formula_cells: list = field(default_factory=list)
    formula_distribution: dict = field(default_factory=dict)
    has_cross_sheet_references: bool = False
    cross_sheet_references: list = field(default_factory=list)
    has_unknown_formulas: bool = False
    notes: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "sheet_name": self.sheet_name,
            "formula_count": self.formula_count,
            "formula_cells": [f.to_dict() for f in self.formula_cells],
            "formula_distribution": self.formula_distribution,
            "has_cross_sheet_references": self.has_cross_sheet_references,
            "cross_sheet_references": [c.to_dict() for c in self.cross_sheet_references],
            "has_unknown_formulas": self.has_unknown_formulas,
            "notes": self.notes,
        }


@dataclass
class QualitySignal:
    type: str = ""
    severity: str = ""
    sheet: str = ""
    location: str = ""
    description: str = ""
    evidence: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "type": self.type,
            "severity": self.severity,
            "sheet": self.sheet,
            "location": self.location,
            "description": self.description,
            "evidence": self.evidence,
        }


@dataclass
class PatternLocation:
    column_index: int = 0
    column_letter: str = ""
    description: str = ""
    evidence: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "column_index": self.column_index,
            "column_letter": self.column_letter,
            "description": self.description,
            "evidence": self.evidence,
        }


@dataclass
class PatternDetection:
    pattern_name: str = ""
    column_indices: list = field(default_factory=list)
    sheet: str = ""
    evidence: list = field(default_factory=list)
    description: str = ""
    locations: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "pattern_name": self.pattern_name,
            "column_indices": self.column_indices,
            "sheet": self.sheet,
            "evidence": self.evidence,
            "description": self.description,
            "locations": [loc.to_dict() for loc in self.locations],
        }


@dataclass
class StructuralRelationship:
    source_sheet: str = ""
    source_cell: str = ""
    target_sheet: str = ""
    target_cell: str = ""
    reference_type: str = ""
    status: str = "DETECTED"
    evidence: list = field(default_factory=list)
    risk_note: str = ""

    def to_dict(self) -> dict:
        return {
            "source_sheet": self.source_sheet,
            "source_cell": self.source_cell,
            "target_sheet": self.target_sheet,
            "target_cell": self.target_cell,
            "reference_type": self.reference_type,
            "status": self.status,
            "evidence": self.evidence,
            "risk_note": self.risk_note,
        }


@dataclass
class Insight:
    insight_id: str = ""
    type: str = ""
    severity: str = ""
    sheet: Optional[str] = None
    location: Optional[str] = None
    description: str = ""
    evidence: list = field(default_factory=list)
    rule: str = ""
    suggested_action: Optional[str] = None
    source: str = ""
    related_to: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "insight_id": self.insight_id,
            "type": self.type,
            "severity": self.severity,
            "sheet": self.sheet,
            "location": self.location,
            "description": self.description,
            "evidence": self.evidence,
            "rule": self.rule,
            "suggested_action": self.suggested_action,
            "source": self.source,
            "related_to": self.related_to,
        }


@dataclass
class SheetSummary:
    sheet_name: str = ""
    physical_rows: int = 0
    physical_columns: int = 0
    data_rows: int = 0
    data_columns: int = 0
    formula_cell_count: int = 0
    quality_signal_count: int = 0
    pattern_count: int = 0
    insight_count: int = 0
    header_status: str = ""

    def to_dict(self) -> dict:
        return {
            "sheet_name": self.sheet_name,
            "physical_rows": self.physical_rows,
            "physical_columns": self.physical_columns,
            "data_rows": self.data_rows,
            "data_columns": self.data_columns,
            "formula_cell_count": self.formula_cell_count,
            "quality_signal_count": self.quality_signal_count,
            "pattern_count": self.pattern_count,
            "insight_count": self.insight_count,
            "header_status": self.header_status,
        }


@dataclass
class QualitySummary:
    total_signals: int = 0
    by_type: dict = field(default_factory=dict)
    by_severity: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "total_signals": self.total_signals,
            "by_type": self.by_type,
            "by_severity": self.by_severity,
        }


@dataclass
class SheetAnalysis:
    sheet_name: str = ""
    row_count: int = 0
    col_count: int = 0
    data_region: Optional[DataRegion] = None
    header: Optional[HeaderDetection] = None
    columns: list = field(default_factory=list)
    formula_profile: Optional[FormulaProfile] = None
    quality_signals: list = field(default_factory=list)
    patterns: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "sheet_name": self.sheet_name,
            "row_count": self.row_count,
            "col_count": self.col_count,
            "data_region": self.data_region.to_dict() if self.data_region else None,
            "header": self.header.to_dict() if self.header else None,
            "columns": [c.to_dict() for c in self.columns],
            "formula_profile": self.formula_profile.to_dict() if self.formula_profile else None,
            "quality_signals": [q.to_dict() for q in self.quality_signals],
            "patterns": [p.to_dict() for p in self.patterns],
        }


@dataclass
class WorkbookSummary:
    file_name: str = ""
    sheet_count: int = 0
    total_physical_rows: int = 0
    total_physical_columns: int = 0
    total_data_rows: int = 0
    total_data_columns: int = 0
    total_formula_cells: int = 0
    total_quality_signals: int = 0
    pattern_count: int = 0
    insight_count: int = 0
    sheet_summaries: list = field(default_factory=list)
    quality_summary: Optional[QualitySummary] = None

    def to_dict(self) -> dict:
        return {
            "file_name": self.file_name,
            "sheet_count": self.sheet_count,
            "total_physical_rows": self.total_physical_rows,
            "total_physical_columns": self.total_physical_columns,
            "total_data_rows": self.total_data_rows,
            "total_data_columns": self.total_data_columns,
            "total_formula_cells": self.total_formula_cells,
            "total_quality_signals": self.total_quality_signals,
            "pattern_count": self.pattern_count,
            "insight_count": self.insight_count,
            "sheet_summaries": [s.to_dict() for s in self.sheet_summaries],
            "quality_summary": self.quality_summary.to_dict() if self.quality_summary else None,
        }


@dataclass
class SpreadsheetAnalysis:
    analysis_metadata: dict = field(default_factory=dict)
    workbook: dict = field(default_factory=dict)
    sheets: list = field(default_factory=list)
    quality_signals: list = field(default_factory=list)
    structural_relationships: list = field(default_factory=list)
    summary: Optional[WorkbookSummary] = None
    insights: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "analysis_metadata": self.analysis_metadata,
            "workbook": self.workbook,
            "sheets": [s.to_dict() for s in self.sheets],
            "quality_signals": [q.to_dict() for q in self.quality_signals],
            "structural_relationships": [r.to_dict() for r in self.structural_relationships],
            "summary": self.summary.to_dict() if self.summary else None,
            "insights": [i.to_dict() for i in self.insights],
        }

    def to_json_compatible(self) -> dict:
        return self.to_dict()


# ---------------------------------------------------------------------------
# Spreadsheet Intelligence Analyzer
# ---------------------------------------------------------------------------

class SpreadsheetIntelligence:
    """MIND-050 Deterministic Spreadsheet Intelligence Analyzer.

    Read-only analysis. Produces SpreadsheetAnalysis from .xlsx files.
    Uses only public APIs from MIND-010, MIND-020, and MIND-030.
    """

    def __init__(self):
        self._engine = CoreExcelEngine()
        self._enhancer = FileEnhancer()
        self._formula_engine = FormulaEngine()

    def analyze(self, input_path: Path) -> SpreadsheetAnalysis:
        """Analyze a workbook and return structured SpreadsheetAnalysis.

        Args:
            input_path: Path to .xlsx file.

        Returns:
            SpreadsheetAnalysis with full workbook profile.
        """
        path = Path(input_path)
        if not path.exists():
            raise SpreadsheetIntelligenceError(
                f"File not found: {input_path}",
                SpreadsheetIntelligenceErrorCode.FILE_NOT_FOUND.value,
                {"path": input_path},
            )

        ext = path.suffix.lower()
        if ext != ".xlsx":
            raise SpreadsheetIntelligenceError(
                f"Unsupported file type: {ext}. Supported: .xlsx",
                SpreadsheetIntelligenceErrorCode.INVALID_FORMAT.value,
                {"extension": ext},
            )

        # Load workbook
        try:
            self._engine.load(input_path)
        except Exception as e:
            raise SpreadsheetIntelligenceError(
                f"Failed to load workbook: {e}",
                SpreadsheetIntelligenceErrorCode.WORKBOOK_CORRUPT.value,
                {"original_error": str(e)},
            ) from e

        try:
            return self._analyze_loaded(input_path)
        finally:
            self._engine.close()

    def _analyze_loaded(self, input_path: str) -> SpreadsheetAnalysis:
        """Run analysis on already-loaded workbook.

        Coordinates MIND-010, MIND-020, MIND-030 analysis.
        """
        # Get workbook structure
        wb_info = self._engine.get_workbook_info()
        sheet_names = self._engine.get_sheet_names()

        # Get MIND-020 quality analysis
        analysis_result_m20 = self._enhancer.analyze(input_path)

        # Get MIND-030 formula analysis
        try:
            formula_infos = self._formula_engine.scan_formulas(input_path)
        except FormulaEngineError:
            formula_infos = []

        # Build analysis piece by piece
        sheets: list[SheetAnalysis] = []
        all_quality_signals: list[QualitySignal] = []
        all_patterns: list[PatternDetection] = []
        all_insights: list[Insight] = []
        all_relationships: list[StructuralRelationship] = []

        total_physical_rows = 0
        total_physical_columns = 0
        total_data_rows = 0
        total_data_columns = 0
        total_formula_cells = 0
        total_quality_signals = 0

        for sheet_name in sheet_names:
            sheet_data = self._engine.get_sheet(sheet_name)
            row_count = sheet_data.row_count or 1
            col_count = sheet_data.col_count or 1

            physical_rows = row_count
            physical_columns = col_count
            total_physical_rows += physical_rows
            total_physical_columns += physical_columns

            # Data region
            data_region = self._detect_data_region(sheet_data)

            # Header detection
            header_detection = self._detect_header(sheet_data, data_region)

            # Column profiling
            columns = self._profile_columns(sheet_data, data_region, header_detection)

            # Formula profile for this sheet
            formula_profile = self._build_formula_profile(
                formula_infos, sheet_name
            )

            # Quality signals (from MIND-020, per sheet)
            sheet_quality_signals = self._aggregate_quality_signals(
                analysis_result_m20, sheet_name
            )

            # Patterns (per sheet)
            sheet_patterns = self._detect_patterns(
                sheet_data, data_region, header_detection, columns, formula_infos
            )

            # Structural relationships (cross-sheet references)
            sheet_relationships = self._detect_relationships(
                formula_infos, sheet_name
            )

            # Insights
            sheet_insights = self._generate_insights(
                sheet_name,
                sheet_data,
                data_region,
                header_detection,
                columns,
                formula_profile,
                sheet_quality_signals,
                sheet_patterns,
                sheet_relationships,
            )

            total_formula_cells += formula_profile.formula_count if formula_profile else 0
            total_quality_signals += len(sheet_quality_signals)

            sheet_summary = SheetSummary(
                sheet_name=sheet_name,
                physical_rows=physical_rows,
                physical_columns=physical_columns,
                data_rows=data_region.max_row - data_region.min_row + 1 if data_region.has_data else 0,
                data_columns=data_region.max_col - data_region.min_col + 1 if data_region.has_data else 0,
                formula_cell_count=formula_profile.formula_count if formula_profile else 0,
                quality_signal_count=len(sheet_quality_signals),
                pattern_count=len(sheet_patterns),
                insight_count=len(sheet_insights),
                header_status=header_detection.status if header_detection else "",
            )

            sheet_analysis = SheetAnalysis(
                sheet_name=sheet_name,
                row_count=row_count,
                col_count=col_count,
                data_region=data_region,
                header=header_detection,
                columns=columns,
                formula_profile=formula_profile,
                quality_signals=sheet_quality_signals,
                patterns=sheet_patterns,
            )

            sheets.append(sheet_analysis)
            all_quality_signals.extend(sheet_quality_signals)
            all_patterns.extend(sheet_patterns)
            all_insights.extend(sheet_insights)
            all_relationships.extend(sheet_relationships)
            total_data_rows += sheet_summary.data_rows
            total_data_columns += sheet_summary.data_columns

        # Aggregate quality summary
        quality_summary = self._build_quality_summary(all_quality_signals)

        # Build workbook summary
        summary = WorkbookSummary(
            file_name=Path(input_path).name,
            sheet_count=len(sheet_names),
            total_physical_rows=total_physical_rows,
            total_physical_columns=total_physical_columns,
            total_data_rows=total_data_rows,
            total_data_columns=total_data_columns,
            total_formula_cells=total_formula_cells,
            total_quality_signals=total_quality_signals,
            pattern_count=len(all_patterns),
            insight_count=len(all_insights),
            sheet_summaries=[
                SheetSummary(
                    sheet_name=s.sheet_name,
                    physical_rows=s.row_count,
                    physical_columns=s.col_count,
                    data_rows=(
                        s.data_region.max_row - s.data_region.min_row + 1
                        if s.data_region and s.data_region.has_data else 0
                    ),
                    data_columns=(
                        s.data_region.max_col - s.data_region.min_col + 1
                        if s.data_region and s.data_region.has_data else 0
                    ),
                    formula_cell_count=(
                        s.formula_profile.formula_count if s.formula_profile else 0
                    ),
                    quality_signal_count=len(s.quality_signals),
                    pattern_count=len(s.patterns),
                    insight_count=0,
                    header_status=(
                        s.header.status if s.header else ""
                    ),
                )
                for s in sheets
            ],
            quality_summary=quality_summary,
        )

        for s in summary.sheet_summaries:
            sheet_insights = [i for i in all_insights if i.sheet == s.sheet_name]
            s.insight_count = len(sheet_insights)

        summary.insight_count = len(all_insights)

        # Analysis metadata
        analysis_metadata = {
            "workbook_path": input_path,
            "analyzed_via": "MIND-050 SpreadsheetIntelligence",
        }

        workbook_info = {
            "file_name": Path(input_path).name,
            "sheet_count": len(sheet_names),
            "sheets": sheet_names,
        }

        return SpreadsheetAnalysis(
            analysis_metadata=analysis_metadata,
            workbook=workbook_info,
            sheets=sheets,
            quality_signals=all_quality_signals,
            structural_relationships=all_relationships,
            summary=summary,
            insights=all_insights,
        )

    # ------------------------------------------------------------------
    # Data region detection
    # ------------------------------------------------------------------

    def _detect_data_region(self, sheet_data: SheetData) -> DataRegion:
        """Detect rectangular bounding box of non-empty cells.

        Non-empty = value is not None and not empty string.
        """
        if not sheet_data.cells:
            return DataRegion()

        min_r, max_r = None, None
        min_c, max_c = None, None

        for r_idx, row in enumerate(sheet_data.cells):
            for c_idx, cell in enumerate(row):
                if cell.value is not None and cell.value != "":
                    if min_r is None or r_idx < min_r:
                        min_r = r_idx
                    if max_r is None or r_idx > max_r:
                        max_r = r_idx
                    if min_c is None or c_idx < min_c:
                        min_c = c_idx
                    if max_c is None or c_idx > max_c:
                        max_c = c_idx

        if min_r is None:
            return DataRegion()

        return DataRegion(
            min_row=min_r,
            max_row=max_r,
            min_col=min_c,
            max_col=max_c,
            has_data=True,
        )

    # ------------------------------------------------------------------
    # Header detection
    # ------------------------------------------------------------------

    def _detect_header(
        self, sheet_data: SheetData, data_region: DataRegion
    ) -> Optional[HeaderDetection]:
        """Detect header row using deterministic rules.

        Returns DETECTED, NOT_DETECTED, or AMBIGUOUS.
        """
        if not data_region.has_data:
            return HeaderDetection(
                status=HeaderStatus.NOT_DETECTED.value,
                evidence=["No non-empty cells in sheet"],
            )

        # Candidate = first row in data region (min_row)
        candidate_row = data_region.min_row
        if candidate_row >= len(sheet_data.cells):
            return HeaderDetection(
                status=HeaderStatus.NOT_DETECTED.value,
                evidence=["Candidate row beyond sheet data"],
            )

        row = sheet_data.cells[candidate_row]
        if not row:
            return HeaderDetection(
                status=HeaderStatus.NOT_DETECTED.value,
                evidence=["Candidate row is empty"],
            )

        column_count = len(row)
        headers: list[HeaderInfo] = []
        text_count = 0
        unique_values: set = set()

        for c_idx, cell in enumerate(row):
            col_letter = self._col_index_to_letter(c_idx)
            header_info = HeaderInfo(
                column_index=c_idx,
                column_letter=col_letter,
                header_value=cell.value,
                header_text=str(cell.value) if cell.value is not None else "",
                is_empty=(cell.value is None or cell.value == ""),
                is_whitespace_only=(
                    isinstance(cell.value, str)
                    and cell.value.strip() == ""
                ),
                surrounding_whitespace=(
                    isinstance(cell.value, str)
                    and cell.value != cell.value.strip()
                ),
            )

            # Check duplicate
            if header_info.header_text and not header_info.is_empty:
                for prev_idx, prev_info in enumerate(headers):
                    if (
                        not prev_info.is_empty
                        and prev_info.header_text.strip().lower()
                        == header_info.header_text.strip().lower()
                    ):
                        header_info.is_duplicate = True
                        header_info.duplicate_of = prev_idx
                        break

            headers.append(header_info)

            if header_info.header_text and not header_info.is_empty:
                unique_values.add(header_info.header_text.strip().lower())
                # Count text values
                if isinstance(cell.value, str) and cell.value.strip():
                    text_count += 1

        # Evaluate
        non_empty_count = sum(
            1 for h in headers
            if h.header_text and not h.is_empty and not h.is_whitespace_only
        )

        if non_empty_count == 0:
            # Try next row
            return self._try_alternate_header(sheet_data, data_region)

        text_density = text_count / max(column_count, 1)

        # Uniqueness check
        uniqueness_score = (
            len(unique_values) / max(non_empty_count, 1)
        )

        evidence: list = []
        if text_density >= 0.5:
            evidence.append(
                f"Text density: {text_density:.2f} (>= 0.5 threshold)"
            )
        else:
            evidence.append(
                f"Text density: {text_density:.2f} (< 0.5 threshold)"
            )

        if uniqueness_score >= 0.5:
            evidence.append(
                f"Uniqueness: {uniqueness_score:.2f} (>= 0.5 threshold)"
            )
        else:
            evidence.append(
                f"Uniqueness: {uniqueness_score:.2f} (< 0.5 threshold)"
            )

        if text_density >= 0.5 and uniqueness_score >= 0.5:
            status = HeaderStatus.DETECTED
            notes = "Header detected by rule: text density >= 0.5 AND uniqueness >= 0.5"
        elif text_density < 0.3:
            status = HeaderStatus.NOT_DETECTED
            notes = (
                "Header not detected: text density < 0.3, "
                "row appears numeric/boolean/empty"
            )
        else:
            status = HeaderStatus.AMBIGUOUS
            notes = (
                "Header ambiguous: does not meet DETECTED or NOT_DETECTED criteria"
            )

        return HeaderDetection(
            status=status.value,
            row_index=candidate_row,
            column_count=column_count,
            headers=[h.to_dict() for h in headers],
            evidence=evidence,
            notes=notes,
        )

    def _try_alternate_header(
        self, sheet_data: SheetData, data_region: DataRegion
    ) -> HeaderDetection:
        """Try next row if first row candidate is all empty/whitespace."""
        for row_idx in range(data_region.min_row + 1, data_region.max_row + 1):
            if row_idx >= len(sheet_data.cells):
                break
            row = sheet_data.cells[row_idx]
            if not row:
                continue
            non_empty = sum(
                1 for cell in row
                if cell.value is not None and cell.value != ""
            )
            if non_empty > 0:
                # Found alternate row — evaluate same way
                return self._detect_header(sheet_data, DataRegion(
                    min_row=row_idx,
                    max_row=data_region.max_row,
                    min_col=data_region.min_col,
                    max_col=data_region.max_col,
                    has_data=True,
                ))
        return HeaderDetection(
            status=HeaderStatus.NOT_DETECTED.value,
            evidence=["No suitable header row found"],
        )

    # ------------------------------------------------------------------
    # Column profiling
    # ------------------------------------------------------------------

    def _profile_columns(
        self,
        sheet_data: SheetData,
        data_region: DataRegion,
        header_detection: Optional[HeaderDetection],
    ) -> list[ColumnProfile]:
        """Profile each column in data region."""
        if not data_region.has_data:
            return []

        start_row = data_region.min_row
        end_row = data_region.max_row
        start_col = data_region.min_col
        end_col = data_region.max_col

        # Determine if header is included in profiling
        header_row = None
        header_start_col = start_col
        header_end_col = end_col
        if header_detection and header_detection.status == HeaderStatus.DETECTED.value:
            hr = header_detection.row_index
            if hr == start_row:
                header_row = hr
                # Exclude header from data profiling
                data_start_row = start_row + 1
            else:
                data_start_row = start_row
        else:
            data_start_row = start_row

        columns: list[ColumnProfile] = []

        for c_idx in range(start_col, end_col + 1):
            col_letter = self._col_index_to_letter(c_idx)
            profile = ColumnProfile(
                index=c_idx,
                column_letter=col_letter,
            )

            # Header info
            if header_row is not None and header_row < len(sheet_data.cells):
                header_cell = sheet_data.cells[header_row][c_idx] if c_idx < len(sheet_data.cells[header_row]) else None
                if header_cell and header_cell.value is not None:
                    profile.header = str(header_cell.value)
                    profile.header_raw = header_cell.value

            # Collect values for this column (data rows only)
            values: list = []
            formula_count = 0
            type_counts: dict = {}

            for r_idx in range(data_start_row, end_row + 1):
                if r_idx >= len(sheet_data.cells):
                    continue
                row = sheet_data.cells[r_idx]
                if c_idx >= len(row):
                    continue
                cell = row[c_idx]
                val = cell.value

                if val is None or val == "":
                    profile.empty_count += 1
                    continue

                profile.non_empty_count += 1
                values.append(val)

                # Type classification
                t = self._classify_value_type(val)
                type_counts[t] = type_counts.get(t, 0) + 1

                if t == "FORMULA":
                    formula_count += 1

            profile.formula_count = formula_count
            profile.type_distribution = type_counts

            # Unique count
            unique_vals = set()
            for v in values:
                if isinstance(v, str) and v.startswith("="):
                    # Formula — consider formula text as value for uniqueness
                    unique_vals.add(v)
                else:
                    unique_vals.add(repr(v))
            profile.unique_count = len(unique_vals)
            profile.duplicate_count = profile.non_empty_count - profile.unique_count

            # Detected type
            profile.detected_type = self._determine_column_type(type_counts)

            # Stats
            if profile.detected_type in (ColumnType.INTEGER.value, ColumnType.FLOAT.value):
                profile.numeric_stats = self._compute_numeric_stats(values)

            if profile.detected_type == ColumnType.DATE_DATETIME.value:
                profile.date_stats = self._compute_date_stats(values)

            if profile.detected_type == ColumnType.TEXT.value:
                profile.value_distribution = self._compute_value_distribution(values)

            # Notes
            if formula_count > 0:
                profile.notes.append(
                    f"Contains {formula_count} formula cell(s)"
                )

            if profile.empty_count > 0:
                profile.notes.append(
                    f"{profile.empty_count} empty cell(s) in data region"
                )

            columns.append(profile)

        return columns

    # ------------------------------------------------------------------
    # Formula profile
    # ------------------------------------------------------------------

    def _build_formula_profile(
        self, formula_infos: list[FormulaInfo], sheet_name: str
    ) -> Optional[FormulaProfile]:
        """Build formula profile for a sheet from MIND-030 FormulaInfo list."""
        sheet_formulas = [fi for fi in formula_infos if fi.sheet == sheet_name]

        if not sheet_formulas:
            return None

        formula_cells = []
        distribution: dict = {}
        cross_sheet_refs: list[CrossSheetRef] = []

        for fi in sheet_formulas:
            formula_cells.append(FormulaLocation(
                cell=fi.cell,
                row=fi.row,
                col=fi.col,
                formula_type=fi.formula_type,
                references=fi.references,
            ))
            distribution[fi.formula_type] = distribution.get(fi.formula_type, 0) + 1

            # Cross-sheet references
            for ref in fi.references:
                if "!" in ref:
                    parts = ref.split("!")
                    if len(parts) == 2:
                        cross_sheet_refs.append(CrossSheetRef(
                            source_sheet=sheet_name,
                            source_cell=fi.cell,
                            target_sheet=parts[0],
                            target_cell=parts[1],
                        ))

        has_unknown = any(
            fi.formula_type == "unknown" for fi in sheet_formulas
        )

        return FormulaProfile(
            sheet_name=sheet_name,
            formula_count=len(sheet_formulas),
            formula_cells=formula_cells,
            formula_distribution=distribution,
            has_cross_sheet_references=len(cross_sheet_refs) > 0,
            cross_sheet_references=cross_sheet_refs,
            has_unknown_formulas=has_unknown,
        )

    # ------------------------------------------------------------------
    # Quality signal aggregation
    # ------------------------------------------------------------------

    def _aggregate_quality_signals(
        self, m20_result: AnalysisResult, sheet_name: str
    ) -> list[QualitySignal]:
        """Aggregate MIND-020 AnalysisResult into QualitySignal list."""
        signals: list[QualitySignal] = []

        # Duplicate rows
        for dup in m20_result.duplicate_rows:
            if dup.sheet == sheet_name:
                signals.append(QualitySignal(
                    type="duplicate_rows",
                    severity="warning",
                    sheet=sheet_name,
                    location=f"rows {dup.row_indices[0]}-{dup.row_indices[-1]}",
                    description=(
                        f"Row {dup.row_indices[0]} duplicated in "
                        f"{', '.join(str(i) for i in dup.row_indices[1:])}"
                    ),
                    evidence=[f"Row values: {dup.values}"],
                ))

        # Missing values
        for mv in m20_result.missing_values:
            if mv.sheet == sheet_name:
                signals.append(QualitySignal(
                    type="missing_values",
                    severity="warning",
                    sheet=sheet_name,
                    location=f"column {mv.column} (index {mv.column_index})",
                    description=f"Column '{mv.column}' has {mv.missing_count} missing value(s)",
                    evidence=[],
                ))

        # Type inconsistency
        for ti in m20_result.type_inconsistencies:
            if ti.sheet == sheet_name:
                signals.append(QualitySignal(
                    type="type_inconsistency",
                    severity="warning",
                    sheet=sheet_name,
                    location=f"column {ti.column} (index {ti.column_index})",
                    description=(
                        f"Column '{ti.column}' has mixed types: "
                        f"{', '.join(ti.types_found)}"
                    ),
                    evidence=[],
                ))

        # Header quality
        severity_map = {
            "empty": "error",
            "duplicate": "warning",
            "whitespace_only": "warning",
            "surrounding_whitespace": "info",
        }
        for hq in m20_result.header_quality_issues:
            if hq.sheet == sheet_name:
                signals.append(QualitySignal(
                    type=f"header_quality:{hq.issue_type}",
                    severity=severity_map.get(hq.issue_type, "info"),
                    sheet=sheet_name,
                    location=f"column {hq.column} (index {hq.column_index})",
                    description=hq.description,
                    evidence=[],
                ))

        # Formula presence
        for fp in m20_result.formula_presence:
            if fp.sheet == sheet_name:
                signals.append(QualitySignal(
                    type="formula_detected",
                    severity="info",
                    sheet=sheet_name,
                    location=f"{fp.formula_count} cell(s)",
                    description=(
                        f"Sheet '{sheet_name}' contains "
                        f"{fp.formula_count} formula cell(s)"
                    ),
                    evidence=[fp.locations[:5]],
                ))

        return signals

    # ------------------------------------------------------------------
    # Pattern detection
    # ------------------------------------------------------------------

    def _detect_patterns(
        self,
        sheet_data: SheetData,
        data_region: DataRegion,
        header_detection: HeaderDetection,
        columns: list[ColumnProfile],
        formula_infos: list[FormulaInfo],
    ) -> list[PatternDetection]:
        """Detect patterns across columns."""
        patterns: list[PatternDetection] = []

        if not data_region.has_data:
            return patterns

        header_row = None
        if header_detection and header_detection.status == HeaderStatus.DETECTED.value:
            if header_detection.row_index == data_region.min_row:
                header_row = header_detection.row_index

        data_start = data_region.min_row + 1 if header_row == data_region.min_row else data_region.min_row

        for col_profile in columns:
            c_idx = col_profile.index

            # Get values for this column (data rows)
            values: list = []
            for r_idx in range(data_start, data_region.max_row + 1):
                if r_idx >= len(sheet_data.cells):
                    continue
                row = sheet_data.cells[r_idx]
                if c_idx >= len(row):
                    continue
                cell = row[c_idx]
                if cell.value is not None and cell.value != "":
                    values.append(cell.value)

            if not values:
                continue

            # SEQUENTIAL_NUMERIC
            seq_pattern = self._detect_sequential_numeric(values, c_idx, col_profile.column_letter, sheet_data.name)
            if seq_pattern:
                patterns.append(seq_pattern)

            # CONSTANT_VALUE
            const_pattern = self._detect_constant_value(values, c_idx, col_profile.column_letter, sheet_data.name)
            if const_pattern:
                patterns.append(const_pattern)

            # REPEATED_VALUE
            rep_pattern = self._detect_repeated_value(values, c_idx, col_profile.column_letter, sheet_data.name)
            if rep_pattern:
                patterns.append(rep_pattern)

            # DATE_SEQUENCE
            date_pattern = self._detect_date_sequence(values, c_idx, col_profile.column_letter, sheet_data.name)
            if date_pattern:
                patterns.append(date_pattern)

            # FORMULA_RELATIVE_PATTERN
            formula_pattern = self._detect_formula_relative_pattern(
                formula_infos, sheet_data.name, c_idx, col_profile.column_letter
            )
            if formula_pattern:
                patterns.append(formula_pattern)

            # DOMINANT_VALUE
            dom_pattern = self._detect_dominant_value(values, c_idx, col_profile.column_letter, sheet_data.name)
            if dom_pattern:
                patterns.append(dom_pattern)

        return patterns

    def _detect_sequential_numeric(
        self, values: list, c_idx: int, col_letter: str, sheet_name: str
    ) -> Optional[PatternDetection]:
        """Detect arithmetic progression in numeric values."""
        numeric_vals: list[float] = []
        for v in values:
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                numeric_vals.append(float(v))
            elif isinstance(v, str) and v.startswith("="):
                continue  # Skip formulas
            else:
                continue

        if len(numeric_vals) < 3:
            return None

        # Check for constant difference
        diffs = []
        for i in range(1, len(numeric_vals)):
            diff = numeric_vals[i] - numeric_vals[i - 1]
            diffs.append(diff)

        # Allow small floating point tolerance
        first_diff = diffs[0]
        is_constant = True
        for d in diffs[1:]:
            if abs(d - first_diff) > 1e-9:
                is_constant = False
                break

        if is_constant and len(numeric_vals) >= 3:
            return PatternDetection(
                pattern_name=PatternType.SEQUENTIAL_NUMERIC.value,
                column_indices=[c_idx],
                sheet=sheet_name,
                evidence=[
                    f"Step: {first_diff}",
                    f"Values: {numeric_vals[:5]}",
                ],
                description=(
                    f"Column {col_letter} contains sequential numeric values "
                    f"with constant step {first_diff}"
                ),
                locations=[PatternLocation(
                    column_index=c_idx,
                    column_letter=col_letter,
                    description=f"Sequential numeric with step {first_diff}",
                    evidence=[f"Step: {first_diff}"],
                )],
            )
        return None

    def _detect_constant_value(
        self, values: list, c_idx: int, col_letter: str, sheet_name: str
    ) -> Optional[PatternDetection]:
        """Detect constant value across column."""
        non_empty: list = [v for v in values if v is not None and v != ""]
        if len(non_empty) < 2:
            return None

        # Check if all values are "equal"
        first = non_empty[0]
        all_same = True
        for v in non_empty[1:]:
            if isinstance(v, str) and v.startswith("="):
                # Formulas not considered equal unless identical
                if v != first:
                    all_same = False
                    break
            elif v != first:
                all_same = False
                break

        if all_same:
            return PatternDetection(
                pattern_name=PatternType.CONSTANT_VALUE.value,
                column_indices=[c_idx],
                sheet=sheet_name,
                evidence=[f"Value: {repr(first)}", f"Count: {len(non_empty)}"],
                description=f"Column {col_letter} has constant value {repr(first)} across {len(non_empty)} cells",
                locations=[PatternLocation(
                    column_index=c_idx,
                    column_letter=col_letter,
                    description=f"Constant value: {repr(first)}",
                    evidence=[f"Value: {repr(first)}"],
                )],
            )
        return None

    def _detect_repeated_value(
        self, values: list, c_idx: int, col_letter: str, sheet_name: str
    ) -> Optional[PatternDetection]:
        """Detect repeated/syclic value pattern."""
        non_empty: list = [v for v in values if v is not None and v != ""]
        if len(non_empty) < 4:
            return None

        # Try to find period
        for period in range(2, min(len(non_empty) // 2 + 1, 10)):
            is_periodic = True
            for i in range(period, len(non_empty)):
                if non_empty[i] != non_empty[i - period]:
                    is_periodic = False
                    break

            if is_periodic:
                cycle = non_empty[:period]
                return PatternDetection(
                    pattern_name=PatternType.REPEATED_VALUE.value,
                    column_indices=[c_idx],
                    sheet=sheet_name,
                    evidence=[
                        f"Period: {period}",
                        f"Cycle: {cycle}",
                        f"Total cells: {len(non_empty)}",
                    ],
                    description=(
                        f"Column {col_letter} repeats every {period} cells: {cycle}"
                    ),
                    locations=[PatternLocation(
                        column_index=c_idx,
                        column_letter=col_letter,
                        description=f"Repeated cycle: {cycle}",
                        evidence=[f"Period: {period}"],
                    )],
                )

        return None

    def _detect_date_sequence(
        self, values: list, c_idx: int, col_letter: str, sheet_name: str
    ) -> Optional[PatternDetection]:
        """Detect daily or weekly date sequence."""
        date_vals: list[date] = []
        for v in values:
            if isinstance(v, (date, datetime)):
                d = v if isinstance(v, date) else v.date()
                date_vals.append(d)
            elif isinstance(v, str) and v.startswith("="):
                continue
            else:
                continue

        if len(date_vals) < 3:
            return None

        # Check daily
        day_diffs = []
        for i in range(1, len(date_vals)):
            diff = (date_vals[i] - date_vals[i - 1]).days
            day_diffs.append(diff)

        # Daily: all diffs == 1
        is_daily = all(d == 1 for d in day_diffs)

        # Weekly: all diffs == 7
        is_weekly = all(d == 7 for d in day_diffs)

        if is_daily:
            return PatternDetection(
                pattern_name=PatternType.DATE_SEQUENCE.value,
                column_indices=[c_idx],
                sheet=sheet_name,
                evidence=[
                    "Interval: DAILY (1 day)",
                    f"Sample dates: {date_vals[:5]}",
                ],
                description=(
                    f"Column {col_letter} contains daily date sequence"
                ),
                locations=[PatternLocation(
                    column_index=c_idx,
                    column_letter=col_letter,
                    description="Daily date sequence",
                    evidence=["Interval: DAILY"],
                )],
            )

        if is_weekly:
            return PatternDetection(
                pattern_name=PatternType.DATE_SEQUENCE.value,
                column_indices=[c_idx],
                sheet=sheet_name,
                evidence=[
                    "Interval: WEEKLY (7 days)",
                    f"Sample dates: {date_vals[:5]}",
                ],
                description=(
                    f"Column {col_letter} contains weekly date sequence"
                ),
                locations=[PatternLocation(
                    column_index=c_idx,
                    column_letter=col_letter,
                    description="Weekly date sequence",
                    evidence=["Interval: WEEKLY"],
                )],
            )

        return None

    def _detect_formula_relative_pattern(
        self,
        formula_infos: list[FormulaInfo],
        sheet_name: str,
        c_idx: int,
        col_letter: str,
    ) -> Optional[PatternDetection]:
        """Detect formula cells that follow relative pattern."""
        sheet_formulas = [
            fi for fi in formula_infos
            if fi.sheet == sheet_name and fi.col == c_idx
        ]

        if len(sheet_formulas) < 2:
            return None

        # Check if formulas are structurally similar with row offset
        # Group by formula structure (strip row numbers)
        from re import sub

        def normalize_formula(formula: str) -> str:
            """Normalize formula by replacing cell refs with placeholder."""
            # Extract row numbers
            def repl_row(m):
                return "R"
            def repl_col(m):
                return "C"
            # Replace A1-style refs with placeholder
            normalized = sub(r'[A-Z]+\d+', 'REF', formula)
            return normalized

        normalized_formulas = [
            (fi, normalize_formula(fi.formula))
            for fi in sheet_formulas
        ]

        # Check if all have same normalized form
        first_norm = normalized_formulas[0][1]
        same_structure = all(
            nf == first_norm for _, nf in normalized_formulas[1:]
        )

        if same_structure and len(sheet_formulas) >= 2:
            sample = sheet_formulas[0]
            return PatternDetection(
                pattern_name=PatternType.FORMULA_RELATIVE_PATTERN.value,
                column_indices=[c_idx],
                sheet=sheet_name,
                evidence=[
                    f"Base formula: {sample.formula}",
                    f"Normalized: {first_norm}",
                    f"Formula count: {len(sheet_formulas)}",
                ],
                description=(
                    f"Column {col_letter} contains {len(sheet_formulas)} "
                    f"formula cells with relative pattern"
                ),
                locations=[PatternLocation(
                    column_index=c_idx,
                    column_letter=col_letter,
                    description="Formula relative pattern",
                    evidence=[f"Base: {sample.formula}"],
                )],
            )

        return None

    def _detect_dominant_value(
        self, values: list, c_idx: int, col_letter: str, sheet_name: str
    ) -> Optional[PatternDetection]:
        """Detect value that appears more than 50% of non-empty cells."""
        non_empty: list = [v for v in values if v is not None and v != ""]
        if len(non_empty) < 2:
            return None

        # Count frequencies
        freq: dict = {}
        for v in non_empty:
            key = repr(v) if not isinstance(v, str) else v
            freq[key] = freq.get(key, 0) + 1

        most_common = max(freq.items(), key=lambda x: x[1])
        dominant_val, dominant_count = most_common

        if dominant_count > len(non_empty) / 2:
            return PatternDetection(
                pattern_name=PatternType.DOMINANT_VALUE.value,
                column_indices=[c_idx],
                sheet=sheet_name,
                evidence=[
                    f"Value: {dominant_val}",
                    f"Frequency: {dominant_count}/{len(non_empty)}",
                    f"Percentage: {dominant_count/len(non_empty)*100:.1f}%",
                ],
                description=(
                    f"Column {col_letter} has dominant value '{dominant_val}' "
                    f"({dominant_count}/{len(non_empty)} = "
                    f"{dominant_count/len(non_empty)*100:.1f}%)"
                ),
                locations=[PatternLocation(
                    column_index=c_idx,
                    column_letter=col_letter,
                    description=f"Dominant value: {dominant_val}",
                    evidence=[f"Count: {dominant_count}/{len(non_empty)}"],
                )],
            )

        return None

    # ------------------------------------------------------------------
    # Relationship detection
    # ------------------------------------------------------------------

    def _detect_relationships(
        self, formula_infos: list[FormulaInfo], sheet_name: str
    ) -> list[StructuralRelationship]:
        """Detect cross-sheet references as CROSS_SHEET_REFERENCE only."""
        relationships: list[StructuralRelationship] = []

        for fi in formula_infos:
            if fi.sheet != sheet_name:
                continue
            for ref in fi.references:
                if "!" in ref:
                    parts = ref.split("!")
                    if len(parts) == 2:
                        target_sheet, target_cell = parts
                        # Only report CROSS_SHEET_REFERENCE, not key relationship
                        relationships.append(StructuralRelationship(
                            source_sheet=sheet_name,
                            source_cell=fi.cell,
                            target_sheet=target_sheet,
                            target_cell=target_cell,
                            reference_type=RelationshipType.CROSS_SHEET_REFERENCE.value,
                            status="DETECTED",
                            evidence=[f"Formula {fi.formula} references {ref}"],
                            risk_note=(
                                "This is a detected cross-sheet reference, not a "
                                "confirmed key relationship. Do not infer foreign key."
                            ),
                        ))

        return relationships

    # ------------------------------------------------------------------
    # Insight generation
    # ------------------------------------------------------------------

    def _generate_insights(
        self,
        sheet_name: str,
        sheet_data: SheetData,
        data_region: DataRegion,
        header_detection: Optional[HeaderDetection],
        columns: list[ColumnProfile],
        formula_profile: Optional[FormulaProfile],
        quality_signals: list[QualitySignal],
        patterns: list[PatternDetection],
        relationships: list[StructuralRelationship],
    ) -> list[Insight]:
        """Generate structured insights from analysis results."""
        insights: list[Insight] = []
        insight_counter = 0

        def next_id() -> str:
            nonlocal insight_counter
            insight_counter += 1
            return f"INS-{insight_counter:04d}"

        # STRUCTURAL insights
        if not data_region.has_data:
            insights.append(Insight(
                insight_id=next_id(),
                type=InsightType.STRUCTURAL.value,
                severity=InsightSeverity.INFO.value,
                sheet=sheet_name,
                location=None,
                description=f"Sheet '{sheet_name}' contains no data",
                evidence=["No non-empty cells found"],
                rule="Data region is empty",
                suggested_action="Add data to sheet",
                source="MIND-050",
            ))

        if header_detection:
            if header_detection.status == HeaderStatus.NOT_DETECTED.value:
                insights.append(Insight(
                    insight_id=next_id(),
                    type=InsightType.STRUCTURAL.value,
                    severity=InsightSeverity.WARNING.value,
                    sheet=sheet_name,
                    location=f"row {header_detection.row_index + 1}",
                    description=f"Header not detected in sheet '{sheet_name}'",
                    evidence=header_detection.evidence,
                    rule="Header detection rule: text density < 0.3",
                    suggested_action="Review header structure",
                    source="MIND-050",
                ))
            elif header_detection.status == HeaderStatus.AMBIGUOUS.value:
                insights.append(Insight(
                    insight_id=next_id(),
                    type=InsightType.STRUCTURAL.value,
                    severity=InsightSeverity.INFO.value,
                    sheet=sheet_name,
                    location=f"row {header_detection.row_index + 1}",
                    description=f"Header detection ambiguous for sheet '{sheet_name}'",
                    evidence=header_detection.evidence,
                    rule="Header detection rule: intermediate text density",
                    suggested_action="Manual header verification recommended",
                    source="MIND-050",
                ))

        if data_region.has_data and data_region.max_row < sheet_data.row_count - 1:
            trailing_empty = sheet_data.row_count - 1 - data_region.max_row
            if trailing_empty > 0:
                insights.append(Insight(
                    insight_id=next_id(),
                    type=InsightType.STRUCTURAL.value,
                    severity=InsightSeverity.INFO.value,
                    sheet=sheet_name,
                    location=f"rows {data_region.max_row + 2}-{sheet_data.row_count}",
                    description=(
                        f"Sheet '{sheet_name}' has {trailing_empty} trailing "
                        f"empty row(s) beyond data region"
                    ),
                    evidence=[f"Data region ends at row {data_region.max_row + 1}"],
                    rule="Physical dimensions exceed data region",
                    suggested_action="Trim empty rows if not needed",
                    source="MIND-050",
                ))

        # DATA_QUALITY insights
        for qs in quality_signals:
            insights.append(Insight(
                insight_id=next_id(),
                type=InsightType.DATA_QUALITY.value,
                severity=qs.severity,
                sheet=qs.sheet,
                location=qs.location,
                description=qs.description,
                evidence=qs.evidence,
                rule=f"MIND-020 {qs.type} detection",
                suggested_action="Review quality signal",
                source="MIND-020",
            ))

        # FORMULA insights
        if formula_profile and formula_profile.formula_count > 0:
            insights.append(Insight(
                insight_id=next_id(),
                type=InsightType.FORMULA.value,
                severity=InsightSeverity.INFO.value,
                sheet=sheet_name,
                location=f"{formula_profile.formula_count} cells",
                description=(
                    f"Sheet '{sheet_name}' contains "
                    f"{formula_profile.formula_count} formula cell(s)"
                ),
                evidence=[
                    f"Distribution: {formula_profile.formula_distribution}",
                ],
                rule="MIND-030 formula detection",
                suggested_action="Review formula cells",
                source="MIND-030",
            ))

            if formula_profile.has_cross_sheet_references:
                for csr in formula_profile.cross_sheet_references:
                    insights.append(Insight(
                        insight_id=next_id(),
                        type=InsightType.RELATIONSHIP.value,
                        severity=InsightSeverity.INFO.value,
                        sheet=sheet_name,
                        location=csr.source_cell,
                        description=(
                            f"Cross-sheet reference: {csr.source_cell} -> "
                            f"{csr.target_sheet}!{csr.target_cell}"
                        ),
                        evidence=[f"Formula references {csr.target_sheet}!{csr.target_cell}"],
                        rule="MIND-030 cross-sheet reference detection",
                        suggested_action="Review cross-sheet reference",
                        source="MIND-030",
                    ))

        if formula_profile and formula_profile.has_unknown_formulas:
            insights.append(Insight(
                insight_id=next_id(),
                type=InsightType.FORMULA.value,
                severity=InsightSeverity.INFO.value,
                sheet=sheet_name,
                location="formula cells",
                description=(
                    f"Sheet '{sheet_name}' contains formula(s) that cannot be "
                    f"classified"
                ),
                evidence=[],
                rule="MIND-030 formula classification",
                suggested_action="Review unknown formula structures",
                source="MIND-030",
            ))

        # PATTERN insights
        for pat in patterns:
            insights.append(Insight(
                insight_id=next_id(),
                type=InsightType.PATTERN.value,
                severity=InsightSeverity.INFO.value,
                sheet=sheet_name,
                location=f"column(s) {pat.column_indices}",
                description=pat.description,
                evidence=pat.evidence,
                rule=f"Pattern detection: {pat.pattern_name}",
                suggested_action="Review detected pattern",
                source="MIND-050",
            ))

        # RELATIONSHIP insights (cross-sheet only, NOT key relationship)
        for rel in relationships:
            insights.append(Insight(
                insight_id=next_id(),
                type=InsightType.RELATIONSHIP.value,
                severity=InsightSeverity.INFO.value,
                sheet=rel.source_sheet,
                location=rel.source_cell,
                description=(
                    f"Cross-sheet reference: {rel.source_cell} -> "
                    f"{rel.target_sheet}!{rel.target_cell}"
                ),
                evidence=rel.evidence,
                rule="MIND-050 cross-sheet reference detection",
                suggested_action="Review reference",
                source="MIND-050",
            ))

        return insights

    # ------------------------------------------------------------------
    # Summary helpers
    # ------------------------------------------------------------------

    def _build_quality_summary(
        self, signals: list[QualitySignal]
    ) -> QualitySummary:
        """Build aggregated quality summary."""
        by_type: dict = {}
        by_severity: dict = {}
        for s in signals:
            by_type[s.type] = by_type.get(s.type, 0) + 1
            by_severity[s.severity] = by_severity.get(s.severity, 0) + 1

        return QualitySummary(
            total_signals=len(signals),
            by_type=by_type,
            by_severity=by_severity,
        )

    # ------------------------------------------------------------------
    # Type helpers
    # ------------------------------------------------------------------

    def _classify_value_type(self, val: Any) -> str:
        """Classify a cell value into a type string."""
        if val is None or val == "":
            return "EMPTY"
        if isinstance(val, bool):
            return "BOOLEAN"
        if isinstance(val, int):
            return "INTEGER"
        if isinstance(val, float):
            return "FLOAT"
        if isinstance(val, (date, datetime)):
            return "DATE_DATETIME"
        if isinstance(val, str):
            if val.startswith("="):
                return "FORMULA"
            return "TEXT"
        return "UNKNOWN"

    def _determine_column_type(self, type_counts: dict) -> str:
        """Determine column type from type distribution."""
        if not type_counts:
            return ColumnType.EMPTY.value

        if len(type_counts) == 1:
            only_type = list(type_counts.keys())[0]
            if only_type == "FORMULA":
                return ColumnType.FORMULA.value
            if only_type == "TEXT":
                return ColumnType.TEXT.value
            if only_type == "INTEGER":
                return ColumnType.INTEGER.value
            if only_type == "FLOAT":
                return ColumnType.FLOAT.value
            if only_type == "BOOLEAN":
                return ColumnType.BOOLEAN.value
            if only_type == "DATE_DATETIME":
                return ColumnType.DATE_DATETIME.value
            if only_type == "EMPTY":
                return ColumnType.EMPTY.value
            return ColumnType.UNKNOWN.value

        # Has multiple types
        # Check if only EMPTY + one other
        types_without_empty = {
            k: v for k, v in type_counts.items() if k != "EMPTY"
        }
        if len(types_without_empty) == 1:
            only = list(types_without_empty.keys())[0]
            if only == "FORMULA":
                return ColumnType.FORMULA.value
            if only == "TEXT":
                return ColumnType.TEXT.value
            if only == "INTEGER":
                return ColumnType.INTEGER.value
            if only == "FLOAT":
                return ColumnType.FLOAT.value
            if only == "BOOLEAN":
                return ColumnType.BOOLEAN.value
            if only == "DATE_DATETIME":
                return ColumnType.DATE_DATETIME.value
            return ColumnType.UNKNOWN.value

        # MIXED
        return ColumnType.MIXED.value

    def _compute_numeric_stats(self, values: list) -> Optional[NumericStats]:
        """Compute numeric statistics for a column."""
        nums: list[float] = []
        for v in values:
            if isinstance(v, bool):
                continue
            if isinstance(v, (int, float)):
                nums.append(float(v))
            elif isinstance(v, str) and v.startswith("="):
                continue  # Skip formulas
            elif isinstance(v, str):
                try:
                    nums.append(float(v))
                except (ValueError, TypeError):
                    continue

        if not nums:
            return None

        return NumericStats(
            min=min(nums),
            max=max(nums),
            sum=sum(nums),
            mean=sum(nums) / len(nums),
            sample_count=len(nums),
        )

    def _compute_date_stats(self, values: list) -> Optional[DateStats]:
        """Compute date statistics for a column."""
        dates: list[date] = []
        for v in values:
            if isinstance(v, (date, datetime)):
                d = v if isinstance(v, date) else v.date()
                dates.append(d)
            elif isinstance(v, str) and v.startswith("="):
                continue
            elif isinstance(v, str):
                continue  # Don't parse strings as dates

        if not dates:
            return None

        return DateStats(
            earliest=min(dates),
            latest=max(dates),
            sample_count=len(dates),
        )

    def _compute_value_distribution(self, values: list) -> Optional[ValueDistribution]:
        """Compute top values distribution for text columns."""
        non_empty: list = [v for v in values if v is not None and v != ""]
        if not non_empty:
            return None

        freq: dict = {}
        for v in non_empty:
            key = v if isinstance(v, str) else repr(v)
            freq[key] = freq.get(key, 0) + 1

        # Sort by frequency
        sorted_vals = sorted(freq.items(), key=lambda x: (-x[1], x[0]))
        top_values = [[k, v] for k, v in sorted_vals[:10]]

        dominant = sorted_vals[0] if sorted_vals else (None, 0)

        return ValueDistribution(
            dominant_value=dominant[0],
            dominant_value_count=dominant[1],
            top_values=top_values,
        )

    # ------------------------------------------------------------------
    # Utility
    # ------------------------------------------------------------------

    def _col_index_to_letter(self, col_idx: int) -> str:
        """Convert 0-based column index to Excel column letter."""
        from openpyxl.utils import get_column_letter
        return get_column_letter(col_idx + 1)

    def close(self) -> None:
        """Close all engines."""
        self._engine.close()
        self._formula_engine.close()
