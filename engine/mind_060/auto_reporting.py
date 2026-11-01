"""
MeMind Auto Reporting — MIND-060

Deterministic Structured Report Composer.

Transforms MIND-050 SpreadsheetAnalysis into a structured Report.
Read-only. No analysis duplication. No rendering.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from engine.mind_050.spreadsheet_intelligence import (
    SpreadsheetAnalysis,
    QualitySignal,
    PatternDetection,
    StructuralRelationship,
    Insight,
    PatternType,
    InsightType,
    InsightSeverity,
)


# ---------------------------------------------------------------------------
# Domain Models
# ---------------------------------------------------------------------------

@dataclass
class ReportMetadata:
    report_type: str = "auto_report"
    source_type: str = "xlsx"
    source_identifier: str = ""
    analyzed_via: str = ""
    schema_version: str = "1.0.0"

    def to_dict(self) -> dict:
        return {
            "report_type": self.report_type,
            "source_type": self.source_type,
            "source_identifier": self.source_identifier,
            "analyzed_via": self.analyzed_via,
            "schema_version": self.schema_version,
        }


@dataclass
class ExecutiveSummary:
    sheet_count: int = 0
    sheets_with_data: int = 0
    sheets_with_formulas: int = 0
    sheets_with_quality_signals: int = 0
    total_formulas: int = 0
    total_quality_signals: int = 0
    total_patterns: int = 0
    total_insights: int = 0
    total_relationships: int = 0

    def to_dict(self) -> dict:
        return {
            "sheet_count": self.sheet_count,
            "sheets_with_data": self.sheets_with_data,
            "sheets_with_formulas": self.sheets_with_formulas,
            "sheets_with_quality_signals": self.sheets_with_quality_signals,
            "total_formulas": self.total_formulas,
            "total_quality_signals": self.total_quality_signals,
            "total_patterns": self.total_patterns,
            "total_insights": self.total_insights,
            "total_relationships": self.total_relationships,
        }


@dataclass
class WorkbookOverview:
    file_name: str = ""
    sheet_count: int = 0
    total_physical_rows: int = 0
    total_physical_columns: int = 0
    total_data_rows: int = 0
    total_data_columns: int = 0

    def to_dict(self) -> dict:
        return {
            "file_name": self.file_name,
            "sheet_count": self.sheet_count,
            "total_physical_rows": self.total_physical_rows,
            "total_physical_columns": self.total_physical_columns,
            "total_data_rows": self.total_data_rows,
            "total_data_columns": self.total_data_columns,
        }


@dataclass
class CompactColumnSummary:
    index: int = 0
    column_letter: str = ""
    header: Optional[str] = None
    detected_type: str = ""
    non_empty_count: int = 0
    empty_count: int = 0
    unique_count: int = 0
    duplicate_count: int = 0
    formula_count: int = 0
    patterns: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "column_letter": self.column_letter,
            "header": self.header,
            "detected_type": self.detected_type,
            "non_empty_count": self.non_empty_count,
            "empty_count": self.empty_count,
            "unique_count": self.unique_count,
            "duplicate_count": self.duplicate_count,
            "formula_count": self.formula_count,
            "patterns": [p.to_dict() for p in self.patterns],
        }


@dataclass
class SheetReport:
    sheet_name: str = ""
    row_count: int = 0
    col_count: int = 0
    data_region: Optional[dict] = None
    header_status: Optional[str] = None
    header_evidence: list = field(default_factory=list)
    columns: list = field(default_factory=list)
    formula_count: int = 0
    formula_distribution: Optional[dict] = None
    has_cross_sheet_references: bool = False
    cross_sheet_references: list = field(default_factory=list)
    quality_signals: list = field(default_factory=list)
    patterns: list = field(default_factory=list)
    insights: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "sheet_name": self.sheet_name,
            "row_count": self.row_count,
            "col_count": self.col_count,
            "data_region": self.data_region,
            "header_status": self.header_status,
            "header_evidence": self.header_evidence,
            "columns": [c.to_dict() for c in self.columns],
            "formula_count": self.formula_count,
            "formula_distribution": self.formula_distribution,
            "has_cross_sheet_references": self.has_cross_sheet_references,
            "cross_sheet_references": [r.to_dict() for r in self.cross_sheet_references],
            "quality_signals": [q.to_dict() for q in self.quality_signals],
            "patterns": [p.to_dict() for p in self.patterns],
            "insights": [i.to_dict() for i in self.insights],
        }


@dataclass
class QualitySection:
    total_issues: int = 0
    by_severity: dict = field(default_factory=dict)
    by_type: dict = field(default_factory=dict)
    issues: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "total_issues": self.total_issues,
            "by_severity": self.by_severity,
            "by_type": self.by_type,
            "issues": [i.to_dict() for i in self.issues],
        }


@dataclass
class FormulaSection:
    formula_count: int = 0
    formula_cells: list = field(default_factory=list)
    formula_distribution: dict = field(default_factory=dict)
    cross_sheet_references: list = field(default_factory=list)
    has_unknown_formulas: bool = False

    def to_dict(self) -> dict:
        return {
            "formula_count": self.formula_count,
            "formula_cells": [c.to_dict() for c in self.formula_cells],
            "formula_distribution": self.formula_distribution,
            "cross_sheet_references": [r.to_dict() for r in self.cross_sheet_references],
            "has_unknown_formulas": self.has_unknown_formulas,
        }


@dataclass
class PatternSection:
    total_patterns: int = 0
    by_type: dict = field(default_factory=dict)
    patterns: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "total_patterns": self.total_patterns,
            "by_type": self.by_type,
            "patterns": [p.to_dict() for p in self.patterns],
        }


@dataclass
class RelationshipSection:
    total_relationships: int = 0
    relationships: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "total_relationships": self.total_relationships,
            "relationships": [r.to_dict() for r in self.relationships],
        }


@dataclass
class InsightSection:
    total_insights: int = 0
    by_type: dict = field(default_factory=dict)
    by_severity: dict = field(default_factory=dict)
    insights: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "total_insights": self.total_insights,
            "by_type": self.by_type,
            "by_severity": self.by_severity,
            "insights": [i.to_dict() for i in self.insights],
        }


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

@dataclass
class Report:
    metadata: ReportMetadata = field(default_factory=ReportMetadata)
    executive_summary: ExecutiveSummary = field(default_factory=ExecutiveSummary)
    workbook_overview: WorkbookOverview = field(default_factory=WorkbookOverview)
    sheet_reports: list = field(default_factory=list)
    quality_section: QualitySection = field(default_factory=QualitySection)
    formula_section: FormulaSection = field(default_factory=FormulaSection)
    pattern_section: PatternSection = field(default_factory=PatternSection)
    relationship_section: RelationshipSection = field(default_factory=RelationshipSection)
    insight_section: InsightSection = field(default_factory=InsightSection)
    limitations: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "metadata": self.metadata.to_dict(),
            "executive_summary": self.executive_summary.to_dict(),
            "workbook_overview": self.workbook_overview.to_dict(),
            "sheet_reports": [s.to_dict() for s in self.sheet_reports],
            "quality_section": self.quality_section.to_dict(),
            "formula_section": self.formula_section.to_dict(),
            "pattern_section": self.pattern_section.to_dict(),
            "relationship_section": self.relationship_section.to_dict(),
            "insight_section": self.insight_section.to_dict(),
            "limitations": self.limitations,
        }

    def to_json_compatible(self) -> dict:
        return self.to_dict()


# ---------------------------------------------------------------------------
# Limitations (verified, frozen)
# ---------------------------------------------------------------------------

LIMITATIONS: list[str] = [
    "Formula calculation is not performed. Formula results are not computed or verified.",
    "Semantic business interpretation is not performed. No conclusions about business performance, financial health, or operational recommendations are generated.",
    "Unsupported Excel structures may not be fully analyzed. Advanced charts, external links, and complex Excel-specific features are outside scope.",
    "Relationship signals do not imply confirmed database relationships. Cross-sheet references are factual observations, not inferred foreign keys or primary keys.",
    "Multiple tables within a single sheet are not detected.",
    "Date sequence patterns are limited to DAILY and WEEKLY intervals. Monthly, quarterly, and yearly semantics are not detected.",
    "This report is a structured transformation of MIND-050 analysis output. It does not perform independent spreadsheet analysis.",
]


# ---------------------------------------------------------------------------
# Column summary builder
# ---------------------------------------------------------------------------

def _build_column_summary(col: Any) -> CompactColumnSummary:
    """Build a compact column summary from a MIND-050 ColumnProfile."""
    return CompactColumnSummary(
        index=col.index,
        column_letter=col.column_letter,
        header=col.header,
        detected_type=col.detected_type,
        non_empty_count=col.non_empty_count,
        empty_count=col.empty_count,
        unique_count=col.unique_count,
        duplicate_count=col.duplicate_count,
        formula_count=col.formula_count,
        patterns=col.patterns,
    )


# ---------------------------------------------------------------------------
# Primary API
# ---------------------------------------------------------------------------

def generate_report(analysis: SpreadsheetAnalysis) -> Report:
    """
    Transform a SpreadsheetAnalysis into a structured Report.

    Pure deterministic transformation. No file I/O. No analysis duplication.
    Operates on already-computed MIND-050 output.

    Args:
        analysis: SpreadsheetAnalysis from MIND-050 SpreadsheetIntelligence.analyze()

    Returns:
        Report with 10 sections organized from the analysis.
    """
    report = Report()

    # --- Metadata ---
    report.metadata = ReportMetadata(
        report_type="auto_report",
        source_type="xlsx",
        source_identifier=analysis.workbook.get("file_name", ""),
        analyzed_via=analysis.analysis_metadata.get("analyzed_via", ""),
        schema_version="1.0.0",
    )

    # --- Workbook Overview ---
    summary = analysis.summary
    if summary:
        report.workbook_overview = WorkbookOverview(
            file_name=summary.file_name,
            sheet_count=summary.sheet_count,
            total_physical_rows=summary.total_physical_rows,
            total_physical_columns=summary.total_physical_columns,
            total_data_rows=summary.total_data_rows,
            total_data_columns=summary.total_data_columns,
        )

    # --- Executive Summary ---
    sheets_with_data = 0
    sheets_with_formulas = 0
    sheets_with_quality = 0

    for sheet in analysis.sheets:
        if sheet.data_region and sheet.data_region.has_data:
            sheets_with_data += 1
        if sheet.formula_profile is not None and sheet.formula_profile.formula_count > 0:
            sheets_with_formulas += 1
        if sheet.quality_signals:
            sheets_with_quality += 1

    report.executive_summary = ExecutiveSummary(
        sheet_count=len(analysis.sheets) if analysis.sheets else 0,
        sheets_with_data=sheets_with_data,
        sheets_with_formulas=sheets_with_formulas,
        sheets_with_quality_signals=sheets_with_quality,
        total_formulas=summary.total_formula_cells if summary else 0,
        total_quality_signals=summary.total_quality_signals if summary else 0,
        total_patterns=summary.pattern_count if summary else 0,
        total_insights=summary.insight_count if summary else 0,
        total_relationships=len(analysis.structural_relationships),
    )

    # --- Sheet Reports ---
    for sheet in analysis.sheets:
        sr = SheetReport()
        sr.sheet_name = sheet.sheet_name
        sr.row_count = sheet.row_count
        sr.col_count = sheet.col_count

        # Data region
        if sheet.data_region:
            sr.data_region = sheet.data_region.to_dict()

        # Header
        if sheet.header:
            sr.header_status = sheet.header.status
            sr.header_evidence = list(sheet.header.evidence)

        # Columns (compact)
        for col in sheet.columns:
            sr.columns.append(_build_column_summary(col))

        # Formula profile
        if sheet.formula_profile:
            sr.formula_count = sheet.formula_profile.formula_count
            sr.formula_distribution = dict(sheet.formula_profile.formula_distribution)
            sr.has_cross_sheet_references = sheet.formula_profile.has_cross_sheet_references
            sr.cross_sheet_references = list(sheet.formula_profile.cross_sheet_references)

        # Quality signals (per-sheet)
        sr.quality_signals = list(sheet.quality_signals)

        # Patterns (per-sheet)
        sr.patterns = list(sheet.patterns)

        # Insights (filter from top-level by sheet)
        sr.insights = [
            ins for ins in analysis.insights
            if ins.sheet == sheet.sheet_name
        ]

        report.sheet_reports.append(sr)

    # --- Quality Section (from top-level) ---
    qs = QualitySection()
    qs.total_issues = len(analysis.quality_signals)
    for sig in analysis.quality_signals:
        qs.by_severity[sig.severity] = qs.by_severity.get(sig.severity, 0) + 1
        qs.by_type[sig.type] = qs.by_type.get(sig.type, 0) + 1
    qs.issues = list(analysis.quality_signals)
    report.quality_section = qs

    # --- Formula Section (aggregated from per-sheet) ---
    fs = FormulaSection()
    fs.formula_count = summary.total_formula_cells if summary else 0
    for sheet in analysis.sheets:
        if sheet.formula_profile:
            for fc in sheet.formula_profile.formula_cells:
                fs.formula_cells.append(fc)
            for ft, count in sorted(sheet.formula_profile.formula_distribution.items()):
                fs.formula_distribution[ft] = fs.formula_distribution.get(ft, 0) + count
            fs.cross_sheet_references.extend(sheet.formula_profile.cross_sheet_references)
            if sheet.formula_profile.has_unknown_formulas:
                fs.has_unknown_formulas = True
    report.formula_section = fs

    # --- Pattern Section (aggregated from per-sheet) ---
    ps = PatternSection()
    for sheet in analysis.sheets:
        for pat in sheet.patterns:
            ps.patterns.append(pat)
            ps.by_type[pat.pattern_name] = ps.by_type.get(pat.pattern_name, 0) + 1
    ps.total_patterns = len(ps.patterns)
    report.pattern_section = ps

    # --- Relationship Section (from top-level) ---
    rs = RelationshipSection()
    rs.total_relationships = len(analysis.structural_relationships)
    rs.relationships = list(analysis.structural_relationships)
    report.relationship_section = rs

    # --- Insight Section (from top-level) ---
    ins_sec = InsightSection()
    ins_sec.total_insights = len(analysis.insights)
    for ins in analysis.insights:
        ins_sec.by_type[ins.type] = ins_sec.by_type.get(ins.type, 0) + 1
        ins_sec.by_severity[ins.severity] = ins_sec.by_severity.get(ins.severity, 0) + 1
    ins_sec.insights = list(analysis.insights)
    report.insight_section = ins_sec

    # --- Limitations ---
    report.limitations = list(LIMITATIONS)

    return report
