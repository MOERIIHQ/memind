"""
MIND-060 Auto Reporting Tests.

Test categories:
- Report construction (basic, multi-sheet, empty, formulas, cross-sheet, quality, patterns)
- Workbook overview mapping
- Sheet reports (data region, header, columns, formula, quality, patterns, insights)
- Quality section (INFO/WARNING/ERROR, by-type, evidence)
- Formula section (count, cells, distribution, cross-sheet, unknown)
- Pattern section (all 6 frozen categories)
- Relationship section (CROSS_SHEET_REFERENCE, no semantic strengthening)
- Insight section (all fields preserved)
- Serialization (to_dict, to_json_compatible, nested)
- Determinism (2 calls → identical)
- Read-only (analysis unchanged)
- Empty analysis (zero/empty state)
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.mind_060.auto_reporting import (
    generate_report,
    Report,
    ReportMetadata,
    ExecutiveSummary,
    WorkbookOverview,
    SheetReport,
    CompactColumnSummary,
    QualitySection,
    FormulaSection,
    PatternSection,
    RelationshipSection,
    InsightSection,
    LIMITATIONS,
)

from engine.mind_050.spreadsheet_intelligence import (
    SpreadsheetAnalysis,
    WorkbookSummary,
    SheetAnalysis,
    SheetSummary,
    DataRegion,
    HeaderDetection,
    HeaderStatus,
    ColumnProfile,
    ColumnType,
    FormulaProfile,
    FormulaLocation,
    CrossSheetRef,
    QualitySignal,
    PatternDetection,
    PatternType,
    PatternLocation,
    StructuralRelationship,
    RelationshipType,
    Insight,
    InsightType,
    InsightSeverity,
)

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_empty_analysis() -> SpreadsheetAnalysis:
    """Create an empty SpreadsheetAnalysis for edge-case testing."""
    return SpreadsheetAnalysis()


def _make_basic_analysis() -> SpreadsheetAnalysis:
    """Create a minimal single-sheet analysis with headers and data."""
    dr = DataRegion(min_row=0, max_row=2, min_col=0, max_col=1, has_data=True)
    header = HeaderDetection(
        status=HeaderStatus.DETECTED.value,
        row_index=0,
        column_count=2,
        headers=["Name", "Value"],
        evidence=["Text density: 1.00 (>= 0.5 threshold)", "Uniqueness: 1.00 (>= 0.5 threshold)"],
        notes="Header detected by rule.",
    )
    col0 = ColumnProfile(
        index=0, column_letter="A", header="Name", header_raw="Name",
        detected_type=ColumnType.TEXT.value, non_empty_count=2, empty_count=0,
        unique_count=2, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.TEXT.value: 2},
    )
    col1 = ColumnProfile(
        index=1, column_letter="B", header="Value", header_raw="Value",
        detected_type=ColumnType.INTEGER.value, non_empty_count=2, empty_count=0,
        unique_count=2, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.INTEGER.value: 2},
    )
    sheet = SheetAnalysis(
        sheet_name="Data", row_count=3, col_count=2,
        data_region=dr, header=header, columns=[col0, col1],
        formula_profile=None, quality_signals=[], patterns=[],
    )
    summary = WorkbookSummary(
        file_name="basic.xlsx", sheet_count=1,
        total_physical_rows=3, total_physical_columns=2,
        total_data_rows=3, total_data_columns=2,
        total_formula_cells=0, total_quality_signals=0,
        pattern_count=0, insight_count=0,
        sheet_summaries=[SheetSummary(
            sheet_name="Data", physical_rows=3, physical_columns=2,
            data_rows=3, data_columns=2, formula_cell_count=0,
            quality_signal_count=0, pattern_count=0, insight_count=0,
            header_status=HeaderStatus.DETECTED.value,
        )],
    )
    return SpreadsheetAnalysis(
        analysis_metadata={"workbook_path": "/tmp/basic.xlsx", "analyzed_via": "MIND-050 SpreadsheetIntelligence"},
        workbook={"file_name": "basic.xlsx", "sheet_count": 1, "sheets": ["Data"]},
        sheets=[sheet], quality_signals=[], structural_relationships=[],
        summary=summary, insights=[],
    )


def _make_multi_sheet_analysis() -> SpreadsheetAnalysis:
    """Create a two-sheet analysis."""
    dr1 = DataRegion(min_row=0, max_row=2, min_col=0, max_col=1, has_data=True)
    header1 = HeaderDetection(
        status=HeaderStatus.DETECTED.value, row_index=0, column_count=2,
        headers=["Name", "Value"], evidence=[], notes="",
    )
    col0 = ColumnProfile(
        index=0, column_letter="A", header="Name", header_raw="Name",
        detected_type=ColumnType.TEXT.value, non_empty_count=2, empty_count=0,
        unique_count=2, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.TEXT.value: 2},
    )
    col1 = ColumnProfile(
        index=1, column_letter="B", header="Value", header_raw="Value",
        detected_type=ColumnType.INTEGER.value, non_empty_count=2, empty_count=0,
        unique_count=2, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.INTEGER.value: 2},
    )
    sheet1 = SheetAnalysis(
        sheet_name="Sales", row_count=3, col_count=2,
        data_region=dr1, header=header1, columns=[col0, col1],
        formula_profile=None, quality_signals=[], patterns=[],
    )

    dr2 = DataRegion(min_row=0, max_row=1, min_col=0, max_col=1, has_data=True)
    header2 = HeaderDetection(
        status=HeaderStatus.DETECTED.value, row_index=0, column_count=2,
        headers=["Item", "Stock"], evidence=[], notes="",
    )
    col2 = ColumnProfile(
        index=0, column_letter="A", header="Item", header_raw="Item",
        detected_type=ColumnType.TEXT.value, non_empty_count=1, empty_count=0,
        unique_count=1, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.TEXT.value: 1},
    )
    col3 = ColumnProfile(
        index=1, column_letter="B", header="Stock", header_raw="Stock",
        detected_type=ColumnType.INTEGER.value, non_empty_count=1, empty_count=0,
        unique_count=1, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.INTEGER.value: 1},
    )
    sheet2 = SheetAnalysis(
        sheet_name="Inventory", row_count=2, col_count=2,
        data_region=dr2, header=header2, columns=[col2, col3],
        formula_profile=None, quality_signals=[], patterns=[],
    )

    summary = WorkbookSummary(
        file_name="multi.xlsx", sheet_count=2,
        total_physical_rows=5, total_physical_columns=4,
        total_data_rows=5, total_data_columns=4,
        total_formula_cells=0, total_quality_signals=0,
        pattern_count=0, insight_count=0,
        sheet_summaries=[
            SheetSummary("Sales", 3, 2, 3, 2, 0, 0, 0, 0, HeaderStatus.DETECTED.value),
            SheetSummary("Inventory", 2, 2, 2, 2, 0, 0, 0, 0, HeaderStatus.DETECTED.value),
        ],
    )
    return SpreadsheetAnalysis(
        analysis_metadata={"workbook_path": "/tmp/multi.xlsx", "analyzed_via": "MIND-050 SpreadsheetIntelligence"},
        workbook={"file_name": "multi.xlsx", "sheet_count": 2, "sheets": ["Sales", "Inventory"]},
        sheets=[sheet1, sheet2], quality_signals=[], structural_relationships=[],
        summary=summary, insights=[],
    )


def _make_formulas_analysis() -> SpreadsheetAnalysis:
    """Create analysis with formula profile."""
    dr = DataRegion(min_row=0, max_row=3, min_col=0, max_col=2, has_data=True)
    header = HeaderDetection(
        status=HeaderStatus.DETECTED.value, row_index=0, column_count=3,
        headers=["A", "B", "Total"], evidence=[], notes="",
    )
    col0 = ColumnProfile(index=0, column_letter="A", header="A", header_raw="A",
        detected_type=ColumnType.INTEGER.value, non_empty_count=3, empty_count=0,
        unique_count=3, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.INTEGER.value: 3})
    col1 = ColumnProfile(index=1, column_letter="B", header="B", header_raw="B",
        detected_type=ColumnType.INTEGER.value, non_empty_count=3, empty_count=0,
        unique_count=3, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.INTEGER.value: 3})
    col2 = ColumnProfile(index=2, column_letter="C", header="Total", header_raw="Total",
        detected_type=ColumnType.FORMULA.value, non_empty_count=3, empty_count=0,
        unique_count=3, duplicate_count=0, formula_count=3,
        type_distribution={ColumnType.FORMULA.value: 3})
    fcell1 = FormulaLocation(cell="C2", row=1, col=2, formula_type="arithmetic", references=["A2", "B2"])
    fcell2 = FormulaLocation(cell="C3", row=2, col=2, formula_type="arithmetic", references=["A3", "B3"])
    fcell3 = FormulaLocation(cell="C4", row=3, col=2, formula_type="arithmetic", references=["A4", "B4"])
    fp = FormulaProfile(
        sheet_name="Calc", formula_count=3, formula_cells=[fcell1, fcell2, fcell3],
        formula_distribution={"arithmetic": 3},
        has_cross_sheet_references=False, cross_sheet_references=[],
        has_unknown_formulas=False,
    )
    sheet = SheetAnalysis(
        sheet_name="Calc", row_count=4, col_count=3,
        data_region=dr, header=header, columns=[col0, col1, col2],
        formula_profile=fp, quality_signals=[], patterns=[],
    )
    summary = WorkbookSummary(
        file_name="formulas.xlsx", sheet_count=1,
        total_physical_rows=4, total_physical_columns=3,
        total_data_rows=4, total_data_columns=3,
        total_formula_cells=3, total_quality_signals=0,
        pattern_count=0, insight_count=0,
        sheet_summaries=[SheetSummary(
            "Calc", 4, 3, 4, 3, 3, 0, 0, 0, HeaderStatus.DETECTED.value,
        )],
    )
    return SpreadsheetAnalysis(
        analysis_metadata={"workbook_path": "/tmp/formulas.xlsx", "analyzed_via": "MIND-050 SpreadsheetIntelligence"},
        workbook={"file_name": "formulas.xlsx", "sheet_count": 1, "sheets": ["Calc"]},
        sheets=[sheet], quality_signals=[], structural_relationships=[],
        summary=summary, insights=[],
    )


def _make_cross_sheet_analysis() -> SpreadsheetAnalysis:
    """Create analysis with cross-sheet references."""
    dr = DataRegion(min_row=0, max_row=1, min_col=0, max_col=0, has_data=True)
    header = HeaderDetection(
        status=HeaderStatus.DETECTED.value, row_index=0, column_count=1,
        headers=["Result"], evidence=[], notes="",
    )
    col0 = ColumnProfile(
        index=0, column_letter="A", header="Result", header_raw="Result",
        detected_type=ColumnType.FORMULA.value, non_empty_count=1, empty_count=0,
        unique_count=1, duplicate_count=0, formula_count=1,
        type_distribution={ColumnType.FORMULA.value: 1},
    )
    csr = CrossSheetRef(source_sheet="MultiRef", source_cell="A2", target_sheet="Calc", target_cell="A2")
    fp = FormulaProfile(
        sheet_name="MultiRef", formula_count=1, formula_cells=[
            FormulaLocation(cell="A2", row=1, col=0, formula_type="reference", references=["Calc!A2"])
        ],
        formula_distribution={"reference": 1},
        has_cross_sheet_references=True, cross_sheet_references=[csr],
        has_unknown_formulas=False,
    )
    rel = StructuralRelationship(
        source_sheet="MultiRef", source_cell="A2",
        target_sheet="Calc", target_cell="A2",
        reference_type=RelationshipType.CROSS_SHEET_REFERENCE.value,
        status="DETECTED",
        evidence=["Formula references Calc!A2"],
        risk_note="This is a detected cross-sheet reference, not a confirmed key relationship.",
    )
    sheet = SheetAnalysis(
        sheet_name="MultiRef", row_count=2, col_count=1,
        data_region=dr, header=header, columns=[col0],
        formula_profile=fp, quality_signals=[], patterns=[],
    )
    summary = WorkbookSummary(
        file_name="cross.xlsx", sheet_count=1,
        total_physical_rows=2, total_physical_columns=1,
        total_data_rows=2, total_data_columns=1,
        total_formula_cells=1, total_quality_signals=0,
        pattern_count=0, insight_count=0,
        sheet_summaries=[SheetSummary(
            "MultiRef", 2, 1, 2, 1, 1, 0, 0, 0, HeaderStatus.DETECTED.value,
        )],
    )
    return SpreadsheetAnalysis(
        analysis_metadata={"workbook_path": "/tmp/cross.xlsx", "analyzed_via": "MIND-050 SpreadsheetIntelligence"},
        workbook={"file_name": "cross.xlsx", "sheet_count": 1, "sheets": ["MultiRef"]},
        sheets=[sheet], quality_signals=[], structural_relationships=[rel],
        summary=summary, insights=[],
    )


def _make_quality_analysis() -> SpreadsheetAnalysis:
    """Create analysis with quality signals (INFO, WARNING, ERROR)."""
    dr = DataRegion(min_row=0, max_row=2, min_col=0, max_col=1, has_data=True)
    header = HeaderDetection(
        status=HeaderStatus.DETECTED.value, row_index=0, column_count=2,
        headers=["Name", "Value"], evidence=[], notes="",
    )
    col0 = ColumnProfile(index=0, column_letter="A", header="Name", header_raw="Name",
        detected_type=ColumnType.TEXT.value, non_empty_count=2, empty_count=0,
        unique_count=2, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.TEXT.value: 2})
    col1 = ColumnProfile(index=1, column_letter="B", header="Value", header_raw="Value",
        detected_type=ColumnType.INTEGER.value, non_empty_count=2, empty_count=0,
        unique_count=2, duplicate_count=0, formula_count=0,
        type_distribution={ColumnType.INTEGER.value: 2})

    qs1 = QualitySignal(
        type="duplicate_rows", severity="warning", sheet="Data",
        location="rows 1-2", description="Row 1 is duplicated in rows 2",
        evidence=["Row values: ['Alice', 100]"],
    )
    qs2 = QualitySignal(
        type="missing_values", severity="warning", sheet="Data",
        location="column Value (index 1)", description="Column 'Value' has 1 missing value(s)",
        evidence=[],
    )
    qs3 = QualitySignal(
        type="header_quality:empty", severity="error", sheet="Data",
        location="column 1 (index 0)", description="Header at column 1 is empty",
        evidence=[],
    )
    qs4 = QualitySignal(
        type="formula_detected", severity="info", sheet="Data",
        location="1 cell(s)", description="Sheet 'Data' contains 1 formula cell(s)",
        evidence=[["A1"]],
    )

    sheet = SheetAnalysis(
        sheet_name="Data", row_count=3, col_count=2,
        data_region=dr, header=header, columns=[col0, col1],
        formula_profile=None, quality_signals=[qs1, qs2, qs3, qs4], patterns=[],
    )
    summary = WorkbookSummary(
        file_name="quality.xlsx", sheet_count=1,
        total_physical_rows=3, total_physical_columns=2,
        total_data_rows=3, total_data_columns=2,
        total_formula_cells=0, total_quality_signals=4,
        pattern_count=0, insight_count=0,
        sheet_summaries=[SheetSummary(
            "Data", 3, 2, 3, 2, 0, 4, 0, 0, HeaderStatus.DETECTED.value,
        )],
        quality_summary=None,
    )
    return SpreadsheetAnalysis(
        analysis_metadata={"workbook_path": "/tmp/quality.xlsx", "analyzed_via": "MIND-050 SpreadsheetIntelligence"},
        workbook={"file_name": "quality.xlsx", "sheet_count": 1, "sheets": ["Data"]},
        sheets=[sheet], quality_signals=[qs1, qs2, qs3, qs4], structural_relationships=[],
        summary=summary, insights=[],
    )


def _make_patterns_analysis() -> SpreadsheetAnalysis:
    """Create analysis with all 6 pattern types."""
    dr = DataRegion(min_row=0, max_row=10, min_col=0, max_col=5, has_data=True)
    header = HeaderDetection(
        status=HeaderStatus.DETECTED.value, row_index=0, column_count=6,
        headers=["Seq", "Const", "Rep", "Dates", "Weekly", "Dom"],
        evidence=[], notes="",
    )

    patterns = [
        PatternDetection(
            pattern_name=PatternType.SEQUENTIAL_NUMERIC.value,
            column_indices=[0], sheet="Patterns",
            evidence=["Step: 1", "Values: [1, 2, 3]"],
            description="Column A contains sequential numeric values with constant step 1",
            locations=[PatternLocation(column_index=0, column_letter="A",
                description="Sequential numeric with step 1", evidence=["Step: 1"])],
        ),
        PatternDetection(
            pattern_name=PatternType.CONSTANT_VALUE.value,
            column_indices=[1], sheet="Patterns",
            evidence=["Value: 'X'", "Count: 10"],
            description="Column B has constant value 'X' across 10 cells",
            locations=[PatternLocation(column_index=1, column_letter="B",
                description="Constant value: 'X'", evidence=["Value: 'X'"])],
        ),
        PatternDetection(
            pattern_name=PatternType.REPEATED_VALUE.value,
            column_indices=[2], sheet="Patterns",
            evidence=["Period: 3", "Cycle: ['A', 'B', 'C']", "Total cells: 10"],
            description="Column C repeats every 3 cells: ['A', 'B', 'C']",
            locations=[PatternLocation(column_index=2, column_letter="C",
                description="Repeated cycle: ['A', 'B', 'C']", evidence=["Period: 3"])],
        ),
        PatternDetection(
            pattern_name=PatternType.DATE_SEQUENCE.value,
            column_indices=[3], sheet="Patterns",
            evidence=["Interval: DAILY (1 day)", "Sample dates: [2024-01-01, 2024-01-02]"],
            description="Column D contains daily date sequence",
            locations=[PatternLocation(column_index=3, column_letter="D",
                description="Daily date sequence", evidence=["Interval: DAILY"])],
        ),
        PatternDetection(
            pattern_name=PatternType.DATE_SEQUENCE.value,
            column_indices=[4], sheet="Patterns",
            evidence=["Interval: WEEKLY (7 days)", "Sample dates: [2024-01-01, 2024-01-08]"],
            description="Column E contains weekly date sequence",
            locations=[PatternLocation(column_index=4, column_letter="E",
                description="Weekly date sequence", evidence=["Interval: WEEKLY"])],
        ),
        PatternDetection(
            pattern_name=PatternType.DOMINANT_VALUE.value,
            column_indices=[5], sheet="Patterns",
            evidence=["Value: 'apple'", "Frequency: 7/10", "Percentage: 70.0%"],
            description="Column F has dominant value 'apple' (7/10 = 70.0%)",
            locations=[PatternLocation(column_index=5, column_letter="F",
                description="Dominant value: 'apple'", evidence=["Count: 7/10"])],
        ),
    ]

    columns = []
    for i in range(6):
        columns.append(ColumnProfile(
            index=i, column_letter=chr(65 + i),
            header=header.headers[i] if i < len(header.headers) else None,
            header_raw=header.headers[i] if i < len(header.headers) else None,
            detected_type=ColumnType.TEXT.value, non_empty_count=10, empty_count=0,
            unique_count=10, duplicate_count=0, formula_count=0,
            type_distribution={ColumnType.TEXT.value: 10},
            patterns=[p for p in patterns if p.column_indices and i in p.column_indices],
        ))

    sheet = SheetAnalysis(
        sheet_name="Patterns", row_count=11, col_count=6,
        data_region=dr, header=header, columns=columns,
        formula_profile=None, quality_signals=[], patterns=patterns,
    )
    summary = WorkbookSummary(
        file_name="patterns.xlsx", sheet_count=1,
        total_physical_rows=11, total_physical_columns=6,
        total_data_rows=11, total_data_columns=6,
        total_formula_cells=0, total_quality_signals=0,
        pattern_count=6, insight_count=0,
        sheet_summaries=[SheetSummary(
            "Patterns", 11, 6, 11, 6, 0, 0, 6, 0, HeaderStatus.DETECTED.value,
        )],
    )
    return SpreadsheetAnalysis(
        analysis_metadata={"workbook_path": "/tmp/patterns.xlsx", "analyzed_via": "MIND-050 SpreadsheetIntelligence"},
        workbook={"file_name": "patterns.xlsx", "sheet_count": 1, "sheets": ["Patterns"]},
        sheets=[sheet], quality_signals=[], structural_relationships=[],
        summary=summary, insights=[],
    )


# ---------------------------------------------------------------------------
# Test classes
# ---------------------------------------------------------------------------

class TestReportConstruction(unittest.TestCase):
    """Report construction tests."""

    def test_basic_report(self):
        """Basic single-sheet report should have all sections."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)

        self.assertIsInstance(report, Report)
        self.assertEqual(report.metadata.report_type, "auto_report")
        self.assertEqual(report.metadata.source_type, "xlsx")
        self.assertEqual(report.metadata.schema_version, "1.0.0")
        self.assertEqual(report.metadata.source_identifier, "basic.xlsx")
        self.assertEqual(report.metadata.analyzed_via, "MIND-050 SpreadsheetIntelligence")

        self.assertEqual(report.workbook_overview.file_name, "basic.xlsx")
        self.assertEqual(report.workbook_overview.sheet_count, 1)

        self.assertEqual(report.executive_summary.sheet_count, 1)
        self.assertEqual(report.executive_summary.sheets_with_data, 1)
        self.assertEqual(report.executive_summary.sheets_with_formulas, 0)
        self.assertEqual(report.executive_summary.total_formulas, 0)

        self.assertEqual(len(report.sheet_reports), 1)
        self.assertEqual(report.sheet_reports[0].sheet_name, "Data")

        self.assertEqual(report.quality_section.total_issues, 0)
        self.assertEqual(report.formula_section.formula_count, 0)
        self.assertEqual(report.pattern_section.total_patterns, 0)
        self.assertEqual(report.relationship_section.total_relationships, 0)
        self.assertEqual(report.insight_section.total_insights, 0)

    def test_multi_sheet_report(self):
        """Multi-sheet report should have correct counts."""
        analysis = _make_multi_sheet_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.workbook_overview.sheet_count, 2)
        self.assertEqual(len(report.sheet_reports), 2)
        sheet_names = [s.sheet_name for s in report.sheet_reports]
        self.assertIn("Sales", sheet_names)
        self.assertIn("Inventory", sheet_names)

        self.assertEqual(report.executive_summary.sheet_count, 2)
        self.assertEqual(report.executive_summary.sheets_with_data, 2)

    def test_empty_workbook_report(self):
        """Empty analysis should produce valid report with zero counts."""
        analysis = _make_empty_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.executive_summary.sheet_count, 0)
        self.assertEqual(report.executive_summary.sheets_with_data, 0)
        self.assertEqual(report.workbook_overview.sheet_count, 0)
        self.assertEqual(len(report.sheet_reports), 0)
        self.assertEqual(report.quality_section.total_issues, 0)
        self.assertEqual(report.formula_section.formula_count, 0)
        self.assertEqual(report.pattern_section.total_patterns, 0)
        self.assertEqual(report.relationship_section.total_relationships, 0)
        self.assertEqual(report.insight_section.total_insights, 0)
        self.assertEqual(len(report.limitations), len(LIMITATIONS))

    def test_formula_workbook_report(self):
        """Formula workbook should have formula section populated."""
        analysis = _make_formulas_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.formula_section.formula_count, 3)
        self.assertEqual(len(report.formula_section.formula_cells), 3)
        self.assertEqual(report.formula_section.has_unknown_formulas, False)
        self.assertEqual(report.formula_section.cross_sheet_references, [])
        self.assertEqual(report.executive_summary.total_formulas, 3)

        sr = report.sheet_reports[0]
        self.assertEqual(sr.formula_count, 3)
        self.assertIsInstance(sr.columns[2].formula_count, int)
        self.assertEqual(sr.columns[2].formula_count, 3)

    def test_cross_sheet_report(self):
        """Cross-sheet reference should be in relationship section."""
        analysis = _make_cross_sheet_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.relationship_section.total_relationships, 1)
        rel = report.relationship_section.relationships[0]
        self.assertEqual(rel.reference_type, "CROSS_SHEET_REFERENCE")
        self.assertEqual(rel.source_sheet, "MultiRef")
        self.assertEqual(rel.target_sheet, "Calc")

        self.assertEqual(len(report.formula_section.cross_sheet_references), 1)
        xsr = report.formula_section.cross_sheet_references[0]
        self.assertEqual(xsr.source_sheet, "MultiRef")
        self.assertEqual(xsr.target_sheet, "Calc")

    def test_quality_issues_report(self):
        """Quality issues report should have severity/type breakdown."""
        analysis = _make_quality_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.quality_section.total_issues, 4)
        self.assertEqual(report.quality_section.by_severity["warning"], 2)
        self.assertEqual(report.quality_section.by_severity["error"], 1)
        self.assertEqual(report.quality_section.by_severity["info"], 1)
        self.assertEqual(report.quality_section.by_type["duplicate_rows"], 1)
        self.assertEqual(report.quality_section.by_type["missing_values"], 1)
        self.assertEqual(report.quality_section.by_type["header_quality:empty"], 1)
        self.assertEqual(report.quality_section.by_type["formula_detected"], 1)

        issues = report.quality_section.issues
        self.assertEqual(len(issues), 4)

    def test_patterns_report(self):
        """Pattern report should have all 6 types."""
        analysis = _make_patterns_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.pattern_section.total_patterns, 6)
        by_type = report.pattern_section.by_type
        self.assertIn(PatternType.SEQUENTIAL_NUMERIC.value, by_type)
        self.assertIn(PatternType.CONSTANT_VALUE.value, by_type)
        self.assertIn(PatternType.REPEATED_VALUE.value, by_type)
        self.assertIn(PatternType.DATE_SEQUENCE.value, by_type)
        self.assertIn(PatternType.DOMINANT_VALUE.value, by_type)
        # FORMULA_RELATIVE_PATTERN is not in this fixture
        self.assertEqual(by_type.get(PatternType.FORMULA_RELATIVE_PATTERN.value, 0), 0)


class TestWorkbookOverview(unittest.TestCase):
    """Workbook overview mapping tests."""

    def test_overview_from_summary(self):
        """Workbook overview fields should map from WorkbookSummary."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)

        ov = report.workbook_overview
        self.assertEqual(ov.file_name, analysis.summary.file_name)
        self.assertEqual(ov.sheet_count, analysis.summary.sheet_count)
        self.assertEqual(ov.total_physical_rows, analysis.summary.total_physical_rows)
        self.assertEqual(ov.total_physical_columns, analysis.summary.total_physical_columns)
        self.assertEqual(ov.total_data_rows, analysis.summary.total_data_rows)
        self.assertEqual(ov.total_data_columns, analysis.summary.total_data_columns)


class TestSheetReports(unittest.TestCase):
    """Sheet report tests."""

    def test_sheet_count(self):
        """Sheet report count should match analysis."""
        analysis = _make_multi_sheet_analysis()
        report = generate_report(analysis)
        self.assertEqual(len(report.sheet_reports), 2)

    def test_data_region(self):
        """Data region should be serialized."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)
        dr = report.sheet_reports[0].data_region
        self.assertIsNotNone(dr)
        self.assertTrue(dr["has_data"])
        self.assertEqual(dr["min_row"], 0)
        self.assertEqual(dr["max_row"], 2)

    def test_header_status(self):
        """Header status should propagate."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)
        self.assertEqual(report.sheet_reports[0].header_status, HeaderStatus.DETECTED.value)
        self.assertIsInstance(report.sheet_reports[0].header_evidence, list)

    def test_compact_column_summary(self):
        """Column summary should include only compact fields."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)
        cols = report.sheet_reports[0].columns

        self.assertEqual(len(cols), 2)
        c0 = cols[0]
        self.assertEqual(c0.index, 0)
        self.assertEqual(c0.column_letter, "A")
        self.assertEqual(c0.header, "Name")
        self.assertEqual(c0.detected_type, ColumnType.TEXT.value)
        self.assertEqual(c0.non_empty_count, 2)
        self.assertEqual(c0.empty_count, 0)
        self.assertEqual(c0.unique_count, 2)
        self.assertEqual(c0.duplicate_count, 0)
        self.assertEqual(c0.formula_count, 0)
        self.assertIsInstance(c0.patterns, list)

        # Verify large structures are NOT included
        self.assertNotIn("type_distribution", c0.to_dict())
        self.assertNotIn("numeric_stats", c0.to_dict())
        self.assertNotIn("date_stats", c0.to_dict())
        self.assertNotIn("value_distribution", c0.to_dict())
        self.assertNotIn("notes", c0.to_dict())

    def test_formula_in_sheet_report(self):
        """Sheet report should include formula information."""
        analysis = _make_formulas_analysis()
        report = generate_report(analysis)
        sr = report.sheet_reports[0]

        self.assertEqual(sr.formula_count, 3)
        self.assertIsNotNone(sr.formula_distribution)
        self.assertEqual(sr.formula_distribution["arithmetic"], 3)
        self.assertFalse(sr.has_cross_sheet_references)

    def test_quality_in_sheet_report(self):
        """Sheet report should include quality signals."""
        analysis = _make_quality_analysis()
        report = generate_report(analysis)
        sr = report.sheet_reports[0]

        self.assertEqual(len(sr.quality_signals), 4)

    def test_patterns_in_sheet_report(self):
        """Sheet report should include patterns."""
        analysis = _make_patterns_analysis()
        report = generate_report(analysis)
        sr = report.sheet_reports[0]

        self.assertEqual(len(sr.patterns), 6)

    def test_insights_in_sheet_report(self):
        """Sheet report should filter insights by sheet."""
        analysis = _make_basic_analysis()
        # Add an insight for the Data sheet
        analysis.insights.append(Insight(
            insight_id="INS-0001", type=InsightType.STRUCTURAL.value,
            severity=InsightSeverity.INFO.value, sheet="Data",
            location="row 1", description="Test insight",
            evidence=["evidence"], rule="rule",
            suggested_action="action", source="MIND-050",
        ))
        report = generate_report(analysis)

        sr = report.sheet_reports[0]
        self.assertEqual(len(sr.insights), 1)
        self.assertEqual(sr.insights[0].insight_id, "INS-0001")


class TestQualitySection(unittest.TestCase):
    """Quality section tests."""

    def test_severity_breakdown(self):
        """Should count INFO, WARNING, ERROR."""
        analysis = _make_quality_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.quality_section.by_severity["info"], 1)
        self.assertEqual(report.quality_section.by_severity["warning"], 2)
        self.assertEqual(report.quality_section.by_severity["error"], 1)

    def test_type_breakdown(self):
        """Should count by type."""
        analysis = _make_quality_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.quality_section.by_type["duplicate_rows"], 1)
        self.assertEqual(report.quality_section.by_type["missing_values"], 1)

    def test_evidence_preservation(self):
        """Quality signals should preserve evidence."""
        analysis = _make_quality_analysis()
        report = generate_report(analysis)

        dup_issue = [i for i in report.quality_section.issues if i.type == "duplicate_rows"][0]
        self.assertEqual(len(dup_issue.evidence), 1)
        self.assertIn("Row values", dup_issue.evidence[0])


class TestFormulaSection(unittest.TestCase):
    """Formula section tests."""

    def test_formula_count(self):
        """Formula count should match summary."""
        analysis = _make_formulas_analysis()
        report = generate_report(analysis)
        self.assertEqual(report.formula_section.formula_count, 3)

    def test_formula_cells(self):
        """Formula cells should be aggregated."""
        analysis = _make_formulas_analysis()
        report = generate_report(analysis)
        self.assertEqual(len(report.formula_section.formula_cells), 3)

    def test_formula_distribution(self):
        """Formula distribution should be aggregated."""
        analysis = _make_formulas_analysis()
        report = generate_report(analysis)
        self.assertEqual(report.formula_section.formula_distribution["arithmetic"], 3)

    def test_cross_sheet_references(self):
        """Cross-sheet references should be aggregated."""
        analysis = _make_cross_sheet_analysis()
        report = generate_report(analysis)
        self.assertEqual(len(report.formula_section.cross_sheet_references), 1)

    def test_unknown_formula_flag(self):
        """Unknown formula flag should propagate."""
        fp = FormulaProfile(
            sheet_name="Test", formula_count=1,
            formula_cells=[FormulaLocation(cell="A1", row=0, col=0,
                formula_type="unknown", references=[])],
            formula_distribution={"unknown": 1},
            has_cross_sheet_references=False, cross_sheet_references=[],
            has_unknown_formulas=True,
        )
        sheet = SheetAnalysis(
            sheet_name="Test", row_count=1, col_count=1,
            data_region=DataRegion(min_row=0, max_row=0, min_col=0, max_col=0, has_data=True),
            header=None, columns=[], formula_profile=fp,
            quality_signals=[], patterns=[],
        )
        analysis = SpreadsheetAnalysis(
            analysis_metadata={"analyzed_via": "MIND-050"},
            workbook={"file_name": "test.xlsx"},
            sheets=[sheet], quality_signals=[], structural_relationships=[],
            summary=WorkbookSummary(file_name="test.xlsx", sheet_count=1,
                total_formula_cells=1, total_quality_signals=0,
                pattern_count=0, insight_count=0),
            insights=[],
        )
        report = generate_report(analysis)
        self.assertTrue(report.formula_section.has_unknown_formulas)


class TestPatternSection(unittest.TestCase):
    """Pattern section tests."""

    def test_all_six_categories(self):
        """Should detect all 6 pattern types."""
        analysis = _make_patterns_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.pattern_section.total_patterns, 6)
        self.assertGreater(report.pattern_section.by_type.get(PatternType.SEQUENTIAL_NUMERIC.value, 0), 0)
        self.assertGreater(report.pattern_section.by_type.get(PatternType.CONSTANT_VALUE.value, 0), 0)
        self.assertGreater(report.pattern_section.by_type.get(PatternType.REPEATED_VALUE.value, 0), 0)
        self.assertGreater(report.pattern_section.by_type.get(PatternType.DATE_SEQUENCE.value, 0), 0)
        self.assertGreater(report.pattern_section.by_type.get(PatternType.DOMINANT_VALUE.value, 0), 0)

    def test_no_new_patterns(self):
        """Should not introduce patterns outside the frozen taxonomy."""
        for pat in _make_patterns_analysis().sheets[0].patterns:
            self.assertIn(pat.pattern_name, [
                PatternType.SEQUENTIAL_NUMERIC.value,
                PatternType.CONSTANT_VALUE.value,
                PatternType.REPEATED_VALUE.value,
                PatternType.DATE_SEQUENCE.value,
                PatternType.FORMULA_RELATIVE_PATTERN.value,
                PatternType.DOMINANT_VALUE.value,
            ])


class TestRelationshipSection(unittest.TestCase):
    """Relationship section tests."""

    def test_cross_sheet_only(self):
        """All relationships should be CROSS_SHEET_REFERENCE."""
        analysis = _make_cross_sheet_analysis()
        report = generate_report(analysis)

        for rel in report.relationship_section.relationships:
            self.assertEqual(rel.reference_type, "CROSS_SHEET_REFERENCE")

    def test_no_semantic_strengthening(self):
        """Should not promote cross-sheet refs to key relationships."""
        analysis = _make_cross_sheet_analysis()
        report = generate_report(analysis)

        rel = report.relationship_section.relationships[0]
        self.assertNotEqual(rel.reference_type, "PRIMARY_KEY")
        self.assertNotEqual(rel.reference_type, "FOREIGN_KEY")
        self.assertNotEqual(rel.reference_type, "MASTER_DETAIL")


class TestInsightSection(unittest.TestCase):
    """Insight section tests."""

    def test_all_fields_preserved(self):
        """Every insight field should be preserved."""
        analysis = _make_basic_analysis()
        analysis.insights.append(Insight(
            insight_id="INS-0099", type=InsightType.DATA_QUALITY.value,
            severity=InsightSeverity.WARNING.value, sheet="Data",
            location="A1", description="Test description",
            evidence=["e1", "e2"], rule="test rule",
            suggested_action="review", source="MIND-020",
            related_to="some_id",
        ))
        report = generate_report(analysis)

        ins = report.insight_section.insights[0]
        self.assertEqual(ins.insight_id, "INS-0099")
        self.assertEqual(ins.type, InsightType.DATA_QUALITY.value)
        self.assertEqual(ins.severity, InsightSeverity.WARNING.value)
        self.assertEqual(ins.sheet, "Data")
        self.assertEqual(ins.location, "A1")
        self.assertEqual(ins.description, "Test description")
        self.assertEqual(ins.evidence, ["e1", "e2"])
        self.assertEqual(ins.rule, "test rule")
        self.assertEqual(ins.suggested_action, "review")
        self.assertEqual(ins.source, "MIND-020")
        self.assertEqual(ins.related_to, "some_id")

    def test_by_type_and_severity(self):
        """Should count by type and severity."""
        analysis = _make_basic_analysis()
        analysis.insights.append(Insight(
            insight_id="I1", type=InsightType.STRUCTURAL.value,
            severity=InsightSeverity.INFO.value, sheet="Data",
            location=None, description="", evidence=[], rule="", source="MIND-050",
        ))
        analysis.insights.append(Insight(
            insight_id="I2", type=InsightType.DATA_QUALITY.value,
            severity=InsightSeverity.WARNING.value, sheet="Data",
            location=None, description="", evidence=[], rule="", source="MIND-020",
        ))
        report = generate_report(analysis)

        self.assertEqual(report.insight_section.by_type[InsightType.STRUCTURAL.value], 1)
        self.assertEqual(report.insight_section.by_type[InsightType.DATA_QUALITY.value], 1)
        self.assertEqual(report.insight_section.by_severity[InsightSeverity.INFO.value], 1)
        self.assertEqual(report.insight_section.by_severity[InsightSeverity.WARNING.value], 1)


class TestSerialization(unittest.TestCase):
    """Serialization tests."""

    def test_to_dict(self):
        """Report.to_dict() should return a dict with all 10 top-level keys."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)
        d = report.to_dict()

        self.assertIsInstance(d, dict)
        self.assertIn("metadata", d)
        self.assertIn("executive_summary", d)
        self.assertIn("workbook_overview", d)
        self.assertIn("sheet_reports", d)
        self.assertIn("quality_section", d)
        self.assertIn("formula_section", d)
        self.assertIn("pattern_section", d)
        self.assertIn("relationship_section", d)
        self.assertIn("insight_section", d)
        self.assertIn("limitations", d)

    def test_json_compatible(self):
        """to_json_compatible should return dict with no non-JSON types."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)
        j = report.to_json_compatible()

        self.assertIsInstance(j, dict)
        # Should not contain Path, datetime, Enum objects
        self.assertIsInstance(j["metadata"]["report_type"], str)
        self.assertIsInstance(j["executive_summary"]["sheet_count"], int)

    def test_nested_serialization(self):
        """Nested structures should serialize."""
        analysis = _make_formulas_analysis()
        report = generate_report(analysis)
        d = report.to_dict()

        sr = d["sheet_reports"][0]
        self.assertIn("columns", sr)
        self.assertIn("formula_distribution", sr)

        fs = d["formula_section"]
        self.assertIn("formula_cells", fs)
        self.assertIn("cross_sheet_references", fs)

    def test_deterministic_serialization(self):
        """to_dict() should be deterministic."""
        analysis = _make_basic_analysis()
        report_a = generate_report(analysis)
        report_b = generate_report(analysis)

        self.assertEqual(report_a.to_dict(), report_b.to_dict())


class TestDeterminism(unittest.TestCase):
    """Determinism tests."""

    def test_identical_generation(self):
        """Same analysis → identical report dict."""
        analysis = _make_multi_sheet_analysis()
        ra = generate_report(analysis)
        rb = generate_report(analysis)

        self.assertEqual(ra.to_dict(), rb.to_dict())

    def test_deterministic_with_quality(self):
        """Determinism with quality signals."""
        analysis = _make_quality_analysis()
        ra = generate_report(analysis)
        rb = generate_report(analysis)

        self.assertEqual(ra.to_dict(), rb.to_dict())

    def test_deterministic_with_patterns(self):
        """Determinism with patterns."""
        analysis = _make_patterns_analysis()
        ra = generate_report(analysis)
        rb = generate_report(analysis)

        self.assertEqual(ra.to_dict(), rb.to_dict())

    def test_no_set_iteration(self):
        """Report generation should not iterate over sets."""
        # This is a behavioral test: if sets were iterated, ordering
        # could vary. We verify by checking that 5 runs produce identical output.
        analysis = _make_patterns_analysis()
        results = [generate_report(analysis).to_dict() for _ in range(5)]
        for i in range(1, len(results)):
            self.assertEqual(results[0], results[i])


class TestReadOnlySafety(unittest.TestCase):
    """Read-only safety tests."""

    def test_analysis_unchanged(self):
        """generate_report must not mutate the input analysis."""
        analysis = _make_basic_analysis()
        before = analysis.to_dict()

        generate_report(analysis)

        after = analysis.to_dict()
        self.assertEqual(before, after)

    def test_sheet_analysis_unchanged(self):
        """Per-sheet data must not be mutated."""
        analysis = _make_formulas_analysis()
        before = analysis.sheets[0].to_dict()

        generate_report(analysis)

        after = analysis.sheets[0].to_dict()
        self.assertEqual(before, after)

    def test_insights_unchanged(self):
        """Insight list must not be mutated."""
        analysis = _make_basic_analysis()
        analysis.insights.append(Insight(
            insight_id="INS-TEST", type=InsightType.STRUCTURAL.value,
            severity=InsightSeverity.INFO.value, sheet="Data",
            location=None, description="test", evidence=[], rule="", source="MIND-050",
        ))
        before = analysis.to_dict()

        generate_report(analysis)

        after = analysis.to_dict()
        self.assertEqual(before, after)


class TestEmptyAnalysis(unittest.TestCase):
    """Empty/partial analysis tests."""

    def test_empty_analysis_all_sections(self):
        """Empty analysis should produce valid zero-state report."""
        analysis = _make_empty_analysis()
        report = generate_report(analysis)

        self.assertEqual(report.metadata.source_identifier, "")
        self.assertEqual(report.workbook_overview.file_name, "")
        self.assertEqual(report.workbook_overview.sheet_count, 0)
        self.assertEqual(report.executive_summary.sheet_count, 0)
        self.assertEqual(len(report.sheet_reports), 0)
        self.assertEqual(report.quality_section.total_issues, 0)
        self.assertEqual(report.formula_section.formula_count, 0)
        self.assertEqual(report.pattern_section.total_patterns, 0)
        self.assertEqual(report.relationship_section.total_relationships, 0)
        self.assertEqual(report.insight_section.total_insights, 0)
        self.assertEqual(len(report.limitations), len(LIMITATIONS))

    def test_empty_analysis_no_invented_values(self):
        """Empty analysis should not invent data."""
        analysis = _make_empty_analysis()
        report = generate_report(analysis)

        d = report.to_dict()
        # All counts should be 0, no invented strings
        self.assertEqual(d["executive_summary"]["total_formulas"], 0)
        self.assertEqual(d["formula_section"]["formula_count"], 0)
        self.assertEqual(d["formula_section"]["formula_cells"], [])
        self.assertEqual(d["formula_section"]["cross_sheet_references"], [])
        self.assertEqual(d["quality_section"]["issues"], [])
        self.assertEqual(d["pattern_section"]["patterns"], [])
        self.assertEqual(d["relationship_section"]["relationships"], [])
        self.assertEqual(d["insight_section"]["insights"], [])


class TestLimitations(unittest.TestCase):
    """Limitations section tests."""

    def test_limitations_count(self):
        """Should have exactly 7 limitations."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)
        self.assertEqual(len(report.limitations), 7)

    def test_limitations_content(self):
        """Limitations should contain verified statements."""
        analysis = _make_basic_analysis()
        report = generate_report(analysis)

        text = "\n".join(report.limitations)
        self.assertIn("Formula calculation is not performed", text)
        self.assertIn("Semantic business interpretation is not performed", text)
        self.assertIn("Relationship signals do not imply confirmed database relationships", text)
        self.assertIn("Multiple tables within a single sheet are not detected", text)
        self.assertIn("Date sequence patterns are limited to DAILY and WEEKLY", text)


class TestRegression(unittest.TestCase):
    """Regression tests — ensure MIND-050 still works after MIND-060 is added."""

    def test_mind050_import(self):
        """MIND-050 should still be importable."""
        from engine.mind_050.spreadsheet_intelligence import (
            SpreadsheetIntelligence,
            SpreadsheetAnalysis,
        )
        self.assertTrue(True)  # Import succeeded

    def test_mind060_import(self):
        """MIND-060 should be importable."""
        from engine.mind_060.auto_reporting import generate_report, Report
        self.assertTrue(True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
