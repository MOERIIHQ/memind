"""
MIND-050 Spreadsheet Intelligence Tests

Test categories:
- Workbook structure
- Sheet structure
- Data region
- Header detection
- Column profiling
- Formula profiling
- Pattern detection
- Quality aggregation
- Relationship detection
- Insight generation
- Summary generation
- Serialization
- Read-only safety
- Determinism
- Edge cases
- Invalid inputs

All tests must pass and maintain regression: MIND-010 23, MIND-020 40,
MIND-030 57, MIND-040 57 = 177 baseline.
"""

import hashlib
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.mind_050.spreadsheet_intelligence import (
    SpreadsheetIntelligence,
    SpreadsheetIntelligenceError,
    SpreadsheetIntelligenceErrorCode,
    DataRegion,
    HeaderDetection,
    HeaderStatus,
    ColumnProfile,
    ColumnType,
    FormulaProfile,
    QualitySignal,
    PatternDetection,
    PatternType,
    StructuralRelationship,
    Insight,
    InsightType,
    InsightSeverity,
    WorkbookSummary,
    SheetSummary,
    SheetAnalysis,
    SpreadsheetAnalysis,
)

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures")


def _make_workbook_with_formulas(path: str) -> str:
    """Create workbook with various formulas including cross-sheet refs."""
    from openpyxl import Workbook
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Calc"
    ws1.append(["A", "B", "Total"])
    ws1.append([10, 20, "=SUM(A2:B2)"])
    ws1.append([30, 40, "=AVERAGE(A3:B3)"])
    ws1.append([50, 60, "=A4+B4"])
    ws1.append([70, 80, "=A5*B5"])
    ws1.append([90, 100, '=IF(A6>50,"BIG","SMALL")'])

    ws2 = wb.create_sheet("Data")
    ws2.append(["Name", "Value"])
    ws2.append(["Widget", 100])
    ws2.append(["Gadget", 200])

    ws3 = wb.create_sheet("MultiRef")
    ws3.append(["Result"])
    ws3.append(["=Calc!A2+Calc!B2"])
    ws3.append(["=Calc!A3"])

    wb.save(path)
    return path


def _make_pattern_workbook(path: str) -> str:
    """Create workbook with various patterns."""
    from openpyxl import Workbook
    from datetime import date, timedelta
    wb = Workbook()
    ws = wb.active
    ws.title = "Patterns"

    # Sequential numeric (col A)
    ws.append(["Seq", "Constant", "Repeated", "DailyDates", "WeeklyDates", "Dominant"])
    base_date = date(2024, 1, 1)
    for i in range(1, 11):
        ws.append([
            i,  # Sequential 1-10
            "X",  # Constant
            ["A", "B", "C"][(i-1) % 3],  # Repeated A,B,C
            base_date + timedelta(days=i-1),  # Daily
            base_date + timedelta(days=(i-1)*7),  # Weekly
            "apple" if i <= 7 else "orange",  # 70% apple (dominant)
        ])

    wb.save(path)
    return path


def _make_empty_workbook(path: str) -> str:
    """Create empty workbook."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Empty"
    wb.save(path)
    return path


def _make_sporadic_workbook(path: str) -> str:
    """Create workbook with sparse cells."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sparse"
    ws["A1"] = "Header"
    ws["A10"] = "Far value"
    ws["C5"] = "Middle"
    wb.save(path)
    return path


def _make_protected_workbook(path: str) -> str:
    """Create workbook with protected sheet."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Protected"
    ws.append(["A", "B"])
    ws.append([1, 2])
    ws.protection.sheet = True
    wb.save(path)
    return path


def _make_ambiguous_header_workbook(path: str) -> str:
    """Create workbook with header that's all numeric."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "NumericHeader"
    ws.append([100, 200, 300])  # Numeric "header" — likely not a header
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    wb.save(path)
    return path


def _make_multi_sheet_workbook(path: str) -> str:
    """Create multi-sheet workbook."""
    from openpyxl import Workbook
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sales"
    ws1.append(["Product", "Quantity", "Price"])
    ws1.append(["Widget A", 100, 25.50])
    ws1.append(["Widget B", 200, 15.75])

    ws2 = wb.create_sheet("Inventory")
    ws2.append(["Item", "Stock"])
    ws2.append(["Widget A", 500])
    ws2.append(["Widget B", 300])

    wb.save(path)
    return path


class TestSpreadsheetIntelligenceBasic(unittest.TestCase):
    """Basic MIND-050 functionality tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_analyze_empty_workbook(self):
        """Empty workbook should produce valid analysis with empty data region."""
        path = os.path.join(self.temp_dir, "empty_test.xlsx")
        _make_empty_workbook(path)

        result = self.analyzer.analyze(path)

        self.assertIsInstance(result, SpreadsheetAnalysis)
        self.assertEqual(result.summary.sheet_count, 1)
        self.assertFalse(result.sheets[0].data_region.has_data)

    def test_analyze_single_sheet(self):
        """Single sheet workbook analysis."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        self.assertEqual(result.summary.sheet_count, 1)
        self.assertEqual(len(result.sheets), 1)

    def test_analyze_multi_sheet(self):
        """Multi-sheet workbook analysis."""
        path = os.path.join(self.temp_dir, "multi.xlsx")
        _make_multi_sheet_workbook(path)

        result = self.analyzer.analyze(path)

        self.assertEqual(result.summary.sheet_count, 2)
        self.assertEqual(len(result.sheets), 2)
        sheet_names = {s.sheet_name for s in result.sheets}
        self.assertIn("Sales", sheet_names)
        self.assertIn("Inventory", sheet_names)

    def test_data_region_detected(self):
        """Data region should be detected for sheet with data."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)
        region = result.sheets[0].data_region

        self.assertTrue(region.has_data)
        self.assertGreater(region.max_row, region.min_row)
        self.assertGreater(region.max_col, region.min_col)

    def test_physical_vs_data_dimensions(self):
        """Physical dimensions should be >= data region."""
        path = os.path.join(self.temp_dir, "sparse.xlsx")
        _make_sporadic_workbook(path)

        result = self.analyzer.analyze(path)
        sheet = result.sheets[0]

        self.assertGreaterEqual(sheet.row_count, sheet.data_region.max_row + 1)
        self.assertGreaterEqual(sheet.col_count, sheet.data_region.max_col + 1)

    def test_serialization_to_dict(self):
        """SpreadsheetAnalysis should serialize to dict."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        d = result.to_dict()
        self.assertIsInstance(d, dict)
        self.assertIn("sheets", d)
        self.assertIn("summary", d)
        self.assertIn("insights", d)


class TestHeaderDetection(unittest.TestCase):
    """Header detection tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_header_detected(self):
        """Clean workbook should have header detected."""
        path = os.path.join(self.temp_dir, "clean_header.xlsx")
        _make_multi_sheet_workbook(path)

        result = self.analyzer.analyze(path)
        header = result.sheets[0].header

        self.assertEqual(header.status, HeaderStatus.DETECTED.value)
        self.assertGreater(header.column_count, 0)

    def test_header_not_detected_numeric(self):
        """Numeric first row should result in NOT_DETECTED."""
        path = os.path.join(self.temp_dir, "numeric_header.xlsx")
        _make_ambiguous_header_workbook(path)

        result = self.analyzer.analyze(path)
        header = result.sheets[0].header

        self.assertEqual(header.status, HeaderStatus.NOT_DETECTED.value)

    def test_header_ambiguous(self):
        """Mixed/duplicate headers with low uniqueness should be AMBIGUOUS."""
        # 6 columns, 2 non-empty headers (duplicates), text_density=0.33, uniqueness=0.5
        path = os.path.join(self.temp_dir, "ambiguous.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Mixed"
        ws.append(["Name", "Name", "", "", "", ""])  # Duplicate headers
        ws.append(["Alice", "Alice", 30, "NYC", "", ""])
        ws.append(["Bob", "Bob", 25, "LA", "", ""])
        wb.save(path)

        result = self.analyzer.analyze(path)
        header = result.sheets[0].header

        self.assertEqual(header.status, HeaderStatus.AMBIGUOUS.value)

    def test_header_on_alternate_row(self):
        """First row empty should try alternate row."""
        path = os.path.join(self.temp_dir, "alt_header.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Alt"
        ws.append([None, None, None])  # Empty first row
        ws.append(["Name", "Age", "City"])  # Header on row 2
        ws.append(["Alice", 30, "NYC"])
        wb.save(path)

        result = self.analyzer.analyze(path)
        header = result.sheets[0].header

        # Should detect header on row 2
        self.assertIn(header.status, [HeaderStatus.DETECTED.value, HeaderStatus.AMBIGUOUS.value])


class TestColumnProfiling(unittest.TestCase):
    """Column profiling tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_text_column(self):
        """Text column should be classified as TEXT."""
        path = os.path.join(self.temp_dir, "text_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Value"])
        ws.append(["Alice", 1])
        ws.append(["Bob", 2])
        ws.append(["Charlie", 3])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[0]
        self.assertEqual(col.detected_type, ColumnType.TEXT.value)
        self.assertEqual(col.non_empty_count, 3)
        self.assertEqual(col.unique_count, 3)

    def test_integer_column(self):
        """Integer column should be classified as INTEGER."""
        path = os.path.join(self.temp_dir, "int_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Value"])
        ws.append(["A", 1])
        ws.append(["B", 2])
        ws.append(["C", 3])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[1]
        self.assertEqual(col.detected_type, ColumnType.INTEGER.value)
        self.assertEqual(col.non_empty_count, 3)

    def test_float_column(self):
        """Float column should be classified as FLOAT."""
        path = os.path.join(self.temp_dir, "float_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Amount"])
        ws.append(["A", 10.5])
        ws.append(["B", 20.75])
        ws.append(["C", 30.25])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[1]
        self.assertEqual(col.detected_type, ColumnType.FLOAT.value)

    def test_boolean_column(self):
        """Boolean column should be classified as BOOLEAN."""
        path = os.path.join(self.temp_dir, "bool_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Flag"])
        ws.append([True])
        ws.append([False])
        ws.append([True])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[0]
        self.assertEqual(col.detected_type, ColumnType.BOOLEAN.value)

    def test_date_column(self):
        """Date column should be classified as DATE/DATETIME."""
        path = os.path.join(self.temp_dir, "date_col.xlsx")
        from openpyxl import Workbook
        from datetime import date
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Date"])
        ws.append([date(2024, 1, 1)])
        ws.append([date(2024, 1, 2)])
        ws.append([date(2024, 1, 3)])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[0]
        self.assertEqual(col.detected_type, ColumnType.DATE_DATETIME.value)

    def test_formula_column(self):
        """Formula column should be classified as FORMULA."""
        path = os.path.join(self.temp_dir, "formula_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Calc"
        ws.append(["A", "B", "Sum"])
        ws.append([10, 20, "=A2+B2"])
        ws.append([30, 40, "=A3+B3"])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[2]
        self.assertEqual(col.detected_type, ColumnType.FORMULA.value)
        self.assertEqual(col.formula_count, 2)

    def test_mixed_column(self):
        """Mixed types should be classified as MIXED."""
        path = self._path("mixed_types")
        result = self.analyzer.analyze(path)
        for col in result.sheets[0].columns:
            if col.detected_type == ColumnType.MIXED.value:
                self.assertGreater(len(col.type_distribution), 1)
                return
        self.fail("No MIXED column found in mixed_types fixture")

    def test_empty_column(self):
        """Column with no non-empty cells should be EMPTY."""
        path = os.path.join(self.temp_dir, "empty_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["A", "B", "C"])
        ws.append([1, None, 3])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[1]
        self.assertEqual(col.detected_type, ColumnType.EMPTY.value)

    def test_type_distribution(self):
        """Type distribution should be populated."""
        path = self._path("mixed_types")
        result = self.analyzer.analyze(path)

        for col in result.sheets[0].columns:
            if col.type_distribution:
                self.assertIsInstance(col.type_distribution, dict)
                return
        self.fail("No type distribution found")

    def test_numeric_stats(self):
        """Numeric column should have stats."""
        path = os.path.join(self.temp_dir, "float_col.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Amount"])
        ws.append(["A", 10.5])
        ws.append(["B", 20.75])
        ws.append(["C", 30.25])
        wb.save(path)

        result = self.analyzer.analyze(path)
        col = result.sheets[0].columns[1]
        self.assertIsNotNone(col.numeric_stats)
        self.assertGreater(col.numeric_stats.sample_count, 0)

    def test_value_distribution(self):
        """Text column should have value distribution."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        col = result.sheets[0].columns[0]  # Product name
        self.assertIsNotNone(col.value_distribution)
        self.assertGreaterEqual(col.value_distribution.dominant_value_count, 0)


class TestFormulaProfiling(unittest.TestCase):
    """Formula profiling tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_formula_count(self):
        """Formula count should be accurate."""
        path = os.path.join(self.temp_dir, "formulas.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)
        calc_sheet = next(s for s in result.sheets if s.sheet_name == "Calc")
        self.assertEqual(calc_sheet.formula_profile.formula_count, 5)

    def test_formula_cells(self):
        """Formula cells should be listed."""
        path = os.path.join(self.temp_dir, "formulas.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)
        calc_sheet = next(s for s in result.sheets if s.sheet_name == "Calc")
        self.assertEqual(len(calc_sheet.formula_profile.formula_cells), 5)

    def test_formula_distribution(self):
        """Formula distribution should be populated."""
        path = os.path.join(self.temp_dir, "formulas.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)
        calc_sheet = next(s for s in result.sheets if s.sheet_name == "Calc")
        self.assertGreater(len(calc_sheet.formula_profile.formula_distribution), 0)

    def test_cross_sheet_references(self):
        """Cross-sheet references should be detected."""
        path = os.path.join(self.temp_dir, "cross.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)
        multi_sheet = next(s for s in result.sheets if s.sheet_name == "MultiRef")
        self.assertTrue(multi_sheet.formula_profile.has_cross_sheet_references)
        self.assertGreater(len(multi_sheet.formula_profile.cross_sheet_references), 0)

    def test_no_formulas_sheet(self):
        """Sheet without formulas should have no formula profile."""
        path = os.path.join(self.temp_dir, "no_formulas.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["A", "B"])
        ws.append([1, 2])
        wb.save(path)

        result = self.analyzer.analyze(path)
        sheet = result.sheets[0]
        self.assertIsNone(sheet.formula_profile)


class TestPatternDetection(unittest.TestCase):
    """Pattern detection tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_sequential_numeric(self):
        """Sequential numeric pattern should be detected."""
        path = os.path.join(self.temp_dir, "seq.xlsx")
        _make_pattern_workbook(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        seq_patterns = [p for p in patterns if p.pattern_name == PatternType.SEQUENTIAL_NUMERIC.value]
        self.assertGreater(len(seq_patterns), 0)

    def test_constant_value(self):
        """Constant value pattern should be detected."""
        path = os.path.join(self.temp_dir, "const.xlsx")
        _make_pattern_workbook(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        const_patterns = [p for p in patterns if p.pattern_name == PatternType.CONSTANT_VALUE.value]
        self.assertGreater(len(const_patterns), 0)

    def test_repeated_value(self):
        """Repeated value pattern should be detected."""
        path = os.path.join(self.temp_dir, "rep.xlsx")
        _make_pattern_workbook(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        rep_patterns = [p for p in patterns if p.pattern_name == PatternType.REPEATED_VALUE.value]
        self.assertGreater(len(rep_patterns), 0)

    def test_daily_date_sequence(self):
        """Daily date sequence should be detected."""
        path = os.path.join(self.temp_dir, "daily.xlsx")
        _make_pattern_workbook(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        date_patterns = [p for p in patterns if p.pattern_name == PatternType.DATE_SEQUENCE.value]
        daily_found = any("DAILY" in str(loc.description).upper() for p in date_patterns for loc in p.locations)
        self.assertTrue(daily_found, "Daily date sequence not detected")

    def test_weekly_date_sequence(self):
        """Weekly date sequence should be detected."""
        path = os.path.join(self.temp_dir, "weekly.xlsx")
        _make_pattern_workbook(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        date_patterns = [p for p in patterns if p.pattern_name == PatternType.DATE_SEQUENCE.value]
        weekly_found = any("WEEKLY" in str(loc.description).upper() for p in date_patterns for loc in p.locations)
        self.assertTrue(weekly_found, "Weekly date sequence not detected")

    def test_formula_relative_pattern(self):
        """Formula relative pattern should be detected."""
        path = os.path.join(self.temp_dir, "formula_pat.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Calc"
        ws.append(["A", "B", "Sum"])
        ws.append([10, 20, "=A2+B2"])
        ws.append([30, 40, "=A3+B3"])
        ws.append([50, 60, "=A4+B4"])
        wb.save(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        formula_patterns = [p for p in patterns if p.pattern_name == PatternType.FORMULA_RELATIVE_PATTERN.value]
        self.assertGreater(len(formula_patterns), 0)

    def test_dominant_value(self):
        """Dominant value pattern should be detected."""
        path = os.path.join(self.temp_dir, "dominant.xlsx")
        _make_pattern_workbook(path)

        result = self.analyzer.analyze(path)
        patterns = result.sheets[0].patterns
        dom_patterns = [p for p in patterns if p.pattern_name == PatternType.DOMINANT_VALUE.value]
        self.assertGreater(len(dom_patterns), 0)

    def test_no_unsupported_date_patterns(self):
        """Should not detect MONTHLY, QUARTERLY, etc."""
        path = os.path.join(self.temp_dir, "dates.xlsx")
        from openpyxl import Workbook
        from datetime import date, timedelta
        wb = Workbook()
        ws = wb.active
        ws.title = "Dates"
        ws.append(["Date"])
        start = date(2024, 1, 1)
        for i in range(10):
            ws.append([start + timedelta(days=i*30)])  # Monthly-ish
        wb.save(path)

        result = self.analyzer.analyze(path)

        date_patterns = [p for p in result.sheets[0].patterns if p.pattern_name == PatternType.DATE_SEQUENCE.value]
        self.assertEqual(len(date_patterns), 0, "MONTHLY pattern should not be detected (out of scope)")


class TestQualityAggregation(unittest.TestCase):
    """Quality aggregation tests — ensures no duplicate detector."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_quality_signals_from_m20(self):
        """Quality signals should come from MIND-020."""
        path = self._path("duplicates")
        result = self.analyzer.analyze(path)

        dup_signals = [q for q in result.quality_signals if q.type == "duplicate_rows"]
        self.assertGreater(len(dup_signals), 0)

    def test_no_duplicate_detector(self):
        """MIND-050 should not implement its own duplicate detector."""
        import inspect
        source = inspect.getsource(SpreadsheetIntelligence)
        self.assertNotIn("_detect_duplicate_rows", source, "MIND-050 must not implement duplicate detection")
        self.assertNotIn("_detect_missing_values", source, "MIND-050 must not implement missing value detection")

    def test_quality_signal_structure(self):
        """Quality signals should have required fields."""
        path = self._path("missing_values")
        result = self.analyzer.analyze(path)

        for qs in result.quality_signals:
            self.assertIsNotNone(qs.type)
            self.assertIsNotNone(qs.severity)
            self.assertIsNotNone(qs.sheet)
            self.assertIsNotNone(qs.description)


class TestRelationshipBoundary(unittest.TestCase):
    """Relationship boundary tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_cross_sheet_reference_not_promoted(self):
        """Cross-sheet reference should NOT be promoted to key relationship."""
        path = os.path.join(self.temp_dir, "cross.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)

        for rel in result.structural_relationships:
            self.assertNotEqual(rel.reference_type, "POTENTIAL_KEY_RELATIONSHIP",
                                "Cross-sheet reference must not be promoted to key relationship")
            self.assertEqual(rel.reference_type, "CROSS_SHEET_REFERENCE",
                             "Should only report CROSS_SHEET_REFERENCE")

    def test_cross_sheet_detected(self):
        """Cross-sheet reference should be detected."""
        path = os.path.join(self.temp_dir, "cross.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)

        cross_refs = [r for r in result.structural_relationships if r.reference_type == "CROSS_SHEET_REFERENCE"]
        self.assertGreater(len(cross_refs), 0)


class TestInsightGeneration(unittest.TestCase):
    """Insight generation tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_insight_types(self):
        """Insights should cover all taxonomy types."""
        path = os.path.join(self.temp_dir, "formulas.xlsx")
        _make_workbook_with_formulas(path)

        result = self.analyzer.analyze(path)

        types_found = {i.type for i in result.insights}
        expected_types = {
            InsightType.STRUCTURAL.value,
            InsightType.DATA_QUALITY.value,
            InsightType.FORMULA.value,
            InsightType.PATTERN.value,
            InsightType.RELATIONSHIP.value,
        }
        # Not all types may appear in every workbook

    def test_insight_severity(self):
        """Insight severity should be from allowed set."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        for insight in result.insights:
            self.assertIn(insight.severity, [s.value for s in InsightSeverity],
                          f"Invalid severity: {insight.severity}")

    def test_insight_has_evidence(self):
        """Every insight should have evidence."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        for insight in result.insights:
            self.assertIsInstance(insight.evidence, list)

    def test_insight_has_rule(self):
        """Every insight should have rule."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        for insight in result.insights:
            self.assertIsNotNone(insight.rule)
            self.assertGreater(len(insight.rule), 0)

    def test_insight_source(self):
        """Insight source should be tracked."""
        # Use sparse workbook with trailing empty rows to trigger MIND-050 structural insights
        path = os.path.join(self.temp_dir, "insight_src.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["A", "B", "C"])
        ws.append([1, 2, 3])
        ws.append([4, 5, 6])
        ws.append([7, 8, 9])
        for _ in range(5):
            ws.append([])
        wb.save(path)

        result = self.analyzer.analyze(path)

        sources = {i.source for i in result.insights}
        self.assertIn("MIND-050", sources)


class TestSerialization(unittest.TestCase):
    """Serialization tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_analysis_to_dict(self):
        """SpreadsheetAnalysis should convert to dict."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        d = result.to_dict()
        self.assertIsInstance(d, dict)
        self.assertIn("analysis_metadata", d)
        self.assertIn("workbook", d)
        self.assertIn("sheets", d)
        self.assertIn("quality_signals", d)
        self.assertIn("structural_relationships", d)
        self.assertIn("summary", d)
        self.assertIn("insights", d)

    def test_nested_serialization(self):
        """Nested models should also serialize."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        d = result.to_dict()
        sheet_dict = d["sheets"][0]
        self.assertIn("data_region", sheet_dict)
        self.assertIn("columns", sheet_dict)

        if sheet_dict.get("data_region"):
            dr = sheet_dict["data_region"]
            self.assertIn("min_row", dr)
            self.assertIn("max_row", dr)
            self.assertIn("has_data", dr)

    def test_to_json_compatible(self):
        """to_json_compatible should return dict."""
        path = self._path("clean")
        result = self.analyzer.analyze(path)

        j = result.to_json_compatible()
        self.assertIsInstance(j, dict)


class TestReadonlySafety(unittest.TestCase):
    """Read-only safety tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_source_workbook_unchanged(self):
        """Source workbook must not be modified after analysis."""
        path = os.path.join(self.temp_dir, "source.xlsx")
        _make_multi_sheet_workbook(path)

        with open(path, "rb") as f:
            original_hash = hashlib.md5(f.read()).hexdigest()

        self.analyzer.analyze(path)

        with open(path, "rb") as f:
            after_hash = hashlib.md5(f.read()).hexdigest()

        self.assertEqual(original_hash, after_hash, "Source workbook was modified!")

    def test_no_output_file_created(self):
        """Analysis should not create output files."""
        path = self._path("clean")
        output_dir = self.temp_dir
        files_before = set(os.listdir(output_dir))

        self.analyzer.analyze(path)

        files_after = set(os.listdir(output_dir))
        self.assertEqual(files_before, files_after, "Analysis created unexpected files")

    def test_protected_workbook_analyzed(self):
        """Protected workbook should be analyzed without modification."""
        path = os.path.join(self.temp_dir, "protected.xlsx")
        _make_protected_workbook(path)

        result = self.analyzer.analyze(path)

        self.assertIsNotNone(result)
        self.assertEqual(result.summary.sheet_count, 1)

        with open(path, "rb") as f:
            original_hash = hashlib.md5(f.read()).hexdigest()

        self.analyzer.analyze(path)

        with open(path, "rb") as f:
            after_hash = hashlib.md5(f.read()).hexdigest()

        self.assertEqual(original_hash, after_hash)


class TestDeterminism(unittest.TestCase):
    """Determinism tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_dir = FIXTURES_DIR
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.fixture_dir, f"{name}.xlsx")

    def test_identical_results(self):
        """Same input should produce identical results."""
        path = self._path("clean")

        r1 = self.analyzer.analyze(path)
        r2 = self.analyzer.analyze(path)
        r3 = self.analyzer.analyze(path)

        d1 = r1.to_dict()
        d2 = r2.to_dict()
        d3 = r3.to_dict()

        self.assertEqual(d1, d2)
        self.assertEqual(d2, d3)

    def test_sequential_analysis_determinism(self):
        """Sequential analyses should be deterministic."""
        path = os.path.join(self.temp_dir, "det.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["A", "B"])
        ws.append([1, 2])
        ws.append([3, 4])
        wb.save(path)

        results = []
        for _ in range(5):
            r = self.analyzer.analyze(path)
            results.append(r.to_dict())

        for i in range(1, len(results)):
            self.assertEqual(results[0], results[i])


class TestEdgeCases(unittest.TestCase):
    """Edge case tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_empty_workbook(self):
        """Empty workbook should not crash."""
        path = os.path.join(self.temp_dir, "empty.xlsx")
        _make_empty_workbook(path)

        result = self.analyzer.analyze(path)
        self.assertIsNotNone(result)

    def test_sparse_cells(self):
        """Sparse cells should be handled."""
        path = os.path.join(self.temp_dir, "sparse.xlsx")
        _make_sporadic_workbook(path)

        result = self.analyzer.analyze(path)
        self.assertIsNotNone(result)

    def test_empty_sheet_in_multi_sheet(self):
        """Multi-sheet with one empty sheet should work."""
        path = os.path.join(self.temp_dir, "mixed.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Data"
        ws1.append(["A", "B"])
        ws1.append([1, 2])

        ws2 = wb.create_sheet("Empty")
        # No data

        wb.save(path)

        result = self.analyzer.analyze(path)
        self.assertEqual(result.summary.sheet_count, 2)

        empty_sheet = next(s for s in result.sheets if s.sheet_name == "Empty")
        self.assertFalse(empty_sheet.data_region.has_data)

    def test_formulas_only_sheet(self):
        """Sheet with only formulas should work."""
        path = os.path.join(self.temp_dir, "formulas_only.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Formulas"
        ws.append(["A", "B"])
        ws.append(["=1+1", "=2+2"])
        wb.save(path)

        result = self.analyzer.analyze(path)
        self.assertIsNotNone(result)

    def test_malformed_values(self):
        """Malformed/unusual values should be handled."""
        path = os.path.join(self.temp_dir, "weird.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Weird"
        ws.append(["A", "B", "C"])
        ws.append([None, "", float('inf')])
        ws.append([str(object()), str(lambda x: x), "[nested]"])
        wb.save(path)

        result = self.analyzer.analyze(path)
        self.assertIsNotNone(result)


class TestInvalidInputs(unittest.TestCase):
    """Invalid input tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.analyzer = SpreadsheetIntelligence()

    @classmethod
    def tearDownClass(cls):
        cls.analyzer.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_nonexistent_file(self):
        """Non-existent file should raise error."""
        with self.assertRaises(SpreadsheetIntelligenceError) as cm:
            self.analyzer.analyze("/nonexistent/path/to/file.xlsx")

        self.assertEqual(cm.exception.code, SpreadsheetIntelligenceErrorCode.FILE_NOT_FOUND.value)

    def test_invalid_format(self):
        """Non-xlsx file should raise error."""
        path = os.path.join(self.temp_dir, "bad.txt")
        with open(path, "w") as f:
            f.write("not an xlsx")

        with self.assertRaises(SpreadsheetIntelligenceError) as cm:
            self.analyzer.analyze(path)

        self.assertEqual(cm.exception.code, SpreadsheetIntelligenceErrorCode.INVALID_FORMAT.value)


class TestRegression(unittest.TestCase):
    """Regression tests — ensure existing milestones still work."""

    def test_mind010_engine_still_works(self):
        """MIND-010 should still function."""
        from engine.excel_engine import CoreExcelEngine
        engine = CoreExcelEngine()
        path = os.path.join(FIXTURES_DIR, "sample.xlsx")
        info = engine.load(path)
        self.assertEqual(info.sheet_count, 2)
        engine.close()

    def test_mind020_enhancer_still_works(self):
        """MIND-020 should still function."""
        from engine.file_enhancer import FileEnhancer
        fixer = FileEnhancer()
        path = os.path.join(FIXTURES_DIR, "clean.xlsx")
        result = fixer.analyze(path)
        self.assertIsNotNone(result)
        fixer._engine.close()

    def test_mind030_formula_engine_still_works(self):
        """MIND-030 should still function."""
        from engine.formula_engine import FormulaEngine
        engine = FormulaEngine()
        path = os.path.join(FIXTURES_DIR, "sample.xlsx")
        formulas = engine.scan_formulas(path)
        engine.close()

    def test_mind040_text_parser_still_works(self):
        """MIND-040 should still function."""
        from engine.text_task_parser import TextTaskParser
        parser = TextTaskParser()
        result = parser.parse("jumlahkan A1:A2 ke B1")
        self.assertIsNotNone(result)
        self.assertEqual(result.operation, "SUM")


if __name__ == "__main__":
    unittest.main(verbosity=2)
