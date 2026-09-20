"""
MIND-020 File Enhancer Tests

Tests the analysis contract and safe whitespace enhancement.
All MIND-010 tests must continue to pass independently.
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.excel_engine import CoreExcelEngine, ExcelError, ExcelErrorCode
from engine.file_enhancer import (
    FileEnhancer,
    FileEnhancerError,
    AnalysisResult,
    EnhancementResult,
    EnhancementIssue,
    DuplicateRowIssue,
    MissingValueIssue,
    TypeInconsistencyIssue,
    HeaderQualityIssue,
    FormulaPresence,
    ColumnTypeInfo,
)
import os

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures")


class TestFileEnhancerAnalysis(unittest.TestCase):
    """MIND-020 Analysis Contract Tests."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.engine = CoreExcelEngine()
        cls.fixer = FileEnhancer()
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_paths = {
            "clean": os.path.join(FIXTURES_DIR, "clean.xlsx"),
            "duplicates": os.path.join(FIXTURES_DIR, "duplicates.xlsx"),
            "missing_values": os.path.join(FIXTURES_DIR, "missing_values.xlsx"),
            "whitespace": os.path.join(FIXTURES_DIR, "whitespace.xlsx"),
            "duplicate_headers": os.path.join(FIXTURES_DIR, "duplicate_headers.xlsx"),
            "mixed_types": os.path.join(FIXTURES_DIR, "mixed_types.xlsx"),
            "formulas": os.path.join(FIXTURES_DIR, "formulas.xlsx"),
            "empty": os.path.join(FIXTURES_DIR, "empty.xlsx"),
        }

    @classmethod
    def tearDownClass(cls) -> None:
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _fixture_path(self, name: str) -> str:
        return self.fixture_paths[name]

    # ------------------------------------------------------------------
    # Clean workbook
    # ------------------------------------------------------------------
    def test_clean_workbook_no_issues(self):
        """Analyze a clean workbook should produce minimal issues."""
        result = self.fixer.analyze(self._fixture_path("clean"))
        self.assertIsInstance(result, AnalysisResult)
        self.assertEqual(len(result.duplicate_rows), 0)
        # Missing values may exist from None vs empty string distinction
        self.assertEqual(len(result.formula_presence), 0)

    # ------------------------------------------------------------------
    # Duplicate rows
    # ------------------------------------------------------------------
    def test_duplicate_rows_detected(self):
        """Analyze duplicate workbook should detect duplicate rows."""
        result = self.fixer.analyze(self._fixture_path("duplicates"))
        self.assertIsInstance(result, AnalysisResult)
        self.assertGreater(len(result.duplicate_rows), 0)

        # Verify structure
        dup = result.duplicate_rows[0]
        self.assertIsInstance(dup, DuplicateRowIssue)
        self.assertEqual(dup.sheet, "Sales")
        self.assertGreater(len(dup.row_indices), 1)  # At least 2 rows are duplicates

    def test_duplicate_rows_header_aware(self):
        """Duplicate detection should skip the header row."""
        result = self.fixer.analyze(self._fixture_path("duplicates"))
        # Row 2 (Widget A) appears at row 2 and row 4 (both data rows)
        # Row 3 (Widget B) appears at row 3 and row 5
        for dup in result.duplicate_rows:
            for idx in dup.row_indices:
                self.assertGreater(idx, 1)  # All duplicate rows should be data rows (> header)

    # ------------------------------------------------------------------
    # Missing values
    # ------------------------------------------------------------------
    def test_missing_values_detected(self):
        """Analyze missing-values workbook should detect missing cells."""
        result = self.fixer.analyze(self._fixture_path("missing_values"))
        self.assertIsInstance(result, AnalysisResult)
        self.assertGreater(len(result.missing_values), 0)

        # Verify structure
        mv = result.missing_values[0]
        self.assertIsInstance(mv, MissingValueIssue)
        self.assertEqual(mv.sheet, "Inventory")
        self.assertIsInstance(mv.missing_count, int)
        self.assertGreater(mv.missing_count, 0)

    # ------------------------------------------------------------------
    # Column type distribution
    # ------------------------------------------------------------------
    def test_type_inconsistency_detected(self):
        """Analyze mixed-types workbook should detect type inconsistency."""
        result = self.fixer.analyze(self._fixture_path("mixed_types"))
        self.assertIsInstance(result, AnalysisResult)
        self.assertGreater(len(result.type_inconsistencies), 0)

        ti = result.type_inconsistencies[0]
        self.assertIsInstance(ti, TypeInconsistencyIssue)
        self.assertGreater(len(ti.types_found), 1)

    def test_column_type_info(self):
        """Analysis should include column type information."""
        result = self.fixer.analyze(self._fixture_path("clean"))
        self.assertIsInstance(result, AnalysisResult)
        self.assertGreater(len(result.column_types), 0)

        ct = result.column_types[0]
        self.assertIsInstance(ct, ColumnTypeInfo)
        self.assertIsInstance(ct.column, str)
        self.assertIsInstance(ct.dominant_type, str)

    # ------------------------------------------------------------------
    # Header quality
    # ------------------------------------------------------------------
    def test_empty_header_detected(self):
        """Duplicate headers workbook with empty header should detect it."""
        result = self.fixer.analyze(self._fixture_path("duplicate_headers"))
        self.assertIsInstance(result, AnalysisResult)

        # The "BadHeaders" sheet has an empty first header
        empty_headers = [h for h in result.header_quality_issues if h.issue_type == "empty"]
        self.assertGreater(len(empty_headers), 0)

    def test_duplicate_header_detected(self):
        """Duplicate headers workbook should detect duplicate headers."""
        result = self.fixer.analyze(self.fixture_paths["duplicate_headers"])
        self.assertIsInstance(result, AnalysisResult)

        dup_headers = [h for h in result.header_quality_issues if h.issue_type == "duplicate"]
        self.assertGreater(len(dup_headers), 0)

    def test_surrounding_whitespace_header_detected(self):
        """Whitespace workbook should detect padded headers."""
        result = self.fixer.analyze(self.fixture_paths["whitespace"])
        self.assertIsInstance(result, AnalysisResult)

        ws_headers = [h for h in result.header_quality_issues if h.issue_type == "surrounding_whitespace"]
        self.assertGreater(len(ws_headers), 0)

    # ------------------------------------------------------------------
    # Formula detection
    # ------------------------------------------------------------------
    def test_formula_detection(self):
        """Formulas workbook should detect formula cells."""
        result = self.fixer.analyze(self.fixture_paths["formulas"])
        self.assertIsInstance(result, AnalysisResult)
        self.assertGreater(len(result.formula_presence), 0)

        fp = result.formula_presence[0]
        self.assertIsInstance(fp, FormulaPresence)
        self.assertEqual(fp.sheet, "Calc")
        self.assertGreater(fp.formula_count, 0)
        self.assertIsInstance(fp.locations, list)

    # ------------------------------------------------------------------
    # Invalid workbook
    # ------------------------------------------------------------------
    def test_invalid_file_raises_error(self):
        """Analyzing a non-existent file should raise FileEnhancerError."""
        with self.assertRaises(FileEnhancerError):
            self.fixer.analyze("/nonexistent/file.xlsx")

    # ------------------------------------------------------------------
    # Empty workbook
    # ------------------------------------------------------------------
    def test_empty_workbook_handled(self):
        """Analyze empty workbook should not crash."""
        result = self.fixer.analyze(self._fixture_path("empty"))
        self.assertIsInstance(result, AnalysisResult)
        self.assertEqual(result.sheet_count, 1)

    # ------------------------------------------------------------------
    # Duplicate detection on clean workbook
    # ------------------------------------------------------------------
    def test_no_duplicates_in_clean_workbook(self):
        """Clean workbook should have no duplicate rows."""
        result = self.fixer.analyze(self._fixture_path("clean"))
        self.assertEqual(len(result.duplicate_rows), 0)


class TestFileEnhancerEnhancement(unittest.TestCase):
    """MIND-020 Safe Enhancement Tests."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.fixer = FileEnhancer()
        cls.temp_dir = tempfile.mkdtemp()
        cls.fixture_paths = {
            "clean": os.path.join(FIXTURES_DIR, "clean.xlsx"),
            "duplicates": os.path.join(FIXTURES_DIR, "duplicates.xlsx"),
            "missing_values": os.path.join(FIXTURES_DIR, "missing_values.xlsx"),
            "whitespace": os.path.join(FIXTURES_DIR, "whitespace.xlsx"),
            "duplicate_headers": os.path.join(FIXTURES_DIR, "duplicate_headers.xlsx"),
            "mixed_types": os.path.join(FIXTURES_DIR, "mixed_types.xlsx"),
            "formulas": os.path.join(FIXTURES_DIR, "formulas.xlsx"),
            "empty": os.path.join(FIXTURES_DIR, "empty.xlsx"),
        }

    @classmethod
    def tearDownClass(cls) -> None:
        cls.fixer._engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _output_path(self, filename: str) -> str:
        return os.path.join(self.temp_dir, filename)

    # ------------------------------------------------------------------
    # Whitespace trimming
    # ------------------------------------------------------------------
    def test_whitespace_trimming_applied(self):
        """Enhance should trim whitespace from text cells."""
        input_path = self.fixture_paths["whitespace"]
        output_path = self._output_path("enhanced_whitespace.xlsx")

        result = self.fixer.enhance(input_path, output_path)
        self.assertIsInstance(result, EnhancementResult)
        self.assertGreater(result.cells_trimmed, 0)
        self.assertTrue(result.valid)

        # Verify the output file can be loaded and values are trimmed
        engine = CoreExcelEngine()
        engine.load(output_path)
        ws = engine.get_active_sheet()
        header_cell = engine.get_cell(ws.name, 1, 1)
        self.assertEqual(header_cell.value, "Name")
        engine.close()

    def test_formulas_not_modified(self):
        """Enhance should not modify formula cells."""
        input_path = self.fixture_paths["formulas"]
        output_path = self._output_path("enhanced_formulas.xlsx")

        result = self.fixer.enhance(input_path, output_path)
        self.assertIsInstance(result, EnhancementResult)

        engine = CoreExcelEngine()
        engine.load(output_path)
        total_cell = engine.get_cell("Calc", 2, 3)
        self.assertIsInstance(total_cell.value, str)
        self.assertTrue(total_cell.value.startswith("="))
        engine.close()

    def test_non_text_values_not_affected(self):
        """Enhance should not modify numeric or other non-text values."""
        input_path = self.fixture_paths["clean"]
        output_path = self._output_path("enhanced_clean.xlsx")

        result = self.fixer.enhance(input_path, output_path)
        self.assertIsInstance(result, EnhancementResult)
        self.assertEqual(result.cells_trimmed, 0)
        self.assertTrue(result.valid)

    # ------------------------------------------------------------------
    # Save + validate
    # ------------------------------------------------------------------
    def test_save_and_reload_valid(self):
        """Enhanced workbook should be valid after save."""
        input_path = self.fixture_paths["clean"]
        output_path = self._output_path("validated.xlsx")

        result = self.fixer.enhance(input_path, output_path)
        self.assertTrue(result.valid)

        engine = CoreExcelEngine()
        info = engine.load(output_path)
        self.assertEqual(info.sheet_count, 1)
        engine.close()

    def test_duplicate_rows_workbook_enhanced(self):
        """Enhance a duplicate-rows workbook — should save and validate."""
        input_path = self.fixture_paths["duplicates"]
        output_path = self._output_path("enhanced_duplicates.xlsx")

        result = self.fixer.enhance(input_path, output_path)
        self.assertTrue(result.valid)

        engine = CoreExcelEngine()
        engine.load(output_path)
        sheet_names = engine.get_sheet_names()
        self.assertIn("Sales", sheet_names)
        row_count = engine.get_row_count("Sales")
        self.assertEqual(row_count, 6)  # Header + 5 data rows
        engine.close()

    # ------------------------------------------------------------------
    # Analysis result structure
    # ------------------------------------------------------------------
    def test_analysis_result_has_all_categories(self):
        """Analysis result should have all issue categories."""
        result = self.fixer.analyze(self.fixture_paths["mixed_types"])
        self.assertIsInstance(result, AnalysisResult)
        self.assertGreater(result.sheet_count, 0)
        self.assertGreater(len(result.sheets_analyzed), 0)
        self.assertGreater(len(result.column_types), 0)

    # ------------------------------------------------------------------
    # Issue structure
    # ------------------------------------------------------------------
    def test_issue_schema(self):
        """EnhancementIssue should have required fields."""
        result = self.fixer.analyze(self.fixture_paths["whitespace"])

        if result.header_quality_issues:
            issue = result.header_quality_issues[0]
            self.assertIsInstance(issue, HeaderQualityIssue)
            self.assertIsInstance(issue.issue_type, str)
            self.assertIsInstance(issue.sheet, str)
            self.assertIsInstance(issue.location, str)
            self.assertIsInstance(issue.description, str)

    # ------------------------------------------------------------------
    # Invalid file during enhance
    # ------------------------------------------------------------------
    def test_invalid_file_enhance_raises_error(self):
        """Enhancing a non-existent file should raise FileEnhancerError."""
        output_path = self._output_path("error.xlsx")
        with self.assertRaises(FileEnhancerError):
            self.fixer.enhance("/nonexistent/file.xlsx", output_path)


class TestFileEnhancerRegression(unittest.TestCase):
    """Regression tests ensuring MIND-010 compatibility."""

    def test_mind010_engine_still_works(self):
        """CoreExcelEngine should still work after MIND-020 changes."""
        engine = CoreExcelEngine()
        fixture = os.path.join(FIXTURES_DIR, "sample.xlsx")
        info = engine.load(fixture)
        self.assertEqual(info.sheet_count, 2)
        names = engine.get_sheet_names()
        self.assertIn("Sales", names)
        self.assertIn("Inventory", names)
        engine.close()

    def test_mind010_pipeline_still_works(self):
        """MIND-010 pipeline should still work."""
        import tempfile
        engine = CoreExcelEngine()
        fixture = os.path.join(FIXTURES_DIR, "sample.xlsx")
        output = os.path.join(tempfile.mkdtemp(), "pipeline_test.xlsx")
        result = engine.pipeline(
            fixture,
            output,
            modifications={"rename_sheet": {"old_name": "Inventory", "new_name": "Stock"}},
        )
        self.assertTrue(result["valid"])
        engine.close()

    def test_mind010_validation_still_works(self):
        """MIND-010 validation should still work."""
        engine = CoreExcelEngine()
        fixture = os.path.join(FIXTURES_DIR, "sample.xlsx")
        output = os.path.join(tempfile.mkdtemp(), "validate_test.xlsx")
        engine.load(fixture)
        engine.save(output)
        validation = engine.validate(output)
        self.assertTrue(validation["valid"])
        engine.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
