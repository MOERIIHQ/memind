"""
MIND-040 Text Task Parser Tests

Test categories:
- Valid SUM commands
- Valid MULTIPLY commands
- Invalid commands
- Unsupported commands
- Ambiguous commands
- Task contract validation
- Determinism
- MIND-030 integration
- Real XLSX end-to-end
- Regression (MIND-010/020/030 still work)
"""

import os
import re
import sys
import tempfile
import unittest
from copy import copy

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.text_task_parser import (
    TextTaskParser,
    TextTask,
    TaskParseError,
    parse,
    execute_task,
    validate_cell_reference,
    validate_range,
    extract_cell_ref,
    extract_range,
)
from engine.formula_engine import (
    FormulaEngine,
    FormulaOperation,
    FormulaEngineError,
)
from engine.excel_engine import CoreExcelEngine, ExcelError

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures")


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_clean_workbook(path: str) -> str:
    """Create a clean workbook for testing."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["A", "B", "C", "D", "E", "F"])
    ws.append([10, 20, None, None, None, None])
    ws.append([5, 10, None, None, None, None])
    wb.save(path)
    return path


def _make_sum_e2e_fixture(path: str) -> str:
    """Create a workbook with A2=10, B2=20 for SUM E2E test.

    Sheet name is "Test" to match verification in test.
    """
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Test"
    ws.append(["A", "B", "C", "D", "E", "F"])
    ws.append([10, 20, None, None, None, None])
    wb.save(path)
    return path


def _make_multiply_e2e_fixture(path: str) -> str:
    """Create a workbook with D2=5, E2=10 for MULTIPLY E2E test."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Test"
    # Row 1: headers
    ws.append(["A", "B", "C", "D", "E", "F"])
    # Row 2: A2=10, B2=20, C2=None, D2=5, E2=10, F2=None
    ws.append([10, 20, None, 5, 10, None])
    wb.save(path)
    return path


def _make_workbook_for_preservation_test(path: str) -> str:
    """Create a workbook specifically for non-target preservation testing.

    Layout:
    - Row 1: headers A, B, C, D, E, F
    - Row 2: A2=10, B2=20, C2=None, D2=None, E2=None, F2=None  (SUM target: C2)
    - Row 3: A3=None, B3=None, C3=None, D3=5, E3=None, F3=None  (verify D3=5 preserved)
    """
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Test"
    # Row 1: headers
    ws.append(["A", "B", "C", "D", "E", "F"])
    # Row 2: values for SUM test (A2=10, B2=20)
    ws.append([10, 20, None, None, None, None])
    # Row 3: D3=5 for preservation verification
    ws.append([None, None, None, 5, None, None])
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Valid SUM tests
# ---------------------------------------------------------------------------

class TestValidSumCommands(unittest.TestCase):
    """Test valid SUM command parsing."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_sum_range_form(self):
        """"jumlahkan B2:C2 ke D2" must produce SUM with range."""
        result = self.parser.parse("jumlahkan B2:C2 ke D2")
        self.assertIsInstance(result, TextTask)
        self.assertEqual(result.operation, "SUM")
        self.assertEqual(result.target, "D2")
        self.assertEqual(result.operands, {"range": "B2:C2"})
        self.assertEqual(result.task_type, "FORMULA")
        self.assertEqual(result.overwrite_policy, "never")

    def test_sum_start_sampai_end_form(self):
        """"jumlahkan B2 sampai C2 ke D2" must produce SUM with range."""
        result = self.parser.parse("jumlahkan B2 sampai C2 ke D2")
        self.assertIsInstance(result, TextTask)
        self.assertEqual(result.operation, "SUM")
        self.assertEqual(result.target, "D2")
        self.assertEqual(result.operands, {"range": "B2:C2"})

    def test_sum_with_dollar_signs(self):
        """Range with $ notation must work."""
        result = self.parser.parse("jumlahkan $B$2:$C$2 ke $D$2")
        self.assertIsInstance(result, TextTask)
        self.assertEqual(result.operation, "SUM")
        self.assertEqual(result.target, "D2")

    def test_sum_range_validation(self):
        """Extracted range must be valid."""
        result = self.parser.parse("jumlahkan A1:A10 ke B1")
        self.assertIsInstance(result, TextTask)
        self.assertTrue(validate_range(result.operands["range"]))


# ---------------------------------------------------------------------------
# Valid MULTIPLY tests
# ---------------------------------------------------------------------------

class TestValidMultiplyCommands(unittest.TestCase):
    """Test valid MULTIPLY command parsing."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_multiply_binary_operands(self):
        """"kalikan D2 dengan E2 ke F2" must produce MULTIPLY."""
        result = self.parser.parse("kalikan D2 dengan E2 ke F2")
        self.assertIsInstance(result, TextTask)
        self.assertEqual(result.operation, "MULTIPLY")
        self.assertEqual(result.target, "F2")
        self.assertEqual(result.operands, {"left": "D2", "right": "E2"})

    def test_multiply_with_dollar_signs(self):
        """Cell refs with $ notation must work."""
        result = self.parser.parse("kalikan $D$2 dengan $E$2 ke $F$2")
        self.assertIsInstance(result, TextTask)
        self.assertEqual(result.operation, "MULTIPLY")

    def test_multiply_reference_validation(self):
        """Extracted refs must be valid."""
        result = self.parser.parse("kalikan A1 dengan B1 ke C1")
        self.assertIsInstance(result, TextTask)
        self.assertTrue(validate_cell_reference(result.operands["left"]))
        self.assertTrue(validate_cell_reference(result.operands["right"]))


# ---------------------------------------------------------------------------
# Invalid command tests
# ---------------------------------------------------------------------------

class TestInvalidCommands(unittest.TestCase):
    """Test that invalid commands are rejected with TaskParseError."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_empty_text(self):
        """Empty text must return INVALID."""
        result = self.parser.parse("")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_whitespace_only(self):
        """Whitespace-only text must return INVALID."""
        result = self.parser.parse("   ")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_missing_target(self):
        """SUM without target after 'ke' must be INVALID."""
        result = self.parser.parse("jumlahkan A1:A2")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_malformed_cell_reference(self):
        """Malformed cell ref must be INVALID."""
        result = self.parser.parse("jumlahkan XYZ ke A1")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_incomplete_sum_without_range(self):
        """SUM without range or cell refs must be INVALID."""
        result = self.parser.parse("jumlahkan ke A1")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_multiply_missing_separator(self):
        """MULTIPLY without 'dengan' must be INVALID."""
        result = self.parser.parse("kalikan D2 E2 ke F2")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_multiply_missing_left_operand(self):
        """MULTIPLY without left operand must be INVALID."""
        result = self.parser.parse("kalikan dengan E2 ke F2")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_multiply_missing_right_operand(self):
        """MULTIPLY without right operand must be INVALID."""
        result = self.parser.parse("kalikan D2 dengan ke F2")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")


# ---------------------------------------------------------------------------
# Unsupported command tests
# ---------------------------------------------------------------------------

class TestUnsupportedCommands(unittest.TestCase):
    """Test that unsupported commands are rejected with UNSUPPORTED kind."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_unsupported_operation(self):
        """"tambahkan A1 dan B1" is outside MVP grammar."""
        result = self.parser.parse("tambahkan A1 dan B1")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "UNSUPPORTED")

    def test_unsupported_row_operation(self):
        """"hapus baris 3" is outside MVP grammar."""
        result = self.parser.parse("hapus baris 3")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "UNSUPPORTED")

    def test_unsupported_pivot(self):
        """"buat pivot" is outside MVP grammar."""
        result = self.parser.parse("buat pivot")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "UNSUPPORTED")

    def test_unsupported_chart(self):
        """"buat grafik" is outside MVP grammar."""
        result = self.parser.parse("buat grafik")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "UNSUPPORTED")

    def test_unsupported_natural_language_request(self):
        """"Hitung total penjualan" is outside MVP grammar."""
        result = self.parser.parse("Hitung total penjualan")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "UNSUPPORTED")

    def test_unsupported_add(self):
        """"tambahkan A1 ke B1" is not in MVP grammar (ADD not supported)."""
        result = self.parser.parse("tambahkan A1 ke B1")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "UNSUPPORTED")


# ---------------------------------------------------------------------------
# Ambiguous command tests
# ---------------------------------------------------------------------------

class TestAmbiguousCommands(unittest.TestCase):
    """Test that ambiguous commands are rejected with AMBIGUOUS kind."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_ambiguous_missing_source(self):
        """"jumlahkan ke C2" without source is ambiguous or invalid."""
        result = self.parser.parse("jumlahkan ke C2")
        # This should be INVALID (no range/source) or AMBIGUOUS
        self.assertIsInstance(result, TaskParseError)
        self.assertIn(result.kind, ["INVALID", "AMBIGUOUS"])

    def test_ambiguous_missing_target(self):
        """"jumlahkan A1:A2" without target is invalid."""
        result = self.parser.parse("jumlahkan A1:A2")
        self.assertIsInstance(result, TaskParseError)
        self.assertEqual(result.kind, "INVALID")

    def test_insufficient_information(self):
        """"Hitung total" without specifics is UNSUPPORTED or AMBIGUOUS."""
        result = self.parser.parse("Hitung total")
        # No spreadsheet keywords → INVALID
        self.assertIsInstance(result, TaskParseError)


# ---------------------------------------------------------------------------
# Task contract tests
# ---------------------------------------------------------------------------

class TestTaskContract(unittest.TestCase):
    """Test TextTask model and contract."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_required_fields_present(self):
        """Parsed task must have all required fields."""
        result = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertIsInstance(result, TextTask)
        self.assertIsNotNone(result.task_type)
        self.assertIsNotNone(result.operation)
        self.assertIsNotNone(result.target)
        self.assertIsNotNone(result.operands)

    def test_operation_restriction(self):
        """Operation must be SUM or MULTIPLY."""
        result = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertIn(result.operation, ["SUM", "MULTIPLY"])

    def test_operands_structure_sum(self):
        """SUM operands must have 'range' key."""
        result = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertIn("range", result.operands)
        self.assertNotIn("left", result.operands)
        self.assertNotIn("right", result.operands)

    def test_operands_structure_multiply(self):
        """MULTIPLY operands must have 'left' and 'right' keys."""
        result = self.parser.parse("kalikan A1 dengan B1 ke C1")
        self.assertIn("left", result.operands)
        self.assertIn("right", result.operands)
        self.assertNotIn("range", result.operands)

    def test_optional_sheet_default(self):
        """Sheet should default to None."""
        result = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertIsNone(result.sheet)

    def test_default_overwrite_policy(self):
        """Default overwrite_policy must be 'never'."""
        result = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertEqual(result.overwrite_policy, "never")

    def test_task_type_is_formula(self):
        """task_type must be 'FORMULA'."""
        result = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertEqual(result.task_type, "FORMULA")


# ---------------------------------------------------------------------------
# Determinism tests
# ---------------------------------------------------------------------------

class TestDeterminism(unittest.TestCase):
    """Verify parser is deterministic."""

    def setUp(self):
        self.parser = TextTaskParser()

    def test_sum_deterministic(self):
        """Same SUM text parsed 5x must produce identical results."""
        text = "jumlahkan B2:C2 ke D2"
        results = [self.parser.parse(text) for _ in range(5)]
        for i in range(1, len(results)):
            self.assertEqual(results[0].operation, results[i].operation)
            self.assertEqual(results[0].target, results[i].target)
            self.assertEqual(results[0].operands, results[i].operands)
            self.assertEqual(results[0].task_type, results[i].task_type)
            self.assertEqual(results[0].sheet, results[i].sheet)
            self.assertEqual(results[0].overwrite_policy, results[i].overwrite_policy)

    def test_multiply_deterministic(self):
        """Same MULTIPLY text parsed 5x must produce identical results."""
        text = "kalikan D2 dengan E2 ke F2"
        results = [self.parser.parse(text) for _ in range(5)]
        for i in range(1, len(results)):
            self.assertEqual(results[0].operation, results[i].operation)
            self.assertEqual(results[0].target, results[i].target)
            self.assertEqual(results[0].operands, results[i].operands)

    def test_error_deterministic(self):
        """Same invalid text parsed 5x must produce identical errors."""
        text = "invalid command"
        results = [self.parser.parse(text) for _ in range(5)]
        for i in range(1, len(results)):
            self.assertEqual(results[0].kind, results[i].kind)
            self.assertEqual(results[0].message, results[i].message)
            self.assertEqual(results[0].details, results[i].details)

    def test_no_state_leakage(self):
        """Repeated parsing must not leak state between calls."""
        text1 = "jumlahkan A1:A2 ke B1"
        text2 = "kalikan C1 dengan D1 ke E1"

        r1_first = self.parser.parse(text1)
        r2 = self.parser.parse(text2)
        r1_second = self.parser.parse(text1)

        # First and second parse of text1 should be identical
        self.assertEqual(r1_first.operation, r1_second.operation)
        self.assertEqual(r1_first.target, r1_second.target)
        self.assertEqual(r1_first.operands, r1_second.operands)


# ---------------------------------------------------------------------------
# MIND-030 integration tests
# ---------------------------------------------------------------------------

class TestMIND030Integration(unittest.TestCase):
    """Verify TextTask → FormulaOperation → MIND-030 pipeline."""

    def setUp(self):
        self.parser = TextTaskParser()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _output(self, name: str) -> str:
        return os.path.join(self.temp_dir, name)

    def test_sum_task_to_formula_operation(self):
        """TextTask for SUM must convert to correct FormulaOperation."""
        task = self.parser.parse("jumlahkan A1:A2 ke B1")
        self.assertIsInstance(task, TextTask)

        operation = FormulaOperation(
            target=task.target,
            operation=task.operation,
            operands=task.operands,
            overwrite_policy=task.overwrite_policy,
            sheet=task.sheet,
        )

        self.assertEqual(operation.target, "B1")
        self.assertEqual(operation.operation, "SUM")
        self.assertEqual(operation.operands, {"range": "A1:A2"})

    def test_multiply_task_to_formula_operation(self):
        """TextTask for MULTIPLY must convert to correct FormulaOperation."""
        task = self.parser.parse("kalikan A1 dengan B1 ke C1")
        self.assertIsInstance(task, TextTask)

        operation = FormulaOperation(
            target=task.target,
            operation=task.operation,
            operands=task.operands,
            overwrite_policy=task.overwrite_policy,
            sheet=task.sheet,
        )

        self.assertEqual(operation.target, "C1")
        self.assertEqual(operation.operation, "MULTIPLY")
        self.assertEqual(operation.operands, {"left": "A1", "right": "B1"})

    def test_execute_sum_via_mind030(self):
        """execute() must delegate to FormulaEngine.apply_formula()."""
        input_path = os.path.join(self.temp_dir, "input.xlsx")
        output_path = self._output("output_sum.xlsx")
        _make_clean_workbook(input_path)

        task = self.parser.parse("jumlahkan A2:B2 ke C2")
        self.assertIsInstance(task, TextTask)

        result = self.parser.execute(task, input_path, output_path)

        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=SUM(A2:B2)")
        self.assertEqual(result["target_cell"], "C2")

    def test_execute_multiply_via_mind030(self):
        """execute() must delegate to FormulaEngine.apply_formula()."""
        input_path = os.path.join(self.temp_dir, "input.xlsx")
        output_path = self._output("output_mult.xlsx")
        _make_clean_workbook(input_path)

        task = self.parser.parse("kalikan A3 dengan B3 ke C3")
        self.assertIsInstance(task, TextTask)

        result = self.parser.execute(task, input_path, output_path)

        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=A3*B3")
        self.assertEqual(result["target_cell"], "C3")


# ---------------------------------------------------------------------------
# Real XLSX end-to-end tests
# ---------------------------------------------------------------------------

class TestRealXLSXEndToEnd(unittest.TestCase):
    """End-to-end tests with real XLSX files."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.parser = TextTaskParser()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _output(self, name: str) -> str:
        return os.path.join(self.temp_dir, name)

    def test_sum_e2e_persistence(self):
        """SUM formula must persist after save/reopen cycle.

        Input: A2=10, B2=20, C2=blank
        Text: "jumlahkan A2 sampai B2 ke C2"
        Expected: C2 = "=SUM(A2:B2)"
        """
        input_path = os.path.join(self.temp_dir, "input_sum.xlsx")
        output_path = self._output("output_sum.xlsx")
        _make_clean_workbook(input_path)
        # Rename sheet to "Test" to match verification
        from openpyxl import load_workbook
        wb = load_workbook(input_path)
        wb.active.title = "Test"
        wb.save(input_path)

        task = self.parser.parse("jumlahkan A2 sampai B2 ke C2")
        self.assertIsInstance(task, TextTask)

        result = self.parser.execute(task, input_path, output_path)

        # Verify result
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=SUM(A2:B2)")
        self.assertEqual(result["target_cell"], "C2")

        # Reopen and verify independently
        verify = CoreExcelEngine()
        verify.load(output_path)
        cell = verify.get_cell("Test", 2, 3)  # C2
        self.assertEqual(cell.formula, "=SUM(A2:B2)")
        self.assertEqual(cell.value, "=SUM(A2:B2)")
        verify.close()

    def test_multiply_e2e_persistence(self):
        """MULTIPLY formula must persist after save/reopen cycle.

        Input: D2=5, E2=10, F2=blank
        Text: "kalikan D2 dengan E2 ke F2"
        Expected: F2 = "=D2*E2"
        """
        input_path = os.path.join(self.temp_dir, "input_mult.xlsx")
        output_path = self._output("output_mult.xlsx")
        _make_multiply_e2e_fixture(input_path)

        task = self.parser.parse("kalikan D2 dengan E2 ke F2")
        self.assertIsInstance(task, TextTask)

        result = self.parser.execute(task, input_path, output_path)

        # Verify result
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=D2*E2")
        self.assertEqual(result["target_cell"], "F2")

        # Reopen and verify independently
        verify = CoreExcelEngine()
        verify.load(output_path)
        cell = verify.get_cell("Test", 2, 6)  # F2 (row 2, col 6)
        self.assertEqual(cell.formula, "=D2*E2")
        self.assertEqual(cell.value, "=D2*E2")
        verify.close()

    def test_non_target_cells_preserved(self):
        """Non-target cells must remain intact after formula application."""
        input_path = os.path.join(self.temp_dir, "input_preserve.xlsx")
        output_path = self._output("output_preserve.xlsx")
        _make_workbook_for_preservation_test(input_path)

        # Apply SUM to C2 only
        task = self.parser.parse("jumlahkan A2 sampai B2 ke C2")
        result = self.parser.execute(task, input_path, output_path)

        self.assertTrue(result["valid"])

        # Verify non-target cells
        verify = CoreExcelEngine()
        verify.load(output_path)

        # A2 should still be 10
        a2 = verify.get_cell("Test", 2, 1)
        self.assertEqual(a2.value, 10)

        # B2 should still be 20
        b2 = verify.get_cell("Test", 2, 2)
        self.assertEqual(b2.value, 20)

        # D3 should still be 5 (row 3, col 4 = D3)
        d3 = verify.get_cell("Test", 3, 4)
        self.assertEqual(d3.value, 5)

        verify.close()


# ---------------------------------------------------------------------------
# Regression tests
# ---------------------------------------------------------------------------

class TestRegression(unittest.TestCase):
    """Verify existing engines still work after MIND-040 addition."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_mind010_engine_still_works(self):
        """CoreExcelEngine must remain functional."""
        engine = CoreExcelEngine()
        fixture = os.path.join(FIXTURES_DIR, "sample.xlsx")
        if os.path.exists(fixture):
            info = engine.load(fixture)
            self.assertGreater(info.sheet_count, 0)
            engine.close()

    def test_mind020_enhancer_still_works(self):
        """FileEnhancer must remain functional."""
        from engine.file_enhancer import FileEnhancer
        fixer = FileEnhancer()
        fixture = os.path.join(FIXTURES_DIR, "clean.xlsx")
        if os.path.exists(fixture):
            result = fixer.analyze(fixture)
            self.assertIsNotNone(result)
            fixer._engine.close()

    def test_mind030_formula_engine_still_works(self):
        """FormulaEngine must remain functional."""
        engine = FormulaEngine()
        fixture = os.path.join(FIXTURES_DIR, "sample.xlsx")
        if os.path.exists(fixture):
            formulas = engine.scan_formulas(fixture)
            # May be empty, but must not error
            engine.close()

    def test_convenience_functions_work(self):
        """Top-level parse() and execute_task() must work."""
        task = parse("jumlahkan A1:A2 ke B1")
        self.assertIsInstance(task, TextTask)

        # Can't execute without a file, but parsing must work
        self.assertEqual(task.operation, "SUM")


# ---------------------------------------------------------------------------
# Cell reference utility tests
# ---------------------------------------------------------------------------

class TestCellReferenceUtilities(unittest.TestCase):
    """Test cell reference validation utilities."""

    def test_valid_single_cell(self):
        self.assertTrue(validate_cell_reference("A1"))
        self.assertTrue(validate_cell_reference("B2"))
        self.assertTrue(validate_cell_reference("AA10"))

    def test_valid_absolute_cell(self):
        self.assertTrue(validate_cell_reference("$A$1"))
        self.assertTrue(validate_cell_reference("$A1"))
        self.assertTrue(validate_cell_reference("A$1"))

    def test_valid_range(self):
        self.assertTrue(validate_cell_reference("A1:B2"))
        self.assertTrue(validate_cell_reference("A1:A10"))
        self.assertTrue(validate_cell_reference("Sheet1!A1"))

    def test_invalid_cell_refs(self):
        self.assertFalse(validate_cell_reference(""))
        self.assertFalse(validate_cell_reference("1A"))
        self.assertFalse(validate_cell_reference("A"))
        self.assertFalse(validate_cell_reference("A:1"))
        self.assertFalse(validate_cell_reference("ABC"))  # No row number

    def test_extract_cell_ref(self):
        self.assertEqual(extract_cell_ref("jumlahkan A1 ke B1"), "A1")
        self.assertEqual(extract_cell_ref("kalikan D2 dengan E2 ke F2"), "D2")

    def test_extract_range(self):
        self.assertEqual(extract_range("jumlahkan A1:B2 ke C1"), "A1:B2")
        self.assertIsNone(extract_range("jumlahkan A1 ke B1"))


# ---------------------------------------------------------------------------
# Parse error model tests
# ---------------------------------------------------------------------------

class TestTaskParseErrorModel(unittest.TestCase):
    """Test TaskParseError model."""

    def test_ambiguous_kind(self):
        error = TaskParseError(
            kind="AMBIGUOUS",
            message="Cannot resolve target",
            details={"text": "jumlahkan ke A1"},
        )
        self.assertEqual(error.kind, "AMBIGUOUS")
        self.assertIn("Cannot resolve target", error.message)

    def test_invalid_kind(self):
        error = TaskParseError(
            kind="INVALID",
            message="Missing target",
            details={"reason": "missing_target"},
        )
        self.assertEqual(error.kind, "INVALID")

    def test_unsupported_kind(self):
        error = TaskParseError(
            kind="UNSUPPORTED",
            message="Command outside grammar",
            details={"text": "buat pivot"},
        )
        self.assertEqual(error.kind, "UNSUPPORTED")


# ---------------------------------------------------------------------------
# Convenience function tests
# ---------------------------------------------------------------------------

class TestConvenienceFunctions(unittest.TestCase):
    """Test top-level parse() and execute_task() functions."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_parse_function(self):
        """parse() function must work like TextTaskParser.parse()."""
        result = parse("jumlahkan A1:A2 ke B1")
        self.assertIsInstance(result, TextTask)
        self.assertEqual(result.operation, "SUM")

    def test_execute_task_function(self):
        """execute_task() function must work."""
        input_path = os.path.join(self.temp_dir, "input.xlsx")
        output_path = os.path.join(self.temp_dir, "output.xlsx")
        _make_clean_workbook(input_path)

        task = parse("jumlahkan A2:B2 ke C2")
        result = execute_task(task, input_path, output_path)

        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
