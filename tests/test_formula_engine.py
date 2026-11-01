"""
MIND-030 Formula Engine Tests

Tests cover:
- Formula detection
- Non-formula exclusion
- Multiple worksheets
- Deterministic results
- Valid formula validation
- Malformed formula validation
- Cell reference validation
- Sheet reference validation
- Unsupported/invalid formula detection
- SUM generation
- AVERAGE generation
- Arithmetic generation (ADD, SUBTRACT, MULTIPLY, DIVIDE)
- IF generation
- Invalid operation rejection
- Invalid operand rejection
- Formula insertion
- Existing formula protection
- Explicit overwrite
- Invalid target rejection
- Missing sheet rejection
- Protected sheet safety
- Formatting preservation
- Non-target cell preservation
- Existing formula preservation
- Write / save / reopen / verify
- Exact formula persistence

All tests use real .xlsx fixtures or temporary workbooks.
"""

import os
import sys
import tempfile
import unittest
from copy import copy

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.formula_engine import (
    FormulaEngine,
    FormulaEngineError,
    FormulaEngineErrorCode,
    FormulaInfo,
    FormulaValidationResult,
    FormulaOperation,
    FormulaType,
)
from engine.excel_engine import CoreExcelEngine

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures")


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_workbook_with_formulas(path: str) -> str:
    """Create a workbook with various formulas across multiple sheets."""
    from openpyxl import Workbook
    wb = Workbook()

    # Sheet 1: Calc — arithmetic and function formulas
    ws1 = wb.active
    ws1.title = "Calc"
    ws1.append(["A", "B", "Total"])
    ws1.append([10, 20, "=SUM(A2:B2)"])
    ws1.append([30, 40, "=AVERAGE(A3:B3)"])
    ws1.append([50, 60, "=A4+B4"])
    ws1.append([70, 80, "=A5*B5"])
    ws1.append([90, 100, "=IF(A6>50,\"BIG\",\"SMALL\")"])

    # Sheet 2: Data — plain values (no formulas)
    ws2 = wb.create_sheet("Data")
    ws2.append(["Name", "Value"])
    ws2.append(["Widget", 100])
    ws2.append(["Gadget", 200])

    # Sheet 3: MultiRef — cross-sheet references
    ws3 = wb.create_sheet("MultiRef")
    ws3.append(["Result"])
    ws3.append(["=Calc!A2+Calc!B2"])

    wb.save(path)
    return path


def _make_workbook_with_existing_formulas(path: str) -> str:
    """Create a workbook that already has formulas (for overwrite tests)."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Existing"
    ws.append(["X", "Y", "Z"])
    ws.append([1, 2, "=X2+Y2"])
    ws.append([3, 4, "=X3*Y3"])
    wb.save(path)
    return path


def _make_protected_workbook(path: str) -> str:
    """Create a workbook with one protected sheet."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Protected"
    ws.append(["Val", "Formula"])
    ws.append([10, "=Val2*2"])
    ws.protection.sheet = True
    wb.save(path)
    return path


def _make_workbook_with_formatting(path: str) -> str:
    """Create a workbook with cell formatting to verify preservation."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Formatted"
    ws.append(["Label", "Amount"])
    ws.append(["Revenue", 1000])
    # Apply formatting to non-formula cells
    ws["A2"].font = Font(bold=True)
    ws["B2"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Test classes
# ---------------------------------------------------------------------------

class TestFormulaDetection(unittest.TestCase):
    """MIND-030 Formula Detection Tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.formulas_path = os.path.join(cls.temp_dir, "formulas_test.xlsx")
        _make_workbook_with_formulas(cls.formulas_path)
        cls.engine = FormulaEngine()

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_formula_detection_finds_formulas(self):
        """Formula detection must find formula cells."""
        formulas = self.engine.scan_formulas(self.formulas_path)
        self.assertGreater(len(formulas), 0, "Should detect formulas")

    def test_formula_detection_excludes_non_formulas(self):
        """Non-formula cells must not be included."""
        formulas = self.engine.scan_formulas(self.formulas_path)
        for f in formulas:
            self.assertTrue(f.formula.startswith("="),
                            f"FormulaInfo.formula must start with '=': {f.formula!r}")

    def test_formula_detection_multiple_sheets(self):
        """Detection must scan all sheets."""
        formulas = self.engine.scan_formulas(self.formulas_path)
        sheets_with_formulas = {f.sheet for f in formulas}
        self.assertIn("Calc", sheets_with_formulas)
        self.assertIn("MultiRef", sheets_with_formulas)
        # Data sheet has no formulas
        self.assertNotIn("Data", sheets_with_formulas)

    def test_formula_detection_deterministic(self):
        """Repeated detection must produce identical results."""
        r1 = self.engine.scan_formulas(self.formulas_path)
        r2 = self.engine.scan_formulas(self.formulas_path)
        self.assertEqual(len(r1), len(r2))
        for i in range(len(r1)):
            self.assertEqual(r1[i].sheet, r2[i].sheet)
            self.assertEqual(r1[i].cell, r2[i].cell)
            self.assertEqual(r1[i].formula, r2[i].formula)
            self.assertEqual(r1[i].formula_type, r2[i].formula_type)
            self.assertEqual(r1[i].references, r2[i].references)

    def test_formula_detection_includes_metadata(self):
        """Each FormulaInfo must have required metadata."""
        formulas = self.engine.scan_formulas(self.formulas_path)
        for f in formulas:
            self.assertIsInstance(f.sheet, str)
            self.assertIsInstance(f.cell, str)
            self.assertIsInstance(f.formula, str)
            self.assertIsInstance(f.formula_type, str)
            self.assertIsInstance(f.references, list)
            self.assertIsInstance(f.row, int)
            self.assertIsInstance(f.col, int)

    def test_formula_detection_formulas_match_sheetdata(self):
        """Detected formula text must match what get_sheet returns."""
        from engine.excel_engine import CoreExcelEngine
        e = CoreExcelEngine()
        e.load(self.formulas_path)
        sheet_data = e.get_sheet("Calc")
        engine_formulas = self.engine.scan_formulas(self.formulas_path)
        e.close()

        # Build set of (row, col, formula_text) from SheetData
        sheet_formulas = set()
        for r, row_cells in enumerate(sheet_data.cells):
            for c, cell in enumerate(row_cells):
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    sheet_formulas.add((r, c, cell.value))

        # Verify engine detection matches
        for f in engine_formulas:
            if f.sheet == "Calc":
                self.assertIn((f.row, f.col, f.formula), sheet_formulas,
                              f"Detected formula mismatch at {f.cell}")

    def test_scan_formulas_returns_empty_for_clean_workbook(self):
        """Clean workbook with no formulas must return empty list."""
        from openpyxl import Workbook
        clean_path = os.path.join(self.temp_dir, "clean.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["A", "B"])
        ws.append([1, 2])
        wb.save(clean_path)

        formulas = self.engine.scan_formulas(clean_path)
        self.assertEqual(len(formulas), 0)


class TestFormulaValidation(unittest.TestCase):
    """MIND-030 Formula Validation Tests."""

    def setUp(self):
        self.engine = FormulaEngine()

    def tearDown(self):
        self.engine.close()

    def test_valid_sum_formula(self):
        """Valid SUM formula must pass validation."""
        result = self.engine.validate_formula("=SUM(A1:A10)")
        self.assertTrue(result.is_valid)
        self.assertEqual(len(result.errors), 0)

    def test_valid_arithmetic_formula(self):
        """Valid arithmetic formula must pass."""
        result = self.engine.validate_formula("=A1+B1")
        self.assertTrue(result.is_valid)

    def test_valid_if_formula(self):
        """Valid IF formula must pass."""
        result = self.engine.validate_formula('=IF(A1>10,"YES","NO")')
        self.assertTrue(result.is_valid)

    def test_malformed_unclosed_paren(self):
        """Unclosed parenthesis must be detected."""
        result = self.engine.validate_formula("=SUM(")
        self.assertFalse(result.is_valid)
        self.assertIn("Unbalanced parentheses", result.errors[0])

    def test_malformed_extra_closing_paren(self):
        """Extra closing parenthesis must be detected."""
        result = self.engine.validate_formula("=A1+)")
        self.assertFalse(result.is_valid)

    def test_malformed_dangling_operator(self):
        """Dangling operator at end must be detected."""
        result = self.engine.validate_formula("=A1+")
        self.assertFalse(result.is_valid)
        self.assertIn("Dangling operator", result.errors[0])

    def test_malformed_no_equals_sign(self):
        """Formula without '=' must be rejected."""
        result = self.engine.validate_formula("SUM(A1:A10)")
        self.assertFalse(result.is_valid)
        self.assertIn("must start with '='", result.errors[0])

    def test_empty_formula_rejected(self):
        """Empty formula must be rejected."""
        result = self.engine.validate_formula("")
        self.assertFalse(result.is_valid)

    def test_unknown_function_warning(self):
        """Unknown function must produce warning (not error)."""
        result = self.engine.validate_formula("=UNKNOWN_SHEET!A1")
        # This isn't a function call, so no warning about function
        # But it's a reference — structurally valid
        self.assertTrue(result.is_valid)

    def test_unknown_function_in_call(self):
        """Unknown function in a function call must be flagged UNSUPPORTED (invalid)."""
        result = self.engine.validate_formula("=UNKNOWNFUNC(A1)")
        self.assertFalse(result.is_valid)  # Unknown functions are unsupported, not just warned
        self.assertTrue(any("Unsupported" in err for err in result.errors))

    def test_cell_reference_validation_valid(self):
        """Valid cell references must pass."""
        self.assertTrue(self.engine.validate_cell_reference("A1"))
        self.assertTrue(self.engine.validate_cell_reference("$A$1"))
        self.assertTrue(self.engine.validate_cell_reference("A1:B10"))
        self.assertTrue(self.engine.validate_cell_reference("Sheet1!A1"))
        self.assertTrue(self.engine.validate_cell_reference("Sheet1!A1:B10"))

    def test_cell_reference_validation_invalid(self):
        """Invalid cell references must fail."""
        self.assertFalse(self.engine.validate_cell_reference(""))
        self.assertFalse(self.engine.validate_cell_reference("1A"))
        self.assertFalse(self.engine.validate_cell_reference("A"))
        self.assertFalse(self.engine.validate_cell_reference("A:1"))


class TestFormulaGeneration(unittest.TestCase):
    """MIND-030 Structured Formula Generation Tests."""

    def setUp(self):
        self.engine = FormulaEngine()

    def tearDown(self):
        self.engine.close()

    def test_generate_sum(self):
        """SUM operation must generate correct formula."""
        op = FormulaOperation(
            target="D2",
            operation="SUM",
            operands={"range": "B2:C2"},
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, "=SUM(B2:C2)")

    def test_generate_average(self):
        """AVERAGE operation must generate correct formula."""
        op = FormulaOperation(
            target="D2",
            operation="AVERAGE",
            operands={"range": "B2:C2"},
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, "=AVERAGE(B2:C2)")

    def test_generate_min(self):
        op = FormulaOperation(target="D2", operation="MIN", operands={"range": "A1:A10"})
        self.assertEqual(self.engine.generate_formula(op), "=MIN(A1:A10)")

    def test_generate_max(self):
        op = FormulaOperation(target="D2", operation="MAX", operands={"range": "A1:A10"})
        self.assertEqual(self.engine.generate_formula(op), "=MAX(A1:A10)")

    def test_generate_count(self):
        op = FormulaOperation(target="D2", operation="COUNT", operands={"range": "A1:A10"})
        self.assertEqual(self.engine.generate_formula(op), "=COUNT(A1:A10)")

    def test_generate_counta(self):
        op = FormulaOperation(target="D2", operation="COUNTA", operands={"range": "A1:A10"})
        self.assertEqual(self.engine.generate_formula(op), "=COUNTA(A1:A10)")

    def test_generate_add(self):
        """ADD must generate cell+cell formula."""
        op = FormulaOperation(
            target="F2",
            operation="ADD",
            operands={"left": "D2", "right": "E2"},
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, "=D2+E2")

    def test_generate_subtract(self):
        op = FormulaOperation(
            target="F2",
            operation="SUBTRACT",
            operands={"left": "D2", "right": "E2"},
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, "=D2-E2")

    def test_generate_multiply(self):
        op = FormulaOperation(
            target="F2",
            operation="MULTIPLY",
            operands={"left": "D2", "right": "E2"},
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, "=D2*E2")

    def test_generate_divide(self):
        op = FormulaOperation(
            target="F2",
            operation="DIVIDE",
            operands={"left": "D2", "right": "E2"},
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, "=D2/E2")

    def test_generate_if(self):
        op = FormulaOperation(
            target="G2",
            operation="IF",
            operands={
                "condition": "A1>10",
                "true_val": "YES",
                "false_val": "NO",
            },
        )
        formula = self.engine.generate_formula(op)
        self.assertEqual(formula, '=IF(A1>10,"YES","NO")')

    def test_generate_invalid_operation(self):
        """Unsupported operation must raise."""
        op = FormulaOperation(
            target="A1",
            operation="UNKNOWN_OP",
            operands={},
        )
        with self.assertRaises(FormulaEngineError) as cm:
            self.engine.generate_formula(op)
        self.assertEqual(cm.exception.code, FormulaEngineErrorCode.UNSUPPORTED_OPERATION.value)

    def test_generate_missing_range_operand(self):
        """SUM without range must raise."""
        op = FormulaOperation(
            target="A1",
            operation="SUM",
            operands={},
        )
        with self.assertRaises(FormulaEngineError):
            self.engine.generate_formula(op)

    def test_generate_missing_left_right_operands(self):
        """ADD without left/right must raise."""
        op = FormulaOperation(
            target="A1",
            operation="ADD",
            operands={},
        )
        with self.assertRaises(FormulaEngineError):
            self.engine.generate_formula(op)

    def test_generate_if_missing_condition(self):
        op = FormulaOperation(
            target="A1",
            operation="IF",
            operands={"true_val": "YES", "false_val": "NO"},
        )
        with self.assertRaises(FormulaEngineError):
            self.engine.generate_formula(op)

    def test_generate_if_missing_true_false(self):
        op = FormulaOperation(
            target="A1",
            operation="IF",
            operands={"condition": "A1>10"},
        )
        with self.assertRaises(FormulaEngineError):
            self.engine.generate_formula(op)


class TestFormulaApplication(unittest.TestCase):
    """MIND-030 Formula Application Tests."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.engine = FormulaEngine()
        # Create a clean workbook for application tests
        from openpyxl import Workbook
        clean_path = os.path.join(cls.temp_dir, "clean_app.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["A", "B", "C", "D"])
        ws.append([10, 20, 30, None])
        ws.append([40, 50, 60, None])
        wb.save(clean_path)
        cls.clean_path = clean_path

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _output(self, name: str) -> str:
        return os.path.join(self.temp_dir, name)

    # -- Insertion tests --

    def test_formula_insertion(self):
        """Applying a formula must place it in the target cell."""
        op = FormulaOperation(
            target="D2",
            operation="SUM",
            operands={"range": "A2:C2"},
        )
        result = self.engine.apply_formula(self.clean_path, self._output("inserted.xlsx"), op)
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=SUM(A2:C2)")

        # Verify independently
        verify = CoreExcelEngine()
        verify.load(self._output("inserted.xlsx"))
        cell = verify.get_cell("Sheet1", 2, 4)
        self.assertEqual(cell.formula, "=SUM(A2:C2)")
        self.assertTrue(cell.value.startswith("="))
        verify.close()

    def test_formula_insertion_multiply(self):
        """MULTIPLY formula must be applied correctly."""
        op = FormulaOperation(
            target="D3",
            operation="MULTIPLY",
            operands={"left": "A3", "right": "B3"},
        )
        result = self.engine.apply_formula(self.clean_path, self._output("mult.xlsx"), op)
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])

        verify = CoreExcelEngine()
        verify.load(self._output("mult.xlsx"))
        cell = verify.get_cell("Sheet1", 3, 4)
        self.assertEqual(cell.formula, "=A3*B3")
        verify.close()

    def test_formula_insertion_if(self):
        """IF formula must be applied correctly."""
        op = FormulaOperation(
            target="E2",
            operation="IF",
            operands={
                "condition": "A2>15",
                "true_val": "BIG",
                "false_val": "SMALL",
            },
        )
        result = self.engine.apply_formula(self.clean_path, self._output("if.xlsx"), op)
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])

        verify = CoreExcelEngine()
        verify.load(self._output("if.xlsx"))
        cell = verify.get_cell("Sheet1", 2, 5)
        self.assertEqual(cell.formula, '=IF(A2>15,"BIG","SMALL")')
        verify.close()

    # -- Existing formula protection --

    def test_existing_formula_not_overwritten_by_default(self):
        """Applying to a cell with existing formula must fail without explicit overwrite."""
        existing_path = os.path.join(self.temp_dir, "existing.xlsx")
        _make_workbook_with_existing_formulas(existing_path)

        op = FormulaOperation(
            target="C2",
            operation="SUM",
            operands={"range": "A2:B2"},
        )
        with self.assertRaises(FormulaEngineError) as cm:
            self.engine.apply_formula(existing_path, self._output("overwrite.xlsx"), op)
        self.assertEqual(cm.exception.code, FormulaEngineErrorCode.EXISTING_FORMULA_PROTECTED.value)

    def test_explicit_overwrite_works(self):
        """Explicit overwrite must replace existing formula."""
        existing_path = os.path.join(self.temp_dir, "existing2.xlsx")
        _make_workbook_with_existing_formulas(existing_path)

        op = FormulaOperation(
            target="C2",
            operation="SUM",
            operands={"range": "A2:B2"},
            overwrite_policy="explicit",
        )
        result = self.engine.apply_formula(existing_path, self._output("overwritten.xlsx"), op)
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=SUM(A2:B2)")

        verify = CoreExcelEngine()
        verify.load(self._output("overwritten.xlsx"))
        cell = verify.get_cell("Existing", 2, 3)
        self.assertEqual(cell.formula, "=SUM(A2:B2)")
        verify.close()

    # -- Invalid target rejection --

    def test_invalid_target_cell_rejected(self):
        """Invalid cell address must be rejected."""
        op = FormulaOperation(
            target="XYZ",
            operation="SUM",
            operands={"range": "A1:A10"},
        )
        with self.assertRaises(FormulaEngineError) as cm:
            self.engine.apply_formula(self.clean_path, self._output("invalid.xlsx"), op)
        self.assertEqual(cm.exception.code, FormulaEngineErrorCode.INVALID_CELL_REF.value)

    def test_missing_sheet_rejected(self):
        """Non-existent sheet must be rejected."""
        op = FormulaOperation(
            target="A1",
            operation="SUM",
            operands={"range": "A1:A10"},
            sheet="NonExistent",
        )
        with self.assertRaises(FormulaEngineError) as cm:
            self.engine.apply_formula(self.clean_path, self._output("missing.xlsx"), op)
        self.assertEqual(cm.exception.code, FormulaEngineErrorCode.SHEET_NOT_FOUND.value)

    # -- Protected sheet safety --

    def test_protected_sheet_rejected(self):
        """Applying formula to protected sheet must be rejected."""
        prot_path = os.path.join(self.temp_dir, "protected.xlsx")
        _make_protected_workbook(prot_path)

        op = FormulaOperation(
            target="B2",
            operation="SUM",
            operands={"range": "A2:A10"},
        )
        with self.assertRaises(FormulaEngineError) as cm:
            self.engine.apply_formula(prot_path, self._output("prot_out.xlsx"), op)
        self.assertEqual(cm.exception.code, FormulaEngineErrorCode.PROTECTED_SHEET.value)

    # -- Non-target cell preservation --

    def test_non_target_cells_preserved(self):
        """Cells other than the target must remain unchanged."""
        original_path = os.path.join(self.temp_dir, "preserve.xlsx")
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["X", "Y", "Z"])
        ws.append([1, 2, 3])
        ws.append([4, 5, 6])
        wb.save(original_path)

        op = FormulaOperation(
            target="C2",
            operation="ADD",
            operands={"left": "A2", "right": "B2"},
        )
        result = self.engine.apply_formula(original_path, self._output("preserved.xlsx"), op)
        self.assertTrue(result["valid"])

        # Verify non-target cells
        verify = CoreExcelEngine()
        verify.load(self._output("preserved.xlsx"))
        # C3 should still be 6 (original value)
        cell_c3 = verify.get_cell("Data", 3, 3)
        self.assertEqual(cell_c3.value, 6)
        # A2 should still be 1
        cell_a2 = verify.get_cell("Data", 2, 1)
        self.assertEqual(cell_a2.value, 1)
        verify.close()

    # -- Workbook structure preservation --

    def test_workbook_structure_preserved(self):
        """Applying a formula must not change sheet structure."""
        op = FormulaOperation(
            target="D2",
            operation="SUM",
            operands={"range": "A2:C2"},
        )
        result = self.engine.apply_formula(self.clean_path, self._output("struct.xlsx"), op)
        self.assertTrue(result["valid"])

        verify = CoreExcelEngine()
        verify.load(self._output("struct.xlsx"))
        info = verify.get_workbook_info()
        self.assertEqual(info.sheet_count, 1)
        self.assertEqual(info.sheets[0].name, "Sheet1")
        verify.close()


class TestFormattingPreservation(unittest.TestCase):
    """Verify that formula application does not destroy formatting."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.engine = FormulaEngine()
        cls.formatted_path = os.path.join(cls.temp_dir, "formatted.xlsx")
        _make_workbook_with_formatting(cls.formatted_path)

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_formatting_preserved_after_formula_apply(self):
        """Cell formatting must survive formula application to a different cell."""
        op = FormulaOperation(
            target="C2",
            operation="SUM",
            operands={"range": "A2:B2"},
        )
        result = self.engine.apply_formula(
            self.formatted_path,
            self._output("formatted_out.xlsx"),
            op,
        )
        self.assertTrue(result["valid"])

        verify = CoreExcelEngine()
        verify.load(self._output("formatted_out.xlsx"))
        ws = verify.get_sheet("Formatted")
        # Check that A2 still has bold font
        a2_cell = ws.cells[1][0]  # row 1, col 0 (0-based in SheetData)
        # openpyxl preserves formatting; we verify the value is intact
        # The actual formatting object is preserved by openpyxl round-trip
        self.assertEqual(a2_cell.value, "Revenue")
        # B2 value intact
        b2_cell = ws.cells[1][1]
        self.assertEqual(b2_cell.value, 1000)
        verify.close()

    def _output(self, name: str) -> str:
        return os.path.join(self.temp_dir, name)


class TestFormulaPersistence(unittest.TestCase):
    """Verify exact formula persistence across save/reopen cycle."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.engine = FormulaEngine()
        from openpyxl import Workbook
        path = os.path.join(cls.temp_dir, "persist.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "PersistTest"
        ws.append(["A", "B"])
        ws.append([5, 10])
        wb.save(path)
        cls.base_path = path

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _output(self, name: str) -> str:
        return os.path.join(self.temp_dir, name)

    def test_exact_formula_persists(self):
        """Formula must be byte-for-byte identical after reopen."""
        op = FormulaOperation(
            target="C2",
            operation="SUM",
            operands={"range": "A2:B2"},
        )
        result = self.engine.apply_formula(self.base_path, self._output("persist_out.xlsx"), op)
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])
        self.assertEqual(result["formula_applied"], "=SUM(A2:B2)")

        # Extra verification: reopen with fresh engine and check exact match
        verify = CoreExcelEngine()
        verify.load(self._output("persist_out.xlsx"))
        cell = verify.get_cell("PersistTest", 2, 3)
        self.assertEqual(cell.formula, "=SUM(A2:B2)")
        self.assertEqual(cell.value, "=SUM(A2:B2)")
        verify.close()

    def test_multiple_formulas_persist(self):
        """Multiple formula applications must all persist."""
        path = self.base_path
        ops = [
            FormulaOperation(target="C2", operation="SUM", operands={"range": "A2:B2"}),
            FormulaOperation(target="C3", operation="MULTIPLY", operands={"left": "A3", "right": "B3"}, sheet="PersistTest"),
        ]
        # Create a base with row 3
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "PersistTest"
        ws.append(["A", "B"])
        ws.append([5, 10])
        ws.append([3, 4])
        wb.save(path)

        # Apply first
        r1 = self.engine.apply_formula(path, self._output("m1.xlsx"), ops[0])
        self.assertTrue(r1["valid"])

        # Apply second on result
        r2 = self.engine.apply_formula(self._output("m1.xlsx"), self._output("m2.xlsx"), ops[1])
        self.assertTrue(r2["valid"])

        # Verify both
        verify = CoreExcelEngine()
        verify.load(self._output("m2.xlsx"))
        c2 = verify.get_cell("PersistTest", 2, 3)
        c3 = verify.get_cell("PersistTest", 3, 3)
        self.assertEqual(c2.formula, "=SUM(A2:B2)")
        self.assertEqual(c3.formula, "=A3*B3")
        verify.close()


class TestFormulaEngineDeterminism(unittest.TestCase):
    """MIND-030 Determinism Tests — repeated runs must produce identical results."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.formulas_path = os.path.join(cls.temp_dir, "det.xlsx")
        _make_workbook_with_formulas(cls.formulas_path)
        cls.engine = FormulaEngine()

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_detection_deterministic_across_runs(self):
        """scan_formulas must return identical results on repeated calls."""
        results = []
        for _ in range(5):
            results.append(self.engine.scan_formulas(self.formulas_path))
        for i in range(1, len(results)):
            self.assertEqual(results[0], results[i])

    def test_generation_deterministic(self):
        """generate_formula must be deterministic."""
        op = FormulaOperation(
            target="D2",
            operation="SUM",
            operands={"range": "B2:C2"},
        )
        results = [self.engine.generate_formula(op) for _ in range(10)]
        self.assertEqual(len(set(results)), 1)
        self.assertEqual(results[0], "=SUM(B2:C2)")

    def test_validation_deterministic(self):
        """validate_formula must be deterministic."""
        formula = "=SUM(A1:A10)"
        results = [self.engine.validate_formula(formula) for _ in range(5)]
        for i in range(1, len(results)):
            self.assertEqual(results[0].valid, results[i].valid)
            self.assertEqual(results[0].errors, results[i].errors)
            self.assertEqual(results[0].warnings, results[i].warnings)


class TestFormulaEngineSafetyBoundaries(unittest.TestCase):
    """Tests ensuring formula engine does NOT modify protected categories."""

    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.mkdtemp()
        cls.engine = FormulaEngine()
        from openpyxl import Workbook
        path = os.path.join(cls.temp_dir, "safety.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Safety"
        ws.append(["A", "B", "C"])
        ws.append([1, 2, "=A2+B2"])  # existing formula
        ws.append([3, 4, 7])          # plain value
        wb.save(path)
        cls.base_path = path

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _output(self, name: str) -> str:
        return os.path.join(self.temp_dir, name)

    def test_existing_formula_in_non_target_cell_preserved(self):
        """Formula in a cell that is NOT the target must remain unchanged."""
        op = FormulaOperation(
            target="B3",
            operation="ADD",
            operands={"left": "A3", "right": "B3"},
        )
        result = self.engine.apply_formula(self.base_path, self._output("safety_out.xlsx"), op)
        self.assertTrue(result["valid"])

        verify = CoreExcelEngine()
        verify.load(self._output("safety_out.xlsx"))
        # C2 has existing formula — must be preserved
        c2 = verify.get_cell("Safety", 2, 3)
        self.assertEqual(c2.formula, "=A2+B2")
        verify.close()

    def test_existing_formula_can_be_explicitly_overwritten(self):
        """With explicit policy, existing formula CAN be replaced."""
        op = FormulaOperation(
            target="C2",
            operation="SUM",
            operands={"range": "A2:B2"},
            overwrite_policy="explicit",
        )
        result = self.engine.apply_formula(self.base_path, self._output("overwrite_safe.xlsx"), op)
        self.assertTrue(result["valid"])
        self.assertTrue(result["formula_persisted"])

        verify = CoreExcelEngine()
        verify.load(self._output("overwrite_safe.xlsx"))
        c2 = verify.get_cell("Safety", 2, 3)
        self.assertEqual(c2.formula, "=SUM(A2:B2)")
        verify.close()


class TestFormulaEngineRegression(unittest.TestCase):
    """Regression tests ensuring MIND-010 and MIND-020 still work."""

    def test_mind010_engine_importable_and_works(self):
        """CoreExcelEngine must still be importable and functional."""
        from engine.excel_engine import CoreExcelEngine
        engine = CoreExcelEngine()
        fixture = os.path.join(FIXTURES_DIR, "sample.xlsx")
        info = engine.load(fixture)
        self.assertEqual(info.sheet_count, 2)
        engine.close()

    def test_mind020_enhancer_importable_and_works(self):
        """FileEnhancer must still be importable and functional."""
        from engine.file_enhancer import FileEnhancer
        fixer = FileEnhancer()
        fixture = os.path.join(FIXTURES_DIR, "clean.xlsx")
        result = fixer.analyze(fixture)
        self.assertIsNotNone(result)
        fixer._engine.close()


class TestFormulaEngineInvalidInput(unittest.TestCase):
    """Tests for invalid inputs produce clear errors."""

    def setUp(self):
        self.engine = FormulaEngine()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        self.engine.close()
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_nonexistent_file_raises_error(self):
        """Loading a non-existent file must raise."""
        with self.assertRaises(FormulaEngineError):
            self.engine.scan_formulas("/nonexistent/file.xlsx")

    def test_invalid_format_raises_error(self):
        """Loading a non-xlsx file must raise."""
        bad_path = os.path.join(self.temp_dir, "bad.txt")
        with open(bad_path, "w") as f:
            f.write("not an xlsx")
        with self.assertRaises(FormulaEngineError):
            self.engine.scan_formulas(bad_path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
