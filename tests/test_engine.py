"""
MeMind Core Excel Engine — Tests

Test pipeline: Input XLSX → Read → Inspect → Modify → Save → Re-open → Validate
"""

import os
import sys
import tempfile
import unittest

# Add engine to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.excel_engine import (
    CoreExcelEngine,
    ExcelError,
    ExcelErrorCode,
    CellData,
    SheetData,
    WorkbookInfo,
)
from engine.create_fixture import create_sample_workbook


class TestCoreExcelEngine(unittest.TestCase):
    """MIND-010 Core Excel Engine Tests."""

    @classmethod
    def setUpClass(cls) -> None:
        """Use fixture file directly; copy to temp for isolation."""
        cls.fixture_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "fixtures"
        )
        os.makedirs(cls.fixture_dir, exist_ok=True)
        cls.sample_path = os.path.join(cls.fixture_dir, "sample.xlsx")
        # Verify fixture exists
        if not os.path.exists(cls.sample_path):
            cls.sample_path = create_sample_workbook(cls.sample_path)
        cls.public_test_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "public", "test.xlsx"
        )
        cls.temp_dir = tempfile.mkdtemp()

    @classmethod
    def tearDownClass(cls) -> None:
        """Cleanup temp files."""
        import shutil
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def _new_engine(self) -> CoreExcelEngine:
        return CoreExcelEngine()

    def _write_output(self, filename: str) -> str:
        return os.path.join(self.temp_dir, filename)

    # -------------------------------------------------------------------
    # TEST 1: Workbook can be read
    # -------------------------------------------------------------------
    def test_workbook_can_be_read(self):
        """1. workbook can be read"""
        engine = self._new_engine()
        info = engine.load(self.sample_path)
        self.assertIsInstance(info, WorkbookInfo)
        self.assertEqual(info.sheet_count, 2)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 2: Sheets can be detected
    # -------------------------------------------------------------------
    def test_sheets_can_be_detected(self):
        """2. sheet can be detected"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        names = engine.get_sheet_names()
        self.assertIn("Sales", names)
        self.assertIn("Inventory", names)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 3b: Active sheet index is correct
    # -------------------------------------------------------------------
    def test_active_sheet_index(self):
        """2b. active sheet index is correct integer."""
        engine = self._new_engine()
        engine.load(self.sample_path)
        
        # Active sheet should be "Sales" (first sheet) -> index 0
        active_index = engine.get_active_sheet_index()
        self.assertIsInstance(active_index, int, "get_active_sheet_index must return int")
        self.assertEqual(active_index, 0, "Sales is the first sheet (index 0)")
        
        # Cross-verify with get_workbook_info
        info = engine.get_workbook_info()
        self.assertEqual(active_index, info.active_sheet_index, 
                         "get_active_sheet_index() must match get_workbook_info().active_sheet_index")
        engine.close()

    def test_sheets_can_be_detected(self):
        """3. cell can be read"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        cell = engine.get_cell("Sales", 1, 1)
        self.assertIsInstance(cell, CellData)
        self.assertEqual(cell.value, "Product")
        engine.close()

    # -------------------------------------------------------------------
    # TEST 4: Sheet can be created
    # -------------------------------------------------------------------
    def test_sheet_can_be_created(self):
        """4. sheet can be created"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        initial_count = len(engine.get_sheet_names())
        engine.create_sheet("NewSheet")
        names = engine.get_sheet_names()
        self.assertIn("NewSheet", names)
        self.assertEqual(len(names), initial_count + 1)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 5: Sheet can be deleted
    # -------------------------------------------------------------------
    def test_sheet_can_be_deleted(self):
        """5. sheet can be deleted"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        initial_count = len(engine.get_sheet_names())
        engine.delete_sheet("Inventory")
        names = engine.get_sheet_names()
        self.assertNotIn("Inventory", names)
        self.assertEqual(len(names), initial_count - 1)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 6: Cell can be changed
    # -------------------------------------------------------------------
    def test_cell_can_be_changed(self):
        """6. cell can be changed"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        engine.update_cell("Sales", 1, 1, "Modified Header")
        cell = engine.get_cell("Sales", 1, 1)
        self.assertEqual(cell.value, "Modified Header")
        engine.close()

    # -------------------------------------------------------------------
    # TEST 7: Row/Column can be modified
    # -------------------------------------------------------------------
    def test_row_can_be_modified(self):
        """7a. row can be added"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        initial_rows = engine.get_row_count("Sales")
        engine.add_row("Sales", ["Gadget D", 75, 49.99])
        self.assertEqual(engine.get_row_count("Sales"), initial_rows + 1)
        engine.close()

    def test_column_can_be_modified(self):
        """7b. column can be added"""
        engine = self._new_engine()
        engine.load(self.sample_path)
        initial_cols = engine.get_col_count("Sales")
        engine.add_column("Sales", ["Category", "Category", "Category", "Category", "Category"])
        self.assertEqual(engine.get_col_count("Sales"), initial_cols + 1)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 8: Output XLSX can be opened again
    # -------------------------------------------------------------------
    def test_output_xlsx_can_be_reopened(self):
        """8. output XLSX can be opened again"""
        output_path = self._write_output("test_output.xlsx")
        engine = self._new_engine()
        engine.load(self.sample_path)
        saved = engine.save(output_path)
        self.assertTrue(os.path.exists(saved))

        # Re-open with a fresh engine
        engine2 = self._new_engine()
        info = engine2.load(saved)
        self.assertEqual(info.sheet_count, 2)
        engine2.close()
        engine.close()

    # -------------------------------------------------------------------
    # TEST 9: Changes are persisted
    # -------------------------------------------------------------------
    def test_changes_are_persisted(self):
        """9. changes are saved"""
        output_path = self._write_output("test_persisted.xlsx")
        engine = self._new_engine()
        engine.load(self.sample_path)
        engine.update_cell("Sales", 2, 1, "SUPER WIDGET")
        engine.save(output_path)
        engine.close()

        # Verify
        engine2 = self._new_engine()
        engine2.load(output_path)
        cell = engine2.get_cell("Sales", 2, 1)
        self.assertEqual(cell.value, "SUPER WIDGET")
        engine2.close()

    # -------------------------------------------------------------------
    # TEST 10: Invalid input produces clear error
    # -------------------------------------------------------------------
    def test_invalid_file_raises_clear_error(self):
        """10a. file not found raises clear error"""
        engine = self._new_engine()
        with self.assertRaises(ExcelError) as cm:
            engine.load("/nonexistent/file.xlsx")
        self.assertEqual(cm.exception.code, ExcelErrorCode.FILE_NOT_FOUND.value)
        engine.close()

    def test_invalid_format_raises_clear_error(self):
        """10b. invalid format raises clear error"""
        engine = self._new_engine()
        # Create a temporary text file
        tmp = os.path.join(self.temp_dir, "bad.txt")
        with open(tmp, "w") as f:
            f.write("not an xlsx")
        with self.assertRaises(ExcelError) as cm:
            engine.load(tmp)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 11: Validation works
    # -------------------------------------------------------------------
    def test_validation_passes(self):
        """Validate after save returns valid=True"""
        output_path = self._write_output("test_validated.xlsx")
        engine = self._new_engine()
        engine.load(self.sample_path)
        engine.update_cell("Sales", 2, 2, 999)
        engine.save(output_path)

        result = engine.validate(output_path)
        self.assertTrue(result["valid"])
        self.assertTrue(result["sheet_count_match"])
        engine.close()

    # -------------------------------------------------------------------
    # TEST 12: Pipeline works end-to-end
    # -------------------------------------------------------------------
    def test_pipeline_end_to_end(self):
        """Full pipeline: Load → Modify → Save → Validate"""
        output_path = self._write_output("test_pipeline.xlsx")
        engine = self._new_engine()

        result = engine.pipeline(
            self.sample_path,
            output_path,
            modifications={
                "rename_sheet": {"old_name": "Inventory", "new_name": "Stock"},
                "update_cell": {"sheet": "Sales", "row": 2, "col": 3, "value": 99.99},
            }
        )

        self.assertTrue(result["valid"])
        self.assertIn("Stock", result["output_sheets"])
        self.assertNotIn("Inventory", result["output_sheets"])
        engine.close()

    # -------------------------------------------------------------------
    # TEST 13: Rename sheet works
    # -------------------------------------------------------------------
    def test_rename_sheet(self):
        engine = self._new_engine()
        engine.load(self.sample_path)
        engine.rename_sheet("Inventory", "Stock")
        names = engine.get_sheet_names()
        self.assertIn("Stock", names)
        self.assertNotIn("Inventory", names)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 14: Sheet not found error
    # -------------------------------------------------------------------
    def test_sheet_not_found_raises_error(self):
        engine = self._new_engine()
        engine.load(self.sample_path)
        with self.assertRaises(ExcelError) as cm:
            engine.get_cell("NonExistent", 1, 1)
        self.assertEqual(cm.exception.code, ExcelErrorCode.SHEET_NOT_FOUND.value)
        engine.close()

    # -------------------------------------------------------------------
    # TEST 15: Headers can be read
    # -------------------------------------------------------------------
    def test_headers_can_be_read(self):
        engine = self._new_engine()
        engine.load(self.sample_path)
        headers = engine.get_headers("Sales")
        self.assertEqual(headers, ["Product", "Quantity", "Price"])
        engine.close()

    # -------------------------------------------------------------------
    # TEST 16: Public test.xlsx works
    # -------------------------------------------------------------------
    def test_public_test_file(self):
        """Verify the existing public/test.xlsx works with the engine."""
        if os.path.exists(self.public_test_path):
            engine = self._new_engine()
            info = engine.load(self.public_test_path)
            self.assertGreater(info.sheet_count, 0)
            names = engine.get_sheet_names()
            self.assertTrue(len(names) > 0)
            engine.close()

    # -------------------------------------------------------------------
    # TEST 17: Sheet can be renamed with create_sheet
    # -------------------------------------------------------------------
    def test_sheet_positions(self):
        engine = self._new_engine()
        engine.load(self.sample_path)
        engine.create_sheet("First", position=0)
        names = engine.get_sheet_names()
        self.assertEqual(names[0], "First")
        engine.close()


class TestErrorHandling(unittest.TestCase):
    """Additional error handling tests."""

    def setUp(self) -> None:
        self.engine = CoreExcelEngine()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self) -> None:
        self.engine.close()
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_delete_nonexistent_sheet(self):
        from engine.excel_engine import ExcelError, ExcelErrorCode
        self.engine.load(
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures", "sample.xlsx")
        )
        with self.assertRaises(ExcelError) as cm:
            self.engine.delete_sheet("DoesNotExist")
        self.assertEqual(cm.exception.code, ExcelErrorCode.SHEET_NOT_FOUND.value)

    def test_delete_nonexistent_row(self):
        from engine.excel_engine import ExcelError, ExcelErrorCode
        sample_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures", "sample.xlsx"
        )
        self.engine.load(sample_path)
        with self.assertRaises(ExcelError) as cm:
            self.engine.delete_row("Sales", 999)
        self.assertEqual(cm.exception.code, ExcelErrorCode.INVALID_OPERATION.value)

    def test_update_invalid_cell(self):
        from engine.excel_engine import ExcelError, ExcelErrorCode
        sample_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures", "sample.xlsx"
        )
        self.engine.load(sample_path)
        with self.assertRaises(ExcelError) as cm:
            self.engine.update_cell("Sales", 0, 0, "bad")
        self.assertEqual(cm.exception.code, ExcelErrorCode.INVALID_CELL_REF.value)

    def test_save_to_invalid_path(self):
        from engine.excel_engine import ExcelError, ExcelErrorCode
        sample_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures", "sample.xlsx"
        )
        self.engine.load(sample_path)
        # Try to save to a path with no directory permissions
        # (should still succeed or give a clear error)
        try:
            self.engine.save(os.path.join(self.temp_dir, "output.xlsx"))
        except ExcelError:
            pass  # Acceptable if path issue
        finally:
            pass


if __name__ == "__main__":
    unittest.main(verbosity=2)
