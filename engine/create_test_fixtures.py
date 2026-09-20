"""
MIND-020 Test Fixtures Generator

Creates test workbooks for the File Enhancer milestone.
Each fixture is saved to the fixtures/ directory.
"""

from openpyxl import Workbook

FIXTURES_DIR = "/data/data/com.termux/files/home/memind/fixtures"


def create_clean_workbook(path: str = f"{FIXTURES_DIR}/clean.xlsx") -> str:
    """Create a clean workbook with no issues."""
    wb = Workbook()
    ws = wb.active
    ws.title = "CleanData"
    ws.append(["Name", "Age", "City"])
    ws.append(["Alice", 30, "New York"])
    ws.append(["Bob", 25, "London"])
    ws.append(["Charlie", 35, "Paris"])
    wb.save(path)
    return path


def create_duplicate_rows_workbook(
    path: str = f"{FIXTURES_DIR}/duplicates.xlsx"
) -> str:
    """Create a workbook with exact duplicate rows."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Product", "Quantity", "Price"])
    ws.append(["Widget A", 100, 25.50])
    ws.append(["Widget B", 200, 15.75])
    ws.append(["Widget A", 100, 25.50])  # duplicate of row 2
    ws.append(["Gadget C", 50, 99.99])
    ws.append(["Widget B", 200, 15.75])  # duplicate of row 3
    wb.save(path)
    return path


def create_missing_values_workbook(
    path: str = f"{FIXTURES_DIR}/missing_values.xlsx"
) -> str:
    """Create a workbook with missing/empty cells."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory"
    ws.append(["Item", "Stock", "Location"])
    ws.append(["Widget A", 500, "Warehouse A"])
    ws.append(["Widget B", None, None])  # missing stock and location
    ws.append(["Gadget C", 300, ""])      # empty string location
    ws.append(["Widget D", None, "Warehouse B"])
    wb.save(path)
    return path


def create_whitespace_workbook(
    path: str = f"{FIXTURES_DIR}/whitespace.xlsx"
) -> str:
    """Create a workbook with whitespace issues."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(["  Name  ", "  Category  ", " Price "])  # padded headers
    ws.append(["  Alpha  ", "Beta", 10.5])
    ws.append(["Gamma  ", "Beta", 20.0])
    ws.append(["  Delta  ", "Alpha", 30.0])
    wb.save(path)
    return path


def create_duplicate_headers_workbook(
    path: str = f"{FIXTURES_DIR}/duplicate_headers.xlsx"
) -> str:
    """Create a workbook with duplicate and empty headers."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Name", "Age", "Name"])  # duplicate header
    ws.append(["Alice", 30, "Engineer"])
    ws.append(["Bob", None, "Manager"])  # empty header row 2 col 2 (actually header row 1 has no empty)
    # Actually let's add an empty header too
    ws2 = wb.create_sheet("BadHeaders")
    ws2.append(["", "Age", "City"])  # empty header
    ws2.append(["Alice", 30, "NYC"])
    ws2.append(["Bob", 25, "LA"])
    wb.save(path)
    return path


def create_mixed_types_workbook(
    path: str = f"{FIXTURES_DIR}/mixed_types.xlsx"
) -> str:
    """Create a workbook with mixed type columns."""
    wb = Workbook()
    ws = wb.active
    ws.title = "MixedData"
    ws.append(["ID", "Value", "Label"])
    ws.append([1, 100, "numeric"])
    ws.append([2, "text", "text"])
    ws.append([3, 200, 42])  # numeric label is actually int
    ws.append([4, 3.14, None])
    wb.save(path)
    return path


def create_formulas_workbook(
    path: str = f"{FIXTURES_DIR}/formulas.xlsx"
) -> str:
    """Create a workbook with formula cells."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws.append(["A", "B", "Total"])
    ws.append([10, 20, "=SUM(A2:B2)"])
    ws.append([30, 40, "=SUM(A3:B3)"])
    ws.append([50, 60, "=A4+B4"])
    wb.save(path)
    return path


def create_empty_workbook(
    path: str = f"{FIXTURES_DIR}/empty.xlsx"
) -> str:
    """Create a workbook with a sheet but no data."""
    wb = Workbook()
    ws = wb.active
    ws.title = "EmptySheet"
    # No append calls — empty sheet
    wb.save(path)
    return path


def create_all_fixtures() -> list[str]:
    """Create all MIND-020 test fixtures."""
    fixtures = [
        create_clean_workbook(),
        create_duplicate_rows_workbook(),
        create_missing_values_workbook(),
        create_whitespace_workbook(),
        create_duplicate_headers_workbook(),
        create_mixed_types_workbook(),
        create_formulas_workbook(),
        create_empty_workbook(),
    ]
    return fixtures


if __name__ == "__main__":
    created = create_all_fixtures()
    for f in created:
        print(f"Created: {f}")
