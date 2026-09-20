"""
MeMind Core Excel Engine — Sample Fixture

Generates sample.xlsx for MIND-010 testing.
"""

from openpyxl import Workbook


def create_sample_workbook(path: str = "/data/data/com.termux/files/home/memind/fixtures/sample.xlsx") -> str:
    """Create a sample .xlsx file with test data."""
    wb = Workbook()

    # Sheet 1: Sales Data
    ws1 = wb.active
    ws1.title = "Sales"
    ws1.append(["Product", "Quantity", "Price"])
    ws1.append(["Widget A", 100, 25.50])
    ws1.append(["Widget B", 200, 15.75])
    ws1.append(["Gadget C", 50, 99.99])

    # Sheet 2: Inventory
    ws2 = wb.create_sheet(title="Inventory")
    ws2.append(["Item", "Stock"])
    ws2.append(["Widget A", 500])
    ws2.append(["Widget B", 300])

    wb.save(path)
    return path


if __name__ == "__main__":
    create_sample_workbook()
    print(f"Created: {create_sample_workbook.__defaults__[0]}")
