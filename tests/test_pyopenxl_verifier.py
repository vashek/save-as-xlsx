from pathlib import Path
from typing import Any

try:
    import openpyxl
except ImportError:
    openpyxl = None


def verify_using_pyopenxl(fn: Path | str, dimensions: str, data: list[tuple[Any, ...]] | None = None, sheet_name: str = "Sheet1", table_name: str = "Table1"):
    if openpyxl is None:
        return
    xl = openpyxl.load_workbook(fn)
    sheet = xl[sheet_name]
    if dimensions:
        assert sheet.dimensions == dimensions
    if data is not None:
        assert list(sheet.values) == data
    table = sheet.tables[table_name]
    assert table.name == table_name
    assert table.displayName == table_name
    if data is not None:
        assert table.column_names == list(data[0])
    if dimensions:
        assert table.ref == dimensions
