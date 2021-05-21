import os.path
import openpyxl

from .utils import (
    get_globally_defined_name,
    get_defined_name,
    get_table,
    add_sheet_to_reference,
    triangulate_cell,
    copy_value,    
)

def get_test_workbook():
    filename = os.path.join(os.path.dirname(__file__), 'test_data', 'source.xlsx')
    return openpyxl.load_workbook(filename, data_only=True)

def test_get_globally_defined_name():
    wb = get_test_workbook()

    assert get_globally_defined_name(wb, "PROFIT_RANGE") is not None
    assert get_globally_defined_name(wb, "FOOBAR") is None

def test_get_defined_name():
    wb = get_test_workbook()
    ws = wb['Report 1']

    assert get_defined_name(wb, ws, "PROFIT_RANGE") is not None
    assert get_defined_name(wb, ws, "FOOBAR") is None

def test_get_table():
    wb = get_test_workbook()
    ws = wb['Report 2']

    assert get_table(ws, "RangleTable") is not None
    assert get_table(ws, "NotFound") is None

def test_add_sheet_to_reference():
    wb = get_test_workbook()
    ws = wb['Report 1']

    assert add_sheet_to_reference(ws, "B3:C4") == "'Report 1'!B3:C4"

def test_triangulate_cell():
    wb = get_test_workbook()
    ws = wb['Report 1']

    row = ws.cell(row=3, column=5)
    col = ws.cell(row=6, column=8)

    cell = triangulate_cell(row, col)

    assert cell.row == 3
    assert cell.column == 8

def test_copy_value():
    wb = get_test_workbook()
    ws = wb['Report 1']

    c1 = ws.cell(row=3, column=2)
    c2 = ws.cell(row=4, column=6)

    assert c1.value == "Date"
    assert c2.value != "Date"

    copy_value(c1, c2)

    assert c1.value == "Date"
    assert c2.value == "Date"

