import os.path
import openpyxl

from .utils import (
    get_globally_defined_name,
    get_defined_name,
    get_named_table,
    add_sheet_to_reference,
    triangulate_cell,
    copy_value,
    get_reference_for_table,
    update_name
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

def test_get_named_table():
    wb = get_test_workbook()
    ws = wb['Report 2']

    assert get_named_table(ws, "RangleTable") is not None
    assert get_named_table(ws, "NotFound") is None

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

    c1 = ws.cell(row=3, column=2)
    c2 = ws.cell(row=4, column=6)
    
    assert c1.value == "Date"
    assert c2.value == "Date"

def test_get_reference_for_table():
    wb = get_test_workbook()
    ws = wb['Report 1']

    table = tuple(ws.iter_rows(min_row=2, min_col=3, max_row=5, max_col=6))
    assert get_reference_for_table(table) == "'Report 1'!$C$2:$F$5"

    table = tuple(ws.iter_rows(min_row=2, min_col=3, max_row=2, max_col=3))
    assert get_reference_for_table(table) == "'Report 1'!$C$2"

def test_update_name():
    wb = get_test_workbook()
    ws = wb['Report 2']

    defined_name = get_defined_name(wb, ws, "PROFIT_RANGE")
    assert defined_name.attr_text == "'Report 3'!$A$1:$E$5"

    table = tuple(ws.iter_rows(min_row=2, min_col=3, max_row=5, max_col=6))
    assert update_name(ws, "PROFIT_RANGE", table) == True
    
    defined_name = get_defined_name(wb, ws, "PROFIT_RANGE")
    assert defined_name.attr_text == "'Report 2'!$C$2:$F$5"

    assert update_name(ws, "NOT_FOUND", table) == False

    named_table = get_named_table(ws, "RangleTable")
    assert named_table.ref == "B10:E13"

    assert update_name(ws, "RangleTable", table) == True
    named_table = get_named_table(ws, "RangleTable")
    assert named_table.ref == "C2:F5"
