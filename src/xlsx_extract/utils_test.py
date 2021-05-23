import datetime
import os.path
import openpyxl

from .utils import (
    get_globally_defined_name,
    get_defined_name,
    get_named_table,
    add_sheet_to_reference,
    resize_table,
    triangulate_cell,
    copy_value,
)

from .range import Range

def get_test_workbook(filename='source.xlsx', data_only=True):
    filename = os.path.join(os.path.dirname(__file__), 'test_data', filename)
    return openpyxl.load_workbook(filename, data_only=data_only)

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

class TestResizeTable:

    def get_table(self):
        wb = get_test_workbook()
        ws = wb['Report 1']

        # Full table: 5 rows by 4  columns
        full_table = tuple(ws.iter_rows(min_row=5, min_col=2, max_row=9, max_col=6))

        assert len(full_table) == 5
        assert [c.value for c in full_table[0]] == [None, "Jan", "Feb", "Mar", "Apr"]
        assert [c.value for c in full_table[1]] == ["Alpha", 1.5, 6, 11, 4.6]
        assert [c.value for c in full_table[2]] == ["Beta", 2, 7, 12, 4.7]
        assert [c.value for c in full_table[3]] == ["Delta", 2.5, 8, 13, 4.8]
        assert [c.value for c in full_table[4]] == ["Gamma", 3, 9, 14, 4.9]

        # Small table: 3 rows by 2 columns
        table = tuple(ws.iter_rows(min_row=5, min_col=3, max_row=7, max_col=4))

        assert len(table) == 3
        assert [c.value for c in table[0]] == ["Jan", "Feb"]
        assert [c.value for c in table[1]] == [1.5, 6]
        assert [c.value for c in table[2]] == [2, 7]

        return ws, Range(table)

    def test_add_rows(self):
        ws, table = self.get_table()

        # Resize small table: add two rows
        new_table = resize_table(table, rows=5, cols=2).cells

        # New table
        assert len(new_table) == 5
        assert [c.value for c in new_table[0]] == ["Jan", "Feb"]
        assert [c.value for c in new_table[1]] == [1.5, 6]
        assert [c.value for c in new_table[2]] == [2, 7]
        assert [c.value for c in new_table[3]] == [None, None]
        assert [c.value for c in new_table[4]] == [None, None]

        full_table = tuple(ws.iter_rows(min_row=5, min_col=2, max_row=11, max_col=6))

        assert len(full_table) == 7
        assert [c.value for c in full_table[0]] == [None, "Jan", "Feb", "Mar", "Apr"]
        assert [c.value for c in full_table[1]] == ["Alpha", 1.5, 6, 11, 4.6]
        assert [c.value for c in full_table[2]] == ["Beta", 2, 7, 12, 4.7]
        assert [c.value for c in full_table[3]] == [None, None, None, None, None]
        assert [c.value for c in full_table[4]] == [None, None, None, None, None]
        assert [c.value for c in full_table[5]] == ["Delta", 2.5, 8, 13, 4.8]
        assert [c.value for c in full_table[6]] == ["Gamma", 3, 9, 14, 4.9]
    
    def test_remove_rows(self):
        ws, table = self.get_table()

        # Resize small table: remove one row
        new_table = resize_table(table, rows=2, cols=2).cells

        # New table
        assert len(new_table) == 2
        assert [c.value for c in new_table[0]] == ["Jan", "Feb"]
        assert [c.value for c in new_table[1]] == [1.5, 6]

        full_table = tuple(ws.iter_rows(min_row=5, min_col=2, max_row=8, max_col=6))

        assert len(full_table) == 4
        assert [c.value for c in full_table[0]] == [None, "Jan", "Feb", "Mar", "Apr"]
        assert [c.value for c in full_table[1]] == ["Alpha", 1.5, 6, 11, 4.6]
        assert [c.value for c in full_table[2]] == ["Delta", 2.5, 8, 13, 4.8]
        assert [c.value for c in full_table[3]] == ["Gamma", 3, 9, 14, 4.9]
    
    def test_add_cols(self):
        ws, table = self.get_table()

        # Resize small table: add two columns
        new_table = resize_table(table, rows=3, cols=4).cells

        # New table
        assert len(new_table) == 3
        assert [c.value for c in new_table[0]] == ["Jan", "Feb", None, None]
        assert [c.value for c in new_table[1]] == [1.5, 6, None, None]
        assert [c.value for c in new_table[2]] == [2, 7, None, None]

        full_table = tuple(ws.iter_rows(min_row=5, min_col=2, max_row=9, max_col=8))

        assert len(full_table) == 5
        assert [c.value for c in full_table[0]] == [None, "Jan", "Feb", None, None, "Mar", "Apr"]
        assert [c.value for c in full_table[1]] == ["Alpha", 1.5, 6, None, None, 11, 4.6]
        assert [c.value for c in full_table[2]] == ["Beta", 2, 7, None, None, 12, 4.7]
        assert [c.value for c in full_table[3]] == ["Delta", 2.5, 8, None, None, 13, 4.8]
        assert [c.value for c in full_table[4]] == ["Gamma", 3, 9, None, None, 14, 4.9]
    
    def test_remove_cols(self):
        ws, table = self.get_table()

        # Resize small table: remove one column
        new_table = resize_table(table, rows=3, cols=1).cells

        # New table
        assert len(new_table) == 3
        assert [c.value for c in new_table[0]] == ["Jan"]
        assert [c.value for c in new_table[1]] == [1.5]
        assert [c.value for c in new_table[2]] == [2]

        full_table = tuple(ws.iter_rows(min_row=5, min_col=2, max_row=9, max_col=5))

        assert len(full_table) == 5
        assert [c.value for c in full_table[0]] == [None, "Jan", "Mar", "Apr"]
        assert [c.value for c in full_table[1]] == ["Alpha", 1.5, 11, 4.6]
        assert [c.value for c in full_table[2]] == ["Beta", 2, 12, 4.7]
        assert [c.value for c in full_table[3]] == ["Delta", 2.5, 13, 4.8]
        assert [c.value for c in full_table[4]] == ["Gamma", 3, 14, 4.9]
    
    def test_change_both_dimensions(self):
        ws, table = self.get_table()
        
        # Remove one row, add two columns
        new_table = resize_table(table, rows=2, cols=4).cells

        assert len(new_table) == 2
        assert [c.value for c in new_table[0]] == ["Jan", "Feb", None, None]
        assert [c.value for c in new_table[1]] == [1.5, 6, None, None]

        full_table = tuple(ws.iter_rows(min_row=5, min_col=2, max_row=8, max_col=8))

        assert len(full_table) == 4
        assert [c.value for c in full_table[0]] == [None, "Jan", "Feb", None, None, "Mar", "Apr"]
        assert [c.value for c in full_table[1]] == ["Alpha", 1.5, 6, None, None, 11, 4.6]
        assert [c.value for c in full_table[2]] == ["Delta", 2.5, 8, None, None, 13, 4.8]
        assert [c.value for c in full_table[3]] == ["Gamma", 3, 9, None, None, 14, 4.9]
    
    def test_resize_defined_name_table(self):
        wb = get_test_workbook()

        defined_name = get_defined_name(wb, None, "PROFIT_RANGE")
        ref = defined_name.attr_text

        sheet_name, (c1, r1, c2, r2) = openpyxl.utils.cell.range_to_tuple(ref)
        table = tuple(wb[sheet_name].iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

        assert len(table) == 5
        assert [c.value for c in table[0]] == [None, 'Profit', None, 'Loss', None]
        assert [c.value for c in table[1]] == [None, '£', 'Plan', '£', 'Plan']
        assert [c.value for c in table[2]] == ['Alpha', 100, 100, 50, 20]
        assert [c.value for c in table[3]] == ['Beta', 200, 150, 50, 20]
        assert [c.value for c in table[4]] == ['Delta', 300, 350, 50, 20]

        # add two columns, remove one row
        new_table = resize_table(Range(table, defined_name=defined_name), rows=4, cols=7).cells

        assert len(new_table) == 4
        assert [c.value for c in new_table[0]] == [None, 'Profit', None, 'Loss', None, None, None]
        assert [c.value for c in new_table[1]] == [None, '£', 'Plan', '£', 'Plan', None, None]
        assert [c.value for c in new_table[2]] == ['Alpha', 100, 100, 50, 20, None, None]
        assert [c.value for c in new_table[3]] == ['Beta', 200, 150, 50, 20, None, None]

        # check that the named range now resolves to the new table
        defined_name = get_defined_name(wb, None, "PROFIT_RANGE")
        ref = defined_name.attr_text

        sheet_name, (c1, r1, c2, r2) = openpyxl.utils.cell.range_to_tuple(ref)
        ref_table = tuple(wb[sheet_name].iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

        assert len(ref_table) == 4
        assert [c.value for c in ref_table[0]] == [None, 'Profit', None, 'Loss', None, None, None]
        assert [c.value for c in ref_table[1]] == [None, '£', 'Plan', '£', 'Plan', None, None]
        assert [c.value for c in ref_table[2]] == ['Alpha', 100, 100, 50, 20, None, None]
        assert [c.value for c in ref_table[3]] == ['Beta', 200, 150, 50, 20, None, None]
    
    def test_resize_named_table(self):
        wb = get_test_workbook()
        ws = wb['Report 2']

        named_table = get_named_table(ws, 'RangleTable')
        ref = add_sheet_to_reference(ws, named_table.ref)

        sheet_name, (c1, r1, c2, r2) = openpyxl.utils.cell.range_to_tuple(ref)
        table = tuple(wb[sheet_name].iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

        assert len(table) == 4
        assert [c.value for c in table[0]] == ['Name', 'Date', 'Range', 'Price']
        assert [c.value for c in table[1]] == ['Bill', datetime.datetime(2021, 1, 1), 9, 15]
        assert [c.value for c in table[2]] == ['Bob', datetime.datetime(2021, 3, 2), 14, 18]
        assert [c.value for c in table[3]] == ['Joan', datetime.datetime(2021, 6, 5), 13, 99]

        # Remove a column, add two rows

        new_table = resize_table(Range(table, named_table=named_table), rows=6, cols=3).cells

        assert len(new_table) == 6
        assert [c.value for c in new_table[0]] == ['Name', 'Date', 'Range']
        assert [c.value for c in new_table[1]] == ['Bill', datetime.datetime(2021, 1, 1), 9]
        assert [c.value for c in new_table[2]] == ['Bob', datetime.datetime(2021, 3, 2), 14]
        assert [c.value for c in new_table[3]] == ['Joan', datetime.datetime(2021, 6, 5), 13]
        assert [c.value for c in new_table[4]] == [None, None, None]
        assert [c.value for c in new_table[5]] == [None, None, None]

        # Check the named reference was updated

        named_table = get_named_table(ws, 'RangleTable')
        ref = add_sheet_to_reference(ws, named_table.ref)

        sheet_name, (c1, r1, c2, r2) = openpyxl.utils.cell.range_to_tuple(ref)
        ref_table = tuple(wb[sheet_name].iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

        assert len(ref_table) == 6
        assert [c.value for c in ref_table[0]] == ['Name', 'Date', 'Range']
        assert [c.value for c in ref_table[1]] == ['Bill', datetime.datetime(2021, 1, 1), 9]
        assert [c.value for c in ref_table[2]] == ['Bob', datetime.datetime(2021, 3, 2), 14]
        assert [c.value for c in ref_table[3]] == ['Joan', datetime.datetime(2021, 6, 5), 13]
        assert [c.value for c in ref_table[4]] == [None, None, None]
        assert [c.value for c in ref_table[5]] == [None, None, None]
