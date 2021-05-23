import os.path
import openpyxl

from . import range, utils

def get_test_workbook():
    filename = os.path.join(os.path.dirname(__file__), 'test_data', 'source.xlsx')
    return openpyxl.load_workbook(filename, data_only=True)

class TestRange:

    def test_empty(self):
        r = range.Range(())
        
        assert r.is_empty
        assert not r.is_cell
        assert not r.is_range

        assert r.workbook is None
        assert r.sheet is None
        assert r.cell is None

        assert r.get_reference() is None
    
    def test_single_cell(self):
        wb = get_test_workbook()
        ws = wb['Report 1']
        cells = tuple(ws.iter_rows(2, 2, 2, 2))

        r = range.Range(cells)

        assert not r.is_empty
        assert r.is_cell
        assert not r.is_range

        assert r.workbook is wb
        assert r.sheet is ws
        assert r.cell is cells[0][0]

        assert r.get_reference() == "'Report 1'!$B$2"
        assert r.get_reference(absolute=False) == "'Report 1'!B2"
        assert r.get_reference(use_sheet=False) == "$B$2"
    
    def test_range_cell(self):
        wb = get_test_workbook()
        ws = wb['Report 1']
        cells = tuple(ws.iter_rows(2, 3, 2, 3))

        r = range.Range(cells)

        assert not r.is_empty
        assert not r.is_cell
        assert r.is_range

        assert r.workbook is wb
        assert r.sheet is ws
        assert r.cell is None

        assert r.get_reference() == "'Report 1'!$B$2:$C$3"
        assert r.get_reference(absolute=False) == "'Report 1'!B2:C3"
        assert r.get_reference(use_sheet=False) == "$B$2:$C$3"
    
    def test_defined_name(self):
        wb = get_test_workbook()
        ws = wb['Report 1']
        
        defined_name = utils.get_defined_name(wb, None, "PROFIT_RANGE")
        ref = defined_name.attr_text

        _, (c1, r1, c2, r2) = openpyxl.utils.cell.range_to_tuple(ref)
        cells = tuple(ws.iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

        r = range.Range(cells, defined_name=defined_name)

        assert not r.is_empty
        assert not r.is_cell
        assert r.is_range

        assert r.workbook is wb
        assert r.sheet is ws
        assert r.cell is None

        assert r.get_reference() == "PROFIT_RANGE"
        assert r.get_reference(use_defined_name=False) == "'Report 1'!$A$1:$E$5"
        assert r.get_reference(use_defined_name=False, absolute=False) == "'Report 1'!A1:E5"
        assert r.get_reference(use_defined_name=False, use_sheet=False) == "$A$1:$E$5"
    
    def test_named_table(self):
        wb = get_test_workbook()
        ws = wb['Report 2']
        
        named_table = utils.get_named_table(ws, 'RangleTable')
        ref = utils.add_sheet_to_reference(ws, named_table.ref)

        _, (c1, r1, c2, r2) = openpyxl.utils.cell.range_to_tuple(ref)
        cells = tuple(ws.iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

        r = range.Range(cells, named_table=named_table)

        assert not r.is_empty
        assert not r.is_cell
        assert r.is_range

        assert r.workbook is wb
        assert r.sheet is ws
        assert r.cell is None

        assert r.get_reference() == "RangleTable"
        assert r.get_reference(use_named_table=False) == "'Report 2'!$B$10:$E$13"
        assert r.get_reference(use_named_table=False, absolute=False) == "'Report 2'!B10:E13"
        assert r.get_reference(use_named_table=False, use_sheet=False) == "$B$10:$E$13"
    
