import os.path
import datetime
import pytest
import openpyxl

from . import match

def get_test_workbook():
    filename = os.path.join(os.path.dirname(__file__), 'test_data', 'source.xlsx')
    return openpyxl.load_workbook(filename, data_only=True)

def test_construct_cell_match():

    sheet = match.Comparator(operator=match.Operator.EQUAL, value="a")

    match.CellMatch(
        name="A",
        sheet=sheet,
        min_row=1,
        max_row=5,
        min_col=1,
        max_col=5,
        reference="A3",
        row_offset=1,
        col_offset=-1
    )

    match.CellMatch(
        name="A",
        sheet=sheet,
        value=match.Comparator(operator=match.Operator.NOT_EMPTY),
    )

    # No match criteria
    with pytest.raises(AssertionError):
        match.CellMatch(
            name="F",
            sheet=sheet,
        )
    
    # Too many match criteria (reference + value)
    with pytest.raises(AssertionError):
        match.CellMatch(
            name="F",
            sheet=sheet,
            reference="A3",
            value=match.Comparator(operator=match.Operator.NOT_EMPTY),
        )
    
def test_construct_range_match():

    sheet = match.Comparator(operator=match.Operator.EQUAL, value="a")

    match.RangeMatch(
        name="A",
        sheet=sheet,
        min_row=1,
        max_row=5,
        min_col=1,
        max_col=5,
        reference="Table1",
    )

    r = match.RangeMatch(
        name="A",
        sheet=sheet,
        start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell")
    )
    assert r.contiguous

    r = match.RangeMatch(
        name="A",
        sheet=sheet,
        start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell"),
        rows=10,
        cols=5,
    )
    assert not r.contiguous

    r = match.RangeMatch(
        name="A",
        sheet=sheet,
        start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell"),
        end_cell=match.CellMatch(name="D", sheet=sheet, reference="B:12"),
    )
    assert not r.contiguous

    # Need start cell or reference
    with pytest.raises(AssertionError):
        match.RangeMatch(
            name="A",
            sheet=sheet,
        )
    
    # ... but not both
    with pytest.raises(AssertionError):
        match.RangeMatch(
            name="A",
            sheet=sheet,
            reference="Table2",
            start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell"),
        )
    
    # Cannot have both end cell and fixed size
    with pytest.raises(AssertionError):
        match.RangeMatch(
            name="A",
            sheet=sheet,
            start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell"),
            end_cell=match.CellMatch(name="D", sheet=sheet, reference="B:12"),
            rows=5,
            cols=5
        )
    
    # Must have both rows and cols
    with pytest.raises(AssertionError):
        match.RangeMatch(
            name="A",
            sheet=sheet,
            start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell"),
            # rows=5,
            cols=5
        )
    
    # Must have both rows and cols
    with pytest.raises(AssertionError):
        match.RangeMatch(
            name="A",
            sheet=sheet,
            start_cell=match.CellMatch(name="C", sheet=sheet, reference="ACell"),
            rows=5,
            # cols=5
        )

class TestMatchValue:

    def mv(self, data, operator, value):
        return match.Comparator(operator=operator, value=value).match(data)

    def test_match_value_requires_regex_to_be_string(self):
        with pytest.raises(AssertionError):
            self.mv(data="foo", operator=match.Operator.REGEX, value=1)

    def test_match_value_requires_consistent_types(self):
        assert self.mv(data="1", operator=match.Operator.EQUAL, value=1) == None

    def test_match_value_empty(self):
        assert self.mv(data="", operator=match.Operator.EMPTY, value=None) == ""
        assert self.mv(data=None, operator=match.Operator.EMPTY, value=None) == ""  # yes, indeed

        assert self.mv(data="a", operator=match.Operator.EMPTY, value=None) == None
        assert self.mv(data=1, operator=match.Operator.EMPTY, value=None) == None

    def test_match_value_not_empty(self):
        assert self.mv(data="", operator=match.Operator.NOT_EMPTY, value=None) == None
        assert self.mv(data=None, operator=match.Operator.NOT_EMPTY, value=None) == None

        assert self.mv(data="a", operator=match.Operator.NOT_EMPTY, value=None) == "a"
        assert self.mv(data=1, operator=match.Operator.NOT_EMPTY, value=None) == 1

    def test_match_value_equal(self):
        assert self.mv(data="foo", operator=match.Operator.EQUAL, value="foo") == "foo"
        assert self.mv(data=1, operator=match.Operator.EQUAL, value=1) == 1
        assert self.mv(data=1.2, operator=match.Operator.EQUAL, value=1.2) == 1.2
        assert self.mv(data=True, operator=match.Operator.EQUAL, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.datetime(2020, 1, 2), operator=match.Operator.EQUAL, value=datetime.date(2020, 1, 2)) == datetime.datetime(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="bar", operator=match.Operator.EQUAL, value="foo") == None
        assert self.mv(data=2, operator=match.Operator.EQUAL, value=1) == None
        assert self.mv(data=2.2, operator=match.Operator.EQUAL, value=1.2) == None
        assert self.mv(data=False, operator=match.Operator.EQUAL, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 3), operator=match.Operator.EQUAL, value=datetime.date(2020, 1, 2)) == None
        assert self.mv(data=datetime.time(14, 1), operator=match.Operator.EQUAL, value=datetime.time(14, 0)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.Operator.EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

    def test_match_value_not_equal(self):
        assert self.mv(data="bar", operator=match.Operator.NOT_EQUAL, value="foo") == "bar"
        assert self.mv(data=2, operator=match.Operator.NOT_EQUAL, value=1) == 2
        assert self.mv(data=2.2, operator=match.Operator.NOT_EQUAL, value=1.2) == 2.2
        assert self.mv(data=False, operator=match.Operator.NOT_EQUAL, value=True) == False
        assert self.mv(data=datetime.date(2020, 1, 3), operator=match.Operator.NOT_EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 3)
        assert self.mv(data=datetime.time(14, 1), operator=match.Operator.NOT_EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 1)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.Operator.NOT_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 1)
        
        assert self.mv(data="foo", operator=match.Operator.NOT_EQUAL, value="foo") == None
        assert self.mv(data=1, operator=match.Operator.NOT_EQUAL, value=1) == None
        assert self.mv(data=1.2, operator=match.Operator.NOT_EQUAL, value=1.2) == None
        assert self.mv(data=True, operator=match.Operator.NOT_EQUAL, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.NOT_EQUAL, value=datetime.date(2020, 1, 2)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.NOT_EQUAL, value=datetime.time(14, 0)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.NOT_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

    def test_match_value_greater_than(self):
        assert self.mv(data="foo", operator=match.Operator.GREATER, value="boo") == "foo"
        assert self.mv(data=2, operator=match.Operator.GREATER, value=1) == 2
        assert self.mv(data=1.2, operator=match.Operator.GREATER, value=1.1) == 1.2
        # assert self.mv(data=True, operator=match.Operator.GREATER, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, value=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.GREATER, value=datetime.time(13, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, value=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="foo", operator=match.Operator.GREATER, value="foo") == None
        assert self.mv(data=1, operator=match.Operator.GREATER, value=1) == None
        assert self.mv(data=1.2, operator=match.Operator.GREATER, value=1.2) == None
        assert self.mv(data=True, operator=match.Operator.GREATER, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, value=datetime.date(2020, 1, 2)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.GREATER, value=datetime.time(14, 0)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

        assert self.mv(data="foo", operator=match.Operator.GREATER, value="goo") == None
        assert self.mv(data=1, operator=match.Operator.GREATER, value=2) == None
        assert self.mv(data=1.2, operator=match.Operator.GREATER, value=1.3) == None
        # assert self.mv(data=True, operator=match.Operator.GREATER, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, value=datetime.date(2020, 1, 3)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.GREATER, value=datetime.time(14, 1)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, value=datetime.datetime(2020, 1, 2, 14, 1)) == None

    def test_match_value_greater_than_equal(self):
        assert self.mv(data="foo", operator=match.Operator.GREATER_EQUAL, value="boo") == "foo"
        assert self.mv(data=2, operator=match.Operator.GREATER_EQUAL, value=1) == 2
        assert self.mv(data=1.2, operator=match.Operator.GREATER_EQUAL, value=1.1) == 1.2
        # assert self.mv(data=True, operator=match.Operator.GREATER_EQUAL, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, value=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.time(13, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="foo", operator=match.Operator.GREATER_EQUAL, value="foo") == "foo"
        assert self.mv(data=1, operator=match.Operator.GREATER_EQUAL, value=1) == 1
        assert self.mv(data=1.2, operator=match.Operator.GREATER_EQUAL, value=1.2) == 1.2
        assert self.mv(data=True, operator=match.Operator.GREATER_EQUAL, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="foo", operator=match.Operator.GREATER_EQUAL, value="goo") == None
        assert self.mv(data=1, operator=match.Operator.GREATER_EQUAL, value=2) == None
        assert self.mv(data=1.2, operator=match.Operator.GREATER_EQUAL, value=1.3) == None
        # assert self.mv(data=True, operator=match.Operator.GREATER_EQUAL, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, value=datetime.date(2020, 1, 3)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.time(14, 1)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 1)) == None

    def test_match_value_less_than(self):
        assert self.mv(data="foo", operator=match.Operator.LESS, value="goo") == "foo"
        assert self.mv(data=2, operator=match.Operator.LESS, value=3) == 2
        assert self.mv(data=1.2, operator=match.Operator.LESS, value=1.3) == 1.2
        # assert self.mv(data=True, operator=match.Operator.LESS, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(15, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="foo", operator=match.Operator.LESS, value="foo") == None
        assert self.mv(data=1, operator=match.Operator.LESS, value=1) == None
        assert self.mv(data=1.2, operator=match.Operator.LESS, value=1.2) == None
        assert self.mv(data=True, operator=match.Operator.LESS, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 2)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(14, 0)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

        assert self.mv(data="foo", operator=match.Operator.LESS, value="boo") == None
        assert self.mv(data=2, operator=match.Operator.LESS, value=1) == None
        assert self.mv(data=1.2, operator=match.Operator.LESS, value=1.1) == None
        # assert self.mv(data=True, operator=match.Operator.LESS, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 1)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(13, 0)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 13, 0)) == None

    def test_match_value_less_than_equal(self):
        assert self.mv(data="foo", operator=match.Operator.LESS, value="goo") == "foo"
        assert self.mv(data=2, operator=match.Operator.LESS, value=3) == 2
        assert self.mv(data=1.2, operator=match.Operator.LESS, value=1.3) == 1.2
        # assert self.mv(data=True, operator=match.Operator.LESS, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(15, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="foo", operator=match.Operator.LESS_EQUAL, value="foo") == "foo"
        assert self.mv(data=1, operator=match.Operator.LESS_EQUAL, value=1) == 1
        assert self.mv(data=1.2, operator=match.Operator.LESS_EQUAL, value=1.2) == 1.2
        assert self.mv(data=True, operator=match.Operator.LESS_EQUAL, value=True) == True
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS_EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.LESS_EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 0)
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

        assert self.mv(data="foo", operator=match.Operator.LESS, value="boo") == None
        assert self.mv(data=2, operator=match.Operator.LESS, value=1) == None
        assert self.mv(data=1.2, operator=match.Operator.LESS, value=1.1) == None
        # assert self.mv(data=True, operator=match.Operator.LESS, value=True) == None
        assert self.mv(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 1)) == None
        assert self.mv(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(13, 0)) == None
        assert self.mv(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 13, 0)) == None

    def test_match_value_regex(self):
        assert self.mv(data="foo bar", operator=match.Operator.REGEX, value="foo") == "foo bar"

class TestSheetMatch:

    def test_cell_match_sheet_match_notfound(self):
        wb = get_test_workbook()
        
        m = match.CellMatch(name="Test", sheet=match.Comparator(match.Operator.EQUAL, "foobar"), reference="A1")

        ws, s = m.get_sheet(wb)

        assert s is None
        assert ws is None

    def test_cell_match_sheet_match_equals(self):
        wb = get_test_workbook()
        m = match.CellMatch(name="Test", sheet=match.Comparator(match.Operator.EQUAL, "Report 1"), reference="A1")

        ws, s = m.get_sheet(wb)

        assert s == "Report 1"
        assert ws.title == "Report 1"
        assert ws.parent is wb

    def test_cell_match_sheet_match_regex(self):
        wb = get_test_workbook()
        m = match.CellMatch(name="Test", sheet=match.Comparator(match.Operator.REGEX, "Report (.+)"), reference="A1")

        ws, s = m.get_sheet(wb)

        assert s == "1"
        assert ws.title == "Report 1"
        assert ws.parent is wb

class TestCellMatch:

    def test_find_by_reference_cell(self):
        wb = get_test_workbook()
        m = match.CellMatch(name="Test", reference="'Report 1'!B3")
        
        v, s = m.match(wb)

        assert v.value == "Date"
        assert s is None

    def test_find_by_reference_cell_with_different_sheet(self):
        wb = get_test_workbook()
        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 2"),
            reference="'Report 1'!B3"
        )
        
        v, s = m.match(wb)

        assert v.value == "Date"
        assert s is None

    def test_find_by_reference_cell_with_sheet(self):
        wb = get_test_workbook()
        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            reference="B3"
        )
        
        v, s = m.match(wb)

        assert v.value == "Date"
        assert s is None

    def test_find_by_reference_range(self):
        wb = get_test_workbook()
        m = match.CellMatch(name="Test", reference="'Report 1'!A3:B3")
        
        v, s = m.match(wb)

        assert v is None
        assert s is None

    def test_find_by_reference_named(self):
        wb = get_test_workbook()
        m = match.CellMatch(name="Test", reference="DATE_CELL")
        v, s = m.match(wb)

        assert v.value == datetime.datetime(2021, 5, 1, 0, 0)
        assert s is None
    
    def test_find_by_reference_named_range(self):
        wb = get_test_workbook()
        m = match.CellMatch(name="Test", reference="PROFIT_RANGE")
        
        v, s = m.match(wb)

        assert v is None
        assert s is None

    def test_find_by_reference_table(self):
        wb = get_test_workbook()
        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 2"),
            reference="RangleTable"
        )
        
        v, s = m.match(wb)

        assert v is None
        assert s is None

    def test_find_by_value_string(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, "Date")
        )
        v, s = m.match(wb)
        assert v.coordinate == 'B3'
        assert v.value == "Date"
        assert s == "Date"

        m = match.CellMatch(name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, "notfound")
        )
        v, s = m.match(wb)
        assert v is None
        assert s is None

    def test_find_by_value_regex(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.REGEX, "^Da(.+)")
        )
        v, s = m.match(wb)
        assert v.coordinate == 'B3'
        assert v.value == "Date"
        assert s == "te"

        m = match.CellMatch(name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.REGEX, "^Da$")
        )
        v, s = m.match(wb)
        assert v is None
        assert s is None

    def test_find_by_value_empty(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EMPTY)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'A1'
        assert v.value is None
        assert s == ""

    def test_find_by_value_empty_bounded(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EMPTY),
            min_row=3,
            min_col=2
        )
        v, s = m.match(wb)
        assert v.coordinate == 'D3'
        assert v.value is None
        assert s == ""

    def test_find_by_value_not_empty(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.NOT_EMPTY)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'B3'
        assert v.value == "Date"
        assert s == "Date"

    def test_find_by_value_not_empty_bounded(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.NOT_EMPTY),
            min_row=4,
            min_col=2
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C5'
        assert v.value == "Jan"
        assert s == "Jan"

    def test_find_by_value_numeric(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.GREATER, 6)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'E6'
        assert v.value == 11
        assert s == 11

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.GREATER_EQUAL, 6)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'D6'
        assert v.value == 6
        assert s == 6

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, 4.6)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'F6'
        assert v.value == 4.6
        assert s == 4.6

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS, 1.5)
        )
        v, s = m.match(wb)
        assert v is None
        assert s == None

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS, 2)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C6'
        assert v.value == 1.5
        assert s == 1.5

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS_EQUAL, 1.5)
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C6'
        assert v.value == 1.5
        assert s == 1.5

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.NOT_EQUAL, 1.5),
            min_row=6,
            min_col=3,
            max_row=9,
            max_col=6,
        )
        v, s = m.match(wb)
        assert v.coordinate == 'D6'
        assert v.value == 6
        assert s == 6

    def test_find_by_value_datetime(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, datetime.datetime(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == datetime.datetime(2021, 5, 1)

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.GREATER, datetime.datetime(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v is None
        assert s is None

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.GREATER_EQUAL, datetime.datetime(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == datetime.datetime(2021, 5, 1)

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS, datetime.datetime(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v is None
        assert s is None

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS_EQUAL, datetime.datetime(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == datetime.datetime(2021, 5, 1)
    
    def test_find_by_value_datet(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, datetime.date(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == datetime.datetime(2021, 5, 1)

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.GREATER, datetime.date(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v is None
        assert s is None

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.GREATER_EQUAL, datetime.date(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == datetime.datetime(2021, 5, 1)

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS, datetime.date(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v is None
        assert s is None

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.LESS_EQUAL, datetime.date(2021, 5, 1))
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == datetime.datetime(2021, 5, 1)

    def test_offset(self):
        wb = get_test_workbook()

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, "Date"),
            col_offset=1,
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C3'
        assert v.value == datetime.datetime(2021, 5, 1)
        assert s == "Date"

        m = match.CellMatch(
            name="Test",
            sheet=match.Comparator(match.Operator.EQUAL, "Report 1"),
            value=match.Comparator(match.Operator.EQUAL, "Date"),
            col_offset=1,
            row_offset=2
        )
        v, s = m.match(wb)
        assert v.coordinate == 'C5'
        assert v.value == "Jan"
        assert s == "Date"

        