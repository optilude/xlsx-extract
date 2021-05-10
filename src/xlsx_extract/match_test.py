import pytest
import datetime

from . import match

def test_construct_cell_match():

    sheet = match.SheetMatch(operator=match.Operator.EQUAL, value="a")

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
        value=match.Match(operator=match.Operator.NOT_EMPTY),
    )

    match.CellMatch(
        name="A",
        sheet=sheet,
        row_index_value=match.Match(operator=match.Operator.NOT_EMPTY),
        col_index_value=match.Match(operator=match.Operator.NOT_EMPTY),
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
            value=match.Match(operator=match.Operator.NOT_EMPTY),
        )
    
    # Row without column index and vice-versa
    with pytest.raises(AssertionError):
        match.CellMatch(
            name="A",
            sheet=sheet,
            row_index_value=match.Match(operator=match.Operator.NOT_EMPTY),
            # col_index_value=match.Match(operator=match.Operator.NOT_EMPTY),
        )
    
    with pytest.raises(AssertionError):
        match.CellMatch(
            name="A",
            sheet=sheet,
            # row_index_value=match.Match(operator=match.Operator.NOT_EMPTY),
            col_index_value=match.Match(operator=match.Operator.NOT_EMPTY),
        )
    
def test_construct_range_match():

    sheet = match.SheetMatch(operator=match.Operator.EQUAL, value="a")

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
    

def test_match_value_requires_regex_to_be_string():
    with pytest.raises(AssertionError):
        match.match_value(data="foo", operator=match.Operator.REGEX, comparator=1)

def test_match_value_requires_consistent_types():
    with pytest.raises(AssertionError):
        match.match_value(data="1", operator=match.Operator.EQUAL, comparator=1)

def test_match_empty():
    assert match.match_value(data="", operator=match.Operator.EMPTY, comparator=None) == ""
    assert match.match_value(data=None, operator=match.Operator.EMPTY, comparator=None) == ""  # yes, indeed

    assert match.match_value(data="a", operator=match.Operator.EMPTY, comparator=None) == None
    assert match.match_value(data=1, operator=match.Operator.EMPTY, comparator=None) == None

def test_match_not_empty():
    assert match.match_value(data="", operator=match.Operator.NOT_EMPTY, comparator=None) == None
    assert match.match_value(data=None, operator=match.Operator.NOT_EMPTY, comparator=None) == None

    assert match.match_value(data="a", operator=match.Operator.NOT_EMPTY, comparator=None) == "a"
    assert match.match_value(data=1, operator=match.Operator.NOT_EMPTY, comparator=None) == 1

def test_match_value_equal():
    assert match.match_value(data="foo", operator=match.Operator.EQUAL, comparator="foo") == "foo"
    assert match.match_value(data=1, operator=match.Operator.EQUAL, comparator=1) == 1
    assert match.match_value(data=1.2, operator=match.Operator.EQUAL, comparator=1.2) == 1.2
    assert match.match_value(data=True, operator=match.Operator.EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="bar", operator=match.Operator.EQUAL, comparator="foo") == None
    assert match.match_value(data=2, operator=match.Operator.EQUAL, comparator=1) == None
    assert match.match_value(data=2.2, operator=match.Operator.EQUAL, comparator=1.2) == None
    assert match.match_value(data=False, operator=match.Operator.EQUAL, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 3), operator=match.Operator.EQUAL, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 1), operator=match.Operator.EQUAL, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.Operator.EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

def test_match_value_not_equal():
    assert match.match_value(data="bar", operator=match.Operator.NOT_EQUAL, comparator="foo") == "bar"
    assert match.match_value(data=2, operator=match.Operator.NOT_EQUAL, comparator=1) == 2
    assert match.match_value(data=2.2, operator=match.Operator.NOT_EQUAL, comparator=1.2) == 2.2
    assert match.match_value(data=False, operator=match.Operator.NOT_EQUAL, comparator=True) == False
    assert match.match_value(data=datetime.date(2020, 1, 3), operator=match.Operator.NOT_EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 3)
    assert match.match_value(data=datetime.time(14, 1), operator=match.Operator.NOT_EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 1)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.Operator.NOT_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 1)
    
    assert match.match_value(data="foo", operator=match.Operator.NOT_EQUAL, comparator="foo") == None
    assert match.match_value(data=1, operator=match.Operator.NOT_EQUAL, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.Operator.NOT_EQUAL, comparator=1.2) == None
    assert match.match_value(data=True, operator=match.Operator.NOT_EQUAL, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.NOT_EQUAL, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.NOT_EQUAL, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.NOT_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

def test_match_value_greater_than():
    assert match.match_value(data="foo", operator=match.Operator.GREATER, comparator="boo") == "foo"
    assert match.match_value(data=2, operator=match.Operator.GREATER, comparator=1) == 2
    assert match.match_value(data=1.2, operator=match.Operator.GREATER, comparator=1.1) == 1.2
    # assert match.match_value(data=True, operator=match.Operator.GREATER, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, comparator=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER, comparator=datetime.time(13, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.Operator.GREATER, comparator="foo") == None
    assert match.match_value(data=1, operator=match.Operator.GREATER, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.Operator.GREATER, comparator=1.2) == None
    assert match.match_value(data=True, operator=match.Operator.GREATER, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

    assert match.match_value(data="foo", operator=match.Operator.GREATER, comparator="goo") == None
    assert match.match_value(data=1, operator=match.Operator.GREATER, comparator=2) == None
    assert match.match_value(data=1.2, operator=match.Operator.GREATER, comparator=1.3) == None
    # assert match.match_value(data=True, operator=match.Operator.GREATER, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, comparator=datetime.date(2020, 1, 3)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER, comparator=datetime.time(14, 1)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, comparator=datetime.datetime(2020, 1, 2, 14, 1)) == None

def test_match_value_greater_than_equal():
    assert match.match_value(data="foo", operator=match.Operator.GREATER_EQUAL, comparator="boo") == "foo"
    assert match.match_value(data=2, operator=match.Operator.GREATER_EQUAL, comparator=1) == 2
    assert match.match_value(data=1.2, operator=match.Operator.GREATER_EQUAL, comparator=1.1) == 1.2
    # assert match.match_value(data=True, operator=match.Operator.GREATER_EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, comparator=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, comparator=datetime.time(13, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.Operator.GREATER_EQUAL, comparator="foo") == "foo"
    assert match.match_value(data=1, operator=match.Operator.GREATER_EQUAL, comparator=1) == 1
    assert match.match_value(data=1.2, operator=match.Operator.GREATER_EQUAL, comparator=1.2) == 1.2
    assert match.match_value(data=True, operator=match.Operator.GREATER_EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.Operator.GREATER_EQUAL, comparator="goo") == None
    assert match.match_value(data=1, operator=match.Operator.GREATER_EQUAL, comparator=2) == None
    assert match.match_value(data=1.2, operator=match.Operator.GREATER_EQUAL, comparator=1.3) == None
    # assert match.match_value(data=True, operator=match.Operator.GREATER_EQUAL, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, comparator=datetime.date(2020, 1, 3)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, comparator=datetime.time(14, 1)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 1)) == None

def test_match_value_less_than():
    assert match.match_value(data="foo", operator=match.Operator.LESS, comparator="goo") == "foo"
    assert match.match_value(data=2, operator=match.Operator.LESS, comparator=3) == 2
    assert match.match_value(data=1.2, operator=match.Operator.LESS, comparator=1.3) == 1.2
    # assert match.match_value(data=True, operator=match.Operator.LESS, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, comparator=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, comparator=datetime.time(15, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, comparator=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.Operator.LESS, comparator="foo") == None
    assert match.match_value(data=1, operator=match.Operator.LESS, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.Operator.LESS, comparator=1.2) == None
    assert match.match_value(data=True, operator=match.Operator.LESS, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

    assert match.match_value(data="foo", operator=match.Operator.LESS, comparator="boo") == None
    assert match.match_value(data=2, operator=match.Operator.LESS, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.Operator.LESS, comparator=1.1) == None
    # assert match.match_value(data=True, operator=match.Operator.LESS, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, comparator=datetime.date(2020, 1, 1)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, comparator=datetime.time(13, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == None

def test_match_value_less_than_equal():
    assert match.match_value(data="foo", operator=match.Operator.LESS, comparator="goo") == "foo"
    assert match.match_value(data=2, operator=match.Operator.LESS, comparator=3) == 2
    assert match.match_value(data=1.2, operator=match.Operator.LESS, comparator=1.3) == 1.2
    # assert match.match_value(data=True, operator=match.Operator.LESS, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, comparator=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, comparator=datetime.time(15, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, comparator=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.Operator.LESS_EQUAL, comparator="foo") == "foo"
    assert match.match_value(data=1, operator=match.Operator.LESS_EQUAL, comparator=1) == 1
    assert match.match_value(data=1.2, operator=match.Operator.LESS_EQUAL, comparator=1.2) == 1.2
    assert match.match_value(data=True, operator=match.Operator.LESS_EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS_EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.LESS_EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.Operator.LESS, comparator="boo") == None
    assert match.match_value(data=2, operator=match.Operator.LESS, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.Operator.LESS, comparator=1.1) == None
    # assert match.match_value(data=True, operator=match.Operator.LESS, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, comparator=datetime.date(2020, 1, 1)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, comparator=datetime.time(13, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == None

def test_match_value_regex():
    assert match.match_value(data="foo bar", operator=match.Operator.REGEX, comparator="foo") == "foo bar"