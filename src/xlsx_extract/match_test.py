import pytest
import datetime

from . import match

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

def match_value(data, operator, value):
    return match.Comparator(operator=operator, value=value).match(data)

def test_match_value_requires_regex_to_be_string():
    with pytest.raises(AssertionError):
        match_value(data="foo", operator=match.Operator.REGEX, value=1)

def test_match_value_requires_consistent_types():
    with pytest.raises(AssertionError):
        match_value(data="1", operator=match.Operator.EQUAL, value=1)

def test_match_empty():
    assert match_value(data="", operator=match.Operator.EMPTY, value=None) == ""
    assert match_value(data=None, operator=match.Operator.EMPTY, value=None) == ""  # yes, indeed

    assert match_value(data="a", operator=match.Operator.EMPTY, value=None) == None
    assert match_value(data=1, operator=match.Operator.EMPTY, value=None) == None

def test_match_not_empty():
    assert match_value(data="", operator=match.Operator.NOT_EMPTY, value=None) == None
    assert match_value(data=None, operator=match.Operator.NOT_EMPTY, value=None) == None

    assert match_value(data="a", operator=match.Operator.NOT_EMPTY, value=None) == "a"
    assert match_value(data=1, operator=match.Operator.NOT_EMPTY, value=None) == 1

def test_match_value_equal():
    assert match_value(data="foo", operator=match.Operator.EQUAL, value="foo") == "foo"
    assert match_value(data=1, operator=match.Operator.EQUAL, value=1) == 1
    assert match_value(data=1.2, operator=match.Operator.EQUAL, value=1.2) == 1.2
    assert match_value(data=True, operator=match.Operator.EQUAL, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="bar", operator=match.Operator.EQUAL, value="foo") == None
    assert match_value(data=2, operator=match.Operator.EQUAL, value=1) == None
    assert match_value(data=2.2, operator=match.Operator.EQUAL, value=1.2) == None
    assert match_value(data=False, operator=match.Operator.EQUAL, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 3), operator=match.Operator.EQUAL, value=datetime.date(2020, 1, 2)) == None
    assert match_value(data=datetime.time(14, 1), operator=match.Operator.EQUAL, value=datetime.time(14, 0)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.Operator.EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

def test_match_value_not_equal():
    assert match_value(data="bar", operator=match.Operator.NOT_EQUAL, value="foo") == "bar"
    assert match_value(data=2, operator=match.Operator.NOT_EQUAL, value=1) == 2
    assert match_value(data=2.2, operator=match.Operator.NOT_EQUAL, value=1.2) == 2.2
    assert match_value(data=False, operator=match.Operator.NOT_EQUAL, value=True) == False
    assert match_value(data=datetime.date(2020, 1, 3), operator=match.Operator.NOT_EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 3)
    assert match_value(data=datetime.time(14, 1), operator=match.Operator.NOT_EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 1)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.Operator.NOT_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 1)
    
    assert match_value(data="foo", operator=match.Operator.NOT_EQUAL, value="foo") == None
    assert match_value(data=1, operator=match.Operator.NOT_EQUAL, value=1) == None
    assert match_value(data=1.2, operator=match.Operator.NOT_EQUAL, value=1.2) == None
    assert match_value(data=True, operator=match.Operator.NOT_EQUAL, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.NOT_EQUAL, value=datetime.date(2020, 1, 2)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.NOT_EQUAL, value=datetime.time(14, 0)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.NOT_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

def test_match_value_greater_than():
    assert match_value(data="foo", operator=match.Operator.GREATER, value="boo") == "foo"
    assert match_value(data=2, operator=match.Operator.GREATER, value=1) == 2
    assert match_value(data=1.2, operator=match.Operator.GREATER, value=1.1) == 1.2
    # assert match_value(data=True, operator=match.Operator.GREATER, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, value=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER, value=datetime.time(13, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, value=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="foo", operator=match.Operator.GREATER, value="foo") == None
    assert match_value(data=1, operator=match.Operator.GREATER, value=1) == None
    assert match_value(data=1.2, operator=match.Operator.GREATER, value=1.2) == None
    assert match_value(data=True, operator=match.Operator.GREATER, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, value=datetime.date(2020, 1, 2)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER, value=datetime.time(14, 0)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

    assert match_value(data="foo", operator=match.Operator.GREATER, value="goo") == None
    assert match_value(data=1, operator=match.Operator.GREATER, value=2) == None
    assert match_value(data=1.2, operator=match.Operator.GREATER, value=1.3) == None
    # assert match_value(data=True, operator=match.Operator.GREATER, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER, value=datetime.date(2020, 1, 3)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER, value=datetime.time(14, 1)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER, value=datetime.datetime(2020, 1, 2, 14, 1)) == None

def test_match_value_greater_than_equal():
    assert match_value(data="foo", operator=match.Operator.GREATER_EQUAL, value="boo") == "foo"
    assert match_value(data=2, operator=match.Operator.GREATER_EQUAL, value=1) == 2
    assert match_value(data=1.2, operator=match.Operator.GREATER_EQUAL, value=1.1) == 1.2
    # assert match_value(data=True, operator=match.Operator.GREATER_EQUAL, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, value=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.time(13, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="foo", operator=match.Operator.GREATER_EQUAL, value="foo") == "foo"
    assert match_value(data=1, operator=match.Operator.GREATER_EQUAL, value=1) == 1
    assert match_value(data=1.2, operator=match.Operator.GREATER_EQUAL, value=1.2) == 1.2
    assert match_value(data=True, operator=match.Operator.GREATER_EQUAL, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="foo", operator=match.Operator.GREATER_EQUAL, value="goo") == None
    assert match_value(data=1, operator=match.Operator.GREATER_EQUAL, value=2) == None
    assert match_value(data=1.2, operator=match.Operator.GREATER_EQUAL, value=1.3) == None
    # assert match_value(data=True, operator=match.Operator.GREATER_EQUAL, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.GREATER_EQUAL, value=datetime.date(2020, 1, 3)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.time(14, 1)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.GREATER_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 1)) == None

def test_match_value_less_than():
    assert match_value(data="foo", operator=match.Operator.LESS, value="goo") == "foo"
    assert match_value(data=2, operator=match.Operator.LESS, value=3) == 2
    assert match_value(data=1.2, operator=match.Operator.LESS, value=1.3) == 1.2
    # assert match_value(data=True, operator=match.Operator.LESS, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(15, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="foo", operator=match.Operator.LESS, value="foo") == None
    assert match_value(data=1, operator=match.Operator.LESS, value=1) == None
    assert match_value(data=1.2, operator=match.Operator.LESS, value=1.2) == None
    assert match_value(data=True, operator=match.Operator.LESS, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 2)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(14, 0)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 14, 0)) == None

    assert match_value(data="foo", operator=match.Operator.LESS, value="boo") == None
    assert match_value(data=2, operator=match.Operator.LESS, value=1) == None
    assert match_value(data=1.2, operator=match.Operator.LESS, value=1.1) == None
    # assert match_value(data=True, operator=match.Operator.LESS, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 1)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(13, 0)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 13, 0)) == None

def test_match_value_less_than_equal():
    assert match_value(data="foo", operator=match.Operator.LESS, value="goo") == "foo"
    assert match_value(data=2, operator=match.Operator.LESS, value=3) == 2
    assert match_value(data=1.2, operator=match.Operator.LESS, value=1.3) == 1.2
    # assert match_value(data=True, operator=match.Operator.LESS, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(15, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="foo", operator=match.Operator.LESS_EQUAL, value="foo") == "foo"
    assert match_value(data=1, operator=match.Operator.LESS_EQUAL, value=1) == 1
    assert match_value(data=1.2, operator=match.Operator.LESS_EQUAL, value=1.2) == 1.2
    assert match_value(data=True, operator=match.Operator.LESS_EQUAL, value=True) == True
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS_EQUAL, value=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.LESS_EQUAL, value=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS_EQUAL, value=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match_value(data="foo", operator=match.Operator.LESS, value="boo") == None
    assert match_value(data=2, operator=match.Operator.LESS, value=1) == None
    assert match_value(data=1.2, operator=match.Operator.LESS, value=1.1) == None
    # assert match_value(data=True, operator=match.Operator.LESS, value=True) == None
    assert match_value(data=datetime.date(2020, 1, 2), operator=match.Operator.LESS, value=datetime.date(2020, 1, 1)) == None
    assert match_value(data=datetime.time(14, 0), operator=match.Operator.LESS, value=datetime.time(13, 0)) == None
    assert match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.Operator.LESS, value=datetime.datetime(2020, 1, 2, 13, 0)) == None

def test_match_value_regex():
    assert match_value(data="foo bar", operator=match.Operator.REGEX, value="foo") == "foo bar"