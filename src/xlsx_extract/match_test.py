import pytest
import datetime

from . import match

def test_construct_cell_match():
    c = match.CellMatch(operator=match.MatchOperator.EQUAL, value="a")
    assert c is not None

    c = match.CellMatch(operator=match.MatchOperator.EQUAL, value="a", min_row=1, min_col="A", max_row=2, max_col="B")
    assert c is not None

def test_construct_direct_cell_match():
    c = match.DirectCellMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.CELL,
        match_type=match.MatchType.DIRECT,
        cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.DirectCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )

def test_construct_separate_cell_match():
    c = match.SeparateCellMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.CELL,
        match_type=match.MatchType.SEPARATE,
        row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        )
    
    with pytest.raises(AssertionError):
        match.SeparateCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        )

def test_construct_named_range_match():
    c = match.NamedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.NAMED,
        name="A"
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.NamedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.NAMED,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.NamedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.NAMED,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.NamedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            name="A"
        )
    
def test_construct_table_range_match():
    c = match.TableRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.TABLE,
        name="A"
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.TableRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.TableRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.TableRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.NAMED,
            name="A"
        )

def test_construct_direct_contiguous_range_match():
    c = match.DirectContiguousRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.CONTIGUOUS,
        start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a")
        )
    
    with pytest.raises(AssertionError):
        match.DirectContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a")
        )
    
    with pytest.raises(AssertionError):
        match.DirectContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a")
        )
    
def test_construct_separate_contiguous_range_match():
    c = match.SeparateContiguousRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.SEPARATE,
        range_size=match.RangeSize.CONTIGUOUS,
        start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.SeparateContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.SeparateContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )

def test_construct_direct_fixed_range_match():
    c = match.DirectFixedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.FIXED,
        start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        range_rows=5,
        range_cols=5
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.FIXED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.DirectFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.FIXED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.DirectFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            range_rows=5,
            range_cols=5
        )
    
def test_construct_separate_fixed_range_match():
    c = match.SeparateFixedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.SEPARATE,
        range_size=match.RangeSize.FIXED,
        start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        range_rows=5,
        range_cols=5
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.FIXED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.SeparateFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.FIXED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.SeparateFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
            range_rows=5,
            range_cols=5
        )

def test_construct_direct_matched_range_match():
    c = match.DirectMatchedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.MATCHED,
        start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.MATCHED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.DirectMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.MATCHED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.DirectMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b")
        )
    
def test_construct_separate_matched_range_match():
    c = match.SeparateMatchedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.SEPARATE,
        range_size=match.RangeSize.MATCHED,
        start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
        end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.MATCHED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
            end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        )
    
    with pytest.raises(AssertionError):
        match.SeparateMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.MATCHED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
            end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        )
    
    with pytest.raises(AssertionError):
        match.SeparateMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUAL, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
            end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="a"),
            end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUAL, value="b"),
        )

def test_match_value_does_not_allow_reference():
    with pytest.raises(AssertionError):
        match.match_value(data="foo", operator=match.MatchOperator.REFERENCE, comparator="foo")

def test_match_value_requires_regex_to_be_string():
    with pytest.raises(AssertionError):
        match.match_value(data="foo", operator=match.MatchOperator.REGEX, comparator=1)

def test_match_value_requires_consistent_types():
    with pytest.raises(AssertionError):
        match.match_value(data="1", operator=match.MatchOperator.EQUAL, comparator=1)

def test_match_empty():
    assert match.match_value(data="", operator=match.MatchOperator.EMPTY, comparator=None) == ""
    assert match.match_value(data=None, operator=match.MatchOperator.EMPTY, comparator=None) == ""  # yes, indeed

    assert match.match_value(data="a", operator=match.MatchOperator.EMPTY, comparator=None) == None
    assert match.match_value(data=1, operator=match.MatchOperator.EMPTY, comparator=None) == None

def test_match_not_empty():
    assert match.match_value(data="", operator=match.MatchOperator.NOT_EMPTY, comparator=None) == None
    assert match.match_value(data=None, operator=match.MatchOperator.NOT_EMPTY, comparator=None) == None

    assert match.match_value(data="a", operator=match.MatchOperator.NOT_EMPTY, comparator=None) == "a"
    assert match.match_value(data=1, operator=match.MatchOperator.NOT_EMPTY, comparator=None) == 1

def test_match_value_equal():
    assert match.match_value(data="foo", operator=match.MatchOperator.EQUAL, comparator="foo") == "foo"
    assert match.match_value(data=1, operator=match.MatchOperator.EQUAL, comparator=1) == 1
    assert match.match_value(data=1.2, operator=match.MatchOperator.EQUAL, comparator=1.2) == 1.2
    assert match.match_value(data=True, operator=match.MatchOperator.EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="bar", operator=match.MatchOperator.EQUAL, comparator="foo") == None
    assert match.match_value(data=2, operator=match.MatchOperator.EQUAL, comparator=1) == None
    assert match.match_value(data=2.2, operator=match.MatchOperator.EQUAL, comparator=1.2) == None
    assert match.match_value(data=False, operator=match.MatchOperator.EQUAL, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 3), operator=match.MatchOperator.EQUAL, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 1), operator=match.MatchOperator.EQUAL, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.MatchOperator.EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

def test_match_value_not_equal():
    assert match.match_value(data="bar", operator=match.MatchOperator.NOT_EQUAL, comparator="foo") == "bar"
    assert match.match_value(data=2, operator=match.MatchOperator.NOT_EQUAL, comparator=1) == 2
    assert match.match_value(data=2.2, operator=match.MatchOperator.NOT_EQUAL, comparator=1.2) == 2.2
    assert match.match_value(data=False, operator=match.MatchOperator.NOT_EQUAL, comparator=True) == False
    assert match.match_value(data=datetime.date(2020, 1, 3), operator=match.MatchOperator.NOT_EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 3)
    assert match.match_value(data=datetime.time(14, 1), operator=match.MatchOperator.NOT_EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 1)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 1), operator=match.MatchOperator.NOT_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 1)
    
    assert match.match_value(data="foo", operator=match.MatchOperator.NOT_EQUAL, comparator="foo") == None
    assert match.match_value(data=1, operator=match.MatchOperator.NOT_EQUAL, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.NOT_EQUAL, comparator=1.2) == None
    assert match.match_value(data=True, operator=match.MatchOperator.NOT_EQUAL, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.NOT_EQUAL, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.NOT_EQUAL, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.NOT_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

def test_match_value_greater_than():
    assert match.match_value(data="foo", operator=match.MatchOperator.GREATER, comparator="boo") == "foo"
    assert match.match_value(data=2, operator=match.MatchOperator.GREATER, comparator=1) == 2
    assert match.match_value(data=1.2, operator=match.MatchOperator.GREATER, comparator=1.1) == 1.2
    # assert match.match_value(data=True, operator=match.MatchOperator.GREATER, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.GREATER, comparator=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.GREATER, comparator=datetime.time(13, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.GREATER, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.MatchOperator.GREATER, comparator="foo") == None
    assert match.match_value(data=1, operator=match.MatchOperator.GREATER, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.GREATER, comparator=1.2) == None
    assert match.match_value(data=True, operator=match.MatchOperator.GREATER, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.GREATER, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.GREATER, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.GREATER, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

    assert match.match_value(data="foo", operator=match.MatchOperator.GREATER, comparator="goo") == None
    assert match.match_value(data=1, operator=match.MatchOperator.GREATER, comparator=2) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.GREATER, comparator=1.3) == None
    # assert match.match_value(data=True, operator=match.MatchOperator.GREATER, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.GREATER, comparator=datetime.date(2020, 1, 3)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.GREATER, comparator=datetime.time(14, 1)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.GREATER, comparator=datetime.datetime(2020, 1, 2, 14, 1)) == None

def test_match_value_greater_than_equal():
    assert match.match_value(data="foo", operator=match.MatchOperator.GREATER_EQUAL, comparator="boo") == "foo"
    assert match.match_value(data=2, operator=match.MatchOperator.GREATER_EQUAL, comparator=1) == 2
    assert match.match_value(data=1.2, operator=match.MatchOperator.GREATER_EQUAL, comparator=1.1) == 1.2
    # assert match.match_value(data=True, operator=match.MatchOperator.GREATER_EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.date(2020, 1, 1)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.time(13, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.MatchOperator.GREATER_EQUAL, comparator="foo") == "foo"
    assert match.match_value(data=1, operator=match.MatchOperator.GREATER_EQUAL, comparator=1) == 1
    assert match.match_value(data=1.2, operator=match.MatchOperator.GREATER_EQUAL, comparator=1.2) == 1.2
    assert match.match_value(data=True, operator=match.MatchOperator.GREATER_EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.MatchOperator.GREATER_EQUAL, comparator="goo") == None
    assert match.match_value(data=1, operator=match.MatchOperator.GREATER_EQUAL, comparator=2) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.GREATER_EQUAL, comparator=1.3) == None
    # assert match.match_value(data=True, operator=match.MatchOperator.GREATER_EQUAL, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.date(2020, 1, 3)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.time(14, 1)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.GREATER_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 1)) == None

def test_match_value_less_than():
    assert match.match_value(data="foo", operator=match.MatchOperator.LESS, comparator="goo") == "foo"
    assert match.match_value(data=2, operator=match.MatchOperator.LESS, comparator=3) == 2
    assert match.match_value(data=1.2, operator=match.MatchOperator.LESS, comparator=1.3) == 1.2
    # assert match.match_value(data=True, operator=match.MatchOperator.LESS, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.LESS, comparator=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.LESS, comparator=datetime.time(15, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.LESS, comparator=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.MatchOperator.LESS, comparator="foo") == None
    assert match.match_value(data=1, operator=match.MatchOperator.LESS, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.LESS, comparator=1.2) == None
    assert match.match_value(data=True, operator=match.MatchOperator.LESS, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.LESS, comparator=datetime.date(2020, 1, 2)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.LESS, comparator=datetime.time(14, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.LESS, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == None

    assert match.match_value(data="foo", operator=match.MatchOperator.LESS, comparator="boo") == None
    assert match.match_value(data=2, operator=match.MatchOperator.LESS, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.LESS, comparator=1.1) == None
    # assert match.match_value(data=True, operator=match.MatchOperator.LESS, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.LESS, comparator=datetime.date(2020, 1, 1)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.LESS, comparator=datetime.time(13, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.LESS, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == None

def test_match_value_less_than_equal():
    assert match.match_value(data="foo", operator=match.MatchOperator.LESS, comparator="goo") == "foo"
    assert match.match_value(data=2, operator=match.MatchOperator.LESS, comparator=3) == 2
    assert match.match_value(data=1.2, operator=match.MatchOperator.LESS, comparator=1.3) == 1.2
    # assert match.match_value(data=True, operator=match.MatchOperator.LESS, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.LESS, comparator=datetime.date(2020, 1, 3)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.LESS, comparator=datetime.time(15, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.LESS, comparator=datetime.datetime(2020, 1, 2, 15, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.MatchOperator.LESS_EQUAL, comparator="foo") == "foo"
    assert match.match_value(data=1, operator=match.MatchOperator.LESS_EQUAL, comparator=1) == 1
    assert match.match_value(data=1.2, operator=match.MatchOperator.LESS_EQUAL, comparator=1.2) == 1.2
    assert match.match_value(data=True, operator=match.MatchOperator.LESS_EQUAL, comparator=True) == True
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.LESS_EQUAL, comparator=datetime.date(2020, 1, 2)) == datetime.date(2020, 1, 2)
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.LESS_EQUAL, comparator=datetime.time(14, 0)) == datetime.time(14, 0)
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.LESS_EQUAL, comparator=datetime.datetime(2020, 1, 2, 14, 0)) == datetime.datetime(2020, 1, 2, 14, 0)

    assert match.match_value(data="foo", operator=match.MatchOperator.LESS, comparator="boo") == None
    assert match.match_value(data=2, operator=match.MatchOperator.LESS, comparator=1) == None
    assert match.match_value(data=1.2, operator=match.MatchOperator.LESS, comparator=1.1) == None
    # assert match.match_value(data=True, operator=match.MatchOperator.LESS, comparator=True) == None
    assert match.match_value(data=datetime.date(2020, 1, 2), operator=match.MatchOperator.LESS, comparator=datetime.date(2020, 1, 1)) == None
    assert match.match_value(data=datetime.time(14, 0), operator=match.MatchOperator.LESS, comparator=datetime.time(13, 0)) == None
    assert match.match_value(data=datetime.datetime(2020, 1, 2, 14, 0), operator=match.MatchOperator.LESS, comparator=datetime.datetime(2020, 1, 2, 13, 0)) == None

def test_match_value_regex():
    assert match.match_value(data="foo bar", operator=match.MatchOperator.REGEX, comparator="foo") == "foo bar"