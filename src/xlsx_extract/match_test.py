import pytest

from . import match

def test_construct_cell_match():
    c = match.CellMatch(operator=match.MatchOperator.EQUALS, value="a")
    assert c is not None

    c = match.CellMatch(operator=match.MatchOperator.EQUALS, value="a", min_row=1, min_col="A", max_row=2, max_col="B")
    assert c is not None

def test_construct_direct_cell_match():
    c = match.DirectCellMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.CELL,
        match_type=match.MatchType.DIRECT,
        cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.DirectCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )

def test_construct_separate_cell_match():
    c = match.SeparateCellMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.CELL,
        match_type=match.MatchType.SEPARATE,
        row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        )
    
    with pytest.raises(AssertionError):
        match.SeparateCellMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        )

def test_construct_named_range_match():
    c = match.NamedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.NAMED,
        name="A"
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.NamedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.NAMED,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.NamedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.NAMED,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.NamedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            name="A"
        )
    
def test_construct_table_range_match():
    c = match.TableRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.TABLE,
        name="A"
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.TableRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.TableRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            name="A"
        )
    
    with pytest.raises(AssertionError):
        match.TableRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.NAMED,
            name="A"
        )

def test_construct_direct_contiguous_range_match():
    c = match.DirectContiguousRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.CONTIGUOUS,
        start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a")
        )
    
    with pytest.raises(AssertionError):
        match.DirectContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a")
        )
    
    with pytest.raises(AssertionError):
        match.DirectContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a")
        )
    
def test_construct_separate_contiguous_range_match():
    c = match.SeparateContiguousRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.SEPARATE,
        range_size=match.RangeSize.CONTIGUOUS,
        start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.SeparateContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.CONTIGUOUS,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.SeparateContiguousRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )

def test_construct_direct_fixed_range_match():
    c = match.DirectFixedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.FIXED,
        start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        range_rows=5,
        range_cols=5
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.FIXED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.DirectFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.FIXED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.DirectFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            range_rows=5,
            range_cols=5
        )
    
def test_construct_separate_fixed_range_match():
    c = match.SeparateFixedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.SEPARATE,
        range_size=match.RangeSize.FIXED,
        start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        range_rows=5,
        range_cols=5
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.FIXED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.SeparateFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.FIXED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
            range_rows=5,
            range_cols=5
        )
    
    with pytest.raises(AssertionError):
        match.SeparateFixedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
            range_rows=5,
            range_cols=5
        )

def test_construct_direct_matched_range_match():
    c = match.DirectMatchedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.DIRECT,
        range_size=match.RangeSize.MATCHED,
        start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.DirectMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.MATCHED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.DirectMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.MATCHED,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )
    
    with pytest.raises(AssertionError):
        match.DirectMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.TABLE,
            start_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            end_cell_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b")
        )
    
def test_construct_separate_matched_range_match():
    c = match.SeparateMatchedRangeMatch(
        sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
        target=match.MatchTarget.RANGE,
        match_type=match.MatchType.SEPARATE,
        range_size=match.RangeSize.MATCHED,
        start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
        end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
    )

    assert c is not None

    with pytest.raises(AssertionError):
        match.SeparateMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.CELL,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.MATCHED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
            end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        )
    
    with pytest.raises(AssertionError):
        match.SeparateMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.DIRECT,
            range_size=match.RangeSize.MATCHED,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
            end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        )
    
    with pytest.raises(AssertionError):
        match.SeparateMatchedRangeMatch(
            sheet=match.SheetMatch(operator=match.MatchOperator.EQUALS, value="a"),
            target=match.MatchTarget.RANGE,
            match_type=match.MatchType.SEPARATE,
            range_size=match.RangeSize.TABLE,
            start_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            start_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
            end_cell_row_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="a"),
            end_cell_col_match=match.CellMatch(operator=match.MatchOperator.EQUALS, value="b"),
        )
