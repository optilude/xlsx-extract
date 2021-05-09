import re

from enum import Enum
from typing import Union
from datetime import datetime, date, time
from dataclasses import dataclass

class MatchTarget(Enum):

    CELL = "cell"   # Find a single cell matching the parameter(s)
    RANGE = "range" # Find a range matching the parameter(s)

class MatchType(Enum):

    DIRECT = "direct"            # Target matches row and column of parameter
    SEPARATE = "separate"        # Separate parameters for row and column

class MatchOperator(Enum):

    EQUAL = "="
    NOT_EQUAL = "!="
    GREATER = ">"
    GREATER_EQUAL = ">="
    LESS = "<"
    LESS_EQUAL = "<="

    EMPTY = "is empty"
    NOT_EMPTY = "is not empty"

    REGEX = "regex"

    REFERENCE = "reference"

class RangeSize(Enum):

    NAMED = "named"             # Find a named range
    TABLE = "table"             # Find a named data table
    FIXED = "fixed"             # Specify a number of rows and column
    MATCHED = "match cell"      # Match the end of the range using match parameters
    CONTIGUOUS = "contiguous"   # Range extends across the header row and down the first column until a blank cell is found

@dataclass
class SheetMatch:
    """Parameters to find a sheet
    """

    operator : MatchOperator
    value : str
    
@dataclass
class CellMatch:
    """Parameters to find a single cell
    """

    operator : MatchOperator
    value : str

    min_row : int = None
    min_col : str = None
    max_row : int = None
    max_col : str = None

@dataclass
class TargetMatch:
    """Find a target cell or range
    """

    sheet : SheetMatch      # which sheet are we looking in
    target : MatchTarget    # looking for a cell or a range
    match_type : MatchType  # one parameter (cell) or two (row, col separate)

@dataclass
class DirectCellMatch(TargetMatch):
    """Target a single cell directly
    """
    
    cell_match : CellMatch

    def __post_init__(self):
        assert self.match_type == MatchType.DIRECT
        assert self.target == MatchTarget.CELL

@dataclass
class SeparateCellMatch(TargetMatch):
    """Target a single cell with separate row/column matches
    """

    row_match : CellMatch
    col_match : CellMatch

    def __post_init__(self):
        assert self.match_type == MatchType.SEPARATE
        assert self.target == MatchTarget.CELL

@dataclass
class RangeMatch(TargetMatch):

    range_size : RangeSize

    def __post_init__(self):
        assert self.target == MatchTarget.RANGE

@dataclass
class NamedRangeMatch(RangeMatch):
    """Target a named range
    """

    name : str

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.NAMED
        assert self.match_type == MatchType.DIRECT

@dataclass
class TableRangeMatch(RangeMatch):
    """Target a named table
    """

    name : str

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.TABLE
        assert self.match_type == MatchType.DIRECT

@dataclass
class DirectContiguousRangeMatch(RangeMatch):
    """Target a contiguous range of cells from a directly targeted start cell
    """

    start_cell_match : CellMatch

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.CONTIGUOUS
        assert self.match_type == MatchType.DIRECT

@dataclass
class SeparateContiguousRangeMatch(RangeMatch):
    """Target a contiguous range of cells from a separately targeted start cell
    """

    start_cell_row_match : CellMatch
    start_cell_col_match : CellMatch

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.CONTIGUOUS
        assert self.match_type == MatchType.SEPARATE

@dataclass
class DirectFixedRangeMatch(RangeMatch):
    """Target a range of fixed size from a directly targeted start cell
    """

    start_cell_match : CellMatch

    range_rows : int
    range_cols : int

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.FIXED
        assert self.match_type == MatchType.DIRECT

@dataclass
class SeparateFixedRangeMatch(RangeMatch):
    """Target a range of fixed size from a separately targeted start cell
    """

    start_cell_row_match : CellMatch
    start_cell_col_match : CellMatch
    
    range_rows : int
    range_cols : int

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.FIXED
        assert self.match_type == MatchType.SEPARATE

@dataclass
class DirectMatchedRangeMatch(RangeMatch):
    """Target a range between directly targeted start and end cells
    """

    start_cell_match : CellMatch
    end_cell_match : CellMatch

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.MATCHED
        assert self.match_type == MatchType.DIRECT

@dataclass
class SeparateMatchedRangeMatch(RangeMatch):
    """Target a range between separately targeted start and end cells
    """

    start_cell_row_match : CellMatch
    start_cell_col_match : CellMatch

    end_cell_row_match : CellMatch
    end_cell_col_match : CellMatch

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.MATCHED
        assert self.match_type == MatchType.SEPARATE

def match_value(
    data : Union[str, int, float, bool, date, time, datetime],
    operator : MatchOperator,
    comparator : Union[str, int, float, bool, date, time, datetime]
) -> Union[str, int, float, bool, date, time, datetime]:
    """Use the `operator` to compare `data` with `comparator`.

    Return value is `None` if not matched, or the matched item.
    For regex matches with match groups, the content of the first
    match group is returned (as a string).
    """

    assert operator != MatchOperator.REFERENCE, "Reference match type should not be used for value comparison"

    if operator == MatchOperator.REGEX:
        assert type(comparator) is str, "Regular expression must be a string"
        comparator = re.compile(comparator)
    elif operator not in (MatchOperator.EMPTY, MatchOperator.NOT_EMPTY):
        assert type(data) is type(comparator), "Cannot compare types %s and %s" % (type(data), type(comparator))
    
    if operator == MatchOperator.EMPTY:
        return "" if (
            (isinstance(data, str) and len(data) == 0) or
            (data is None)
        ) else None
    elif operator == MatchOperator.NOT_EMPTY:
        return data if (
            (isinstance(data, str) and len(data) > 0) or
            (not isinstance(data, str) and data is not None)
        ) else None
    elif operator == MatchOperator.EQUAL:
        return data if data == comparator else None
    elif operator == MatchOperator.NOT_EQUAL:
        return data if data != comparator else None
    elif operator == MatchOperator.GREATER:
        return data if data > comparator else None
    elif operator == MatchOperator.GREATER_EQUAL:
        return data if data >= comparator else None
    elif operator == MatchOperator.LESS:
        return data if data < comparator else None
    elif operator == MatchOperator.LESS_EQUAL:
        return data if data <= comparator else None
    elif operator == MatchOperator.REGEX:
        match = re.search(comparator, data, re.I)
        if match is None:
            return None
        groups = match.groups()
        return groups[0] if len(groups) > 0 else data
    
    return None  # no match

# Find the right matcher class
MATCH_LOOKUP = {
    MatchTarget.CELL: {
        MatchType.DIRECT: DirectCellMatch,
        MatchType.SEPARATE: SeparateCellMatch
    },
    MatchTarget.RANGE: {
        RangeSize.NAMED: {
            MatchType.DIRECT: NamedRangeMatch,
            MatchType.SEPARATE: None,
        },
        RangeSize.TABLE: {
            MatchType.DIRECT: TableRangeMatch,
            MatchType.SEPARATE: None,
        },
        RangeSize.FIXED: {
            MatchType.DIRECT: DirectFixedRangeMatch,
            MatchType.SEPARATE: SeparateFixedRangeMatch
        },
        RangeSize.MATCHED: {
            MatchType.DIRECT: DirectMatchedRangeMatch,
            MatchType.SEPARATE: SeparateMatchedRangeMatch
        },
        RangeSize.CONTIGUOUS: {
            MatchType.DIRECT: DirectContiguousRangeMatch,
            MatchType.SEPARATE: SeparateContiguousRangeMatch
        },
    }
}

