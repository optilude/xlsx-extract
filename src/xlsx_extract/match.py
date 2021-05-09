from enum import Enum
from dataclasses import dataclass

class MatchTarget(Enum):

    CELL = "cell"   # Find a single cell matching the parameter(s)
    RANGE = "range" # Find a range matching the parameter(s)

class MatchType(Enum):

    DIRECT = "direct"            # Target matches row and column of parameter
    SEPARATE = "separate"        # Separate parameters for row and column

class MatchOperator(Enum):

    EQUALS = "equals"
    GREATER = "greater than"
    GREATER_EQUAL = "greater than or equal to"
    LESS = "less than"
    LESS_EQUAL = "less than or equal to"

    EMPTY = "empty"
    NOT_EMPTY = "not empty"

    REFERENCE = "cell reference"
    NAMED_REFERENCE = "named reference"

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