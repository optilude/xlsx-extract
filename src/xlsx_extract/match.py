from enum import Enum
from dataclasses import dataclass

class MatchTarget(Enum):

    CELL = "Cell"   # Find a single cell matching the parameter(s)
    RANGE = "Range" # Find a range matching the parameter(s)

class MatchType(Enum):

    DIRECT = "Direct"            # Target matches row and column of parameter
    SEPARATE = "Separate"        # Separate parameters for row and column

class MatchOperator(Enum):

    EQUALS = "Equals"
    GREATER = "Greater Than"
    GREATER_EQUAL = "Greater Than or Equal To"
    LESS = "Less Than"
    LESS_EQUAL = "Less Than or Equal To"

    EMPTY = "Empty"
    NOT_EMPTY = "Not Empty"

    REFERENCE = "Cell reference"
    NAMED_REFERENCE = "Named reference"

class RangeSize(Enum):

    NAMED = "Named"             # Find a named range
    FIXED = "Fixed"             # Specify a number of rows and column
    MATCHED = "Match cell"      # Match the end of the range using match parameters
    CONTIGUOUS = "Contiguous"   # Range extends across the header row and down the first column until a blank cell is found

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

@dataclass
class TargetMatch:
    """Find a target cell or range
    """

    sheet : SheetMatch      # which sheet are we looking in
    target : MatchTarget    # looking for a cell or a range
    match_type : MatchType  # one parameter (cell) or two (row, col separate)

    min_row : int
    min_col : str
    max_row : int
    max_col : str

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

    range_start_match : TargetMatch
    range_size : RangeSize

    def __post_init__(self):
        assert self.match_type == self.range_start_match.match_type
        assert self.target == MatchTarget.RANGE

@dataclass
class NamedRangeMatch(RangeMatch):
    """Target a named range
    """

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.NAMED
        assert self.match_type == MatchType.DIRECT
        assert self.range_start_match.cell_match.operator == MatchOperator.NAMED_REFERENCE

@dataclass
class FixedRangeMatch(RangeMatch):
    """Target a range of fixed size
    """

    range_rows : int
    range_cols : int

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.FIXED
        assert self.range_rows != None and self.range_rows > 0
        assert self.range_cols != None and self.range_cols > 0

@dataclass
class MatchedRangeMatch(RangeMatch):
    """Target a range of fixed size
    """

    range_end_match : TargetMatch

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.MATCHED

@dataclass
class ContiguousRangeMatch(RangeMatch):
    """Target a contiguous range of cells
    """

    def __post_init__(self):
        super().__post_init__()
        assert self.range_size == RangeSize.CONTIGUOUS
