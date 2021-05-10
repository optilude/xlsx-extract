import re

from enum import Enum
from typing import Union
from datetime import datetime, date, time
from dataclasses import dataclass

class Operator(Enum):

    EQUAL = "="
    NOT_EQUAL = "!="
    GREATER = ">"
    GREATER_EQUAL = ">="
    LESS = "<"
    LESS_EQUAL = "<="

    EMPTY = "is empty"
    NOT_EMPTY = "is not empty"

    REGEX = "regex"

@dataclass
class SheetMatch:
    """Parameters to find a sheet
    """

    operator : Operator
    value : str
    
@dataclass
class Match:
    """Parameters to find a single cell
    """

    operator : Operator
    value : str = None

@dataclass
class TargetMatch:

    name : str
    sheet : SheetMatch

    # Define a search area
    min_row : int = None
    min_col : int = None
    max_row : int = None
    max_col : int = None

@dataclass
class CellMatch(TargetMatch):
    """Target a single cell
    """
    
    # Search by cell reference (name or coordinate)
    reference : str = None

    # Search by cell contents
    value : Match = None
    
    # Find a cell by contents and use its row, and a separate cell and use its column
    row_index_value : Match = None
    col_index_value : Match = None

    # Find value in offset from the matched cell (can be positive or negative)
    row_offset : int = 0
    col_offset : int = 0

    def __post_init__(self):

        assert self.reference is not None or self.value is not None or \
                (self.row_index_value is not None and self.col_index_value is not None), \
                "Either cell reference, cell value or row- and column index value must be given to identify a cell"

        if self.reference is not None:
            assert self.value is None, "Cell value cannot be specified if cell reference is given"
            assert self.row_index_value is None, "Row index value cannot be specified if cell reference is given"
            assert self.col_index_value is None, "Column index value cannot be specified if cell reference is given"
        
        if self.value is not None:
            assert self.reference is None, "Cell value cannot be specified if cell value is given"
            assert self.row_index_value is None, "Row index value cannot be specified if cell value is given"
            assert self.col_index_value is None, "Column index value cannot be specified if cell value is given"

        if self.row_index_value is not None:
            assert self.col_index_value is not None, "If row index value is specified, column index value must also be spcified"
        if self.col_index_value is not None:
            assert self.row_index_value is not None, "If column index value is specified, row index value must also be spcified"
        
        if self.row_index_value is not None:
            assert self.value is None, "Cell value cannot be specified if row and column index are given"
            assert self.reference is None, "Cell value cannot be specified if row and column index are given"

@dataclass
class RangeMatch(TargetMatch):

    # Cell reference, defined named or table name
    reference : str = None

    # Find start of table by cell
    start_cell : CellMatch = None

    # Range extends until specified end cell
    end_cell : CellMatch = None
    
    # Range extends for a set number of rows and columns
    rows : int = None
    cols : int = None

    # Range extends until blank header row and blank first column (or later)
    # As a special case, allows top-left cell to be blank if the rest of the header is defined
    contiguous : bool = False

    def __post_init__(self):

        assert self.reference is not None or self.start_cell is not None, \
            "Either a reference or a start cell must be specified"

        if self.reference is not None:
            assert self.start_cell is None, "Start cell cannot be specified if a reference is used"
            assert self.end_cell is None, "End cell cannot be specified if a reference is used"
            assert self.rows is None, "Row count cannot be specified if a reference is used"
            assert self.cols is None, "Column count cannot be specified if a reference is used"
            assert self.contiguous == False, "Contiguousness cannot be specified if a reference is used"
        
        if self.start_cell is not None:
            assert self.reference is None, "A cell reference cannot be specified if a start cell is given"
            
            # Default to contiguous mode if neither end cell or rows/cols are specified
            if self.end_cell is None and self.rows is None and self.cols is None:
                self.contiguous = True

            if self.end_cell is not None:
                assert self.rows is None and self.cols is None, "Fixed row and column counts cannot be specified if an end cell is given"
                assert self.contiguous == False, "Contiguousness cannot be specified if an end cell is given"
            
            if self.rows is not None:
                assert self.cols is not None, "If a fixed row count is given, a fixed column count must also be specified"
            if self.cols is not None:
                assert self.rows is not None, "If a fixed column count is given, a fixed row count must also be specified"

            if self.rows is not None and self.cols is not None:
                assert self.end_cell is None, "An end cell cannot be specified if fixed row and column counts are given"
                assert self.contiguous == False, "Contiguousness cannot be specified if fixed row and column counts are given"

            if self.contiguous:
                assert self.rows is None and self.cols is None, "Fixed row and column counts cannot be specified if contiguousness is specified"
                assert self.end_cell is None, "An end cell cannot be specified if contiguousness is specified"

def match_value(
    data : Union[str, int, float, bool, date, time, datetime],
    operator : Operator,
    comparator : Union[str, int, float, bool, date, time, datetime]
) -> Union[str, int, float, bool, date, time, datetime]:
    """Use the `operator` to compare `data` with `comparator`.

    Return value is `None` if not matched, or the matched item.
    For regex matches with match groups, the content of the first
    match group is returned (as a string).
    """

    if operator == Operator.REGEX:
        assert type(comparator) is str, "Regular expression must be a string"
    elif operator not in (Operator.EMPTY, Operator.NOT_EMPTY):
        assert type(data) is type(comparator), "Cannot compare types %s and %s" % (type(data), type(comparator))
    
    if operator == Operator.EMPTY:
        return "" if (
            (isinstance(data, str) and len(data) == 0) or
            (data is None)
        ) else None
    elif operator == Operator.NOT_EMPTY:
        return data if (
            (isinstance(data, str) and len(data) > 0) or
            (not isinstance(data, str) and data is not None)
        ) else None
    elif operator == Operator.EQUAL:
        return data if data == comparator else None
    elif operator == Operator.NOT_EQUAL:
        return data if data != comparator else None
    elif operator == Operator.GREATER:
        return data if data > comparator else None
    elif operator == Operator.GREATER_EQUAL:
        return data if data >= comparator else None
    elif operator == Operator.LESS:
        return data if data < comparator else None
    elif operator == Operator.LESS_EQUAL:
        return data if data <= comparator else None
    elif operator == Operator.REGEX:
        match = re.search(comparator, data, re.IGNORECASE)
        if match is None:
            return None
        groups = match.groups()
        return groups[0] if len(groups) > 0 else data
    
    return None  # no match

