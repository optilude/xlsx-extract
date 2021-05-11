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
class Comparator:
    """Parameters to find a single cell
    """

    operator : Operator
    value : str = None

@dataclass
class Match:

    name : str
    sheet : SheetMatch

    # Search by cell reference (name or coordinate)
    reference : str = None

    # Define a search area
    min_row : int = None
    min_col : int = None
    max_row : int = None
    max_col : int = None

@dataclass
class CellMatch(Match):
    """Target a single cell
    """
    
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
                "%s: Either cell reference, cell value or row- and column index value must be given to identify a cell" % self.name

        if self.reference is not None:
            assert self.value is None, "%s: Cell value cannot be specified if cell reference is given" % self.name
            assert self.row_index_value is None, "%s: Row index value cannot be specified if cell reference is given" % self.name
            assert self.col_index_value is None, "%s: Column index value cannot be specified if cell reference is given" % self.name
        
        if self.value is not None:
            assert self.reference is None, "%s: Cell value cannot be specified if cell value is given" % self.name
            assert self.row_index_value is None, "%s: Row index value cannot be specified if cell value is given" % self.name
            assert self.col_index_value is None, "%s: Column index value cannot be specified if cell value is given" % self.name

        if self.row_index_value is not None:
            assert self.col_index_value is not None, "%s: If row index value is specified, column index value must also be spcified" % self.name
        if self.col_index_value is not None:
            assert self.row_index_value is not None, "%s: If column index value is specified, row index value must also be spcified" % self.name
        
        if self.row_index_value is not None:
            assert self.value is None, "%s: Cell value cannot be specified if row and column index are given" % self.name
            assert self.reference is None, "%s: Cell value cannot be specified if row and column index are given" % self.name

@dataclass
class RangeMatch(Match):

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
            "%s: Either a reference or a start cell must be specified" % self.name

        if self.reference is not None:
            assert self.start_cell is None, "%s: Start cell cannot be specified if a reference is used" % self.name
            assert self.end_cell is None, "%s: End cell cannot be specified if a reference is used" % self.name
            assert self.rows is None, "%s: Row count cannot be specified if a reference is used" % self.name
            assert self.cols is None, "%s: Column count cannot be specified if a reference is used" % self.name
            assert self.contiguous == False, "%s: Contiguousness cannot be specified if a reference is used" % self.name
        
        if self.start_cell is not None:
            assert self.reference is None, "%s: A cell reference cannot be specified if a start cell is given" % self.name
            
            # Default to contiguous mode if neither end cell or rows/cols are specified
            if self.end_cell is None and self.rows is None and self.cols is None:
                self.contiguous = True

            if self.end_cell is not None:
                assert self.rows is None and self.cols is None, "%s: Fixed row and column counts cannot be specified if an end cell is given" % self.name
                assert self.contiguous == False, "%s: Contiguousness cannot be specified if an end cell is given" % self.name
            
            if self.rows is not None:
                assert self.cols is not None, "%s: If a fixed row count is given, a fixed column count must also be specified" % self.name
            if self.cols is not None:
                assert self.rows is not None, "%s: If a fixed column count is given, a fixed row count must also be specified" % self.name

            if self.rows is not None and self.cols is not None:
                assert self.end_cell is None, "%s: An end cell cannot be specified if fixed row and column counts are given" % self.name
                assert self.contiguous == False, "%s: Contiguousness cannot be specified if fixed row and column counts are given" % self.name

            if self.contiguous:
                assert self.rows is None and self.cols is None, "%s: Fixed row and column counts cannot be specified if contiguousness is specified" % self.name
                assert self.end_cell is None, "%s: An end cell cannot be specified if contiguousness is specified" % self.name

def match_value(
    data : Union[str, int, float, bool, date, time, datetime],
    operator : Operator,
    value : Union[str, int, float, bool, date, time, datetime]
) -> Union[str, int, float, bool, date, time, datetime]:
    """Use the `operator` to compare `data` with `value`.

    Return value is `None` if not matched, or the matched item.
    For regex matches with match groups, the content of the first
    match group is returned (as a string).
    """

    if operator == Operator.REGEX:
        assert type(value) is str, "Regular expression must be a string"
    elif operator not in (Operator.EMPTY, Operator.NOT_EMPTY):
        assert type(data) is type(value), "Cannot compare types %s and %s" % (type(data), type(value))
    
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
        return data if data == value else None
    elif operator == Operator.NOT_EQUAL:
        return data if data != value else None
    elif operator == Operator.GREATER:
        return data if data > value else None
    elif operator == Operator.GREATER_EQUAL:
        return data if data >= value else None
    elif operator == Operator.LESS:
        return data if data < value else None
    elif operator == Operator.LESS_EQUAL:
        return data if data <= value else None
    elif operator == Operator.REGEX:
        match = re.search(value, data, re.IGNORECASE)
        if match is None:
            return None
        groups = match.groups()
        return groups[0] if len(groups) > 0 else data
    
    return None  # no match

