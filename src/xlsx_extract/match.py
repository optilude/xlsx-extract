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
class Comparator:
    """Parameters to find a single cell
    """

    operator : Operator
    value : Union[str, int, float, bool, date, time, datetime] = None

    def __post_init__(self):
        if self.operator == Operator.REGEX:
            assert type(self.value) is str, "Regular expression must be a string"

    def match(self, data : Union[str, int, float, bool, date, time, datetime]):
        """Use the `operator` to compare `data` with `value`.

        Return value is `None` if not matched, or the matched item.
        For regex matches with match groups, the content of the first
        match group is returned (as a string).
        """

        if self.operator not in (Operator.EMPTY, Operator.NOT_EMPTY):
            assert type(data) is type(self.value), "Cannot compare types %s and %s" % (type(data), type(self.value))
        
        if self.operator == Operator.EMPTY:
            return "" if (
                (isinstance(data, str) and len(data) == 0) or
                (data is None)
            ) else None
        elif self.operator == Operator.NOT_EMPTY:
            return data if (
                (isinstance(data, str) and len(data) > 0) or
                (not isinstance(data, str) and data is not None)
            ) else None
        elif self.operator == Operator.EQUAL:
            return data if data == self.value else None
        elif self.operator == Operator.NOT_EQUAL:
            return data if data != self.value else None
        elif self.operator == Operator.GREATER:
            return data if data > self.value else None
        elif self.operator == Operator.GREATER_EQUAL:
            return data if data >= self.value else None
        elif self.operator == Operator.LESS:
            return data if data < self.value else None
        elif self.operator == Operator.LESS_EQUAL:
            return data if data <= self.value else None
        elif self.operator == Operator.REGEX:
            match = re.search(self.value, data, re.IGNORECASE)
            if match is None:
                return None
            groups = match.groups()
            return groups[0] if len(groups) > 0 else data
        
        return None  # no match

@dataclass
class Match:

    name : str
    sheet : Comparator

    # Search by cell/range reference (name or coordinate)
    reference : str = None

    # Define a search area
    min_row : int = None
    min_col : int = None
    max_row : int = None
    max_col : int = None

    def get_sheet(self, workbook):
        """
        """
        pass

    def _iter_rows(self, worksheet):
        pass

@dataclass
class CellMatch(Match):
    """Target a single cell
    """
    
    # Search by cell contents
    value : Comparator = None
    
    # Find value in offset from the matched cell (can be positive or negative)
    row_offset : int = 0
    col_offset : int = 0

    def __post_init__(self):

        assert self.reference is not None or self.value is not None, \
                "%s: Either cell reference or cell value must be given to identify a cell" % self.name

        if self.reference is not None:
            assert self.value is None, "%s: Cell value cannot be specified if cell reference is given" % self.name
        
        if self.value is not None:
            assert self.reference is None, "%s: Cell value cannot be specified if cell value is given" % self.name

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
