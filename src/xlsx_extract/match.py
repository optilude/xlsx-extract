import re

from enum import Enum
from typing import Any, Union, Tuple, Generator
from datetime import datetime, date, time
from dataclasses import dataclass

from openpyxl import Workbook
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import quote_sheetname, range_to_tuple

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

    # What sheet are we on? Optional if `reference` is set
    sheet : Comparator = None

    # Search by cell/range reference (name or coordinate)
    reference : str = None

    # Define a search area
    min_row : int = None
    min_col : int = None
    max_row : int = None
    max_col : int = None

    def match(self, workbook : Workbook, worksheet : Worksheet = None) -> Tuple[Union[Cell,Tuple[Cell]], Any]:
        """Match current parameters in worksheet and return a tuple of
        `(matched cell(s), matched value)`.
        """
        # Subclasses will implement
        raise NotImplementedError

    def get_sheet(self, workbook : Workbook) -> Tuple[Worksheet, str]:
        """Return the worksheet matching `self.sheet`, returning
        a tuple of the worksheet object and the matched title
        """
        for ws in workbook.worksheets:
            match = self.sheet.match(ws.title)
            if match is not None:
                return (ws, match)
        return (None, None)

    def find_by_reference(self, workbook : Workbook, worksheet : Worksheet = None) -> Union[Cell, Tuple[Cell]]:
        """Find the cell or range matching `self.reference`.
        """

        if self.reference is None:
            return None
        
        ref = self.reference
        found_locally = False

        # First try a locally defined name
        if worksheet is not None:
            sheet_id = workbook.get_index(worksheet)
            if self.reference in workbook.defined_names.localnames(sheet_id):
                defined_name = workbook.defined_names.get(self.reference, sheet_id)
                if defined_name is not None:
                    ref = defined_name.attr_text
                    found_locally = True
        
        # Then try a globally defined name
        if not found_locally and self.reference in workbook.defined_names:
            ref = workbook.defined_names[self.reference].attr_text

        if ref is None or ref == "":
            return None

        # Add sheet name to reference if needed
        if '!' not in ref:
            assert worksheet is not None, "Sheet must be given if reference does not contain a sheet name"
            ref = "%s!%s" % (quote_sheetname(worksheet.title), ref)

        # Get range numerically
        sheet_name, (r1, c1, r2, c2) = range_to_tuple(ref)
        
        sheet = workbook[sheet_name]

        # Single cell
        if r1 == r2 and c1 == c2:
            return sheet.cell(r1, c1)
        # Range
        else:
            return tuple(sheet.iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2))

    def _iter_rows(self, worksheet : Worksheet) -> Generator[Tuple[Cell], None, None]:
        """Iterate over rows (list of cells) in the worksheet within the 
        min/max row/col boundaries.
        """
        return worksheet.iter_rows(
            min_row=self.min_row,
            max_row=self.max_row,
            min_col=self.min_col,
            max_col=self.max_col
        )

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
            assert self.sheet is not None, "%s: Sheet is required if matching by value" % self.name
            assert self.reference is None, "%s: Cell value cannot be specified if cell value is given" % self.name

    def match(self, workbook : Workbook, worksheet : Worksheet = None) -> Tuple[Cell, Any]:
        """Match a single cell
        """

        cell = None
        match = None

        if self.reference is not None:
            cell = self.find_by_reference(workbook, worksheet)
        elif self.value is not None:
            cell, match = self.find_by_value(worksheet)
        
        if cell is not None:
            cell = self.apply_offset(cell)
        
        return (cell, match)

    def find_by_value(self, worksheet : Worksheet) -> Tuple[Cell, Any]:
        """Search the worksheet for a cell by value comparator, returning
        a tuple of `(cell, match)` or `(None, None)`.
        """
        if self.value is None or worksheet is None:
            return (None, None)

        for row in self._iter_rows(worksheet):
            for cell in row:
                match_value = self.value.match(cell.value)
                if match_value is not None:
                    return (cell, match_value)
    
    def apply_offset(self, cell : Cell) -> Cell:
        """Return a cell at the current row/col offset from the input cell
        """
        row = cell.row + self.row_offset
        col = cell.col_idx + self.col_offset
        return cell.parent.cell(row, col)

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

    def match(self, workbook : Workbook, worksheet : Worksheet = None) -> Tuple[Tuple[Cell], Any]:
        """Match a range cell
        """
        # TODO
    
    def find_by_end_cell(self, worksheet : Worksheet) -> Tuple[Cell]:
        """Find range by `self.start_cell` and `self.end_cell`.
        """
        # TODO
    
    def find_by_dimensions(self, worksheet : Worksheet) -> Tuple[Cell]:
        """Find range by `self.start_cell`, `self.rows` and `self.cols`.
        """
        # TODO
    
    def find_by_contiguous_region(self, worksheet : Worksheet) -> Tuple[Cell]:
        """Find range contiguously from `self.start_cell`.
        """
        # TODO
