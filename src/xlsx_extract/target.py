from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Union, Tuple

from openpyxl import Workbook
from openpyxl.cell import Cell

from .match import Match, CellMatch, RangeMatch
from .utils import copy_value, triangulate_cell, update_table, update_vector

def locate_cell_in_range(workbook : Workbook, range_cells : Tuple[Tuple[Cell]], cell_match : CellMatch) -> Cell:
    """Use `cell_match` to find a cell within the range
    """

    if len(range_cells) == 0 or len(range_cells[0]) == 0 or len(range_cells[-1]) == 0:
        return None

    m = deepcopy(cell_match)
                
    m.min_row = range_cells[0][0].row
    m.min_col = range_cells[0][0].column
    m.max_row = range_cells[-1][-1].row
    m.max_col = range_cells[-1][-1].column

    cell, _ = m.match(workbook)

    if (
        cell is not None and (
            cell.column < m.min_col or cell.column > m.max_col or 
            cell.row < m.min_row or cell.row > m.max_row
    )):
        return None

    return cell

@dataclass
class Target:
    """Define where a matched value should go.

    There are several scenarios:

    - Target is a single cell, source is a single cell: Copy value from source to the target
    - Target is a single cell, source is a table: Use `source_row` and `source_col` to find
        a single cell and copy its value to the target.
    - Target is a range/table, source is a single cell: Use `target_row` and `target_col`
        to target a specific cell.
    - Target is a range/table, source is a range/table: There are three scenarios.
    
        1. Replace the whole target table with the source: Set `reference` and `source` only.
        2. Replace an entire row or column in the target with a row or column in the source:
                Set one of `source_row` or `source_col` (but not both) and one of `target_row`
                and `target_col` (but not both), and keep `align=False`. The target table will
                be expanded to accommodate the source if `expand` is `True`, otherwise the source
                vector may be truncated. Note that it is possible to transpose a column into a
                row or vice-versa with this approach, by setting `source_row` and `target_col`
                or vice-versa.
        3. Populate a table by aligning row/column labels from the source to corresponding
                row/column labels in the target: Set `source_row` or `source_col` (but not both)
                and `target_row` and `target_col` (but not both) and set `align=True`. The
                label is the text from the first row/column in the tables. Again, this may be
                used to transpose data, i.e. labels corresponding to a source row can be aligned
                to labels in a target column, and vice-versa.
    
    `source_row`, `source_col`, `target_row` and `target_col` are found with cell matches.
    These are used to identify a row/column number within the source/target table range,
    but can be found in any way, e.g. it's possible to use a cell match to search for
    a cell in the middle of the table and extrapolate its row or column number.
    """
    
    # Source cell or range we are reading from
    source : Match

    # Target cell or range we are writing to
    target : Match

    # If targetting a subset of a source range, look up row and/or column by header
    # These will be matched inside the range of the table only, and used to identify
    # a row and/or column number
    source_row : CellMatch = None
    source_col : CellMatch = None

    # If populating a single row or column in a table, look up row or column by header.
    # Again, these are matched within the target range only and used to identify row/col number.
    target_row : CellMatch = None
    target_col : CellMatch = None
    
    # How to insert range: align to target table row/column labels or copy directly
    align : bool = False

    # Whether to expand table or truncate values when copying (if align == False)
    expand : bool = True

    def __post_init__(self):
        
        # Range -> Cell: Need source row and column
        if isinstance(self.source, RangeMatch) and isinstance(self.target, CellMatch):
            assert self.source_row is not None and self.source_col is not None, \
                "A source row and column must be specified if the source is a range and the target is a cell"

        # Cell -> Range: Need target row and column
        if isinstance(self.source, CellMatch) and isinstance(self.target, RangeMatch):
            assert self.target_row is not None and self.target_col is not None, \
                "A target row and column must be specified if the source is a cell and the target is a range"

        if self.source_row is not None and self.source.sheet is not None:
            self.source_row.sheet = self.source.sheet
        if self.source_col is not None and self.source.sheet is not None:
            self.source_col.sheet = self.source.sheet
        if self.target_row is not None and self.target.sheet is not None:
            self.target_row.sheet = self.target.sheet
        if self.target_col is not None and self.target.sheet is not None:
            self.target_col.sheet = self.target.sheet

    def extract(self, source_workbook : Workbook, target_workbook : Workbook) -> Tuple[Union[Cell,Tuple[Cell]], Any]:
        """Extract source cell from the source workbook and update target workbook.
        Returns source match.
        """

        NOT_FOUND = (None, None,)

        source_is_range = isinstance(self.source, RangeMatch)
        target_is_range = isinstance(self.target, RangeMatch)

        source_c, _ = original_match = self.source.match(source_workbook)
        if source_c is None or len(source_c) == 0:
            return NOT_FOUND

        target_c, _ = self.target.match(target_workbook)
        if target_c is None or len(target_c) == 0:
            return NOT_FOUND

        source_row_cell = None
        source_col_cell = None
        target_row_cell = None
        target_col_cell = None

        # Look for the cells that define rows and columns

        if source_is_range:
            if self.source_row is not None:
                source_row_cell = locate_cell_in_range(source_workbook, source_c, self.source_row)
                if source_row_cell is None:
                    return NOT_FOUND
            if self.source_col is not None:
                source_col_cell = locate_cell_in_range(source_workbook, source_c, self.source_col)
                if source_col_cell is None:
                    return NOT_FOUND
        
        if target_is_range:
            if self.target_row is not None:
                target_row_cell = locate_cell_in_range(target_workbook, target_c, self.target_row)
                if target_row_cell is None:
                    return NOT_FOUND
            if self.target_col is not None:
                target_col_cell = locate_cell_in_range(target_workbook, target_c, self.target_col)
                if target_col_cell is None:
                    return NOT_FOUND

        # If we have a range and two locators, resolve to a single cell

        if source_is_range and source_row_cell is not None and source_col_cell is not None:
            source_c = ((triangulate_cell(source_row_cell, source_col_cell),),)
            source_is_range = False
        
        if target_is_range and target_row_cell is not None and target_col_cell is not None:
            target_c = ((triangulate_cell(target_row_cell, target_col_cell),),)
            target_is_range = False
        
        # Both source and target might now have changed, but both should be the same type
        assert source_is_range == target_is_range, \
            "%s: Cannot copy a table to a single cell or vice-versa" % self.source.name
        
        # They should also be non-empty
        assert len(source_c) > 0 and len(source_c[0]) > 0, \
            "%s: Source cell range is empty (this is a bug - it should never happen)" % self.source.name
        assert len(target_c) > 0 and len(target_c[0]) > 0, \
            "%s: Target cell range is empty (this is a bug - it should never happen)" % self.source.name

        # We also should have at most one of source_row and source_col set (to identify a vector)
        assert not source_is_range or (
            (source_row_cell is None and source_col_cell is None) or 
            (source_row_cell is None and source_col_cell is not None) or
            (source_row_cell is not None and source_col_cell is None)
        ), "%s: Both source row and source column are set but cell has not been located (this is a bug - it should never happen)" % self.source.name
        
        # And the same for target_row and target_col
        assert not target_is_range or (
            (target_row_cell is None and target_col_cell is None) or 
            (target_row_cell is None and target_col_cell is not None) or
            (target_row_cell is not None and target_col_cell is None)
        ), "%s: Both target row and target column are set but cell has not been located (this is a bug - it should never happen)" % self.source.name

        if source_is_range:
            if all(c is None for c in (source_row_cell, source_col_cell, target_row_cell, target_col_cell,)):
                # Update/replace an entire table
                update_table(
                    source=source_c,
                    target=target_c,
                    target_reference=self.target.reference,
                    expand=self.expand
                )
            else:
                # Update/replace a row or column (possibly transposed)
                update_vector(
                    source=source_c,
                    target=target_c,
                    target_reference=self.target.reference,
                    source_in_row=(source_row_cell is not None),
                    source_idx=source_row_cell.row - source_c[0][0].row if source_row_cell is not None else source_col_cell.column  - source_c[0][0].column if source_col_cell is not None else None,
                    target_in_row=(target_col_cell is not None),
                    target_idx=target_row_cell.row - target_c[0][0].row if target_row_cell is not None else target_col_cell.column  - target_c[0][0].column if target_col_cell is not None else None,
                    align=self.align,
                    expand=self.expand
                )
        else:
            # Update/replace a single cell
            copy_value(source_c[0][0], target_c[0][0])

        return original_match


            
        