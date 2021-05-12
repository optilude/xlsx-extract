from dataclasses import dataclass
from enum import Enum
from .match import CellMatch, Match

@dataclass
class Target:
    """Define where a matched value should go.

    There are several scenarios:

    - Target reference is a single cell, source is a single cell: Copy value from source to the target
    - Target reference is a single cell, source is a table: Use `source_row` and `source_col` to find
      a single cell and copy its value to the target.
    - Target reference is a range/table, source is a single cell: Invalid. target a single cell
      instead.
    - Target reference is a range/table, source is a range/table: There are three scenarios.
    
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
    
    # We only target cells/ranges by reference (range, name, table name)
    reference : str

    # Source cell or range we are reading from
    source : Match

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
