from typing import Tuple

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table
from openpyxl.cell import Cell
from openpyxl.utils.cell import quote_sheetname, absolute_coordinate

def get_defined_name(workbook : Workbook, worksheet : Worksheet, name : str) -> DefinedName:
    """Get a locally or globally defined name object
    """

    defined_name = None
    if worksheet is not None:
        defined_name = get_locally_defined_name(worksheet, name)

    if defined_name is None:
        defined_name = get_globally_defined_name(workbook, name)
    
    return defined_name

def get_locally_defined_name(worksheet : Worksheet, name : str) -> DefinedName:
    """Look up a defined name local to the worksheet
    """

    workbook = worksheet.parent
    
    # First try a locally defined name
    sheet_id = workbook.index(worksheet)
    if name in workbook.defined_names.localnames(sheet_id):
        defined_name = workbook.defined_names.get(name, sheet_id)
        if defined_name is not None:
            return defined_name
    
    return None

def get_globally_defined_name(workbook : Workbook, name : str) -> DefinedName:
    """Look up a defined name global to the workbook
    """

    return workbook.defined_names.get(name, None)

def get_named_table(worksheet : Worksheet, name : str) -> Table:
    """Look up a named table
    """

    if worksheet is None or name not in worksheet.tables:
        return None

    return worksheet.tables[name]

def add_sheet_to_reference(worksheet : Worksheet, ref : str) -> str:
    """Add worksheet name to table if needed
    """
    if '!' not in ref:
        assert worksheet is not None, "Sheet must be given if reference does not contain a sheet name"
        ref = "%s!%s" % (quote_sheetname(worksheet.title), ref)
    return ref

def triangulate_cell(row : Cell, col : Cell) -> Cell:
    """Find the cell at the intersection of the row of `row`
    and the column of `col`.
    """
    assert row.parent is col.parent
    return row.parent.cell(row.row, col.column)

def copy_value(source : Cell, target : Cell):
    """Copy a single value from source to target
    """
    target.value = source.value

def get_reference_for_table(table : Tuple[Tuple[Cell]]) -> str:
    """Get an absolute cell range reference for the given table
    """
    assert len(table) > 0 and len(table[0]) > 0, \
        "Cannot create a reference for an empty table"
    
    first_cell = table[0][0]
    last_cell = table[-1][-1]

    first_cell_coordinate = absolute_coordinate(first_cell.coordinate)
    last_cell_coordinate = absolute_coordinate(last_cell.coordinate)

    worksheet = first_cell.parent

    if first_cell.coordinate == last_cell.coordinate:
        return "%s!%s" % (quote_sheetname(worksheet.title), first_cell_coordinate)
    else:
        return "%s!%s:%s" % (quote_sheetname(worksheet.title), first_cell_coordinate, last_cell_coordinate)

def update_name(worksheet : Worksheet, reference : str, table : Tuple[Tuple[Cell]]) -> bool:
    """Update a named reference or table to point to the table.
    """

    assert len(table) > 0 and len(table[0]) > 0, \
        "Cannot update name for an empty table"

    workbook = worksheet.parent

    defined_name = get_defined_name(workbook, worksheet, reference)
    if defined_name is not None:
        defined_name.attr_text = get_reference_for_table(table)
        return True
    
    named_table = get_named_table(worksheet, reference)
    if named_table is not None:
        named_table.ref = "%s:%s" % (table[0][0].coordinate, table[-1][-1].coordinate)
        return True
    
    return False

def resize_table(table : Tuple[Tuple[Cell]], rows : int, cols : int, reference : str = None) -> Tuple[Tuple[Cell]]:
    """Add or remove rows or columns at the end of of `table` so that it has the dimensions
    `rows` x `cols`. If `reference` is a named table or named table, update it in the parent
    workshet/workbook to reflect the new dimensions. Return new table with the correct
    dimensions.
    """

    assert len(table) > 0 and len(table[0]) > 0, \
        "Cannot resize an empty range"
    
    # No change
    if rows == len(table) and cols == len(table[0]):
        return table

    rows_delta = rows - len(table)
    cols_delta = cols - len(table[0])

    first_cell = table[0][0]
    last_cell = table[-1][-1]

    sheet = first_cell.parent

    # Add new rows to the bottom
    if rows_delta > 0:
        sheet.insert_rows(last_cell.row + 1, rows_delta)
    
    # Remove rows from the bottom
    if rows_delta < 0:
        sheet.delete_rows(first_cell.row + rows, -rows_delta)
    
    # Add new columns to the end
    if cols_delta > 0:
        sheet.insert_cols(last_cell.column + 1, cols_delta)
    
    # Remove columns from the top
    if cols_delta < 0:
        sheet.delete_cols(first_cell.column + cols, -cols_delta)
    
    new_table = tuple(
        sheet.iter_rows(
            min_row=first_cell.row,
            min_col=first_cell.column,
            max_row=first_cell.row + (rows - 1),
            max_col=first_cell.column + (cols - 1)
        )
    )

    if reference is not None and (rows_delta != 0 or cols_delta != 0):
        update_name(sheet, reference, new_table)
    
    return new_table

def update_table(
    source : Tuple[Tuple[Cell]],
    target : Tuple[Tuple[Cell]],
    target_reference : str,
    expand : bool = True,
):
    """Update target table with source table
    """

    assert len(target) > 0 and len(target[0]) > 0, \
        "Cannot target an empty table (this is a bug - it should not happen)"
    
    assert len(source) > 0 and len(source[0]) > 0, \
        "Cannot copy an empty table (this is a bug - it should not happen)"

    if expand:
        target = resize_table(target, len(source), len(source[0]), target_reference)
    
    for source_row, target_row in zip(source, target):
        for source_cell, target_cell in zip(source_row, target_row):
            copy_value(source_cell, target_cell)

def extract_vector(table : Tuple[Tuple[Cell]], in_row : bool, index : int) -> Tuple[Cell]:
    """Get a tuple of the cells in the row at `index` if `in_row`, or in the
    column at `index` if not `in_row`.
    """
    return (table[index] if in_row else [c[index] for c in table])

def align_vectors(
    source : Tuple[Tuple[Cell]],
    source_in_row : bool,
    source_idx : int,
    target : Tuple[Tuple[Cell]],
    target_in_row : bool,
    target_idx : int,
):
    """Replace a vector in `target` with a vector in `source` by matching
    labels in the first row/column of each.
    """

    assert len(target) > 0 and len(target[0]) > 0, \
        "Cannot target an empty table (this is a bug - it should not happen)"

    assert source_idx is not None, \
        "One of source row and column index must be set (this is a bug - it should not happen)"
    assert target_idx is not None, \
        "One of target row and column index must be set (this is a bug - it should not happen)"

    to_label = lambda s: s.strip().lower() if isinstance(s, (str, bytes,)) else s

    # Get the relevant source and target row or column into a single list
    source_vector = extract_vector(source, source_in_row, source_idx)
    target_vector = extract_vector(target, target_in_row, target_idx)

    assert source_idx > 0 and source_idx < len(source_vector), \
        "Source row/column is outside of the target table (this is a bug - it should not happen)"
    assert target_idx > 0 and target_idx < len(target_vector), \
        "Target row/column is outside of the target table (this is a bug - it should not happen)"

    # Find first row or column and use as labels
    source_labels = (to_label(c.value) for c in extract_vector(source, source_in_row, 0))
    target_labels = (to_label(c.value) for c in extract_vector(target, target_in_row, 0))

    source_lookup = dict(zip(source_labels, source_vector))

    # For each target label, find and copy the corresponding source cell
    for target_label, target_cell in zip(target_labels, target_vector):
        source_cell = source_lookup.get(target_label, None)
        if source_cell is not None:
            copy_value(source_cell, target_cell)


def replace_vector(
    source : Tuple[Tuple[Cell]],
    source_in_row : bool,
    source_idx : int,
    target : Tuple[Tuple[Cell]],
    target_in_row : bool,
    target_idx : int,
    target_reference : str,
    expand : bool,
):
    """Replace a single row or column in target with a single row or column in source.
    """

    assert len(target) > 0 and len(target[0]) > 0, \
        "Cannot target an empty table (this is a bug - it should not happen)"

    assert source_idx is not None, \
        "One of source row and column index must be set (this is a bug - it should not happen)"
    assert target_idx is not None, \
        "One of target row and column index must be set (this is a bug - it should not happen)"

    # Get the relevant source and target row or column into a single list
    source_vector = extract_vector(source, source_in_row, source_idx)
    target_vector = extract_vector(target, target_in_row, target_idx)

    if expand:
        rows = len(source_vector) if not target_in_row else len(target)
        cols = len(source_vector) if target_in_row else len(target[0])

        target = resize_table(target, rows, cols, target_reference)
        target_vector = extract_vector(target, target_in_row, target_idx)

    assert source_idx > 0 and source_idx < len(source_vector), \
        "Source row/column is outside of the target table (this is a bug - it should not happen)"
    assert target_idx > 0 and target_idx < len(target_vector), \
        "Target row/column is outside of the target table (this is a bug - it should not happen)"

    # Replace each value in the target bector with the corresponding value
    # in the target vector
    for source_cell, target_cell in zip(source_vector, target_vector):
        copy_value(source_cell, target_cell)

