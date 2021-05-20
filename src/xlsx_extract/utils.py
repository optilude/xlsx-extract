from typing import Tuple

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table
from openpyxl.cell import Cell
from openpyxl.utils.cell import quote_sheetname

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
    sheet_id = workbook.get_index(worksheet)
    if name in workbook.defined_names.localnames(sheet_id):
        defined_name = workbook.defined_names.get(name, sheet_id)
        if defined_name is not None:
            return defined_name
    
    return None

def get_globally_defined_name(workbook : Workbook, name : str) -> DefinedName:
    """Look up a defined name global to the workbook
    """

    return workbook.defined_names.get(name, None)

def get_table(worksheet : Worksheet, name : str) -> Table:
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

def update_vector(
    source : Tuple[Tuple[Cell]],
    target : Tuple[Tuple[Cell]],
    target_reference : str,
    source_in_row : bool,
    source_idx : int,
    target_in_row : bool,
    target_idx : int,
    align : bool = False,
    expand : bool = True,
):
    """Replace a single row or column in target with a single row or
    column in source.
    """

    assert source_idx is not None, \
        "One of source row and column index must be set (this is a bug - it should not happen)"
    assert target_idx is not None, \
        "One of target row and column index must be set (this is a bug - it should not happen)"

    # Get the relevant source and target row or column into a single list
    source_vector = source[source_idx] if source_in_row else [c[source_idx] for c in source]
    target_vector = target[target_idx] if target_in_row else [c[target_idx] for c in target]

    if align:
        # Find first row or column and use as labels
        source_labels = [c.value for c in [source[0] if source_in_row else [c[0] for c in source]]]
        target_labels = [c.value for c in [target[0] if target_in_row else [c[0] for c in target]]]

        source_lookup = dict(zip(source_labels, source_vector))

        # For each target label, find and copy the corresponding source cell
        for target_label, target_cell in zip(target_labels, target_vector):
            if target_label is None or target_label == "":
                continue
            
            source_cell = source_lookup.get(target_label, None)
            if source_cell is not None:
                copy_value(source_cell, target_cell)

    else:        

        if expand:
            # TODO: `expand` support - may involve updating named references
            pass

        # Replace each value in the target bector with the corresponding value
        # in the target vector
        for source_cell, target_cell in zip(source_vector, target_vector):
            copy_value(source_cell, target_cell)

def update_table(
    source : Tuple[Tuple[Cell]],
    target : Tuple[Tuple[Cell]],
    target_reference : str,
    expand : bool = True,
):
    """Update target table with source table
    """

    if expand:
        # TODO: `expand` support - may involve updating named references
        pass
    
    for source_row, target_row in zip(source, target):
        for source_cell, target_cell in zip(source_row, target_row):
            copy_value(source_cell, target_cell)
