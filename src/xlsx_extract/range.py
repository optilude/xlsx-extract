from dataclasses import dataclass
from typing import Any, Union, Tuple, Generator

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.cell import Cell

from openpyxl.utils.cell import absolute_coordinate, quote_sheetname
from openpyxl.worksheet.worksheet import Worksheet

@dataclass
class Range:
    """One or multiple contiguous cells, possibly identified by a name,
    in a tuple (rows) of tuples (columns).
    """

    cells : Tuple[Tuple[Cell]]
    
    defined_name : DefinedName = None
    named_table : Table = None

    def __post_init__(self):
        assert not (self.defined_name is not None and self.named_table is not None), \
            "A results range cannot have both a defined name and a table name"

    @property
    def is_empty(self) -> bool:
        return len(self.cells) == 0 or len(self.cells[0]) == 0

    @property
    def is_cell(self) -> bool:
        return not self.is_empty and len(self.cells) == 1 and len(self.cells[0]) == 1
    
    @property
    def is_range(self) -> bool:
        return not self.is_empty and (len(self.cells) > 1 or len(self.cells[0]) > 1)

    @property
    def cell(self) -> Cell:
        return self.cells[0][0] if self.is_cell else None
    
    @property
    def first_cell(self) -> Cell:
        return self.cells[0][0] if not self.is_empty else None
    
    @property
    def last_cell(self) -> Cell:
        return self.cells[-1][-1] if not self.is_empty else None

    @property
    def rows(self) -> int:
        return len(self.cells) if not self.is_empty else 0
    
    @property
    def columns(self) -> int:
        return len(self.cells[0]) if not self.is_empty else 0

    @property
    def sheet(self) -> Worksheet:
        return self.cells[0][0].parent if not self.is_empty else None
    
    @property
    def workbook(self) -> Workbook:
        return self.cells[0][0].parent.parent if not self.is_empty else None

    def get_reference(self, absolute=True, use_sheet=True, use_defined_name=True, use_named_table=True) -> str:
        if self.is_empty:
            return None
        
        if use_defined_name and self.defined_name is not None:
            return self.defined_name.name

        if use_named_table and self.named_table is not None:
            return self.named_table.name
        
        prefix = "%s!" % quote_sheetname(self.sheet.title) if use_sheet else ""
        start = absolute_coordinate(self.cells[0][0].coordinate) if absolute else self.cells[0][0].coordinate

        if self.is_cell:
            return prefix + start
        
        end = absolute_coordinate(self.cells[-1][-1].coordinate) if absolute else self.cells[-1][-1].coordinate
        return prefix + start + ":" + end

    def get_values(self) -> Tuple[Tuple[Cell]]:
        return tuple(
            tuple(c.value for c in r) for r in self.cells
        )
