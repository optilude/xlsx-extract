import os
from enum import Enum
from typing import Any, Dict, List, Match, Union

from openpyxl.workbook.workbook import Workbook

from .range import Range
from .match import Comparator, RangeMatch, CellMatch, Operator
from .target import Target

class GlobalKeys(Enum):
    """Keys used outside a target/match block
    """

    # Optional: Sets search directory for source files
    DIRECTORY = "directory"

    # Set source file
    FILE = "file"

class MatchKeys(Enum):
    """Keys used for any type of match
    """

    # Starts a new match block
    NAME = "name"

    # Sheet to search
    SHEET = "sheet"

    # A cell reference, defined name, or named table
    REFERENCE = "reference"

    # Target reference
    TARGET = "target"

    # Whether to resize target table to match source (or truncate)
    EXPAND = "expand"

    # Whether to align source and target table
    ALIGN = "align"

class CellMatchKeys(Enum):
    """Keys specific to cell matches
    """

    VALUE = "value"

    MIN_ROW = "min row"
    MAX_ROW = "max row"
    MIN_COL = "min column"
    MAX_CAL = "max column"

    ROW_OFFSET = "row offset"
    COL_OFFSET = "column offset"

class RangeMatchKeys(Enum):
    """Keys specific to range matches
    """

    ROWS = "rows"
    COLS = "columns"

class CellMatchPrefixes(Enum):
    """Prefixes used for finding cells within a table
    """

    START = "start"
    END = "end"
    
    SOURCE_ROW = "source row"
    SOURCE_COL = "source column"
    
    TARGET_ROW = "target row"
    TARGET_COL = "target column"

Operators = {
    "is": Operator.EQUAL,
    "=": Operator.EQUAL,
    "==": Operator.EQUAL,
    "is not": Operator.NOT_EQUAL,
    "!=": Operator.NOT_EQUAL,
    "matches": Operator.REGEX,
    "regex": Operator.REGEX,
    "<": Operator.LESS,
    "<=": Operator.LESS_EQUAL,
    ">": Operator.GREATER,
    ">=": Operator.GREATER_EQUAL,
    "is empty": Operator.EMPTY,
    "empty": Operator.EMPTY,
    "is not empty": Operator.NOT_EMPTY,
    "not empty": Operator.EMPTY,
}

def load(workbook : Workbook, source_directory : str, config_sheet : str = "Config") -> List[Union[Target, Match]]:
    """Load configuration from the given sheet in workbook.

    Returns a list of either targets or matches, in the order they were defined.
    """

    if config_sheet not in workbook.sheetnames:
        return None
    
    blocks = []

    source_file = None

    # Find contiguous range matches for each of "directory", "file" or "name",
    # until sheet is exhausted, and put into the `blocks` list

    block_match = RangeMatch(
        name="block",
        sheet=Comparator(Operator.EQUAL, config_sheet),
        start_cell=CellMatch("key", value=Comparator(
                Operator.REGEX, r'^\s*(' + '|'.join(
                    (GlobalKeys.DIRECTORY, GlobalKeys.FILE, MatchKeys.NAME,)
                ) + r')\s*$')
            ),
        min_row=1,
        )
    
    while (match := block_match.match(workbook)) is not None:
        block_range, _ = match

        # prepare for next match (put this up top so we can use `continue` safely!)
        block_match.start_cell.min_row = block_range.last_cell.row + 1

        source_match = None
        target = None

        block = parse_block(block_range)
        if block is None:
            continue
        
        if GlobalKeys.DIRECTORY in block:
            source_directory = extract_directory(block)
        
        if GlobalKeys.FILE in block:
            source_file = extract_filename(block, source_directory)
        
        # Don't parse target blocks if we don't yet have a file
        if source_file is None:
            continue
        
        if MatchKeys.NAME in block:
            source_match = extract_source_match(block)
        
        if source_match is not None:
            target = extract_target(block, source_match, source_file)

            if target is None:
                blocks.append(source_match)
            else:
                blocks.append(target)
    
    return blocks

def parse_block(match_range : Range) -> Dict[str, Comparator]:
    """Turn a 3-column range into a dict of lowercase string keys to comparators
    """

    if match_range.is_empty or not match_range.is_range or match_range.columns < 3:
        return None

    blocks = {}

    for row in match_range.get_values():
        name, operator, value = row[:3]
        if (
            not isinstance(name, (str, bytes,)) or name == "" or
            not isinstance(operator, (str, bytes,)) or operator == ""
        ):
            continue
        
        comparator = parse_comparator(operator, value)
        if comparator is None:
            continue

        blocks[name.strip().lower()] = comparator

    return blocks

def parse_comparator(operator : str, value : Any) -> Comparator:
    """Build a comparator from an operator string name and a value
    """

    op = Operators.get(operator.strip().lower(), None)
    if op is None:
        return None
    
    return Comparator(op, value)

def extract_directory(block : Dict[str, Comparator]) -> str:
    """Extract directory path from block
    """

    comp = block.get(GlobalKeys.DIRECTORY, None)
    if comp is None or not isinstance(comp.value, (str, bytes,)) or comp.operator != Operator.EQUAL:
        return None
    
    # When on Windows, do like the Windowsians
    return comp.value.replace('/', os.path.sep)

def extract_filename(block : Dict[str, Comparator], current_directory : str) -> str:
    """Extract filename from block. May involve matching filesystem filenames.
    """

    comp = block.get(GlobalKeys.FILE, None)
    if comp is None or not isinstance(comp.value, (str, bytes,)) or comp.operator not in (Operator.EQUAL, Operator.REGEX,):
        return None
    
    filename = comp.value

    # TODO: How do we extract filename match for later variable interpolation (esp if in a regex group)?
    if comp.operator == Operator.REGEX:
        filename = find_source_path(current_directory, comp)

    return filename


def extract_source_match(block : Dict[str, Comparator]) -> Match:
    """
    """

    return None

def extract_target(block : Dict[str, Comparator], source_match : Match, source_file : str) -> Target:
    """
    """

    return None

def find_source_path(directory : str, filename : Comparator) -> str:
    """Locate a file in the directory and return a valid file path
    """
