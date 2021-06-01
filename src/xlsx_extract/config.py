import os

from dataclasses import dataclass
from string import Template
from typing import Any, Dict, List, Match, Tuple, Union
import openpyxl

from openpyxl.workbook.workbook import Workbook
from openpyxl.utils.exceptions import InvalidFileException

from .range import Range
from .match import Comparator, RangeMatch, CellMatch, Operator
from .target import Target

class GlobalKeys:
    """Keys used outside a target/match block
    """

    # Optional: Sets search directory for source files
    DIRECTORY = "directory"

    # Set source file
    FILE = "file"

class MatchKeys:
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

class CellMatchKeys:
    """Keys specific to cell matches
    """

    VALUE = "value"

    MIN_ROW = "min row"
    MAX_ROW = "max row"
    MIN_COL = "min column"
    MAX_CAL = "max column"

    ROW_OFFSET = "row offset"
    COL_OFFSET = "column offset"

class RangeMatchKeys:
    """Keys specific to range matches
    """

    ROWS = "rows"
    COLS = "columns"

class CellMatchPrefixes:
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
    "not empty": Operator.NOT_EMPTY,
}

class LowerDict(dict):
    """Dict where keys are always fetched in lowercase
    """

    def __getitem__(self, name):
        return super().__getitem__(name.lower())

class VariableTemplate(Template):
    """Case-insensitive version of String.Template
    """

    def safe_substitute(self, mapping=None, **kws):
        if mapping is None:
            mapping = {}
        m = LowerDict((k.lower(), v) for k, v in mapping.items())
        m.update(LowerDict((k.lower(), v) for k, v in kws.items()))
        return super().safe_substitute(m)

@dataclass
class Action:
    """Used for logging actions
    """

    name : str
    success : bool
    message : str

    comparator : Comparator = None
    match : Match = None
    target : Target = None


def run(target_workbook : Workbook, source_directory : str, config_sheet : str = "Config") -> List[Action]:
    """Load configuration from the given sheet in the target workbook and execute
    each step. Returns a history of what's happened. Tries pretty hard not to raise
    exceptions.
    """

    if config_sheet not in target_workbook.sheetnames:
        return None
    
    history = []

    source_workbook = None
    variables = {}

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
    
    while (match := block_match.match(target_workbook)) is not None:
        block_range, _ = match

        # prepare for next match (put this up top so we can use `continue` safely!)
        block_match.start_cell.min_row = block_range.last_cell.row + 1

        block = None
        source_match = None
        target = None

        try:
            block = parse_block(block_range, variables)
        except AssertionError as e:
            # Block contained an explicit parsing error (e.g. invalid operator)
            history.append(Action(block_range.first_cell.value, False, str(e)))
            continue
        
        if block is None:
            # Block was not a block after all - ignore
            continue
        
        if GlobalKeys.DIRECTORY in block:
            try:
                source_directory = extract_directory(block)
            except AssertionError as e:
                # Malformed block
                history.append(Action(GlobalKeys.DIRECTORY, False, str(e)))
                continue

            if source_directory is None:
                # Block was not a block after all - ignore
                continue
            
            history.append(Action(GlobalKeys.DIRECTORY, True, "Obtained %s" % source_directory, comparator=block[GlobalKeys.DIRECTORY]))
            variables[GlobalKeys.DIRECTORY] = source_directory

        if GlobalKeys.FILE in block:
            source_file = None
            file_match = None

            try:
                source_file, file_match = extract_filename(block, source_directory)
            except AssertionError as e:
                # Directory or file not found, or malformed block
                history.append(Action(GlobalKeys.FILE, False, str(e)))
                continue

            if source_file is None:
                # Block was not a block after all - ignore
                continue
            
            try:
                source_workbook = openpyxl.load_workbook(source_file, data_only=True)
            except (InvalidFileException, FileNotFoundError,) as e:
                history.append(Action(GlobalKeys.FILE, False, str(e), comparator=block[GlobalKeys.FILE]))
                continue
            
            history.append(Action(GlobalKeys.DIRECTORY, True, "Obtained %s" % source_file, comparator=block[GlobalKeys.FILE]))
            variables[GlobalKeys.FILE] = file_match

        if MatchKeys.NAME in block:
            block_name = block[MatchKeys.NAME].value

            # Don't parse target blocks if we don't yet have a file
            if source_workbook is None:
                history.append(Action(block_name, False, "No filename set ahead of %s" % block_name))
                continue

            try:
                source_match = extract_source_match(block, source_workbook)
            except AssertionError as e:
                # An assertion failed during match construction
                history.append(Action(block_name, False, str(e)))
                continue

            if source_match is None:
                history.append(Action(block_name, False, "Could not extract source match from block %s" % block_name))
                continue
        
            try:
                target = extract_target(block, source_match)
            except AssertionError as e:
                # An assertion failed during target construction
                history.append(Action(block_name, False, str(e)))
                continue
            
            match_range = None
            match_value = None

            # Source match only (used to define variables)
            if target is None:
                try:
                    match_range, match_value = source_match.match(source_workbook)
                except AssertionError as e:
                    # An assertion failed during match execution
                    history.append(Action(block_name, False, str(e), match=source_match))
                    continue
            # Target with source match
            else:
                try:
                    match_range, match_value = target.extract(source_workbook, target_workbook)
                except AssertionError as e:
                    # An assertion failed during target execution
                    history.append(Action(block_name, False, str(e), match=source_match, target=target))
                    continue
            
            if match_range is None:
                history.append(Action(block_name, False, "%s failed to match", match=source_match, target=target))
            else:
                history.append(Action(block_name, True, "Matched", match=source_match, target=target))
            
            if match_value is not None:
                variables[block_name] = match_value
    
    return history

def parse_block(match_range : Range, variables : Dict[str, Any]) -> Dict[str, Comparator]:
    """Turn a 3-column range into a dict of lowercase string keys to comparators
    """

    if match_range.is_empty or not match_range.is_range or match_range.columns < 3:
        return None

    block = {}

    for row in match_range.get_values():

        if len(row) < 3:
            continue

        name, operator, value = row[:3]
        if (
            not isinstance(name, (str, bytes,)) or name == "" or
            not isinstance(operator, (str, bytes,)) or operator == ""
        ):
            continue
        
        value = interpolate_variables(value, variables)

        comparator = parse_comparator(operator, value)
        if comparator is None:
            continue

        block[name.strip().lower()] = comparator

    return block

def parse_comparator(operator : str, value : Any) -> Comparator:
    """Build a comparator from an operator string name and a value
    """

    op = Operators.get(operator.strip().lower(), None)
    assert op is not None, "Operator `%s` not recognised" % operator
    
    return Comparator(op, value)

def extract_directory(block : Dict[str, Comparator]) -> str:
    """Extract directory path from block
    """

    comp = block.get(GlobalKeys.DIRECTORY, None)
    if comp is None:
        return None
        
    assert isinstance(comp.value, (str, bytes,)) and comp.operator == Operator.EQUAL, \
        "Directory block must use operator `is` and a string value"
    
    # When on Windows, do like the Windowsians
    return comp.value.replace('/', os.path.sep)

def extract_filename(block : Dict[str, Comparator], current_directory : str) -> Tuple[str, str]:
    """Extract filename from block. May involve matching filesystem filenames.
    Returns a tuple of validated file path and filename match. May raise an
    AssertionError if file or directory not found.
    """

    comp = block.get(GlobalKeys.FILE, None)
    if comp is None:
        return (None, None,)
    
    assert isinstance(comp.value, (str, bytes,)) and comp.operator in (Operator.EQUAL, Operator.REGEX,), \
        "File block must use operator `is` or `matches` and a string value"
    
    assert os.path.isdir(current_directory), "Directory `%s` not found" % current_directory

    with_dir = lambda f: os.path.join(current_directory, f)
    is_file = lambda f: os.path.isfile(with_dir(f))

    filename = None
    match = None

    if comp.operator == Operator.EQUAL:
        filename = match = comp.value
    elif comp.operator == Operator.REGEX:
        time_sort = lambda f: os.path.getmtime(with_dir(f))
        files = sorted(filter(is_file, os.listdir(current_directory)), key=time_sort)
    
        for f in reversed(files):
            match = comp.match(f)
            if match is not None:
                filename = f
                break

        assert filename is not None, "No matching file found for `%s` in `%s`" % (comp.value, current_directory,)

    assert is_file(filename), "File `%s` not found" % comp.value
    return (with_dir(filename), match,)


def extract_source_match(block : Dict[str, Comparator], source_workbook : Workbook) -> Match:
    """
    """

    # TODO: Construct CellMatch or RangeMatch

    return None

def extract_target(block : Dict[str, Comparator], source_match : Match) -> Target:
    """
    """

    # TODO: Construct Target

    return None

def interpolate_variables(value : str, variables : Dict[str, Any]) -> str:
    """Interpolate variables into the value string
    """

    if value is None or not isinstance(value, (str, bytes,)) or value == "":
        return value

    template = VariableTemplate(value)
    return template.safe_substitute(variables)
