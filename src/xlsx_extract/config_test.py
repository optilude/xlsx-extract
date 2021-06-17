import time
import datetime
import tempfile
import os.path
import pytest
import openpyxl

from dataclasses import dataclass
from typing import Any, Tuple

from .match import CellMatch, Comparator, Operator, RangeMatch
from .target import Target

from .config import (
    Prefix,
    interpolate_variables,
    extract_directory,
    extract_filename,
    parse_comparator,
    parse_block,
    cast_col,
    contains_cell_match,
    build_cell_match,
    build_range_match,
    extract_source_match,
    extract_target,
    run,
)

def test_interpolate_variables():

    assert interpolate_variables("$foo and ${Bar}", dict(foo=1, bar=2)) == "1 and 2"
    assert interpolate_variables("$foo and ${Bar}", dict(foo=1, BAR=2)) == "1 and 2"
    assert interpolate_variables("$foo and ${Bar}", dict(foo=1)) == "1 and ${Bar}"

def test_extract_directory():

    assert extract_directory({
        'directory': Comparator(Operator.EQUAL, "/foo/bar")
    }) == "/foo/bar"

    with pytest.raises(AssertionError):
        extract_directory({
            'directory': Comparator(Operator.EQUAL, 13)
        })

    with pytest.raises(AssertionError):
        extract_directory({
            'directory': Comparator(Operator.REGEX, "/foo/bar/${stuff}/bar")
        })

    with pytest.raises(AssertionError):
        extract_directory({
            'directory': Comparator(Operator.NOT_EQUAL, "/foo/bar")
        })

def test_parse_comparator():

    assert parse_comparator("is", "foo") == Comparator(Operator.EQUAL, "foo")
    assert parse_comparator("=", "foo") == Comparator(Operator.EQUAL, "foo")
    assert parse_comparator("==", "foo") == Comparator(Operator.EQUAL, "foo")
    assert parse_comparator("is not", "foo") == Comparator(Operator.NOT_EQUAL, "foo")
    assert parse_comparator("!=", "foo") == Comparator(Operator.NOT_EQUAL, "foo")
    assert parse_comparator("matches", "foo") == Comparator(Operator.REGEX, "foo")
    assert parse_comparator("<", 1) == Comparator(Operator.LESS, 1)
    assert parse_comparator("<=", 1) == Comparator(Operator.LESS_EQUAL, 1)
    assert parse_comparator(">", 1) == Comparator(Operator.GREATER, 1)
    assert parse_comparator(">=", 1) == Comparator(Operator.GREATER_EQUAL, 1)
    assert parse_comparator("is empty", None) == Comparator(Operator.EMPTY, None)
    assert parse_comparator("empty", None) == Comparator(Operator.EMPTY, None)
    assert parse_comparator("is not empty", None) == Comparator(Operator.NOT_EMPTY, None)
    assert parse_comparator("not empty", None) == Comparator(Operator.NOT_EMPTY, None)

def test_parse_block():

    @dataclass
    class FauxRange:
        
        values : Tuple[Tuple[Any]]
        is_empty : bool = False
        is_range : bool = True
        
        @property
        def columns(self):
            return len(self.values[0])
        
        def get_values(self):
            return self.values

    assert parse_block(FauxRange((
    ), is_empty=True), dict()) is None

    assert parse_block(FauxRange((
        ("foo",),
    )), dict()) is None

    assert parse_block(FauxRange((
        ("foo", "is", 9),
    )), dict()) == dict(
        foo=Comparator(Operator.EQUAL, 9)
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9),
        ("Bar", "!=", "thirteen"),
    )), dict()) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "thirteen"),
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9),
        ("ringer", None, None),
        ("Bar", "!=", "thirteen"),
    )), dict()) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "thirteen"),
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9, "extra"),
        ("ringer", None, None, "extra"),
        ("Bar", "!=", "thirteen", "extra"),
    )), dict()) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "thirteen"),
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9, "extra"),
        ("ringer",),
        ("Bar", "!=", "thirteen", "extra"),
    )), dict()) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "thirteen"),
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9),
        ("Bar", "!=", "${foo} bar"),
    )), dict()) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "${foo} bar"),
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9),
        ("Bar", "!=", "${foo} bar"),
    )), dict(foo=3)) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "3 bar"),
    )

    assert parse_block(FauxRange((
        ("foo", "is", 9),
        ("Bar", "!=", "${Foo} bar"),
    )), dict(foo="four")) == dict(
        foo=Comparator(Operator.EQUAL, 9),
        bar=Comparator(Operator.NOT_EQUAL, "four bar"),
    )

def test_extract_filename():
    
    with tempfile.TemporaryDirectory() as current_directory:

        d = lambda f: os.path.join(current_directory, f)

        # Create some test files
        for filename in ('test1.xlsx', 'test2.xlsx', 'foo.xlsx', 'bar.txt', 'baz.xlsx',):
            time.sleep(0.01) # space out modified time - regex match should use most recent
            with open(d(filename), 'w') as fp:
                fp.write('test')
        
        # invalid arguments

        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.EQUAL, 1)
            ), current_directory)
        
        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.EQUAL, None)
            ), current_directory)
        
        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.NOT_EQUAL, "test1.xlsx")
            ), current_directory)

        # equality match

        assert extract_filename(dict(
            file=Comparator(Operator.EQUAL, "test1.xlsx")
        ), current_directory) == (d('test1.xlsx'), 'test1.xlsx')

        assert extract_filename(dict(
            file=Comparator(Operator.EQUAL, "TEST1.xlsx")
        ), current_directory) == (d('test1.xlsx'), 'test1.xlsx')

        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.EQUAL, "notfound.xlsx")
            ), current_directory)

        # do not allow directory inline

        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.EQUAL, d('test1.xlsx'))
            ), current_directory)
        
        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.EQUAL, '../test1.xlsx')
            ), current_directory)

        # regex match (test2 is a tiny but more recently modified than test 1)

        assert extract_filename(dict(
            file=Comparator(Operator.REGEX, r"(test)[0-9]\.xlsx")
        ), current_directory) == (d('test2.xlsx'), 'test')

        assert extract_filename(dict(
            file=Comparator(Operator.REGEX, r"(TEST)[0-9]\.xlsx")
        ), current_directory) == (d('test2.xlsx'), 'test')

        with pytest.raises(AssertionError):
            extract_filename(dict(
                file=Comparator(Operator.REGEX, r"notfound\.xlsx")
            ), current_directory)

def test_cast_col():

    assert cast_col("C") == 3
    assert cast_col(4) == 4
    assert cast_col(None) is None

    with pytest.raises(AssertionError):
        cast_col(3.2)
    
    with pytest.raises(AssertionError):
        cast_col("zebra")

def test_contains_cell_match():

    assert contains_cell_match({
        'cell': Comparator(Operator.EQUAL, "B3"),
        'foo': Comparator(Operator.EQUAL, "bar"),
    }) == True

    assert contains_cell_match({
        'value': Comparator(Operator.EQUAL, "foo"),
        'foo': Comparator(Operator.EQUAL, "bar"),
    }) == True

    assert contains_cell_match({
        'foo': Comparator(Operator.EQUAL, "bar"),
    }) == False

    assert contains_cell_match({
        'start cell': Comparator(Operator.EQUAL, "B3"),
        'foo': Comparator(Operator.EQUAL, "bar"),
    }) == False

    assert contains_cell_match({
        'start cell': Comparator(Operator.EQUAL, "B3"),
        'foo': Comparator(Operator.EQUAL, "bar"),
    }, Prefix.start) == True

    assert contains_cell_match(dict()) == False

def test_build_cell_match():

    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'cell': Comparator(Operator.EQUAL, "B3"),
        # 'value': Comparator(Operator.EQUAL, "baz"),
        'row offset': Comparator(Operator.EQUAL, 1),
        'column offset': Comparator(Operator.EQUAL, 2),
        'min row': Comparator(Operator.EQUAL, 3),
        'max row': Comparator(Operator.EQUAL, 4),
        'min column': Comparator(Operator.EQUAL, 5),
        'max column': Comparator(Operator.EQUAL, 'F'),
    }) == CellMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3",
        value=None,
        row_offset=1,
        col_offset=2,
        min_row=3,
        max_row=4,
        min_col=5,
        max_col=6
    )

    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        # 'cell': Comparator(Operator.EQUAL, "B3"),
        'value': Comparator(Operator.EQUAL, "baz"),
        'row offset': Comparator(Operator.EQUAL, 1),
        'column offset': Comparator(Operator.EQUAL, 2),
        'min row': Comparator(Operator.EQUAL, 3),
        'max row': Comparator(Operator.EQUAL, 4),
        'min column': Comparator(Operator.EQUAL, 5),
        'max column': Comparator(Operator.EQUAL, 'F'),
    }) == CellMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        # reference="B3",
        value=Comparator(Operator.EQUAL, "baz"),
        row_offset=1,
        col_offset=2,
        min_row=3,
        max_row=4,
        min_col=5,
        max_col=6
    )

    # passed-in name overrides name in block
    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'cell': Comparator(Operator.EQUAL, "B3"),
    }, name="quux") == CellMatch(
        name="quux",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3",
    )

    assert build_cell_match({
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'cell': Comparator(Operator.EQUAL, "B3"),
    }, name="quux") == CellMatch(
        name="quux",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3",
    )

    # passed-in sheet does not override sheet in block
    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'cell': Comparator(Operator.EQUAL, "B3"),
    }, sheet=Comparator(Operator.EQUAL, "zoo")) == CellMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3",
    )

    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'cell': Comparator(Operator.EQUAL, "B3"),
    }, sheet=Comparator(Operator.EQUAL, "zoo")) == CellMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "zoo"),
        reference="B3",
    )

    # prefix + name + sheet (this is what table and target matches willl do)

    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'start sheet': Comparator(Operator.EQUAL, "bar"),
        'start cell': Comparator(Operator.EQUAL, "B3"),
        # 'start value': Comparator(Operator.EQUAL, "baz"),
        'start row offset': Comparator(Operator.EQUAL, 1),
        'start column offset': Comparator(Operator.EQUAL, 2),
        'start min row': Comparator(Operator.EQUAL, 3),
        'start max row': Comparator(Operator.EQUAL, 4),
        'start min column': Comparator(Operator.EQUAL, 5),
        'start max column': Comparator(Operator.EQUAL, 'F'),
    }, name="foo:start", sheet=Comparator(Operator.EQUAL, "zoo"), prefix=Prefix.start) == CellMatch(
        name="foo:start",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3",
        value=None,
        row_offset=1,
        col_offset=2,
        min_row=3,
        max_row=4,
        min_col=5,
        max_col=6
    )

    assert build_cell_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'start value': Comparator(Operator.EQUAL, "baz"),
    }, name="foo:start", sheet=Comparator(Operator.EQUAL, "zoo"), prefix=Prefix.start) == CellMatch(
        name="foo:start",
        sheet=Comparator(Operator.EQUAL, "zoo"),
        value=Comparator(Operator.EQUAL, "baz"),
    )

def test_build_range_match():

    assert build_range_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'table': Comparator(Operator.EQUAL, "B3:D6"),
        # 'rows': Comparator(Operator.EQUAL, 4),
        # 'columns': Comparator(Operator.EQUAL, 5),
        # 'start cell': Comparator(Operator.EQUAL, "S1"),
        # 'end value': Comparator(Operator.EQUAL, "V1"),
    }) == RangeMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3:D6",
    )

    assert build_range_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        # 'table': Comparator(Operator.EQUAL, "B3:D6"),
        'rows': Comparator(Operator.EQUAL, 4),
        'columns': Comparator(Operator.EQUAL, 5),
        'start cell': Comparator(Operator.EQUAL, "S1"),
        # 'end value': Comparator(Operator.EQUAL, "V1"),
    }) == RangeMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        start_cell=CellMatch(
            name="foo:start",
            sheet=Comparator(Operator.EQUAL, "bar"),
            reference="S1",
        ),
        rows=4,
        cols=5,
    )

    assert build_range_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        # 'table': Comparator(Operator.EQUAL, "B3:D6"),
        # 'rows': Comparator(Operator.EQUAL, 4),
        # 'columns': Comparator(Operator.EQUAL, 5),
        'start cell': Comparator(Operator.EQUAL, "S1"),
        'end value': Comparator(Operator.EQUAL, "V1"),
        'end row offset': Comparator(Operator.EQUAL, 4)
    }) == RangeMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        start_cell=CellMatch(
            name="foo:start",
            sheet=Comparator(Operator.EQUAL, "bar"),
            reference="S1",
        ),
        end_cell=CellMatch(
            name="foo:end",
            sheet=Comparator(Operator.EQUAL, "bar"),
            value=Comparator(Operator.EQUAL, "V1"),
            row_offset=4,
        ),
    )

def test_extract_source_match():

    assert extract_source_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'cell': Comparator(Operator.EQUAL, "B3"),
        'row offset': Comparator(Operator.EQUAL, 1),
        'column offset': Comparator(Operator.EQUAL, 2),
        'min row': Comparator(Operator.EQUAL, 3),
        'max row': Comparator(Operator.EQUAL, 4),
        'min column': Comparator(Operator.EQUAL, 5),
        'max column': Comparator(Operator.EQUAL, 'F'),
    }) == CellMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3",
        row_offset=1,
        col_offset=2,
        min_row=3,
        max_row=4,
        min_col=5,
        max_col=6
    )

    assert extract_source_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'table': Comparator(Operator.EQUAL, "B3:D6"),
    }) == RangeMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        reference="B3:D6",
    )

    assert extract_source_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'start cell': Comparator(Operator.EQUAL, "S1"),
        'end value': Comparator(Operator.EQUAL, "V1"),
        'end row offset': Comparator(Operator.EQUAL, 4)
    }) == RangeMatch(
        name="foo",
        sheet=Comparator(Operator.EQUAL, "bar"),
        start_cell=CellMatch(
            name="foo:start",
            sheet=Comparator(Operator.EQUAL, "bar"),
            reference="S1",
        ),
        end_cell=CellMatch(
            name="foo:end",
            sheet=Comparator(Operator.EQUAL, "bar"),
            value=Comparator(Operator.EQUAL, "V1"),
            row_offset=4,
        ),
    )

    assert extract_source_match({
        'name': Comparator(Operator.EQUAL, "foo"),
        'sheet': Comparator(Operator.EQUAL, "bar"),
    }) == None

    assert extract_source_match({
        'sheet': Comparator(Operator.EQUAL, "bar"),
        'table': Comparator(Operator.EQUAL, "B3:D6"),
    }) == None

    with pytest.raises(AssertionError):
        extract_source_match({
            'name': Comparator(Operator.EQUAL, "foo"),
            'sheet': Comparator(Operator.EQUAL, "bar"),
            'table': Comparator(Operator.EQUAL, "B3:D6"),
            'cell': Comparator(Operator.EQUAL, "C1"),
        })

def test_extract_target():

    source_cell_match = CellMatch(name="foo", reference="Foo")
    source_range_match = RangeMatch(name="bar", reference="Bar")

    assert extract_target({
        'foo': Comparator(Operator.EQUAL, 'Baz'),
    }, source_cell_match) == None

    assert extract_target({
        'target cell': Comparator(Operator.EQUAL, 'Baz'),
    }, source_cell_match) == Target(
        source=source_cell_match,
        target=CellMatch(name="foo:target", reference="Baz"),
    )

    assert extract_target({
        'target table': Comparator(Operator.EQUAL, 'Baz'),
    }, source_range_match) == Target(
        source=source_range_match,
        target=RangeMatch(name="bar:target", reference="Baz"),
    )

    assert extract_target({
        'target table': Comparator(Operator.EQUAL, 'Baz'),
        'align': Comparator(Operator.EQUAL, True),
        'expand': Comparator(Operator.EQUAL, True),
        'source row value': Comparator(Operator.EQUAL, "alpha"),
        'source column value': Comparator(Operator.EQUAL, "beta"),
        'target row value': Comparator(Operator.EQUAL, "delta"),
        'target column value': Comparator(Operator.EQUAL, "gamma"),
    }, source_range_match) == Target(
        source=source_range_match,
        target=RangeMatch(name="bar:target", reference="Baz"),
        align=True,
        expand=True,
        source_row=CellMatch(name="bar:source_row", value=Comparator(Operator.EQUAL, "alpha")),
        source_col=CellMatch(name="bar:source_col", value=Comparator(Operator.EQUAL, "beta")),
        target_row=CellMatch(name="bar:target_row", value=Comparator(Operator.EQUAL, "delta")),
        target_col=CellMatch(name="bar:target_col", value=Comparator(Operator.EQUAL, "gamma")),
    )

    with pytest.raises(AssertionError):
        extract_target({
            'target cell': Comparator(Operator.EQUAL, 'Baz'),
            'target table': Comparator(Operator.EQUAL, 'Baz'),
        }, source_cell_match)

def test_run():
    directory = os.path.join(os.path.dirname(__file__), 'test_data')
    target_file = os.path.join(directory, 'target.xlsx')
    
    target_workbook = openpyxl.load_workbook(target_file, data_only=False)

    history = run(target_workbook, directory)

    # we could make more assertions here, but in reality it's more useful to eyeball this
    assert target_workbook['Summary']['C3'].value == datetime.datetime(2021, 5, 1)
    assert target_workbook['Summary']['F3'].value == 'input'
    assert target_workbook['Summary']['C5'].value == 12

    assert len(history) > 3  # start, file, multiple blocks, then finish 
    assert all((a.success for a in history))
    