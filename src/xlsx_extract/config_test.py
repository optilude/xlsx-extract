from dataclasses import dataclass
from typing import Any, Tuple
from .match import Comparator, Operator

from .config import (
    interpolate_variables,
    extract_directory,
    parse_comparator,
    parse_block,
)

def test_interpolate_variables():

    assert interpolate_variables("$foo and ${Bar}", dict(foo=1, bar=2)) == "1 and 2"
    assert interpolate_variables("$foo and ${Bar}", dict(foo=1, BAR=2)) == "1 and 2"
    assert interpolate_variables("$foo and ${Bar}", dict(foo=1)) == "1 and ${Bar}"

def test_extract_directory():

    assert extract_directory({
        'directory': Comparator(Operator.EQUAL, "/foo/bar")
    }) == "/foo/bar"

    assert extract_directory({
        'foo': Comparator(Operator.EQUAL, "/foo/bar")
    }) is None

    assert extract_directory({
        'directory': Comparator(Operator.EQUAL, 13)
    }) is None

    assert extract_directory({
        'directory': Comparator(Operator.REGEX, "/foo/bar/${stuff}/bar")
    }) is None

    assert extract_directory({
        'directory': Comparator(Operator.NOT_EMPTY, "/foo/bar/${stuff}/bar")
    }) is None

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

