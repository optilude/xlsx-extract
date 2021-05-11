import pytest
import datetime

from . import match, target

def test_construct_target():
    target.Target(
        reference="Table1",
        source=match.RangeMatch(name="Source", sheet=None, reference="Table2"),
    )
