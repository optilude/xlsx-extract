import pytest
import datetime

from . import match, target

def test_construct_target():
    target.Target(
        source=match.CellMatch(name="Source", reference="SourceRef"),
        target=match.CellMatch(name="Target", reference="TargetRef"),
    )

    target.Target(
        source=match.RangeMatch(name="Source", reference="SourceRef"),
        target=match.RangeMatch(name="Target", reference="TargetRef"),
    )

    with pytest.raises(AssertionError):
        target.Target(
            source=match.RangeMatch(name="Source", reference="SourceRef"),
            target=match.CellMatch(name="Target", reference="TargetRef"),
        )
    
    target.Target(
        source=match.RangeMatch(name="Source", reference="SourceRef"),
        target=match.CellMatch(name="Target", reference="TargetRef"),
        source_col=match.CellMatch("S1", reference="C1"),
        source_row=match.CellMatch("S2", reference="C2"),
    )

    with pytest.raises(AssertionError):
        target.Target(
            source=match.CellMatch(name="Source", reference="SourceRef"),
            target=match.RangeMatch(name="Target", reference="TargetRef"),
        )
    
    target.Target(
        source=match.CellMatch(name="Source", reference="SourceRef"),
        target=match.RangeMatch(name="Target", reference="TargetRef"),
        target_col=match.CellMatch("S1", reference="C1"),
        target_row=match.CellMatch("S2", reference="C2"),
    )

