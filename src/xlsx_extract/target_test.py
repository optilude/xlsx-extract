import pytest
import datetime

from . import match, target

class TestInit:

    def test_construct_target(self):
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

    def test_init_copies_sheet_match(self):
        source_sheet = match.Comparator(operator=match.Operator.EQUAL, value="S")
        target_sheet = match.Comparator(operator=match.Operator.EQUAL, value="T")

        t = target.Target(
            source=match.CellMatch(name="Source", sheet=source_sheet, reference="A1:B2"),
            target=match.RangeMatch(name="Target", sheet=target_sheet, reference="A2:B3"),
            source_col=match.CellMatch("S1", reference="C1"),
            source_row=match.CellMatch("S2", reference="C2"),
            target_col=match.CellMatch("S1", reference="C1"),
            target_row=match.CellMatch("S2", reference="C2"),
        )

        assert t.source_row.sheet is source_sheet
        assert t.source_col.sheet is source_sheet
        assert t.target_row.sheet is target_sheet
        assert t.target_col.sheet is target_sheet


