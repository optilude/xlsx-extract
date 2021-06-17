from datetime import datetime
import os.path

import pytest
import openpyxl

from .match import (
    Operator,
    Comparator,
    CellMatch,
    RangeMatch
)
from .target import Target
from .range import Range

def get_test_workbook(filename='source.xlsx', data_only=True):
    filename = os.path.join(os.path.dirname(__file__), 'test_data', filename)
    return openpyxl.load_workbook(filename, data_only=data_only)

class TestInit:

    def test_construct_target(self):
        Target(
            source=CellMatch(name="Source", reference="SourceRef"),
            target=CellMatch(name="Target", reference="TargetRef"),
        )

        Target(
            source=RangeMatch(name="Source", reference="SourceRef"),
            target=RangeMatch(name="Target", reference="TargetRef"),
        )

        with pytest.raises(AssertionError):
            Target(
                source=RangeMatch(name="Source", reference="SourceRef"),
                target=CellMatch(name="Target", reference="TargetRef"),
            )
        
        Target(
            source=RangeMatch(name="Source", reference="SourceRef"),
            target=CellMatch(name="Target", reference="TargetRef"),
            source_col=CellMatch("S1", reference="C1"),
            source_row=CellMatch("S2", reference="C2"),
        )

        with pytest.raises(AssertionError):
            Target(
                source=CellMatch(name="Source", reference="SourceRef"),
                target=RangeMatch(name="Target", reference="TargetRef"),
            )
        
        Target(
            source=CellMatch(name="Source", reference="SourceRef"),
            target=RangeMatch(name="Target", reference="TargetRef"),
            target_col=CellMatch("S1", reference="C1"),
            target_row=CellMatch("S2", reference="C2"),
        )

    def test_init_copies_sheet_match(self):
        source_sheet = Comparator(operator=Operator.EQUAL, value="S")
        target_sheet = Comparator(operator=Operator.EQUAL, value="T")

        t = Target(
            source=CellMatch(name="Source", sheet=source_sheet, reference="A1:B2"),
            target=RangeMatch(name="Target", sheet=target_sheet, reference="A2:B3"),
            source_col=CellMatch("S1", reference="C1"),
            source_row=CellMatch("S2", reference="C2"),
            target_col=CellMatch("S1", reference="C1"),
            target_row=CellMatch("S2", reference="C2"),
        )

        assert t.source_row.sheet is source_sheet
        assert t.source_col.sheet is source_sheet
        assert t.target_row.sheet is target_sheet
        assert t.target_col.sheet is target_sheet


def test_single_cell():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    assert source_wb['Report 1']['C3'].value == datetime(2021, 5, 1)
    assert target_wb['Summary']['C3'].value != datetime(2021, 5, 1)

    t = Target(
        source=CellMatch("Date:source", reference="'Report 1'!C3"),
        target=CellMatch("Date:target", reference="'Summary'!C3"),
    )

    t.extract(source_wb, target_wb)

    assert target_wb['Summary']['C3'].value == datetime(2021, 5, 1)

def test_single_cell_triangulated_source():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    t = Target(
        source=RangeMatch("Date:source", reference="'Report 1'!B5:F9"),
        target=CellMatch("Date:target", reference="'Summary'!C3"),
        source_row=CellMatch("", value=Comparator(Operator.EQUAL, "Beta")),
        source_col=CellMatch("", value=Comparator(Operator.EQUAL, "Feb")),
    )

    t.extract(source_wb, target_wb)

    assert target_wb['Summary']['C3'].value == 7

def test_single_cell_triangulated_target():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    t = Target(
        source=RangeMatch("Date:source", reference="'Report 1'!B5:F9"),
        target=RangeMatch("Date:target", reference="'Summary'!B7:E9"),
        source_row=CellMatch("", value=Comparator(Operator.EQUAL, "Beta")),
        source_col=CellMatch("", value=Comparator(Operator.EQUAL, "Feb")),
        target_row=CellMatch("", value=Comparator(Operator.EQUAL, "Profit")),
        target_col=CellMatch("", value=Comparator(Operator.EQUAL, "Delta")),
    )

    t.extract(source_wb, target_wb)

    assert target_wb['Summary']['D8'].value == 7

def test_replace_table():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    source = Range(source_wb['Report 1']['B5:F9'])
    target = Range(target_wb['Summary']['B7:E9'])

    assert source.get_values() == (
        (None, 'Jan', 'Feb', 'Mar', 'Apr',),
        ('Alpha', 1.5, 6, 11, 4.6,),
        ('Beta', 2, 7, 12, 4.7,),
        ('Delta', 2.5, 8, 13, 4.8,),
        ('Gamma', 3, 9, 14, 4.9,),
    )

    assert target.get_values() == (
        (None, 'Alpha', 'Delta', 'Beta',),
        ('Profit', None, None, None,),
        ('Loss', None, None, None,),
    )

    assert target_wb['Summary']['B11'].value == "Area"

    t = Target(
        source=RangeMatch("Table", reference="'Report 1'!B5:F9"),
        target=RangeMatch("", reference="'Summary'!B7:E9"),
    )

    t.extract(source_wb, target_wb)    

    confirm_target = Range(target_wb['Summary']['B7:E9'])

    assert confirm_target.get_values() == (
        (None, 'Jan', 'Feb', 'Mar',),
        ('Alpha', 1.5, 6, 11,),
        ('Beta', 2, 7, 12,),
    )

    assert target_wb['Summary']['B11'].value == "Area"

def test_replace_table_expand():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    source = Range(source_wb['Report 1']['B5:F9'])
    target = Range(target_wb['Summary']['B7:E9'])

    assert source.get_values() == (
        (None, 'Jan', 'Feb', 'Mar', 'Apr',),
        ('Alpha', 1.5, 6, 11, 4.6,),
        ('Beta', 2, 7, 12, 4.7,),
        ('Delta', 2.5, 8, 13, 4.8,),
        ('Gamma', 3, 9, 14, 4.9,),
    )

    assert target.get_values() == (
        (None, 'Alpha', 'Delta', 'Beta',),
        ('Profit', None, None, None,),
        ('Loss', None, None, None,),
    )

    # Will be pushed down
    assert target_wb['Summary']['B11'].value == "Area"

    t = Target(
        source=RangeMatch("Table", reference="'Report 1'!B5:F9"),
        target=RangeMatch("", reference="'Summary'!B7:E9"),
        expand=True,
    )

    t.extract(source_wb, target_wb)

    confirm_target = Range(target_wb['Summary']['B7:F11'])

    assert confirm_target.get_values() == (
        (None, 'Jan', 'Feb', 'Mar', 'Apr',),
        ('Alpha', 1.5, 6, 11, 4.6,),
        ('Beta', 2, 7, 12, 4.7,),
        ('Delta', 2.5, 8, 13, 4.8,),
        ('Gamma', 3, 9, 14, 4.9,),
    )

    # Has been pushed down
    assert target_wb['Summary']['B13'].value == "Area"

def test_align_vector():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    source = Range(source_wb['Report 1']['B5:F9'])
    target = Range(target_wb['Summary']['B7:E9'])

    assert source.get_values() == (
        (None,   'Jan', 'Feb', 'Mar', 'Apr',),
        ('Alpha', 1.5,     6,     11,   4.6,),
        ('Beta',    2,     7,     12,   4.7,),
        ('Delta', 2.5,     8,     13,   4.8,),
        ('Gamma',   3,     9,     14,   4.9,),
    )

    assert target.get_values() == (
        (None,     'Alpha', 'Delta', 'Beta',),
        ('Profit',    None,    None,  None,),
        ('Loss',      None,    None,  None,),
    )

    # Align third column of source to second row of target
    t = Target(
        source=RangeMatch("Table", reference="'Report 1'!B5:F9"),
        target=RangeMatch("", reference="'Summary'!B7:E9"),
        source_col=CellMatch("", value=Comparator(Operator.EQUAL, "Feb")),
        target_row=CellMatch("", value=Comparator(Operator.EQUAL, "Profit")),
        align=True,
    )

    t.extract(source_wb, target_wb)

    assert Range(target_wb['Summary']['B7:E9']).get_values() == (
        (None,     'Alpha', 'Delta', 'Beta',),
        ('Profit',       6,       8,      7,),
        ('Loss',      None,    None,   None,),
    )

def test_replace_vector():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    source = Range(source_wb['Report 1']['B5:F9'])
    target = Range(target_wb['Summary']['B7:E9'])

    assert source.get_values() == (
        (None,   'Jan', 'Feb', 'Mar', 'Apr',),
        ('Alpha', 1.5,     6,     11,   4.6,),
        ('Beta',    2,     7,     12,   4.7,),
        ('Delta', 2.5,     8,     13,   4.8,),
        ('Gamma',   3,     9,     14,   4.9,),
    )

    assert target.get_values() == (
        (None,     'Alpha', 'Delta', 'Beta',),
        ('Profit',    None,    None,  None,),
        ('Loss',      None,    None,  None,),
    )

    # Replace third column of source to second row of target
    t = Target(
        source=RangeMatch("Table", reference="'Report 1'!B5:F9"),
        target=RangeMatch("", reference="'Summary'!B7:E9"),
        source_col=CellMatch("", value=Comparator(Operator.EQUAL, "Feb")),
        target_row=CellMatch("", value=Comparator(Operator.EQUAL, "Profit")),
    )

    t.extract(source_wb, target_wb)

    assert Range(target_wb['Summary']['B7:F9']).get_values() == (
        (None,   'Alpha', 'Delta', 'Beta', None,),
        ('Feb',   6,            7,      8, None,),
        ('Loss',  None,      None,   None, None,),
    )

def test_replace_vector_expand():
    source_wb = get_test_workbook('source.xlsx')
    target_wb = get_test_workbook('target.xlsx', data_only=False)

    source = Range(source_wb['Report 1']['B5:F9'])
    target = Range(target_wb['Summary']['B7:E9'])

    assert source.get_values() == (
        (None,   'Jan', 'Feb', 'Mar', 'Apr',),
        ('Alpha', 1.5,     6,     11,   4.6,),
        ('Beta',    2,     7,     12,   4.7,),
        ('Delta', 2.5,     8,     13,   4.8,),
        ('Gamma',   3,     9,     14,   4.9,),
    )

    assert target.get_values() == (
        (None,     'Alpha', 'Delta', 'Beta',),
        ('Profit',    None,    None,  None,),
        ('Loss',      None,    None,  None,),
    )

    # Replace third column of source to second row of target
    t = Target(
        source=RangeMatch("Table", reference="'Report 1'!B5:F9"),
        target=RangeMatch("", reference="'Summary'!B7:E9"),
        source_col=CellMatch("", value=Comparator(Operator.EQUAL, "Feb")),
        target_row=CellMatch("", value=Comparator(Operator.EQUAL, "Profit")),
        expand=True,
    )

    t.extract(source_wb, target_wb)

    assert Range(target_wb['Summary']['B7:F9']).get_values() == (
        (None,   'Alpha', 'Delta', 'Beta', None,),
        ('Feb',   6,            7,      8,    9,),
        ('Loss',  None,      None,   None, None,),
    )

