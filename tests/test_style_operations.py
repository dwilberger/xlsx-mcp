from __future__ import annotations

from pathlib import Path

import pytest

from xlsx_agent.core.columns import column_dimensions_set
from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.rows import row_dimensions_set
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.style import (
    conditional_formatting_add,
    conditional_formatting_list,
    range_style_set,
)
from xlsx_agent.core.workbook import workbook_open


@pytest.fixture
def store_with_workbook(copied_workbook_path: Path) -> WorkbookStore:
    store = WorkbookStore()
    workbook_open(store, path=str(copied_workbook_path))
    return store


def _workbook_id(store: WorkbookStore) -> str:
    return next(iter(store._items))


def test_range_style_set_applies_extended_cell_style(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = _workbook_id(store)

    result = range_style_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        range="A1:B1",
        bold=True,
        fill_color="FF112233",
        alignment="center",
        vertical_alignment="center",
        wrap_text=True,
        text_rotation=45,
        shrink_to_fit=True,
        indent=2,
        border_style="thin",
        border_color="FF445566",
    )

    cell = store.get(workbook_id).workbook["Sheet1"]["A1"]
    assert result["styled_count"] == 2
    assert cell.font.bold is True
    assert cell.fill.fgColor.rgb == "FF112233"
    assert cell.alignment.horizontal == "center"
    assert cell.alignment.vertical == "center"
    assert cell.alignment.wrap_text is True
    assert cell.alignment.text_rotation == 45
    assert cell.alignment.shrink_to_fit is True
    assert cell.alignment.indent == 2
    assert cell.border.left.style == "thin"
    assert cell.border.left.color.rgb == "FF445566"


def test_row_and_column_dimensions_set_apply_to_contiguous_dimensions(
    store_with_workbook: WorkbookStore,
) -> None:
    store = store_with_workbook
    workbook_id = _workbook_id(store)

    row_result = row_dimensions_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        start_row=2,
        count=2,
        height=24,
        hidden=True,
    )
    column_result = column_dimensions_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        start_column="B",
        count=2,
        width=18,
        hidden=True,
    )

    worksheet = store.get(workbook_id).workbook["Sheet1"]
    assert row_result["height"] == 24
    assert row_result["hidden"] is True
    assert worksheet.row_dimensions[2].height == 24
    assert worksheet.row_dimensions[3].hidden is True
    assert column_result["start_column"] == "B"
    assert column_result["width"] == 18
    assert worksheet.column_dimensions["B"].width == 18
    assert worksheet.column_dimensions["C"].hidden is True


def test_dimensions_require_a_value(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = _workbook_id(store)

    with pytest.raises(XlsxAgentError, match="Specify height or hidden"):
        row_dimensions_set(store, workbook_id=workbook_id, sheet="Sheet1", start_row=1)
    with pytest.raises(XlsxAgentError, match="Specify width or hidden"):
        column_dimensions_set(store, workbook_id=workbook_id, sheet="Sheet1", start_column="A")


def test_conditional_formatting_add_applies_visual_cell_rule_and_priority(
    store_with_workbook: WorkbookStore,
) -> None:
    store = store_with_workbook
    workbook_id = _workbook_id(store)

    result = conditional_formatting_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        range="B2:B10",
        type="cellIs",
        operator="greaterThan",
        formula="100",
        priority=5,
        fill_color="FFFF0000",
        font_color="FFFFFFFF",
        stop_if_true=True,
    )

    listed = conditional_formatting_list(store, workbook_id=workbook_id, sheet="Sheet1")
    rule = listed["rules"][0]
    assert result["priority"] == 5
    assert rule["type"] == "cellIs"
    assert rule["operator"] == "greaterThan"
    assert rule["formula"] == ["100"]
    assert rule["priority"] == 5
    assert rule["fill_color"] == "FFFF0000"
    assert rule["font_color"] == "FFFFFFFF"
    assert rule["stop_if_true"] is True


@pytest.mark.parametrize(
    ("rule_type", "expected_type"),
    [
        ("colorScale", "colorScale"),
        ("dataBar", "dataBar"),
        ("iconSet", "iconSet"),
    ],
)
def test_conditional_formatting_add_supports_visual_rule_types(
    store_with_workbook: WorkbookStore,
    rule_type: str,
    expected_type: str,
) -> None:
    store = store_with_workbook
    workbook_id = _workbook_id(store)

    conditional_formatting_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        range="B2:B10",
        type=rule_type,
    )

    listed = conditional_formatting_list(store, workbook_id=workbook_id, sheet="Sheet1")
    assert listed["rules"][0]["type"] == expected_type


def test_conditional_formatting_add_rejects_unknown_type_and_priority(
    store_with_workbook: WorkbookStore,
) -> None:
    store = store_with_workbook
    workbook_id = _workbook_id(store)

    with pytest.raises(XlsxAgentError, match="type must be one of"):
        conditional_formatting_add(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            range="B2:B10",
            type="unknown",
        )
    with pytest.raises(XlsxAgentError, match="priority must be greater"):
        conditional_formatting_add(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            range="B2:B10",
            type="formula",
            priority=0,
        )
