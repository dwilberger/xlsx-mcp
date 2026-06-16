from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from xlsx_agent.mcp_server import (
    cells_write,
    range_read,
    rows_append,
    sheet_list,
    workbook_close,
    workbook_open,
    workbook_save,
)


def test_mcp_workbook_open_and_sheet_list_succeeds(copied_workbook_path: Path) -> None:
    open_result = workbook_open(path=str(copied_workbook_path))

    assert open_result["ok"] is True
    workbook_id = open_result["data"]["workbook_id"]

    list_result = sheet_list(workbook_id=workbook_id)

    assert list_result["ok"] is True
    assert [sheet["name"] for sheet in list_result["data"]] == [
        "Sheet1",
        "Hidden",
        "VeryHidden",
        "Blank",
    ]

    close_result = workbook_close(workbook_id=workbook_id)
    assert close_result["ok"] is True


def test_mcp_write_read_save_flow_persists_changes(copied_workbook_path: Path) -> None:
    open_result = workbook_open(path=str(copied_workbook_path))
    workbook_id = open_result["data"]["workbook_id"]

    write_result = cells_write(workbook_id=workbook_id, sheet="Sheet1", cells={"A3": "Bob", "B3": 20})
    assert write_result["ok"] is True

    append_result = rows_append(workbook_id=workbook_id, sheet="Sheet1", rows=[["Carol", 30]])
    assert append_result["ok"] is True

    read_result = range_read(workbook_id=workbook_id, sheet="Sheet1", range="A3:B4", format="matrix")
    assert read_result["ok"] is True
    assert read_result["data"]["data"] == [["Bob", 20], ["Carol", 30]]

    save_result = workbook_save(workbook_id=workbook_id)
    assert save_result["ok"] is True
    assert save_result["data"]["dirty"] is False

    close_result = workbook_close(workbook_id=workbook_id)
    assert close_result["ok"] is True

    persisted = load_workbook(copied_workbook_path)
    try:
        assert persisted["Sheet1"]["A3"].value == "Bob"
        assert persisted["Sheet1"]["B3"].value == 20
        assert persisted["Sheet1"]["A4"].value == "Carol"
        assert persisted["Sheet1"]["B4"].value == 30
    finally:
        persisted.close()


def test_mcp_returns_error_envelope_for_dirty_close(copied_workbook_path: Path) -> None:
    open_result = workbook_open(path=str(copied_workbook_path))
    workbook_id = open_result["data"]["workbook_id"]
    cells_write(workbook_id=workbook_id, sheet="Sheet1", cells={"A3": "Bob"})

    close_result = workbook_close(workbook_id=workbook_id)

    assert close_result["ok"] is False
    assert close_result["error"]["code"] == "invalid_input"

    forced_close_result = workbook_close(workbook_id=workbook_id, force=True)
    assert forced_close_result["ok"] is True


def test_mcp_returns_validation_error_for_invalid_format(copied_workbook_path: Path) -> None:
    open_result = workbook_open(path=str(copied_workbook_path))
    workbook_id = open_result["data"]["workbook_id"]

    read_result = range_read(workbook_id=workbook_id, sheet="Sheet1", range="A1:B2", format="invalid")

    assert read_result["ok"] is False
    assert read_result["error"]["code"] == "invalid_input"
    assert read_result["error"]["details"]["errors"]

    workbook_close(workbook_id=workbook_id, force=True)
