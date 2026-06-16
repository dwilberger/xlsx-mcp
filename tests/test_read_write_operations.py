from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.read import range_read
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.workbook import workbook_open, workbook_save
from xlsx_agent.core.write import cells_write, rows_append


def _create_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "Name"
    sheet["B1"] = "Amount"
    sheet["A2"] = "Alice"
    sheet["B2"] = 10
    sheet["C2"] = True
    sheet["D2"] = date(2026, 6, 11)
    sheet["E2"] = "=B2*2"
    workbook.save(path)
    workbook.close()


def test_range_read_cells_returns_typed_values(tmp_path: Path) -> None:
    workbook_path = tmp_path / "typed.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))
    result = range_read(store, workbook_id=handle.workbook_id, sheet="Sheet1", range="A2:E2")

    assert result.resolved_range == "A2:E2"
    assert result.row_count == 1
    assert result.column_count == 5
    assert result.data[0][0]["value"]["type"] == "string"
    assert result.data[0][1]["value"]["type"] == "number"
    assert result.data[0][2]["value"]["type"] == "boolean"
    assert result.data[0][3]["value"]["type"] == "date"
    assert result.data[0][4]["value"]["type"] == "formula"


def test_range_read_matrix_returns_primitives(tmp_path: Path) -> None:
    workbook_path = tmp_path / "matrix.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))
    result = range_read(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        range="A1:B2",
        format="matrix",
    )

    assert result.data == [["Name", "Amount"], ["Alice", 10]]


def test_range_read_records_uses_first_row_as_headers(tmp_path: Path) -> None:
    workbook_path = tmp_path / "records.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))
    result = range_read(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        range="A1:B2",
        format="records",
    )

    assert result.data == [{"Name": "Alice", "Amount": 10}]


def test_range_read_records_rejects_duplicate_headers(tmp_path: Path) -> None:
    workbook_path = tmp_path / "duplicate-headers.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "Name"
    sheet["B1"] = "Name"
    sheet["A2"] = "Alice"
    sheet["B2"] = 10
    workbook.save(workbook_path)
    workbook.close()
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))

    try:
        range_read(
            store,
            workbook_id=handle.workbook_id,
            sheet="Sheet1",
            range="A1:B2",
            format="records",
        )
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"
    else:
        raise AssertionError("Expected duplicate headers to fail.")


def test_cells_write_marks_workbook_dirty_and_persists_literal_string(tmp_path: Path) -> None:
    workbook_path = tmp_path / "write.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))
    result = cells_write(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        cells={"A3": "Bob", "B3": 20, "C3": "=literal"},
    )
    saved = workbook_save(store, workbook_id=handle.workbook_id)
    reread = workbook_open(WorkbookStore(), path=saved.path)
    reread_store = WorkbookStore()
    reread_handle = workbook_open(reread_store, path=saved.path)
    reread_result = range_read(
        reread_store,
        workbook_id=reread_handle.workbook_id,
        sheet="Sheet1",
        range="A3:C3",
    )

    assert result["written_count"] == 3
    assert result["dirty"] is True
    assert reread_result.data[0][0]["value"]["value"] == "Bob"
    assert reread_result.data[0][1]["value"]["value"] == 20
    assert reread_result.data[0][2]["value"]["type"] == "string"
    assert reread_result.data[0][2]["value"]["value"] == "=literal"


def test_cells_write_rejects_missing_sheet(tmp_path: Path) -> None:
    workbook_path = tmp_path / "missing-sheet.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))

    try:
        cells_write(store, workbook_id=handle.workbook_id, sheet="Missing", cells={"A1": 1})
    except XlsxAgentError as exc:
        assert exc.code == "sheet_not_found"
    else:
        raise AssertionError("Expected write to missing sheet to fail.")


def test_rows_append_accepts_array_rows(tmp_path: Path) -> None:
    workbook_path = tmp_path / "append-arrays.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))
    result = rows_append(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        rows=[["Bob", 20], ["Carol", 30]],
    )
    read_result = range_read(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        range="A3:B4",
        format="matrix",
    )

    assert result["appended_count"] == 2
    assert read_result.data == [["Bob", 20], ["Carol", 30]]


def test_rows_append_accepts_object_rows(tmp_path: Path) -> None:
    workbook_path = tmp_path / "append-objects.xlsx"
    _create_workbook(workbook_path)
    store = WorkbookStore()

    handle = workbook_open(store, path=str(workbook_path))
    result = rows_append(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        rows=[{"Name": "Bob", "Amount": 20}, {"Name": "Carol", "Amount": 30}],
    )
    read_result = range_read(
        store,
        workbook_id=handle.workbook_id,
        sheet="Sheet1",
        range="A3:B4",
        format="matrix",
    )

    assert result["appended_count"] == 2
    assert read_result.data == [["Bob", 20], ["Carol", 30]]
