from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook
from typer.testing import CliRunner

from xlsx_agent.cli import app
from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.read import range_read
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.workbook import workbook_open, workbook_save
from xlsx_agent.core.write import cells_write, rows_append


runner = CliRunner()


def test_workbook_save_detects_deleted_original_file(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path))
    cells_write(store, workbook_id=handle.workbook_id, sheet="Sheet1", cells={"A3": "Bob"})
    copied_workbook_path.unlink()

    try:
        workbook_save(store, workbook_id=handle.workbook_id)
    except XlsxAgentError as exc:
        assert exc.code == "save_conflict"
    else:
        raise AssertionError("Expected save to fail when the original file was deleted.")


def test_range_read_rejects_invalid_range_string(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path))

    try:
        range_read(store, workbook_id=handle.workbook_id, sheet="Sheet1", range="not-a-range")
    except XlsxAgentError as exc:
        assert exc.code == "invalid_range"
    else:
        raise AssertionError("Expected invalid range input to fail.")


def test_range_read_records_rejects_empty_headers(copied_workbook_path: Path) -> None:
    workbook = load_workbook(copied_workbook_path)
    workbook["Sheet1"]["A1"] = None
    workbook.save(copied_workbook_path)
    workbook.close()

    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path))

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
        raise AssertionError("Expected records mode to reject empty headers.")


def test_cells_write_rejects_invalid_cell_address(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path))

    try:
        cells_write(store, workbook_id=handle.workbook_id, sheet="Sheet1", cells={"A0": 1})
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"
    else:
        raise AssertionError("Expected invalid cell address to fail.")


def test_cells_write_rejects_read_only_workbook(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path), read_only=True)

    try:
        cells_write(store, workbook_id=handle.workbook_id, sheet="Sheet1", cells={"A3": "Bob"})
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    else:
        raise AssertionError("Expected cells_write to fail for read-only workbooks.")


def test_rows_append_rejects_mixed_row_shapes(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path))

    try:
        rows_append(
            store,
            workbook_id=handle.workbook_id,
            sheet="Sheet1",
            rows=[["Bob", 20], {"Name": "Carol", "Amount": 30}],
        )
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"
    else:
        raise AssertionError("Expected mixed row shapes to fail.")


def test_rows_append_rejects_read_only_workbook(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    handle = workbook_open(store, path=str(copied_workbook_path), read_only=True)

    try:
        rows_append(store, workbook_id=handle.workbook_id, sheet="Sheet1", rows=[["Bob", 20]])
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    else:
        raise AssertionError("Expected rows_append to fail for read-only workbooks.")


def test_cli_run_rejects_invalid_json_input(tmp_path: Path) -> None:
    input_path = tmp_path / "invalid.json"
    input_path.write_text("{\n  \"path\":\n", encoding="utf-8")

    result = runner.invoke(app, ["run", "workbook_open", "--input", str(input_path)])

    assert result.exit_code == 1
    payload = json.loads(result.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "invalid_input"
    assert payload["error"]["details"]["line"] > 0
    assert payload["error"]["details"]["column"] > 0
