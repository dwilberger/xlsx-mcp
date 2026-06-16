import json
from pathlib import Path

from openpyxl import Workbook
from typer.testing import CliRunner

from xlsx_agent.cli import app


runner = CliRunner()


def test_version_command_returns_json() -> None:
    result = runner.invoke(app, ["version"])

    assert result.exit_code == 0
    assert '"ok": true' in result.stdout
    assert '"version": "1.0.0"' in result.stdout


def test_unimplemented_run_command_fails(tmp_path: Path) -> None:
    input_path = tmp_path / "unknown.json"
    _write_json(input_path, {})
    result = runner.invoke(app, ["run", "unknown_command", "--input", str(input_path)])

    assert result.exit_code == 1
    assert '"code": "unsupported_operation"' in result.stdout


def _write_json(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload), encoding="utf-8")


def _create_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "Name"
    sheet["B1"] = "Amount"
    sheet["A2"] = "Alice"
    sheet["B2"] = 10
    workbook.save(path)
    workbook.close()


def test_run_workbook_open_and_sheet_list(tmp_path: Path) -> None:
    workbook_path = tmp_path / "sample.xlsx"
    open_input = tmp_path / "open.json"
    list_input = tmp_path / "list.json"
    _create_workbook(workbook_path)
    _write_json(open_input, {"path": str(workbook_path)})

    open_result = runner.invoke(app, ["run", "workbook_open", "--input", str(open_input)])

    assert open_result.exit_code == 0
    open_payload = json.loads(open_result.stdout)
    workbook_id = open_payload["data"]["workbook_id"]
    _write_json(list_input, {"workbook_id": workbook_id})

    list_result = runner.invoke(app, ["run", "sheet_list", "--input", str(list_input)])

    assert list_result.exit_code == 0
    list_payload = json.loads(list_result.stdout)
    assert list_payload["data"][0]["name"] == "Sheet1"
    assert list_payload["data"][0]["used_range"] == "A1:B2"


def test_run_cells_write_and_range_read(tmp_path: Path) -> None:
    workbook_path = tmp_path / "sample.xlsx"
    open_input = tmp_path / "open.json"
    write_input = tmp_path / "write.json"
    read_input = tmp_path / "read.json"
    _create_workbook(workbook_path)
    _write_json(open_input, {"path": str(workbook_path)})

    open_result = runner.invoke(app, ["run", "workbook_open", "--input", str(open_input)])
    workbook_id = json.loads(open_result.stdout)["data"]["workbook_id"]

    _write_json(
        write_input,
        {"workbook_id": workbook_id, "sheet": "Sheet1", "cells": {"A3": "Bob", "B3": 20}},
    )
    write_result = runner.invoke(app, ["run", "cells_write", "--input", str(write_input)])

    assert write_result.exit_code == 0

    _write_json(
        read_input,
        {"workbook_id": workbook_id, "sheet": "Sheet1", "range": "A3:B3", "format": "matrix"},
    )
    read_result = runner.invoke(app, ["run", "range_read", "--input", str(read_input)])

    assert read_result.exit_code == 0
    read_payload = json.loads(read_result.stdout)
    assert read_payload["data"]["data"] == [["Bob", 20]]
