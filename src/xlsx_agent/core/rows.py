from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def rows_insert(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    start_row: int,
    count: int = 1,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    if start_row < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message="start_row must be greater than or equal to 1.",
        )
    if count < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message="count must be greater than or equal to 1.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.insert_rows(start_row, count)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "start_row": start_row,
        "inserted_count": count,
        "dirty": session.dirty,
    }


def rows_delete(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    start_row: int,
    count: int = 1,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    if start_row < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message="start_row must be greater than or equal to 1.",
        )
    if count < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message="count must be greater than or equal to 1.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.delete_rows(start_row, count)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "start_row": start_row,
        "deleted_count": count,
        "dirty": session.dirty,
    }


def rows_write(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    start_row: int,
    rows: list[list[object] | dict[str, object]],
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    if start_row < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message="start_row must be greater than or equal to 1.",
        )

    if not rows:
        return {
            "workbook_id": workbook_id,
            "sheet": sheet,
            "start_row": start_row,
            "modified_count": 0,
            "dirty": session.dirty,
        }

    first_row = rows[0]
    if isinstance(first_row, dict):
        columns = list(first_row.keys())
        for r_idx, row in enumerate(rows):
            if not isinstance(row, dict):
                raise XlsxAgentError(
                    code="invalid_input",
                    message="All rows must use the same shape.",
                )
            for c_idx, key in enumerate(columns):
                worksheet.cell(row=start_row + r_idx, column=c_idx + 1).value = row.get(key)
    else:
        for r_idx, row in enumerate(rows):
            if isinstance(row, dict):
                raise XlsxAgentError(
                    code="invalid_input",
                    message="All rows must use the same shape.",
                )
            for c_idx, value in enumerate(row):
                worksheet.cell(row=start_row + r_idx, column=c_idx + 1).value = value

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "start_row": start_row,
        "modified_count": len(rows),
        "dirty": session.dirty,
    }
