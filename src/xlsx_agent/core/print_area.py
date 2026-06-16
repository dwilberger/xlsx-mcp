from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def print_area_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    range: str | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.print_area = range

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "print_area": worksheet.print_area or None,
        "dirty": session.dirty,
    }


def print_titles_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    rows: str | None = None,
    columns: str | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    if rows is not None:
        worksheet.print_title_rows = rows
    if columns is not None:
        worksheet.print_title_cols = columns

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "rows": worksheet.print_title_rows,
        "columns": worksheet.print_title_cols,
        "dirty": session.dirty,
    }
