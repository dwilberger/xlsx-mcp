from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def freeze_panes_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    cell: str = "A1",
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.freeze_panes = cell

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "freeze_panes": worksheet.freeze_panes,
        "dirty": session.dirty,
    }


def freeze_panes_clear(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.freeze_panes = None

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "freeze_panes": None,
        "dirty": session.dirty,
    }
