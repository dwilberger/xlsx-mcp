from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def merged_cells_list(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    worksheet = get_worksheet(session.workbook, sheet)

    merges = [str(m) for m in worksheet.merged_cells.ranges]

    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "merged_ranges": merges,
        "count": len(merges),
    }


def merge_cells(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    range: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.merge_cells(range)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "range": range,
        "dirty": session.dirty,
    }


def unmerge_cells(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    range: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    worksheet.unmerge_cells(range)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "range": range,
        "dirty": session.dirty,
    }
