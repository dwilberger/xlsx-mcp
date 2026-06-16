from __future__ import annotations

from openpyxl.comments import Comment

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def comment_add(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    cell: str,
    text: str,
    author: str = "",
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    try:
        cell_obj = worksheet[cell]
    except KeyError as exc:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell '{cell}' is not valid.",
        ) from exc

    comment = Comment(text=text, author=author)
    cell_obj.comment = comment

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "cell": cell,
        "text": text,
        "author": author,
        "dirty": session.dirty,
    }


def comment_read(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    cell: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    worksheet = get_worksheet(session.workbook, sheet)

    try:
        cell_obj = worksheet[cell]
    except KeyError as exc:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell '{cell}' is not valid.",
        ) from exc

    if cell_obj.comment is None:
        raise XlsxAgentError(
            code="not_found",
            message=f"No comment found on cell '{cell}'.",
        )

    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "cell": cell,
        "text": cell_obj.comment.text,
        "author": cell_obj.comment.author,
    }


def comment_remove(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    cell: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    try:
        cell_obj = worksheet[cell]
    except KeyError as exc:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell '{cell}' is not valid.",
        ) from exc

    if cell_obj.comment is None:
        raise XlsxAgentError(
            code="not_found",
            message=f"No comment found on cell '{cell}'.",
        )

    cell_obj.comment = None

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "cell": cell,
        "removed": True,
        "dirty": session.dirty,
    }


def comment_list(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    worksheet = get_worksheet(session.workbook, sheet)

    comments = []
    for row in worksheet.iter_rows():
        for cell_obj in row:
            if cell_obj.comment is not None:
                comments.append({
                    "cell": cell_obj.coordinate,
                    "text": cell_obj.comment.text,
                    "author": cell_obj.comment.author,
                })

    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "comments": comments,
        "count": len(comments),
    }
