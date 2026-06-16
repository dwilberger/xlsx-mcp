from __future__ import annotations

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def _check_read_only(session) -> None:
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )


def cell_read(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    cell: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    worksheet = get_worksheet(session.workbook, sheet)

    try:
        row, column = coordinate_to_tuple(cell)
    except ValueError as exc:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell address '{cell}' is invalid.",
        ) from exc

    if row < 1 or column < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell address '{cell}' is invalid.",
        )

    c = worksheet[cell.upper()]
    value = c.value
    data_type = c.data_type

    if value is None:
        typed = {"type": "empty"}
    elif data_type == "f":
        typed = {"type": "formula", "value": None, "formula": str(value)}
    elif data_type == "e":
        typed = {"type": "error", "value": str(value)}
    elif isinstance(value, bool):
        typed = {"type": "boolean", "value": value}
    elif isinstance(value, (int, float)):
        typed = {"type": "number", "value": value}
    else:
        typed = {"type": "string", "value": str(value)}

    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "cell": cell.upper(),
        "value": typed,
    }


def cell_write(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    cell: str,
    value: object,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    try:
        row, column = coordinate_to_tuple(cell)
    except ValueError as exc:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell address '{cell}' is invalid.",
        ) from exc

    if row < 1 or column < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Cell address '{cell}' is invalid.",
        )

    c = worksheet[cell.upper()]
    c.value = value
    if isinstance(value, str) and value.startswith("="):
        c.data_type = "s"

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "cell": cell.upper(),
        "dirty": session.dirty,
    }
