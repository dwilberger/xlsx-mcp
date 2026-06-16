from __future__ import annotations

from openpyxl.utils.cell import coordinate_to_tuple

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def formula_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    formulas: dict[str, str],
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    if not formulas:
        return {
            "workbook_id": workbook_id,
            "sheet": sheet,
            "written_count": 0,
            "dirty": session.dirty,
        }

    normalized_items = []
    for address, formula in formulas.items():
        if not isinstance(formula, str) or not formula.startswith("=") or len(formula) < 2:
            raise XlsxAgentError(
                code="invalid_input",
                message=f"Formula for cell '{address}' must be a string starting with '='.",
            )
        try:
            row, column = coordinate_to_tuple(address)
        except ValueError as exc:
            raise XlsxAgentError(
                code="invalid_input",
                message=f"Cell address '{address}' is invalid.",
            ) from exc
        if row < 1 or column < 1:
            raise XlsxAgentError(
                code="invalid_input",
                message=f"Cell address '{address}' is invalid.",
            )
        normalized_items.append((address.upper(), formula))

    for address, formula in normalized_items:
        cell = worksheet[address]
        cell.value = formula
        cell.data_type = "f"

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "written_count": len(normalized_items),
        "dirty": session.dirty,
    }
