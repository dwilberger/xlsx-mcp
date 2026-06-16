from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def data_validation_list(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    worksheet = get_worksheet(session.workbook, sheet)

    validations = []
    for dv in worksheet.data_validations.dataValidation:
        validations.append({
            "type": dv.type,
            "operator": dv.operator,
            "formula1": dv.formula1,
            "formula2": dv.formula2,
            "ranges": str(dv.sqref) if dv.sqref else None,
            "allow_blank": dv.allowBlank,
            "showErrorMessage": dv.showErrorMessage,
            "showInputMessage": dv.showInputMessage,
        })

    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "validations": validations,
        "count": len(validations),
    }


def data_validation_add(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    type: str,
    range: str,
    operator: str | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type=type, operator=operator, formula1=formula1, formula2=formula2)
    dv.add(range)
    worksheet.add_data_validation(dv)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "type": type,
        "range": range,
        "dirty": session.dirty,
    }


def data_validation_remove(
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

    removed = 0
    for dv in list(worksheet.data_validations.dataValidation):
        if range in str(dv.sqref):
            worksheet.data_validations.dataValidation.remove(dv)
            removed += 1

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "removed_count": removed,
        "dirty": session.dirty,
    }
