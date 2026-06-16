from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.store import WorkbookStore


def defined_names_list(
    store: WorkbookStore,
    *,
    workbook_id: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    workbook = session.workbook

    names = []
    for name in workbook.defined_names.definedName:
        names.append({
            "name": name.name,
            "value": name.attr_text,
            "local_sheet_id": name.localSheetId,
        })

    return {
        "workbook_id": workbook_id,
        "names": names,
        "count": len(names),
    }


def defined_name_create(
    store: WorkbookStore,
    *,
    workbook_id: str,
    name: str,
    value: str,
    sheet: str | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    workbook = session.workbook

    if name in [n.name for n in workbook.defined_names.definedName]:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Defined name '{name}' already exists.",
        )

    from openpyxl.workbook.defined_name import DefinedName
    new_name = DefinedName(name, attr_text=value)
    if sheet:
        new_name.local_sheet_id = workbook.sheetnames.index(sheet)
    workbook.defined_names.add(new_name)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "name": name,
        "value": value,
        "dirty": session.dirty,
    }


def defined_name_delete(
    store: WorkbookStore,
    *,
    workbook_id: str,
    name: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    workbook = session.workbook

    if name not in [n.name for n in workbook.defined_names.definedName]:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Defined name '{name}' does not exist.",
        )

    del workbook.defined_names[name]

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "deleted_name": name,
        "dirty": session.dirty,
    }


def external_links_list(
    store: WorkbookStore,
    *,
    workbook_id: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    workbook = session.workbook

    links = []
    if hasattr(workbook, '_external_links'):
        for link in workbook._external_links:
            links.append(str(link))

    return {
        "workbook_id": workbook_id,
        "links": links,
        "count": len(links),
    }
