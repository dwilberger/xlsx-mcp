from __future__ import annotations

from openpyxl.worksheet.header_footer import HeaderFooterItem

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def _build_item(text: str) -> HeaderFooterItem:
    item = HeaderFooterItem()
    item.text = text
    return item


def headers_footers_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    odd_header: str | None = None,
    odd_footer: str | None = None,
    even_header: str | None = None,
    even_footer: str | None = None,
    first_header: str | None = None,
    first_footer: str | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    if odd_header is not None:
        worksheet.oddHeader = _build_item(odd_header)
    if odd_footer is not None:
        worksheet.oddFooter = _build_item(odd_footer)
    if even_header is not None:
        worksheet.evenHeader = _build_item(even_header)
    if even_footer is not None:
        worksheet.evenFooter = _build_item(even_footer)
    if first_header is not None:
        worksheet.firstHeader = _build_item(first_header)
    if first_footer is not None:
        worksheet.firstFooter = _build_item(first_footer)

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "dirty": session.dirty,
    }
