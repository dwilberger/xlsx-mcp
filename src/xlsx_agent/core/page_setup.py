from __future__ import annotations

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore

PAPER_SIZES: dict[str, int] = {
    "letter": 1,
    "letter_small": 2,
    "tabloid": 3,
    "ledger": 4,
    "legal": 5,
    "statement": 6,
    "executive": 7,
    "a3": 8,
    "a4": 9,
    "a5": 10,
    "b4": 12,
    "b5": 13,
    "japanese_postcard": 11,
    "envelope_9": 19,
    "envelope_10": 20,
    "envelope_dl": 27,
    "envelope_c5": 28,
    "envelope_b5": 35,
    "envelope_b4": 36,
}


def _resolve_paper_size(paper_size: str | int) -> int:
    if isinstance(paper_size, int):
        return paper_size
    key = paper_size.lower().strip()
    if key not in PAPER_SIZES:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Unknown paper size '{paper_size}'. Supported: {', '.join(sorted(PAPER_SIZES.keys()))}.",
        )
    return PAPER_SIZES[key]


def page_setup_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    orientation: str | None = None,
    paper_size: str | int | None = None,
    scale: int | None = None,
    fit_to_page: bool | None = None,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    ps = worksheet.page_setup

    if orientation is not None:
        normalized = orientation.lower().strip()
        if normalized not in ("portrait", "landscape"):
            raise XlsxAgentError(
                code="invalid_input",
                message=f"Invalid orientation '{orientation}'. Must be 'portrait' or 'landscape'.",
            )
        ps.orientation = normalized

    if paper_size is not None:
        ps.paperSize = _resolve_paper_size(paper_size)

    if scale is not None:
        if not 1 <= scale <= 400:
            raise XlsxAgentError(
                code="invalid_input",
                message=f"Scale must be between 1 and 400, got {scale}.",
            )
        ps.scale = scale

    if fit_to_page is not None:
        ps.fitToPage = fit_to_page

    if fit_to_width is not None:
        ps.fitToWidth = fit_to_width

    if fit_to_height is not None:
        ps.fitToHeight = fit_to_height

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "orientation": ps.orientation,
        "paper_size": ps.paperSize,
        "scale": ps.scale,
        "fit_to_page": ps.fitToPage,
        "fit_to_width": ps.fitToWidth,
        "fit_to_height": ps.fitToHeight,
        "dirty": session.dirty,
    }


def page_margins_set(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    left: float | None = None,
    right: float | None = None,
    top: float | None = None,
    bottom: float | None = None,
    header: float | None = None,
    footer: float | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)
    margins = worksheet.page_margins

    if left is not None:
        margins.left = left
    if right is not None:
        margins.right = right
    if top is not None:
        margins.top = top
    if bottom is not None:
        margins.bottom = bottom
    if header is not None:
        margins.header = header
    if footer is not None:
        margins.footer = footer

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "left": margins.left,
        "right": margins.right,
        "top": margins.top,
        "bottom": margins.bottom,
        "header": margins.header,
        "footer": margins.footer,
        "dirty": session.dirty,
    }
