from __future__ import annotations

from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.schemas.common import SheetMetadata
from xlsx_agent.core.sheets import _sheet_metadata


def workbook_metadata(
    store: WorkbookStore,
    *,
    workbook_id: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    workbook = session.workbook

    sheets = [
        _sheet_metadata(workbook, ws, idx + 1).model_dump()
        for idx, ws in enumerate(workbook.worksheets)
    ]

    return {
        "workbook_id": workbook_id,
        "path": str(session.path),
        "read_only": session.read_only,
        "dirty": session.dirty,
        "sheet_count": len(workbook.worksheets),
        "sheet_names": workbook.sheetnames,
        "active_sheet": workbook.active.title if workbook.active else None,
        "sheets": sheets,
    }
