from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class WorkbookOpenRequest(BaseModel):
    path: str
    read_only: bool = False


class WorkbookSaveRequest(BaseModel):
    workbook_id: str
    path: str | None = None
    overwrite: bool = False


class WorkbookCloseRequest(BaseModel):
    workbook_id: str
    force: bool = False


class SheetListRequest(BaseModel):
    workbook_id: str


class SheetCreateRequest(BaseModel):
    workbook_id: str
    name: str
    position: int | None = None


class RangeReadRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    format: Literal["matrix", "records", "cells"] = "cells"


class CellsWriteRequest(BaseModel):
    workbook_id: str
    sheet: str
    cells: dict[str, object] = Field(default_factory=dict)


class RowsAppendRequest(BaseModel):
    workbook_id: str
    sheet: str
    rows: list[list[object] | dict[str, object]] = Field(default_factory=list)
