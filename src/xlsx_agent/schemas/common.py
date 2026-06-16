from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class ErrorDetail(BaseModel):
    code: str
    message: str
    details: dict[str, Any] = Field(default_factory=dict)


class SuccessEnvelope(BaseModel):
    ok: Literal[True] = True
    data: Any = Field(default_factory=dict)


class ErrorEnvelope(BaseModel):
    ok: Literal[False] = False
    error: ErrorDetail


class WorkbookHandle(BaseModel):
    workbook_id: str
    path: str
    read_only: bool = False
    dirty: bool = False


class SheetMetadata(BaseModel):
    name: str
    position: int
    visibility: Literal["visible", "hidden", "very_hidden"]
    is_active: bool
    used_range: str | None = None
    max_row: int
    max_column: int


class CellValue(BaseModel):
    type: Literal["string", "number", "boolean", "date", "datetime", "empty", "formula", "error"]
    value: Any = None
    formula: str | None = None


class CellRead(BaseModel):
    address: str
    value: CellValue


class RangeReadResult(BaseModel):
    sheet: str
    requested_range: str
    resolved_range: str
    row_count: int
    column_count: int
    data: list[Any]
