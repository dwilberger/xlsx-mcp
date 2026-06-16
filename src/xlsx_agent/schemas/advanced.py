"""Schemas for the extended command surface."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class CellReadRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str


class CellWriteRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str
    value: object = None


class RangeWriteRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    data: list[list[object]] = Field(default_factory=list)


class RangeClearRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class RowsInsertRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_row: int
    count: int = 1


class RowsDeleteRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_row: int
    count: int = 1


class RowsWriteRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_row: int
    rows: list[list[object] | dict[str, object]] = Field(default_factory=list)


class ColumnsInsertRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_column: int | str
    count: int = 1


class ColumnsDeleteRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_column: int | str
    count: int = 1


class SheetDeleteRequest(BaseModel):
    workbook_id: str
    sheet: str


class SheetRenameRequest(BaseModel):
    workbook_id: str
    sheet: str
    new_name: str


class SheetCopyRequest(BaseModel):
    workbook_id: str
    sheet: str
    new_name: str | None = None
    position: int | None = None


class SheetUsedRangeGetRequest(BaseModel):
    workbook_id: str
    sheet: str


class SheetQueryRequest(BaseModel):
    workbook_id: str
    sheet: str


class RangeFindRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    value: object = None
    match_type: str = "exact"


class RangeReplaceRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    find: str
    replace_with: str


class DistinctValuesRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class DuplicatesFindRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class BlanksFindRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class ErrorsFindRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class RangeSortRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    sort_column: int | str
    ascending: bool = True


class RangeDeduplicateRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class FormulaReadRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str


class FormulaWriteRequest(BaseModel):
    workbook_id: str
    sheet: str
    formulas: dict[str, str] = Field(default_factory=dict)


class FormulaCopyRequest(BaseModel):
    workbook_id: str
    sheet: str
    source: str
    target_range: str


class FormulaFillDownRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_cell: str
    end_row: int


class FormulaFillRightRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_cell: str
    end_column: int | str


class FormulaToValuesRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class ArrayFormulaListRequest(BaseModel):
    workbook_id: str
    sheet: str


class DefinedNamesListRequest(BaseModel):
    workbook_id: str


class DefinedNameCreateRequest(BaseModel):
    workbook_id: str
    name: str
    value: str
    sheet: str | None = None


class DefinedNameDeleteRequest(BaseModel):
    workbook_id: str
    name: str


class ExternalLinksListRequest(BaseModel):
    workbook_id: str


class TableListRequest(BaseModel):
    workbook_id: str
    sheet: str | None = None


class TableCreateRequest(BaseModel):
    workbook_id: str
    sheet: str
    name: str
    range: str


class TableResizeRequest(BaseModel):
    workbook_id: str
    sheet: str
    name: str
    range: str


class TableDeleteRequest(BaseModel):
    workbook_id: str
    sheet: str
    name: str


class DataValidationListRequest(BaseModel):
    workbook_id: str
    sheet: str


class DataValidationAddRequest(BaseModel):
    workbook_id: str
    sheet: str
    type: str
    range: str
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None


class DataValidationRemoveRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class MergedCellsListRequest(BaseModel):
    workbook_id: str
    sheet: str


class MergeCellsRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class UnmergeCellsRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class WorkbookMetadataRequest(BaseModel):
    workbook_id: str


class RangeStyleSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = None
    fill_color: str | None = None
    font_color: str | None = None
    number_format: str | None = None
    alignment: str | None = None


class ConditionalFormattingListRequest(BaseModel):
    workbook_id: str
    sheet: str


class ConditionalFormattingAddRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    type: str
    formula: str | None = None
    priority: int = 1


class ConditionalFormattingRemoveRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str


class SheetToCsvRequest(BaseModel):
    workbook_id: str
    sheet: str
    delimiter: str = ","


class RangeToCsvRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str
    delimiter: str = ","


class CsvToSheetRequest(BaseModel):
    workbook_id: str
    sheet: str
    csv_data: str
    start_cell: str = "A1"
    delimiter: str = ","


class SheetToJsonRequest(BaseModel):
    workbook_id: str
    sheet: str
    use_header_row: bool = True


class JsonToSheetRequest(BaseModel):
    workbook_id: str
    sheet: str
    data: list[dict[str, object] | list[object]] = Field(default_factory=list)
    start_cell: str = "A1"


class SheetDiffRequest(BaseModel):
    workbook_id: str
    sheet_a: str
    sheet_b: str


class WorkbookDiffRequest(BaseModel):
    workbook_id: str
    path_b: str


class TableWriteRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_cell: str
    rows: list[list[object] | dict[str, object]] = Field(default_factory=list)
    columns: list[str] | None = None
    overwrite: Literal["error", "replace"] = "error"


class RowsModifyRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_row: int
    rows: list[list[object] | dict[str, object]] = Field(default_factory=list)


class ColumnsModifyRequest(BaseModel):
    workbook_id: str
    sheet: str
    start_column: int | str
    columns: list[list[object] | dict[str, object]] = Field(default_factory=list)


class ImageAddRequest(BaseModel):
    workbook_id: str
    sheet: str
    path: str
    anchor: str
    width: int | None = None
    height: int | None = None


class ImageListRequest(BaseModel):
    workbook_id: str
    sheet: str


class ImageRemoveRequest(BaseModel):
    workbook_id: str
    sheet: str
    index: int


class CommentAddRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str
    text: str
    author: str = ""


class CommentReadRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str


class CommentRemoveRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str


class CommentListRequest(BaseModel):
    workbook_id: str
    sheet: str


class PageSetupSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    orientation: str | None = None
    paper_size: str | int | None = None
    scale: int | None = None
    fit_to_page: bool | None = None
    fit_to_width: int | None = None
    fit_to_height: int | None = None


class PageMarginsSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    left: float | None = None
    right: float | None = None
    top: float | None = None
    bottom: float | None = None
    header: float | None = None
    footer: float | None = None


class PrintAreaSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    range: str | None = None


class PrintTitlesSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    rows: str | None = None
    columns: str | None = None


class HeadersFootersSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    odd_header: str | None = None
    odd_footer: str | None = None
    even_header: str | None = None
    even_footer: str | None = None
    first_header: str | None = None
    first_footer: str | None = None


class FreezePanesSetRequest(BaseModel):
    workbook_id: str
    sheet: str
    cell: str = "A1"


class FreezePanesClearRequest(BaseModel):
    workbook_id: str
    sheet: str
