from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ValidationError

from mcp.server.fastmcp import FastMCP

from xlsx_agent.core.cell import cell_read as core_cell_read
from xlsx_agent.core.cell import cell_write as core_cell_write
from xlsx_agent.core.columns import columns_delete as core_columns_delete
from xlsx_agent.core.columns import columns_insert as core_columns_insert
from xlsx_agent.core.comments import comment_add as core_comment_add
from xlsx_agent.core.comments import comment_list as core_comment_list
from xlsx_agent.core.comments import comment_read as core_comment_read
from xlsx_agent.core.comments import comment_remove as core_comment_remove
from xlsx_agent.core.data import blanks_find as core_blanks_find
from xlsx_agent.core.data import distinct_values as core_distinct_values
from xlsx_agent.core.data import duplicates_find as core_duplicates_find
from xlsx_agent.core.data import errors_find as core_errors_find
from xlsx_agent.core.data import range_deduplicate as core_range_deduplicate
from xlsx_agent.core.data import range_find as core_range_find
from xlsx_agent.core.data import range_replace as core_range_replace
from xlsx_agent.core.data import range_sort as core_range_sort
from xlsx_agent.core.diff import sheet_diff as core_sheet_diff
from xlsx_agent.core.diff import workbook_diff as core_workbook_diff
from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.export import csv_to_sheet as core_csv_to_sheet
from xlsx_agent.core.export import range_to_csv as core_range_to_csv
from xlsx_agent.core.export import sheet_to_csv as core_sheet_to_csv
from xlsx_agent.core.export import sheet_to_json as core_sheet_to_json
from xlsx_agent.core.export import json_to_sheet as core_json_to_sheet
from xlsx_agent.core.formula import formula_set as core_formula_set
from xlsx_agent.core.freeze_panes import freeze_panes_clear as core_freeze_panes_clear
from xlsx_agent.core.freeze_panes import freeze_panes_set as core_freeze_panes_set
from xlsx_agent.core.headers_footers import headers_footers_set as core_headers_footers_set
from xlsx_agent.core.images import image_add as core_image_add
from xlsx_agent.core.images import image_list as core_image_list
from xlsx_agent.core.images import image_remove as core_image_remove
from xlsx_agent.core.formula_ops import array_formula_list as core_array_formula_list
from xlsx_agent.core.formula_ops import formula_copy as core_formula_copy
from xlsx_agent.core.formula_ops import formula_fill_down as core_formula_fill_down
from xlsx_agent.core.formula_ops import formula_fill_right as core_formula_fill_right
from xlsx_agent.core.formula_ops import formula_read as core_formula_read
from xlsx_agent.core.formula_ops import formula_to_values as core_formula_to_values
from xlsx_agent.core.merges import merge_cells as core_merge_cells
from xlsx_agent.core.merges import merged_cells_list as core_merged_cells_list
from xlsx_agent.core.merges import unmerge_cells as core_unmerge_cells
from xlsx_agent.core.names import defined_name_create as core_defined_name_create
from xlsx_agent.core.names import defined_name_delete as core_defined_name_delete
from xlsx_agent.core.names import defined_names_list as core_defined_names_list
from xlsx_agent.core.names import external_links_list as core_external_links_list
from xlsx_agent.core.page_setup import page_margins_set as core_page_margins_set
from xlsx_agent.core.page_setup import page_setup_set as core_page_setup_set
from xlsx_agent.core.print_area import print_area_set as core_print_area_set
from xlsx_agent.core.print_area import print_titles_set as core_print_titles_set
from xlsx_agent.core.range import range_clear as core_range_clear
from xlsx_agent.core.range import range_write as core_range_write
from xlsx_agent.core.read import range_read as core_range_read
from xlsx_agent.core.rows import rows_delete as core_rows_delete
from xlsx_agent.core.rows import rows_insert as core_rows_insert
from xlsx_agent.core.rows import rows_write as core_rows_write
from xlsx_agent.core.sheet_ops import sheet_copy as core_sheet_copy
from xlsx_agent.core.sheet_ops import sheet_query as core_sheet_query
from xlsx_agent.core.sheet_ops import sheet_used_range_get as core_sheet_used_range_get
from xlsx_agent.core.sheets import sheet_create as core_sheet_create
from xlsx_agent.core.sheets import sheet_delete as core_sheet_delete
from xlsx_agent.core.sheets import sheet_list as core_sheet_list
from xlsx_agent.core.sheets import sheet_rename as core_sheet_rename
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.style import conditional_formatting_add as core_conditional_formatting_add
from xlsx_agent.core.style import conditional_formatting_list as core_conditional_formatting_list
from xlsx_agent.core.style import conditional_formatting_remove as core_conditional_formatting_remove
from xlsx_agent.core.style import range_style_set as core_range_style_set
from xlsx_agent.core.tables import table_create as core_table_create
from xlsx_agent.core.tables import table_delete as core_table_delete
from xlsx_agent.core.tables import table_list as core_table_list
from xlsx_agent.core.tables import table_resize as core_table_resize
from xlsx_agent.core.validation_ops import data_validation_add as core_data_validation_add
from xlsx_agent.core.validation_ops import data_validation_list as core_data_validation_list
from xlsx_agent.core.validation_ops import data_validation_remove as core_data_validation_remove
from xlsx_agent.core.workbook import workbook_close as core_workbook_close
from xlsx_agent.core.workbook import workbook_open as core_workbook_open
from xlsx_agent.core.workbook import workbook_save as core_workbook_save
from xlsx_agent.core.workbook_ops import workbook_metadata as core_workbook_metadata
from xlsx_agent.core.write import cells_write as core_cells_write
from xlsx_agent.core.write import columns_modify as core_columns_modify
from xlsx_agent.core.write import rows_append as core_rows_append
from xlsx_agent.core.write import rows_modify as core_rows_modify
from xlsx_agent.core.write import table_write as core_table_write
from xlsx_agent.schemas.advanced import (
    ArrayFormulaListRequest,
    BlanksFindRequest,
    CellReadRequest,
    CellWriteRequest,
    ColumnsDeleteRequest,
    ColumnsInsertRequest,
    ColumnsModifyRequest,
    CommentAddRequest,
    CommentListRequest,
    CommentReadRequest,
    CommentRemoveRequest,
    ConditionalFormattingAddRequest,
    ConditionalFormattingListRequest,
    ConditionalFormattingRemoveRequest,
    CsvToSheetRequest,
    DataValidationAddRequest,
    DataValidationListRequest,
    DataValidationRemoveRequest,
    DefinedNameCreateRequest,
    DefinedNameDeleteRequest,
    DefinedNamesListRequest,
    DistinctValuesRequest,
    DuplicatesFindRequest,
    ErrorsFindRequest,
    ExternalLinksListRequest,
    FormulaCopyRequest,
    FormulaFillDownRequest,
    FormulaFillRightRequest,
    FormulaReadRequest,
    FormulaToValuesRequest,
    FormulaWriteRequest,
    FreezePanesClearRequest,
    FreezePanesSetRequest,
    HeadersFootersSetRequest,
    ImageAddRequest,
    ImageListRequest,
    ImageRemoveRequest,
    JsonToSheetRequest,
    MergeCellsRequest,
    MergedCellsListRequest,
    PageMarginsSetRequest,
    PageSetupSetRequest,
    PrintAreaSetRequest,
    PrintTitlesSetRequest,
    RangeClearRequest,
    RangeDeduplicateRequest,
    RangeFindRequest,
    RangeReplaceRequest,
    RangeSortRequest,
    RangeStyleSetRequest,
    RangeToCsvRequest,
    RangeWriteRequest,
    RowsDeleteRequest,
    RowsInsertRequest,
    RowsModifyRequest,
    RowsWriteRequest,
    SheetCopyRequest,
    SheetDeleteRequest,
    SheetDiffRequest,
    SheetQueryRequest,
    SheetRenameRequest,
    SheetToJsonRequest,
    SheetToCsvRequest,
    SheetUsedRangeGetRequest,
    TableCreateRequest,
    TableDeleteRequest,
    TableListRequest,
    TableResizeRequest,
    TableWriteRequest,
    UnmergeCellsRequest,
    WorkbookDiffRequest,
    WorkbookMetadataRequest,
)
from xlsx_agent.schemas.common import ErrorEnvelope, ErrorDetail, SuccessEnvelope
from xlsx_agent.schemas.essentials import (
    CellsWriteRequest,
    RangeReadRequest,
    RowsAppendRequest,
    SheetCreateRequest,
    SheetListRequest,
    WorkbookCloseRequest,
    WorkbookOpenRequest,
    WorkbookSaveRequest,
)

mcp = FastMCP("xlsx-mcp")
store = WorkbookStore()


def _serialize(value: Any) -> Any:
    if isinstance(value, BaseModel):
        return value.model_dump()
    if isinstance(value, list):
        return [_serialize(item) for item in value]
    if isinstance(value, dict):
        return {key: _serialize(item) for key, item in value.items()}
    return value


def _success(data: Any) -> dict[str, Any]:
    return SuccessEnvelope(data=_serialize(data)).model_dump()


def _error(*, code: str, message: str, details: dict[str, Any] | None = None) -> dict[str, Any]:
    return ErrorEnvelope(
        error=ErrorDetail(code=code, message=message, details=details or {}),
    ).model_dump()


def _execute(payload: dict[str, Any], schema: type[BaseModel], callback) -> dict[str, Any]:
    try:
        request = schema.model_validate(payload)
        result = callback(request)
    except ValidationError as exc:
        return _error(
            code="invalid_input",
            message="Input payload failed schema validation.",
            details={"errors": exc.errors(include_url=False)},
        )
    except XlsxAgentError as exc:
        return _error(code=exc.code, message=exc.message, details=exc.details)

    return _success(result)


# --- Workbook lifecycle ---


@mcp.tool()
def workbook_open(path: str, read_only: bool = False) -> dict[str, Any]:
    """Open a workbook from disk and return a process-scoped handle."""
    return _execute(
        {"path": path, "read_only": read_only},
        WorkbookOpenRequest,
        lambda request: core_workbook_open(store, path=request.path, read_only=request.read_only),
    )


@mcp.tool()
def workbook_save(workbook_id: str, path: str | None = None, overwrite: bool = False) -> dict[str, Any]:
    """Save an opened workbook to its current path or a new path."""
    return _execute(
        {"workbook_id": workbook_id, "path": path, "overwrite": overwrite},
        WorkbookSaveRequest,
        lambda request: core_workbook_save(store, workbook_id=request.workbook_id, path=request.path, overwrite=request.overwrite),
    )


@mcp.tool()
def workbook_close(workbook_id: str, force: bool = False) -> dict[str, Any]:
    """Close an opened workbook and optionally discard unsaved changes."""
    return _execute(
        {"workbook_id": workbook_id, "force": force},
        WorkbookCloseRequest,
        lambda request: core_workbook_close(store, workbook_id=request.workbook_id, force=request.force),
    )


@mcp.tool()
def workbook_metadata(workbook_id: str) -> dict[str, Any]:
    """Get metadata about an opened workbook including all sheet summaries."""
    return _execute(
        {"workbook_id": workbook_id},
        WorkbookMetadataRequest,
        lambda request: core_workbook_metadata(store, workbook_id=request.workbook_id),
    )


# --- Sheet operations ---


@mcp.tool()
def sheet_list(workbook_id: str) -> dict[str, Any]:
    """List sheet metadata for an opened workbook."""
    return _execute(
        {"workbook_id": workbook_id},
        SheetListRequest,
        lambda request: core_sheet_list(store, workbook_id=request.workbook_id),
    )


@mcp.tool()
def sheet_create(workbook_id: str, name: str, position: int | None = None) -> dict[str, Any]:
    """Create a sheet in an opened workbook."""
    return _execute(
        {"workbook_id": workbook_id, "name": name, "position": position},
        SheetCreateRequest,
        lambda request: core_sheet_create(store, workbook_id=request.workbook_id, name=request.name, position=request.position),
    )


@mcp.tool()
def sheet_delete(workbook_id: str, sheet: str) -> dict[str, Any]:
    """Delete a sheet from an opened workbook."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        SheetDeleteRequest,
        lambda request: core_sheet_delete(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def sheet_rename(workbook_id: str, sheet: str, new_name: str) -> dict[str, Any]:
    """Rename a sheet in an opened workbook."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "new_name": new_name},
        SheetRenameRequest,
        lambda request: core_sheet_rename(store, workbook_id=request.workbook_id, sheet=request.sheet, new_name=request.new_name),
    )


@mcp.tool()
def sheet_copy(workbook_id: str, sheet: str, new_name: str | None = None, position: int | None = None) -> dict[str, Any]:
    """Copy a sheet in an opened workbook."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "new_name": new_name, "position": position},
        SheetCopyRequest,
        lambda request: core_sheet_copy(store, workbook_id=request.workbook_id, sheet=request.sheet, new_name=request.new_name, position=request.position),
    )


@mcp.tool()
def sheet_used_range_get(workbook_id: str, sheet: str) -> dict[str, Any]:
    """Get the used range and dimensions of a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        SheetUsedRangeGetRequest,
        lambda request: core_sheet_used_range_get(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def sheet_query(workbook_id: str, sheet: str) -> dict[str, Any]:
    """Query detailed metadata about a specific sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        SheetQueryRequest,
        lambda request: core_sheet_query(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


# --- Cell operations ---


@mcp.tool()
def cell_read(workbook_id: str, sheet: str, cell: str) -> dict[str, Any]:
    """Read a single cell value with type information."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell},
        CellReadRequest,
        lambda request: core_cell_read(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell),
    )


@mcp.tool()
def cell_write(workbook_id: str, sheet: str, cell: str, value: object = None) -> dict[str, Any]:
    """Write a value to a single cell."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell, "value": value},
        CellWriteRequest,
        lambda request: core_cell_write(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell, value=request.value),
    )


# --- Range operations ---


@mcp.tool()
def range_read(workbook_id: str, sheet: str, range: str, format: str = "cells") -> dict[str, Any]:
    """Read a range from a sheet using cells, matrix, or records format."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "format": format},
        RangeReadRequest,
        lambda request: core_range_read(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, format=request.format),
    )


@mcp.tool()
def range_write(workbook_id: str, sheet: str, range: str, data: list[list[object]]) -> dict[str, Any]:
    """Write a 2D array of data to a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "data": data},
        RangeWriteRequest,
        lambda request: core_range_write(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, data=request.data),
    )


@mcp.tool()
def range_clear(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Clear all values in a range while preserving formatting."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        RangeClearRequest,
        lambda request: core_range_clear(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def range_find(workbook_id: str, sheet: str, range: str, value: object = None, match_type: str = "exact") -> dict[str, Any]:
    """Find cells matching a value in a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "value": value, "match_type": match_type},
        RangeFindRequest,
        lambda request: core_range_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, value=request.value, match_type=request.match_type),
    )


@mcp.tool()
def range_replace(workbook_id: str, sheet: str, range: str, find: str, replace_with: str) -> dict[str, Any]:
    """Replace text in all cells within a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "find": find, "replace_with": replace_with},
        RangeReplaceRequest,
        lambda request: core_range_replace(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, find=request.find, replace_with=request.replace_with),
    )


@mcp.tool()
def range_sort(workbook_id: str, sheet: str, range: str, sort_column: int | str, ascending: bool = True) -> dict[str, Any]:
    """Sort rows in a range by a column."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "sort_column": sort_column, "ascending": ascending},
        RangeSortRequest,
        lambda request: core_range_sort(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, sort_column=request.sort_column, ascending=request.ascending),
    )


@mcp.tool()
def range_deduplicate(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Remove duplicate rows from a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        RangeDeduplicateRequest,
        lambda request: core_range_deduplicate(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def range_style_set(
    workbook_id: str,
    sheet: str,
    range: str,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: str | None = None,
    fill_color: str | None = None,
    font_color: str | None = None,
    number_format: str | None = None,
    alignment: str | None = None,
) -> dict[str, Any]:
    """Apply font and cell styling to a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "font_name": font_name, "font_size": font_size, "bold": bold, "italic": italic, "underline": underline, "fill_color": fill_color, "font_color": font_color, "number_format": number_format, "alignment": alignment},
        RangeStyleSetRequest,
        lambda request: core_range_style_set(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, font_name=request.font_name, font_size=request.font_size, bold=request.bold, italic=request.italic, underline=request.underline, fill_color=request.fill_color, font_color=request.font_color, number_format=request.number_format, alignment=request.alignment),
    )


# --- Row operations ---


@mcp.tool()
def cells_write(workbook_id: str, sheet: str, cells: dict[str, object]) -> dict[str, Any]:
    """Write explicit literal values to specific cells."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cells": cells},
        CellsWriteRequest,
        lambda request: core_cells_write(store, workbook_id=request.workbook_id, sheet=request.sheet, cells=request.cells),
    )


@mcp.tool()
def rows_append(workbook_id: str, sheet: str, rows: list[list[object] | dict[str, object]]) -> dict[str, Any]:
    """Append array rows or object rows to the end of a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "rows": rows},
        RowsAppendRequest,
        lambda request: core_rows_append(store, workbook_id=request.workbook_id, sheet=request.sheet, rows=request.rows),
    )


@mcp.tool()
def rows_modify(workbook_id: str, sheet: str, start_row: int, rows: list[list[object] | dict[str, object]]) -> dict[str, Any]:
    """Overwrite existing rows at a specific row number (alias for rows_write)."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_row": start_row, "rows": rows},
        RowsModifyRequest,
        lambda request: core_rows_modify(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, rows=request.rows),
    )


@mcp.tool()
def rows_write(workbook_id: str, sheet: str, start_row: int, rows: list[list[object] | dict[str, object]]) -> dict[str, Any]:
    """Overwrite existing rows starting at a specific row number."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_row": start_row, "rows": rows},
        RowsWriteRequest,
        lambda request: core_rows_write(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, rows=request.rows),
    )


@mcp.tool()
def rows_insert(workbook_id: str, sheet: str, start_row: int, count: int = 1) -> dict[str, Any]:
    """Insert blank rows, shifting existing rows down."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_row": start_row, "count": count},
        RowsInsertRequest,
        lambda request: core_rows_insert(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, count=request.count),
    )


@mcp.tool()
def rows_delete(workbook_id: str, sheet: str, start_row: int, count: int = 1) -> dict[str, Any]:
    """Delete rows, shifting remaining rows up."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_row": start_row, "count": count},
        RowsDeleteRequest,
        lambda request: core_rows_delete(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, count=request.count),
    )


@mcp.tool()
def table_write(
    workbook_id: str,
    sheet: str,
    start_cell: str,
    rows: list[list[object] | dict[str, object]],
    columns: list[str] | None = None,
    overwrite: str = "error",
) -> dict[str, Any]:
    """Write a table of data starting at a specific cell."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_cell": start_cell, "rows": rows, "columns": columns, "overwrite": overwrite},
        TableWriteRequest,
        lambda request: core_table_write(store, workbook_id=request.workbook_id, sheet=request.sheet, start_cell=request.start_cell, rows=request.rows, columns=request.columns, overwrite=request.overwrite),
    )


# --- Column operations ---


@mcp.tool()
def columns_insert(workbook_id: str, sheet: str, start_column: int | str, count: int = 1) -> dict[str, Any]:
    """Insert blank columns, shifting existing columns right."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_column": start_column, "count": count},
        ColumnsInsertRequest,
        lambda request: core_columns_insert(store, workbook_id=request.workbook_id, sheet=request.sheet, start_column=request.start_column, count=request.count),
    )


@mcp.tool()
def columns_delete(workbook_id: str, sheet: str, start_column: int | str, count: int = 1) -> dict[str, Any]:
    """Delete columns, shifting remaining columns left."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_column": start_column, "count": count},
        ColumnsDeleteRequest,
        lambda request: core_columns_delete(store, workbook_id=request.workbook_id, sheet=request.sheet, start_column=request.start_column, count=request.count),
    )


@mcp.tool()
def columns_modify(workbook_id: str, sheet: str, start_column: int | str, columns: list[list[object] | dict[str, object]]) -> dict[str, Any]:
    """Overwrite existing columns starting at a specific column."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_column": start_column, "columns": columns},
        ColumnsModifyRequest,
        lambda request: core_columns_modify(store, workbook_id=request.workbook_id, sheet=request.sheet, start_column=request.start_column, columns=request.columns),
    )


# --- Data query operations ---


@mcp.tool()
def distinct_values(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Get distinct values from a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        DistinctValuesRequest,
        lambda request: core_distinct_values(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def duplicates_find(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Find duplicate values in a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        DuplicatesFindRequest,
        lambda request: core_duplicates_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def blanks_find(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Find blank cells in a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        BlanksFindRequest,
        lambda request: core_blanks_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def errors_find(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Find cells with errors in a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        ErrorsFindRequest,
        lambda request: core_errors_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


# --- Formula operations ---


@mcp.tool()
def formula_set(workbook_id: str, sheet: str, formulas: dict[str, str]) -> dict[str, Any]:
    """Set formulas in cells. Each formula must start with '='."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "formulas": formulas},
        FormulaWriteRequest,
        lambda request: core_formula_set(store, workbook_id=request.workbook_id, sheet=request.sheet, formulas=request.formulas),
    )


@mcp.tool()
def formula_read(workbook_id: str, sheet: str, cell: str) -> dict[str, Any]:
    """Read the formula from a cell."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell},
        FormulaReadRequest,
        lambda request: core_formula_read(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell),
    )


@mcp.tool()
def formula_copy(workbook_id: str, sheet: str, source: str, target_range: str) -> dict[str, Any]:
    """Copy a formula from one cell to a range of cells."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "source": source, "target_range": target_range},
        FormulaCopyRequest,
        lambda request: core_formula_copy(store, workbook_id=request.workbook_id, sheet=request.sheet, source=request.source, target_range=request.target_range),
    )


@mcp.tool()
def formula_fill_down(workbook_id: str, sheet: str, start_cell: str, end_row: int) -> dict[str, Any]:
    """Fill a formula down a column from start_cell to end_row."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_cell": start_cell, "end_row": end_row},
        FormulaFillDownRequest,
        lambda request: core_formula_fill_down(store, workbook_id=request.workbook_id, sheet=request.sheet, start_cell=request.start_cell, end_row=request.end_row),
    )


@mcp.tool()
def formula_fill_right(workbook_id: str, sheet: str, start_cell: str, end_column: int | str) -> dict[str, Any]:
    """Fill a formula right across a row from start_cell to end_column."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "start_cell": start_cell, "end_column": end_column},
        FormulaFillRightRequest,
        lambda request: core_formula_fill_right(store, workbook_id=request.workbook_id, sheet=request.sheet, start_cell=request.start_cell, end_column=request.end_column),
    )


@mcp.tool()
def formula_to_values(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Convert formulas in a range to their cached values."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        FormulaToValuesRequest,
        lambda request: core_formula_to_values(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def array_formula_list(workbook_id: str, sheet: str) -> dict[str, Any]:
    """List array formulas in a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        ArrayFormulaListRequest,
        lambda request: core_array_formula_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


# --- Name operations ---


@mcp.tool()
def defined_names_list(workbook_id: str) -> dict[str, Any]:
    """List all defined names in the workbook."""
    return _execute(
        {"workbook_id": workbook_id},
        DefinedNamesListRequest,
        lambda request: core_defined_names_list(store, workbook_id=request.workbook_id),
    )


@mcp.tool()
def defined_name_create(workbook_id: str, name: str, value: str, sheet: str | None = None) -> dict[str, Any]:
    """Create a defined name (named range)."""
    return _execute(
        {"workbook_id": workbook_id, "name": name, "value": value, "sheet": sheet},
        DefinedNameCreateRequest,
        lambda request: core_defined_name_create(store, workbook_id=request.workbook_id, name=request.name, value=request.value, sheet=request.sheet),
    )


@mcp.tool()
def defined_name_delete(workbook_id: str, name: str) -> dict[str, Any]:
    """Delete a defined name."""
    return _execute(
        {"workbook_id": workbook_id, "name": name},
        DefinedNameDeleteRequest,
        lambda request: core_defined_name_delete(store, workbook_id=request.workbook_id, name=request.name),
    )


@mcp.tool()
def external_links_list(workbook_id: str) -> dict[str, Any]:
    """List external links in the workbook."""
    return _execute(
        {"workbook_id": workbook_id},
        ExternalLinksListRequest,
        lambda request: core_external_links_list(store, workbook_id=request.workbook_id),
    )


# --- Table operations ---


@mcp.tool()
def table_list(workbook_id: str, sheet: str | None = None) -> dict[str, Any]:
    """List tables in the workbook or a specific sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        TableListRequest,
        lambda request: core_table_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def table_create(workbook_id: str, sheet: str, name: str, range: str) -> dict[str, Any]:
    """Create an Excel table from a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "name": name, "range": range},
        TableCreateRequest,
        lambda request: core_table_create(store, workbook_id=request.workbook_id, sheet=request.sheet, name=request.name, range=request.range),
    )


@mcp.tool()
def table_resize(workbook_id: str, sheet: str, name: str, range: str) -> dict[str, Any]:
    """Resize an existing table to a new range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "name": name, "range": range},
        TableResizeRequest,
        lambda request: core_table_resize(store, workbook_id=request.workbook_id, sheet=request.sheet, name=request.name, range=request.range),
    )


@mcp.tool()
def table_delete(workbook_id: str, sheet: str, name: str) -> dict[str, Any]:
    """Delete a table from a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "name": name},
        TableDeleteRequest,
        lambda request: core_table_delete(store, workbook_id=request.workbook_id, sheet=request.sheet, name=request.name),
    )


# --- Validation operations ---


@mcp.tool()
def data_validation_list(workbook_id: str, sheet: str) -> dict[str, Any]:
    """List data validation rules in a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        DataValidationListRequest,
        lambda request: core_data_validation_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def data_validation_add(
    workbook_id: str,
    sheet: str,
    type: str,
    range: str,
    operator: str | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
) -> dict[str, Any]:
    """Add a data validation rule to a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "type": type, "range": range, "operator": operator, "formula1": formula1, "formula2": formula2},
        DataValidationAddRequest,
        lambda request: core_data_validation_add(store, workbook_id=request.workbook_id, sheet=request.sheet, type=request.type, range=request.range, operator=request.operator, formula1=request.formula1, formula2=request.formula2),
    )


@mcp.tool()
def data_validation_remove(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Remove data validation rules from a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        DataValidationRemoveRequest,
        lambda request: core_data_validation_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


# --- Merge operations ---


@mcp.tool()
def merged_cells_list(workbook_id: str, sheet: str) -> dict[str, Any]:
    """List all merged cell ranges in a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        MergedCellsListRequest,
        lambda request: core_merged_cells_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def merge_cells(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Merge cells in a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        MergeCellsRequest,
        lambda request: core_merge_cells(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def unmerge_cells(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Unmerge cells in a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        UnmergeCellsRequest,
        lambda request: core_unmerge_cells(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


# --- Conditional formatting ---


@mcp.tool()
def conditional_formatting_list(workbook_id: str, sheet: str) -> dict[str, Any]:
    """List conditional formatting rules in a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        ConditionalFormattingListRequest,
        lambda request: core_conditional_formatting_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def conditional_formatting_add(workbook_id: str, sheet: str, range: str, type: str, formula: str | None = None, priority: int = 1) -> dict[str, Any]:
    """Add a conditional formatting rule to a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "type": type, "formula": formula, "priority": priority},
        ConditionalFormattingAddRequest,
        lambda request: core_conditional_formatting_add(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, type=request.type, formula=request.formula, priority=request.priority),
    )


@mcp.tool()
def conditional_formatting_remove(workbook_id: str, sheet: str, range: str) -> dict[str, Any]:
    """Remove conditional formatting rules from a range."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        ConditionalFormattingRemoveRequest,
        lambda request: core_conditional_formatting_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


# --- Export operations ---


@mcp.tool()
def sheet_to_csv(workbook_id: str, sheet: str, delimiter: str = ",") -> dict[str, Any]:
    """Export an entire sheet as CSV."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "delimiter": delimiter},
        SheetToCsvRequest,
        lambda request: core_sheet_to_csv(store, workbook_id=request.workbook_id, sheet=request.sheet, delimiter=request.delimiter),
    )


@mcp.tool()
def range_to_csv(workbook_id: str, sheet: str, range: str, delimiter: str = ",") -> dict[str, Any]:
    """Export a range as CSV."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range, "delimiter": delimiter},
        RangeToCsvRequest,
        lambda request: core_range_to_csv(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, delimiter=request.delimiter),
    )


@mcp.tool()
def csv_to_sheet(workbook_id: str, sheet: str, csv_data: str, start_cell: str = "A1", delimiter: str = ",") -> dict[str, Any]:
    """Import CSV data into a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "csv_data": csv_data, "start_cell": start_cell, "delimiter": delimiter},
        CsvToSheetRequest,
        lambda request: core_csv_to_sheet(store, workbook_id=request.workbook_id, sheet=request.sheet, csv_data=request.csv_data, start_cell=request.start_cell, delimiter=request.delimiter),
    )


@mcp.tool()
def sheet_to_json(workbook_id: str, sheet: str, use_header_row: bool = True) -> dict[str, Any]:
    """Export a sheet as JSON (array of objects or arrays)."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "use_header_row": use_header_row},
        SheetToJsonRequest,
        lambda request: core_sheet_to_json(store, workbook_id=request.workbook_id, sheet=request.sheet, use_header_row=request.use_header_row),
    )


@mcp.tool()
def json_to_sheet(workbook_id: str, sheet: str, data: list[dict[str, object] | list[object]], start_cell: str = "A1") -> dict[str, Any]:
    """Import JSON data into a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "data": data, "start_cell": start_cell},
        JsonToSheetRequest,
        lambda request: core_json_to_sheet(store, workbook_id=request.workbook_id, sheet=request.sheet, data=request.data, start_cell=request.start_cell),
    )


# --- Diff operations ---


@mcp.tool()
def sheet_diff(workbook_id: str, sheet_a: str, sheet_b: str) -> dict[str, Any]:
    """Compare two sheets in the same workbook and find differences."""
    return _execute(
        {"workbook_id": workbook_id, "sheet_a": sheet_a, "sheet_b": sheet_b},
        SheetDiffRequest,
        lambda request: core_sheet_diff(store, workbook_id=request.workbook_id, sheet_a=request.sheet_a, sheet_b=request.sheet_b),
    )


@mcp.tool()
def workbook_diff(workbook_id: str, path_b: str) -> dict[str, Any]:
    """Compare the current workbook with another workbook file."""
    return _execute(
        {"workbook_id": workbook_id, "path_b": path_b},
        WorkbookDiffRequest,
        lambda request: core_workbook_diff(store, workbook_id=request.workbook_id, path_b=request.path_b),
    )


# --- Image operations ---


@mcp.tool()
def image_add(workbook_id: str, sheet: str, path: str, anchor: str, width: int | None = None, height: int | None = None) -> dict[str, Any]:
    """Insert an image into a sheet at a cell anchor."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "path": path, "anchor": anchor, "width": width, "height": height},
        ImageAddRequest,
        lambda request: core_image_add(store, workbook_id=request.workbook_id, sheet=request.sheet, path=request.path, anchor=request.anchor, width=request.width, height=request.height),
    )


@mcp.tool()
def image_list(workbook_id: str, sheet: str) -> dict[str, Any]:
    """List images in a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        ImageListRequest,
        lambda request: core_image_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


@mcp.tool()
def image_remove(workbook_id: str, sheet: str, index: int) -> dict[str, Any]:
    """Remove an image from a sheet by index."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "index": index},
        ImageRemoveRequest,
        lambda request: core_image_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, index=request.index),
    )


# --- Comment operations ---


@mcp.tool()
def comment_add(workbook_id: str, sheet: str, cell: str, text: str, author: str = "") -> dict[str, Any]:
    """Add a comment to a cell."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell, "text": text, "author": author},
        CommentAddRequest,
        lambda request: core_comment_add(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell, text=request.text, author=request.author),
    )


@mcp.tool()
def comment_read(workbook_id: str, sheet: str, cell: str) -> dict[str, Any]:
    """Read a comment from a cell."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell},
        CommentReadRequest,
        lambda request: core_comment_read(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell),
    )


@mcp.tool()
def comment_remove(workbook_id: str, sheet: str, cell: str) -> dict[str, Any]:
    """Remove a comment from a cell."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell},
        CommentRemoveRequest,
        lambda request: core_comment_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell),
    )


@mcp.tool()
def comment_list(workbook_id: str, sheet: str) -> dict[str, Any]:
    """List all comments in a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        CommentListRequest,
        lambda request: core_comment_list(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


# --- Layout and Print ---


@mcp.tool()
def page_setup_set(
    workbook_id: str,
    sheet: str,
    orientation: str | None = None,
    paper_size: str | int | None = None,
    scale: int | None = None,
    fit_to_page: bool | None = None,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
) -> dict[str, Any]:
    """Set page orientation, paper size, scale, and fit-to-page options."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "orientation": orientation, "paper_size": paper_size, "scale": scale, "fit_to_page": fit_to_page, "fit_to_width": fit_to_width, "fit_to_height": fit_to_height},
        PageSetupSetRequest,
        lambda request: core_page_setup_set(store, workbook_id=request.workbook_id, sheet=request.sheet, orientation=request.orientation, paper_size=request.paper_size, scale=request.scale, fit_to_page=request.fit_to_page, fit_to_width=request.fit_to_width, fit_to_height=request.fit_to_height),
    )


@mcp.tool()
def page_margins_set(
    workbook_id: str,
    sheet: str,
    left: float | None = None,
    right: float | None = None,
    top: float | None = None,
    bottom: float | None = None,
    header: float | None = None,
    footer: float | None = None,
) -> dict[str, Any]:
    """Set page margins for printing."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "left": left, "right": right, "top": top, "bottom": bottom, "header": header, "footer": footer},
        PageMarginsSetRequest,
        lambda request: core_page_margins_set(store, workbook_id=request.workbook_id, sheet=request.sheet, left=request.left, right=request.right, top=request.top, bottom=request.bottom, header=request.header, footer=request.footer),
    )


@mcp.tool()
def print_area_set(workbook_id: str, sheet: str, range: str | None = None) -> dict[str, Any]:
    """Set or clear the print area for a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "range": range},
        PrintAreaSetRequest,
        lambda request: core_print_area_set(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range),
    )


@mcp.tool()
def print_titles_set(workbook_id: str, sheet: str, rows: str | None = None, columns: str | None = None) -> dict[str, Any]:
    """Set repeating rows and/or columns for printing."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "rows": rows, "columns": columns},
        PrintTitlesSetRequest,
        lambda request: core_print_titles_set(store, workbook_id=request.workbook_id, sheet=request.sheet, rows=request.rows, columns=request.columns),
    )


@mcp.tool()
def headers_footers_set(
    workbook_id: str,
    sheet: str,
    odd_header: str | None = None,
    odd_footer: str | None = None,
    even_header: str | None = None,
    even_footer: str | None = None,
    first_header: str | None = None,
    first_footer: str | None = None,
) -> dict[str, Any]:
    """Configure header and footer content for printing."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "odd_header": odd_header, "odd_footer": odd_footer, "even_header": even_header, "even_footer": even_footer, "first_header": first_header, "first_footer": first_footer},
        HeadersFootersSetRequest,
        lambda request: core_headers_footers_set(store, workbook_id=request.workbook_id, sheet=request.sheet, odd_header=request.odd_header, odd_footer=request.odd_footer, even_header=request.even_header, even_footer=request.even_footer, first_header=request.first_header, first_footer=request.first_footer),
    )


@mcp.tool()
def freeze_panes_set(workbook_id: str, sheet: str, cell: str = "A1") -> dict[str, Any]:
    """Set freeze panes at a cell reference."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet, "cell": cell},
        FreezePanesSetRequest,
        lambda request: core_freeze_panes_set(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell),
    )


@mcp.tool()
def freeze_panes_clear(workbook_id: str, sheet: str) -> dict[str, Any]:
    """Remove freeze panes from a sheet."""
    return _execute(
        {"workbook_id": workbook_id, "sheet": sheet},
        FreezePanesClearRequest,
        lambda request: core_freeze_panes_clear(store, workbook_id=request.workbook_id, sheet=request.sheet),
    )


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
