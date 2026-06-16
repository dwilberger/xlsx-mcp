from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import typer
from pydantic import BaseModel, ValidationError

from xlsx_agent import __version__
from xlsx_agent.core.cell import cell_read, cell_write
from xlsx_agent.core.columns import columns_delete, columns_insert
from xlsx_agent.core.comments import comment_add, comment_list, comment_read, comment_remove
from xlsx_agent.core.data import (
    blanks_find,
    distinct_values,
    duplicates_find,
    errors_find,
    range_deduplicate,
    range_find,
    range_replace,
    range_sort,
)
from xlsx_agent.core.diff import sheet_diff, workbook_diff
from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.export import (
    csv_to_sheet,
    range_to_csv,
    sheet_to_csv,
    sheet_to_json,
    json_to_sheet,
)
from xlsx_agent.core.formula import formula_set
from xlsx_agent.core.freeze_panes import freeze_panes_clear, freeze_panes_set
from xlsx_agent.core.headers_footers import headers_footers_set
from xlsx_agent.core.images import image_add, image_list, image_remove
from xlsx_agent.core.formula_ops import (
    array_formula_list,
    formula_copy,
    formula_fill_down,
    formula_fill_right,
    formula_read,
    formula_to_values,
)
from xlsx_agent.core.merges import merge_cells, merged_cells_list, unmerge_cells
from xlsx_agent.core.names import (
    defined_name_create,
    defined_name_delete,
    defined_names_list,
    external_links_list,
)
from xlsx_agent.core.page_setup import page_margins_set, page_setup_set
from xlsx_agent.core.print_area import print_area_set, print_titles_set
from xlsx_agent.core.range import range_clear, range_write
from xlsx_agent.core.read import range_read
from xlsx_agent.core.rows import rows_delete, rows_insert, rows_write
from xlsx_agent.core.sheet_ops import sheet_copy, sheet_query, sheet_used_range_get
from xlsx_agent.core.sheets import sheet_create, sheet_delete, sheet_list, sheet_rename
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.style import (
    conditional_formatting_add,
    conditional_formatting_list,
    conditional_formatting_remove,
    range_style_set,
)
from xlsx_agent.core.tables import table_create, table_delete, table_list, table_resize
from xlsx_agent.core.validation_ops import (
    data_validation_add,
    data_validation_list,
    data_validation_remove,
)
from xlsx_agent.core.workbook import workbook_close, workbook_open, workbook_save
from xlsx_agent.core.workbook_ops import workbook_metadata
from xlsx_agent.core.write import cells_write, columns_modify, rows_append, rows_modify, table_write
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

app = typer.Typer(
    no_args_is_help=True,
    add_completion=False,
    help="Deterministic local Excel automation.",
)

store = WorkbookStore()


def _serialize(value: Any) -> Any:
    if isinstance(value, BaseModel):
        return value.model_dump()
    if isinstance(value, list):
        return [_serialize(item) for item in value]
    if isinstance(value, dict):
        return {key: _serialize(item) for key, item in value.items()}
    return value


def _print_json(payload: dict) -> None:
    typer.echo(json.dumps(payload, indent=2, sort_keys=True))


def _print_success(data: Any) -> None:
    _print_json(SuccessEnvelope(data=_serialize(data)).model_dump())


def _print_error(*, code: str, message: str, details: dict[str, Any] | None = None) -> None:
    _print_json(
        ErrorEnvelope(
            error=ErrorDetail(code=code, message=message, details=details or {}),
        ).model_dump()
    )


def _read_input_payload(input_path: Path | None) -> dict[str, Any]:
    if input_path is None:
        raise XlsxAgentError(
            code="invalid_input",
            message="The --input option is required for run commands.",
        )

    try:
        return json.loads(input_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise XlsxAgentError(
            code="file_not_found",
            message=f"Input file '{input_path}' was not found.",
        ) from exc
    except json.JSONDecodeError as exc:
        raise XlsxAgentError(
            code="invalid_input",
            message=f"Input file '{input_path}' does not contain valid JSON.",
            details={"line": exc.lineno, "column": exc.colno},
        ) from exc


def _dispatch_run_command(command: str, payload: dict[str, Any]) -> Any:
    if command == "workbook_open":
        request = WorkbookOpenRequest.model_validate(payload)
        return workbook_open(store, path=request.path, read_only=request.read_only)
    if command == "workbook_save":
        request = WorkbookSaveRequest.model_validate(payload)
        return workbook_save(store, workbook_id=request.workbook_id, path=request.path, overwrite=request.overwrite)
    if command == "workbook_close":
        request = WorkbookCloseRequest.model_validate(payload)
        return workbook_close(store, workbook_id=request.workbook_id, force=request.force)
    if command == "sheet_list":
        request = SheetListRequest.model_validate(payload)
        return sheet_list(store, workbook_id=request.workbook_id)
    if command == "sheet_create":
        request = SheetCreateRequest.model_validate(payload)
        return sheet_create(store, workbook_id=request.workbook_id, name=request.name, position=request.position)
    if command == "range_read":
        request = RangeReadRequest.model_validate(payload)
        return range_read(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, format=request.format)
    if command == "cells_write":
        request = CellsWriteRequest.model_validate(payload)
        return cells_write(store, workbook_id=request.workbook_id, sheet=request.sheet, cells=request.cells)
    if command == "rows_append":
        request = RowsAppendRequest.model_validate(payload)
        return rows_append(store, workbook_id=request.workbook_id, sheet=request.sheet, rows=request.rows)

    # Additional commands beyond the minimal workbook/read/write loop.
    if command == "sheet_delete":
        request = SheetDeleteRequest.model_validate(payload)
        return sheet_delete(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "sheet_rename":
        request = SheetRenameRequest.model_validate(payload)
        return sheet_rename(store, workbook_id=request.workbook_id, sheet=request.sheet, new_name=request.new_name)
    if command == "table_write":
        request = TableWriteRequest.model_validate(payload)
        return table_write(store, workbook_id=request.workbook_id, sheet=request.sheet, start_cell=request.start_cell, rows=request.rows, columns=request.columns, overwrite=request.overwrite)
    if command == "formula_set":
        request = FormulaWriteRequest.model_validate(payload)
        return formula_set(store, workbook_id=request.workbook_id, sheet=request.sheet, formulas=request.formulas)
    if command == "rows_modify":
        request = RowsModifyRequest.model_validate(payload)
        return rows_modify(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, rows=request.rows)
    if command == "columns_modify":
        request = ColumnsModifyRequest.model_validate(payload)
        return columns_modify(store, workbook_id=request.workbook_id, sheet=request.sheet, start_column=request.start_column, columns=request.columns)

    # Extended command surface.
    if command == "cell_read":
        request = CellReadRequest.model_validate(payload)
        return cell_read(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell)
    if command == "cell_write":
        request = CellWriteRequest.model_validate(payload)
        return cell_write(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell, value=request.value)
    if command == "range_write":
        request = RangeWriteRequest.model_validate(payload)
        return range_write(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, data=request.data)
    if command == "range_clear":
        request = RangeClearRequest.model_validate(payload)
        return range_clear(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "rows_insert":
        request = RowsInsertRequest.model_validate(payload)
        return rows_insert(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, count=request.count)
    if command == "rows_delete":
        request = RowsDeleteRequest.model_validate(payload)
        return rows_delete(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, count=request.count)
    if command == "rows_write":
        request = RowsWriteRequest.model_validate(payload)
        return rows_write(store, workbook_id=request.workbook_id, sheet=request.sheet, start_row=request.start_row, rows=request.rows)
    if command == "columns_insert":
        request = ColumnsInsertRequest.model_validate(payload)
        return columns_insert(store, workbook_id=request.workbook_id, sheet=request.sheet, start_column=request.start_column, count=request.count)
    if command == "columns_delete":
        request = ColumnsDeleteRequest.model_validate(payload)
        return columns_delete(store, workbook_id=request.workbook_id, sheet=request.sheet, start_column=request.start_column, count=request.count)
    if command == "sheet_copy":
        request = SheetCopyRequest.model_validate(payload)
        return sheet_copy(store, workbook_id=request.workbook_id, sheet=request.sheet, new_name=request.new_name, position=request.position)
    if command == "sheet_used_range_get":
        request = SheetUsedRangeGetRequest.model_validate(payload)
        return sheet_used_range_get(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "sheet_query":
        request = SheetQueryRequest.model_validate(payload)
        return sheet_query(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "range_find":
        request = RangeFindRequest.model_validate(payload)
        return range_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, value=request.value, match_type=request.match_type)
    if command == "range_replace":
        request = RangeReplaceRequest.model_validate(payload)
        return range_replace(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, find=request.find, replace_with=request.replace_with)
    if command == "distinct_values":
        request = DistinctValuesRequest.model_validate(payload)
        return distinct_values(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "duplicates_find":
        request = DuplicatesFindRequest.model_validate(payload)
        return duplicates_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "blanks_find":
        request = BlanksFindRequest.model_validate(payload)
        return blanks_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "errors_find":
        request = ErrorsFindRequest.model_validate(payload)
        return errors_find(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "range_sort":
        request = RangeSortRequest.model_validate(payload)
        return range_sort(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, sort_column=request.sort_column, ascending=request.ascending)
    if command == "range_deduplicate":
        request = RangeDeduplicateRequest.model_validate(payload)
        return range_deduplicate(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "formula_read":
        request = FormulaReadRequest.model_validate(payload)
        return formula_read(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell)
    if command == "formula_copy":
        request = FormulaCopyRequest.model_validate(payload)
        return formula_copy(store, workbook_id=request.workbook_id, sheet=request.sheet, source=request.source, target_range=request.target_range)
    if command == "formula_fill_down":
        request = FormulaFillDownRequest.model_validate(payload)
        return formula_fill_down(store, workbook_id=request.workbook_id, sheet=request.sheet, start_cell=request.start_cell, end_row=request.end_row)
    if command == "formula_fill_right":
        request = FormulaFillRightRequest.model_validate(payload)
        return formula_fill_right(store, workbook_id=request.workbook_id, sheet=request.sheet, start_cell=request.start_cell, end_column=request.end_column)
    if command == "formula_to_values":
        request = FormulaToValuesRequest.model_validate(payload)
        return formula_to_values(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "array_formula_list":
        request = ArrayFormulaListRequest.model_validate(payload)
        return array_formula_list(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "defined_names_list":
        request = DefinedNamesListRequest.model_validate(payload)
        return defined_names_list(store, workbook_id=request.workbook_id)
    if command == "defined_name_create":
        request = DefinedNameCreateRequest.model_validate(payload)
        return defined_name_create(store, workbook_id=request.workbook_id, name=request.name, value=request.value, sheet=request.sheet)
    if command == "defined_name_delete":
        request = DefinedNameDeleteRequest.model_validate(payload)
        return defined_name_delete(store, workbook_id=request.workbook_id, name=request.name)
    if command == "external_links_list":
        request = ExternalLinksListRequest.model_validate(payload)
        return external_links_list(store, workbook_id=request.workbook_id)
    if command == "table_list":
        request = TableListRequest.model_validate(payload)
        return table_list(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "table_create":
        request = TableCreateRequest.model_validate(payload)
        return table_create(store, workbook_id=request.workbook_id, sheet=request.sheet, name=request.name, range=request.range)
    if command == "table_resize":
        request = TableResizeRequest.model_validate(payload)
        return table_resize(store, workbook_id=request.workbook_id, sheet=request.sheet, name=request.name, range=request.range)
    if command == "table_delete":
        request = TableDeleteRequest.model_validate(payload)
        return table_delete(store, workbook_id=request.workbook_id, sheet=request.sheet, name=request.name)
    if command == "data_validation_list":
        request = DataValidationListRequest.model_validate(payload)
        return data_validation_list(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "data_validation_add":
        request = DataValidationAddRequest.model_validate(payload)
        return data_validation_add(store, workbook_id=request.workbook_id, sheet=request.sheet, type=request.type, range=request.range, operator=request.operator, formula1=request.formula1, formula2=request.formula2)
    if command == "data_validation_remove":
        request = DataValidationRemoveRequest.model_validate(payload)
        return data_validation_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "merged_cells_list":
        request = MergedCellsListRequest.model_validate(payload)
        return merged_cells_list(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "merge_cells":
        request = MergeCellsRequest.model_validate(payload)
        return merge_cells(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "unmerge_cells":
        request = UnmergeCellsRequest.model_validate(payload)
        return unmerge_cells(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "workbook_metadata":
        request = WorkbookMetadataRequest.model_validate(payload)
        return workbook_metadata(store, workbook_id=request.workbook_id)
    if command == "range_style_set":
        request = RangeStyleSetRequest.model_validate(payload)
        return range_style_set(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, font_name=request.font_name, font_size=request.font_size, bold=request.bold, italic=request.italic, underline=request.underline, fill_color=request.fill_color, font_color=request.font_color, number_format=request.number_format, alignment=request.alignment)
    if command == "conditional_formatting_list":
        request = ConditionalFormattingListRequest.model_validate(payload)
        return conditional_formatting_list(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "conditional_formatting_add":
        request = ConditionalFormattingAddRequest.model_validate(payload)
        return conditional_formatting_add(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, type=request.type, formula=request.formula, priority=request.priority)
    if command == "conditional_formatting_remove":
        request = ConditionalFormattingRemoveRequest.model_validate(payload)
        return conditional_formatting_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "sheet_to_csv":
        request = SheetToCsvRequest.model_validate(payload)
        return sheet_to_csv(store, workbook_id=request.workbook_id, sheet=request.sheet, delimiter=request.delimiter)
    if command == "range_to_csv":
        request = RangeToCsvRequest.model_validate(payload)
        return range_to_csv(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range, delimiter=request.delimiter)
    if command == "csv_to_sheet":
        request = CsvToSheetRequest.model_validate(payload)
        return csv_to_sheet(store, workbook_id=request.workbook_id, sheet=request.sheet, csv_data=request.csv_data, start_cell=request.start_cell, delimiter=request.delimiter)
    if command == "sheet_to_json":
        request = SheetToJsonRequest.model_validate(payload)
        return sheet_to_json(store, workbook_id=request.workbook_id, sheet=request.sheet, use_header_row=request.use_header_row)
    if command == "json_to_sheet":
        request = JsonToSheetRequest.model_validate(payload)
        return json_to_sheet(store, workbook_id=request.workbook_id, sheet=request.sheet, data=request.data, start_cell=request.start_cell)
    if command == "sheet_diff":
        request = SheetDiffRequest.model_validate(payload)
        return sheet_diff(store, workbook_id=request.workbook_id, sheet_a=request.sheet_a, sheet_b=request.sheet_b)
    if command == "workbook_diff":
        request = WorkbookDiffRequest.model_validate(payload)
        return workbook_diff(store, workbook_id=request.workbook_id, path_b=request.path_b)

    # Image and comment commands.
    if command == "image_add":
        request = ImageAddRequest.model_validate(payload)
        return image_add(store, workbook_id=request.workbook_id, sheet=request.sheet, path=request.path, anchor=request.anchor, width=request.width, height=request.height)
    if command == "image_list":
        request = ImageListRequest.model_validate(payload)
        return image_list(store, workbook_id=request.workbook_id, sheet=request.sheet)
    if command == "image_remove":
        request = ImageRemoveRequest.model_validate(payload)
        return image_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, index=request.index)
    if command == "comment_add":
        request = CommentAddRequest.model_validate(payload)
        return comment_add(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell, text=request.text, author=request.author)
    if command == "comment_read":
        request = CommentReadRequest.model_validate(payload)
        return comment_read(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell)
    if command == "comment_remove":
        request = CommentRemoveRequest.model_validate(payload)
        return comment_remove(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell)
    if command == "comment_list":
        request = CommentListRequest.model_validate(payload)
        return comment_list(store, workbook_id=request.workbook_id, sheet=request.sheet)

    # Layout and print commands.
    if command == "page_setup_set":
        request = PageSetupSetRequest.model_validate(payload)
        return page_setup_set(store, workbook_id=request.workbook_id, sheet=request.sheet, orientation=request.orientation, paper_size=request.paper_size, scale=request.scale, fit_to_page=request.fit_to_page, fit_to_width=request.fit_to_width, fit_to_height=request.fit_to_height)
    if command == "page_margins_set":
        request = PageMarginsSetRequest.model_validate(payload)
        return page_margins_set(store, workbook_id=request.workbook_id, sheet=request.sheet, left=request.left, right=request.right, top=request.top, bottom=request.bottom, header=request.header, footer=request.footer)
    if command == "print_area_set":
        request = PrintAreaSetRequest.model_validate(payload)
        return print_area_set(store, workbook_id=request.workbook_id, sheet=request.sheet, range=request.range)
    if command == "print_titles_set":
        request = PrintTitlesSetRequest.model_validate(payload)
        return print_titles_set(store, workbook_id=request.workbook_id, sheet=request.sheet, rows=request.rows, columns=request.columns)
    if command == "headers_footers_set":
        request = HeadersFootersSetRequest.model_validate(payload)
        return headers_footers_set(store, workbook_id=request.workbook_id, sheet=request.sheet, odd_header=request.odd_header, odd_footer=request.odd_footer, even_header=request.even_header, even_footer=request.even_footer, first_header=request.first_header, first_footer=request.first_footer)
    if command == "freeze_panes_set":
        request = FreezePanesSetRequest.model_validate(payload)
        return freeze_panes_set(store, workbook_id=request.workbook_id, sheet=request.sheet, cell=request.cell)
    if command == "freeze_panes_clear":
        request = FreezePanesClearRequest.model_validate(payload)
        return freeze_panes_clear(store, workbook_id=request.workbook_id, sheet=request.sheet)

    raise XlsxAgentError(
        code="unsupported_operation",
        message=f"Command '{command}' is not implemented yet.",
    )


@app.command("version")
def version() -> None:
    """Print the current package version."""
    _print_success({"version": __version__})


@app.command("run")
def run(command: str, input: Path | None = None) -> None:
    """Run a command using a JSON input payload."""
    try:
        payload = _read_input_payload(input)
        result = _dispatch_run_command(command, payload)
    except ValidationError as exc:
        _print_error(
            code="invalid_input",
            message="Input payload failed schema validation.",
            details={"errors": exc.errors(include_url=False)},
        )
        raise typer.Exit(code=1)
    except XlsxAgentError as exc:
        _print_error(code=exc.code, message=exc.message, details=exc.details)
        raise typer.Exit(code=1)

    _print_success(result)


def main() -> None:
    app()


if __name__ == "__main__":
    main()
