from __future__ import annotations

from copy import copy

from openpyxl.utils import range_boundaries

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.sheets import get_worksheet
from xlsx_agent.core.store import WorkbookStore


def range_style_set(
    store: WorkbookStore,
    *,
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
    vertical_alignment: str | None = None,
    wrap_text: bool | None = None,
    text_rotation: int | None = None,
    shrink_to_fit: bool | None = None,
    indent: int | None = None,
    border_style: str | None = None,
    border_color: str | None = None,
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    try:
        min_col, min_row, max_col, max_row = range_boundaries(range)
    except ValueError as exc:
        raise XlsxAgentError(
            code="invalid_range",
            message=f"Range '{range}' is invalid.",
        ) from exc

    from openpyxl.styles import Border, PatternFill, Side

    count = 0
    for row in worksheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font_name or font_size is not None or bold is not None or italic is not None or underline or font_color:
                font = copy(cell.font)
                if font_name:
                    font.name = font_name
                if font_size is not None:
                    font.sz = font_size
                if bold is not None:
                    font.b = bold
                if italic is not None:
                    font.i = italic
                if underline:
                    font.u = underline
                if font_color:
                    font.color = font_color
                cell.font = font

            if fill_color:
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

            if number_format:
                cell.number_format = number_format

            if (
                alignment
                or vertical_alignment
                or wrap_text is not None
                or text_rotation is not None
                or shrink_to_fit is not None
                or indent is not None
            ):
                cell_alignment = copy(cell.alignment)
                if alignment:
                    cell_alignment.horizontal = alignment
                if vertical_alignment:
                    cell_alignment.vertical = vertical_alignment
                if wrap_text is not None:
                    cell_alignment.wrap_text = wrap_text
                if text_rotation is not None:
                    cell_alignment.text_rotation = text_rotation
                if shrink_to_fit is not None:
                    cell_alignment.shrink_to_fit = shrink_to_fit
                if indent is not None:
                    cell_alignment.indent = indent
                cell.alignment = cell_alignment

            if border_style or border_color:
                style = None if border_style == "none" else border_style
                side = Side(style=style, color=border_color)
                cell.border = Border(left=side, right=side, top=side, bottom=side)

            count += 1

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "range": range,
        "styled_count": count,
        "dirty": session.dirty,
    }


def conditional_formatting_list(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
) -> dict[str, object]:
    session = store.get(workbook_id)
    worksheet = get_worksheet(session.workbook, sheet)

    rules = []
    for cf in worksheet.conditional_formatting:
        for rule in cf.rules:
            item = {
                "type": rule.type,
                "priority": rule.priority,
                "ranges": str(cf.sqref) if cf.sqref else None,
                "formula": getattr(rule, "formula", None),
                "operator": getattr(rule, "operator", None),
                "stop_if_true": getattr(rule, "stopIfTrue", None),
            }
            if rule.dxf:
                if rule.dxf.fill and rule.dxf.fill.fgColor:
                    item["fill_color"] = rule.dxf.fill.fgColor.rgb
                if rule.dxf.font and rule.dxf.font.color:
                    item["font_color"] = rule.dxf.font.color.rgb
            if rule.colorScale:
                item["color_scale_colors"] = [color.rgb for color in rule.colorScale.color]
            if rule.dataBar:
                item["data_bar_color"] = rule.dataBar.color.rgb
            if rule.iconSet:
                item["icon_style"] = rule.iconSet.iconSet
            rules.append(item)

    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "rules": rules,
        "count": len(rules),
    }


def conditional_formatting_add(
    store: WorkbookStore,
    *,
    workbook_id: str,
    sheet: str,
    range: str,
    type: str,
    formula: str | None = None,
    formula2: str | None = None,
    operator: str = "equal",
    priority: int = 1,
    fill_color: str | None = None,
    font_color: str | None = None,
    stop_if_true: bool | None = None,
    min_color: str = "FFF8696B",
    mid_color: str = "FFFFEB84",
    max_color: str = "FF63BE7B",
    data_bar_color: str = "FF638EC6",
    icon_style: str = "3TrafficLights1",
) -> dict[str, object]:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    worksheet = get_worksheet(session.workbook, sheet)

    if priority < 1:
        raise XlsxAgentError(
            code="invalid_input",
            message="priority must be greater than or equal to 1.",
        )

    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule
    from openpyxl.styles import Font, PatternFill

    fill = PatternFill(fill_type="solid", fgColor=fill_color) if fill_color else None
    font = Font(color=font_color) if font_color else None
    normalized_type = {
        "cellIs": "cell_is",
        "cell_is": "cell_is",
        "formula": "formula",
        "colorScale": "color_scale",
        "color_scale": "color_scale",
        "dataBar": "data_bar",
        "data_bar": "data_bar",
        "iconSet": "icon_set",
        "icon_set": "icon_set",
    }.get(type)

    if normalized_type == "cell_is":
        formulas = [value for value in (formula, formula2) if value is not None]
        rule = CellIsRule(
            operator=operator,
            formula=formulas,
            stopIfTrue=stop_if_true,
            fill=fill,
            font=font,
        )
    elif normalized_type == "formula":
        rule = FormulaRule(
            formula=[formula] if formula is not None else [],
            stopIfTrue=stop_if_true,
            fill=fill,
            font=font,
        )
    elif normalized_type == "color_scale":
        rule = ColorScaleRule(
            start_type="min",
            start_color=min_color,
            mid_type="percentile",
            mid_value=50,
            mid_color=mid_color,
            end_type="max",
            end_color=max_color,
        )
    elif normalized_type == "data_bar":
        rule = DataBarRule(
            start_type="min",
            end_type="max",
            color=data_bar_color,
        )
    elif normalized_type == "icon_set":
        rule = IconSetRule(
            icon_style=icon_style,
            type="percent",
            values=[0, 33, 67],
        )
    else:
        raise XlsxAgentError(
            code="invalid_input",
            message="type must be one of: cellIs, formula, colorScale, dataBar, or iconSet.",
        )

    worksheet.conditional_formatting.add(range, rule)
    rule.priority = priority

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "range": range,
        "type": rule.type,
        "priority": rule.priority,
        "dirty": session.dirty,
    }


def conditional_formatting_remove(
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
    for cf in list(worksheet.conditional_formatting):
        if str(cf.sqref) == range:
            worksheet.conditional_formatting.remove(cf)
            removed += 1

    session.dirty = True
    return {
        "workbook_id": workbook_id,
        "sheet": sheet,
        "removed_count": removed,
        "dirty": session.dirty,
    }
