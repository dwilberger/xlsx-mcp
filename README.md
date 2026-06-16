# xlsx-mcp

**Your AI agent needs to read a spreadsheet. Now what?**

`xlsx-mcp` gives coding agents a clean, deterministic way to open, read, edit, and save `.xlsx` files -- no `openpyxl` gymnastics, no surprises. It ships as a Python core with two thin interfaces: an MCP server (recommended) and a JSON CLI runner.

[![PyPI version](https://img.shields.io/pypi/v/xlsx-mcp?style=flat-square&logo=pypi&logoColor=white)](https://pypi.org/project/xlsx-mcp/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://github.com/dwilberger/xlsx-mcp/blob/main/LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg?style=flat-square)](https://www.python.org/downloads/)

---

## The problem

Spreadsheets are everywhere in real workflows -- financial models, reports, data exports, dashboards. But when a coding agent needs to manipulate one, the options are bad:

- **Raw `openpyxl`**: low-level, verbose, easy to break. Agents waste tokens on boilerplate.
- **Manual workarounds**: CSV exports, Pandas round-trips, brittle Python scripts. Lose formatting, formulas, and structure.
- **No contracts**: every tool invents its own shape for requests and responses. Agents guess at error handling.

`xlsx-mcp` solves this with a small, stable command surface. Every command has a defined JSON schema, returns a predictable envelope, and works identically through MCP or the CLI. Your agent opens a workbook, reads cells, writes formulas, saves, and moves on -- no magic, no footguns.

## What it is

- **A focused automation layer** -- 80 commands covering workbook, sheet, cell, range, row, column, formula, table, validation, merge, formatting, export, diff, image, comment, layout, and print workflows.
- **Two interfaces, one core** -- MCP server (stateful, recommended for real workflows) and CLI runner (stateless, good for one-off commands and testing).
- **Deterministic JSON contracts** -- every request and response is schema-defined. No guessing.
- **Explicit by design** -- writes stay in memory until you call save. Dirty workbooks refuse to close. No auto-save, no implicit anything.

## Why xlsx-mcp

Most Excel automation tools fall into two camps: they either require Microsoft Excel installed (Windows-only, COM-based), or they cover only basic read/write operations. `xlsx-mcp` takes a different approach:

- **Cross-platform, zero dependencies on Excel** -- runs anywhere Python runs. Linux servers, macOS, Windows, CI/CD pipelines, Docker containers. No Excel installation needed.
- **Works with the file format, not the application** -- reads and writes `.xlsx` directly via `openpyxl`. No COM automation, no process control, no GUI required.
- **Unique operations** -- image insertion/removal, cell comments, workbook diffing, merge cell management, print layout control (margins, print area, repeating titles, headers/footers), freeze panes, and data validation rules. Features that go beyond basic cell manipulation.
- **Dual interface** -- MCP server for stateful multi-step workflows, CLI for one-off commands and scripting. Same JSON contracts on both.
- **Deterministic contracts** -- every request and response is schema-defined. Agents branch on `error.code`, not guesswork.

## What it's not

- Not a full `openpyxl` wrapper. It covers common automation tasks, not every `openpyxl` API surface.
- Not a spreadsheet UI. No rendering, no formulas engine, no charting.
- Not cloud-connected. It works on local `.xlsx` files only.
- Not a data analysis tool. It moves data around; what you do with it is your problem.

---

## Quick Start

### 1. Install

```bash
pip install xlsx-mcp
```

### 2. Start the MCP server

```bash
xlsx-mcp-server
```

### 3. Point your agent at it

**OpenCode** (`opencode.json`):

```json
{
  "mcp": {
    "xlsx": {
      "type": "local",
      "command": ["xlsx-mcp-server"],
      "enabled": true
    }
  }
}
```

**Codex** (`~/.codex/config.toml`):

```toml
[mcp_servers.xlsx]
command = "xlsx-mcp-server"
```

Your agent can now open workbooks, read cells, write data, set formulas, and save changes.

---

## See it in action

A typical workflow through MCP:

```
1.  workbook_open  { "path": "report.xlsx" }    ->  { "workbook_id": "abc123" }
2.  sheet_list     { "workbook_id": "abc123" }   ->  { "sheets": ["Data", "Summary"] }
3.  cell_read      { ..., "sheet": "Data", "cell": "A1" }  ->  { "value": "Revenue" }
4.  cell_write     { ..., "cell": "B1", "value": 42000 }
5.  formula_set    { ..., "formulas": { "C1": "=SUM(A1:B1)" } }
6.  workbook_save  { "workbook_id": "abc123" }
7.  workbook_close { "workbook_id": "abc123" }
```

Every call returns `{"ok": true, "data": {...}}` on success or `{"ok": false, "error": {"code": "...", "message": "..."}}` on failure. Agents branch on `error.code`, not message text.

---

## Command Surface

All commands are available through MCP tools and through `xlsx-mcp run <command> --input <file>`.

| Area | Commands |
|------|----------|
| Workbook | `workbook_open`, `workbook_save`, `workbook_close`, `workbook_metadata` |
| Sheets | `sheet_list`, `sheet_create`, `sheet_delete`, `sheet_rename`, `sheet_copy`, `sheet_used_range_get`, `sheet_query` |
| Cells | `cell_read`, `cell_write` |
| Ranges | `range_read`, `range_write`, `range_clear`, `range_find`, `range_replace`, `range_sort`, `range_deduplicate`, `range_style_set` |
| Rows | `cells_write`, `rows_append`, `rows_write`, `rows_modify`, `rows_insert`, `rows_delete`, `table_write` |
| Columns | `columns_insert`, `columns_delete`, `columns_modify` |
| Data Queries | `distinct_values`, `duplicates_find`, `blanks_find`, `errors_find` |
| Formulas | `formula_set`, `formula_read`, `formula_copy`, `formula_fill_down`, `formula_fill_right`, `formula_to_values`, `array_formula_list` |
| Names | `defined_names_list`, `defined_name_create`, `defined_name_delete`, `external_links_list` |
| Tables | `table_list`, `table_create`, `table_resize`, `table_delete` |
| Validation | `data_validation_list`, `data_validation_add`, `data_validation_remove` |
| Merged Cells | `merged_cells_list`, `merge_cells`, `unmerge_cells` |
| Conditional Formatting | `conditional_formatting_list`, `conditional_formatting_add`, `conditional_formatting_remove` |
| Export / Import | `sheet_to_csv`, `range_to_csv`, `csv_to_sheet`, `sheet_to_json`, `json_to_sheet` |
| Diff | `sheet_diff`, `workbook_diff` |
| Images | `image_add`, `image_list`, `image_remove` |
| Comments | `comment_add`, `comment_read`, `comment_remove`, `comment_list` |
| Layout / Print | `page_setup_set`, `page_margins_set`, `print_area_set`, `print_titles_set`, `headers_footers_set`, `freeze_panes_set`, `freeze_panes_clear` |

---

## MCP Setup

The MCP server is the recommended interface. It keeps workbook handles alive in the server process, so multi-step workflows (open, edit, save, close) work naturally.

Start the server:

```bash
xlsx-mcp-server
```

### OpenCode

```json
{
  "mcp": {
    "xlsx": {
      "type": "local",
      "command": ["xlsx-mcp-server"],
      "enabled": true
    }
  }
}
```

See `examples/opencode.json`.

### Codex

```toml
[mcp_servers.xlsx]
command = "xlsx-mcp-server"
```

See `examples/codex-config.toml`.

---

## CLI Usage

The CLI is a JSON runner. Useful for validating one command at a time, smoke testing payloads, and scripting simple operations.

```bash
xlsx-mcp run workbook_open --input examples/open.json
```

Each invocation is stateless. The `workbook_id` returned by `workbook_open` lives only in the current process -- separate `xlsx-mcp run ...` calls do not share workbook handles. Use MCP for multi-step workflows.

---

## Error Codes

Clients should branch on `error.code`, not on message text.

| Code | Meaning |
|------|---------|
| `file_not_found` | The file does not exist at the given path |
| `invalid_path` | The path is invalid or inaccessible |
| `workbook_not_open` | The workbook handle is invalid or the session ended |
| `sheet_not_found` | The sheet name was not found in the workbook |
| `invalid_range` | The range string is malformed |
| `invalid_input` | The request payload is missing or has invalid fields |
| `file_locked` | The file is open in another process |
| `unsupported_operation` | The operation is not supported for this object |
| `save_conflict` | The file changed on disk since it was opened |
| `overwrite_required` | The target file already exists and overwrite was not set |
| `target_not_empty` | The destination range already contains data |

---

## Development

From a source checkout:

```bash
pip install -e .[dev]
python -m pytest
```

## Documentation

- `ARCHITECTURE.md` -- implementation structure and boundaries
- `DECISIONS.md` -- stable product and API decisions
- `CONTRIBUTING.md` -- contribution workflow
- `CHANGELOG.md` -- release history
- `SECURITY.md` -- vulnerability reporting and security notes
- `examples/` -- example payloads and MCP client snippets

## License

MIT
