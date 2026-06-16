# Architecture

`xlsx-mcp` is a small automation layer with one shared implementation and two thin interfaces.

## Layers

1. Core library in `src/xlsx_agent/core/`
2. CLI runner in `src/xlsx_agent/cli.py`
3. MCP server in `src/xlsx_agent/mcp_server.py`

The core owns workbook behavior. The CLI and MCP layers validate transport input, call the core, and return structured output.

## Design Goals

- deterministic JSON contracts
- explicit errors with stable error codes
- local-first file access
- no implicit saves
- no raw `openpyxl` object exposure as public API
- minimal interface-specific logic

## State Model

- `workbook_open` returns an opaque `workbook_id`.
- Workbook handles are process-local and kept in memory.
- MCP is stateful while the server process runs.
- The CLI is stateless across separate invocations.
- Writes stay in memory until `workbook_save` is called.
- Closing a dirty workbook fails unless the caller explicitly discards changes.

## Contracts

Commands use request schemas from `src/xlsx_agent/schemas/` and return deterministic dictionaries.

Typical success envelope:

```json
{
  "ok": true,
  "data": {}
}
```

Typical error envelope:

```json
{
  "ok": false,
  "error": {
    "code": "sheet_not_found",
    "message": "Sheet 'Report' was not found"
  }
}
```

Callers should branch on `error.code` instead of parsing human-readable messages.

## Package Layout

```text
src/
  xlsx_agent/
    __init__.py
    cli.py
    mcp_server.py
    schemas/
      common.py
      essentials.py
      advanced.py
    core/
      cell.py
      columns.py
      comments.py
      data.py
      diff.py
      errors.py
      export.py
      formula.py
      formula_ops.py
      freeze_panes.py
      headers_footers.py
      images.py
      merges.py
      names.py
      page_setup.py
      print_area.py
      range.py
      rows.py
      sheet_ops.py
      sheets.py
      store.py
      style.py
      tables.py
      validation.py
      validation_ops.py
      workbook.py
      workbook_ops.py
```

## Boundaries

`xlsx-mcp` intentionally covers common, automatable workbook tasks. It does not attempt to wrap the complete `openpyxl` object model, author pivots, manipulate VBA, edit arbitrary workbook XML, or provide cloud synchronization.
