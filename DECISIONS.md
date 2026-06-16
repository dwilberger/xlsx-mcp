# Decision Record

This file records the stable product and API decisions for `xlsx-mcp` 1.0.

## Product Identity

- Use `xlsx-mcp` as the package and product name.
- Use `xlsx_agent` as the Python import path.
- Use `xlsx-mcp` for the CLI entrypoint.
- Use `xlsx-mcp-server` for the MCP server entrypoint.

## Interface Model

- MCP is the primary interface for real multi-step workflows.
- The CLI is a stateless JSON runner for one command at a time.
- Command names are the same in the core, CLI, MCP tools, examples, and docs.

## Workbook State

- `workbook_open` returns a process-local opaque `workbook_id`.
- Handles are stored in memory and do not survive process restart.
- The handle store is intentionally simple and bounded.
- Callers must explicitly save and close workbooks.

## Save Behavior

- Writes mutate the in-memory workbook session.
- `workbook_save` is always explicit.
- Save-as operations require an explicit overwrite decision when the target exists.
- Save conflict detection should protect callers from overwriting files changed outside the current session.

## Write Behavior

- Write commands do not create sheets implicitly.
- `cells_write` writes literal values only.
- Formula authoring belongs to `formula_set`.
- Operations should fail explicitly instead of silently repairing ambiguous input.

## Error Model

- Error codes are stable, flat, and machine-friendly.
- Clients should branch on `error.code`, not on message text.
- Public errors should avoid leaking raw implementation details.

## Scope

- The project is a narrow automation layer, not a full `openpyxl` wrapper.
- Raw workbook XML access, VBA/macro manipulation, pivot authoring, and cloud sync are out of scope.
- New feature areas are not planned after 1.0 unless there is clear user demand and a small, deterministic contract.
