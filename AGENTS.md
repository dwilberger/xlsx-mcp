# AGENTS

## Working Baseline

- This repo uses a `src/` layout. Install it in editable mode before running tests or importing `xlsx_agent`: `python -m pip install -e .[dev]`.
- Focused verification command: `python -m pytest`.
- CI runs the pytest suite through GitHub Actions. There is no lint or typecheck workflow configured.

## Current Entry Points

- CLI entrypoint: `xlsx-mcp` -> `src/xlsx_agent/cli.py`
- MCP entrypoint: `xlsx-mcp-server` -> `src/xlsx_agent/mcp_server.py`
- The 1.0 command surface is implemented in both the CLI JSON runner and the MCP server.
- Important CLI limitation: `workbook_id` state is process-local. Separate `xlsx-mcp run ...` invocations do not share handles, so real multi-step workflows should use MCP.

## Architecture Rules

- Keep workbook behavior in `src/xlsx_agent/core/`.
- Keep schemas in `src/xlsx_agent/schemas/`.
- Keep CLI and MCP layers thin: argument/transport handling only. Do not duplicate workbook logic there.
- If a feature needs interface-specific handling, adapt in the CLI/MCP layer and keep the core as the source of truth.

## Product Constraints

- The product is intentionally a narrow automation layer, not a full `openpyxl` wrapper.
- Do not expose raw `openpyxl` internals as public API.
- Favor deterministic JSON contracts and explicit error codes over convenience shortcuts.
- Write project documentation, examples, user-facing text, code comments, and implementation-facing docs in English.

## Workflow Files

- `CHANGELOG.md` tracks public releases.
- `DECISIONS.md` and `ARCHITECTURE.md` define product direction and should be checked before changing public behavior.

## Testing And Docs Expectations

- When behavior changes, add or update tests in `tests/`.
- When public behavior changes, also update the relevant docs and examples, especially `README.md` and files under `examples/`.
