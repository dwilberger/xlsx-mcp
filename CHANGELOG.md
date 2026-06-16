# Changelog

All notable public changes to this project are documented here.

## 1.0.0 - 2026-06-16

Initial public release.

### Included

- Shared Python core for deterministic `.xlsx` automation.
- CLI JSON runner through `xlsx-mcp`.
- MCP server through `xlsx-mcp-server`.
- 80 commands across workbook, sheet, cell, range, row, column, formula, table, validation, merge, formatting, export, diff, image, comment, layout, and print workflows.
- Explicit JSON contracts and stable error codes.
- Tests for core behavior, CLI smoke paths, MCP smoke paths, and regression coverage.

### Notes

- MCP is the recommended interface for stateful multi-step workflows.
- The standalone CLI is intentionally stateless across separate invocations.
- The project is considered feature-complete for its current scope and is maintained primarily for fixes, docs, packaging, compatibility, and security.
