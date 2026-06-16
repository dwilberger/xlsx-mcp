# Contributing

## Goal

Keep `xlsx-mcp` small, deterministic, and easy to reason about.

The 1.0 release is considered feature-complete for the current product scope. Contributions are still welcome, especially for bugs, tests, docs, packaging, compatibility, and security.

## Development Principles

- prefer the smallest correct change
- keep contracts explicit
- avoid hidden behavior
- favor reliability over feature count
- do not expose raw `openpyxl` internals as public API

## Repository Workflow

1. Open an issue or discussion for non-trivial behavior changes.
2. Keep changes scoped to a clear objective.
3. Add or update tests when behavior changes.
4. Update docs when public behavior changes.
5. Run `python -m pytest` before submitting changes.

## Code Organization

- Put workbook behavior in `src/xlsx_agent/core/`.
- Keep schemas in `src/xlsx_agent/schemas/`.
- Keep CLI code focused on argument and output handling.
- Keep MCP code focused on tool registration and transport.
- Avoid duplicating logic across interfaces.

## Testing Expectations

Before considering work complete:

- schema validation should be covered where relevant
- core behavior should be tested directly
- CLI changes should include smoke coverage when practical
- MCP changes should include tool coverage when practical
- regressions should be captured with focused tests

## Public Contract Expectations

When changing public behavior, check:

- command names
- JSON request and response shapes
- error codes
- examples in `examples/`
- user-facing documentation in `README.md`

## Scope Discipline

Do not expand scope casually.

If a change pushes the project toward full `openpyxl` surface coverage, it likely does not fit the current product direction.

## Pull Requests

When opening work for review, describe:

- what changed
- why it changed
- what was tested
- any follow-up work still needed
