from xlsx_agent.mcp_server import main, mcp, workbook_open


def test_mcp_server_object_exists() -> None:
    assert mcp is not None


def test_mcp_tool_returns_error_envelope_for_missing_file() -> None:
    result = workbook_open(path="missing.xlsx")

    assert result["ok"] is False
    assert result["error"]["code"] == "file_not_found"


def test_mcp_main_exists() -> None:
    assert callable(main)
