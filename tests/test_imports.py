from xlsx_agent import __version__
from xlsx_agent.cli import app
from xlsx_agent.core.workbook import workbook_open
from xlsx_agent.core.store import WorkbookStore


def test_package_version_exists() -> None:
    assert __version__ == "1.0.0"


def test_cli_app_exists() -> None:
    assert app is not None


def test_workbook_module_exports_open_function() -> None:
    assert callable(workbook_open)


def test_store_starts_empty() -> None:
    store = WorkbookStore()
    assert store.count() == 0
