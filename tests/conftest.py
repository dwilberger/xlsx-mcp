from __future__ import annotations

import shutil
from pathlib import Path

import pytest

FIXTURE_DIR = Path(__file__).parent / "fixtures"
BASE_WORKBOOK_PATH = FIXTURE_DIR / "base-workbook.xlsx"


@pytest.fixture
def copied_workbook_path(tmp_path: Path) -> Path:
    target_path = tmp_path / BASE_WORKBOOK_PATH.name
    shutil.copy2(BASE_WORKBOOK_PATH, target_path)
    return target_path


@pytest.fixture(autouse=True)
def clear_mcp_store() -> None:
    from xlsx_agent.mcp_server import store

    for session in list(store._items.values()):
        session.workbook.close()
    store._items.clear()

    yield

    for session in list(store._items.values()):
        session.workbook.close()
    store._items.clear()
