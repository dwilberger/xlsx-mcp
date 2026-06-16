from __future__ import annotations

from pathlib import Path

import pytest

from xlsx_agent.core.comments import comment_add, comment_list, comment_read, comment_remove
from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.workbook import workbook_close, workbook_open


@pytest.fixture
def store_with_workbook(tmp_path: Path, copied_workbook_path: Path) -> WorkbookStore:
    store = WorkbookStore()
    workbook_open(store, path=str(copied_workbook_path))
    return store


def test_comment_add(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = comment_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="A1",
        text="Test comment",
        author="Test Author",
    )

    assert result["workbook_id"] == workbook_id
    assert result["sheet"] == "Sheet1"
    assert result["cell"] == "A1"
    assert result["text"] == "Test comment"
    assert result["author"] == "Test Author"
    assert result["dirty"] is True


def test_comment_read(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    comment_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="B2",
        text="Readable comment",
        author="Author",
    )

    result = comment_read(store, workbook_id=workbook_id, sheet="Sheet1", cell="B2")

    assert result["text"] == "Readable comment"
    assert result["author"] == "Author"
    assert result["cell"] == "B2"


def test_comment_read_not_found(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        comment_read(store, workbook_id=workbook_id, sheet="Sheet1", cell="A1")
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "not_found"


def test_comment_remove(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    comment_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="C3",
        text="To be removed",
        author="Author",
    )

    result = comment_remove(store, workbook_id=workbook_id, sheet="Sheet1", cell="C3")

    assert result["removed"] is True
    assert result["dirty"] is True

    try:
        comment_read(store, workbook_id=workbook_id, sheet="Sheet1", cell="C3")
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "not_found"


def test_comment_remove_not_found(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        comment_remove(store, workbook_id=workbook_id, sheet="Sheet1", cell="A1")
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "not_found"


def test_comment_list(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    comment_add(store, workbook_id=workbook_id, sheet="Sheet1", cell="A1", text="Comment 1")
    comment_add(store, workbook_id=workbook_id, sheet="Sheet1", cell="B2", text="Comment 2")

    result = comment_list(store, workbook_id=workbook_id, sheet="Sheet1")

    assert result["count"] == 2
    texts = [c["text"] for c in result["comments"]]
    assert "Comment 1" in texts
    assert "Comment 2" in texts


def test_comment_list_empty(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = comment_list(store, workbook_id=workbook_id, sheet="Sheet1")

    assert result["count"] == 0
    assert result["comments"] == []


def test_comment_add_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        comment_add(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            cell="A1",
            text="Test",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)
