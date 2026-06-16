from __future__ import annotations

from pathlib import Path

import pytest

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.images import image_add, image_list, image_remove
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.workbook import workbook_close, workbook_open


@pytest.fixture
def test_image_path(tmp_path: Path) -> Path:
    try:
        from PIL import Image as PILImage

        img_path = tmp_path / "test_image.png"
        img = PILImage.new("RGB", (100, 100), color="red")
        img.save(str(img_path))
        return img_path
    except ImportError:
        pytest.skip("Pillow not installed, skipping image tests")


@pytest.fixture
def store_with_workbook(tmp_path: Path, copied_workbook_path: Path) -> WorkbookStore:
    store = WorkbookStore()
    workbook_open(store, path=str(copied_workbook_path))
    return store


def test_image_add(store_with_workbook: WorkbookStore, test_image_path: Path) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = image_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        path=str(test_image_path),
        anchor="A1",
        width=50,
        height=50,
    )

    assert result["workbook_id"] == workbook_id
    assert result["sheet"] == "Sheet1"
    assert result["anchor"] == "A1"
    assert result["width"] == 50
    assert result["height"] == 50
    assert result["dirty"] is True


def test_image_list(store_with_workbook: WorkbookStore, test_image_path: Path) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    image_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        path=str(test_image_path),
        anchor="B2",
    )

    result = image_list(store, workbook_id=workbook_id, sheet="Sheet1")

    assert result["count"] == 1
    assert result["images"][0]["anchor"] == "B2"


def test_image_remove(store_with_workbook: WorkbookStore, test_image_path: Path) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    image_add(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        path=str(test_image_path),
        anchor="C3",
    )

    result = image_remove(store, workbook_id=workbook_id, sheet="Sheet1", index=0)

    assert result["removed_index"] == 0
    assert result["anchor"] == "C3"
    assert result["dirty"] is True

    list_result = image_list(store, workbook_id=workbook_id, sheet="Sheet1")
    assert list_result["count"] == 0


def test_image_remove_invalid_index(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        image_remove(store, workbook_id=workbook_id, sheet="Sheet1", index=0)
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"


def test_image_add_file_not_found(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        image_add(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            path="/nonexistent/image.png",
            anchor="A1",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "file_not_found"


def test_image_add_read_only(copied_workbook_path: Path, test_image_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        image_add(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            path=str(test_image_path),
            anchor="A1",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)
