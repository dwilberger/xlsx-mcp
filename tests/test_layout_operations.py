from __future__ import annotations

from pathlib import Path

import pytest

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.freeze_panes import freeze_panes_clear, freeze_panes_set
from xlsx_agent.core.headers_footers import headers_footers_set
from xlsx_agent.core.page_setup import page_margins_set, page_setup_set
from xlsx_agent.core.print_area import print_area_set, print_titles_set
from xlsx_agent.core.store import WorkbookStore
from xlsx_agent.core.workbook import workbook_close, workbook_open


@pytest.fixture
def store_with_workbook(tmp_path: Path, copied_workbook_path: Path) -> WorkbookStore:
    store = WorkbookStore()
    workbook_open(store, path=str(copied_workbook_path))
    return store


# --- page_setup_set ---


def test_page_setup_set_orientation(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_setup_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        orientation="landscape",
    )

    assert result["orientation"] == "landscape"
    assert result["dirty"] is True


def test_page_setup_set_paper_size_by_name(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_setup_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        paper_size="a4",
    )

    assert result["paper_size"] == 9


def test_page_setup_set_paper_size_by_int(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_setup_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        paper_size=1,
    )

    assert result["paper_size"] == 1


def test_page_setup_set_scale(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_setup_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        scale=75,
    )

    assert result["scale"] == 75


def test_page_setup_set_fit_to_page(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_setup_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        fit_to_page=True,
        fit_to_width=1,
        fit_to_height=0,
    )

    assert result["fit_to_page"] is True
    assert result["fit_to_width"] == 1
    assert result["fit_to_height"] == 0


def test_page_setup_set_invalid_orientation(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        page_setup_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            orientation="diagonal",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"


def test_page_setup_set_invalid_paper_size(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        page_setup_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            paper_size="nonexistent",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"


def test_page_setup_set_invalid_scale(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    try:
        page_setup_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            scale=500,
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "invalid_input"


def test_page_setup_set_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        page_setup_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            orientation="landscape",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)


# --- page_margins_set ---


def test_page_margins_set(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_margins_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        left=1.0,
        right=1.0,
        top=1.5,
        bottom=1.5,
        header=0.5,
        footer=0.5,
    )

    assert result["left"] == 1.0
    assert result["right"] == 1.0
    assert result["top"] == 1.5
    assert result["bottom"] == 1.5
    assert result["header"] == 0.5
    assert result["footer"] == 0.5
    assert result["dirty"] is True


def test_page_margins_set_partial(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = page_margins_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        left=2.0,
    )

    assert result["left"] == 2.0
    assert result["dirty"] is True


def test_page_margins_set_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        page_margins_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            left=1.0,
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)


# --- print_area_set ---


def test_print_area_set(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = print_area_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        range="A1:C10",
    )

    assert result["print_area"] is not None
    assert "$A$1:$C$10" in result["print_area"]
    assert result["dirty"] is True


def test_print_area_clear(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    print_area_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        range="A1:C10",
    )

    result = print_area_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        range=None,
    )

    assert result["print_area"] is None
    assert result["dirty"] is True


def test_print_area_set_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        print_area_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            range="A1:C10",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)


# --- print_titles_set ---


def test_print_titles_set(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = print_titles_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        rows="1:1",
        columns="A:A",
    )

    assert result["rows"] == "$1:$1"
    assert result["columns"] == "$A:$A"
    assert result["dirty"] is True


def test_print_titles_set_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        print_titles_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            rows="$1",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)


# --- headers_footers_set ---


def test_headers_footers_set(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = headers_footers_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        odd_header="Page &P of &N",
        odd_footer="Confidential",
    )

    assert result["dirty"] is True


def test_headers_footers_set_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        headers_footers_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            odd_header="Test",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)


# --- freeze_panes_set / freeze_panes_clear ---


def test_freeze_panes_set(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = freeze_panes_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="B2",
    )

    assert result["freeze_panes"] == "B2"
    assert result["dirty"] is True


def test_freeze_panes_set_top_row(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = freeze_panes_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="A2",
    )

    assert result["freeze_panes"] == "A2"


def test_freeze_panes_set_left_column(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    result = freeze_panes_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="B1",
    )

    assert result["freeze_panes"] == "B1"


def test_freeze_panes_clear(store_with_workbook: WorkbookStore) -> None:
    store = store_with_workbook
    workbook_id = list(store._items.keys())[0]

    freeze_panes_set(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
        cell="B2",
    )

    result = freeze_panes_clear(
        store,
        workbook_id=workbook_id,
        sheet="Sheet1",
    )

    assert result["freeze_panes"] is None
    assert result["dirty"] is True


def test_freeze_panes_set_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        freeze_panes_set(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
            cell="A1",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)


def test_freeze_panes_clear_read_only(copied_workbook_path: Path) -> None:
    store = WorkbookStore()
    session = workbook_open(store, path=str(copied_workbook_path), read_only=True)
    workbook_id = session.workbook_id

    try:
        freeze_panes_clear(
            store,
            workbook_id=workbook_id,
            sheet="Sheet1",
        )
        assert False, "Expected XlsxAgentError"
    except XlsxAgentError as exc:
        assert exc.code == "unsupported_operation"
    finally:
        workbook_close(store, workbook_id=workbook_id, force=True)
