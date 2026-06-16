from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from xlsx_agent.core.errors import XlsxAgentError
from xlsx_agent.core.store import WorkbookSession, WorkbookStore, build_revision_token
from xlsx_agent.schemas.common import WorkbookHandle


def _resolve_existing_workbook_path(path: str) -> Path:
    resolved_path = Path(path).expanduser().resolve()
    if not resolved_path.exists():
        raise XlsxAgentError(
            code="file_not_found",
            message=f"Workbook '{resolved_path}' was not found.",
        )
    if not resolved_path.is_file():
        raise XlsxAgentError(
            code="invalid_path",
            message=f"Path '{resolved_path}' is not a file.",
        )
    return resolved_path


def _handle_from_session(session: WorkbookSession) -> WorkbookHandle:
    return WorkbookHandle(
        workbook_id=session.workbook_id,
        path=str(session.path),
        read_only=session.read_only,
        dirty=session.dirty,
    )


def workbook_open(store: WorkbookStore, *, path: str, read_only: bool = False) -> WorkbookHandle:
    resolved_path = _resolve_existing_workbook_path(path)
    try:
        workbook = load_workbook(resolved_path, read_only=read_only)
    except PermissionError as exc:
        raise XlsxAgentError(
            code="file_locked",
            message=f"Workbook '{resolved_path}' could not be opened.",
        ) from exc
    except OSError as exc:
        raise XlsxAgentError(
            code="invalid_path",
            message=f"Workbook '{resolved_path}' could not be opened.",
        ) from exc

    session = store.create(
        path=resolved_path,
        workbook=workbook,
        revision_token=build_revision_token(resolved_path),
        read_only=read_only,
    )
    return _handle_from_session(session)


def workbook_save(
    store: WorkbookStore,
    *,
    workbook_id: str,
    path: str | None = None,
    overwrite: bool = False,
) -> WorkbookHandle:
    session = store.get(workbook_id)
    if session.read_only:
        raise XlsxAgentError(
            code="unsupported_operation",
            message=f"Workbook '{workbook_id}' was opened in read-only mode.",
        )

    target_path = session.path if path is None else Path(path).expanduser().resolve()
    saving_to_original_path = target_path == session.path

    if saving_to_original_path:
        if not session.path.exists():
            raise XlsxAgentError(
                code="save_conflict",
                message=f"Workbook '{session.path}' no longer exists on disk.",
            )
        current_revision = build_revision_token(session.path)
        if current_revision != session.revision_token:
            raise XlsxAgentError(
                code="save_conflict",
                message=f"Workbook '{session.path}' changed on disk after it was opened.",
            )
    elif target_path.exists() and not overwrite:
        raise XlsxAgentError(
            code="overwrite_required",
            message=f"Workbook '{target_path}' already exists.",
        )

    try:
        session.workbook.save(target_path)
    except PermissionError as exc:
        raise XlsxAgentError(
            code="file_locked",
            message=f"Workbook '{target_path}' could not be saved.",
        ) from exc
    except OSError as exc:
        raise XlsxAgentError(
            code="invalid_path",
            message=f"Workbook '{target_path}' could not be saved.",
        ) from exc

    session.path = target_path
    session.revision_token = build_revision_token(target_path)
    session.dirty = False
    return _handle_from_session(session)


def workbook_close(
    store: WorkbookStore,
    *,
    workbook_id: str,
    force: bool = False,
) -> WorkbookHandle:
    session = store.get(workbook_id)
    if session.dirty and not force:
        raise XlsxAgentError(
            code="invalid_input",
            message=(
                f"Workbook '{workbook_id}' has unsaved changes. "
                "Save it first or close with force=true."
            ),
        )

    session = store.close(workbook_id)
    session.workbook.close()
    return _handle_from_session(session)
