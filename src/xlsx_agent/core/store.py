from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook

from xlsx_agent.core.errors import XlsxAgentError


def build_revision_token(path: Path) -> str:
    stat = path.stat()
    return f"{stat.st_mtime_ns}:{stat.st_size}"


@dataclass(slots=True)
class WorkbookSession:
    workbook_id: str
    path: Path
    workbook: Workbook
    revision_token: str
    read_only: bool = False
    dirty: bool = False


class WorkbookStore:
    """In-memory workbook handle store for process-scoped workbook sessions."""

    def __init__(self, max_open_workbooks: int = 10) -> None:
        self.max_open_workbooks = max_open_workbooks
        self._items: dict[str, WorkbookSession] = {}

    def count(self) -> int:
        return len(self._items)

    def create(
        self,
        *,
        path: Path,
        workbook: Workbook,
        revision_token: str,
        read_only: bool = False,
    ) -> WorkbookSession:
        if self.count() >= self.max_open_workbooks:
            raise XlsxAgentError(
                code="unsupported_operation",
                message=(
                    f"Cannot open more than {self.max_open_workbooks} workbooks "
                    "in the current process."
                ),
                details={"max_open_workbooks": self.max_open_workbooks},
            )

        workbook_id = f"wb_{uuid4().hex[:12]}"
        session = WorkbookSession(
            workbook_id=workbook_id,
            path=path,
            workbook=workbook,
            revision_token=revision_token,
            read_only=read_only,
        )
        self._items[workbook_id] = session
        return session

    def get(self, workbook_id: str) -> WorkbookSession:
        try:
            return self._items[workbook_id]
        except KeyError as exc:
            raise XlsxAgentError(
                code="workbook_not_open",
                message=f"Workbook '{workbook_id}' is not open.",
            ) from exc

    def close(self, workbook_id: str) -> WorkbookSession:
        session = self.get(workbook_id)
        del self._items[workbook_id]
        return session

    def mark_dirty(self, workbook_id: str, *, dirty: bool = True) -> WorkbookSession:
        session = self.get(workbook_id)
        session.dirty = dirty
        return session
