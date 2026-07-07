"""Common interface implemented by excel_adapter and word_adapter.

A DocumentAdapter loads a document, extracts zh/en Segments from it (config-driven,
see config.py), and can write translated text back to the exact location it came from
without disturbing anything else in the file (styles, formulas, other sheets/paragraphs).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol

from doc_update_tool.models import Segment, SegmentUpdate


class DocumentAdapter(Protocol):
    def load(self, path: Path) -> Any: ...

    def extract_segments(self, doc: Any) -> list[Segment]: ...

    def apply_translations(self, doc: Any, updates: list[SegmentUpdate]) -> None: ...

    def save(self, doc: Any, path: Path) -> None: ...
