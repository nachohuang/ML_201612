"""The pluggable translation interface referenced throughout the plan.

Automated translators (API/LLM-backed, aware of a customer glossary, etc.) implement
this directly. The current placeholder (manual_file_translator.py) is human-in-the-loop
and therefore doesn't fit this synchronous shape — see its module docstring.
"""

from __future__ import annotations

from typing import Protocol

from doc_update_tool.models import TranslationItem, TranslationResult


class Translator(Protocol):
    def translate(self, items: list[TranslationItem]) -> list[TranslationResult]: ...
