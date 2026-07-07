"""python-docx-backed adapter supporting three zh/en layouts, picked via WordMapping.mode:

- alternating_paragraphs: a zh paragraph followed by its en paragraph (v1 documents are
  zh-only, so the en paragraph may not exist yet — see insert_after handling below)
- same_paragraph_split: a single paragraph holding "zh<delimiter>en"
- table_based: a docx table with dedicated zh/en columns (like the Excel layout)

Real customer paragraph layout is unknown, so language is detected by CJK character
ratio rather than any hardcoded pattern. Use `docmgr inspect-word` to see how a real
document's paragraphs classify before picking a mode/config.
"""

from __future__ import annotations

from pathlib import Path

import docx
from docx.document import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from doc_update_tool.config import WordMapping
from doc_update_tool.models import Segment, SegmentUpdate


def is_chinese_text(text: str, threshold: float = 0.3) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    cjk = sum(1 for ch in stripped if "一" <= ch <= "鿿")
    return (cjk / len(stripped)) >= threshold


def _set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    """Overwrite a paragraph's visible text. Keeps the first run's formatting (font,
    bold, etc.) and clears any other runs — full run-level fidelity isn't preserved
    when the translated text doesn't align with original run boundaries."""
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def _insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    """python-docx has no public API for this; standard recipe is to build a bare <w:p>
    and splice it into the XML tree right after the reference paragraph."""
    new_p = paragraph._p.makeelement(qn("w:p"), {})
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.add_run(text)
    return new_para


class WordAdapter:
    def __init__(self, mapping: WordMapping):
        self.mapping = mapping

    def load(self, path: Path) -> Document:
        return docx.Document(str(path))

    def save(self, doc: Document, path: Path) -> None:
        doc.save(str(path))

    def extract_segments(self, doc: Document) -> list[Segment]:
        if self.mapping.mode == "alternating_paragraphs":
            return self._extract_alternating(doc)
        if self.mapping.mode == "same_paragraph_split":
            return self._extract_same_split(doc)
        if self.mapping.mode == "table_based":
            return self._extract_table_based(doc)
        raise ValueError(f"Unknown word_mapping.mode: {self.mapping.mode!r}")

    def apply_translations(self, doc: Document, updates: list[SegmentUpdate]) -> None:
        direct, tails, inserts, tables = [], [], [], []
        for u in updates:
            parts = u.en_location_id.split(":")
            if parts[0] == "para" and len(parts) == 2:
                direct.append(u)
            elif parts[0] == "para" and parts[-1] == "tail":
                tails.append(u)
            elif parts[0] == "para" and parts[-1] == "insert_after":
                inserts.append(u)
            elif parts[0] == "table":
                tables.append(u)
            else:
                raise ValueError(f"Unrecognized en_location_id: {u.en_location_id!r}")

        for u in direct:
            idx = int(u.en_location_id.split(":")[1])
            _set_paragraph_text(doc.paragraphs[idx], u.en_text)

        for u in tails:
            idx = int(u.en_location_id.split(":")[1])
            para = doc.paragraphs[idx]
            delimiter = self.mapping.split_delimiter or "\n"
            zh_part = para.text.split(delimiter, 1)[0]
            _set_paragraph_text(para, f"{zh_part}{delimiter}{u.en_text}")

        # Insertions change paragraph indices below the insertion point, so process
        # highest-index-first: every not-yet-processed index is still below every
        # insertion point already handled, so it can't have shifted yet.
        for u in sorted(inserts, key=lambda u: int(u.en_location_id.split(":")[1]), reverse=True):
            idx = int(u.en_location_id.split(":")[1])
            _insert_paragraph_after(doc.paragraphs[idx], u.en_text)

        for u in tables:
            _, t_idx, r_idx, c_idx = u.en_location_id.split(":")
            cell = doc.tables[int(t_idx)].rows[int(r_idx)].cells[int(c_idx)]
            cell.text = u.en_text

    def _extract_alternating(self, doc: Document) -> list[Segment]:
        paragraphs = doc.paragraphs
        n = len(paragraphs)
        segments: list[Segment] = []
        i = 0
        while i < n:
            text = paragraphs[i].text
            if not is_chinese_text(text):
                i += 1
                continue
            zh_idx = i
            j = i + 1
            while j < n and not paragraphs[j].text.strip():
                j += 1
            if j < n and paragraphs[j].text.strip() and not is_chinese_text(paragraphs[j].text):
                en_idx = j
                next_i = j + 1
            else:
                en_idx = None  # no en paragraph yet (e.g. first-time zh-only document)
                next_i = i + 1
            segments.append(
                Segment(
                    location_id=f"para:{zh_idx}",
                    en_location_id=f"para:{en_idx}" if en_idx is not None else f"para:{zh_idx}:insert_after",
                    row_key=None,
                    field_name=None,
                    zh_text=paragraphs[zh_idx].text,
                    en_text=paragraphs[en_idx].text if en_idx is not None else "",
                )
            )
            i = next_i
        return segments

    def _extract_same_split(self, doc: Document) -> list[Segment]:
        delimiter = self.mapping.split_delimiter or "\n"
        segments = []
        for idx, para in enumerate(doc.paragraphs):
            text = para.text
            if delimiter in text:
                zh_part, en_part = text.split(delimiter, 1)
            elif is_chinese_text(text):
                zh_part, en_part = text, ""
            else:
                continue
            segments.append(
                Segment(
                    location_id=f"para:{idx}:head",
                    en_location_id=f"para:{idx}:tail",
                    row_key=None,
                    field_name=None,
                    zh_text=zh_part.strip(),
                    en_text=en_part.strip(),
                )
            )
        return segments

    def _extract_table_based(self, doc: Document) -> list[Segment]:
        segments = []
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                if r_idx == 0:
                    continue  # header row
                for zh_c, en_c in zip(self.mapping.zh_columns, self.mapping.en_columns):
                    zh_cell = row.cells[int(zh_c)]
                    en_cell = row.cells[int(en_c)]
                    if not zh_cell.text.strip():
                        continue
                    segments.append(
                        Segment(
                            location_id=f"table:{t_idx}:{r_idx}:{zh_c}",
                            en_location_id=f"table:{t_idx}:{r_idx}:{en_c}",
                            row_key=None,
                            field_name=None,
                            zh_text=zh_cell.text,
                            en_text=en_cell.text,
                        )
                    )
        return segments


def inspect(path: Path, sample_chars: int = 60) -> list[dict]:
    """List each paragraph's index, char count, CJK ratio and a text preview, to help
    pick word_mapping.mode/split_delimiter for a real document without guessing."""
    doc = docx.Document(str(path))
    result = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        cjk = sum(1 for ch in text if "一" <= ch <= "鿿")
        result.append(
            {
                "index": idx,
                "length": len(text),
                "cjk_ratio": round(cjk / len(text), 2) if text else 0.0,
                "is_chinese": is_chinese_text(text),
                "preview": text[:sample_chars],
            }
        )
    return result
