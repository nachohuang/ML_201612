"""Align old/new Segment sequences and classify each pairing as added/modified/deleted/
unchanged.

Because the customer edits the bilingual file we sent them (old English + new Chinese
mixed in the same file — see plan Context), only the Chinese side drives alignment and
change detection: English is never customer-authored, so comparing it would just
surface our own translation choices as false "changes".

When Segment.row_key is available (a configured key_columns value) it's used to align
rows/paragraphs by content, correctly surviving reordering/insertion/deletion. Without
it, alignment falls back to position, which cannot distinguish "row 3 was deleted" from
"everything after row 3 shifted" — this is a known limitation, not a bug: configure
key_columns whenever the document has a stable identifier column.
"""

from __future__ import annotations

from difflib import SequenceMatcher

from doc_update_tool.diffing.text_normalize import normalize
from doc_update_tool.models import ChangeRecord, ChangeType, Segment


def _keys(segments: list[Segment]) -> list[str]:
    keys = []
    for i, s in enumerate(segments):
        if s.row_key is not None:
            keys.append(f"key:{normalize(s.row_key)}:{s.field_name or ''}")
        else:
            keys.append(f"pos:{i}:{s.field_name or ''}")
    return keys


def align_and_diff(old_segments: list[Segment], new_segments: list[Segment]) -> list[ChangeRecord]:
    old_keys = _keys(old_segments)
    new_keys = _keys(new_segments)
    matcher = SequenceMatcher(a=old_keys, b=new_keys, autojunk=False)

    records: list[ChangeRecord] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for old, new in zip(old_segments[i1:i2], new_segments[j1:j2]):
                records.append(_diff_pair(old, new))
        elif tag == "insert":
            for new in new_segments[j1:j2]:
                records.append(_added(new))
        elif tag == "delete":
            for old in old_segments[i1:i2]:
                records.append(_deleted(old))
        elif tag == "replace":
            old_slice = old_segments[i1:i2]
            new_slice = new_segments[j1:j2]
            if len(old_slice) == len(new_slice):
                for old, new in zip(old_slice, new_slice):
                    records.append(_diff_pair(old, new))
            else:
                records.extend(_deleted(old) for old in old_slice)
                records.extend(_added(new) for new in new_slice)
    return records


def _diff_pair(old: Segment, new: Segment) -> ChangeRecord:
    if normalize(old.zh_text) == normalize(new.zh_text):
        change_type = ChangeType.UNCHANGED
    else:
        change_type = ChangeType.MODIFIED
    return ChangeRecord(
        location_id=new.location_id,
        field_name=new.field_name,
        change_type=change_type,
        old_zh=old.zh_text,
        new_zh=new.zh_text,
        old_en=old.en_text,
        new_en=old.en_text,  # carried forward until a fresh translation is applied
        new_en_location_id=new.en_location_id,
    )


def _added(new: Segment) -> ChangeRecord:
    return ChangeRecord(
        location_id=new.location_id,
        field_name=new.field_name,
        change_type=ChangeType.ADDED,
        new_zh=new.zh_text,
        new_en_location_id=new.en_location_id,
    )


def _deleted(old: Segment) -> ChangeRecord:
    return ChangeRecord(
        location_id=old.location_id,
        field_name=old.field_name,
        change_type=ChangeType.DELETED,
        old_zh=old.zh_text,
        old_en=old.en_text,
    )
