from doc_update_tool.diffing.align import align_and_diff
from doc_update_tool.models import ChangeType, Segment


def _seg(row_key, zh, en="", field="需求說明", loc=None):
    loc = loc or f"C{row_key}"
    return Segment(
        location_id=loc,
        en_location_id=f"D{row_key}",
        row_key=row_key,
        field_name=field,
        zh_text=zh,
        en_text=en,
    )


def test_modified_and_added_with_row_key():
    old = [_seg("1", "A"), _seg("2", "B")]
    new = [_seg("1", "A"), _seg("2", "B2"), _seg("3", "C")]

    records = align_and_diff(old, new)
    by_key = {r.location_id: r for r in records}

    assert by_key["C1"].change_type == ChangeType.UNCHANGED
    assert by_key["C2"].change_type == ChangeType.MODIFIED
    assert by_key["C2"].old_zh == "B"
    assert by_key["C2"].new_zh == "B2"
    assert by_key["C2"].new_en_location_id == "D2"
    assert by_key["C3"].change_type == ChangeType.ADDED
    assert by_key["C3"].new_zh == "C"


def test_deleted_row():
    old = [_seg("1", "A"), _seg("2", "B"), _seg("3", "C")]
    new = [_seg("1", "A"), _seg("3", "C")]

    records = align_and_diff(old, new)
    deleted = [r for r in records if r.change_type == ChangeType.DELETED]
    assert len(deleted) == 1
    assert deleted[0].old_zh == "B"
    assert deleted[0].new_en_location_id is None


def test_unchanged_carries_forward_old_english():
    old = [_seg("1", "A", en="Translated A")]
    new = [_seg("1", "A")]

    records = align_and_diff(old, new)
    assert records[0].change_type == ChangeType.UNCHANGED
    assert records[0].new_en == "Translated A"


def test_normalization_ignores_full_width_and_whitespace_differences():
    old = [_seg("1", "第一項")]
    new = [_seg("1", "第一項 ")]  # trailing space only

    records = align_and_diff(old, new)
    assert records[0].change_type == ChangeType.UNCHANGED


def test_position_fallback_without_row_key():
    old = [_seg(None, "A", loc="C2"), _seg(None, "B", loc="C3")]
    new = [_seg(None, "A2", loc="C2"), _seg(None, "B", loc="C3")]

    records = align_and_diff(old, new)
    assert any(r.change_type == ChangeType.MODIFIED for r in records)
