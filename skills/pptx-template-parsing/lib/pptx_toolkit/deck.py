"""Deck-level operations: deleting slides.

Extracted from scripts/generate_strategy_ppt.py (delete_slide).
python-pptx has no public slide-delete API, so we edit the slide id list and
drop the relationship.
"""
from __future__ import annotations


def delete_slide(prs, index: int) -> None:
    """Delete the slide at 0-based `index` from the presentation."""
    slide_id_list = prs.slides._sldIdLst  # noqa: SLF001
    slides = list(slide_id_list)
    slide = slides[index]
    rel_id = slide.rId
    prs.part.drop_rel(rel_id)
    slide_id_list.remove(slide)


def delete_slides(prs, indexes) -> None:
    """Delete multiple slides given 0-based indexes (handled high-to-low so
    earlier deletions don't shift later indexes)."""
    for index in sorted(indexes, reverse=True):
        delete_slide(prs, index)
