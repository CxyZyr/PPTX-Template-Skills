"""Geometry helpers: EMU/point/inch conversion and shape resizing.

Extracted and generalized from scripts/generate_strategy_ppt.py.
All PowerPoint internal measurements are in EMU (English Metric Units).
"""
from __future__ import annotations

EMU_PER_PT = 12700
EMU_PER_INCH = 914400


def emu_to_pt(value: int) -> float:
    return value / EMU_PER_PT


def pt_to_emu(value: float) -> int:
    return int(value * EMU_PER_PT)


def emu_to_inch(value: int) -> float:
    return value / EMU_PER_INCH


def inch_to_emu(value: float) -> int:
    return int(value * EMU_PER_INCH)


def resize_shape(shape, *, left=None, top=None, width=None, height=None) -> None:
    """Move/resize a shape in place. Only the provided dimensions change.

    Values are EMU integers (use pt_to_emu / inch_to_emu to build them).
    """
    if left is not None:
        shape.left = left
    if top is not None:
        shape.top = top
    if width is not None:
        shape.width = width
    if height is not None:
        shape.height = height
