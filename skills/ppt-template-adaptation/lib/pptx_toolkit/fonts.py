"""Font resolution and text measurement.

The proven approach (from scripts/generate_strategy_ppt.py) measures real glyph
widths/heights with Pillow against an installed CJK font, so autofit decisions
match what LibreOffice will actually render.

We default to Noto Sans CJK (confirmed installed in this environment) but fall
back gracefully so the toolkit never hard-crashes on a missing font file.
"""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

# Candidate font files, tried in order. First existing wins.
_BOLD_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
]
_REGULAR_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
]


def _first_existing(candidates: list[str]) -> str | None:
    for path in candidates:
        if Path(path).exists():
            return path
    return None


BOLD_FONT_FILE = _first_existing(_BOLD_CANDIDATES)
REGULAR_FONT_FILE = _first_existing(_REGULAR_CANDIDATES)


@lru_cache(maxsize=256)
def _load_font(bold: bool, pixel_size: int) -> ImageFont.FreeTypeFont:
    path = BOLD_FONT_FILE if bold else REGULAR_FONT_FILE
    if path is None:
        # Last resort: PIL default bitmap font (measurement still works, less precise).
        return ImageFont.load_default()
    return ImageFont.truetype(path, pixel_size)


def measure_line_width_pt(text: str, size_pt: float, *, bold: bool = False) -> float:
    """Width of a single (non-wrapping) line, in points."""
    pixel_size = max(1, round(size_pt * 96 / 72))
    font = _load_font(bold, pixel_size)
    bbox = font.getbbox(text or " ")
    return (bbox[2] - bbox[0]) * 72 / 96


def measure_line_height_pt(size_pt: float, *, bold: bool = False) -> float:
    """Representative glyph height (mixed CJK + latin), in points."""
    pixel_size = max(1, round(size_pt * 96 / 72))
    font = _load_font(bold, pixel_size)
    bbox = font.getbbox("战略Ag")
    return (bbox[3] - bbox[1]) * 72 / 96
