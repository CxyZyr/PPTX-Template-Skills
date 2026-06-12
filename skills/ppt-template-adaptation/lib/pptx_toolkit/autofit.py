"""Autofit: wrap, break, and size text to fit a shape using real measurement.

This is the heart of the proven layout quality in scripts/generate_strategy_ppt.py.
The strategy (per the playbook): try the template box first, widen/reposition,
try a better wrap, and only then reduce font size — never shrink as the first move.

Depends on geometry (conversions/resize), fonts (measurement), text (writing).
"""
from __future__ import annotations

import re

from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Pt

from .fonts import measure_line_height_pt, measure_line_width_pt
from .geometry import emu_to_pt, pt_to_emu, resize_shape
from .text import set_text


def wrap_line(text: str, size_pt: int, max_width_pt: float, *, bold: bool = False) -> list[str]:
    """Greedy character-level wrap (correct for CJK; fine for mixed CJK/latin)."""
    lines: list[str] = []
    current = ""
    for char in text:
        trial = current + char
        if not current or measure_line_width_pt(trial, size_pt, bold=bold) <= max_width_pt:
            current = trial
        else:
            lines.append(current)
            current = char
    if current:
        lines.append(current)
    return lines


def _rebalance_line_orphans(lines: list[str], size_pt: int, max_width_pt: float, *,
                            bold: bool = False) -> list[str]:
    """Avoid visually weak CJK wraps such as a line starting with `和`."""
    if len(lines) < 2:
        return lines
    bad_starts = set("，。；：、！？)]）】和与")
    bad_ends = set("不单和与及并等")
    out = list(lines)
    for i in range(1, len(out)):
        prev = out[i - 1]
        cur = out[i]
        if not prev or not cur or len(prev.strip()) <= 2:
            continue
        should_rebalance = len(cur.strip()) <= 2 or cur[0] in bad_starts
        if not should_rebalance:
            should_rebalance = prev[-1] in bad_ends
        if not should_rebalance:
            continue
        move_count = 2 if len(cur.strip()) <= 2 and len(prev.strip()) > 4 else 1
        moved = prev[-move_count:]
        trial = moved + cur
        if measure_line_width_pt(trial, size_pt, bold=bold) <= max_width_pt:
            out[i - 1] = prev[:-move_count]
            out[i] = trial
    return [line for line in out if line]


def candidate_breaks(text: str) -> list[str]:
    """Propose natural 2-line break points for a title (at punctuation, else midpoint)."""
    candidates = []

    def add_candidate(left: str, right: str) -> None:
        left = left.strip()
        right = right.strip()
        if left and right:
            candidate = f"{left}\n{right}"
            if candidate not in candidates:
                candidates.append(candidate)

    for token in (" | ", "|", "——", "·", "：", "，", " "):
        start = 0
        while True:
            idx = text.find(token, start)
            if idx < 0:
                break
            split_idx = idx + len(token)
            add_candidate(text[:split_idx], text[split_idx:])
            start = split_idx
    for match in re.finditer(r"(?:19|20)\d{2}年?", text):
        start, end = match.span()
        if start >= 3 and len(text) - start >= 4:
            add_candidate(text[:start], text[start:])
        if end < len(text) and end >= 4 and len(text) - end >= 3:
            add_candidate(text[:end], text[end:])
    if len(text) >= 10:
        half = len(text) // 2
        while 0 < half < len(text) and text[half - 1].isdigit() and text[half].isdigit():
            half += 1
        add_candidate(text[:half], text[half:])
    return candidates


def fit_title(
    shape,
    text: str,
    *,
    base_size_pt: int,
    min_size_pt: int,
    font_name=None,
    color=None,
    align=PP_ALIGN.LEFT,
    bold: bool = False,
    min_left: int | None = None,
    max_right: int | None = None,
    max_bottom: int | None = None,
    allow_break: bool = False,
    prefer_break: bool = False,
    respect_bounds: bool = False,
) -> None:
    """Fit a (usually short) title: prefer width expansion + optional 2-line break,
    reduce size only as needed. Resizes the shape box, then writes the text."""
    current_width_emu = max(1, abs(int(shape.width)))
    current_height_emu = max(1, abs(int(shape.height)))
    if min_left is not None and max_right is not None and align in (PP_ALIGN.CENTER, PP_ALIGN.DISTRIBUTE):
        max_width_emu = max_right - min_left
    else:
        max_width_emu = (max_right - shape.left) if max_right else current_width_emu
    if max_width_emu <= 0:
        max_width_emu = current_width_emu
    current_width_pt = max(1.0, emu_to_pt(current_width_emu))
    current_height_pt = max(1.0, emu_to_pt(current_height_emu))
    max_width_pt = max(1.0, emu_to_pt(max_width_emu))
    max_height_pt = emu_to_pt(max_bottom - shape.top) if max_bottom is not None and max_bottom > shape.top else None
    if respect_bounds:
        current_width_pt = min(current_width_pt, max_width_pt)
        if max_height_pt is not None:
            current_height_pt = min(current_height_pt, max_height_pt)
    best_text = text
    best_size = min_size_pt
    best_width = current_width_pt
    best_height = current_height_pt
    best_align = align
    fallback = None
    original_center = shape.left + shape.width // 2
    options = [text]
    if allow_break:
        breaks = candidate_breaks(text)
        if prefer_break:
            base_width = measure_line_width_pt(text, base_size_pt, bold=True)
            if base_width > max_width_pt - 10:
                options = breaks + [text]
            else:
                options.extend(breaks)
        else:
            options.extend(breaks)

    for candidate in options:
        lines = candidate.split("\n")
        candidate_fit = False
        for size_pt in range(base_size_pt, min_size_pt - 1, -2):
            line_width = max(measure_line_width_pt(line, size_pt, bold=True) for line in lines)
            line_height = measure_line_height_pt(size_pt, bold=True)
            # Pillow measurement can be slightly narrower than LibreOffice's
            # final font substitution. Keep a real safety margin so left titles
            # expand right instead of wrapping at the last character.
            width_margin = max(24, size_pt * 0.9)
            needed_width = min(max_width_pt, max(current_width_pt, line_width * 1.08 + width_margin))
            needed_height = max(current_height_pt, line_height * len(lines) * 1.18 + 10)
            height_ok = max_height_pt is None or needed_height <= max_height_pt
            if needed_width <= max_width_pt and line_width <= needed_width - 10:
                overflow = max(0, needed_height - (max_height_pt or needed_height))
                candidate_record = (candidate, size_pt, needed_width, needed_height, overflow)
                if fallback is None or overflow < fallback[-1]:
                    fallback = candidate_record
                if height_ok and (len(lines) == 1 or needed_height <= current_height_pt * 1.9):
                    best_text = candidate
                    best_size = size_pt
                    best_width = needed_width
                    best_height = needed_height
                    # DISTRIBUTE alignment with long text -> switch to CENTER
                    if align == PP_ALIGN.DISTRIBUTE and len(text) > 8:
                        best_align = PP_ALIGN.CENTER
                    else:
                        best_align = align
                    candidate_fit = True
                    break
        if candidate_fit:
            break
    if not candidate_fit and fallback is not None:
        best_text, best_size, best_width, best_height, _overflow = fallback

    new_width = max(1, pt_to_emu(best_width))
    if respect_bounds and max_right is not None:
        available_width = max_right - shape.left
        if available_width > 0:
            new_width = min(new_width, available_width)
    new_height = max(1, pt_to_emu(best_height))
    if respect_bounds and max_bottom is not None:
        available_height = max_bottom - shape.top
        if available_height > 0:
            new_height = min(new_height, available_height)
    new_left = None
    if min_left is not None and max_right is not None and align in (PP_ALIGN.CENTER, PP_ALIGN.DISTRIBUTE):
        bounds_width = max_right - min_left
        if bounds_width > 0:
            new_width = min(new_width, bounds_width)
            new_left = int(original_center - new_width / 2)
            new_left = max(min_left, min(new_left, max_right - new_width))
    resize_shape(shape, left=new_left, width=new_width, height=new_height)
    set_text(
        shape,
        best_text.split("\n"),
        font_name=font_name,
        size=Pt(best_size),
        bold=bold,
        color=color,
        align=best_align,
        line_spacing=1.0,
        space_after=0,
    )


def fit_body(
    shape,
    paragraphs: list[str] | str,
    *,
    base_size_pt: int,
    min_size_pt: int,
    font_name=None,
    color=None,
    align=PP_ALIGN.LEFT,
    bold: bool = False,
    max_right: int | None = None,
    respect_bounds: bool = False,
    margins=(4, 4, 4, 4),
    line_spacing=1.12,
    measure_width_emu: int | None = None,
    measure_height_emu: int | None = None,
    max_measure_width_emu: int | None = None,
    resize_width: bool = True,
    strict_height: bool = False,
    max_height_emu: int | None = None,
) -> None:
    """Fit body copy: wrap at the box width, grow width modestly if it helps,
    reduce size step by step until height fits. Writes the wrapped lines."""
    if isinstance(paragraphs, str):
        paragraphs = [paragraphs]
    paragraphs = [
        line.strip()
        for paragraph in paragraphs
        for line in str(paragraph).splitlines()
        if line.strip()
    ]
    measure_width = measure_width_emu or shape.width
    measure_height = measure_height_emu or shape.height
    current_width_pt = emu_to_pt(measure_width)
    current_height_pt = emu_to_pt(measure_height)
    height_limit_pt = current_height_pt if strict_height else current_height_pt * 1.25
    if max_measure_width_emu is not None:
        max_width_pt = emu_to_pt(max_measure_width_emu)
    else:
        max_width_pt = emu_to_pt((max_right - shape.left) if max_right else measure_width)
    if respect_bounds:
        current_width_pt = min(current_width_pt, max_width_pt)
    # Pillow metrics are systematically optimistic for CJK under LibreOffice
    # fallback fonts. Wrap against a conservative width, then disable renderer
    # auto-wrap after writing explicit lines so final output does not invent
    # one-character lines.
    render_slack = max(18, base_size_pt * 1.5)
    inner_max_width = max(20, max_width_pt - margins[0] - margins[2] - 8 - render_slack)
    best_lines = paragraphs
    best_size = min_size_pt
    best_width = current_width_pt
    fallback = None

    for size_pt in range(base_size_pt, min_size_pt - 1, -1):
        wrapped: list[str] = []
        for para in paragraphs:
            wrapped.extend(wrap_line(para, size_pt, inner_max_width, bold=bold))
        wrapped = _rebalance_line_orphans(
            wrapped,
            size_pt,
            inner_max_width,
            bold=bold,
        )
        if not wrapped:
            break
        line_height = measure_line_height_pt(size_pt, bold=bold)
        needed_height = line_height * len(wrapped) * line_spacing + margins[1] + margins[3] + 6
        longest = max(measure_line_width_pt(line, size_pt, bold=bold) for line in wrapped)
        needed_width = min(max_width_pt, max(current_width_pt, longest + margins[0] + margins[2] + 12))
        fallback = (wrapped, size_pt, needed_width)
        if needed_height <= height_limit_pt and longest <= needed_width - margins[0] - margins[2] - 4:
            best_lines = wrapped
            best_size = size_pt
            best_width = needed_width
            break
    else:
        # Even when height cannot be satisfied, preserve width-safe explicit
        # CJK wrapping. A single fallback paragraph renders as horizontal
        # overflow in narrow template columns.
        if fallback is not None:
            best_lines, best_size, best_width = fallback

    if resize_width:
        resize_shape(shape, width=pt_to_emu(best_width))
    if strict_height:
        target_height = max_height_emu if max_height_emu is not None else pt_to_emu(height_limit_pt)
        if target_height and target_height > 0:
            resize_shape(shape, height=max(1, int(target_height)))
    set_text(
        shape,
        best_lines,
        font_name=font_name,
        size=Pt(best_size),
        bold=bold,
        color=color,
        align=align,
        margins=margins,
        line_spacing=line_spacing,
        space_after=0,
    )
    try:
        shape.text_frame.word_wrap = False
        if strict_height:
            shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    except Exception:
        pass
