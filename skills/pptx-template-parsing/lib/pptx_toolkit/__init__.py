"""pptx_toolkit: shared, proven helpers for parsing and generating PPTX decks.

Two skills consume this library:
  - pptx-template-parsing  (reads a template -> spec.json)
  - ppt-template-adaptation (spec.json + content_plan.json -> filled deck)

Import submodules directly, e.g.:
    from lib.pptx_toolkit import autofit, text, render, walk
or pull common names from the package root.
"""
from __future__ import annotations

from . import assets, autofit, content, deck, fonts, geometry, render, shapes, text, walk

# Frequently used names, re-exported for convenience.
from .geometry import emu_to_inch, emu_to_pt, inch_to_emu, pt_to_emu, resize_shape
from .text import replace_first_run_text, set_mixed_text, set_run_style, set_text
from .autofit import candidate_breaks, fit_body, fit_title, wrap_line
from .shapes import (
    add_centered_picture,
    delete_shape,
    hide_shape_visual,
    replace_shape_with_picture,
)
from .assets import build_mask_icon, crop_to_fill, ensure_png_from_svg, get_icon
from .deck import delete_slide, delete_slides
from .render import export_pdf, render_deck, render_pdf_to_pngs
from .fonts import measure_line_height_pt, measure_line_width_pt
from .content import apply_page, resolve_fill_plan
from .walk import (
    geometry_of,
    get_shape_by_path,
    is_group,
    iter_shapes,
    iter_top_level,
    placeholder_info,
    resolve_paths,
    shape_type_name,
    text_info,
)

__all__ = [
    "assets", "autofit", "content", "deck", "fonts", "geometry", "render", "shapes", "text", "walk",
    "emu_to_inch", "emu_to_pt", "inch_to_emu", "pt_to_emu", "resize_shape",
    "replace_first_run_text", "set_mixed_text", "set_run_style", "set_text",
    "candidate_breaks", "fit_body", "fit_title", "wrap_line",
    "add_centered_picture", "delete_shape", "hide_shape_visual", "replace_shape_with_picture",
    "build_mask_icon", "crop_to_fill", "ensure_png_from_svg", "get_icon",
    "delete_slide", "delete_slides",
    "export_pdf", "render_deck", "render_pdf_to_pngs",
    "measure_line_height_pt", "measure_line_width_pt",
    "apply_page", "resolve_fill_plan",
    "geometry_of", "get_shape_by_path", "is_group", "iter_shapes", "iter_top_level",
    "placeholder_info", "resolve_paths", "shape_type_name", "text_info",
]
