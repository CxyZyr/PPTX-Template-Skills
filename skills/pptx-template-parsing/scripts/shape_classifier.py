"""Per-shape role classification (the heart of automatic template parsing).

Builds a feature node tree for a slide, then assigns each shape a semantic role
using multiple signals in priority order — see references/classification-signals.md.
No single signal is trusted alone; every node records the signals that fired and
a confidence, so a human/agent can review against the rendered slide.

Roles: title, subtitle, body, label, number, section_number, logo, icon, image,
       background, chart, table, decoration, footer, unknown
"""
from __future__ import annotations

import pathlib
import re
import sys

SKILL_ROOT = pathlib.Path(__file__).resolve().parents[1]
REPO_ROOT = pathlib.Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(SKILL_ROOT))

from lib.pptx_toolkit.walk import (  # noqa: E402
    absolute_geometry_of_path,
    geometry_of,
    is_group,
    placeholder_info,
    picture_fill_info,
    shape_type_name,
    text_info,
    visual_info,
)

# ---- stock placeholder text patterns (strongest text signal) --------------
STOCK_PATTERNS = [
    ("title", re.compile(r"(点击|单击|请|在此)?.{0,4}(输入|添加).{0,2}标题|标题文字|大标题|主标题")),
    ("subtitle", re.compile(r"副标题|副标題|输入副标题|sub\s*title", re.I)),
    ("body", re.compile(
        r"(点击|单击|请|在此).{0,6}(输入|添加).{0,4}(文字|文本|内容|正文)"
        r"|文字内容|此处输入|添加文本|详细内容|内容描述|输入内容")),
    ("logo", re.compile(r"\b(your\s*)?logo\b|您的\s*logo|公司\s*logo", re.I)),
    ("contents_hint", re.compile(r"目录|CONTENTS|AGENDA|目\s*录", re.I)),
    ("section_number", re.compile(r"\bPART\s*\.?\s*\d+|第.{1,2}部分|第.{1,2}章|章节", re.I)),
]

PURE_NUMBER = re.compile(r"^\s*\d{1,4}\s*$")
SECTION_NUMBERISH = re.compile(r"^\s*0?\d\s*$")
YEARISH = re.compile(r"^\s*(19|20)\d{2}\s*$")
ENDING_HINT = re.compile(r"谢谢|感谢|THANKS?|THANK\s*YOU", re.I)
CONTENTS_HINT = STOCK_PATTERNS[4][1]

# ---- geometry thresholds (ratios of slide unless noted) -------------------
BG_AREA = 0.62            # >= this area ratio (and broadly covering) -> background
SMALL_DIM_IN = 1.3        # max dimension (inches) under which a shape is "small"
TINY_AREA = 0.0025        # below this area ratio -> decoration dot
ICON_ASPECT_LO, ICON_ASPECT_HI = 0.45, 2.2   # near-square-ish
UPPER_BAND = 0.30         # cy_ratio below this == upper region
BOTTOM_BAND = 0.88        # cy_ratio above this == bottom region (footer)
CORNER_X = 0.16           # within this of a vertical edge == corner-ish


def _stock_role(text_full: str):
    for role, pat in STOCK_PATTERNS:
        if pat.search(text_full):
            return role
    return None


def build_nodes(shapes, slide_w: int, slide_h: int) -> list[dict]:
    """Recursively build feature nodes (no roles yet) for a shape collection."""
    nodes = []
    for index, shape in enumerate(shapes):
        node = {
            "path": [index] if not isinstance(shapes, list) else [index],
            "shape_type": shape_type_name(shape),
            "placeholder": placeholder_info(shape),
            "picture_fill": picture_fill_info(shape),
            "geometry": geometry_of(shape, slide_w, slide_h),
            "text": text_info(shape),
            "visual": visual_info(shape),
            "is_group": is_group(shape),
            "children": [],
        }
        if node["is_group"]:
            node["children"] = build_nodes(shape.shapes, slide_w, slide_h)
        nodes.append(node)
    return nodes


def _fix_paths(nodes: list[dict], prefix: tuple = ()):
    """Rewrite each node['path'] to its absolute index path within the slide."""
    for index, node in enumerate(nodes):
        node["path"] = list(prefix + (index,))
        if node["children"]:
            _fix_paths(node["children"], tuple(node["path"]))


def _fix_absolute_geometry(nodes: list[dict], root_shapes, slide_w: int, slide_h: int) -> None:
    """Rewrite node geometry from group-local values to slide-level values."""
    for node in _flatten(nodes):
        geo = absolute_geometry_of_path(root_shapes, node["path"], slide_w, slide_h)
        if geo is not None:
            node["geometry"] = geo


def _flatten(nodes: list[dict]):
    for node in nodes:
        yield node
        if node["children"]:
            yield from _flatten(node["children"])


def _slide_max_font(nodes: list[dict]) -> float:
    sizes = [
        n["text"]["max_font_pt"]
        for n in _flatten(nodes)
        if n["text"] and n["text"]["max_font_pt"]
    ]
    return max(sizes) if sizes else 0.0


def _classify_text_shape(node, ctx) -> tuple[str, float, list[str]]:
    t = node["text"]
    g = node["geometry"]
    signals = []
    full = t["full"]

    stock = _stock_role(full)
    if stock == "contents_hint":
        return "title", 0.85, ["stock:contents"]
    if stock:
        return stock, 0.9, [f"stock:{stock}"]

    # numeric content
    if full and PURE_NUMBER.match(full):
        if YEARISH.match(full):
            return "label", 0.6, ["year"]
        size = t["max_font_pt"] or 0
        if SECTION_NUMBERISH.match(full) and size >= 28:
            return "section_number", 0.6, ["digit", "large"]
        if size >= 24:
            return "number", 0.7, ["pure_digit", "large"]
        return "label", 0.5, ["digit", "small"]

    size = t["max_font_pt"]
    upper = g["cy_ratio"] <= UPPER_BAND
    short = t["char_count"] <= 12
    multiline = t["n_paragraphs"] >= 2 or t["char_count"] > 40
    narrow_tall = g["aspect"] and g["aspect"] <= 0.35 and g["height_in"] >= 1.2

    # footer / page number: bottom band, short, small
    if g["cy_ratio"] >= BOTTOM_BAND and t["char_count"] <= 20 and (size is None or size <= 16):
        return "footer", 0.5, ["bottom", "short"]

    # Vertical/narrow side text is usually a design label, not a body slot.
    # Treat it as fillable optional text so generation does not pour long body
    # copy into a decorative rail.
    if (t.get("vertical") or narrow_tall) and t["char_count"] <= 24:
        sig = ["vertical" if t.get("vertical") else "narrow_tall", "short"]
        return "label", 0.6, sig

    if size is not None and ctx["max_font"]:
        ratio = size / ctx["max_font"]
        if ratio >= 0.92 and size >= 28 and upper:
            return "title", 0.75, ["max_font", "upper"]
        if ratio >= 0.92 and size >= 28:
            return "title", 0.6, ["max_font"]
        if 0.55 <= ratio < 0.92 and short and upper:
            return "subtitle", 0.55, ["mid_font", "upper", "short"]
        if multiline:
            return "body", 0.6, ["multiline"]
        if short:
            return "label", 0.5, ["short"]
        return "body", 0.45, ["text_fallback"]

    # font size unknown (inherited) -> rely on geometry + length
    if upper and short and g["width_in"] >= 3:
        return "title", 0.45, ["geo_upper_wide", "no_size"]
    if multiline:
        return "body", 0.45, ["multiline", "no_size"]
    if short:
        return "label", 0.4, ["short", "no_size"]
    return "body", 0.35, ["text_fallback", "no_size"]


def _classify_picture(node, ctx) -> tuple[str, float, list[str]]:
    g = node["geometry"]
    big = g["area_ratio"] >= BG_AREA
    near_origin = g["left_in"] <= 0.3 and g["top_in"] <= 0.3
    max_dim = max(g["width_in"], g["height_in"])
    corner = (g["cx_ratio"] <= CORNER_X or g["cx_ratio"] >= 1 - CORNER_X)

    if big and (near_origin or not g["on_canvas"] or g["area_ratio"] >= 0.85):
        return "background", 0.7, ["picture", "large", "origin" if near_origin else "cover"]
    if max_dim <= SMALL_DIM_IN and corner:
        return "logo", 0.55, ["picture", "small", "corner"]
    if max_dim <= SMALL_DIM_IN:
        return "icon", 0.45, ["picture", "small"]
    return "image", 0.7, ["picture", "content"]


def _classify_picture_fill(node, ctx) -> tuple[str, float, list[str]]:
    role, conf, sig = _classify_picture(node, ctx)
    return role, max(conf, 0.7), ["picture_fill"] + sig


def _classify_blank_shape(node, ctx) -> tuple[str, float, list[str]]:
    """AUTO_SHAPE / FREEFORM with no meaningful text."""
    g = node["geometry"]
    if not g["on_canvas"]:
        return "decoration", 0.6, ["off_canvas"]
    if g["area_ratio"] >= BG_AREA:
        return "background", 0.55, ["large_shape"]
    if g["area_ratio"] <= TINY_AREA:
        return "decoration", 0.6, ["tiny"]
    max_dim = max(g["width_in"], g["height_in"])
    aspect = g["aspect"] or 1
    if max_dim <= SMALL_DIM_IN and ICON_ASPECT_LO <= aspect <= ICON_ASPECT_HI:
        # near-square small vector -> likely an icon glyph (confirmed if grouped w/ text)
        return "icon", 0.4, ["small", "square"]
    return "decoration", 0.35, ["shape_fallback"]


def classify_node(node, ctx) -> None:
    """Assign role/confidence/signals to a single (non-group) node."""
    st = node["shape_type"]
    ph = node["placeholder"]

    # 1) hard shape-type rules
    if st == "TABLE":
        node.update(role="table", role_confidence=0.95, signals=["shape:table"])
        return
    if st == "CHART":
        node.update(role="chart", role_confidence=0.95, signals=["shape:chart"])
        return
    if st in ("LINE", "CONNECTOR"):
        node.update(role="decoration", role_confidence=0.5, signals=["line"])
        return

    # 2) real placeholder type (rare but authoritative)
    if ph:
        mapping = {
            "TITLE": "title", "CENTER_TITLE": "title", "SUBTITLE": "subtitle",
            "BODY": "body", "OBJECT": "body", "PICTURE": "image",
            "TABLE": "table", "CHART": "chart",
        }
        role = mapping.get(ph)
        if role:
            node.update(role=role, role_confidence=0.95, signals=[f"placeholder:{ph}"])
            return

    # 3) shape with bitmap fill (e.g. a circular cropped photo in an oval)
    if node.get("picture_fill"):
        role, conf, sig = _classify_picture_fill(node, ctx)
        node.update(role=role, role_confidence=conf, signals=sig)
        return

    # 4) text-bearing shapes (TEXT_BOX, or AUTO_SHAPE/FREEFORM carrying real text)
    if node["text"] and node["text"]["full"]:
        role, conf, sig = _classify_text_shape(node, ctx)
        node.update(role=role, role_confidence=conf, signals=sig)
        return

    # 5) pictures
    if st == "PICTURE":
        role, conf, sig = _classify_picture(node, ctx)
        node.update(role=role, role_confidence=conf, signals=sig)
        return

    # 6) blank shapes
    role, conf, sig = _classify_blank_shape(node, ctx)
    node.update(role=role, role_confidence=conf, signals=sig)


def _descendant_leaf_roles(node) -> list[str]:
    roles = []
    for c in node["children"]:
        if c["children"]:
            roles += _descendant_leaf_roles(c)
        else:
            roles.append(c.get("role"))
    return roles


def _promote_icons(node) -> None:
    """Within a content group, treat small near-square blank shapes as icons."""
    for c in node["children"]:
        if c["children"]:
            _promote_icons(c)
        elif c["role"] == "decoration" and "small" in (c.get("signals") or []) \
                and "square" in (c.get("signals") or []):
            c["role"] = "icon"
            c["role_confidence"] = 0.55
            c["signals"] = c["signals"] + ["in_card"]


def _classify_group(node) -> None:
    """A group's role is inferred from its descendants after they're classified.

    A group is a *card* only when it carries substantive text (title/subtitle/body).
    A group that holds only a number/label (e.g. an '01' badge circle) or only
    icons is NOT a card — this keeps nested badge groups from masking the real
    outer card and leaking card titles up to the slide title slot.
    """
    leaf_roles = _descendant_leaf_roles(node)
    substantive = any(r in ("title", "subtitle", "body") for r in leaf_roles)
    has_icon = any(r in ("icon", "image") for r in leaf_roles)
    has_minor_text = any(r in ("label", "number", "section_number") for r in leaf_roles)

    if substantive:
        _promote_icons(node)
        node.update(role="card", role_confidence=0.6,
                    signals=["group", "card+icon" if has_icon else "card"])
    elif has_icon:
        node.update(role="icon_group", role_confidence=0.4, signals=["group", "icons"])
    elif has_minor_text:
        node.update(role="badge", role_confidence=0.4, signals=["group", "badge"])
    else:
        node.update(role="decoration", role_confidence=0.3, signals=["group", "blank"])


def classify_tree(nodes: list[dict], ctx) -> None:
    for node in nodes:
        if node["is_group"]:
            classify_tree(node["children"], ctx)
            _classify_group(node)
        else:
            classify_node(node, ctx)


def classify_slide(shapes, slide_w: int, slide_h: int) -> list[dict]:
    """Top-level entry: build feature nodes and classify the whole slide tree.
    Returns the node list (each node carries role/role_confidence/signals)."""
    nodes = build_nodes(shapes, slide_w, slide_h)
    _fix_paths(nodes)
    _fix_absolute_geometry(nodes, shapes, slide_w, slide_h)
    ctx = {"max_font": _slide_max_font(nodes), "slide_w": slide_w, "slide_h": slide_h}
    classify_tree(nodes, ctx)
    return nodes
