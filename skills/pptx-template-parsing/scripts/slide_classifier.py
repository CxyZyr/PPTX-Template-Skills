"""Slide-level analysis: detect repeated card structures, infer the slide's
business role, and distill a ready-to-fill `fill_plan` from classified shapes.

Consumes the node tree from shape_classifier.classify_slide.
See references/slide-roles.md and references/spec-schema.md.
"""
from __future__ import annotations

import re

CONTENTS_RE = re.compile(r"目录|CONTENTS|AGENDA", re.I)
ENDING_RE = re.compile(r"谢谢|感谢|THANKS?|THANK\s*YOU", re.I)
TEMPLATE_META_RE = re.compile(
    r"(PPT|幻灯片|演示文稿).{0,8}(模板|模版)|(?:模板|模版).{0,8}(PPT|幻灯片|演示文稿)",
    re.I,
)
BODY_PLACEHOLDER_RE = re.compile(
    r"(点击|单击|请|在此).{0,6}(输入|添加).{0,4}(文字|文本|内容|正文)"
    r"|文字内容|详细内容|内容描述|输入内容"
)
TITLE_PLACEHOLDER_RE = re.compile(r"标题|title|heading", re.I)

TEXT_ROLES = {"title", "subtitle", "body", "label", "number", "section_number"}
DECOR_ROLES = {"decoration", "background"}
ROW_BUCKET_EMU = 457200  # 0.5 inch; groups visually same-row cards together.


def _flatten(nodes):
    for n in nodes:
        yield n
        if n["children"]:
            yield from _flatten(n["children"])


def _leaves(nodes):
    for n in nodes:
        if n["children"]:
            yield from _leaves(n["children"])
        else:
            yield n


def _font_of(node):
    return (node["text"]["max_font_pt"] or 0) if node.get("text") else 0


def _path_tuple(path) -> tuple:
    return tuple(path or [])


def _same_path(a, b) -> bool:
    return bool(a and b) and _path_tuple(a) == _path_tuple(b)


def _residual_card_text_paths(leaves: list[dict], used_paths: set[tuple]) -> list[list[int]]:
    """Text leaves inside a card that are neither filled nor structural.

    Some templates put small helper labels (for example an English category
    tag) inside each card group. If they are not exposed in the card slot,
    generation has no contract for clearing them and stale template text leaks
    into the output deck.
    """
    residual = []
    for node in leaves:
        if _path_tuple(node.get("path")) in used_paths:
            continue
        if node.get("role") not in {"title", "subtitle", "body", "label"}:
            continue
        text = node.get("text") or {}
        if not (text.get("sample") or "").strip():
            continue
        residual.append(node["path"])
    return residual


def _center(geo: dict) -> tuple[float, float]:
    return (
        geo["left"] + geo["width"] / 2,
        geo["top"] + geo["height"] / 2,
    )


def _right(geo: dict) -> int:
    return geo["left"] + geo["width"]


def _bottom(geo: dict) -> int:
    return geo["top"] + geo["height"]


def _union_geometry(nodes: list[dict]) -> dict:
    """Return a lightweight geometry box that can sort virtual cards."""
    geos = [n["geometry"] for n in nodes if n and n.get("geometry")]
    left = min(g["left"] for g in geos)
    top = min(g["top"] for g in geos)
    right = max(_right(g) for g in geos)
    bottom = max(_bottom(g) for g in geos)
    return {
        "left": left,
        "top": top,
        "width": right - left,
        "height": bottom - top,
    }


def _common_parent_path(a: list[int], b: list[int]) -> list[int]:
    prefix = []
    for left, right in zip(a or [], b or []):
        if left != right:
            break
        prefix.append(left)
    return prefix


def _is_mixed_title_body_slot(node) -> bool:
    text = node.get("text") or {}
    full = text.get("full") or ""
    paragraphs = [p.strip() for p in full.splitlines() if p.strip()]
    if len(paragraphs) < 2:
        return False
    first = paragraphs[0]
    rest = "\n".join(paragraphs[1:])
    first_compact = re.sub(r"\s+", "", first)
    rest_compact = re.sub(r"\s+", "", rest)
    if not first_compact or not rest_compact:
        return False
    if len(first_compact) > 18:
        return False
    first_is_title_placeholder = bool(TITLE_PLACEHOLDER_RE.search(first))
    if BODY_PLACEHOLDER_RE.search(first) and not first_is_title_placeholder:
        return False
    if first_is_title_placeholder:
        return bool(BODY_PLACEHOLDER_RE.search(rest) or len(rest_compact) >= 6)
    return bool(BODY_PLACEHOLDER_RE.search(rest))


def _is_slide_title_node(node: dict) -> bool:
    """Main page titles should not be reinterpreted as card titles."""
    if node.get("role") != "title":
        return False
    geo = node.get("geometry") or {}
    size = _font_of(node)
    return geo.get("top_in", 99) <= 1.25 and size >= 24


def _is_page_header_title_candidate(node: dict) -> bool:
    """Prefer the template's page header over large content placeholders."""
    if node.get("role") != "title":
        return False
    if not ((node.get("text") or {}).get("sample") or "").strip():
        return False
    geo = node.get("geometry") or {}
    return (
        geo.get("top_in", 99) <= 1.15
        and geo.get("height_in", 99) <= 0.9
        and geo.get("width_in", 0) >= 1.0
    )


def _is_page_header_section_number_candidate(node: dict) -> bool:
    if node.get("role") != "section_number":
        return False
    geo = node.get("geometry") or {}
    return (
        geo.get("top_in", 99) <= 1.15
        and geo.get("height_in", 99) <= 0.9
        and geo.get("width_in", 0) >= 0.5
    )


def _overlap_ratio(a: dict, b: dict) -> float:
    overlap = max(0, min(_right(a), _right(b)) - max(a["left"], b["left"]))
    return overlap / max(1, min(a["width"], b["width"]))


def _same_text_column(title_geo: dict, body_geo: dict) -> bool:
    if _overlap_ratio(title_geo, body_geo) >= 0.28:
        return True
    left_delta_in = abs(title_geo["left_in"] - body_geo["left_in"])
    return left_delta_in <= 0.35


def _vertical_overlap_ratio(a: dict, b: dict) -> float:
    overlap = max(0, min(_bottom(a), _bottom(b)) - max(a["top"], b["top"]))
    return overlap / max(1, min(a["height"], b["height"]))


def _pair_score(title_node: dict, body_node: dict) -> float:
    title_geo = title_node["geometry"]
    body_geo = body_node["geometry"]
    y_gap = max(0, body_geo["top"] - _bottom(title_geo)) / max(1, title_geo["height"])
    x_gap = abs(title_geo["left"] - body_geo["left"]) / max(1, title_geo["width"])
    overlap_bonus = _overlap_ratio(title_geo, body_geo)
    return y_gap * 2.0 + x_gap - overlap_bonus


def _row_major_key(geo: dict) -> tuple[int, int]:
    return (round(geo.get("top", 0) / ROW_BUCKET_EMU), geo.get("left", 0))


def _default_text_action(node) -> str:
    """What generation should do if this text slot is not explicitly filled."""
    text = (node.get("text") or {}).get("full") or ""
    signals = node.get("signals") or []
    role = node.get("role")
    compact = re.sub(r"\s+", "", text)

    if TEMPLATE_META_RE.search(text):
        return "clear"
    if any(s.startswith("stock:") for s in signals):
        return "clear"
    if role == "footer":
        return "preserve"
    if role in {"title", "subtitle", "label", "number", "section_number"} and re.fullmatch(r"[A-Z]", compact):
        return "preserve"
    if role == "label" and any(s in signals for s in ("vertical", "narrow_tall", "year", "digit", "small")):
        return "preserve"
    return "clear"


def _text_style(node) -> dict | None:
    """Carry template text styling into fill_plan slots.

    Raw shape nodes already record style, but generation mostly consumes
    fill_plan. If slot-level style is omitted, theme colors such as
    BACKGROUND_1 can be lost when the text frame is rebuilt. Mixed-format
    text boxes also need paragraph/run styles so generation can replace text
    in-place without flattening numbers, titles, and bodies to one format.
    """
    text = node.get("text") or {}
    style = {}
    if text.get("max_font_pt") is not None:
        style["font_pt"] = text["max_font_pt"]
    if text.get("min_font_pt") is not None:
        style["min_font_pt"] = text["min_font_pt"]
    if text.get("font_name"):
        style["font_name"] = text["font_name"]
    if text.get("bold") is not None:
        style["bold"] = text["bold"]
    elif text.get("any_bold"):
        style["bold"] = True
    if text.get("italic") is not None:
        style["italic"] = text["italic"]
    if text.get("align"):
        style["align"] = text["align"]
    if text.get("vertical") is not None:
        style["vertical"] = bool(text.get("vertical"))
    if text.get("color"):
        style["color"] = text["color"]
    if text.get("paragraphs"):
        style["paragraphs"] = text["paragraphs"]
    return style or None


def _visual_color(node: dict | None) -> dict | None:
    visual = (node or {}).get("visual") or {}
    for key in ("fill", "line"):
        color = (visual.get(key) or {}).get("color")
        if color:
            return color
    return None


def _rgb_tuple(color: dict | None) -> tuple[int, int, int] | None:
    if not color:
        return None
    rgb = color.get("rgb")
    if rgb:
        value = str(rgb).lstrip("#")
        if len(value) == 6:
            try:
                return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
            except ValueError:
                return None
    theme = str(color.get("theme_color") or "").upper()
    if theme in {"BACKGROUND_1", "LIGHT_1"}:
        return 255, 255, 255
    if theme in {"TEXT_1", "DARK_1"}:
        return 0, 0, 0
    return None


def _contrast_icon_color(frame_color: dict | None) -> dict | None:
    rgb = _rgb_tuple(frame_color)
    if rgb is None:
        return None
    r, g, b = rgb
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return {"type": "RGB", "rgb": "FFFFFF" if luminance < 150 else "000000"}


def _node_tree_color(node: dict | None, flat: list[dict]) -> dict | None:
    """Find a visible foreground color on a node or its descendants."""
    if not node:
        return None
    prefix = _path_tuple(node.get("path"))
    descendants = [
        n for n in flat
        if _path_tuple(n.get("path"))[: len(prefix)] == prefix
    ]
    # Small descendants are usually glyph strokes inside an icon group.
    descendants.sort(key=lambda n: (n.get("geometry") or {}).get("area_ratio", 1))
    for candidate in descendants:
        color = _visual_color(candidate)
        if color:
            return color
    return None


def _icon_style_for_slot(slot: dict, by_path: dict[tuple, dict], flat: list[dict]) -> dict | None:
    """Infer replacement icon foreground from the original template glyph."""
    for overlay_path in slot.get("icon_overlays") or []:
        color = _node_tree_color(by_path.get(_path_tuple(overlay_path)), flat)
        if color:
            return {"color": color, "source": "overlay"}

    icon_node = by_path.get(_path_tuple(slot.get("icon")))
    if icon_node and not _is_icon_frame(icon_node):
        color = _node_tree_color(icon_node, flat)
        if color:
            return {"color": color, "source": "icon"}

    if icon_node and _is_icon_frame(icon_node):
        contrast = _contrast_icon_color(_visual_color(icon_node))
        if contrast:
            return {"color": contrast, "source": "frame_contrast"}
    return None


# ---------------------------------------------------------------------------
# Card / card-list detection
# ---------------------------------------------------------------------------
def _is_card_container(node) -> bool:
    """A group that holds child groups which are themselves cards."""
    return any(c["is_group"] and c.get("role") == "card" for c in node["children"])


def _is_decorative_side_label(label_node: dict, body_node: dict | None) -> bool:
    if not label_node or not body_node:
        return False
    label_geo = label_node.get("geometry") or {}
    body_geo = body_node.get("geometry") or {}
    text = (label_node.get("text") or {}).get("sample") or ""
    if len(re.sub(r"\s+", "", text)) > 8:
        return False
    if label_geo.get("width", 0) > body_geo.get("width", 0) * 0.55:
        return False
    if label_geo.get("height", 0) > body_geo.get("height", 0) * 0.8:
        return False
    if label_geo.get("left", 0) >= body_geo.get("left", 0):
        return False
    return _vertical_overlap_ratio(label_geo, body_geo) >= 0.35


def _extract_slots(card_node) -> dict:
    """Pick title/body/icon/number slot paths from a card subtree's leaves."""
    leaves = list(_leaves([card_node]))
    title = sorted(
        [l for l in leaves if l["role"] in ("title", "subtitle")],
        key=_font_of, reverse=True,
    )
    label = sorted([l for l in leaves if l["role"] == "label"], key=_font_of, reverse=True)
    body = [l for l in leaves if l["role"] == "body"]
    icon = [l for l in leaves if l["role"] == "icon"]
    number = [l for l in leaves if l["role"] in ("number", "section_number")]

    body_node = body[0] if body else None
    label_title_candidates = [
        node for node in label
        if not _is_decorative_side_label(node, body_node)
    ]
    title_node = title[0] if title else (label_title_candidates[0] if label_title_candidates else None)
    if body_node and title_node is None and _is_mixed_title_body_slot(body_node):
        title_node = body_node
    if body_node is None and title_node and _is_mixed_title_body_slot(title_node):
        body_node = title_node
    if title_node and body_node:
        title_body_mode = "shared_textbox" if _same_path(title_node["path"], body_node["path"]) else "separate_textboxes"
    else:
        title_body_mode = None
    used_paths = {
        _path_tuple(node["path"])
        for node in (title_node, body_node, icon[0] if icon else None, number[0] if number else None)
        if node and node.get("path")
    }
    slot = {
        "group_path": card_node["path"],
        "title": title_node["path"] if title_node else None,
        "title_sample": title_node["text"]["sample"] if title_node and title_node.get("text") else None,
        "title_style": _text_style(title_node) if title_node else None,
        "body": body_node["path"] if body_node else None,
        "body_sample": body_node["text"]["sample"] if body_node and body_node.get("text") else None,
        "body_style": _text_style(body_node) if body_node else None,
        "title_body_mode": title_body_mode,
        "icon": icon[0]["path"] if icon else None,
        "number": number[0]["path"] if number else None,
        "residual_text": _residual_card_text_paths(leaves, used_paths),
    }
    return slot


def _similar_dimension_values(values: list[int], tolerance: float = 0.28) -> bool:
    if len(values) < 2:
        return False
    lo, hi = min(values), max(values)
    return hi > 0 and (hi - lo) <= hi * tolerance


def _same_row_or_column_geometries(geometries: list[dict]) -> bool:
    if len(geometries) < 2:
        return False
    centers_x = [g["left"] + g["width"] / 2 for g in geometries]
    centers_y = [g["top"] + g["height"] / 2 for g in geometries]
    widths = [g["width"] for g in geometries]
    heights = [g["height"] for g in geometries]
    same_row = max(centers_y) - min(centers_y) <= max(heights) * 0.35
    same_column = max(centers_x) - min(centers_x) <= max(widths) * 0.35
    return same_row or same_column


def _repeated_text_only_slots(card_node) -> list[dict]:
    """Split a grouped list of peer text boxes into one virtual card per box.

    A contents page often stores all agenda items in one PowerPoint group, with
    each item as a multi-paragraph text box (`01` / title / subtitle). Treating
    the whole group as one card loses the remaining items. Split only when the
    direct children look like repeated peer slots, not title/body parts of one
    card.
    """
    text_nodes = [
        child for child in card_node.get("children", [])
        if not child.get("children")
        and child.get("text")
        and child.get("role") in {"title", "subtitle", "label"}
        and (child["text"].get("sample") or "").strip()
    ]
    if len(text_nodes) < 2:
        return []

    geometries = [node["geometry"] for node in text_nodes if node.get("geometry")]
    if len(geometries) != len(text_nodes):
        return []
    if not _same_row_or_column_geometries(geometries):
        return []
    if not (
        _similar_dimension_values([g["width"] for g in geometries])
        and _similar_dimension_values([g["height"] for g in geometries], tolerance=0.18)
    ):
        return []

    slots = []
    for node in sorted(text_nodes, key=lambda n: _row_major_key(n["geometry"])):
        slots.append({
            "group_path": node["path"],
            "title": node["path"],
            "title_sample": node["text"]["sample"],
            "title_style": _text_style(node),
            "body": None,
            "body_sample": None,
            "body_style": None,
            "title_body_mode": None,
            "icon": None,
            "number": None,
            "card_geometry": node["geometry"],
            "residual_text": [],
        })
    return slots


def _repeated_number_body_slots(card_node) -> list[dict]:
    """Split grouped number/body rows into one virtual card per row."""
    leaves = [
        node for node in _leaves([card_node])
        if node.get("text")
        and (node["text"].get("sample") or "").strip()
        and node.get("geometry")
    ]
    body_nodes = [n for n in leaves if n.get("role") == "body"]
    number_nodes = [n for n in leaves if n.get("role") in {"number", "section_number"}]
    if len(body_nodes) < 2 or len(number_nodes) < 2:
        return []

    pairs = []
    used_numbers: set[tuple] = set()
    for body in sorted(body_nodes, key=lambda n: _row_major_key(n["geometry"])):
        body_geo = body["geometry"]
        viable = [
            number for number in number_nodes
            if _path_tuple(number["path"]) not in used_numbers
            and number["geometry"].get("left", 0) < body_geo.get("left", 0)
            and _vertical_overlap_ratio(number["geometry"], body_geo) >= 0.35
        ]
        if not viable:
            continue
        number = max(viable, key=lambda n: _vertical_overlap_ratio(n["geometry"], body_geo))
        used_numbers.add(_path_tuple(number["path"]))
        pairs.append((number, body))

    if len(pairs) < 2:
        return []
    if len(pairs) != min(len(body_nodes), len(number_nodes)):
        return []

    slots = []
    for number, body in pairs:
        slots.append({
            "group_path": body["path"],
            "title": None,
            "title_sample": None,
            "title_style": None,
            "body": body["path"],
            "body_sample": body["text"]["sample"],
            "body_style": _text_style(body),
            "title_body_mode": None,
            "icon": None,
            "number": number["path"],
            "number_sample": number["text"]["sample"] if number.get("text") else None,
            "card_geometry": _union_geometry([number, body]),
            "residual_text": [],
        })
    return sorted(slots, key=lambda s: _row_major_key(s["card_geometry"]))


def _is_container_scale_card(card_node) -> bool:
    """Large layout containers should not be treated as one semantic card."""
    geo = card_node.get("geometry") or {}
    if geo.get("area_ratio", 0) < 0.28:
        return False
    text_leaves = [
        node for node in _leaves([card_node])
        if node.get("text")
        and (node["text"].get("sample") or "").strip()
        and node.get("role") in {"title", "subtitle", "body", "label"}
        and node.get("geometry")
    ]
    if len(text_leaves) < 4:
        return False
    body_like_count = sum(1 for node in text_leaves if node.get("role") == "body")
    if body_like_count < 2:
        return False
    centers_x = [
        node["geometry"]["left"] + node["geometry"]["width"] / 2
        for node in text_leaves
    ]
    centers_y = [
        node["geometry"]["top"] + node["geometry"]["height"] / 2
        for node in text_leaves
    ]
    multi_column = max(centers_x) - min(centers_x) >= geo.get("width", 0) * 0.28
    multi_row = max(centers_y) - min(centers_y) >= geo.get("height", 0) * 0.28
    return multi_column and multi_row


def _parallel_peer_text_nodes(nodes: list[dict]) -> list[dict]:
    """Return repeated peer text boxes that should stay together as free slots."""
    text_nodes = [
        node for node in nodes
        if node.get("text")
        and (node["text"].get("sample") or "").strip()
        and node.get("geometry")
    ]
    if len(text_nodes) < 2:
        return []
    geometries = [node["geometry"] for node in text_nodes]
    if not _same_row_or_column_geometries(geometries):
        return []
    if not (
        _similar_dimension_values([g["width"] for g in geometries])
        and _similar_dimension_values([g["height"] for g in geometries], tolerance=0.18)
    ):
        return []
    return sorted(text_nodes, key=lambda n: _row_major_key(n["geometry"]))


def collect_cards(nodes) -> list[dict]:
    """Find leaf 'card' groups anywhere in the tree and extract their slots."""
    cards = []

    def walk(ns):
        for n in ns:
            if n["is_group"] and n.get("role") == "card" and not _is_card_container(n):
                expanded = _repeated_number_body_slots(n) or _repeated_text_only_slots(n)
                if expanded:
                    cards.extend(expanded)
                elif _is_container_scale_card(n):
                    walk(n.get("children") or [])
                else:
                    cards.append(_extract_slots(n))
            elif n["children"]:
                walk(n["children"])

    walk(nodes)
    cards.sort(key=lambda slot: _row_major_key(slot.get("card_geometry") or {}))
    return cards


def collect_text_pair_cards(nodes) -> list[dict]:
    """Detect repeated title/body text pairs even when PPT did not group them.

    Some templates draw repeated card rows as loose text boxes, or put multiple
    title/body pairs inside one large group. If parsing only looks for group
    cards, those meaningful boxes are later treated as unfilled placeholders and
    get cleared. This pass promotes repeated text pairs to card slots.
    """
    leaves = list(_leaves(nodes))
    title_nodes = [
        n for n in leaves
        if n.get("role") in ("title", "subtitle")
        and n.get("text")
        and not _is_slide_title_node(n)
    ]
    body_nodes = [
        n for n in leaves
        if n.get("role") == "body"
        and n.get("text")
    ]
    mixed_body_nodes = [
        n for n in body_nodes
        if _is_mixed_title_body_slot(n)
    ]
    used_bodies: set[tuple] = set()
    slots: list[dict] = []

    def add_slot(title_node, body_node, *, shared: bool) -> None:
        parent = _common_parent_path(title_node["path"], body_node["path"])
        group_path = parent if parent else title_node["path"]
        slot = {
            "group_path": group_path,
            "title": title_node["path"],
            "title_sample": title_node["text"]["sample"],
            "title_style": _text_style(title_node),
            "body": body_node["path"],
            "body_sample": body_node["text"]["sample"],
            "body_style": _text_style(body_node),
            "title_body_mode": "shared_textbox" if shared else "separate_textboxes",
            "icon": None,
            "number": None,
            "card_geometry": _union_geometry([title_node, body_node]),
            "residual_text": [],
        }
        slots.append(slot)

    for title_node in sorted(title_nodes + mixed_body_nodes, key=lambda n: (n["geometry"]["top"], n["geometry"]["left"])):
        title_path = _path_tuple(title_node["path"])
        if _is_mixed_title_body_slot(title_node):
            add_slot(title_node, title_node, shared=True)
            used_bodies.add(title_path)
            continue

        title_geo = title_node["geometry"]
        viable = []
        for body_node in body_nodes:
            body_path = _path_tuple(body_node["path"])
            if body_path in used_bodies:
                continue
            body_geo = body_node["geometry"]
            if body_geo["top"] < title_geo["top"]:
                continue
            if body_geo["top_in"] - title_geo["top_in"] > 1.25:
                continue
            if not _same_text_column(title_geo, body_geo):
                continue
            viable.append(body_node)
        if not viable:
            continue
        body_node = min(viable, key=lambda n: _pair_score(title_node, n))
        used_bodies.add(_path_tuple(body_node["path"]))
        add_slot(title_node, body_node, shared=False)

    # A single title/body pair can be a page subtitle/body pattern rather than
    # a repeated card layout. Promote only repeated structures.
    if len(slots) < 2:
        return []
    slots.sort(key=lambda s: _row_major_key(s["card_geometry"]))
    return slots


def _pair_key(slot: dict) -> tuple[tuple, tuple]:
    return (_path_tuple(slot.get("title")), _path_tuple(slot.get("body")))


def merge_card_slots(group_slots: list[dict], text_slots: list[dict]) -> list[dict]:
    """Prefer text-pair coverage, but preserve group/card icon metadata."""
    if not text_slots:
        return group_slots
    merged = list(text_slots)
    by_pair = {_pair_key(slot): slot for slot in merged}
    for group_slot in group_slots:
        key = _pair_key(group_slot)
        text_slot = by_pair.get(key)
        if text_slot:
            for meta_key in ("icon", "icon_overlays", "number", "residual_text"):
                value = group_slot.get(meta_key)
                if value and not text_slot.get(meta_key):
                    text_slot[meta_key] = value
            if group_slot.get("group_path") and len(group_slot["group_path"]) < len(text_slot["group_path"]):
                text_slot["group_path"] = group_slot["group_path"]
            continue
        merged.append(group_slot)
    merged.sort(key=lambda s: _row_major_key(s.get("card_geometry") or {}))
    return merged


def _is_under(path, prefix) -> bool:
    path = _path_tuple(path)
    prefix = _path_tuple(prefix)
    return path[: len(prefix)] == prefix


def _is_external_icon_candidate(node, card_prefixes: set[tuple]) -> bool:
    if any(_is_under(node["path"], prefix) for prefix in card_prefixes):
        return False
    role = node.get("role")
    if role in ("icon", "icon_group"):
        return True
    # Some PPT vector glyphs are imported as tiny freeforms and are too small
    # for the general classifier. They can still be a card icon when spatially
    # paired with a card.
    geo = node.get("geometry") or {}
    if role == "decoration" and node.get("shape_type") == "GROUP":
        return (
            geo.get("area_ratio", 0) >= 0.00045
            and 0.45 <= (geo.get("aspect") or 0) <= 2.2
            and max(geo.get("width_in", 0), geo.get("height_in", 0)) <= 0.5
        )
    if role == "decoration" and node.get("shape_type") in ("FREEFORM", "GROUP"):
        return (
            geo.get("area_ratio", 0) >= 0.0015
            and 0.45 <= (geo.get("aspect") or 0) <= 2.2
            and max(geo.get("width_in", 0), geo.get("height_in", 0)) <= 1.1
        )
    return False


def _is_icon_overlay_candidate(node: dict) -> bool:
    """Small non-text glyphs that may sit on top of an icon frame."""
    if (node.get("text") or {}).get("full"):
        return False
    role = node.get("role")
    if role not in ("icon", "icon_group", "decoration"):
        return False
    if node.get("shape_type") not in ("FREEFORM", "GROUP", "AUTO_SHAPE", "PICTURE"):
        return False
    geo = node.get("geometry") or {}
    if not geo.get("on_canvas", True):
        return False
    return (
        geo.get("area_ratio", 0) >= 0.00015
        and geo.get("area_ratio", 0) <= 0.035
        and max(geo.get("width_in", 0), geo.get("height_in", 0)) <= 1.2
    )


def _candidate_score(card_geo: dict, icon_node: dict) -> float:
    icon_geo = icon_node["geometry"]
    card_cx, card_cy = _center(card_geo)
    icon_cx, icon_cy = _center(icon_geo)
    y_score = abs(icon_cy - card_cy) / max(1, card_geo["height"])
    x_gap = max(0, abs(icon_cx - card_cx) - card_geo["width"] / 2)
    x_score = x_gap / max(1, card_geo["width"])
    area_score = icon_geo.get("area_ratio", 0) / max(0.0001, card_geo.get("area_ratio", 0.0001))
    role_penalty = {"icon_group": 0.0, "decoration": 0.02, "icon": 0.12}.get(icon_node.get("role"), 0.2)
    # Prefer same-row icons first, then nearby/smaller glyphs. This prevents a
    # large decorative circle from winning over the actual small glyph inside it.
    return y_score * 3.0 + x_score * 0.35 + area_score * 0.8 + role_penalty


def _icon_scale_ok(card_geo: dict, icon_geo: dict) -> bool:
    """Reject media-sized shapes when looking for small semantic card icons."""
    return (
        icon_geo["width"] <= card_geo["width"] * 0.75
        and icon_geo["height"] <= card_geo["height"] * 1.35
    )


def _is_icon_left_of_card(card_geo: dict, icon_geo: dict) -> bool:
    card_cx, card_cy = _center(card_geo)
    icon_cx, icon_cy = _center(icon_geo)
    if icon_cx >= card_geo["left"]:
        return False
    if abs(icon_cy - card_cy) > card_geo["height"] * 0.7:
        return False
    if card_geo["left"] - icon_cx > card_geo["width"] * 0.55:
        return False
    return True


def _is_icon_above_card(card_geo: dict, icon_geo: dict) -> bool:
    icon_cx, _ = _center(icon_geo)
    if not (card_geo["left"] <= icon_cx <= _right(card_geo)):
        return False
    vertical_gap = card_geo["top"] - _bottom(icon_geo)
    if vertical_gap < -card_geo["height"] * 0.12:
        return False
    if vertical_gap > card_geo["height"] * 0.55:
        return False
    return True


def _is_icon_below_card(card_geo: dict, icon_geo: dict) -> bool:
    icon_cx, _ = _center(icon_geo)
    if not (card_geo["left"] <= icon_cx <= _right(card_geo)):
        return False
    vertical_gap = icon_geo["top"] - _bottom(card_geo)
    if vertical_gap < -card_geo["height"] * 0.12:
        return False
    if vertical_gap > card_geo["height"] * 0.55:
        return False
    return True


def _is_near_card_icon(card_geo: dict, icon_node: dict) -> bool:
    icon_geo = icon_node["geometry"]
    if not _icon_scale_ok(card_geo, icon_geo):
        return False
    return (
        _is_icon_left_of_card(card_geo, icon_geo)
        or _is_icon_above_card(card_geo, icon_geo)
        or _is_icon_below_card(card_geo, icon_geo)
    )


def _is_icon_frame(node: dict) -> bool:
    """True when a shape is the visual slot/frame where an icon image belongs."""
    geo = node.get("geometry") or {}
    if node.get("role") != "icon":
        return False
    if node.get("shape_type") in ("AUTO_SHAPE", "PICTURE"):
        return geo.get("area_ratio", 0) >= 0.003
    return False


def _inside_frame(candidate: dict, frame: dict) -> bool:
    geo = candidate.get("geometry") or {}
    frame_geo = frame.get("geometry") or {}
    cx, cy = _center(geo)
    pad_x = frame_geo["width"] * 0.12
    pad_y = frame_geo["height"] * 0.12
    return (
        frame_geo["left"] - pad_x <= cx <= frame_geo["left"] + frame_geo["width"] + pad_x
        and frame_geo["top"] - pad_y <= cy <= frame_geo["top"] + frame_geo["height"] + pad_y
        and geo.get("area_ratio", 0) < frame_geo.get("area_ratio", 0)
    )


def _prune_nested_paths(paths: list[list[int]]) -> list[list[int]]:
    """Keep the outermost overlay paths so deleting a group also removes children."""
    kept: list[tuple] = []
    for path in sorted((_path_tuple(p) for p in paths), key=lambda p: (len(p), p)):
        if any(path[: len(parent)] == parent for parent in kept):
            continue
        kept.append(path)
    return [list(p) for p in kept]


def _is_text_badge_group(node: dict) -> bool:
    """True when a group already carries its own text, so it is an occupied
    visual element rather than an empty icon slot.

    Templates often draw a decorated circle that holds its own concentric
    label: a step number (``01/02/03``) or a statistic ring (``85%``). Either
    way the text is the content, and dropping an icon image into the group
    overlaps it. A real, fillable icon frame holds no text of its own, so the
    presence of any non-empty textual leaf marks the group as already taken.
    """
    if not node.get("is_group"):
        return False
    for leaf in _leaves([node]):
        if leaf.get("role") not in {"title", "subtitle", "body", "label",
                                    "number", "section_number"}:
            continue
        if (leaf.get("text") or {}).get("sample", "").strip():
            return True
    return False


def associate_external_card_icons(nodes, card_slots) -> None:
    """Attach standalone icon/icon_group shapes to nearby cards.

    Some templates keep the card text in one group and draw the icon as a
    separate shape to the left. The generated contract should still expose that
    icon under the matching card instead of leaving card.icon null.
    """
    flat = list(_flatten(nodes))
    by_group = {_path_tuple(n["path"]): n for n in flat if n.get("is_group")}
    card_prefixes = {_path_tuple(s["group_path"]) for s in card_slots}
    # A group that already carries its own text (a numbered/labeled badge) is a
    # finished visual element, not an empty icon frame. Excluding the whole
    # subtree keeps a badge's inner circle from being chosen as an icon slot.
    badge_prefixes = {
        _path_tuple(n["path"]) for n in flat if _is_text_badge_group(n)
    }

    def under_text_badge(node: dict) -> bool:
        path = _path_tuple(node["path"])
        return any(path[: len(prefix)] == prefix for prefix in badge_prefixes)

    candidates = [
        n for n in flat
        if _is_external_icon_candidate(n, card_prefixes) and not under_text_badge(n)
    ]
    used: set[tuple] = set()
    by_path = {_path_tuple(n["path"]): n for n in flat}

    def under_large_icon_cluster(node: dict) -> bool:
        path = _path_tuple(node["path"])
        for depth in range(1, len(path)):
            ancestor = by_path.get(path[:depth])
            if (
                ancestor
                and ancestor.get("role") == "icon_group"
                and (ancestor.get("geometry") or {}).get("area_ratio", 0) > 0.05
            ):
                return True
        return False

    candidates = [n for n in candidates if not under_large_icon_cluster(n)]
    overlay_candidates = [
        n for n in flat
        if _is_icon_overlay_candidate(n) and not under_large_icon_cluster(n)
    ]

    for slot in card_slots:
        if slot.get("icon"):
            icon_node = by_path.get(_path_tuple(slot["icon"]))
            if icon_node and _is_icon_frame(icon_node):
                overlays = [
                    n["path"] for n in overlay_candidates
                    if _path_tuple(n["path"]) != _path_tuple(slot["icon"])
                    and _inside_frame(n, icon_node)
                ]
                if overlays:
                    slot["icon_overlays"] = _prune_nested_paths(overlays)
                    for path in slot["icon_overlays"]:
                        used.add(_path_tuple(path))
            continue
        card_node = by_group.get(_path_tuple(slot["group_path"]))
        card_geo = slot.get("card_geometry") or (card_node["geometry"] if card_node else None)
        if not card_geo:
            continue
        viable = []
        for cand in candidates:
            path = _path_tuple(cand["path"])
            if path in used:
                continue
            if not _is_near_card_icon(card_geo, cand):
                continue
            viable.append(cand)
        if not viable:
            continue
        frame_candidates = [n for n in viable if _is_icon_frame(n)]
        if frame_candidates:
            # The frame carries the actual image placement contract: geometry,
            # size, and shape. Prefer it over the smaller glyph drawn inside it.
            best = min(frame_candidates, key=lambda n: _candidate_score(card_geo, n))
            overlays = [
                n["path"] for n in overlay_candidates
                if n is not best and _inside_frame(n, best)
            ]
            slot["icon_overlays"] = _prune_nested_paths(overlays)
        else:
            best = min(viable, key=lambda n: _candidate_score(card_geo, n))
        slot["icon"] = best["path"]
        used.add(_path_tuple(best["path"]))
        for path in slot.get("icon_overlays", []):
            used.add(_path_tuple(path))


def _similar(a, b, tol=0.18) -> bool:
    if not a or not b:
        return False
    return abs(a - b) <= tol * max(a, b)


def detect_card_list(nodes, card_slots) -> list[dict]:
    """Group cards into card_lists when ≥2 cards share similar width & height."""
    # map group_path -> geometry width/height
    geo = {tuple(n["path"]): n["geometry"] for n in _flatten(nodes) if n["is_group"]}

    def slot_geo(slot: dict) -> dict | None:
        return slot.get("card_geometry") or geo.get(tuple(slot["group_path"]))

    buckets: list[list[dict]] = []
    for slot in card_slots:
        g = slot_geo(slot)
        placed = False
        for b in buckets:
            bg = slot_geo(b[0])
            if g and bg and _similar(g["width"], bg["width"]) and _similar(g["height"], bg["height"]):
                b.append(slot)
                placed = True
                break
        if not placed:
            buckets.append([slot])
    result = []
    for b in buckets:
        if len(b) >= 2:
            result.append({"kind": "card_list", "count": len(b), "items": b})
        else:
            result.append({"kind": "card", "count": 1, "items": b})
    return result


# ---------------------------------------------------------------------------
# Slide role
# ---------------------------------------------------------------------------
def classify_slide_role(nodes, index: int, total: int) -> tuple[str, float, list[str]]:
    flat = list(_flatten(nodes))
    leaves = list(_leaves(nodes))
    text_all = " ".join(l["text"]["full"] for l in leaves if l.get("text") and l["text"]["full"])
    roles = [n.get("role") for n in flat]

    n_cards = sum(1 for n in flat if n.get("role") == "card")
    n_body = sum(1 for n in leaves if n.get("role") == "body")
    n_image = sum(1 for n in leaves if n.get("role") == "image")
    has_section_no = any(r == "section_number" for r in roles)
    has_chart_table = any(r in ("chart", "table") for r in roles)
    meaningful = [l for l in leaves if l.get("role") not in DECOR_ROLES]

    is_contents = bool(CONTENTS_RE.search(text_all)) or any(
        "stock:contents" in (n.get("signals") or []) for n in flat
    )
    is_ending = bool(ENDING_RE.search(text_all))

    if is_contents and (n_cards >= 2 or index <= 3):
        return "contents", 0.8, ["contents_text"]
    if index == total - 1 and (is_ending or len(meaningful) <= 8):
        return "ending", 0.7, ["last_slide", "ending_text" if is_ending else "sparse"]
    if index == 0 and not is_contents:
        return "cover", 0.75, ["first_slide"]
    if (
        has_section_no
        and n_cards == 0
        and n_image == 0
        and not has_chart_table
        and n_body <= 2
        and len(meaningful) <= 9
    ):
        return "section_divider", 0.7, ["section_number", "sparse"]
    if n_cards >= 1 or n_body >= 1 or n_image >= 1 or has_chart_table:
        return "content", 0.6, ["cards" if n_cards else "body/media"]
    if is_ending:
        return "ending", 0.55, ["ending_text"]
    return "content", 0.35, ["fallback"]


# ---------------------------------------------------------------------------
# fill_plan synthesis
# ---------------------------------------------------------------------------
def _card_member_paths(card_slots) -> set:
    """All paths that belong to a card (so we don't double-count at slide level)."""
    members = set()
    for slot in card_slots:
        members.add(tuple(slot["group_path"]))
        for key in ("title", "body", "icon", "number"):
            if slot[key]:
                members.add(tuple(slot[key]))
    return members


def _under_any(path, prefixes) -> bool:
    p = tuple(path)
    return any(p[: len(pref)] == pref for pref in prefixes)


def _promote_contents_body_labels(body_nodes: list[dict], label_nodes: list[dict],
                                  *, slide_role: str | None) -> tuple[list[dict], list[dict]]:
    """Treat short lower-row contents labels as item body/subtitle slots.

    Some contents pages use short English subtitles (`Use Cases`) below the
    section title row. The generic text classifier sees those as labels, but
    generation needs all item subtitles in `body[]` so their font treatment is
    parallel.
    """
    if slide_role != "contents" or len(body_nodes) < 2 or not label_nodes:
        return body_nodes, label_nodes

    body_by_row: dict[int, list[dict]] = {}
    for node in body_nodes:
        geo = node.get("geometry") or {}
        row = round(geo.get("top", 0) / ROW_BUCKET_EMU)
        body_by_row.setdefault(row, []).append(node)
    target_row, target_nodes = max(body_by_row.items(), key=lambda item: len(item[1]))
    if len(target_nodes) < 2:
        return body_nodes, label_nodes

    body_font = max((_font_of(n) for n in target_nodes), default=0)
    promoted: list[dict] = []
    remaining: list[dict] = []
    for node in label_nodes:
        geo = node.get("geometry") or {}
        row = round(geo.get("top", 0) / ROW_BUCKET_EMU)
        font = _font_of(node)
        same_row = abs(row - target_row) <= 0
        body_like_size = body_font == 0 or font <= body_font * 1.15
        if same_row and body_like_size:
            promoted.append(node)
        else:
            remaining.append(node)

    if promoted:
        body_nodes = sorted(body_nodes + promoted, key=lambda n: _row_major_key(n["geometry"]))
    return body_nodes, remaining


def _is_visual_obstacle(node: dict, card_prefixes: set[tuple]) -> bool:
    """Large non-text visual regions that text should not expand into."""
    if (node.get("text") or {}).get("full"):
        return False
    if _under_any(node.get("path"), card_prefixes):
        return False
    if node.get("role") not in {"background", "decoration"}:
        return False
    if node.get("shape_type") in {"LINE", "CONNECTOR"}:
        return False
    geo = node.get("geometry") or {}
    if not geo.get("on_canvas", True):
        return False
    area = geo.get("area_ratio", 0)
    if area < 0.015:
        return False
    # Whole-slide backgrounds are layout substrate, not obstacles.
    if area >= 0.78 and geo.get("width_in", 0) >= 10.5 and geo.get("height_in", 0) >= 6.2:
        return False
    return True


def build_fill_plan(nodes, card_groups, *, slide_role: str | None = None) -> dict:
    flat = list(_flatten(nodes))
    card_slots = [item for g in card_groups for item in g["items"]]
    group_geo = {_path_tuple(n["path"]): n["geometry"] for n in flat if n.get("is_group")}
    by_path = {_path_tuple(n["path"]): n for n in flat}

    def slot_geo(slot: dict) -> dict:
        return slot.get("card_geometry") or group_geo.get(_path_tuple(slot["group_path"]), {})

    card_slots.sort(key=lambda s: _row_major_key(slot_geo(s)))
    card_prefixes = {tuple(s["group_path"]) for s in card_slots}
    card_member_paths = _card_member_paths(card_slots)

    def free(role):
        out = []
        for n in flat:
            if n.get("role") != role:
                continue
            path = _path_tuple(n["path"])
            if path in card_member_paths or _under_any(n["path"], card_prefixes):
                continue
            out.append(n)
        return out

    def slot_for(n) -> dict:
        slot = {"path": n["path"], "sample": n["text"]["sample"] if n.get("text") else None}
        if n.get("text"):
            slot["default_action"] = _default_text_action(n)
            style = _text_style(n)
            if style:
                slot["style"] = style
        return slot

    title_header_paths = {
        _path_tuple(n["path"])
        for n in free("title")
        if _is_page_header_title_candidate(n)
    }
    section_header_paths = {
        _path_tuple(n["path"])
        for n in free("section_number")
        if _is_page_header_section_number_candidate(n)
    }
    peer_title_paths: set[tuple] = set()
    if not title_header_paths:
        peer_title_paths = {
            _path_tuple(n["path"])
            for n in _parallel_peer_text_nodes(free("title"))
        }

    def best(role):
        cands = free(role)
        if not cands:
            return None
        if role == "title":
            header_cands = [n for n in cands if _path_tuple(n["path"]) in title_header_paths]
            if header_cands:
                cands = header_cands
            elif peer_title_paths:
                cands = [n for n in cands if _path_tuple(n["path"]) not in peer_title_paths]
                if not cands:
                    return None
        if role == "section_number" and section_header_paths:
            header_cands = [n for n in cands if _path_tuple(n["path"]) in section_header_paths]
            if header_cands:
                cands = header_cands
        cands.sort(key=lambda n: (n.get("role_confidence", 0), _font_of(n)), reverse=True)
        return slot_for(cands[0])

    images = [
        {"path": n["path"], "aspect": n["geometry"]["aspect"]}
        for n in free("image")
    ]
    charts = [
        {"path": n["path"], "sample": None}
        for n in free("chart")
    ]
    tables = [
        {"path": n["path"], "sample": None}
        for n in free("table")
    ]
    visual_obstacles = [
        {
            "path": n["path"],
            "kind": n.get("role"),
            "geometry": n.get("geometry"),
        }
        for n in flat
        if _is_visual_obstacle(n, card_prefixes)
    ]
    title_slot = best("title")
    subtitle_slot = best("subtitle")
    section_number_slot = best("section_number")
    logo_slot = best("logo")
    footer_slot = best("footer")
    used_single_paths = {
        _path_tuple(slot["path"])
        for slot in (title_slot, subtitle_slot, section_number_slot, logo_slot, footer_slot)
        if slot and slot.get("path")
    }

    body_nodes = list(free("body"))
    label_nodes = list(free("label"))
    for n in [*free("title"), *free("subtitle"), *free("section_number"), *free("footer")]:
        path = _path_tuple(n["path"])
        if path in used_single_paths:
            continue
        if not ((n.get("text") or {}).get("sample") or "").strip():
            continue
        label_nodes.append(n)
    body_nodes, label_nodes = _promote_contents_body_labels(
        body_nodes,
        label_nodes,
        slide_role=slide_role,
    )
    body_nodes.sort(key=lambda n: _row_major_key(n["geometry"]))
    bodies = [slot_for(n) for n in body_nodes]
    label_nodes.sort(key=lambda n: _row_major_key(n["geometry"]))
    labels = [slot_for(n) for n in label_nodes]

    plan = {
        "title": title_slot,
        "subtitle": subtitle_slot,
        "section_number": section_number_slot,
        "logo": logo_slot,
        "footer": footer_slot,
        "body": bodies,
        "labels": labels,
        "images": images,
        "charts": charts,
        "tables": tables,
        "cards": [
            {
                "group_path": s["group_path"],
                "title": s["title"], "title_sample": s["title_sample"],
                "title_style": s.get("title_style"),
                "body": s["body"], "body_sample": s.get("body_sample"),
                "body_style": s.get("body_style"),
                "title_body_mode": s.get("title_body_mode"),
                "card_geometry": slot_geo(s) or None,
                "icon": s["icon"], "icon_overlays": s.get("icon_overlays", []),
                "icon_style": _icon_style_for_slot(s, by_path, flat) if s.get("icon") else None,
                "number": s["number"],
                "residual_text": s.get("residual_text", []),
            }
            for s in card_slots
        ],
        "visual_obstacles": visual_obstacles,
    }
    return plan


def analyze_slide(nodes, index: int, total: int) -> dict:
    """Full slide analysis: groups + role + fill_plan."""
    group_cards = collect_cards(nodes)
    text_pair_cards = collect_text_pair_cards(nodes)
    cards = merge_card_slots(group_cards, text_pair_cards)
    associate_external_card_icons(nodes, cards)
    card_groups = detect_card_list(nodes, cards)
    role, role_conf, role_sig = classify_slide_role(nodes, index, total)
    fill_plan = build_fill_plan(nodes, card_groups, slide_role=role)
    return {
        "role": role,
        "role_confidence": role_conf,
        "role_signals": role_sig,
        "groups": card_groups,
        "fill_plan": fill_plan,
    }
