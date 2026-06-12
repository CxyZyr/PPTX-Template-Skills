# spec.json schema

`spec.json` is the contract between parsing and generation. The generation skill
(`ppt-template-adaptation`) reads it to find *where* each semantic slot lives on every slide, then fills
those slots from a `content_plan.json`. This file documents every field.

Produced by `scripts/parse_template.py` (`SCHEMA_VERSION = "1.0"`).

## The index-path convention (most important)

Every shape is addressed by an **index-path**: a list of integers walking the shape tree.

- `[4]` → `slide.shapes[4]` (5th top-level shape; z-order == iteration order).
- `[4, 1]` → `slide.shapes[4].shapes[1]` (2nd child of group 4).

Resolve a path to a live shape with `lib.pptx_toolkit.get_shape_by_path(slide, path)` (returns `None` if
out of range). Always resolve **all** paths you need on a slide to shape objects in one read pass *before*
mutating it — deleting a sibling keeps held refs valid, but re-resolving by index afterward would point at
the wrong shape. `lib.pptx_toolkit.resolve_paths` / `content.resolve_fill_plan` do this.

## Top level

```jsonc
{
  "schema_version": "1.0",
  "template": {
    "path": "PPT模板/大气人工智能科技感PPT模板.pptx",
    "name": "大气人工智能科技感PPT模板",
    "slide_width_emu": 12192000,            // pass to apply_page as slide_width
    "slide_height_emu": 6858000,
    "slide_count": 16,
    "theme": {
      "common_fonts":  ["微软雅黑"],          // most frequent run fonts (use as brand defaults)
      "font_availability": {
        "微软雅黑": {"available": true, "matched": "Microsoft YaHei"}
      },                                      // best-effort render-environment font check
      "common_colors": ["00FFF3", "00FFFD"]  // most frequent explicit run colors (hex, no '#')
    }
  },
  "slides": [ /* one object per parsed slide, see below */ ]
}
```

All sizes are EMU (English Metric Units): 914400 EMU = 1 inch, 12700 EMU = 1 pt.

## Per-slide object

```jsonc
{
  "index": 3,                       // 0-based; matches the original template slide order
  "role": "content",               // cover | contents | section_divider | content | ending
  "role_confidence": 0.6,           // 0..1 — review anything < 0.5
  "role_signals": ["cards"],       // why this role fired
  "render": "renders/slide-03.png", // path (relative to spec dir) or null
  "fill_plan": { /* the slot map — see below */ },
  "groups":    [ /* card / card_list detection — see below */ ],
  "shapes":    [ /* raw classified shape tree — see below */ ]
}
```

## fill_plan (what generation fills)

The distilled, ready-to-fill slot map. Generation reads **this** almost exclusively.

```jsonc
{
  "title":          {"path": [0], "sample": "点击输入标题", "style": {...}} | null,
  "subtitle":       {"path": [...], "sample": "..."} | null,
  "section_number": {"path": [...], "sample": "PART 01"} | null,
  "logo":           {"path": [...], "sample": "..."} | null,
  "footer":         {"path": [...], "sample": "...", "default_action": "preserve"} | null,
  "body":   [ {"path": [...], "sample": "...", "style": {...}}, ... ],   // free body boxes
  "labels": [ {"path": [...], "sample": "...", "style": {...}}, ... ],   // free short labels / section tags
  "images": [ {"path": [...], "aspect": 1.78}, ... ],    // picture slots (aspect = w/h)
  "charts": [ {"path": [...], "sample": null}, ... ],    // native PPT chart slots
  "tables": [ {"path": [...], "sample": null}, ... ],    // native PPT table slots
  "cards":  [ {
                "group_path":   [2],          // the card's group
                "title":        [2, 0] | null,
                "title_sample": "点击输入文字标题" | null,
                "title_style":  {"color": {...}, "font_pt": 14.0, ...} | null,
                "body":         [2, 1] | null,
                "body_sample":  "点击输入文字内容..." | null,
                "body_style":   {"color": {...}, "font_pt": 11.0, ...} | null,
                "title_body_mode": "shared_textbox" | "separate_textboxes" | null,
                "card_geometry": {"left": 0, "top": 0, "width": 0, "height": 0, ...} | null,
                "icon":         [2, 3] | null, // existing icon shape to replace, if any
                "icon_overlays": [[2, 4], ...], // old glyphs inside the icon frame to hide
                "icon_style":    {"color": {...}, "source": "overlay" | "icon" | "frame_contrast"} | null,
                "number":       [2, 2] | null, // badge number, if any
                "residual_text": [[2, 5], ...] // other text inside this card to clear after filling
              }, ... ],
  "visual_obstacles": [
    {"path": [4], "kind": "decoration", "geometry": {...}}, ...
  ]
}
```

Notes for generation:
- Any slot may be `null` — the slide simply doesn't have it. Fill only what exists; skip the rest.
- `cards` are emitted in detected order; a `content_plan` page's `cards` list maps 1:1 onto them.
- `labels` includes ordinary label shapes plus non-stock extra title/subtitle/footer
  shapes that were not selected as the slide's main title/subtitle/footer. This keeps
  template section tags and secondary footer text editable instead of leaving them as residual text.
- `title_body_mode` tells generation whether a card title/body share one text box (`shared_textbox`) or
  occupy separate template boxes (`separate_textboxes`). Generation must preserve that structure.
- `card_geometry` is the slide-level card container/union box. Generation should
  keep card titles and bodies inside this boundary instead of expanding into
  neighboring media or decorative regions.
- `style` / `title_style` / `body_style` carry parser-observed text styling
  including `font_name`, `font_pt`, `min_font_pt`, `bold`, `italic`,
  `align`, `vertical`, `color`, and optional paragraph/run structure in `paragraphs`.
  Generation must preserve these when rebuilding text frames, especially theme
  colors such as `BACKGROUND_1`; otherwise body text can become visually
  invisible on dark templates even when the text string was written.
- `default_action` tells generation what to do when the plan leaves a slot
  unfilled. Stock placeholders usually use `clear`; stable template metadata
  such as dates/page footers and decorative vertical labels use `preserve`.
- `paragraphs` exists for mixed-format text boxes. Each item records paragraph
  text, paragraph alignment, and run-level style. Generation should prefer
  in-place paragraph replacement for these slots so number/title/body styling is
  not flattened.
- A card's `icon` is the path of the existing image/icon slot. When a template has a larger frame
  shape plus smaller glyphs inside it, `icon` points to the frame and `icon_overlays` lists the old
  glyphs to hide before placing the new semantic icon. If `icon` is `null` but the content asks for
  one, generation logs a warning and skips it.
- `icon_overlays` may contain freeforms or grouped vector glyphs. If a small group sits inside an
  icon frame, the parser should emit the group path so generation can remove the whole old glyph while
  preserving the frame/background.
- `icon_style` carries the old icon foreground color when it can be inferred
  from overlay glyphs. If no overlay color is available but the icon slot is a
  filled frame, the parser may provide a contrast foreground derived from the
  frame fill.
- `residual_text` lists text-bearing shapes inside a parsed card group that were
  not selected as title, body, number, or icon overlay. Generation clears these
  after the card is filled so stale template labels/subtitles do not leak into
  the delivered deck. Do not use this as a substitute for exposing meaningful
  extra content slots when the page actually needs them.
- `visual_obstacles` lists large non-text visual regions such as partial
  backgrounds, image masks, and central decorative panels. Generation uses them
  as expansion boundaries for nearby text; whole-slide backgrounds are omitted.
- Shape `geometry` values in `shapes[]` are slide-level absolute coordinates, including shapes inside
  nested PowerPoint groups. Generation must preserve that visual geometry when replacing icons/images;
  group-local `shape.left/top` values are not a safe placement contract.
- `sample` strings are the template's original (often placeholder) text, shown so you can recognize the
  slot against the render. They are truncated to ~60 chars and newline-joined with ` | `.

## groups (card / card_list detection)

How repeated structures were grouped. Each entry is either a single `card` or a `card_list` of ≥2 cards
that share a similar width & height (so you know "these 3 cards are parallel").

```jsonc
[
  {"kind": "card_list", "count": 2, "items": [ {card slot...}, {card slot...} ]},
  {"kind": "card",      "count": 1, "items": [ {card slot...} ]}
]
```
`items` carry the same per-card slot shape as `fill_plan.cards`. `fill_plan.cards` is the flattened union
of all card items; `groups` is the grouping view.

## shapes (raw classified tree)

The full per-shape classification, for when `fill_plan` isn't enough (custom layouts, debugging a wrong
role). Recursive: a group node carries `children`.

```jsonc
{
  "path": [2, 0],
  "shape_type": "AUTO_SHAPE",        // TEXT_BOX | AUTO_SHAPE | PICTURE | GROUP | TABLE | CHART | ...
  "role": "title",                   // see classification-signals.md for the full role set
  "role_confidence": 0.9,
  "signals": ["stock:title"],
  "geometry": {
    "left": 928974, "top": 370742, "width": 2767771, "height": 584775,   // EMU
    "left_in": 1.016, "top_in": 0.405, "width_in": 3.027, "height_in": 0.64,
    "area_ratio": 0.0194,            // shape area / slide area
    "aspect": 4.733,                 // width / height
    "on_canvas": true                // false = off-slide (decoration/bleed)
  },
  "placeholder": "TITLE" | null,     // real placeholder type, if any
  "text": {                          // present only for text-bearing shapes
    "sample": "点击输入标题", "char_count": 6,
    "max_font_pt": 32.0, "min_font_pt": 12.0,
    "font_name": "微软雅黑", "any_bold": false,
    "bold": null, "italic": null,
    "align": null, "color": null,
    "paragraphs": [
      {"text": "01", "align": "CENTER", "runs": [
        {"text": "01", "font_name": "微软雅黑", "font_pt": 40.0, "color": {...}}
      ]}
    ],
    "vertical": false
  },
  "visual": {                         // best-effort fill/line style for non-text visuals
    "fill": {"type": "SOLID", "color": {"rgb": "1F70B3"}},
    "line": {"color": {"theme_color": "BACKGROUND_1"}}
  },
  "children": [ ... ]                // present only for groups
}
```

## Stability

`index` is the **original** template slide index. Generation fills by original index and only drops
unused slides at the very end (high-to-low), so spec indices stay valid throughout the fill pass.
