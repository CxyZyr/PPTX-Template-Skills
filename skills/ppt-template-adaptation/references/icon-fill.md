# Semantic Icon Fill

Use this step after the page text is stable and before logo/image replacement.

## Boundary

- `cards[].icon` is for small semantic icons next to card titles or bullets.
- `images[]` is for larger photos, illustrations, screenshots, or media slots.
- Do not use an icon slot for a large picture area, and do not use an image slot for a small semantic glyph.
- Number badges such as `01/02/03/04` are structural markers, not semantic icons. Preserve them unless the user explicitly asks to redesign the numbering system.

## Inputs from parsing

The parser provides icon slots in `spec.json`:

- `fill_plan.cards[].icon`: the original icon frame or glyph position.
- `fill_plan.cards[].icon_overlays`: old glyphs inside a frame that should be removed when replacing.
- `fill_plan.cards[].icon_style`: the original glyph foreground color when the parser can infer it.
- Geometry comes from the template and must be preserved.
- Icon paths may point inside nested PowerPoint groups. Generation must place
  replacement pictures using slide-level absolute geometry, not group-local
  `shape.left/top` values.
- If the parsed icon slot is a visual frame/background such as a circle or
  rounded rectangle, preserve that frame and replace only the inner glyph.
  Deleting the frame changes the template contrast and is a layout break.
- Old glyphs may be freeforms or grouped vector shapes. Treat both as overlays
  when they sit inside the icon frame.

If a page has a larger picture filled into an auto shape, it appears under `fill_plan.images[]`.

## Selection rule

Choose icons from a stable front-end icon library. This repo uses Bootstrap Icons through
`lib/pptx_toolkit/assets.py`.

Pick by the final card text, not by the template sample:

- data, asset, database -> `database`
- AI, compute, platform, technology -> `cpu`
- growth, revenue, trend -> `graph-up-arrow`
- implementation, protection, guarantee -> `shield-check`
- cycle, iteration, review -> `repeat`
- positioning, direction, strategy -> `compass`

When uncertain, prefer a simple, generic icon that matches the noun or action in the title.

## Execution

Put the icon name in the content plan:

```json
{
  "title": "业务布局与增长",
  "body": "聚焦核心场景，形成可复制的业务增长路径",
  "icon": "graph-up-arrow"
}
```

`skills/ppt-template-adaptation/scripts/apply_content_plan.py` calls `apply_page()`.
`lib/pptx_toolkit/content.py` then:

1. Resolves the parsed icon slot.
2. Calls `get_icon()` to fetch/cache the Bootstrap SVG and render it to PNG.
3. Recolors the icon using `brand.icon_color` when explicitly set; otherwise
   inherits `fill_plan.cards[].icon_style.color` when available. If neither is
   present, it keeps the icon's monochrome transparent PNG.
4. Removes `icon_overlays` when present.
5. Places the new icon inside the original icon frame/position.

Do not infer `icon_color` from `brand.accent`. Accent often describes title or
frame color. The parser should infer icon foreground from the old glyph overlay
or, when only a filled frame is available, a contrast color based on that frame.

## Acceptance

- Icon meaning matches the filled text.
- Icon uses the parsed slot's original geometry.
- Icon foreground keeps enough contrast against its original frame/background.
- Original icon frames/backgrounds remain unless the user explicitly asks for a redesign.
- Old template glyphs do not remain on top of the new icon.
- Number badges and section markers remain intact.
- Larger image slots remain separate and are not treated as icons.
