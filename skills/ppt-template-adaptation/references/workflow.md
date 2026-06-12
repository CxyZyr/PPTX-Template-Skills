# Workflow Summary

## Goal
Adapt an existing PPT template into a new deck without breaking the template's original layout logic.
This reference is the complete local workflow. Do not depend on old external project docs.

## Mandatory operating mode

- Start from the original template and the current `spec.json`, not from a failed generated deck.
- One global outline before page text.
- One page or a small page batch at a time; keep original slide indices stable during incremental work.
- Text first, icons second, logo third, images fourth.
- Validate structurally and render-check before moving on. If the user does visual review manually, still
  produce render artifacts and complete structural checks.

## Step order

1. Parse the template with `pptx-template-parsing`.
2. Read `summary.txt`; select pages whose roles and slot counts match the requested deck.
3. Inspect selected slides in `spec.json`, especially `fill_plan.cards`, `labels`, `body`, `images`,
   `charts`, `tables`, `style.paragraphs`, `visual_obstacles`, and `font_availability`.
4. If the parser missed a meaningful slot, wrong card count, or mixed text style, fix parsing and rerun.
   Do not compensate by hard-coding a special generation rule.
5. Scaffold a `content_plan.json`.
6. Establish `content_plan.outline[]` or derive it from the contents page.
7. Fill page text from the outline and page story. Preserve template paragraph structure when the parsed
   slot exposes mixed formatting.
8. Validate contents/section consistency.
9. Replace small semantic icons only when the parsed icon slot exists.
10. Replace logo only with a valid brand asset.
11. Replace images only with validated local files or conservative provider results.
12. Leave chart/table hooks explicit if not implemented; otherwise provide matching chart/table data.
13. Apply the plan to the current page(s) with `--pages`.
14. Render and inspect the current page, or provide the render to the user for manual inspection.
15. Iterate until the page passes, then move to the next page.
16. Run the final full build without `--pages` to drop unused slides via `keep_slides`.

## Common failure modes

- Starting from a previous failed draft instead of the original template
- Writing page text independently so contents and section pages drift apart
- Shrinking font size before trying box expansion or better wrapping
- Replacing icons without checking semantics
- Confusing small semantic icons with larger image/media slots
- Recoloring semantic icons to the accent color without checking contrast
- Placing icons from group-local coordinates instead of slide-level geometry
- Replacing a template image with a weak web image instead of keeping original
- Leaving stock authoring prompts in unparsed decorative groups after generation
- Recoloring a logo to match the template instead of preserving brand color
- Exporting once for the whole deck and skipping per-page checks

## Script map

- `scripts/scaffold_content_plan.py`: create a plan skeleton from `spec.json`
- `scripts/validate_content_plan.py`: validate global text coherence
- `scripts/apply_content_plan.py`: fill pages and render optionally
- `scripts/render_pages.py`: render PPTX pages for review
- `scripts/tavily_image_service.py`: optional Tavily image resolver that updates image paths

## Minimal command sequence

```bash
python3 skills/pptx-template-parsing/scripts/parse_template.py \
  --template "PPT模板/<template>.pptx" \
  --out workspace/specs/<template>

python3 skills/ppt-template-adaptation/scripts/scaffold_content_plan.py \
  --spec workspace/specs/<template>/spec.json \
  --out workspace/plans/<deck>.content_plan.json \
  --output-pptx workspace/out/<deck>.pptx

python3 skills/ppt-template-adaptation/scripts/validate_content_plan.py \
  --plan workspace/plans/<deck>.content_plan.json

python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/<deck>.content_plan.json \
  --pages 0 \
  --render

python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/<deck>.content_plan.json \
  --render
```

Use `--no-render` only for quick parser diagnostics. Use `--skip-global-text-check`,
`--skip-image-resolve`, and `--no-clear-unfilled` only for diagnostics, not for acceptance.

## Icon replacement rule

Small icons belong to the current page's text semantics. After text is stable,
fill `content_plan.pages[].cards[].icon` with a front-end icon-library name
such as Bootstrap Icons `database`, `cpu`, `graph-up-arrow`, or `shield-check`.
The generator must use the parsed icon slot's original size and position; do not
invent a new icon layout. If the slot is inside a PowerPoint group, resolve its
absolute slide position before placing the replacement. Larger picture/media
slots stay under `images[]`. Icon foreground color should either be inherited
from the transparent monochrome icon asset or explicitly set with
`brand.icon_color`; do not derive it from `brand.accent` automatically.

## Text coherence rule

Contents pages, section divider pages, and section-level page titles must be
derived from one global outline. Before generating, run:

```bash
python3 skills/ppt-template-adaptation/scripts/validate_content_plan.py --plan <plan.json>
```

`apply_content_plan.py` runs this check by default. Fix any error before
rendering; do not treat the current page as passed when it conflicts with the
deck outline.

## Text and style preservation rule

Use parser-observed style before brand defaults. Preserve the template's
geometry, alignment, theme colors, font names, bold/italic flags, and
paragraph/run structure unless the user explicitly asks for redesign.

For mixed-format boxes, such as number/title/body in one text frame, replace
paragraph text in place. Do not flatten the text frame into one style. Preserve
number badges and structural markers unless the user asks to remove or redesign
them.

If rendering differs because a font is missing, report the missing font from
`spec.template.theme.font_availability`. Do not replace the PPTX font name just
to match the local render environment.

## Image replacement rule

Images are replaced only when there is a validated local file. New
`content_plan.pages[].images[]` entries should use object form:

```json
{"source": "web", "provider": "tavily", "query": "...", "path": "", "fallback": "keep_original"}
```

`apply_content_plan.py` resolves Tavily image candidates automatically when
`provider: "tavily"` is present and `path` is empty. If it does not produce a
valid `path`, keep the original template image/placeholder. Do not force a
generic or low-confidence image into the deck.

## Placeholder cleanup rule

After text, icon, logo, and image filling, generation must recursively remove
stock template prompts such as `点击输入标题`, `单击此处输入标题`, and
`点击添加标题`, including prompts inside nested groups that were not exposed as
semantic fill slots. This is a final safety net, not a substitute for improving
the parser when a meaningful slot is missing.

## Chart/table page rule

If a parsed slide exposes `fill_plan.charts[]` or `fill_plan.tables[]`, treat it
as a chart/table page. Do not select that slide as an ordinary text page. Either
provide matching `content_plan.pages[].charts` / `tables` data and replace the
native object, or choose a different template page whose semantic slots can be
fully filled.

## Page acceptance

A page passes only when:

- the selected template slide's parsed slots match the intended content;
- contents and section wording align with the global outline;
- all filled text fits without overlapping neighboring slots;
- mixed-format text keeps its paragraph/run style;
- semantic icons and images use existing parsed slots;
- unresolved images keep the original template image;
- no stock authoring prompt remains in visible content;
- chart/table slots are either intentionally filled or the slide is not used.
