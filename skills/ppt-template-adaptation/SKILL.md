---
name: ppt-template-adaptation
description: "Use to generate or debug a new deck from an existing PPTX template and a parsed spec.json. Enforces the second half of the parse→generate pipeline: select pages, build one global outline, author content_plan.json, fill text while preserving parsed template style, replace icons/logo/images conservatively, validate page by page, and keep chart/table hooks explicit."
---

# PPT Template Adaptation

Use this skill when the task is to rewrite an existing PPT template into a new deck while preserving the
template's visual logic. It consumes `spec.json` from `pptx-template-parsing`; do not start from a failed
generated deck or manually inferred shape indices.

This skill is intentionally self-contained. Do not read old external handbooks from previous projects.

## What this skill enforces

- Parse first, generate second. If a slot, card count, or style is missing, fix parsing before adding
  generation fallback logic.
- Work page by page. Incremental runs keep original slide indices stable until the final full build.
- Keep one global outline as the source of truth for contents pages, section dividers, and section titles.
- Fill text before icons, logo, and images.
- Preserve parsed template style: geometry, alignment, theme colors, paragraph/run structure, and number
  badges unless the user asks for redesign.
- Replace icons and images only when semantics require it and a valid slot exists.
- Validate each page before moving on.

## Inputs

- Original template `.pptx`.
- Parsed `spec.json` from `pptx-template-parsing`.
- A `content_plan.json` with `spec`, `output`, `asset_dir`, optional `brand`, `keep_slides`, optional
  `outline`, and `pages[]`.

## Required workflow

1. If no current `spec.json` exists, run `pptx-template-parsing` first.
2. Read the parser `summary.txt`; choose template slides whose roles and slot counts match the needed
   story. Avoid chart/table slides unless chart/table data will be supplied.
3. Scaffold a plan:

```bash
python3 skills/ppt-template-adaptation/scripts/scaffold_content_plan.py \
  --spec workspace/specs/<template>/spec.json \
  --out workspace/plans/<deck>.content_plan.json \
  --output-pptx workspace/out/<deck>.pptx
```

4. Edit `keep_slides`, `outline`, and `pages[]`. Map each page to existing slots only: titles to
   `title`, free labels to `labels[]`, body boxes to `body[]`, repeated structures to `cards[]`,
   semantic glyphs to `cards[].icon`, and photos/screenshots to `images[]`.
5. Validate global text coherence:

```bash
python3 skills/ppt-template-adaptation/scripts/validate_content_plan.py \
  --plan workspace/plans/<deck>.content_plan.json
```

6. Fill one or a few original slide indices at a time:

```bash
python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/<deck>.content_plan.json \
  --pages 0,1 \
  --render
```

7. Review the rendered page when visual review is available. If the user will do visual review, provide
   the output files and continue with structural checks; do not use image tools when the user forbids it.
8. Repeat the page loop until selected pages pass. Then run a full build without `--pages` so unused
   slides are dropped according to `keep_slides`.

## Debugging order

1. Inspect parser output first: `summary.txt`, `spec.json` `fill_plan`, card count, `style.paragraphs`,
   font availability, and visual obstacles.
2. If parsing is wrong, fix `pptx-template-parsing`, rerun parsing, and rebuild the plan or affected
   pages.
3. If parsing is right but generation breaks style or layout, fix `ppt-template-adaptation` application
   logic.
4. Use `--no-clear-unfilled`, `--skip-global-text-check`, or `--skip-image-resolve` only for deliberate
   diagnostics. Do not use them to declare a page complete.

## Reference read order

Read only the references needed for the current task:

- `references/workflow.md` — complete page-by-page workflow and acceptance criteria.
- `references/text-coherence.md` — contents/section outline consistency and text fitting.
- `references/icon-fill.md` — semantic icon replacement rules.
- `references/image-fill.md` — conservative image replacement and Tavily usage.

## Example scripts

- `scripts/scaffold_content_plan.py`
- `scripts/validate_content_plan.py`
- `scripts/apply_content_plan.py`
- `scripts/render_pages.py`
- `scripts/tavily_image_service.py`

## Usage rule

Do not jump to the next page until the current page has passed validation. For visual validation, either
review the render when allowed or hand the render artifact to the user for manual review.
