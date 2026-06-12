---
name: pptx-template-parsing
description: "Use before adapting or debugging any PPTX template-driven deck. Parses a template into a portable spec.json contract: classified shape tree, slide roles, repeated card structures, per-slide fill_plan, text style/paragraph metadata, font availability, and optional renders. This is the required first half of the parse→generate pipeline consumed by ppt-template-adaptation."
---

# PPTX Template Parsing

Turn a `.pptx` template into a semantic `spec.json` that downstream generation can fill
mechanically — without the agent ever hand-reading shape indices off the slide.

This is the **first half** of the pipeline:

```
template.pptx ─[THIS SKILL]→ spec.json (+ summary.txt + renders/*.png)
                                   │
                                   ▼
                          ppt-template-adaptation  (authors content_plan.json, fills the deck)
```

## When to use

- A user provides a PPT template that will be reused (not redesigned).
- Before any generation: generation needs the `spec.json` slot map this skill produces.
- When you need to know, per slide: what role it plays, where the title/body/cards/icons/logo/images
  live, and which slides repeat a card structure.
- When a generated deck has layout, style, overlap, or missing-slot problems. Check whether parsing
  captured the right slots and text styles before adding generation fallbacks.

## Operating rule

The parser output is the contract. Do not hand-pick PowerPoint shape indices from memory, screenshots,
or a previous failed deck. If `fill_plan` is wrong, fix parsing/classification first, rerun this skill,
then regenerate with `ppt-template-adaptation`.

## What it produces

| Output | Purpose |
|---|---|
| `spec.json` | The machine-readable contract consumed by generation. See `references/spec-schema.md`. |
| `summary.txt` | One line per slide: role + confidence + detected slots. Read this first. |
| `renders/*.png` | One PNG per slide when rendering is enabled. Use for human/visual review when allowed. |

## How to run

```bash
python3 skills/pptx-template-parsing/scripts/parse_template.py \
  --template "PPT模板/大气人工智能科技感PPT模板.pptx" \
  --out workspace/specs/ai_tech
```

Flags:
- `--no-render` — skip LibreOffice rendering (faster; do this for a quick structural pass).
- `--limit N` — only parse the first N slides (useful while iterating).
- `--dpi 110` — render resolution (default 110).

Rendering needs `libreoffice` + `pdftoppm` on PATH. If rendering fails the parser still emits
`spec.json` (renders are best-effort).

## Required workflow

1. Run `parse_template.py` against the original template, not a previously generated deck.
2. Read `summary.txt` and flag any low-confidence slide (`role_confidence < 0.5`) or unexpected slot
   count, especially `contents` and repeated-card slides.
3. Read the relevant `fill_plan` entries in `spec.json`: `title`, `labels`, `body`, `cards`,
   `images`, `charts`, `tables`, `logo`, and `visual_obstacles`.
4. For text-sensitive templates, inspect each slot's `style.paragraphs` / `title_style.paragraphs`.
   Mixed-format boxes must preserve paragraph/run style instead of being flattened during generation.
5. Check `template.theme.font_availability`. Missing fonts explain render differences, but the PPTX
   should still preserve the original font names in `style`.
6. Only after the slot map looks coherent, hand `spec.json` to `ppt-template-adaptation` and create a
   `content_plan.json`.

## Debugging checklist

- Contents page: card count must match visible directory entries. If four entries parse as one large
  card, fix card-list/group extraction here.
- Shared text boxes: if a single text box contains number/title/body with different styles, the parsed
  style must include `paragraphs` with run-level sizes/colors.
- Alignment: preserve parser values such as `CENTER`, `LEFT`, and `DISTRIBUTE`. Do not reinterpret
  WPS distributed alignment as ordinary center.
- Icon/image boundary: small semantic glyphs belong in `cards[].icon`; photos/screenshots belong in
  `images[]`.
- Chart/table pages: if native chart/table slots are present, expose them as `charts[]` / `tables[]`;
  do not treat them as ordinary text pages.

## How the classifier works (no single signal is trusted)

Each shape gets a `role`, a `role_confidence`, and the list of `signals` that fired, so its decision is
reviewable against the render. Roles are assigned in priority order — hard shape types (table/chart) →
real placeholders → stock placeholder text (`点击输入标题`) → font-size/geometry heuristics. Groups become
`card` only when they carry substantive text. Full rules: `references/classification-signals.md`.

Slide roles (cover / contents / section_divider / content / ending) and the per-slide `fill_plan` are
derived on top of the shape roles. Full rules: `references/slide-roles.md`.

## Reference read order

Read only what the task needs:

- `references/spec-schema.md` — required when consuming or debugging `spec.json`.
- `references/classification-signals.md` — read when a shape role is wrong.
- `references/slide-roles.md` — read when a slide role, card count, or `fill_plan` is wrong.

## Scripts

- `scripts/parse_template.py` — CLI entry: template → `spec.json` + `summary.txt` + renders.
- `scripts/shape_classifier.py` — per-shape role classification (builds the node tree).
- `scripts/slide_classifier.py` — slide role + card-list detection + `fill_plan` synthesis.

These scripts depend on `lib/pptx_toolkit` (shape walking, geometry, rendering).

## Accuracy expectations

This is heuristic classification, not ground truth. `role_confidence` and `signals` exist precisely so a
reviewer can catch mistakes. Always sanity-check `summary.txt` and the relevant `spec.json` entries
before authoring a `content_plan.json` against the spec. When visual review is available, compare renders
against the template; otherwise leave the render artifacts for human review and rely on structural checks.
