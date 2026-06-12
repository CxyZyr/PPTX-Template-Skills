# PPTX Template Skills

Two composable skills that turn a PowerPoint **template** into a freshly filled deck —
without an agent ever hand-reading shape indices off a slide. Parsing produces a portable,
machine-readable contract; generation fills that contract page by page while preserving the
template's original visual logic.

```
template.pptx ─[pptx-template-parsing]→ spec.json ─[ppt-template-adaptation]→ filled deck.pptx
```

## The two skills

| Skill | Role | Output |
|---|---|---|
| [`pptx-template-parsing`](skills/pptx-template-parsing) | **Parse** a template into a semantic contract | `spec.json` + `summary.txt` + optional `renders/*.png` |
| [`ppt-template-adaptation`](skills/ppt-template-adaptation) | **Generate** a new deck from `spec.json` + a `content_plan.json` | filled `deck.pptx` + renders |

The parser classifies every shape (title / body / cards / icons / logo / images / charts /
tables), infers each slide's role, detects repeated card structures, and emits a per-slide
`fill_plan` slot map. The generator reads that slot map and fills text first, then icons,
logo, and images — keeping geometry, alignment, theme colors, and paragraph/run structure
intact.

## Why parse first

The parser output is the contract. Generation never picks PowerPoint shape indices from
memory or a previous failed deck. If a slot, card count, or style is wrong, the fix belongs
in parsing — not in a per-case generation fallback. This keeps both skills portable and the
failure modes diagnosable.

## Quick start

Requirements: Python 3.12+, [`python-pptx`](https://python-pptx.readthedocs.io/). Rendering
needs `libreoffice` + `pdftoppm` on PATH (best-effort; the parser still emits `spec.json`
without them). Semantic icons use Bootstrap Icons via `rsvg-convert`. Optional web images use
a Tavily API key (`TAVILY_API_KEY`); without it, original template images are kept.

```bash
# 1. Parse a template into spec.json
python3 skills/pptx-template-parsing/scripts/parse_template.py \
  --template "path/to/template.pptx" \
  --out workspace/specs/my_template

# 2. Scaffold a content plan from the spec
python3 skills/ppt-template-adaptation/scripts/scaffold_content_plan.py \
  --spec workspace/specs/my_template/spec.json \
  --out workspace/plans/my_deck.content_plan.json \
  --output-pptx workspace/out/my_deck.pptx

# 3. Edit the plan: fill keep_slides, outline, and pages[]

# 4. Validate global text coherence
python3 skills/ppt-template-adaptation/scripts/validate_content_plan.py \
  --plan workspace/plans/my_deck.content_plan.json

# 5. Fill page by page (incremental), then a full build
python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/my_deck.content_plan.json --pages 0,1 --render

python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/my_deck.content_plan.json --render
```

## Repository layout

```
skills/
  pptx-template-parsing/      # template.pptx -> spec.json
    SKILL.md                  # when/how to use
    scripts/                  # parse_template, shape_classifier, slide_classifier
    references/               # spec schema, slide roles, classification signals
    lib/pptx_toolkit/         # shape walking, geometry, rendering helpers
  ppt-template-adaptation/    # spec.json + content_plan.json -> filled deck
    SKILL.md
    scripts/                  # scaffold / validate / apply_content_plan, render, tavily
    references/               # workflow, text-coherence, icon-fill, image-fill
    lib/pptx_toolkit/
```

Each skill is intentionally self-contained: it carries its own copy of `lib/pptx_toolkit`
so it can be used independently. See each skill's `SKILL.md` for the full workflow and
reference read order.

## License

[MIT](LICENSE) © 2026 JerryChou

## Community
> This open-source project links to and endorses the [LINUX DO Community](https://linux.do).