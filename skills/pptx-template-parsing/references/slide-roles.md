# Slide roles and fill_plan

How `scripts/slide_classifier.py` turns the classified shape tree (from `shape_classifier.py`) into a
slide `role`, a card/card-list grouping, and the ready-to-fill `fill_plan`.

## Slide roles

Five business roles: `cover`, `contents`, `section_divider`, `content`, `ending`. Assigned by
`classify_slide_role(nodes, index, total)` using shape-role counts, slide text, and slide position.
First match wins:

| Condition | Role (confidence) |
|---|---|
| contents text (`目录`/`CONTENTS`/`AGENDA` or a `stock:contents` signal) **and** (≥2 cards or `index ≤ 3`) | `contents` (0.8) |
| last slide **and** (ending text `谢谢`/`THANK YOU`, or ≤8 meaningful shapes) | `ending` (0.7) |
| `index == 0` and not contents | `cover` (0.75) |
| has a `section_number`, no cards/images/chart/table, ≤2 body slots, and ≤9 meaningful shapes | `section_divider` (0.7) |
| ≥1 card, or ≥1 body, or ≥1 image, or a chart/table | `content` (0.6) |
| ending text present | `ending` (0.55) |
| fallback | `content` (0.35) |

"Meaningful" = leaves whose role is not `decoration`/`background`. Review any slide at `< 0.5`, and any
`content` slide whose detected card count disagrees with the render.

## Card and card-list detection

1. `collect_cards` walks the tree for **leaf** `card` groups — a `card` group that does *not* itself
   contain child card groups (so an outer container of cards isn't mistaken for a single card). For each,
   `_extract_slots` picks its slots from the card's leaves:
   - **title**: the highest-font `title`/`subtitle` leaf; if none, the highest-font `label`.
   - **body**: the first `body` leaf. **icon**: the first `icon` leaf. **number**: the first
     `number`/`section_number` leaf.
2. `detect_card_list` buckets cards whose group width **and** height are similar (within 18%). A bucket of
   ≥2 becomes a `card_list` (these are parallel cards); a lone card stays a `card`. This is the `groups`
   array in the spec.

## fill_plan synthesis (`build_fill_plan`)

The distilled slot map generation actually fills. Key rule: **shapes that belong to a card are excluded
from the slide-level free slots**, so a card's title can never also be picked as the slide title.

- `title` / `subtitle` / `section_number` / `logo` / `footer`: the best **free** (non-card) shape of that
  role — highest `role_confidence`, then largest font (`best(role)`).
- `body`: **all** free `body` shapes (a list — slides can have several parallel body boxes).
- `labels`: all free `label` shapes, plus non-stock extra `title`/`subtitle`
  shapes that were not selected as the slide's main title/subtitle. These are
  common template section tags and must remain fillable or clearable.
- If multiple free `title` shapes form a same-size peer row/column and no
  page-header title exists, none is promoted to the single slide `title`; the
  whole peer group stays in `labels[]` so generation can fill and normalize the
  repeated slots together.
- On `contents` slides, short labels that sit on the same lower row as other
  body/subtitle items are promoted into `body[]`. This keeps directory item
  subtitles parallel even when one item is short enough to be initially
  classified as a label.
- `images`: all free `image` shapes, each with its `aspect`.
- `cards`: the flattened slots of every detected card (the same slots grouped under `groups`).

See `spec-schema.md` for the exact JSON shape of each of these.

## What this means for generation

- The presence/absence of each slot tells you what a slide can hold. A `section_divider` typically exposes
  `section_number` + `title` and little else; a `content` slide exposes `title` + `cards` and/or
  `body`/`images`.
- Section divider templates often include an English subtitle plus a one-line explanatory body. Those
  sparse pages should still classify as `section_divider`; the role should not depend on only one body
  slot being present.
- A `content_plan` page's `cards` list maps 1:1 onto `fill_plan.cards` in order. If you supply more cards
  than the template has slots, the extras are dropped (and the applier logs it) — pick a template slide
  whose card count matches your content, or split across slides.
- `role` is advisory. You choose which template slide carries which content in the `content_plan`; the
  role just tells you what the slide was designed for.
