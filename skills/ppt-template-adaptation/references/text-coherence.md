# Global Text Coherence

Deck text must be authored from a single outline. Do not treat each slide as an
isolated writing task.

## Source of truth

Use this priority:

1. `content_plan.outline[]`, if present.
2. The first `role: "contents"` page's `cards[]`.
3. If neither exists, stop and create an outline before writing section pages.

Each outline entry should represent one section:

```json
{
  "title": "战略共识与定位",
  "subtitle": "明确2026战略主题与数智能源价值主线",
  "number": "01"
}
```

## Required consistency

- Contents-page entries and section-divider pages must match by order.
- If a contents page exposes section names through `labels[]` rather than
  `cards[]`, every outline title must appear among the non-number labels in
  outline order (an ordered subsequence). Decorative labels such as a
  `CONTENTS`/`目录` banner or English captions between section names are skipped,
  not treated as section names. A missing or out-of-order section title fails.
- A section divider's first title line must match the corresponding outline
  title.
- If the template exposes the visible section title through `labels[]` instead
  of `title`, use the first non-number label as the section title. Use body text
  as the subtitle/description when no title subtitle exists.
- If both sides have a subtitle, the subtitle must match too.
- `PART 01/02/...` must match the section order.
- Body/content pages under a section should use titles and wording that support
  that section, not introduce an unrelated section topic.

## Validation

Run this before rendering:

```bash
python3 skills/ppt-template-adaptation/scripts/validate_content_plan.py \
  --plan workspace/plans/your.content_plan.json
```

`apply_content_plan.py` runs the same validation by default. Use
`--skip-global-text-check` only for a deliberate diagnostic run.

## Text fit

- Do not shorten meaningful text just to avoid wrapping. Preserve the authored
  content, then let the applier expand or refit the template text box.
- Use single-line fitting only for isolated titles and labels that have room to
  expand without touching neighboring slots.
- For repeated title-only card lists such as contents pages, keep the original
  card box as the layout boundary and allow natural line breaks at semantic
  separators such as `|`.
- For contents pages with parallel section title/description rows, fit each
  row as a peer group and use the smallest fitted font size across that row.
  Do not let the final item grow larger just because it has no right-side
  neighbor.
- For same-size parallel text peers, normalize by geometry rather than by slot
  name alone. If a parser version splits a visual peer group across `title` and
  `labels[]`, generation still uses the smallest fitted font size across the
  peer group.
- Before expanding a text box, respect slide-level neighboring text slots. A
  right-side neighbor limits horizontal growth; a lower overlapping neighbor
  limits title height. Prefer this order: bounded expansion, natural wrapping,
  then font reduction. Do not solve long text by shortening authored content.
- After rendering, check extracted PDF text for weak CJK wraps such as a
  one-character line split from a phrase. These indicate a layout fit problem,
  not a content-plan wording problem.

## Acceptance

- The contents page can be used as a reliable map of the deck.
- Every section divider corresponds to exactly one contents/outline entry.
- Repeated section wording is intentionally reused, not manually rephrased into
  drift.
- Page-level text fits the layout after global consistency has passed.
