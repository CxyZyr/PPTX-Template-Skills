# Classification signals

How `scripts/shape_classifier.py` assigns each shape a `role`. The guiding principle: **no single signal
is trusted alone.** Every shape records the `signals` that fired and a `role_confidence`, so a reviewer
can check the decision against the rendered slide.

## Role set

`title`, `subtitle`, `body`, `label`, `number`, `section_number`, `footer`, `logo`, `icon`, `image`,
`background`, `chart`, `table`, `decoration` — plus group roles `card`, `icon_group`, `badge`.

## Priority order (first match wins)

A non-group shape is classified by `classify_node` in this order:

1. **Hard shape type** — authoritative.
   - `TABLE` → `table` (0.95); `CHART` → `chart` (0.95); `LINE`/`CONNECTOR` → `decoration` (0.5).
2. **Real placeholder type** — rare in stock templates but authoritative when present (0.95).
   `TITLE`/`CENTER_TITLE`→title, `SUBTITLE`→subtitle, `BODY`/`OBJECT`→body, `PICTURE`→image,
   `TABLE`→table, `CHART`→chart.
3. **Text-bearing shape** (TEXT_BOX, or AUTO_SHAPE/FREEFORM carrying real text) → see *Text rules*.
4. **Picture** → see *Picture rules*.
5. **Blank shape** (AUTO_SHAPE/FREEFORM, no meaningful text) → see *Blank-shape rules*.

## Text rules (`_classify_text_shape`)

Strongest text signal is **stock placeholder text** — these patterns map almost directly to a role (0.9;
a `目录`/`CONTENTS` hint maps to `title` at 0.85):

| Pattern (regex, abbreviated) | Role |
|---|---|
| `点击/单击/请/在此 + 输入/添加 + 标题`, `标题文字`, `主标题` | `title` |
| `副标题`, `sub title` | `subtitle` |
| `点击…输入…文字/文本/内容/正文`, `此处输入`, `添加文本`, `详细内容` | `body` |
| `your logo`, `您的/公司 logo` | `logo` |
| `目录`, `CONTENTS`, `AGENDA` | (contents hint → `title`) |
| `PART.N`, `第N部分`, `第N章`, `章节` | `section_number` |

Then **pure-number** content: a year (`19xx`/`20xx`) → `label`; a single `0?\d` at ≥28pt → `section_number`;
any number at ≥24pt → `number`; otherwise → `label`.

Then **footer**: bottom band (`cy_ratio ≥ 0.88`), ≤20 chars, ≤16pt → `footer`.

Otherwise use **font size relative to the slide's max font** (`ratio = size / slide_max_font`) plus
position. `upper` means `cy_ratio ≤ 0.30`; `short` means ≤12 chars; `multiline` means ≥2 paragraphs or
>40 chars:

| Condition | Role (confidence) |
|---|---|
| `ratio ≥ 0.92` & `size ≥ 28` & upper | `title` (0.75) |
| `ratio ≥ 0.92` & `size ≥ 28` | `title` (0.6) |
| `0.55 ≤ ratio < 0.92` & short & upper | `subtitle` (0.55) |
| multiline | `body` (0.6) |
| short | `label` (0.5) |
| else | `body` (0.45) |

When font size is unknown (inherited, not set on the run), it falls back to geometry + length
(`upper + short + wide` → title; multiline → body; short → label; else body), all at lower confidence.

## Picture rules (`_classify_picture`)

Thresholds: `BG_AREA = 0.62` area ratio, `SMALL_DIM_IN = 1.3` inches, `CORNER_X = 0.16` of the width
from a vertical edge.

| Condition | Role (confidence) |
|---|---|
| area ≥ 0.62 **and** (near origin, or off-canvas, or area ≥ 0.85) | `background` (0.7) |
| max dimension ≤ 1.3in **and** in a corner | `logo` (0.55) |
| max dimension ≤ 1.3in | `icon` (0.45) |
| otherwise | `image` (0.7) |

## Blank-shape rules (`_classify_blank_shape`)

For AUTO_SHAPE/FREEFORM with no meaningful text. Thresholds: `TINY_AREA = 0.0025`,
icon aspect band `0.45–2.2` (near-square).

| Condition | Role (confidence) |
|---|---|
| off-canvas | `decoration` (0.6) |
| area ≥ 0.62 | `background` (0.55) |
| area ≤ 0.0025 (a dot) | `decoration` (0.6) |
| small (≤1.3in) **and** near-square | `icon` (0.4) |
| otherwise | `decoration` (0.35) |

## Group roles (`_classify_group`)

A group is classified **after** its descendants, from their leaf roles:

- Holds substantive text (`title`/`subtitle`/`body`) → **`card`** (0.6). Inside a card, small near-square
  `decoration` shapes are promoted to `icon` (`_promote_icons`, 0.55) — that's how a card's glyph icon is
  recognized even when it carried no text signal of its own.
- Holds only icons/images → **`icon_group`** (0.4).
- Holds only minor text (`label`/`number`/`section_number`, e.g. an `01` badge circle) → **`badge`** (0.4).
  Keeping badge groups out of `card` is deliberate: it stops a nested badge from masking the real outer
  card and leaking a card title up into the slide-title slot.
- Otherwise → **`decoration`** (0.3).

## Reviewing

Confidence is a review aid, not a guarantee. When a slide's render disagrees with the parse:
- check the shape's `signals` to see which rule fired,
- the most common misfires are: an inherited-font title landing as `body`, a large foreground photo read
  as `background`, and a decorative ring read as `icon`. Fix by editing the `content_plan`'s slot mapping,
  not the classifier, for a one-off; adjust thresholds here only if a misfire is systematic.
