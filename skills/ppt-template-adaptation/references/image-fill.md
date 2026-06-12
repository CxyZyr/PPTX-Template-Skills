# Image Fill

Image replacement is conservative. A weak image is worse than the original
template visual, so replacement requires a high-confidence local file.

## Slot categories

Classify each `images[]` slot before replacing:

- `semantic_photo`: real-world scene/photo related to the page topic.
- `product_screenshot`: product, platform, dashboard, website, or app UI.
- `brand_asset`: logo or brand-owned asset; do not use generic web search.
- `decorative_background`: template atmosphere/background; usually preserve.
- `placeholder_only`: explicit placeholder; replace only when a strong asset is available.

## Content Plan Contract

Use structured image items. Empty `path` plus `fallback: "keep_original"` means
generation preserves the template image/slot.

```json
{
  "source": "web",
  "provider": "tavily",
  "query": "新能源 数据资产平台 可视化 dashboard",
  "required_terms": ["新能源", "数据资产"],
  "reject_terms": ["物流", "跟踪", "iot"],
  "allowed_domains": ["xxx.com"],
  "require_page_url": true,
  "path": "",
  "fallback": "keep_original",
  "category": "product_screenshot"
}
```

Existing string paths are still accepted for backward compatibility, but new
plans should use the object form.

## Tavily Provider

Tavily is used to find web pages and image candidates. It is not treated as a
dedicated image-search oracle; the service still filters by URL, dimensions,
image validity, and placeholder/icon signals.

`query` is only a search seed. The provider still adds current page and section
context. Use `required_terms` for page-specific must-match concepts and
`reject_terms` for nearby but wrong domains.

The provider uses a strict-first acceptance strategy:

- Default strict pass: candidate metadata/source text must satisfy `required_terms`,
  `reject_terms`, dimensions, aspect ratio, and non-logo/non-icon checks.
- Relaxed fallback is opt-in only: set `semantic_fallback: "relaxed"` when the
  user explicitly accepts looser matching. If used, the item status must be
  `resolved_relaxed` and include a semantic warning.

By default `require_page_url` is true. Anonymous image-CDN direct links are not
accepted unless the plan explicitly sets `require_page_url: false` for a
diagnostic or low-risk draft.

For brand-owned product assets, prefer `allowed_domains`. If Tavily returns an
official image CDN direct link without `page_url`, the provider may accept it
only when the image URL domain itself matches `allowed_domains`. Direct links
from unapproved domains still fail the traceability check.

Normal generation runs the provider automatically:

```bash
TAVILY_API_KEY=... python3 skills/ppt-template-adaptation/scripts/apply_content_plan.py \
  --plan workspace/plans/your.content_plan.json --render
```

Use `--skip-image-resolve` only for diagnostics or when the user explicitly
wants no web lookup.

You can also run the provider script directly to inspect or prewrite resolved
paths:

```bash
TAVILY_API_KEY=... python3 skills/ppt-template-adaptation/scripts/tavily_image_service.py \
  --plan workspace/plans/your.content_plan.json \
  --out workspace/plans/your.with_images.content_plan.json
```

The API key must come from the environment or a local project `.env`. Do not
write it into `SKILL.md`, `content_plan.json`, shell history examples, or
generated reports.

The service writes accepted images under `asset_dir/tavily_images/` and updates
only the image item that has a validated local `path`. If no candidate passes,
the item stays `path: ""` and `fallback: "keep_original"`.

## Generation Rule

`apply_content_plan.py` replaces an image only when the matching image item has
a valid local `path`. It preserves slot order:

- slot 1 can keep original while slot 2 is replaced;
- empty strings do not shift later image slots;
- missing files are warnings and keep the original image.

## Acceptance

- The image supports the global outline and the current page topic.
- The source is recorded through `url`, and `page_url` is recorded when Tavily
  provides one.
- The image has enough resolution for the target slot.
- Cropping preserves the intended subject.
- If no suitable image is found, the original template image remains.
