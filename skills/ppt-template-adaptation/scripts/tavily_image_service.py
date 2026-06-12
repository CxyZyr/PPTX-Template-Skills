"""Tavily-backed image candidate service for content_plan images.

This is an optional provider layer. It never hardcodes credentials and never
forces replacement: if no high-confidence image is found, the plan keeps
`fallback: keep_original` so apply_content_plan.py preserves the template image.

Usage:
  TAVILY_API_KEY=... python3 skills/ppt-template-adaptation/scripts/tavily_image_service.py \
      --plan workspace/plans/deck.content_plan.json \
      --out workspace/plans/deck.with_images.content_plan.json
"""
from __future__ import annotations

import argparse
import io
import json
import os
import pathlib
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Any

from PIL import Image

TAVILY_SEARCH_URL = "https://api.tavily.com/search"
DEFAULT_ASSET_SUBDIR = "tavily_images"
BAD_IMAGE_TOKENS = (
    "logo", "icon", "favicon", "avatar", "sprite", "wechat", "qrcode",
    "qr-code", "blank", "placeholder", "loader", "spinner",
)
GENERIC_QUERY_TERMS = {
    "2026", "战略", "目标", "平台", "可视化", "dashboard", "服务",
    "建设", "产品", "业务", "发展", "管理", "系统", "方案",
}
OWNED_SOURCE_ONLY_CATEGORIES = {
    "product_screenshot", "platform_screenshot", "ui_screenshot",
    "dashboard_screenshot", "app_screenshot",
}


@dataclass
class ImageCandidate:
    url: str
    description: str = ""
    page_url: str = ""
    title: str = ""


def load_dotenv(path: str | pathlib.Path = ".env") -> None:
    """Load simple KEY=VALUE lines without overwriting existing environment."""
    dotenv = pathlib.Path(path)
    if not dotenv.exists():
        return
    for raw_line in dotenv.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def _safe_name(value: str, *, fallback: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-")
    return value[:80] or fallback


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return " ".join(str(v).strip() for v in value if str(v).strip())
    return str(value).strip()


def _slot_aspects(plan: dict) -> dict[int, list[float]]:
    spec_path = plan.get("spec")
    if not spec_path:
        return {}
    try:
        spec = json.loads(pathlib.Path(spec_path).read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[int, list[float]] = {}
    for slide in spec.get("slides", []):
        images = slide.get("fill_plan", {}).get("images", [])
        out[slide.get("index")] = [float(img.get("aspect") or 1.0) for img in images]
    return out


def _target_ratio(aspects: dict[int, list[float]], slide_index: int | None,
                  slot_index: int) -> float:
    values = aspects.get(slide_index) or []
    if 0 <= slot_index < len(values):
        return values[slot_index] or 1.0
    return 1.0


def _page_text(page: dict) -> str:
    parts = [
        _as_text(page.get("title")),
        _as_text(page.get("subtitle")),
        _as_text(page.get("body")),
    ]
    for card in page.get("cards") or []:
        parts.append(_as_text(card.get("title")))
        parts.append(_as_text(card.get("body")))
    return " ".join(p for p in parts if p)


def _outline_text(plan: dict) -> str:
    outline = plan.get("outline") or []
    parts = []
    for item in outline:
        if isinstance(item, dict):
            parts.append(_as_text(item.get("title")))
            parts.append(_as_text(item.get("subtitle") or item.get("body")))
        else:
            parts.append(_as_text(item))
    if parts:
        return " ".join(p for p in parts if p)
    for page in plan.get("pages", []):
        if page.get("role") == "contents":
            return _page_text(page)
    return ""


def _list_text(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    return [str(value).strip()]


def _domain_of(url: str) -> str:
    try:
        return urllib.parse.urlparse(url).netloc.lower().split("@")[-1].split(":")[0]
    except Exception:
        return ""


def _domain_allowed(domain: str, allowed_domains: list[str]) -> bool:
    if not domain:
        return False
    allowed = [d.lower().lstrip(".") for d in allowed_domains if d]
    return any(domain == d or domain.endswith("." + d) for d in allowed)


def _allowed_domains(item: dict) -> list[str]:
    return _list_text(item.get("allowed_domains") or item.get("allowed_page_domains"))


def _requires_owned_source(item: dict) -> bool:
    category = _as_text(item.get("category")).lower()
    if category in OWNED_SOURCE_ONLY_CATEGORIES:
        return True
    return bool(item.get("require_owned_source"))


def _section_context(plan: dict, page: dict) -> str:
    slide_index = page.get("slide")
    if slide_index is None:
        return ""
    section = ""
    for candidate in plan.get("pages", []):
        if candidate.get("slide") is None or candidate.get("slide") > slide_index:
            continue
        if candidate.get("role") == "section_divider":
            section = _page_text(candidate)
    return section


def build_query(plan: dict, page: dict, image_item: dict, *, query_prefix: str = "") -> str:
    explicit = _as_text(image_item.get("query"))
    topic = _as_text(plan.get("topic") or plan.get("title"))
    page_text = _page_text(page)
    section_text = _section_context(plan, page)
    fallback_outline = _outline_text(plan) if not page_text and not section_text else ""
    text = " ".join(p for p in [query_prefix, explicit, topic, section_text, page_text, fallback_outline] if p)
    # Keep queries focused enough for Tavily and avoid dumping full slide text.
    words = text.split()
    return " ".join(words[:28])


def semantic_terms(plan: dict, page: dict, image_item: dict) -> tuple[list[str], list[str]]:
    required = _list_text(image_item.get("required_terms"))
    reject = _list_text(image_item.get("reject_terms"))
    if required:
        return required, reject

    text = " ".join([
        _as_text(image_item.get("query")),
        _section_context(plan, page),
        _page_text(page),
    ])
    parts = re.split(r"[\s,，;；/|、：:()（）\[\]{}<>《》\"']+", text)
    for part in parts:
        part = part.strip()
        if not part or part.lower() in GENERIC_QUERY_TERMS:
            continue
        if re.fullmatch(r"\d+", part):
            continue
        if len(part) >= 2:
            required.append(part)
    return required[:8], reject


def tavily_search_images(api_key: str, query: str, *, max_results: int,
                         search_depth: str) -> list[ImageCandidate]:
    payload = {
        "api_key": api_key,
        "query": query,
        "search_depth": search_depth,
        "max_results": max_results,
        "include_images": True,
        "include_image_descriptions": True,
    }
    req = urllib.request.Request(
        TAVILY_SEARCH_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return extract_candidates(data)


def extract_candidates(data: dict) -> list[ImageCandidate]:
    candidates: list[ImageCandidate] = []
    images = data.get("images") or []
    for item in images:
        if isinstance(item, str):
            candidates.append(ImageCandidate(url=item))
        elif isinstance(item, dict):
            url = item.get("url") or item.get("image_url") or item.get("src")
            if url:
                candidates.append(ImageCandidate(
                    url=url,
                    description=_as_text(item.get("description") or item.get("alt")),
                    page_url=_as_text(item.get("page_url") or item.get("source_url")),
                    title=_as_text(item.get("title")),
                ))
    for result in data.get("results") or []:
        page_url = _as_text(result.get("url"))
        title = _as_text(result.get("title"))
        description = _as_text(result.get("content") or result.get("raw_content"))
        for key in ("image", "image_url", "thumbnail", "thumbnail_url", "og_image"):
            url = result.get(key)
            if isinstance(url, str):
                candidates.append(ImageCandidate(
                    url=url,
                    page_url=page_url,
                    title=title,
                    description=description,
                ))
    seen = set()
    unique = []
    for cand in candidates:
        if cand.url in seen:
            continue
        seen.add(cand.url)
        unique.append(cand)
    return unique


def _bad_image_url(url: str) -> bool:
    lowered = urllib.parse.unquote(url).lower()
    return any(token in lowered for token in BAD_IMAGE_TOKENS)


def _candidate_text(cand: ImageCandidate) -> str:
    return " ".join([cand.title, cand.description, cand.page_url, cand.url]).lower()


def _semantic_ok(cand: ImageCandidate, required_terms: list[str],
                 reject_terms: list[str], *, require_page_url: bool = True,
                 allowed_domains: list[str] | None = None) -> tuple[bool, str]:
    page_domain = _domain_of(cand.page_url)
    image_domain = _domain_of(cand.url)
    if require_page_url and not cand.page_url:
        if not (allowed_domains and _domain_allowed(image_domain, allowed_domains)):
            return False, "missing traceable source page"
    if allowed_domains:
        if not (
            _domain_allowed(page_domain, allowed_domains)
            or _domain_allowed(image_domain, allowed_domains)
        ):
            return False, "source domain not allowed"
    text = _candidate_text(cand)
    for term in reject_terms:
        if term and term.lower() in text:
            return False, f"semantic reject term matched: {term}"
    if not required_terms:
        return True, "ok"
    matched = [term for term in required_terms if term.lower() in text]
    if matched:
        return True, "ok"
    return False, "semantic mismatch: missing required terms"


def _aspect_ok(width: int, height: int, target_ratio: float, tolerance: float) -> bool:
    if not width or not height or not target_ratio:
        return False
    ratio = width / height
    drift = max(ratio / target_ratio, target_ratio / ratio)
    return drift <= tolerance


def download_candidate(cand: ImageCandidate, out_path: pathlib.Path, *,
                       min_width: int, min_height: int, target_ratio: float,
                       aspect_tolerance: float, max_bytes: int,
                       required_terms: list[str] | None = None,
                       reject_terms: list[str] | None = None,
                       require_page_url: bool = True,
                       allowed_domains: list[str] | None = None) -> tuple[bool, str]:
    ok, reason = _semantic_ok(
        cand,
        required_terms or [],
        reject_terms or [],
        require_page_url=require_page_url,
        allowed_domains=allowed_domains,
    )
    if not ok:
        return False, reason
    if _bad_image_url(cand.url):
        return False, "url rejected as logo/icon/placeholder"
    req = urllib.request.Request(cand.url, headers={"User-Agent": "PPT-Agent image resolver"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            blob = resp.read(max_bytes + 1)
    except urllib.error.URLError as exc:
        return False, f"download failed: {exc}"
    if len(blob) > max_bytes:
        return False, "image too large"
    try:
        with Image.open(io.BytesIO(blob)) as img:
            img.load()
            width, height = img.size
            if width < min_width or height < min_height:
                return False, f"too small: {width}x{height}"
            if not _aspect_ok(width, height, target_ratio, aspect_tolerance):
                return False, f"aspect mismatch: {width / height:.2f} vs {target_ratio:.2f}"
            out_path.parent.mkdir(parents=True, exist_ok=True)
            if img.mode in ("RGBA", "LA"):
                img.convert("RGBA").save(out_path)
            else:
                img.convert("RGB").save(out_path)
    except Exception as exc:
        return False, f"invalid image: {exc}"
    return True, "ok"


def _download_first_acceptable(candidates: list[ImageCandidate], *, page: dict,
                               slot_index: int, query: str, asset_dir: pathlib.Path,
                               min_width: int, min_height: int,
                               target_ratio: float, aspect_tolerance: float,
                               max_bytes: int, required_terms: list[str],
                               reject_terms: list[str], require_page_url: bool,
                               allowed_domains: list[str],
                               status: str) -> tuple[dict | None, list[str]]:
    reasons = []
    for rank, cand in enumerate(candidates, start=1):
        filename = (
            f"slide{int(page.get('slide', 0)):02d}_image{slot_index + 1:02d}_"
            f"{_safe_name(query, fallback='query')}_{rank}.png"
        )
        out_path = asset_dir / DEFAULT_ASSET_SUBDIR / filename
        ok, reason = download_candidate(
            cand, out_path,
            min_width=min_width,
            min_height=min_height,
            target_ratio=target_ratio,
            aspect_tolerance=aspect_tolerance,
            max_bytes=max_bytes,
            required_terms=required_terms,
            reject_terms=reject_terms,
            require_page_url=require_page_url,
            allowed_domains=allowed_domains,
        )
        if ok:
            return {
                "path": str(out_path),
                "url": cand.url,
                "page_url": cand.page_url,
                "description": cand.description,
                "status": status,
            }, reasons
        reasons.append(f"{rank}: {reason}")
    return None, reasons


def _ensure_image_item(item: Any) -> dict:
    if isinstance(item, dict):
        out = dict(item)
    elif isinstance(item, str) and item.strip():
        out = {"source": "local", "path": item.strip(), "fallback": "keep_original"}
    else:
        out = {"source": "web", "provider": "tavily", "query": "", "path": "", "fallback": "keep_original"}
    out.setdefault("fallback", "keep_original")
    return out


def adapt_plan_images(plan: dict, *, api_key: str, asset_dir: pathlib.Path,
                      selected_pages: set[int] | None, query_prefix: str,
                      max_results: int, search_depth: str, min_width: int,
                      min_height: int, aspect_tolerance: float, max_bytes: int,
                      force: bool, dry_run: bool) -> dict:
    aspects = _slot_aspects(plan)
    out_plan = dict(plan)
    pages = []

    for page in plan.get("pages", []):
        page_out = dict(page)
        if selected_pages is not None and page.get("slide") not in selected_pages:
            pages.append(page_out)
            continue
        if "images" not in page:
            pages.append(page_out)
            continue

        image_items = []
        for slot_index, raw_item in enumerate(page.get("images") or []):
            item = _ensure_image_item(raw_item)
            if item.get("path") and not force:
                image_items.append(item)
                continue
            if item.get("source") == "keep_original":
                image_items.append(item)
                continue
            provider = item.get("provider") or item.get("source")
            if provider not in ("tavily", "web", ""):
                image_items.append(item)
                continue

            query = build_query(plan, page, item, query_prefix=query_prefix)
            item["source"] = "web"
            item["provider"] = "tavily"
            item["query"] = query
            target_ratio = _target_ratio(aspects, page.get("slide"), slot_index)
            allowed_domains = _allowed_domains(item)
            if _requires_owned_source(item) and not allowed_domains:
                item["path"] = ""
                item["status"] = "needs_owned_source"
                item["reject_reasons"] = [
                    "product/platform screenshots require allowed_domains; keeping original"
                ]
                image_items.append(item)
                continue
            if dry_run:
                item["status"] = "dry_run"
                image_items.append(item)
                continue

            try:
                candidates = tavily_search_images(
                    api_key, query, max_results=max_results, search_depth=search_depth
                )
            except Exception as exc:
                item["status"] = f"search_failed: {exc}"
                image_items.append(item)
                continue

            required_terms, reject_terms = semantic_terms(plan, page, item)
            item["required_terms"] = required_terms
            item.setdefault("require_page_url", True)
            if reject_terms:
                item["reject_terms"] = reject_terms
            accepted, reasons = _download_first_acceptable(
                candidates,
                page=page,
                slot_index=slot_index,
                query=query,
                asset_dir=asset_dir,
                min_width=min_width,
                min_height=min_height,
                target_ratio=target_ratio,
                aspect_tolerance=aspect_tolerance,
                max_bytes=max_bytes,
                required_terms=required_terms,
                reject_terms=reject_terms,
                require_page_url=bool(item.get("require_page_url", True)),
                allowed_domains=allowed_domains,
                status="resolved",
            )
            if not accepted and item.get("semantic_fallback", "strict") == "relaxed":
                accepted, relaxed_reasons = _download_first_acceptable(
                    candidates,
                    page=page,
                    slot_index=slot_index,
                    query=query,
                    asset_dir=asset_dir,
                    min_width=min_width,
                    min_height=min_height,
                    target_ratio=target_ratio,
                    aspect_tolerance=aspect_tolerance,
                    max_bytes=max_bytes,
                    required_terms=[],
                    reject_terms=reject_terms,
                    require_page_url=False,
                    allowed_domains=allowed_domains,
                    status="resolved_relaxed",
                )
                if accepted:
                    item["semantic_warning"] = "strict required_terms did not match; relaxed fallback accepted candidate"
                reasons = reasons + [f"relaxed {reason}" for reason in relaxed_reasons]
            if accepted:
                item.update(accepted)
            else:
                item["path"] = ""
                item["status"] = "no_acceptable_image"
                item["reject_reasons"] = reasons[:6]
            image_items.append(item)
        page_out["images"] = image_items
        pages.append(page_out)

    out_plan["pages"] = pages
    return out_plan


def _selected_pages(value: str | None) -> set[int] | None:
    if not value:
        return None
    return {int(x) for x in value.split(",") if x.strip()}


def main() -> int:
    load_dotenv()
    ap = argparse.ArgumentParser(description="Resolve content_plan image slots through Tavily.")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--pages", default=None, help="comma-separated slide indices")
    ap.add_argument("--query-prefix", default="")
    ap.add_argument("--max-results", type=int, default=8)
    ap.add_argument("--search-depth", default="basic", choices=["basic", "advanced"])
    ap.add_argument("--min-width", type=int, default=640)
    ap.add_argument("--min-height", type=int, default=360)
    ap.add_argument("--aspect-tolerance", type=float, default=2.2)
    ap.add_argument("--max-bytes", type=int, default=8_000_000)
    ap.add_argument("--force", action="store_true", help="replace existing image path fields")
    ap.add_argument("--dry-run", action="store_true", help="build queries without calling Tavily")
    args = ap.parse_args()

    plan_path = pathlib.Path(args.plan)
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    asset_dir = pathlib.Path(plan.get("asset_dir") or "workspace/assets")
    api_key = os.environ.get("TAVILY_API_KEY", "")
    if not api_key and not args.dry_run:
        raise SystemExit("TAVILY_API_KEY is required; do not put the key in the plan or skill files")

    out_plan = adapt_plan_images(
        plan,
        api_key=api_key,
        asset_dir=asset_dir,
        selected_pages=_selected_pages(args.pages),
        query_prefix=args.query_prefix,
        max_results=args.max_results,
        search_depth=args.search_depth,
        min_width=args.min_width,
        min_height=args.min_height,
        aspect_tolerance=args.aspect_tolerance,
        max_bytes=args.max_bytes,
        force=args.force,
        dry_run=args.dry_run,
    )

    out_path = pathlib.Path(args.out) if args.out else plan_path.with_suffix(".with_images.content_plan.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(out_plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"image plan -> {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
