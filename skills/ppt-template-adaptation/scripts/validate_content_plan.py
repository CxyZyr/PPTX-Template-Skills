"""Validate cross-page text coherence in a content_plan.

The generator fills pages one at a time, but deck text is not page-local:
contents pages, section dividers, and section titles must agree. This module
derives a canonical outline from plan["outline"] when present, otherwise from
the first contents page, then checks section-divider pages against it.
"""
from __future__ import annotations

import argparse
import json
import pathlib
import re
import sys
from dataclasses import dataclass
from typing import Iterable


NUMBER_RE = re.compile(r"^\s*(?:0?\d+|PART\s*0?\d+|第\s*[一二三四五六七八九十0-9]+\s*[章节部分])\s*$", re.I)
SECTION_NO_RE = re.compile(r"(?:PART\s*)?0*(\d+)|第\s*([一二三四五六七八九十]+)\s*[章节部分]", re.I)

CN_NUM = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


@dataclass
class Issue:
    level: str
    message: str


@dataclass
class OutlineEntry:
    title: str
    subtitle: str = ""
    number: str = ""
    source_slide: int | None = None


def _as_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return " ".join(str(v).strip() for v in value if str(v).strip())
    return str(value).strip()


def _first_line(value) -> str:
    text = _as_text(value)
    return next((line.strip() for line in text.splitlines() if line.strip()), "")


def _rest_lines(value) -> str:
    lines = [line.strip() for line in _as_text(value).splitlines() if line.strip()]
    return " ".join(lines[1:]) if len(lines) > 1 else ""


def _norm(value) -> str:
    return re.sub(r"[\s\W_]+", "", _as_text(value), flags=re.UNICODE).lower()


def _is_number_label(value) -> bool:
    return bool(NUMBER_RE.match(_as_text(value)))


def _section_number_index(value) -> int | None:
    text = _as_text(value)
    match = SECTION_NO_RE.search(text)
    if not match:
        return None
    if match.group(1):
        return int(match.group(1))
    cn = match.group(2)
    if cn in CN_NUM:
        return CN_NUM[cn]
    if cn and cn.startswith("十") and len(cn) == 2:
        return 10 + CN_NUM.get(cn[1], 0)
    if cn and cn.endswith("十") and len(cn) == 2:
        return CN_NUM.get(cn[0], 0) * 10
    return None


def _matches(expected: str, actual: str) -> bool:
    exp = _norm(expected)
    got = _norm(actual)
    if not exp or not got:
        return False
    return exp == got or exp in got or got in exp


def _section_title(page: dict) -> str:
    title = _first_line(page.get("title"))
    if title:
        return title
    for label in page.get("labels") or []:
        text = _as_text(label)
        if text and not _is_number_label(text):
            return text
    return ""


def _section_subtitle(page: dict) -> str:
    subtitle = _rest_lines(page.get("title"))
    if subtitle:
        return subtitle
    body = _as_text(page.get("body"))
    if body:
        return body
    labels = [
        _as_text(label)
        for label in page.get("labels") or []
        if _as_text(label) and not _is_number_label(label)
    ]
    return " ".join(labels[1:]) if len(labels) > 1 else ""


def _explicit_outline(plan: dict) -> list[OutlineEntry]:
    outline = plan.get("outline") or []
    entries: list[OutlineEntry] = []
    for i, item in enumerate(outline, start=1):
        if isinstance(item, str):
            entries.append(OutlineEntry(title=item, number=str(i)))
        elif isinstance(item, dict):
            title = _as_text(item.get("title"))
            if title:
                entries.append(OutlineEntry(
                    title=title,
                    subtitle=_as_text(item.get("subtitle") or item.get("body")),
                    number=_as_text(item.get("number") or item.get("section_number") or i),
                ))
    return entries


def _contents_outline(plan: dict) -> list[OutlineEntry]:
    for page in plan.get("pages", []):
        if page.get("role") != "contents" or not page.get("cards"):
            continue
        entries: list[OutlineEntry] = []
        for i, card in enumerate(page.get("cards", []), start=1):
            title = _as_text(card.get("title"))
            body = _as_text(card.get("body"))
            if _is_number_label(title) and body:
                entries.append(OutlineEntry(
                    title=body,
                    number=_as_text(card.get("number") or title or i),
                    source_slide=page.get("slide"),
                ))
            elif title:
                entries.append(OutlineEntry(
                    title=title,
                    subtitle=body,
                    number=_as_text(card.get("number") or i),
                    source_slide=page.get("slide"),
                ))
        if entries:
            return entries
    return []


def _contents_nonnumber_labels(page: dict) -> list[str]:
    """All non-number labels on a contents page, in slot order.

    Section names and decorative words (e.g. a 'CONTENTS'/'目录' banner) both
    land in labels[]. The caller matches outline titles as an ordered
    subsequence so extra decorative labels do not cause positional drift.
    """
    labels = []
    for label in page.get("labels") or []:
        text = _as_text(label)
        if not text or _is_number_label(text):
            continue
        labels.append(text)
    return labels


def canonical_outline(plan: dict) -> list[OutlineEntry]:
    return _explicit_outline(plan) or _contents_outline(plan)


def _iter_pages(plan: dict, selected_pages: set[int] | None) -> Iterable[dict]:
    for page in plan.get("pages", []):
        if selected_pages is not None and page.get("slide") not in selected_pages:
            continue
        yield page


def _page_body_lines(page: dict) -> list[str]:
    value = page.get("body")
    if isinstance(value, list):
        return [str(v) for v in value if str(v).strip()]
    if value is not None and str(value).strip():
        return [str(value)]
    return []


def _image_item_active(item) -> bool:
    """An image item that intends to fill a slot (has a path or a web provider)."""
    if isinstance(item, str):
        return bool(item.strip())
    if isinstance(item, dict):
        if str(item.get("path") or "").strip():
            return True
        provider = item.get("provider") or item.get("source")
        return provider in ("tavily", "web")
    return bool(item)


def _structural_issues(plan: dict, spec: dict,
                       selected_pages: set[int] | None) -> list[Issue]:
    """Preflight the plan against the parsed slots so content that the template
    has no slot for is blocked before apply, instead of being silently dropped
    with a runtime warning. Mirrors the slot-matching acceptance criteria in
    references/workflow.md. Missing-icon is intentionally left to apply-time
    (warn + skip) per references/spec-schema.md."""
    issues: list[Issue] = []
    fill_plan_by_index = {s.get("index"): s.get("fill_plan", {}) for s in spec.get("slides", [])}
    for page in _iter_pages(plan, selected_pages):
        idx = page.get("slide")
        fp = fill_plan_by_index.get(idx)
        if fp is None:
            continue

        page_cards = page.get("cards") or []
        slot_cards = fp.get("cards") or []
        if len(page_cards) > len(slot_cards):
            issues.append(Issue(
                "error",
                f"slide {idx}: content_plan has {len(page_cards)} card(s) but the template slide has "
                f"{len(slot_cards)} card slot(s); extra cards would be dropped — pick a page with enough slots",
            ))

        if _page_body_lines(page) and not (fp.get("body") or []):
            issues.append(Issue(
                "error",
                f"slide {idx}: content_plan has body text but the template slide has no body slot; "
                f"the text would be dropped — pick a page with a body slot",
            ))

        page_images = [it for it in (page.get("images") or []) if _image_item_active(it)]
        slot_images = fp.get("images") or []
        if len(page_images) > len(slot_images):
            issues.append(Issue(
                "error",
                f"slide {idx}: content_plan has {len(page_images)} image(s) but the template slide has "
                f"{len(slot_images)} image slot(s); extra images would be dropped",
            ))

        if fp.get("charts") and not page.get("charts"):
            issues.append(Issue(
                "error",
                f"slide {idx}: template has chart slot(s), but content_plan has no charts data",
            ))
        if fp.get("tables") and not page.get("tables"):
            issues.append(Issue(
                "error",
                f"slide {idx}: template has table slot(s), but content_plan has no tables data",
            ))
    return issues


def validate_plan(plan: dict, *, selected_pages: set[int] | None = None,
                  spec: dict | None = None) -> list[Issue]:
    issues: list[Issue] = []
    outline = canonical_outline(plan)
    if not outline:
        needs_outline = [
            page for page in plan.get("pages", [])
            if page.get("role") in ("section_divider", "contents")
        ]
        if needs_outline:
            slides = ", ".join(str(p.get("slide")) for p in needs_outline)
            issues.append(Issue(
                "error",
                f"deck has section/contents page(s) (slides {slides}) but no outline source; "
                f"add content_plan.outline[] or a contents page with section cards before writing section pages",
            ))
        else:
            issues.append(Issue(
                "warn",
                "no explicit outline or contents cards found; global text coherence was not checked",
            ))
        if spec:
            issues.extend(_structural_issues(plan, spec, selected_pages))
        return issues

    for page in _iter_pages(plan, selected_pages):
        if page.get("role") != "contents" or page.get("cards"):
            continue
        labels = _contents_nonnumber_labels(page)
        # Match outline titles as an ordered subsequence: each outline title must
        # appear in order, but decorative labels (banners, English captions)
        # between them are skipped rather than treated as section names.
        cursor = 0
        for entry in outline:
            if not entry.title:
                cursor += 1
                continue
            hit = next(
                (j for j in range(cursor, len(labels)) if _matches(entry.title, labels[j])),
                None,
            )
            if hit is None:
                issues.append(Issue(
                    "error",
                    f"slide {page.get('slide')}: contents page is missing section '{entry.title}' "
                    f"from the outline (labels: {labels})",
                ))
            else:
                cursor = hit + 1

    all_section_pages = [
        page for page in plan.get("pages", [])
        if page.get("role") == "section_divider"
    ]
    section_pages = [
        page for page in all_section_pages
        if selected_pages is None or page.get("slide") in selected_pages
    ]
    # Section-divider consistency runs only when this scope actually contains
    # dividers, but structural slot checks (below) must always run.
    if section_pages:
        for position, page in enumerate(all_section_pages, start=1):
            if selected_pages is not None and page.get("slide") not in selected_pages:
                continue
            if position > len(outline):
                issues.append(Issue(
                    "error",
                    f"slide {page.get('slide')}: section divider has no matching contents/outline entry #{position}",
                ))
                continue
            expected = outline[position - 1]
            actual_title = _section_title(page)
            actual_subtitle = _section_subtitle(page)
            actual_no = _as_text(page.get("section_number"))
            actual_no_index = _section_number_index(actual_no)

            if actual_no_index is not None and actual_no_index != position:
                issues.append(Issue(
                    "error",
                    f"slide {page.get('slide')}: section number '{actual_no}' points to #{actual_no_index}, expected #{position}",
                ))
            if expected.title and actual_title and not _matches(expected.title, actual_title):
                issues.append(Issue(
                    "error",
                    f"slide {page.get('slide')}: section title '{actual_title}' does not match outline '{expected.title}'",
                ))
            if expected.subtitle and actual_subtitle and not _matches(expected.subtitle, actual_subtitle):
                issues.append(Issue(
                    "error",
                    f"slide {page.get('slide')}: section subtitle '{actual_subtitle}' does not match outline subtitle '{expected.subtitle}'",
                ))

        if selected_pages is None and len(section_pages) < len(outline):
            issues.append(Issue(
                "warn",
                f"outline has {len(outline)} entries but only {len(section_pages)} section divider page(s) are present",
            ))

    if spec:
        issues.extend(_structural_issues(plan, spec, selected_pages))
    return issues


def print_report(issues: list[Issue]) -> None:
    if not issues:
        print("global text coherence: ok")
        return
    print("global text coherence:")
    for issue in issues:
        print(f"  {issue.level}: {issue.message}")


def main() -> int:
    ap = argparse.ArgumentParser(description="Validate global text coherence in a content_plan.json.")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--pages", default=None,
                    help="optional comma-separated slide indices for incremental validation")
    args = ap.parse_args()

    plan = json.loads(pathlib.Path(args.plan).read_text(encoding="utf-8"))
    spec = None
    if plan.get("spec"):
        spec_path = pathlib.Path(plan["spec"])
        if spec_path.exists():
            spec = json.loads(spec_path.read_text(encoding="utf-8"))
    selected = None
    if args.pages:
        selected = {int(x) for x in args.pages.split(",") if x.strip()}
    issues = validate_plan(plan, selected_pages=selected, spec=spec)
    print_report(issues)
    return 1 if any(issue.level == "error" for issue in issues) else 0


if __name__ == "__main__":
    sys.exit(main())
