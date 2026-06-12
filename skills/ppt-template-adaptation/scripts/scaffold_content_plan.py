"""scaffold_content_plan.py — generate a skeleton content_plan.json from a spec.json.

Reads a parsed spec and emits a content_plan with one page entry per slide, each
carrying exactly the slots that slide actually has (N cards, M body boxes, K images,
logo flag, ...). You then fill in the text/icons/images. A `slide_guide` block (ignored
by apply_content_plan.py) records each slide's role + the template's original sample
text, so you can match slots to the render while editing.

Usage:
  python3 scaffold_content_plan.py --spec workspace/specs/ai_tech/spec.json \
      --out workspace/plans/ai_tech.content_plan.json \
      --output-pptx workspace/out/ai_tech.pptx
"""
from __future__ import annotations

import argparse
import json
import pathlib


def page_skeleton(slide: dict) -> dict:
    fp = slide["fill_plan"]
    page: dict = {"slide": slide["index"], "role": slide["role"]}
    if fp.get("title"):
        page["title"] = ""
    if fp.get("subtitle"):
        page["subtitle"] = ""
    if fp.get("section_number"):
        page["section_number"] = ""
    if fp.get("footer"):
        page["footer"] = fp["footer"].get("sample") or ""
    if fp.get("logo"):
        page["logo"] = ""          # path to a logo image (relative to asset_dir)
    if fp.get("labels"):
        page["labels"] = [
            lab.get("sample") if lab.get("default_action") == "preserve" else ""
            for lab in fp["labels"]
        ]
    if fp.get("body"):
        page["body"] = ["" for _ in fp["body"]]
    if fp.get("images"):
        page["images"] = [
            {
                "source": "web",
                "provider": "tavily",
                "query": "",
                "path": "",
                "fallback": "keep_original",
            }
            for _ in fp["images"]
        ]
    if fp.get("cards"):
        cards = []
        for c in fp["cards"]:
            card: dict = {}
            if c.get("title"):
                card["title"] = ""
            if c.get("body"):
                card["body"] = ""
            if c.get("number"):
                card["number"] = ""
            if c.get("icon"):
                card["icon"] = ""   # a bootstrap-icons name, e.g. "database"
            cards.append(card)
        page["cards"] = cards
    return page


def slide_hint(slide: dict) -> dict:
    fp = slide["fill_plan"]
    hint: dict = {
        "slide": slide["index"],
        "role": slide["role"],
        "role_confidence": slide["role_confidence"],
    }
    if fp.get("title"):
        hint["title_sample"] = fp["title"].get("sample")
    if fp.get("section_number"):
        hint["section_number_sample"] = fp["section_number"].get("sample")
    if fp.get("body"):
        hint["body_slots"] = len(fp["body"])
    if fp.get("labels"):
        hint["label_samples"] = [lab.get("sample") for lab in fp["labels"]]
    if fp.get("images"):
        hint["image_slots"] = len(fp["images"])
    if fp.get("cards"):
        hint["card_title_samples"] = [c.get("title_sample") for c in fp["cards"]]
        hint["card_has_icon"] = [bool(c.get("icon")) for c in fp["cards"]]
    return hint


def main() -> None:
    ap = argparse.ArgumentParser(description="Scaffold a content_plan.json from a spec.json.")
    ap.add_argument("--spec", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--output-pptx", default="workspace/out/deck.pptx",
                    help="where apply_content_plan.py should save the filled deck")
    ap.add_argument("--asset-dir", default="workspace/assets")
    args = ap.parse_args()

    spec = json.loads(pathlib.Path(args.spec).read_text(encoding="utf-8"))
    theme = spec["template"].get("theme", {})

    plan = {
        "spec": args.spec,
        "output": args.output_pptx,
        "asset_dir": args.asset_dir,
        "brand": {
            "title_font": (theme.get("common_fonts") or [None])[0],
            "accent": (theme.get("common_colors") or [None])[0],
        },
        "keep_slides": [s["index"] for s in spec["slides"]],
        "slide_guide": [slide_hint(s) for s in spec["slides"]],
        "pages": [page_skeleton(s) for s in spec["slides"]],
    }

    out = pathlib.Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"scaffold -> {out}  ({len(plan['pages'])} pages)")
    print("edit `pages` to fill content; trim `keep_slides` to drop unused slides.")
    print("`slide_guide` shows each slide's role + original sample text (ignored by the applier).")


if __name__ == "__main__":
    main()
