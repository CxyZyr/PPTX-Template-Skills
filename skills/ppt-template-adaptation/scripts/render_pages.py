"""render_pages.py — render a deck (or a page range) to PNG for visual validation.

Rendering, not code inspection, is the source of truth for whether a page passes.
Use this after apply_content_plan.py to inspect pages one at a time, following the
page-by-page validation discipline (see ../SKILL.md and ../references/workflow.md).

Usage:
  python3 render_pages.py --pptx workspace/out/gcl_strategy.pptx                 # all pages
  python3 render_pages.py --pptx workspace/out/gcl_strategy.pptx --first 4 --last 4   # one page
"""
from __future__ import annotations

import argparse
import pathlib
import sys

SKILL_ROOT = pathlib.Path(__file__).resolve().parents[1]
REPO_ROOT = pathlib.Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(SKILL_ROOT))

from lib.pptx_toolkit import render  # noqa: E402

PAGE_PASS_CHECKLIST = """\
PAGE-PASS CHECKLIST (read the rendered PNG, not the code):
  text      : title readable, body readable, no overflow, no ugly wrap, no drift after font changes
  graphics  : logo correct or intentionally absent; icons match the page meaning; images on-topic & balanced
  template  : page still feels like the template; no placeholder text; no stale sample number / year / PART
  decision  : if any issue -> send the page back (edit only this page's plan entry) and re-render; else pass
"""


def main() -> None:
    ap = argparse.ArgumentParser(description="Render a deck to PNG for per-page validation.")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--out", default=None, help="output dir for PNGs (default: <pptx_dir>/renders)")
    ap.add_argument("--dpi", type=int, default=130)
    ap.add_argument("--first", type=int, default=None, help="1-based first page (inclusive)")
    ap.add_argument("--last", type=int, default=None, help="1-based last page (inclusive)")
    args = ap.parse_args()

    pptx_path = pathlib.Path(args.pptx)
    out_dir = pathlib.Path(args.out) if args.out else pptx_path.parent / "renders"

    pdf_path = render.export_pdf(pptx_path, out_dir)
    pngs = render.render_pdf_to_pngs(pdf_path, out_dir, dpi=args.dpi,
                                     first=args.first, last=args.last)
    print(f"rendered {len(pngs)} page(s) -> {out_dir}")
    for p in pngs:
        print(f"  {p}")
    print()
    print(PAGE_PASS_CHECKLIST)


if __name__ == "__main__":
    main()
