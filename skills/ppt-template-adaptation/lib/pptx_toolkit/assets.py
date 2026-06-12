"""Icon and image asset preparation.

Extracted from scripts/generate_strategy_ppt_ai_tech.py (ensure_png_from_svg,
build_mask_icon, prepare_brand_assets pattern) and the image-crop helper from
scripts/generate_strategy_ppt.py's media path.

Pipeline for semantic icons: Bootstrap-Icons SVG -> PNG (rsvg-convert) ->
optional recolor (alpha mask). Everything is cached on disk by name.
"""
from __future__ import annotations

import subprocess
import urllib.request
from pathlib import Path

from PIL import Image

BOOTSTRAP_ICONS_VERSION = "1.11.3"


def bootstrap_icon_url(name: str) -> str:
    return (
        f"https://cdn.jsdelivr.net/npm/bootstrap-icons@{BOOTSTRAP_ICONS_VERSION}"
        f"/icons/{name}.svg"
    )


def download_svg(name: str, svg_path: Path) -> None:
    svg_path.parent.mkdir(parents=True, exist_ok=True)
    urllib.request.urlretrieve(bootstrap_icon_url(name), str(svg_path))


def ensure_png_from_svg(svg_path: Path, png_path: Path, *, size: int = 512) -> None:
    if png_path.exists():
        return
    png_path.parent.mkdir(parents=True, exist_ok=True)
    with png_path.open("wb") as fp:
        subprocess.run(
            ["rsvg-convert", "-w", str(size), "-h", str(size), str(svg_path)],
            check=True,
            stdout=fp,
        )


def build_mask_icon(source_path: Path, output_path: Path, rgb: tuple[int, int, int]) -> None:
    """Recolor a PNG to a solid `rgb`, keeping its alpha (monochrome icon variant)."""
    if output_path.exists():
        return
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(source_path).convert("RGBA") as img:
        alpha = img.getchannel("A")
        masked = Image.new("RGBA", img.size, (*rgb, 0))
        masked.putalpha(alpha)
        masked.save(output_path)


def get_icon(name: str, asset_dir: Path, *, color: tuple[int, int, int] | None = None,
             size: int = 512) -> Path:
    """High-level: fetch a Bootstrap icon by name, cache SVG+PNG under asset_dir,
    optionally produce a recolored variant. Returns the PNG path to insert."""
    asset_dir = Path(asset_dir)
    svg_path = asset_dir / f"{name}.svg"
    png_path = asset_dir / f"{name}.png"
    if not svg_path.exists():
        download_svg(name, svg_path)
    ensure_png_from_svg(svg_path, png_path, size=size)
    if color is None:
        return png_path
    tag = "_".join(str(c) for c in color)
    colored = asset_dir / f"{name}_{tag}.png"
    build_mask_icon(png_path, colored, color)
    return colored


def crop_to_fill(image_path: Path, output_path: Path, *, target_ratio: float) -> Path:
    """Center-crop an image to a target width/height ratio (no stretching)."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(image_path).convert("RGB") as img:
        w, h = img.size
        current_ratio = w / h
        if current_ratio > target_ratio:
            new_w = int(h * target_ratio)
            left = (w - new_w) // 2
            cropped = img.crop((left, 0, left + new_w, h))
        else:
            new_h = int(w / target_ratio)
            top = (h - new_h) // 2
            cropped = img.crop((0, top, w, top + new_h))
        cropped.save(output_path)
    return output_path
