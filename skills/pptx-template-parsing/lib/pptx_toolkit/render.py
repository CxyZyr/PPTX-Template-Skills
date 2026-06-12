"""Rendering for visual validation: PPTX -> PDF -> per-slide PNG.

Extracted from skills/ppt-generation-playbook/scripts/05_visual_validation.py.
Uses LibreOffice headless (PDF export) + pdftoppm (rasterize). Rendering, not
code inspection, is the source of truth for whether a page is acceptable.
"""
from __future__ import annotations

import os
import subprocess
from pathlib import Path

DEFAULT_LO_PROFILE = "file:///tmp/libreoffice-profile-pptx-toolkit"


def export_pdf(pptx_path, out_dir, *, profile: str = DEFAULT_LO_PROFILE) -> Path:
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = out_dir / f"{pptx_path.stem}.pdf"
    pptx_mtime = pptx_path.stat().st_mtime
    if pdf_path.exists():
        pdf_path.unlink()
    env = os.environ.copy()
    env["HOME"] = "/tmp"
    env["XDG_RUNTIME_DIR"] = "/tmp"
    env["GSETTINGS_BACKEND"] = "memory"
    result = subprocess.run(
        [
            "libreoffice",
            f"-env:UserInstallation={profile}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(pptx_path),
        ],
        check=False,
        env=env,
        capture_output=True,
        text=True,
    )
    pdf_is_fresh = pdf_path.exists() and pdf_path.stat().st_mtime >= pptx_mtime
    if result.returncode != 0:
        if pdf_is_fresh:
            print(f"[warn] libreoffice returned {result.returncode}, but PDF was produced: {pdf_path}")
        else:
            if result.stdout:
                print(result.stdout)
            if result.stderr:
                print(result.stderr)
            result.check_returncode()
    elif not pdf_is_fresh:
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
        raise FileNotFoundError(f"libreoffice did not produce a fresh PDF: {pdf_path}")
    return pdf_path


def render_pdf_to_pngs(pdf_path, out_dir, *, dpi: int = 150, first: int | None = None,
                       last: int | None = None, prefix: str = "slide") -> list[Path]:
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob(f"{prefix}*.png"):
        old.unlink()
    cmd = ["pdftoppm", "-r", str(dpi), "-png"]
    if first is not None:
        cmd += ["-f", str(first)]
    if last is not None:
        cmd += ["-l", str(last)]
    cmd += [str(pdf_path), str(out_dir / prefix)]
    subprocess.run(cmd, check=True)
    return sorted(out_dir.glob(f"{prefix}*.png"))


def render_deck(pptx_path, out_dir, *, dpi: int = 150, prefix: str = "slide") -> list[Path]:
    """Convenience: export to PDF and rasterize every page to PNG.
    Returns the list of PNG paths in slide order."""
    out_dir = Path(out_dir)
    pdf_path = export_pdf(pptx_path, out_dir)
    return render_pdf_to_pngs(pdf_path, out_dir, dpi=dpi, prefix=prefix)
