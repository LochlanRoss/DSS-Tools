#!/usr/bin/env python3
"""Create repo-root ``dss_tools.ico`` for Windows (PyInstaller PE icon + Inno Setup shortcuts).

Order of resolution (repository root only):
1. If ``dss_tools.ico`` already exists — done.
2. If there is exactly one other ``*.ico`` file — copy it to ``dss_tools.ico``.
3. If a known PNG exists — convert to a multi-resolution ICO (requires Pillow).
4. Otherwise copy ``tools/default_dss_tools.ico`` (built-in placeholder) unless ``--strict``.

PNG file names checked first: ``dss_tools.png``, ``DSS-Tools Icon.png``, ``DSSTools Icon.png``,
then any single root ``*.png`` whose name contains ``icon`` (case-insensitive). More than one
such loose match is an error.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
TOOLS_DIR = Path(__file__).resolve().parent
PNG_CANDIDATE_NAMES = ("dss_tools.png", "DSS-Tools Icon.png", "DSSTools Icon.png")
ICO_SIZES = (16, 24, 32, 48, 64, 128, 256)


def _find_lone_other_ico(repo: Path, target: Path) -> Path | None:
    icos = sorted(
        p for p in repo.iterdir() if p.is_file() and p.suffix.lower() == ".ico" and p.resolve() != target.resolve()
    )
    if len(icos) == 1:
        return icos[0]
    return None


def _resolve_png(repo: Path) -> Path | None:
    for name in PNG_CANDIDATE_NAMES:
        candidate = repo / name
        if candidate.is_file():
            return candidate
    loose = sorted(
        p
        for p in repo.iterdir()
        if p.is_file() and p.suffix.lower() == ".png" and "icon" in p.name.lower()
    )
    if len(loose) == 1:
        return loose[0]
    if len(loose) > 1:
        names = ", ".join(p.name for p in loose)
        print(
            f"error: multiple icon-like PNG files in repo root ({names}). "
            f"Keep one, or use a fixed name from: {', '.join(PNG_CANDIDATE_NAMES)}.",
            file=sys.stderr,
        )
        raise SystemExit(1)
    return None


def _png_to_ico(png: Path, ico: Path) -> None:
    from PIL import Image

    image = Image.open(png).convert("RGBA")
    ico.parent.mkdir(parents=True, exist_ok=True)
    image.save(ico, format="ICO", sizes=[(s, s) for s in ICO_SIZES])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Do not use the built-in placeholder; exit 1 if no real icon or PNG source is available.",
    )
    parser.add_argument(
        "--repo",
        type=Path,
        default=REPO_ROOT,
        help="Repository root (default: parent of tools/).",
    )
    args = parser.parse_args(argv)
    repo = args.repo.resolve()
    target = repo / "dss_tools.ico"

    if target.is_file():
        print(f"OK: {target.name} already present.")
        return 0

    lone = _find_lone_other_ico(repo, target)
    if lone is not None:
        shutil.copy2(lone, target)
        print(f"OK: copied {lone.name} -> {target.name}")
        return 0

    png = _resolve_png(repo)
    if png is not None:
        try:
            _png_to_ico(png, target)
        except ImportError:
            print("error: Pillow is required to convert PNG to ICO (`pip install pillow`).", file=sys.stderr)
            return 1
        except OSError as exc:
            print(f"error: could not write {target}: {exc}", file=sys.stderr)
            return 1
        print(f"OK: wrote {target.name} from {png.name}")
        return 0

    if not args.strict:
        fallback = TOOLS_DIR / "default_dss_tools.ico"
        if fallback.is_file():
            shutil.copy2(fallback, target)
            print(
                "WARN: using built-in placeholder for dss_tools.ico. "
                "Commit DSS-Tools Icon.png (or dss_tools.ico) at the repo root for your branded icon.",
                file=sys.stderr,
            )
            return 0

    print(
        "error: no dss_tools.ico, lone .ico, PNG icon source, or (when --strict) no fallback allowed.",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
