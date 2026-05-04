#!/usr/bin/env python3
"""Create repo-root ``dss_tools.ico`` for Windows (PyInstaller PE icon + Inno Setup shortcuts).

Order of resolution (repository root only):
1. If ``dss_tools.ico`` already exists — done (unless ``--force``).
2. If there is exactly one other ``*.ico`` file — copy it to ``dss_tools.ico``.
3. If a known PNG exists — convert to a multi-resolution ICO (requires Pillow). The ICO writer
   saves the **256×256** frame first (Pillow skips larger ``sizes`` if the primary image is small).
   A quick PNG↔ICO pixel check rejects outputs that do not match the source visually.
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


def _aspect_fit_square(source_rgba, edge: int):
    """Return an ``edge``×``edge`` RGBA image: *source* scaled with aspect preserved, centered on transparent."""
    from PIL import Image

    canvas = Image.new("RGBA", (edge, edge), (0, 0, 0, 0))
    layer = source_rgba.copy()
    layer.thumbnail((edge, edge), Image.Resampling.LANCZOS)
    ox = (edge - layer.width) // 2
    oy = (edge - layer.height) // 2
    canvas.paste(layer, (ox, oy), layer)
    return canvas


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
    """Write a multi-size ICO with **square** entries only (Windows/Inno are picky).

    Pillow's ICO writer uses the **first** image's width/height as an upper bound: if the
    primary frame is 16×16, larger ``sizes`` are skipped and only a tiny icon is stored (often
    looks like a solid smear in the installer). The primary frame must therefore be the
    largest size (256×256 here).

    A non-square PNG is first letterboxed onto a square canvas before scaling so frames are
    never 256×255-style sizes that confuse the shell.
    """
    from PIL import Image

    master = Image.open(png).convert("RGBA")
    ico.parent.mkdir(parents=True, exist_ok=True)
    frames = {s: _aspect_fit_square(master, s) for s in ICO_SIZES}
    for s, fr in frames.items():
        if fr.size != (s, s):
            raise ValueError(f"internal error: frame for {s} is {fr.size}, expected {(s, s)}")
    primary = frames[256]
    append = [frames[s] for s in ICO_SIZES if s != 256]
    primary.save(
        ico,
        format="ICO",
        sizes=[(s, s) for s in ICO_SIZES],
        append_images=append,
    )


def _verify_ico_embeds_multi_square_sizes(ico: Path) -> None:
    """Fail if the ICO is missing sizes or still uses a tiny primary (common cause of blue smears)."""
    from PIL import Image

    with Image.open(ico) as im:
        im.load()
        sizes = im.info.get("sizes") or {(im.width, im.height)}
        if len(sizes) < 6:
            raise ValueError(f"ICO should embed many sizes; got {len(sizes)}: {sorted(sizes)}")
        if not all(w == h for w, h in sizes):
            raise ValueError(f"ICO must only contain square frames: {sorted(sizes)}")
        if im.size != (256, 256):
            raise ValueError(
                f"ICO default frame must be 256×256 (Pillow uses the first image as max canvas); got {im.size}"
            )


def _verify_ico_matches_png(png: Path, ico: Path, *, max_mean_abs_rgb: float = 32.0) -> None:
    """Re-read the ICO: structural checks always; RGB comparison only for non-trivial artwork."""
    from PIL import Image

    _verify_ico_embeds_multi_square_sizes(ico)

    ref = _aspect_fit_square(Image.open(png).convert("RGBA"), 64)
    ref_rgb = ref.convert("RGB")
    diversity = len(set(ref_rgb.getdata()))
    if diversity < 64:
        # Solid / test patterns: ICO BMP/PNG round-trip differs a lot from RGBA thumbnails; skip pixels.
        return

    with Image.open(ico) as produced:
        produced.load()
        cand = _aspect_fit_square(produced.convert("RGBA"), 64)
    total = 0
    count = 0
    for a, b in zip(ref_rgb.getdata(), cand.convert("RGB").getdata()):
        total += abs(a[0] - b[0]) + abs(a[1] - b[1]) + abs(a[2] - b[2])
        count += 3
    mae = total / max(count, 1)
    if mae > max_mean_abs_rgb:
        raise ValueError(
            f"generated ICO does not match source PNG visually (mean abs RGB error {mae:.1f}, "
            f"max allowed {max_mean_abs_rgb:.1f}). Try a higher-resolution PNG or check transparency."
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Do not use the built-in placeholder; exit 1 if no real icon or PNG source is available.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Regenerate dss_tools.ico even if it already exists (PNG or lone .ico still wins over placeholder).",
    )
    parser.add_argument(
        "--repo",
        type=Path,
        default=REPO_ROOT,
        help="Repository root (default: parent of tools/).",
    )
    parser.add_argument(
        "--no-verify",
        action="store_true",
        help="Skip PNG vs ICO visual sanity check after conversion (not recommended).",
    )
    args = parser.parse_args(argv)
    repo = args.repo.resolve()
    target = repo / "dss_tools.ico"

    if target.is_file() and not args.force:
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
            if not args.no_verify:
                _verify_ico_matches_png(png, target)
        except ImportError:
            print("error: Pillow is required to convert PNG to ICO (`pip install pillow`).", file=sys.stderr)
            return 1
        except (OSError, ValueError) as exc:
            try:
                target.unlink(missing_ok=True)
            except OSError:
                pass
            print(f"error: could not write or verify {target}: {exc}", file=sys.stderr)
            return 1
        print(f"OK: wrote {target.name} from {png.name} (square frames + visual check)")
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
