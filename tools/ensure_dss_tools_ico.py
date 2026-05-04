#!/usr/bin/env python3
"""Create repo-root ``dss_tools.ico`` for Windows (PyInstaller PE icon + Inno Setup).

**Branded repositories:** If any **priority** branding PNG exists at the repo root (see
``BRAND_PNG_PRIORITY``), ``dss_tools.ico`` is **always** produced from the first match only
(no lone ``*.ico`` shortcut, no placeholder). Existing
``dss_tools.ico`` is verified against the PNG and regenerated when it does not match. Each ICO
size strips a **near-white outer band** to full transparency (removes matte/halos), then paints
a **1px opaque blue rim** sampled from saturated blues in the source art so the outermost pixels
read as brand blue in shell / taskbar previews.

**Other repositories:** Without the canonical PNG, resolution follows (unless ``--strict``):

1. If ``dss_tools.ico`` already exists — done (unless ``--force``).
2. If there is exactly one other ``*.ico`` file — copy it to ``dss_tools.ico``.
3. If another known PNG exists — convert to a multi-resolution ICO (requires Pillow).
4. Otherwise copy ``tools/default_dss_tools.ico`` (built-in placeholder) unless ``--strict``.

PNG file names (non-canonical): ``dss_tools.png``, ``DSSTools Icon.png``, then any single root
``*.png`` whose name contains ``icon`` (case-insensitive). More than one loose match is an error.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
TOOLS_DIR = Path(__file__).resolve().parent
# First existing file wins (must match .gitignore ``!`` exceptions for tracked PNGs).
BRAND_PNG_PRIORITY = (
    "DSS-Tools Icon.png",
    "DSS Tools Icon.png",
    "DSS-Tools-Icon.png",
    "app-icon.png",
    "DSSTools Icon.png",
    "dss_tools.png",
)
CANONICAL_BRAND_PNG = BRAND_PNG_PRIORITY[0]  # preferred name for docs / CI messaging
PNG_CANDIDATE_NAMES = ("dss_tools.png", "DSS-Tools Icon.png", "DSS Tools Icon.png", "DSSTools Icon.png")
ICO_SIZES = (16, 24, 32, 48, 64, 128, 256)
# Fallback BGR when the PNG has no obvious saturated blues to sample.
_FALLBACK_BRAND_BLUE = (15, 76, 129)


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


def _sample_brand_blue_rgb(master_rgba) -> tuple[int, int, int]:
    """Pick an opaque blue from the artwork for the ICO perimeter (RGB)."""
    from PIL import Image

    im = master_rgba.convert("RGBA")
    w, h = im.size
    px = im.load()
    blues: list[tuple[int, int, int]] = []
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a < 160:
                continue
            if b >= max(r, g) + 12 and b >= 90:
                blues.append((r, g, b))
    if len(blues) < 24:
        return _FALLBACK_BRAND_BLUE
    blues.sort(key=lambda t: (t[2], t[0], t[1]))
    return blues[len(blues) // 2]


def _transparent_near_white_border_band(source_rgba, *, band_px: int, white_thresh: int = 242) -> None:
    """Turn near-white matte in an outer band fully transparent (mutates *source_rgba* RGBA)."""
    w, h = source_rgba.size
    if w < 2 or h < 2:
        return
    band = max(1, min(int(band_px), min(w, h) // 2))
    px = source_rgba.load()
    for y in range(h):
        for x in range(w):
            d = min(x, y, w - 1 - x, h - 1 - y)
            if d > band:
                continue
            r, g, b, a = px[x, y]
            if a < 40:
                continue
            if r >= white_thresh and g >= white_thresh and b >= white_thresh:
                px[x, y] = (r, g, b, 0)


def _paint_blue_perimeter(source_rgba, blue_rgb: tuple[int, int, int]) -> None:
    """Set the outermost pixel row/column to opaque *blue_rgb* (mutates RGBA)."""
    w, h = source_rgba.size
    if w < 2 or h < 2:
        return
    r0, g0, b0 = blue_rgb
    px = source_rgba.load()
    for x in range(w):
        px[x, 0] = (r0, g0, b0, 255)
        px[x, h - 1] = (r0, g0, b0, 255)
    for y in range(h):
        px[0, y] = (r0, g0, b0, 255)
        px[w - 1, y] = (r0, g0, b0, 255)


def _postprocess_square_frame(frame_rgba, *, edge: int, blue_rgb: tuple[int, int, int]) -> None:
    """Remove white/halation in the outer band, then draw a 1px blue rim (all ICO sizes)."""
    band = max(1, min(edge // 6, 36))
    _transparent_near_white_border_band(frame_rgba, band_px=band, white_thresh=242)
    _paint_blue_perimeter(frame_rgba, blue_rgb)


def _build_square_frame(master_rgba, edge: int, blue_rgb: tuple[int, int, int]):
    """Square aspect-fit from *master_rgba*, then matte strip + blue perimeter."""
    fr = _aspect_fit_square(master_rgba, edge)
    _postprocess_square_frame(fr, edge=edge, blue_rgb=blue_rgb)
    return fr


def _resolve_brand_png(repo: Path) -> Path | None:
    """Return the first existing branding PNG under *repo* (``BRAND_PNG_PRIORITY`` order)."""
    for name in BRAND_PNG_PRIORITY:
        candidate = repo / name
        if candidate.is_file():
            return candidate
    return None


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
    blue_rgb = _sample_brand_blue_rgb(master)
    ico.parent.mkdir(parents=True, exist_ok=True)
    frames = {s: _build_square_frame(master, s, blue_rgb) for s in ICO_SIZES}
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

    master = Image.open(png).convert("RGBA")
    blue_rgb = _sample_brand_blue_rgb(master)
    ref_div = _aspect_fit_square(master, 64)
    diversity = len(set(ref_div.convert("RGB").getdata()))
    if diversity < 64:
        # Solid / test patterns: ICO BMP/PNG round-trip differs a lot from RGBA thumbnails; skip pixels.
        return

    ref = _build_square_frame(master, 256, blue_rgb)
    ref_rgb = ref.convert("RGB")
    with Image.open(ico) as produced:
        produced.load()
        cand = produced.convert("RGBA")
    if cand.size != (256, 256):
        raise ValueError(f"ICO primary frame must be 256×256 for verify; got {cand.size}")
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


def _write_branded_ico_from_canonical(target: Path, canonical: Path, *, no_verify: bool) -> int:
    """Regenerate ``dss_tools.ico`` from ``DSS-Tools Icon.png`` and verify."""
    try:
        _png_to_ico(canonical, target)
        if not no_verify:
            _verify_ico_matches_png(canonical, target)
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
    print(f"OK: wrote {target.name} from {canonical.name} (branding PNG; square frames + check)")
    return 0


def _ensure_branded_icon(repo: Path, target: Path, args: argparse.Namespace) -> int | None:
    """If a priority branding PNG exists, enforce ICO from that file only. Returns None if not branded."""
    brand_png = _resolve_brand_png(repo)
    if brand_png is None:
        return None

    need_write = args.force or not target.is_file()
    if not need_write and not args.no_verify:
        try:
            _verify_ico_matches_png(brand_png, target)
        except (ValueError, OSError, ImportError, FileNotFoundError):
            need_write = True

    if need_write:
        return _write_branded_ico_from_canonical(target, brand_png, no_verify=args.no_verify)

    print(f"OK: {target.name} already matches {brand_png.name}.")
    return 0


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
        help="Regenerate dss_tools.ico even if it already matches (canonical PNG or unbranded flow).",
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

    branded = _ensure_branded_icon(repo, target, args)
    if branded is not None:
        return int(branded)

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
                f"Add one of {', '.join(BRAND_PNG_PRIORITY)} at the repo root for release branding.",
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
