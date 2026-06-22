#!/usr/bin/env python3
"""Convert site raster images from assets/img to optimized WebP copies."""

from __future__ import annotations

import argparse
import shutil
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageOps, UnidentifiedImageError

try:
    from pillow_heif import register_heif_opener
except ImportError:
    register_heif_opener = None
else:
    register_heif_opener()


SCRIPT_DIR = Path(__file__).resolve().parent
SITE_DIR = SCRIPT_DIR.parent

DEFAULT_SOURCE = SITE_DIR / "assets/img"
DEFAULT_OUTPUT = SITE_DIR / "assets/img-webp"
RASTER_EXTENSIONS = {".jpg", ".jpeg", ".png"}


@dataclass(frozen=True)
class ResizeRule:
    label: str
    mode: str
    size: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="Source image directory.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Output WebP directory.")
    parser.add_argument("--quality", type=int, default=75, help="WebP quality for all converted images.")
    parser.add_argument("--avatar-size", type=int, default=500, help="Maximum avatar side length in pixels.")
    parser.add_argument("--cover-width", type=int, default=1000, help="Maximum cover/group image width in pixels.")
    parser.add_argument("--force", action="store_true", help="Rebuild files even when output is newer.")
    parser.add_argument("--clean", action="store_true", help="Remove the output directory before converting.")
    return parser.parse_args()


def rule_for(relative_path: Path, avatar_size: int, cover_width: int) -> ResizeRule:
    first = relative_path.parts[0] if relative_path.parts else ""
    if first == "members":
        return ResizeRule("avatar", "max_side", avatar_size)
    if first in {"group", "videos"}:
        return ResizeRule("cover", "width", cover_width)
    return ResizeRule("general", "max_width", cover_width)


def output_path_for(source_file: Path, source_root: Path, output_root: Path) -> Path:
    relative = source_file.relative_to(source_root)
    return output_root / relative.with_suffix(".webp")


def should_skip(source_file: Path, output_file: Path, force: bool) -> bool:
    return not force and output_file.exists() and output_file.stat().st_mtime >= source_file.stat().st_mtime


def target_size(width: int, height: int, rule: ResizeRule) -> tuple[int, int]:
    if width <= 0 or height <= 0:
        return width, height
    if rule.mode == "max_side":
        scale = rule.size / max(width, height)
    elif rule.mode == "width":
        scale = rule.size / width
    elif rule.mode == "max_width":
        scale = min(1.0, rule.size / width)
    else:
        scale = 1.0
    return max(1, round(width * scale)), max(1, round(height * scale))


def normalize_mode(image: Image.Image) -> Image.Image:
    if image.mode in {"RGBA", "RGB"}:
        return image
    if image.mode in {"LA", "PA"} or "transparency" in image.info:
        return image.convert("RGBA")
    return image.convert("RGB")


def convert_one(source_file: Path, source_root: Path, output_root: Path, quality: int, avatar_size: int, cover_width: int, force: bool) -> tuple[str, int, int]:
    output_file = output_path_for(source_file, source_root, output_root)
    if should_skip(source_file, output_file, force):
        return "skipped", 0, output_file.stat().st_size

    relative_path = source_file.relative_to(source_root)
    rule = rule_for(relative_path, avatar_size, cover_width)

    with Image.open(source_file) as original:
        image = ImageOps.exif_transpose(original)
        image = normalize_mode(image)
        width, height = image.size
        next_size = target_size(width, height, rule)
        if next_size != image.size:
            image = image.resize(next_size, Image.Resampling.LANCZOS)

        output_file.parent.mkdir(parents=True, exist_ok=True)
        image.save(output_file, "WEBP", quality=quality, method=6)

    return "converted", source_file.stat().st_size, output_file.stat().st_size


def iter_source_images(source_root: Path) -> list[Path]:
    return sorted(
        path
        for path in source_root.rglob("*")
        if path.is_file() and path.suffix.lower() in RASTER_EXTENSIONS
    )


def main() -> int:
    args = parse_args()
    source_root = args.source.resolve()
    output_root = args.output.resolve()

    if not source_root.is_dir():
        raise SystemExit(f"Source directory not found: {source_root}")
    if args.clean and output_root.exists():
        shutil.rmtree(output_root)

    converted = 0
    skipped = 0
    failed = 0
    input_bytes = 0
    output_bytes = 0

    for source_file in iter_source_images(source_root):
        try:
            status, before, after = convert_one(
                source_file,
                source_root,
                output_root,
                args.quality,
                args.avatar_size,
                args.cover_width,
                args.force,
            )
        except (OSError, UnidentifiedImageError) as error:
            failed += 1
            print(f"failed {source_file.relative_to(source_root)}: {error}")
            continue

        if status == "converted":
            converted += 1
            input_bytes += before
        else:
            skipped += 1
        output_bytes += after

    print(f"converted={converted} skipped={skipped} failed={failed}")
    if converted:
        print(f"converted_input_mb={input_bytes / 1024 / 1024:.2f}")
    print(f"output_total_mb={output_bytes / 1024 / 1024:.2f}")
    print(f"output_dir={output_root}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
