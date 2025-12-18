#!/usr/bin/env python3
"""
Convert all .docx Word documents in a folder (recursively) to Markdown.

Requires:
  - pandoc installed and on PATH
  - pypandoc (pip install pypandoc)

Usage:
  python docx_to_md.py /path/to/input_folder /path/to/output_folder
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path

import pypandoc


def ensure_pandoc_available() -> None:
    if shutil.which("pandoc") is None:
        raise RuntimeError(
            "pandoc not found on PATH. Install pandoc and ensure it's available as 'pandoc'."
        )


def convert_docx_to_md(
    docx_path: Path,
    out_dir: Path,
    media_dir_name: str = "media",
) -> Path:
    """
    Converts one .docx to Markdown and extracts embedded media to:
      out_dir/<stem>_<media_dir_name>/

    Returns the output .md path.
    """
    rel_stem = docx_path.stem
    out_md = out_dir / f"{rel_stem}.md"
    out_media = out_dir / f"{rel_stem}_{media_dir_name}"

    out_dir.mkdir(parents=True, exist_ok=True)
    out_media.mkdir(parents=True, exist_ok=True)

    extra_args = [
        "--wrap=none",
        "--extract-media", str(out_media),
    ]

    md_text = pypandoc.convert_file(
        str(docx_path),
        to="gfm",              # GitHub-flavored Markdown
        format="docx",
        extra_args=extra_args,
    )

    out_md.write_text(md_text, encoding="utf-8")
    return out_md


def iter_docx_files(root: Path) -> list[Path]:
    # Exclude temporary Word lock files like "~$foo.docx"
    return sorted(
        p for p in root.rglob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_folder", type=Path)
    parser.add_argument("output_folder", type=Path)
    parser.add_argument(
        "--recursive",
        action="store_true",
        default=True,
        help="Included for compatibility; script always recurses.",
    )
    parser.add_argument(
        "--preserve-subdirs",
        action="store_true",
        help="Mirror input subfolder structure under output.",
    )
    parser.add_argument(
        "--media-dir-name",
        default="media",
        help="Suffix for extracted media directories (default: media).",
    )
    args = parser.parse_args()

    ensure_pandoc_available()

    in_root: Path = args.input_folder.resolve()
    out_root: Path = args.output_folder.resolve()

    if not in_root.exists() or not in_root.is_dir():
        print(f"Input folder not found or not a directory: {in_root}", file=sys.stderr)
        return 2

    docx_files = iter_docx_files(in_root)
    if not docx_files:
        print(f"No .docx files found under: {in_root}", file=sys.stderr)
        return 0

    converted = 0
    failed: list[tuple[Path, str]] = []

    for docx in docx_files:
        try:
            if args.preserve_subdirs:
                rel_parent = docx.parent.relative_to(in_root)
                out_dir = out_root / rel_parent
            else:
                out_dir = out_root

            out_md = convert_docx_to_md(docx, out_dir, media_dir_name=args.media_dir_name)
            converted += 1
            print(f"OK  {docx} -> {out_md}")
        except Exception as e:
            failed.append((docx, str(e)))
            print(f"ERR {docx}: {e}", file=sys.stderr)

    print(f"\nConverted: {converted}/{len(docx_files)}")
    if failed:
        print("Failures:", file=sys.stderr)
        for path, err in failed:
            print(f"  - {path}: {err}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
