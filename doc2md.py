#!/usr/bin/env python3
"""
doc2md.py - Convert all Word documents in the script's folder (and subfolders) to Markdown.

Default behavior (no arguments):
- Recurses from the folder where this script file lives.
- Converts every .doc and .docx it finds.
- Writes Markdown into: <script_folder>/markdown_out/
- Extracts embedded media into per-file folders inside markdown_out.

Optional overrides:
- --input  : choose a different starting folder to scan
- --output : choose a different output folder
- --no-mirror : do NOT mirror subfolders (flatten into output folder)
- --media-suffix : change the extracted media folder suffix

Requirements:
- pandoc installed and available as "pandoc"
- pip install pypandoc
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pypandoc


def ensure_pandoc_available() -> None:
    """Stop early if Pandoc isn't installed or not on PATH."""
    if shutil.which("pandoc") is None:
        raise RuntimeError(
            "pandoc not found on PATH. Install pandoc and ensure the 'pandoc' command works."
        )


def iter_word_files(root_folder: Path) -> list[Path]:
    """
    Recursively collect .doc and .docx files under root_folder.

    Skips Word's temporary lock files that start with '~$'.
    """
    files: list[Path] = []
    for ext in ("*.docx", "*.doc"):
        for p in root_folder.rglob(ext):
            if p.is_file() and not p.name.startswith("~$"):
                files.append(p)
    return sorted(files)


def convert_word_to_markdown(
    word_path: Path,
    input_root: Path,
    output_root: Path,
    mirror_subdirs: bool,
    media_suffix: str,
) -> Path:
    """
    Convert one Word file to Markdown.

    Folder rules:
    - If mirror_subdirs=True, recreate the input folder structure under output_root.
      Example:
        input_root/notes/a.docx -> output_root/notes/a.md
    - If mirror_subdirs=False, put everything directly in output_root (flatten).
    """
    # Decide where this file's Markdown should go.
    if mirror_subdirs:
        # Keep the same subfolder layout under the output folder.
        rel_parent = word_path.parent.relative_to(input_root)
        out_dir = output_root / rel_parent
    else:
        # Put every Markdown file directly into the output folder.
        out_dir = output_root

    out_dir.mkdir(parents=True, exist_ok=True)

    # Markdown file path: same base name, but .md extension.
    out_md = out_dir / f"{word_path.stem}.md"

    # Media extraction folder: "<stem>_<suffix>" next to the .md file.
    out_media = out_dir / f"{word_path.stem}_{media_suffix}"
    out_media.mkdir(parents=True, exist_ok=True)

    # Input format for Pandoc depends on file extension.
    # Note: .doc support can vary by system; if it fails, convert to .docx first.
    suffix = word_path.suffix.lower()
    input_format = "docx" if suffix == ".docx" else "doc"

    extra_args = [
        "--wrap=none",
        "--extract-media",
        str(out_media),
    ]

    md_text: str = pypandoc.convert_file(
        str(word_path),
        to="gfm",
        format=input_format,
        extra_args=extra_args,
    )

    out_md.write_text(md_text, encoding="utf-8")
    return out_md


def build_arg_parser() -> argparse.ArgumentParser:
    """
    Build command line options.

    Key point:
    - No required positional arguments.
    - Running 'python3 doc2md.py' should work with defaults.
    """
    parser = argparse.ArgumentParser(
        prog="doc2md.py",
        description="Recursively convert .doc/.docx files to Markdown.",
    )

    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Folder to scan. Default: the folder where doc2md.py is located.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Folder to write Markdown into. Default: <script_folder>/markdown_out",
    )
    parser.add_argument(
        "--no-mirror",
        action="store_true",
        help="Do not mirror subfolders. Flatten all .md outputs into the output folder.",
    )
    parser.add_argument(
        "--media-suffix",
        default="media",
        help="Suffix used for extracted media folders (default: media).",
    )

    return parser


def main() -> int:
    """
    Orchestrates the whole run.

    Process (in order):
    1) Decide input folder (default: where the script lives).
    2) Decide output folder (default: <script_folder>/markdown_out).
    3) Find Word files.
    4) Convert them one by one.
    5) Print a summary and return an exit code.
    """
    parser = build_arg_parser()
    args = parser.parse_args()

    # Where is this script file located?
    script_path = Path(__file__).resolve()
    script_folder = script_path.parent

    # Default input: the script's folder (NOT the current working directory).
    input_root = (args.input or script_folder).resolve()

    # Default output: a "markdown_out" folder next to the script.
    output_root = (args.output or (script_folder / "markdown_out")).resolve()

    # Mirror subfolders unless user disables it.
    mirror_subdirs = not args.no_mirror
    media_suffix = args.media_suffix

    ensure_pandoc_available()

    if not input_root.exists() or not input_root.is_dir():
        print(f"Input folder not found or not a directory: {input_root}", file=sys.stderr)
        return 2

    word_files = iter_word_files(input_root)
    if not word_files:
        print(f"No .doc or .docx files found under: {input_root}")
        return 0

    output_root.mkdir(parents=True, exist_ok=True)

    converted = 0
    failures: list[tuple[Path, str]] = []

    for word_path in word_files:
        try:
            out_md = convert_word_to_markdown(
                word_path=word_path,
                input_root=input_root,
                output_root=output_root,
                mirror_subdirs=mirror_subdirs,
                media_suffix=media_suffix,
            )
            converted += 1

            # Print nice relative paths when possible (easier to read).
            try:
                in_rel = word_path.relative_to(input_root)
            except Exception:
                in_rel = word_path

            try:
                out_rel = out_md.relative_to(script_folder)
            except Exception:
                out_rel = out_md

            print(f"OK  {in_rel} -> {out_rel}")

        except Exception as e:
            failures.append((word_path, str(e)))
            try:
                in_rel = word_path.relative_to(input_root)
            except Exception:
                in_rel = word_path
            print(f"ERR {in_rel}: {e}", file=sys.stderr)

    print(f"\nConverted: {converted}/{len(word_files)}")
    print(f"Scanned:    {input_root}")
    print(f"Output:     {output_root}")

    if failures:
        print("\nFailures:", file=sys.stderr)
        for path, err in failures:
            try:
                in_rel = path.relative_to(input_root)
            except Exception:
                in_rel = path
            print(f"  - {in_rel}: {err}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
