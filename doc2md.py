#!/usr/bin/env python3
"""
doc2md.py - Convert all Word documents in the script's folder (and subfolders) to Markdown.

Default behavior (no arguments):
- Recurses from the folder where this script file lives (NOT the terminalI mean: not the current working directory).
- Converts every .doc and .docx it finds.
- Writes Markdown into: <script_folder>/markdown_out/
- Extracts embedded media into per-file folders inside markdown_out.

Optional overrides:
- --input        : choose a different starting folder to scan
- --output       : choose a different output folder
- --no-mirror    : do NOT mirror subfolders (flatten into output folder)
- --media-suffix : change the extracted media folder suffix

Requirements:
- pandoc executable installed and runnable (the real program, not just a Python package)
- pip install pypandoc
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

import pypandoc


# ----------------------------
# Step 0: Prove pandoc is usable
# ----------------------------
def ensure_pandoc_available() -> str:
    """
    Robust pandoc check.

    Why this exists:
    - A PATH-based check like shutil.which("pandoc") is fragile.
    - Your shell PATH and Python's PATH can differ.
    - Also: `pip install pandoc` installs a Python package, not necessarily the pandoc executable.

    What we do (in order):
    1) Try to discover pandoc using a login shell: `command -v pandoc`
       - This mirrors how your terminal normally finds executables.
    2) If that fails, try to run `pandoc --version` directly anyway.
       - Sometimes discovery fails but execution still works.
    3) Once we have a pandoc path (or "pandoc"), ensure pypandoc is pointed at it.
       - If pypandoc cannot locate pandoc itself, we set it explicitly.

    Returns:
      A string representing the pandoc executable path (best effort).

    Raises:
      RuntimeError if pandoc is not runnable.
    """

    def can_run(cmd: list[str]) -> bool:
        """Return True if running cmd works (exit code 0)."""
        try:
            subprocess.run(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
            )
            return True
        except Exception:
            return False

    def resolve_pandoc_from_shell() -> str | None:
        """
        Ask a shell to locate pandoc.

        We use:
          sh -lc "command -v pandoc"

        - `-l` means "login shell" (loads the user's normal environment)
        - `-c` means "run this command"
        """
        try:
            r = subprocess.run(
                ["sh", "-lc", "command -v pandoc"],
                capture_output=True,
                text=True,
                check=True,
            )
            p = r.stdout.strip()
            return p if p else None
        except Exception:
            return None

    # Try to get an absolute path like /usr/bin/pandoc
    pandoc_path = resolve_pandoc_from_shell()

    # If shell lookup failed, try executing pandoc anyway.
    if pandoc_path is None and can_run(["pandoc", "--version"]):
        pandoc_path = "pandoc"

    # If we still have nothing, pandoc is not runnable from this process.
    if pandoc_path is None:
        raise RuntimeError(
            "pandoc is not runnable from this Python process.\n"
            "Fix options:\n"
            "  - Install pandoc as an executable (apt/brew/dnf/etc)\n"
            "  - Ensure its directory is on PATH for the user running python\n"
            f"Current PATH seen by Python:\n  {os.environ.get('PATH','')}\n"
        )

    # Make sure pypandoc is using a pandoc it can actually find.
    # If pypandoc can't find one, we set it explicitly.
    try:
        seen = pypandoc.get_pandoc_path()
    except OSError:
        seen = ""

    if not seen:
        # If we only have "pandoc", try to make it absolute for pypandoc.
        abs_path = pandoc_path
        if pandoc_path == "pandoc":
            abs_found = resolve_pandoc_from_shell()
            abs_path = abs_found or "pandoc"

        pypandoc.set_pandoc_path(abs_path)

        # Verify after setting; if this fails it will raise, which is correct.
        _ = pypandoc.get_pandoc_path()

    return str(pandoc_path)


# ----------------------------
# Step 1: Find Word files (recursive)
# ----------------------------
def iter_word_files(root_folder: Path) -> list[Path]:
    """
    Recursively collect .doc and .docx files under root_folder.

    Skips:
    - Word temporary lock files that start with "~$" (not real documents)

    Returns:
    - A sorted list of Paths to Word documents.
    """
    files: list[Path] = []

    # We search twice: once for .docx and once for .doc
    for ext in ("*.docx", "*.doc"):
        for p in root_folder.rglob(ext):
            if p.is_file() and not p.name.startswith("~$"):
                files.append(p)

    # Sorted output makes logs stable and predictable.
    return sorted(files)


# ----------------------------
# Step 2: Convert ONE Word file to Markdown
# ----------------------------
def convert_word_to_markdown(
    word_path: Path,
    input_root: Path,
    output_root: Path,
    mirror_subdirs: bool,
    media_suffix: str,
) -> Path:
    """
    Convert a single Word document to Markdown.

    Inputs:
    - word_path: the file to convert
    - input_root: the folder we are scanning
    - output_root: where Markdown output should go
    - mirror_subdirs: whether to preserve subfolders in output
    - media_suffix: suffix for extracted media folders (default "media")

    Output behavior:
    - If mirror_subdirs=True:
        input_root/notes/a.docx -> output_root/notes/a.md
      plus:
        output_root/notes/a_media/  (images, etc.)
    - If mirror_subdirs=False:
        input_root/notes/a.docx -> output_root/a.md
      plus:
        output_root/a_media/
    """

    # Decide which folder the output Markdown should go into.
    if mirror_subdirs:
        # Keep the same folder structure (relative to input_root).
        rel_parent = word_path.parent.relative_to(input_root)
        out_dir = output_root / rel_parent
    else:
        # Flatten: everything goes directly into output_root.
        out_dir = output_root

    # Ensure the folder exists.
    out_dir.mkdir(parents=True, exist_ok=True)

    # Markdown file name: same base name, but .md extension.
    out_md = out_dir / f"{word_path.stem}.md"

    # Media folder name: "<stem>_<media_suffix>"
    out_media = out_dir / f"{word_path.stem}_{media_suffix}"
    out_media.mkdir(parents=True, exist_ok=True)

    # Decide the input format for pandoc based on the file extension.
    # Note: .doc support varies by system. If .doc fails, convert it to .docx first.
    suffix = word_path.suffix.lower()
    input_format = "docx" if suffix == ".docx" else "doc"

    # Extra args to control pandoc output.
    extra_args = [
        "--wrap=none",            # Keep lines as-is, don't force wrapping.
        "--extract-media", str(out_media),  # Pull images/media into this folder.
    ]

    # Convert the file.
    md_text: str = pypandoc.convert_file(
        str(word_path),
        to="gfm",                 # GitHub-Flavored Markdown.
        format=input_format,      # "docx" or "doc"
        extra_args=extra_args,
    )

    # Write the Markdown text to disk.
    out_md.write_text(md_text, encoding="utf-8")

    return out_md


# ----------------------------
# Step 3: Command line arguments (optional overrides)
# ----------------------------
def build_arg_parser() -> argparse.ArgumentParser:
    """
    Make command line options.

    Critical design choice:
    - There are NO required positional arguments.
    - Running: `python3 doc2md.py` must work with defaults.
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


# ----------------------------
# Step 4: Main program
# ----------------------------
def main() -> int:
    """
    The program runner.

    Process:
    1) Determine the script folder.
    2) Decide input folder (default: script folder).
    3) Decide output folder (default: script_folder/markdown_out).
    4) Confirm pandoc works.
    5) Find Word files.
    6) Convert each file.
    7) Print a summary and exit status.
    """
    parser = build_arg_parser()
    args = parser.parse_args()

    # Where is this script file stored?
    script_path = Path(__file__).resolve()
    script_folder = script_path.parent

    # Default input: the folder the script is in.
    input_root = (args.input or script_folder).resolve()

    # Default output: markdown_out next to the script.
    output_root = (args.output or (script_folder / "markdown_out")).resolve()

    # Mirror subfolders unless user disables it.
    mirror_subdirs = not args.no_mirror
    media_suffix = args.media_suffix

    # Validate pandoc early so we fail before doing work.
    _pandoc_path = ensure_pandoc_available()

    # Sanity check input folder.
    if not input_root.exists() or not input_root.is_dir():
        print(f"Input folder not found or not a directory: {input_root}", file=sys.stderr)
        return 2

    # Gather Word files.
    word_files = iter_word_files(input_root)

    # No files means nothing to do.
    if not word_files:
        print(f"No .doc or .docx files found under: {input_root}")
        return 0

    # Make sure output folder exists.
    output_root.mkdir(parents=True, exist_ok=True)

    # Track results.
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

            # Print user-friendly paths.
            in_rel = word_path.relative_to(input_root)
            try:
                out_rel = out_md.relative_to(script_folder)
            except Exception:
                out_rel = out_md

            print(f"OK  {in_rel} -> {out_rel}")

        except Exception as e:
            failures.append((word_path, str(e)))
            in_rel = word_path.relative_to(input_root)
            print(f"ERR {in_rel}: {e}", file=sys.stderr)

    # Print summary.
    print(f"\nConverted: {converted}/{len(word_files)}")
    print(f"Scanned:    {input_root}")
    print(f"Output:     {output_root}")

    if failures:
        print("\nFailures:", file=sys.stderr)
        for path, err in failures:
            print(f"  - {path.relative_to(input_root)}: {err}", file=sys.stderr)
        return 1

    return 0


# Standard Python entry point:
# - If you run "python3 doc2md.py", __name__ == "__main__" and it runs main().
# - If you import this file from another script, main() will not auto-run.
if __name__ == "__main__":
    raise SystemExit(main())
