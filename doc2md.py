#!/usr/bin/env python3
"""
DOC/DOCX -> Markdown Converter (recursive, run-from-here)

What this script does:
- Looks in the folder where THIS script is sitting (and all subfolders).
- Finds every Word file that ends with .doc or .docx.
- Converts each one into a .md (Markdown) file.
- Saves results into a new folder next to the script: ./markdown_out/
- Also pulls out pictures from the Word files into folders inside ./markdown_out/

Important notes (simple):
- Pandoc does the actual converting.
- .doc files (older Word format) might fail unless Pandoc can read them on your machine.
  If .doc fails, you can convert those files to .docx in Word first.

Requirements:
- Install pandoc and make sure the command "pandoc" works in Terminal / Command Prompt.
- pip install pypandoc

Run:
- python doc_to_md.py
"""

from __future__ import annotations

# These are built-in Python tools (they come with Python).
import shutil          # Helps us check if a program exists on the computer (pandoc).
import sys             # Lets us exit with an error code if something goes wrong.
from pathlib import Path  # A clean way to work with files and folders (paths).

# This is the extra library we install with pip.
# It lets Python call Pandoc without us building command strings by hand.
import pypandoc


# ----------------------------
# Step 1: Make sure Pandoc exists
# ----------------------------
def ensure_pandoc_available() -> None:
    """
    Pandoc is a separate program (not part of Python).
    This function checks: "Can the computer find a program named 'pandoc'?"

    If it cannot, we stop the script and show a clear message.
    """
    if shutil.which("pandoc") is None:
        raise RuntimeError(
            "pandoc not found. Install pandoc and make sure 'pandoc' works in your terminal."
        )


# ----------------------------
# Step 2: Find Word files (recursive)
# ----------------------------
def iter_word_files(root_folder: Path) -> list[Path]:
    """
    This searches through the folder and every subfolder.

    It returns a list of paths that end with:
    - .docx (newer Word files)
    - .doc  (older Word files)

    We also skip weird temporary Word files that start with "~$"
    because those are not real documents (Word creates them while editing).
    """
    word_files: list[Path] = []

    # Look for .docx files anywhere under root_folder.
    for p in root_folder.rglob("*.docx"):
        if p.is_file() and not p.name.startswith("~$"):
            word_files.append(p)

    # Look for .doc files anywhere under root_folder.
    for p in root_folder.rglob("*.doc"):
        if p.is_file() and not p.name.startswith("~$"):
            word_files.append(p)

    # Sort the list so the output is predictable (same order each run).
    return sorted(word_files)


# ----------------------------
# Step 3: Convert ONE Word file to Markdown
# ----------------------------
def convert_word_to_markdown(
    word_path: Path,
    output_root: Path,
    script_root: Path,
    media_suffix: str = "media",
) -> Path:
    """
    Convert a single Word document into Markdown.

    Inputs:
    - word_path: the Word file we are converting
    - output_root: the folder where we put Markdown results (./markdown_out)
    - script_root: the folder where this script is located (the "starting point")
    - media_suffix: a label used for the folder where pictures are extracted

    What we do:
    1) We recreate the same subfolder structure inside markdown_out.
       Example:
         If the Word file is:  ./notes/meetings/week1.docx
         Then the Markdown goes: ./markdown_out/notes/meetings/week1.md

    2) We also extract images to a folder next to the .md file:
         ./markdown_out/notes/meetings/week1_media/
    """
    # Figure out the Word file’s location *relative* to the script folder.
    # Example:
    #   script_root = /home/me/project
    #   word_path   = /home/me/project/notes/a.docx
    #   relative    = notes/a.docx
    relative_path = word_path.relative_to(script_root)

    # Create the output folder path that matches the input structure.
    # Example:
    #   output_root = /home/me/project/markdown_out
    #   relative_path.parent = notes
    #   output_dir = /home/me/project/markdown_out/notes
    output_dir = output_root / relative_path.parent

    # Make sure that folder exists.
    output_dir.mkdir(parents=True, exist_ok=True)

    # The Markdown filename should match the Word filename but end in .md
    # Example: a.docx -> a.md
    output_md_path = output_dir / f"{word_path.stem}.md"

    # Create a folder for extracted images and other embedded files.
    # Example: a_media
    media_dir = output_dir / f"{word_path.stem}_{media_suffix}"
    media_dir.mkdir(parents=True, exist_ok=True)

    # These are extra instructions we pass to Pandoc.
    extra_args = [
        "--wrap=none",                 # Do not hard-wrap lines at a certain width.
        "--extract-media", str(media_dir),  # Save images and media to this folder.
    ]

    # Pandoc needs to know what format we want:
    # - "gfm" means GitHub-Flavored Markdown.
    # - The input format is based on file type.
    #
    # We set input format explicitly:
    # - If it ends with .docx, use "docx"
    # - If it ends with .doc, use "doc" (may or may not work depending on your setup)
    if word_path.suffix.lower() == ".docx":
        input_format = "docx"
    else:
        input_format = "doc"  # Warning: older .doc files can fail on some systems.

    # This is the main conversion:
    # - convert_file reads the Word file
    # - Pandoc converts it
    # - We get back Markdown text as a Python string
    md_text: str = pypandoc.convert_file(
        str(word_path),
        to="gfm",
        format=input_format,
        extra_args=extra_args,
    )

    # Write that Markdown text into the .md file.
    output_md_path.write_text(md_text, encoding="utf-8")

    # Return the path to the new Markdown file (useful for printing/logging).
    return output_md_path


# ----------------------------
# Step 4: Main program (the part that runs)
# ----------------------------
def main() -> int:
    """
    This is the "controller" function.

    Big picture:
    - Decide what folder to scan (the script's folder).
    - Create an output folder.
    - Find Word files.
    - Convert each one.
    - Print results.
    - Return an exit code:
        0 = success
        1 = some files failed
        2 = setup problem (like missing folder)
    """
    # Find the folder where this script is located.
    # Path(__file__) is the path to this script file.
    # .resolve() makes it an absolute path (no ".." parts).
    script_path = Path(__file__).resolve()
    script_root = script_path.parent

    # Decide where outputs will go.
    output_root = script_root / "markdown_out"

    # Make sure Pandoc exists before doing anything else.
    ensure_pandoc_available()

    # Find all Word files under the script folder.
    word_files = iter_word_files(script_root)

    # If there are none, we just tell the user and exit normally.
    if not word_files:
        print(f"No .doc or .docx files found under: {script_root}")
        return 0

    # Make sure the main output folder exists.
    output_root.mkdir(parents=True, exist_ok=True)

    # Keep track of how many succeed and which ones fail.
    converted_count = 0
    failures: list[tuple[Path, str]] = []

    # Convert each Word file.
    for word_path in word_files:
        try:
            out_md = convert_word_to_markdown(
                word_path=word_path,
                output_root=output_root,
                script_root=script_root,
            )
            converted_count += 1
            print(f"OK  {word_path.relative_to(script_root)} -> {out_md.relative_to(script_root)}")
        except Exception as e:
            # If something goes wrong for this file, we record the error and keep going.
            failures.append((word_path, str(e)))
            print(f"ERR {word_path.relative_to(script_root)}: {e}", file=sys.stderr)

    # Summary at the end.
    total = len(word_files)
    print(f"\nConverted: {converted_count}/{total}")
    print(f"Output folder: {output_root}")

    # If any failed, print a list and return error code 1.
    if failures:
        print("\nFailures:", file=sys.stderr)
        for path, err in failures:
            print(f"  - {path.relative_to(script_root)}: {err}", file=sys.stderr)
        return 1

    # All good.
    return 0


# This is the standard Python way to say:
# "Only run main() if this file is being executed directly,
#  not if it is being imported by another Python file."
if __name__ == "__main__":
    raise SystemExit(main())
