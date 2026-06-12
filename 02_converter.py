#!/usr/bin/env python3
"""
Windows-only Word/HTML → PDF converter with progress bar and cleanup.

Behavior
- Scans the current working directory and all subfolders.
- Targets .doc, .docx, .rtf, .htm, .html.
- Exports each to a same-name .pdf next to the source.
- Deletes the source file only after a successful conversion.
- Skips items whose target PDF already exists.
- Shows a running progress bar.
- --dry-run performs no conversion and no deletion.

Prereqs
- Windows with Microsoft Word installed.
- Python 3.8+ and pywin32:  pip install pywin32

Usage
- Open PowerShell in the target folder, then:
    python convert_to_pdf.py
- Dry run:
    python convert_to_pdf.py --dry-run
"""

import os
import sys
import argparse
from pathlib import Path
from typing import Iterable, Tuple, List

DOC_EXTS = {".doc", ".docx", ".rtf"}
HTML_EXTS = {".htm", ".html"}
ALL_EXTS = DOC_EXTS | HTML_EXTS
TARGET_EXT = ".pdf"

# ---------- discovery ----------

def find_eligible(root: Path) -> List[Path]:
    """
    Return a list of eligible files under root recursively.
    """
    files: List[Path] = []
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in ALL_EXTS:
            files.append(p)
    return files

def target_path(src: Path) -> Path:
    return src.with_suffix(TARGET_EXT)

# ---------- progress bar ----------

def render_progress(done: int, total: int, width: int = 40, prefix: str = "Progress") -> None:
    """
    Simple in-place progress bar: [#####.....] 12/100
    """
    if total <= 0:
        total = 1
    ratio = min(max(done / total, 0.0), 1.0)
    filled = int(ratio * width)
    bar = "#" * filled + "." * (width - filled)
    msg = f"\r{prefix} [{bar}] {done}/{total}"
    sys.stdout.write(msg)
    sys.stdout.flush()
    if done == total:
        sys.stdout.write("\n")

# ---------- conversion via Word COM ----------

def convert_one(word_app, src: Path, dst: Path, dry_run: bool) -> Tuple[bool, str]:
    """
    Convert a single file using Word's ExportAsFixedFormat.
    Returns (success, error_message_if_any).
    """
    # Skip if PDF exists
    if dst.exists():
        return False, "pdf-exists"

    if dry_run:
        return True, ""

    # Open and export
    doc = None
    try:
        # ReadOnly avoids prompts. ConfirmConversions False, AddToRecentFiles False.
        doc = word_app.Documents.Open(str(src), ReadOnly=True)
        # 17 = wdExportFormatPDF; OptimizeFor=0 print quality; CreateBookmarks=1 headings
        doc.ExportAsFixedFormat(
            OutputFileName=str(dst),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1
        )
        doc.Close(False)
        return True, ""
    except Exception as e:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        return False, f"word-error: {e}"

# ---------- main ----------

def main():
    if os.name != "nt":
        print("ERROR: Windows only. Microsoft Word required.", file=sys.stderr)
        sys.exit(1)

    ap = argparse.ArgumentParser(description="Convert DOC/DOCX/RTF/HTML to PDF in current folder recursively, delete sources after success.")
    ap.add_argument("--dry-run", action="store_true", help="List actions and show progress only. No PDFs written. No deletions.")
    args = ap.parse_args()

    root = Path.cwd()
    print(f"Scanning: {root}")

    # Build the full workload first
    workload = find_eligible(root)
    total = len(workload)
    if total == 0:
        print("No DOC/DOCX/RTF/HTML files found.")
        return
    print(f"Found {total} eligible files")

    # Import pywin32 and start Word
    try:
        import win32com.client
    except Exception:
        print("ERROR: pywin32 not available. Install with: pip install pywin32", file=sys.stderr)
        sys.exit(3)

    try:
        word = win32com.client.DispatchEx("Word.Application")
    except Exception:
        print("ERROR: Cannot start Word via COM. Is Microsoft Word installed?", file=sys.stderr)
        sys.exit(4)

    # Quiet Word
    try:
        word.Visible = False
        word.DisplayAlerts = 0
    except Exception:
        pass

    converted = 0
    skipped_pdf_exists = 0
    failed = 0
    failures: List[Tuple[Path, str]] = []

    # Process with progress bar
    done = 0
    render_progress(done, total)

    for src in workload:
        dst = target_path(src)
        ok, info = convert_one(word, src, dst, args.dry_run)
        if ok:
            converted += 1
            # Delete source only if we actually exported a new PDF, not when skipping due to existing PDF, and not in dry-run
            if not args.dry_run:
                try:
                    src.unlink()
                except Exception as e:
                    # Non-fatal: record but continue
                    failures.append((src, f"delete-failed: {e}"))
                    failed += 1
        else:
            if info == "pdf-exists":
                skipped_pdf_exists += 1
            else:
                failed += 1
                failures.append((src, info))

        done += 1
        render_progress(done, total)

    # Cleanup Word
    try:
        word.Quit()
    except Exception:
        pass

    # Summary
    print("Summary")
    print(f"  Converted and deleted sources: {converted}")
    print(f"  Skipped (PDF already exists): {skipped_pdf_exists}")
    print(f"  Failed: {failed}")

    if failures:
        print("Failures:")
        for p, msg in failures:
            print(f"  {p} -> {msg}")

if __name__ == "__main__":
    main()
