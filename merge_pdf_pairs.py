"""
merge_pdf_pairs.py

Drop this script into a folder of PDFs and run it. It pairs the PDFs by the
trailing digit run in their filenames (...1.pdf + ...2.pdf, ...3.pdf + ...4.pdf,
...) and writes each merged pair to a "merged" subfolder, named after the
person identified inside the pair.

The trailing index is any run of digits sitting immediately before the .pdf
extension, with or without a separator before it. All of these match:

    Letters1.pdf, Letters_1.pdf, Letters 1.pdf, Letters-99.pdf

Naming preference within each pair:
    1. "Name: First Last"  (employment-details letter)
    2. "Dear First"        (cover letter; first name only)
    3. Stem of the first PDF in the pair (last-resort fallback)

Run from PowerShell or cmd, with the script sitting in the target folder:
    pip install pypdf
    python merge_pdf_pairs.py
"""

import re
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    sys.stderr.write(
        "Missing dependency 'pypdf'. Install with:  pip install pypdf\n"
    )
    sys.exit(1)


SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_SUBFOLDER = "merged"

NAME_FROM_DETAILS = re.compile(
    r"Name\s*:\s*([A-Za-z][A-Za-z\-'\.]+(?:\s+[A-Za-z][A-Za-z\-'\.]+)+)"
)
NAME_FROM_LETTER = re.compile(r"Dear\s+([A-Za-z][A-Za-z\-']+)")
TRAILING_INT = re.compile(r"(\d+)\.pdf$", re.IGNORECASE)
INVALID_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def extract_text(pdf_path: Path) -> str:
    try:
        reader = PdfReader(str(pdf_path))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        sys.stderr.write(f"  ! could not read {pdf_path.name}: {exc}\n")
        return ""


def derive_name(pair: list) -> str:
    full_name = None
    first_name = None
    for pdf_path in pair:
        text = extract_text(pdf_path)
        if not full_name:
            match = NAME_FROM_DETAILS.search(text)
            if match:
                full_name = " ".join(match.group(1).split())
        if not first_name:
            match = NAME_FROM_LETTER.search(text)
            if match:
                first_name = match.group(1).strip()
        if full_name:
            break
    return full_name or first_name or pair[0].stem


def safe_filename(name: str) -> str:
    cleaned = INVALID_FS_CHARS.sub("", name).strip(" .")
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "unnamed"


def numbered_pdfs(folder: Path) -> list:
    items = []
    for pdf_path in folder.glob("*.pdf"):
        match = TRAILING_INT.search(pdf_path.name)
        if match:
            items.append((int(match.group(1)), pdf_path))
    items.sort(key=lambda entry: entry[0])
    return [path for _, path in items]


def merge_pair(pair: list, out_path: Path) -> None:
    writer = PdfWriter()
    for pdf_path in pair:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            writer.add_page(page)
    with open(out_path, "wb") as handle:
        writer.write(handle)


def unique_path(out_folder: Path, base: str) -> Path:
    candidate = out_folder / f"{base}.pdf"
    counter = 2
    while candidate.exists():
        candidate = out_folder / f"{base}_{counter}.pdf"
        counter += 1
    return candidate


def main() -> int:
    in_folder = SCRIPT_DIR
    out_folder = in_folder / OUTPUT_SUBFOLDER
    out_folder.mkdir(parents=True, exist_ok=True)

    pdfs = numbered_pdfs(in_folder)
    if not pdfs:
        sys.stderr.write(
            f"No PDFs ending in a digit run found in {in_folder}\n"
        )
        return 1
    if len(pdfs) % 2 != 0:
        sys.stderr.write(
            f"WARNING: odd PDF count ({len(pdfs)}). Last unpaired file skipped: "
            f"{pdfs[-1].name}\n"
        )

    paired_count = len(pdfs) - (len(pdfs) % 2)
    pairs = [pdfs[i:i + 2] for i in range(0, paired_count, 2)]

    print(f"Input:  {in_folder}")
    print(f"Output: {out_folder}")
    print(f"Pairs:  {len(pairs)}")
    print()

    for index, pair in enumerate(pairs, start=1):
        name = derive_name(pair)
        out_path = unique_path(out_folder, safe_filename(name))
        merge_pair(pair, out_path)
        print(
            f"[{index:>3}] {pair[0].name}  +  {pair[1].name}  ->  {out_path.name}"
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
