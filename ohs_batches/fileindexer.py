# rename_students_dvf_guarded.py
# Purpose:
#   Rename student folders so each folder keeps its 8-digit ID and appends the student's name.
#   Example output name: "12304570 - McNally, Morgan"
#
# Problem this solves:
#   Student names can appear in many messy ways:
#     - In file names like "McNally, Morgan.pdf" or "Morgan McNally.jpg"
#     - Inside PDFs, especially on "Data Verification Forms" that show "Student Name: Last, First"
#   We only want real student names - not labels like "Student Name, Signature" or "Contact Person's Name".
#
# How it works in plain steps:
#   1) Walk each child folder under ROOT that looks like 8 digits.
#   2) Try to find a reliable "Last, First" name from filenames first.
#   3) If not found, read up to the first few pages of any PDFs and look for:
#        - DVF line: "Student Name: Last, First" - special, high trust
#        - Other name-like patterns, but reject surrounding admin text
#   4) Pick the best candidate with a simple score system and rename the folder.
#
# Requirements:
#   pip install PyPDF2
#
# Important toggles:
#   - Set DRY_RUN = True to preview changes without renaming anything.
#   - Set DRY_RUN = False to actually rename folders.

from pathlib import Path
from PyPDF2 import PdfReader
import re

# ---------- configuration ----------

# ROOT is the parent folder that contains many subfolders named with 8 digits.
# Example structure:
#   U:\General Portfolio\No_Match\12304570
#   U:\General Portfolio\No_Match\12306082
ROOT = Path(r"U:\General Portfolio\No_Match")

# DRY_RUN controls whether we just print what would happen (True) or actually rename (False).
DRY_RUN = True

# We only read the first MAX_PDF_PAGES pages of each PDF for speed.
MAX_PDF_PAGES = 3

# A simple check to accept folder names that are exactly 8 digits long.
ID8 = re.compile(r"^\d{8}$")

# ---------- vocabulary and patterns ----------

# STOPWORDS is a list of tokens that should never be part of a real student's name.
# Why? Many PDFs include labels or admin words near the name that can trick the extractor.
# We keep this list aggressive to reduce false positives. It is OK to add more if needed.
STOPWORDS = {
    # generic admin
    "Address","E-mail","Email","Application","Admission","Admissions","Form","Forms","Report",
    "Release","Records","Request","Authorization","Agreement","Information","Data","Verification",
    "Transcript","Testing","Tests","Scores","Interims","Standardized","Conference","Conferences",
    "PSAT","Gifted","Health","Immunization","Immunizations","Record","Plan","Accommodation","Accomodation",
    "Phys","PHY","Eval","Evaluation","Schedule","Status","Honor","Roll","Number","Phone","Comments",
    "Absence","Administering","Interim","Birth","Certificate","Certificates","Faxed","Documents","Sent",
    # DVF lines and headings that must not enter the name
    "Student","Students","Student's","Name","Signature","Grade","Date","Page","Lower","School","Data","Verification","Form",
    # context that created false hits
    "Contact","Person","Persons","Person's","Guardian","Parent","Emergency","Please","Subject",
    # school and common words
    "Field","Trip","Event","Athletic","Service","Hours","Community","Training","Weight",
    "Creative","Writing","Science","Physical","Chinese","Government","History","AP","Pre-ap","Teacher",
    "Grade","Grades","Lower","Middle","Upper","Hall","Prep","Tampa",
    # address tokens
    "Street","St","Avenue","Ave","Road","Rd","Lane","Ln","Court","Ct","Drive","Dr","Boulevard","Blvd",
    "Suite","Ste","Apt","SW","NW","NE","SE",
    # misc frequent noise - short codes that appear in your dataset
    "Mr.","Ms.","Mrs.","Dr.","Drs.","JK","K","US","MS","HS","BC","Im","Imm","St","Rc","Pt"
}

# CONTEXT_REJECT is used to ignore name-like matches that appear next to admin phrases.
# Example we do not want: "Contact Person's Name: John Smith"
CONTEXT_REJECT = re.compile(
    r"(Contact\s+Person'?s\s+Name|Guardian|Parent|Emergency|Student\s+Name\s*,?\s*Signature)"
    r"|(\bAddress\b|\bE-?mail\b|\bComments?\b|\bAbsence\b|\bInterim\b|\bBirth\b|\bCertificate)",
    re.I
)

# NAME_PART describes one piece of a human name:
# - Starts with a capital letter A-Z
# - Followed by letters or common name punctuation like apostrophes or hyphens
NAME_PART = r"[A-Z][a-zA-Z'.\-]+"

# Patterns to find likely names in text:
#   1) "Last, First" - safest
#   2) "First Last" - allowed but lower priority
#   3) "Last; First" - sometimes appears in DVFs
PAT_COMMA = re.compile(rf"\b({NAME_PART}),\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")
PAT_SPACE = re.compile(rf"\b({NAME_PART})\s+({NAME_PART}(?:\s+{NAME_PART})*)\b")
PAT_SEMI  = re.compile(rf"\b({NAME_PART});\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# DVF-specific pattern:
#   We expect "Student Name : Last, First"
#   The look-ahead at the end stops the First name before labels like "Grade:" or "Date:".
PAT_DVF = re.compile(
    rf"""
    Student\s*'?s?\s*Name\s*[:\-]\s*                # label, flexible spaces and punctuation
    (?P<last>{NAME_PART}(?:\s+{NAME_PART})*)        # last name - may contain spaces, hyphens, etc.
    \s*[;,]\s*                                      # separator is a comma or semicolon
    (?P<first>{NAME_PART}(?:\s+{NAME_PART})*)       # first name - may contain middle parts
    (?=\s*(?:$|[,;]|Grade(?:\s*:|\b)|Date(?:\s*:|\b)|Page\b|Lower\s+School|Data\s+Verification|Form\b))
    """,
    re.I | re.X
)

# LABEL_TAIL removes any stray labels that accidentally attach to the captured name.
# Example:
#   "Morgan Grade : 7" -> "Morgan"
LABEL_TAIL = re.compile(r"\s+(Grade|Date|Page|Lower|School|Data|Verification|Form)\b.*$", re.I)

# SHORT_ALLOW lists valid 2-letter surnames or given names that should not be rejected for being short.
SHORT_ALLOW = {"Li","Lu","Xu","Yu","Su","Wu","Ng","Ho","Hu","Ko","Do","He"}

# ---------- helpers ----------

def collapse_ws(s: str) -> str:
    """
    Make whitespace normal so our patterns work better.
    - Replace non-breaking spaces and various dashes with simple forms.
    - Convert runs of spaces, tabs, and newlines to a single space.
    Input: raw text from file names or PDFs
    Output: simplified one-line string
    """
    s = s.replace("\u00A0"," ")  # non-breaking space -> normal space
    # unify different dash characters
    s = s.replace("\u2010","-").replace("\u2011","-").replace("\u2013","-").replace("\u2014","-")
    # collapse multiple spaces to one space
    return re.sub(r"\s+", " ", s)

def norm_cap(s: str) -> str:
    """
    Normalize capitalization for names.
    Example: "mCNAllY, mORGAN" -> "Mcnally, Morgan"
    """
    return " ".join(p[:1].upper() + p[1:].lower() if p else p for p in s.split())

def token_ok(tok: str) -> bool:
    """
    Check if a single word looks like it belongs in a real name.
    Rules:
      - No digits
      - Not in STOPWORDS
      - If very short, must be whitelisted
      - Avoid long fully-upper tokens which are usually labels
    """
    if any(ch.isdigit() for ch in tok):
        return False
    if tok in STOPWORDS:
        return False
    if len(tok) <= 2 and tok not in SHORT_ALLOW:
        return False
    if tok.isupper() and len(tok) > 3:
        return False
    # special case for common email spellings - reject
    t = tok.lower().replace("-","-").replace("–","-")
    if t in {"email","e-mail"}:
        return False
    return True

def looks_like_name(last: str, first: str) -> bool:
    """
    Decide if "last, first" together looks like a real human name.
    We split the text into words and make sure each word passes token_ok.
    """
    toks = (last + " " + first).replace(",", " ").split()
    return all(token_ok(t) for t in toks)

def score_candidate(last: str, first: str, *, source: str, dvf_hit: bool, had_comma: bool, near_anchor: bool) -> int:
    """
    Give a score to a name guess so we can choose the best one later.
    Factors:
      - had_comma: "Last, First" is best
      - near_anchor: close to a helpful label
      - source: file name vs pdf
      - dvf_hit: matched the "Student Name: Last, First" rule - very strong
      - tiny penalty if first name is three or more words - often noise
    """
    score = 10
    if had_comma:
        score += 6
    if near_anchor:
        score += 4
    if source == "filename":
        score += 2
    if source == "pdf":
        score += 1
    if dvf_hit:
        score += 12
    if len(first.split()) >= 3:
        score -= 2
    return score

def strip_labels(s: str) -> str:
    """
    Remove trailing labels that sometimes stick to names in PDFs.
    Example: "Morgan Grade : 7" -> "Morgan"
    """
    return LABEL_TAIL.sub("", s).strip()

# ---------- extraction from file names ----------

def candidates_from_filename(fname: str):
    """
    Look for names in a single file name.
    Strategy:
      - Prefer "Last, First"
      - Allow "First Last" as a backup
      - Clean punctuation and spacing before matching
    Yields tuples: (last, first, score)
    """
    # Path(fname).stem keeps the name without extension, then we normalize separators
    base = collapse_ws(Path(fname).stem.replace("_"," ").replace("-"," "))

    # 1) Strong pattern "Last, First"
    for m in PAT_COMMA.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))

    # 2) Backup pattern "First Last" - lower priority
    for m in PAT_SPACE.finditer(base):
        f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
        last, first = l, f
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=False, near_anchor=False))

# ---------- extraction from PDFs ----------

def pdf_text(pdf: Path, max_pages=MAX_PDF_PAGES) -> str:
    """
    Read up to max_pages of a PDF and return the combined text as one string.
    If the PDF cannot be read, return an empty string.
    """
    try:
        r = PdfReader(str(pdf))
    except Exception:
        return ""
    chunks = []
    for p in r.pages[:max_pages]:
        try:
            chunks.append(p.extract_text() or "")
        except Exception:
            # Some pages cannot be extracted - skip them
            pass
    return collapse_ws(" ".join(chunks))

def candidates_from_pdf(pdf: Path):
    """
    Look for names inside a PDF's text.
    Steps:
      1) Try the DVF-specific rule "Student Name: Last, First" - strongest
      2) Try general patterns, but reject when context smells like admin labels
    Yields tuples: (last, first, score)
    """
    blob = pdf_text(pdf)
    if not blob:
        return []

    out = []

    # 1) DVF-targeted - very strong and guarded against trailing labels like "Grade:"
    for m in PAT_DVF.finditer(blob):
        last = strip_labels(norm_cap(m.group("last")))
        first = strip_labels(norm_cap(m.group("first")))
        if looks_like_name(last, first):
            out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=True, had_comma=True, near_anchor=True)))

    # 2) General fallbacks with context rejection
    for pat, had_comma in ((PAT_COMMA, True), (PAT_SEMI, True), (PAT_SPACE, False)):
        for m in pat.finditer(blob):
            if had_comma:
                last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
            else:
                f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
                last, first = l, f

            # Grab nearby text - if it contains admin phrases like "Contact Person's Name", reject it
            ctx = blob[max(0, m.start()-80): m.end()+80]
            if CONTEXT_REJECT.search(ctx):
                continue

            if looks_like_name(last, first):
                out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=False, had_comma=had_comma, near_anchor=False)))

    return out

# ---------- candidate selection ----------

def pick_best(cands):
    """
    Pick the best name candidate among many.
    Rules:
      - Combine duplicate names by keeping the highest score
      - Sort by score descending
      - If scores tie, prefer the longer full name length
    Returns: (last, first) or None if no candidates
    """
    if not cands:
        return None

    best = {}
    for last, first, sc in cands:
        key = (last, first)
        best[key] = max(sc, best.get(key, -10**9))

    return sorted(best.items(), key=lambda kv: (-kv[1], -(len(kv[0][0]) + len(kv[0][1]))))[0][0]

# ---------- renaming helpers ----------

def target_name(id8: str, last_first):
    """
    Build the final folder name using the 8-digit ID and the chosen "Last, First".
    """
    last, first = last_first
    return f"{id8} - {last}, {first}"

def safe_rename(folder: Path, new_base: str):
    """
    Safely rename a folder to new_base.
    - If DRY_RUN is True, only print the plan.
    - If a target with the same name exists, append " (1)", " (2)", etc.
    """
    target = folder.with_name(new_base)

    # If the name is unchanged, there is nothing to do
    if target.name == folder.name:
        print(f"[SAME] {folder.name}")
        return

    # Avoid overwriting an existing folder name - add a counter suffix
    if target.exists():
        i = 1
        while folder.with_name(f"{new_base} ({i})").exists():
            i += 1
        target = folder.with_name(f"{new_base} ({i})")

    if DRY_RUN:
        print(f"[DRY]  {folder.name} -> {target.name}")
    else:
        folder.rename(target)
        print(f"[OK]   {folder.name} -> {target.name}")

# ---------- main workflow ----------

def derive_name_for_folder(folder: Path):
    """
    Try to find the student's name for a folder by:
      1) Checking file names in the folder
      2) If not found, scanning PDFs inside the folder
    Returns: (last, first) or None
    """
    cands = []

    # Step 1 - mine file names
    for f in folder.iterdir():
        if f.is_file():
            cands.extend(candidates_from_filename(f.name))

    # Step 2 - mine PDFs if file names did not help
    if not cands:
        for f in folder.iterdir():
            if f.suffix.lower() == ".pdf":
                cands.extend(candidates_from_pdf(f))
                if cands:
                    break  # stop after first PDF that yields candidates

    return pick_best(cands)

def main():
    """
    Walk the ROOT directory and rename each 8-digit folder when a reliable name is found.
    Assumptions:
      - Each student folder is named with exactly 8 digits at the start.
      - A correct match should be in a file name or inside a PDF within the folder.
      - The DVF pattern is authoritative when present.
    """
    for child in ROOT.iterdir():
        if not child.is_dir():
            continue
        if not ID8.fullmatch(child.name):
            continue

        best = derive_name_for_folder(child)

        if best:
            safe_rename(child, target_name(child.name, best))
        else:
            print(f"[SKIP] {child.name}  no reliable name")

if __name__ == "__main__":
    main()
