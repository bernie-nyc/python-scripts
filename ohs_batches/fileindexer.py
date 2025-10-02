"""
fileindexer.py  —  FULLY COMMENTED VERSION (friendly for middle school CS students)

WHAT THIS PROGRAM DOES (in plain words):
---------------------------------------
You have many folders, each named with an 8-digit student ID like "12305932".
Inside each folder are files (PDFs and others). We want to rename the folder
so it becomes: "12305932 - Last, First" (student's last and first name).

HOW WE FIND THE NAME:
1) We FIRST look at the NAMES of ALL the files in the folder.
   - If any filename already contains a name like "Wynn, Hannah" or "Hannah Wynn",
     we use that. This is fast and usually correct.

2) If filenames are not helpful, we look INSIDE ALL the PDF files in that folder.
   - We read a little bit of text from the first few pages.
   - We try to find patterns like:
       • "Student Name: Last, First" (this special rule is for Data Verification Forms)
       • "Last, First"
       • "First Last" (which we turn into "Last, First")
   - We avoid getting tricked by generic words like "All Divisions, DVF".

SAFETY:
- We use a big list of “STOPWORDS” so non-name words don’t fool us.
- We give higher points to very name-looking patterns like "Last, First".
- We do a DRY RUN by default (it prints what it WOULD rename, but doesn’t
  actually change anything). You can turn DRY_RUN to False to really rename.

HOW TO RUN:
- Put this file somewhere you can run it (like alongside your folders).
- Make sure the ROOT path below points to the top folder with the 8-digit folders.
- Run:  py .\fileindexer.py  (in PowerShell)

NOTE:
- This script prints lines like:
    [DRY]  12305932 -> 12305932 - Wynn, Hannah
  That shows “old name -> new name”. If DRY_RUN is False, it will actually rename.
"""

from pathlib import Path               # Path lets us work with folders/files easily
import re                               # re is for Regular Expressions (pattern matching)
from PyPDF2 import PdfReader            # PdfReader lets us read text from PDF files


# ---------------------------- Settings ----------------------------

# ROOT is the top folder that contains many 8-digit student folders.
ROOT = Path(r"U:\General Portfolio\No_Match")

# DRY_RUN = True means "don't actually rename"; just print what we would do.
# Set to False to actually rename folders.
DRY_RUN = True

# When reading a PDF, how many pages do we scan for text?
# We don’t need to read the whole file to find a name, and this saves time.
MAX_PDF_PAGES = 3

# This pattern matches exactly 8 digits (like "12345678").
# We use it to decide which subfolders are student folders.
ID8 = re.compile(r"^\d{8}$")


# ------------------ Noise words (single tokens) -------------------
# STOPWORDS is a set (like a list, but faster for checking) of words that
# should NOT appear in a real name. If we see these words, we’ll think
# “this is probably not a student name.”
STOPWORDS = {
    # admin/common labels
    "Address","Application","Admissions","Admission","Form","Forms","Report","Release","Records",
    "Request","Authorization","Agreement","Information","Transcript","Testing","Tests","Scores",
    "Interim","Interims","Conference","Conferences","Consent","Plan","Accommodation","Accomodation",
    "Health","Immunization","Immunizations","Phys","Phy","Evaluation","Schedule","Status","Number",
    "Phone","Comments","Absence","Birth","Certificate","Certificates","Faxed","Documents","Sent",
    "Standardized","Gifted","PSAT","Bc","Im","Imm","Rc","Pt","St","Jk",
    # DVF and division words that caused false matches
    "All","Division","Divisions","DVF","Dvf",
    "Lower","Middle","Upper","School","Data","Verification","Form","Page","Grade","Date",
    # contact/guardian labels
    "Contact","Person","Persons","Guardian","Parent","Emergency","Subject","Please",
    # frequent campus words
    "Service","Hours","Community","Athletic","Event","Trip","Training","Weight",
    "Creative","Writing","Science","Government","History","Chinese","English",
    "American","Teacher","Pre-ap","AP","Honors","Honor","Roll",
    # place/organization noise seen in files
    "Oak","Hall","Hall's","Tampa","Gainesville","Meridian",
    # address bits
    "Street","St","Avenue","Ave","Road","Rd","Lane","Ln","Court","Ct","Drive","Dr","Boulevard","Blvd",
    "Suite","Ste","Apt","SW","NW","NE","SE",
    # generic placeholders
    "Last","First","Sample","Summary","Year","End",
}

# Some very short, but real last names we should allow (2-letter names are usually blocked).
SHORT_ALLOW = {"Li","Lu","Xu","Yu","Su","Wu","Ng","Ho","Hu","Ko","Do","He"}


# ------------------------- Regex patterns -------------------------
# A “regular expression” (regex) is a special pattern to match text.
# NAME_PART matches one word of a name: starts with a capital letter and can include
# letters, apostrophes, periods, or hyphens. Example: O'Neil, J.R., Anne-Marie, etc.
NAME_PART = r"[A-Z][a-zA-Z'.\-]+"

# Pattern for "Last, First" (the strongest signal for a name)
PAT_COMMA = re.compile(rf"\b({NAME_PART}),\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# Pattern for "First Last" (we will flip it into "Last, First" later)
PAT_SPACE = re.compile(rf"\b({NAME_PART})\s+({NAME_PART}(?:\s+{NAME_PART})*)\b")

# Pattern for "Last; First" (sometimes people use semicolons)
PAT_SEMI  = re.compile(rf"\b({NAME_PART});\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# SPECIAL DVF (Data Verification Form) PATTERN:
# On these forms, "Student Name: Last, First" is the rule you told us.
# We match that exact style and grab the pieces.
PAT_DVF = re.compile(
    rf"""
    Student\s*'?s?\s*Name\s*[:\-]\s*                # looks for "Student Name:" with optional spaces
    (?P<last>{NAME_PART}(?:\s+{NAME_PART})*)       # last name part(s)
    \s*[;,]\s*                                     # a comma or semicolon between last and first
    (?P<first>{NAME_PART}(?:\s+{NAME_PART})*)      # first name part(s)
    (?=\s*(?:$|[,;]|Grade(?:\s*:|\b)|Date(?:\s*:|\b)|Page\b|Lower\s+School|Data\s+Verification|Form\b))
    """,
    re.I | re.X
)

# CONTEXT_REJECT:
# If our match is sitting near these words (like "Guardian" or "Contact Person's Name"),
# it’s probably not the student’s name, so we ignore it.
CONTEXT_REJECT = re.compile(
    r"(Contact\s+Person'?s\s+Name|Guardian|Parent|Emergency|Student\s+Name\s*,?\s*Signature)"
    r"|(\bAddress\b|\bE-?mail\b|\bComments?\b|\bAbsence\b|\bInterim\b|\bBirth\b|\bCertificate)",
    re.I
)

# LABEL_TAIL:
# Sometimes real names have unwanted labels stuck after them (like "Form" or "Grade").
# This pattern helps us cut off those trailing labels.
LABEL_TAIL = re.compile(r"\s+(Grade|Date|Page|Lower|School|Data|Verification|Form)\b.*$", re.I)


# ------------------------- Small helpers --------------------------

def collapse_ws(s: str) -> str:
    """
    Replace weird spaces and dashes with normal ones, and collapse multiple spaces into one.
    This makes it easier to match names without being confused by strange characters.
    """
    s = s.replace("\u00A0"," ")  # non-breaking space -> normal space
    # turn different dash types into a normal hyphen
    s = s.replace("\u2010","-").replace("\u2011","-").replace("\u2013","-").replace("\u2014","-")
    # replace any group of whitespace with a single space
    return re.sub(r"\s+", " ", s)

def norm_cap(s: str) -> str:
    """
    Normalize capitalization: "hANNAH" -> "Hannah".
    We do this for each word in the string.
    """
    return " ".join(p[:1].upper() + p[1:].lower() if p else p for p in s.split())

def strip_labels(s: str) -> str:
    """
    Remove trailing labels like 'Form', 'Grade', etc. from the end of a name string.
    """
    return LABEL_TAIL.sub("", s).strip()

def token_ok(tok: str) -> bool:
    """
    Decide if a single word (token) looks like it could be part of a real name.
    Rules:
      - No digits
      - Not a STOPWORD (like 'Form' or 'Records')
      - Very short tokens (<=2 chars) are usually bad unless they are in SHORT_ALLOW
      - Very long ALL-CAPS tokens are suspicious (probably not a name)
    """
    if any(ch.isdigit() for ch in tok):
        return False
    if tok in STOPWORDS:
        return False
    if len(tok) <= 2 and tok not in SHORT_ALLOW:
        return False
    if tok.isupper() and len(tok) > 3:
        return False
    t = tok.lower()
    if t in {"email","e-mail"}:
        return False
    return True

def looks_like_name(last: str, first: str) -> bool:
    """
    Check if 'last' and 'first' together pass our token_ok checks.
    """
    toks = (last + " " + first).replace(",", " ").split()
    return all(token_ok(t) for t in toks)

def score_candidate(last: str, first: str, *, source: str, dvf_hit: bool, had_comma: bool, near_anchor: bool) -> int:
    """
    Give a numeric score to a possible (last, first) name.
    Higher scores mean “more likely to be the correct student name”.

    Rules that add points:
      +6 if it used a comma "Last, First" (this is very name-like)
      +4 if it was near helpful labels (anchors)
      +2 if found in a filename (filenames are pretty trustworthy)
      +1 if found in a PDF
     +12 if it matched the special DVF rule ("Student Name: Last, First")

    Small adjustments:
      +1 if the name is short/normal length (under 4 words total)
      -2 if the last name has spaces (multi-word last names are fine, but slightly riskier)
    """
    score = 10  # base score
    if had_comma:   score += 6
    if near_anchor: score += 4
    if source == "filename": score += 2
    if source == "pdf":      score += 1
    if dvf_hit:              score += 12
    # name length preference
    parts = len((last + " " + first).split())
    if parts <= 3:   score += 1
    if parts >= 5:   score -= 2
    # prefer single-word surnames a bit
    if " " in last:  score -= 2
    return score


# ------------------- Filename-based extraction --------------------

def candidates_from_filename(fname: str):
    """
    Look for names inside a FILE NAME (not the contents).
    We try three patterns:
      1) Last, First
      2) First Last  (we flip it to Last, First)
      3) Last; First
    For each match, we check if it looks like a real name and then yield a scored candidate.
    """
    # Path(...).stem gives the filename without the extension.
    # We also replace _ and - with spaces, and collapse multiple spaces.
    base = collapse_ws(Path(fname).stem.replace("_"," ").replace("-"," "))

    # 1) Prefer "Last, First"
    for m in PAT_COMMA.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))

    # 2) Also allow "First Last" (flip to Last, First)
    for m in PAT_SPACE.finditer(base):
        f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
        last, first = l, f
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=False, near_anchor=False))

    # 3) And "Last; First"
    for m in PAT_SEMI.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))


# -------------------- PDF-based extraction ------------------------

def pdf_text(pdf: Path, max_pages=MAX_PDF_PAGES) -> str:
    """
    Read up to 'max_pages' pages of text from a PDF file and return it as one string.
    If we fail to read the PDF, we return an empty string.
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
            # If a page can’t be read, skip it.
            pass
    # Join pages and clean up spaces/dashes for easier matching
    return collapse_ws(" ".join(chunks))

def candidates_from_pdf(pdf: Path):
    """
    Look for names INSIDE a PDF’s text.
    Steps:
      - Use the special DVF rule first: "Student Name: Last, First"
      - Then try "Last, First", "Last; First", and "First Last"
      - Ignore matches that sit near admin words (CONTEXT_REJECT)
      - Score everything and return all candidates we find
    """
    blob = pdf_text(pdf)
    if not blob:
        return []
    out = []

    # 1) Strong DVF rule (you told us DVFs always use Last, First after "Student Name:")
    for m in PAT_DVF.finditer(blob):
        last = strip_labels(norm_cap(m.group("last")))
        first = strip_labels(norm_cap(m.group("first")))
        if looks_like_name(last, first):
            out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=True, had_comma=True, near_anchor=True)))

    # 2) General fallbacks
    for pat, had_comma in ((PAT_COMMA, True), (PAT_SEMI, True), (PAT_SPACE, False)):
        for m in pat.finditer(blob):
            # If the pattern has a comma/semi, group(1)=last, group(2)=first
            if had_comma:
                last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
            else:
                # "First Last" -> flip to "Last, First"
                f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
                last, first = l, f

            # Take some surrounding text to check for bad context words.
            ctx = blob[max(0, m.start()-80): m.end()+80]
            if CONTEXT_REJECT.search(ctx):
                continue

            if looks_like_name(last, first):
                out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=False, had_comma=had_comma, near_anchor=False)))
    return out


# -------------------- Picking the best match ----------------------

def pick_best(cands):
    """
    We might find the SAME (last, first) pair multiple times with different scores.
    This function:
      1) Keeps only the HIGHEST score for each unique (last, first) pair.
      2) Sorts by best score (highest first).
      3) If scores tie, prefer fewer total words (shorter names are safer).
      4) If still tied, prefer shorter total character length.
    Returns:
      (last, first) of the winner, or None if no candidates.
    """
    if not cands:
        return None

    # Keep the highest score for each unique name pair
    best = {}
    for last, first, sc in cands:
        key = (last, first)
        if key not in best or sc > best[key]:
            best[key] = sc

    # Sort using our rules
    def sort_key(kv):
        (last, first), sc = kv
        tokens = len((last + " " + first).split())
        total_len = len(last.replace(" ","")) + len(first.replace(" ",""))
        return (-sc, tokens, total_len)

    # Return the (last, first) with the best key
    return sorted(best.items(), key=sort_key)[0][0]


# --------------------------- Renaming -----------------------------

def target_name(id8: str, last_first):
    """
    Build the final folder name string: "######## - Last, First"
    """
    last, first = last_first
    return f"{id8} - {last}, {first}"

def safe_rename(folder: Path, new_base: str):
    """
    Actually rename the folder (or just print what we would do if DRY_RUN is True).

    Safety checks:
      - If the new name is exactly the same, do nothing.
      - If a folder with the new name already exists, add " (1)", " (2)", etc.
    """
    target = folder.with_name(new_base)

    # If it's already the same name, say so.
    if target.name == folder.name:
        print(f"[SAME] {folder.name}")
        return

    # Avoid collisions: if target exists, append (1), (2), ...
    if target.exists():
        i = 1
        while folder.with_name(f"{new_base} ({i})").exists():
            i += 1
        target = folder.with_name(f"{new_base} ({i})")

    # Do the rename or just print it (DRY RUN)
    if DRY_RUN:
        print(f"[DRY]  {folder.name} -> {target.name}")
    else:
        folder.rename(target)
        print(f"[OK]   {folder.name} -> {target.name}")


# ---------------------------- Main --------------------------------

def derive_name_for_folder(folder: Path):
    """
    This is the brain for one folder:
      1) Look at ALL filenames in this folder for a good name.
      2) If that fails, read INSIDE ALL PDFs in this folder to find a name.
      3) Return the best (last, first) we can find, or None if nothing is reliable.
    """
    # --- Step 1: check ALL file NAMES first (fast and usually enough) ---
    file_name_cands = []
    for f in folder.iterdir():
        if f.is_file():
            file_name_cands.extend(candidates_from_filename(f.name))

    best_from_names = pick_best(file_name_cands)
    if best_from_names:
        return best_from_names

    # --- Step 2: if filenames failed, check INSIDE ALL PDFs ---
    pdf_cands = []
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() == ".pdf":
            pdf_cands.extend(candidates_from_pdf(f))

    return pick_best(pdf_cands)

def main():
    """
    Go through each item in ROOT.
    If it’s a directory named with exactly 8 digits (like '12345678'),
    try to figure out the student's name and rename the folder.
    """
    for child in ROOT.iterdir():
        # Skip things that aren't folders
        if not child.is_dir():
            continue

        # We only process folders with names like "12345678"
        if not ID8.fullmatch(child.name):
            continue

        # Try to derive the name for this specific folder
        best = derive_name_for_folder(child)

        if best:
            # If we found something, build the new folder name and rename safely
            safe_rename(child, target_name(child.name, best))
        else:
            # If nothing looked reliable, we skip (and tell the user)
            print(f"[SKIP] {child.name}  no reliable name")


# This runs 'main()' when the file is executed directly.
if __name__ == "__main__":
    main()
