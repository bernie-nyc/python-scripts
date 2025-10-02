"""
fileindexer.py — FULLY COMMENTED (middle-school friendly) + safer renaming

WHAT THIS PROGRAM DOES:
-----------------------
You have a big folder full of student folders named like "12305932".
Inside each student folder are files (PDFs, etc.). We want to rename each
student folder to: "######## - Last, First" (example: "12305932 - Wynn, Hannah").

HOW WE FIND THE NAME (for EACH student folder):
1) Look at the NAMES of ALL files in that folder.
   - If a filename already includes a name (like “Wynn, Hannah” or “Hannah Wynn”),
     we use that.
2) If we still don’t know the name, we look INSIDE ALL the PDFs.
   - We read text from the first few pages and search for patterns like:
       • “Student Name: Last, First”  (special rule for Data Verification Forms)
       • “Last, First”
       • “First Last” (we flip it to “Last, First”)
3) We AVOID being tricked by generic words like “All Divisions, DVF”.

SAFETY / RELIABILITY:
---------------------
- A big list of “STOPWORDS” keeps non-name words out.
- We “score” better-looking names (like ones that have a comma “Last, First”).
- DRY_RUN mode shows what would happen WITHOUT renaming for real.
- Safer renaming:
  • We catch Windows “Access is denied” (WinError 5) which usually means a file
    in that folder is open in another program (Explorer, PDF viewer, sync tool).
  • We retry a few times with short pauses instead of crashing the whole run.
  • If a target name exists already, we add " (1)", " (2)", etc.

HOW TO USE:
-----------
1) Set ROOT (below) to the top “No_Match” folder with the student folders.
2) Start with DRY_RUN = True (prints what it WOULD do).
3) When happy, set DRY_RUN = False to actually rename.
4) Close Explorer windows and any open files inside the folders before running,
   to avoid “Access is denied” errors.
"""

from pathlib import Path
import re
import time
import os
from PyPDF2 import PdfReader


# ---------------------------- Settings ----------------------------

ROOT = Path(r"U:\General Portfolio\No_Match")  # top-level folder that holds the 8-digit folders
DRY_RUN = True                                  # True = preview only; False = actually rename

MAX_PDF_PAGES = 3                               # how many pages to read from each PDF (speeds things up)
ID8 = re.compile(r"^\d{8}$")                    # matches folder names like "12345678"

# When trying to rename, we’ll retry if Windows says “Access is denied”
RENAMING_RETRIES = 5
RETRY_WAIT_SECONDS = 1.5                        # wait this many seconds between retries


# ------------------ Words we do NOT want in names -----------------

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

# allow legit two-letter last names (normally we reject <=2 letter tokens)
SHORT_ALLOW = {"Li","Lu","Xu","Yu","Su","Wu","Ng","Ho","Hu","Ko","Do","He"}


# ------------------------- Patterns (Regex) ------------------------

# NAME_PART is one “word” of a name: starts with capital letter and can include letters,
# apostrophes, periods, or hyphens (e.g., O'Neil, J.R., Anne-Marie).
NAME_PART = r"[A-Z][a-zA-Z'.\-]+"

# 1) “Last, First”
PAT_COMMA = re.compile(rf"\b({NAME_PART}),\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# 2) “First Last”  (we’ll flip to “Last, First”)
PAT_SPACE = re.compile(rf"\b({NAME_PART})\s+({NAME_PART}(?:\s+{NAME_PART})*)\b")

# 3) “Last; First”
PAT_SEMI  = re.compile(rf"\b({NAME_PART});\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# Special DVF rule: “Student Name: Last, First”
PAT_DVF = re.compile(
    rf"""
    Student\s*'?s?\s*Name\s*[:\-]\s*                # "Student Name:" (flexible spaces)
    (?P<last>{NAME_PART}(?:\s+{NAME_PART})*)       # last name
    \s*[;,]\s*
    (?P<first>{NAME_PART}(?:\s+{NAME_PART})*)      # first name
    (?=\s*(?:$|[,;]|Grade(?:\s*:|\b)|Date(?:\s*:|\b)|Page\b|Lower\s+School|Data\s+Verification|Form\b))
    """,
    re.I | re.X
)

# If these words are near a match, it’s probably not the student name (skip it).
CONTEXT_REJECT = re.compile(
    r"(Contact\s+Person'?s\s+Name|Guardian|Parent|Emergency|Student\s+Name\s*,?\s*Signature)"
    r"|(\bAddress\b|\bE-?mail\b|\bComments?\b|\bAbsence\b|\bInterim\b|\bBirth\b|\bCertificate)",
    re.I
)

# Remove trailing labels like “… Form”, “… Grade”, etc., from a matched name.
LABEL_TAIL = re.compile(r"\s+(Grade|Date|Page|Lower|School|Data|Verification|Form)\b.*$", re.I)


# ------------------------- Small helpers --------------------------

def collapse_ws(s: str) -> str:
    """Normalize spaces/dashes so patterns match more easily."""
    s = s.replace("\u00A0"," ")  # non-breaking space
    s = (s.replace("\u2010","-").replace("\u2011","-")
           .replace("\u2013","-").replace("\u2014","-"))
    return re.sub(r"\s+", " ", s)

def norm_cap(s: str) -> str:
    """Capitalize nicely: 'hANNAH wYNN' -> 'Hannah Wynn'."""
    return " ".join(p[:1].upper() + p[1:].lower() if p else p for p in s.split())

def strip_labels(s: str) -> str:
    """Cut off tail labels like 'Form', 'Grade', etc."""
    return LABEL_TAIL.sub("", s).strip()

def token_ok(tok: str) -> bool:
    """
    Decide if one word could be part of a real name.
    (No digits, not a STOPWORD, not random ALLCAPS, short words allowed only if in SHORT_ALLOW.)
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
    """Check every word in 'Last, First' with token_ok()."""
    toks = (last + " " + first).replace(",", " ").split()
    return all(token_ok(t) for t in toks)

def score_candidate(last: str, first: str, *, source: str, dvf_hit: bool, had_comma: bool, near_anchor: bool) -> int:
    """
    Give a higher score to more trustworthy matches.
    - comma style “Last, First” is great
    - being near good anchors helps
    - filenames are trustworthy
    - DVF rule is very strong
    """
    score = 10
    if had_comma:   score += 6
    if near_anchor: score += 4
    if source == "filename": score += 2
    if source == "pdf":      score += 1
    if dvf_hit:              score += 12

    parts = len((last + " " + first).split())
    if parts <= 3: score += 1
    if parts >= 5: score -= 2
    if " " in last: score -= 2
    return score


# ------------------- Filename-based extraction --------------------

def candidates_from_filename(fname: str):
    """
    Search a FILE NAME for names (not the contents). We try patterns:
      1) Last, First
      2) First Last  (flipped to Last, First)
      3) Last; First
    We return every good-looking match with a score.
    """
    base = collapse_ws(Path(fname).stem.replace("_"," ").replace("-"," "))

    # Last, First
    for m in PAT_COMMA.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))

    # First Last (flip)
    for m in PAT_SPACE.finditer(base):
        f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
        last, first = l, f
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=False, near_anchor=False))

    # Last; First
    for m in PAT_SEMI.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))


# -------------------- PDF-based extraction ------------------------

def pdf_text(pdf: Path, max_pages=MAX_PDF_PAGES) -> str:
    """
    Read text from up to 'max_pages' of a PDF. If we can’t read it, return "".
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
            pass
    return collapse_ws(" ".join(chunks))

def candidates_from_pdf(pdf: Path):
    """
    Find names INSIDE a PDF:
      - Try the special DVF rule (“Student Name: Last, First”).
      - Then try general patterns (“Last, First”, “Last; First”, “First Last”).
      - Ignore matches in bad contexts (like near “Guardian”).
    """
    blob = pdf_text(pdf)
    if not blob:
        return []
    out = []

    # Strong DVF rule first
    for m in PAT_DVF.finditer(blob):
        last = strip_labels(norm_cap(m.group("last")))
        first = strip_labels(norm_cap(m.group("first")))
        if looks_like_name(last, first):
            out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=True, had_comma=True, near_anchor=True)))

    # General patterns
    for pat, had_comma in ((PAT_COMMA, True), (PAT_SEMI, True), (PAT_SPACE, False)):
        for m in pat.finditer(blob):
            if had_comma:
                last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
            else:
                f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
                last, first = l, f

            ctx = blob[max(0, m.start()-80): m.end()+80]  # neighborhood text for context check
            if CONTEXT_REJECT.search(ctx):
                continue

            if looks_like_name(last, first):
                out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=False, had_comma=had_comma, near_anchor=False)))
    return out


# -------------------- Picking the best match ----------------------

def pick_best(cands):
    """
    From many candidates (some duplicates), keep the highest score for each unique (last, first),
    then sort and pick the winner.
    """
    if not cands:
        return None

    best = {}
    for last, first, sc in cands:
        key = (last, first)
        if key not in best or sc > best[key]:
            best[key] = sc

    def sort_key(kv):
        (last, first), sc = kv
        tokens = len((last + " " + first).split())
        total_len = len(last.replace(" ","")) + len(first.replace(" ",""))
        return (-sc, tokens, total_len)

    return sorted(best.items(), key=sort_key)[0][0]


# --------------------------- Renaming -----------------------------

def target_name(id8: str, last_first):
    """Build final folder name like: '12345678 - Last, First'."""
    last, first = last_first
    return f"{id8} - {last}, {first}"

def safe_rename(folder: Path, new_base: str):
    """
    Try to rename the folder safely.
    - If the new name already exists, append " (1)", " (2)", ...
    - If Windows says "Access is denied" (WinError 5), we retry a few times,
      pausing between tries, and then give up gracefully (don’t crash the whole script).
    """
    target = folder.with_name(new_base)

    # No-op if already correct
    if target.name == folder.name:
        print(f"[SAME] {folder.name}")
        return

    # Avoid name collisions
    if target.exists():
        i = 1
        while folder.with_name(f"{new_base} ({i})").exists():
            i += 1
        target = folder.with_name(f"{new_base} ({i})")

    if DRY_RUN:
        print(f"[DRY]  {folder.name} -> {target.name}")
        return

    # Real rename with retries (helps when files are briefly “locked” by other apps)
    last_err = None
    for attempt in range(1, RENAMING_RETRIES + 1):
        try:
            folder.rename(target)
            print(f"[OK]   {folder.name} -> {target.name}")
            return
        except PermissionError as e:
            last_err = e
            # On Windows, WinError 5 = “Access is denied”.
            # Usually a file is open in Explorer, a PDF viewer, antivirus scan, or sync tool.
            print(f"[WAIT] {folder.name} locked (attempt {attempt}/{RENAMING_RETRIES}). "
                  f"Close files or Explorer windows. Retrying in {RETRY_WAIT_SECONDS}s...")
            time.sleep(RETRY_WAIT_SECONDS)
        except OSError as e:
            # Other OS errors (print and stop retrying—usually won’t fix themselves)
            print(f"[FAIL] {folder.name} -> {target.name}  ({e})")
            return

    # If we’re here, all retries failed with PermissionError.
    print(f"[FAIL] {folder.name} -> {target.name}  (Access denied after retries: {last_err})")


# ---------------------------- Main --------------------------------

def derive_name_for_folder(folder: Path):
    """
    Figure out the student’s name for ONE folder.
    1) Try ALL filenames.
    2) If needed, look INSIDE ALL PDFs.
    Return (last, first) or None.
    """
    # Step 1: filenames
    file_name_cands = []
    for f in folder.iterdir():
        if f.is_file():
            file_name_cands.extend(candidates_from_filename(f.name))

    best_from_names = pick_best(file_name_cands)
    if best_from_names:
        return best_from_names

    # Step 2: PDF contents
    pdf_cands = []
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() == ".pdf":
            pdf_cands.extend(candidates_from_pdf(f))

    return pick_best(pdf_cands)

def main():
    """
    Walk through ROOT and process every subfolder named like 8 digits (e.g., "12345678").
    For each one, try to find “Last, First” and rename the folder safely.
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
