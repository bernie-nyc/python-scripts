"""
WHAT THIS PROGRAM DOES
----------------------
You have lots of student folders named with 8 digits like "12305932".
Inside each folder are files (PDFs, etc.). We want to rename each folder to:
"######## - Last, First"  (example: "12305932 - Wynn, Hannah").

HOW WE FIND THE NAME (for EACH folder)
--------------------------------------
1) Look at the NAMES of ALL files in that folder.
   - If a filename contains a person name (like “Wynn, Hannah—Immunizations.pdf”
     or “Hannah Wynn.pdf”), we use it.
2) If that fails, look INSIDE ALL the PDFs (first few pages).
   - Find patterns like “Last, First” or “First Last” (then flip to “Last, First”).
   - Special rule for “Data Verification Form” (DVF): “Student Name: Last, First”.

HOW WE AVOID BAD NAMES
----------------------
We block generic words from becoming names (case-insensitive):
- Admin words: “verification”, “data”, “school”, “form”, “recordsrequest”, “releaseforrecords”,
  “psat”, “medical”, “social security card”, months, weekdays, subjects, etc.
- Obvious junk PAIRS like “verification, data” or “school, lower”.
- If a token is in the STOPWORDS list, we drop it.
- If, after cleaning, a side of the name is empty → reject.

SAFETY / RELIABILITY
--------------------
- We “score” better-looking name matches (comma style is good; DVF rule is very strong).
- DRY_RUN mode prints changes without renaming.
- Safer renaming:
  • Retry a few times on Windows “Access is denied” (WinError 5).
  • If the new name already exists, add " (1)", " (2)", etc.

HOW TO USE
----------
1) Set ROOT below to your top folder.
2) Start with DRY_RUN = True to preview.
3) When happy, set DRY_RUN = False and run again.
4) Close any Explorer windows or files open inside the folders to avoid lock errors.
"""

from pathlib import Path
import re
import time
import os
import logging
from PyPDF2 import PdfReader

# Quiet PyPDF2 logs (some PDFs may still print harmless “unknown widths” lines).
logging.getLogger("PyPDF2").setLevel(logging.ERROR)

# ---------------------------- Settings ----------------------------

# Top folder containing the 8-digit student folders:
ROOT = Path(r"U:\General Portfolio\No_Match")

# True = preview only (no actual rename). False = do the renames.
DRY_RUN = True

# Read up to this many pages from each PDF when searching inside files.
MAX_PDF_PAGES = 3

# Subfolder must be exactly 8 digits to be considered a student folder.
ID8 = re.compile(r"^\d{8}$")

# Retry counts for Windows "Access is denied" during rename.
RENAMING_RETRIES = 5
RETRY_WAIT_SECONDS = 1.5

# ------------------ Stopwords / Banned tokens (case-insensitive) ------------------
# All entries are lower-case; we compare using .lower() so “PSAT”, “Psat”, “psat” all match.
STOPWORDS_LOWER = {
    # admin/common
    "address","application","admissions","admission","admin","administrative","administration",
    "form","forms","report","release","records","record","request","requests","authorization",
    "agreement","information","transcript","transcripts","testing","tests","scores","interim",
    "interims","conference","conferences","consent","plan","accommodation","accomodation",
    "acommodation",  # common misspellings
    "health","immunization","immunizations","phys","phy","evaluation","evaluations","eval",
    "schedule","status","number","phone","comments","absence","absences","birth","certificate",
    "faxed","documents","sent","standardized","gifted","psat","bc","im","imm","rc","pt","stk",
    "dvf","division","divisions","all","student",

    # contact/guardian labels
    "contact","person","persons","guardian","parent","emergency","subject","please",

    # subjects / campus-y words
    "service","hours","community","athletic","event","trip","training","weight","creative","writing",
    "science","government","history","chinese","english","american","teacher","pre-ap","ap","honors",
    "honor","roll","kindergarten","junior","sample",

    # school/org/place noise
    "school","lower","upper","middle","oak","hall","hall's","tampa","gainesville","meridian","yonge",
    "p.k.","p.k","pk","pk.",

    # address bits
    "street","st","avenue","ave","road","rd","lane","ln","court","ct","drive","dr","boulevard","blvd",
    "suite","ste","apt","sw","nw","ne","se",

    # generic placeholders / traps
    "last","first","name","signature",

    # tests/program words
    "iowa","sat","act","math","eng","physical","pe","eval","sample",

    # social security card bits
    "soc","sec","card","social","security",

    # months (block “month, day” fake names)
    "january","february","march","april","may","june","july","august","september","october",
    "november","december","jan","feb","mar","apr","jun","jul","aug","sep","sept","oct","nov","dec",

    # weekdays
    "monday","tuesday","wednesday","thursday","friday","saturday","sunday",
    "mon","tue","tues","wed","thu","thur","thurs","fri","sat","sun",

    # known fused or misspelled admin blobs
    "recordsrequest","recordrequest","requestrecords","requestreocrds",
    "recordreleaseauthorization","recordrelease","recordsrelease",
    "administering","administered","admis","admis eval","admiseval","socsec","socseccard",

    # other junk we saw in outputs
    "verification","data","medical","flvs","lms","justice","press","squats","president",
    "representative","average","avg","score","test","releaseforrecords","requestreocrds",
}

# HARD reject pairs like (“verification”, “data”) regardless of anything else.
BAD_PAIR_LOWER = {
    ("verification", "data"),
    ("school", "lower"),
    ("lower", "school"),
    ("score", "test"),
    ("average", "avg"),
    ("student", "name"),
}

# allow legit two-letter surnames
SHORT_ALLOW = {"li","lu","xu","yu","su","wu","ng","ho","hu","ko","do","he"}

# common suffixes
SUFFIX_ALLOW = {"jr","sr","ii","iii","iv","v"}

# ------------------------- Patterns (Regex) ------------------------

# One “word” of a name: starts with a capital; may contain letters, apostrophes, periods, hyphens.
NAME_PART = r"[A-Z][a-zA-Z'.\-]+"

# 1) “Last, First”
PAT_COMMA = re.compile(rf"\b({NAME_PART}),\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# 2) “First Last” (we’ll flip to “Last, First”)
PAT_SPACE = re.compile(rf"\b({NAME_PART})\s+({NAME_PART}(?:\s+{NAME_PART})*)\b")

# 3) “Last; First”
PAT_SEMI  = re.compile(rf"\b({NAME_PART});\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# Special DVF rule: “Student Name: Last, First” with typical words following it.
PAT_DVF = re.compile(
    rf"""
    Student\s*'?s?\s*Name\s*[:\-]\s*
    (?P<last>{NAME_PART}(?:\s+{NAME_PART})*)
    \s*[;,]\s*
    (?P<first>{NAME_PART}(?:\s+{NAME_PART})*)
    (?=\s*(?:$|[,;]|Grade(?:\s*:|\b)|Date(?:\s*:|\b)|Page\b|Lower\s+School|Data\s+Verification|Form\b))
    """,
    re.I | re.X
)

# If these appear near a match in a PDF, it’s probably not the student’s name.
CONTEXT_REJECT = re.compile(
    r"(Contact\s+Person'?s\s+Name|Guardian|Parent|Emergency|Student\s+Name\s*,?\s*Signature)"
    r"|(\bAddress\b|\bE-?mail\b|\bComments?\b|\bAbsence\b|\bInterim\b|\bBirth\b|\bCertificate)"
    r"|(\bAdmission(?:s)?\b|\bAdministering\b|\bEvaluation(?:s)?\b|\bEval\b)"
    r"|(\bRecord(?:s)?(?:request|release|authorization)\b|\bRequestreocrds\b|\bRequestrecords\b)"
    r"|(\bData\s+Verification\b|\bLower\s+School\b|\bVerification\b|\bForm\b)",
    re.I
)

# Remove trailing labels from detected sides (extra cleanup). Made more flexible.
LABEL_TAIL = re.compile(
    r"\s+(grade|date|page|lower|school|data|verification|form|admission(?:s)?|administering|"
    r"evaluation(?:s)?|eval|records?request|requestreocrds|recordrelease(?:authorization)?|"
    r"soc(?:ial)?\s*sec(?:urity)?\s*card|iowa|english|math|medical|physical)\b.*$",
    re.I
)

# Pre-clean obvious garbage phrases from FILENAMES before scanning for names.
FILENAME_NOISE = re.compile(
    r"\b(All\s+Divisions?,?\s*DVF|DVF|Admission(?:s)?\s+Eval(?:uation)?s?|Administering|"
    r"Records?request|Requestreocrds|Recordrelease(?:authorization)?|Soc(?:ial)?\s*Sec(?:urity)?\s*Card|"
    r"Iowa|English|Math|Data\s+Verification|Lower\s+School|Student\s+Name)\b",
    re.I
)

# ------------------------- Small helpers --------------------------

def collapse_ws(s: str) -> str:
    """Normalize spaces/dashes so patterns match more easily."""
    s = s.replace("\u00A0"," ")  # non-breaking space
    s = (s.replace("\u2010","-").replace("\u2011","-")
           .replace("\u2013","-").replace("\u2014","-"))
    return re.sub(r"\s+", " ", s).strip()

def norm_cap(s: str) -> str:
    """Capitalize nicely: 'hANNAH wYNN' -> 'Hannah Wynn'."""
    return " ".join(p[:1].upper() + p[1:].lower() if p else p for p in s.split())

def is_stopword_token(tok: str) -> bool:
    """Case-insensitive stopword check."""
    return tok.lower() in STOPWORDS_LOWER

def strip_labels(side: str) -> str:
    """
    Remove trailing labels like 'Form', 'Eval', 'Iowa', 'Medical', etc.
    Also, if the entire side is just a stopword (e.g., 'School'), blank it out.
    """
    side = side.strip()
    if not side:
        return side
    # remove trailing label chunks
    side = LABEL_TAIL.sub("", side).strip()
    # if the whole side is a stopword, drop it
    if is_stopword_token(side):
        return ""
    return side

def token_ok(tok: str) -> bool:
    """
    Is this word allowed to be part of a real name?
    - No digits
    - Not a stopword
    - Very short (<=2) only allowed if in SHORT_ALLOW
    - Not random ALLCAPS (except short pieces)
    """
    if not tok:
        return False
    if any(ch.isdigit() for ch in tok):
        return False
    tlo = tok.lower()
    if tlo in STOPWORDS_LOWER:
        return False
    if len(tlo) <= 2 and tlo not in SHORT_ALLOW and tlo not in {s.lower() for s in SUFFIX_ALLOW}:
        return False
    if tok.isupper() and len(tok) > 3:
        return False
    if tlo in {"email","e-mail"}:
        return False
    return True

def clean_name_side(side_text: str) -> str:
    """
    Clean one side of a name (“Last” OR “First Middle”):
    - Remove trailing labels (strip_labels)
    - Keep only tokens that look like real name pieces
    - Allow suffixes like Jr, III
    If nothing real remains, return "".
    """
    side_text = strip_labels(side_text)
    tokens = side_text.replace(",", " ").split()
    keep = []
    for t in tokens:
        # normalize caps (keep apostrophes/hyphens)
        t_norm = t[:1].upper() + t[1:] if t else t
        tlo = t_norm.lower()
        if tlo in {s.lower() for s in SUFFIX_ALLOW} or token_ok(t_norm):
            keep.append(t_norm)
    return " ".join(keep)

def looks_like_name(last: str, first: str) -> bool:
    """
    Final test for “Last, First” after cleaning:
    - Both sides must still have content
    - Neither side can be a lone stopword
    - Pair cannot be in BAD_PAIR_LOWER (e.g., "verification, data")
    - Every token must pass token_ok or be a suffix
    """
    last = clean_name_side(last)
    first = clean_name_side(first)
    if not last or not first:
        return False

    last_lo = last.lower()
    first_lo = first.lower()
    if (last_lo, first_lo) in BAD_PAIR_LOWER:
        return False

    toks = (last + " " + first).replace(",", " ").split()
    if not all(token_ok(t) or t.lower() in {s.lower() for s in SUFFIX_ALLOW} for t in toks):
        return False
    return True

def score_candidate(last: str, first: str, *, source: str, dvf_hit: bool, had_comma: bool, near_anchor: bool) -> int:
    """
    Score how trustworthy a candidate is:
    + comma style helps
    + near good anchors helps
    + DVF rule is very strong
    + filenames are somewhat strong; PDFs are fine too
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

def preclean_filename_text(raw: str) -> str:
    """
    Remove obvious garbage phrases from a filename BEFORE scanning for names.
    Example: “Wynn, Hannah - Immunizations.pdf” stays useful, but
             “... Data Verification ...”, “... Lower School ...”, etc., are scrubbed.
    """
    base = Path(raw).stem
    # turn underscores/dashes into spaces to expose words
    base = base.replace("_"," ").replace("-"," ")
    base = collapse_ws(base)
    # delete obvious noise chunks
    base = FILENAME_NOISE.sub("", base)
    # collapse extra spaces again
    base = collapse_ws(base)
    return base

def candidates_from_filename(fname: str):
    """
    Find names in the FILE NAME (not the contents).
    We try:
      1) Last, First
      2) First Last  (we flip to Last, First)
      3) Last; First
    Return every good-looking match with a score.
    """
    base = preclean_filename_text(fname)

    # Last, First
    for m in PAT_COMMA.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (clean_name_side(last), clean_name_side(first),
                   score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))

    # First Last (flip)
    for m in PAT_SPACE.finditer(base):
        f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
        last, first = l, f
        if looks_like_name(last, first):
            yield (clean_name_side(last), clean_name_side(first),
                   score_candidate(last, first, source="filename", dvf_hit=False, had_comma=False, near_anchor=False))

    # Last; First
    for m in PAT_SEMI.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (clean_name_side(last), clean_name_side(first),
                   score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))

# -------------------- PDF-based extraction ------------------------

def pdf_text(pdf: Path, max_pages=MAX_PDF_PAGES) -> str:
    """Read text from up to 'max_pages' of a PDF. If reading fails, return ""."""
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
      - Special DVF rule (“Student Name: Last, First”).
      - General patterns (“Last, First”, “Last; First”, “First Last”).
      - Reject if near bad context words (guardian, address, etc.).
      - Clean out trailing labels like “Form”, “Iowa”, “PSAT”, “Medical”.
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
            out.append((clean_name_side(last), clean_name_side(first),
                        score_candidate(last, first, source="pdf", dvf_hit=True, had_comma=True, near_anchor=True)))

    # General patterns
    for pat, had_comma in ((PAT_COMMA, True), (PAT_SEMI, True), (PAT_SPACE, False)):
        for m in pat.finditer(blob):
            if had_comma:
                last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
            else:
                f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
                last, first = l, f

            # neighborhood text for context check
            ctx = blob[max(0, m.start()-80): m.end()+80]
            if CONTEXT_REJECT.search(ctx):
                continue

            last_clean = clean_name_side(last)
            first_clean = clean_name_side(first)
            if last_clean and first_clean and looks_like_name(last_clean, first_clean):
                out.append((last_clean, first_clean,
                            score_candidate(last_clean, first_clean, source="pdf", dvf_hit=False, had_comma=had_comma, near_anchor=False)))
    return out

# -------------------- Picking the best match ----------------------

def pick_best(cands):
    """
    From many candidates, keep the highest score for each unique (last, first),
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
    - If Windows says "Access is denied" (WinError 5), retry a few times.
    - Fail gracefully instead of crashing the whole script.
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

    last_err = None
    for attempt in range(1, RENAMING_RETRIES + 1):
        try:
            folder.rename(target)
            print(f"[OK]   {folder.name} -> {target.name}")
            return
        except PermissionError as e:
            last_err = e
            print(f"[WAIT] {folder.name} locked (attempt {attempt}/{RENAMING_RETRIES}). "
                  f"Close files or Explorer windows. Retrying in {RETRY_WAIT_SECONDS}s...")
            time.sleep(RETRY_WAIT_SECONDS)
        except OSError as e:
            print(f"[FAIL] {folder.name} -> {target.name}  ({e})")
            return

    print(f"[FAIL] {folder.name} -> {target.name}  (Access denied after retries: {last_err})")

# ---------------------------- Main --------------------------------

def derive_name_for_folder(folder: Path):
    """
    Figure out the student’s name for ONE folder.

    IMPORTANT: We check ALL files in the folder.
      1) Try ALL filenames first (fast).
      2) If needed, look INSIDE ALL PDFs (slower but powerful).

    Return (last, first) or None.
    """
    # Step 1: filenames (ALL files)
    file_name_cands = []
    for f in folder.iterdir():
        if f.is_file():
            file_name_cands.extend(candidates_from_filename(f.name))

    best_from_names = pick_best(file_name_cands)
    if best_from_names:
        return best_from_names

    # Step 2: PDF contents (ALL PDFs)
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
