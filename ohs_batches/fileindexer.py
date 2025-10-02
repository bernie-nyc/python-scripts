# fileindexer.py
# Renames 8-digit student folders to: "######## - Last, First"
# Order of operations (per folder):
#   1) Look at ALL filenames for "Last, First" (preferred) or "First Last".
#   2) If none found, read text from ALL PDFs (first few pages) and try again.
# Guards:
#   - Blocks generic/admin phrases (e.g., "All Divisions, DVF") from being
#     mistaken as names.

from pathlib import Path
import re
from PyPDF2 import PdfReader

# ---------------------------- Settings ----------------------------

ROOT = Path(r"U:\General Portfolio\No_Match")
DRY_RUN = True
MAX_PDF_PAGES = 3
ID8 = re.compile(r"^\d{8}$")

# ------------------ Noise words (single tokens) -------------------
# Any token here cannot appear in a real student's name.
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

# Some short, valid name tokens we should allow
SHORT_ALLOW = {"Li","Lu","Xu","Yu","Su","Wu","Ng","Ho","Hu","Ko","Do","He"}

# ------------------------- Regex patterns -------------------------

NAME_PART = r"[A-Z][a-zA-Z'.\-]+"

# Strongest: "Last, First"
PAT_COMMA = re.compile(rf"\b({NAME_PART}),\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# Backup: "First Last"
PAT_SPACE = re.compile(rf"\b({NAME_PART})\s+({NAME_PART}(?:\s+{NAME_PART})*)\b")

# Sometimes: "Last; First"
PAT_SEMI  = re.compile(rf"\b({NAME_PART});\s*({NAME_PART}(?:\s+{NAME_PART})*)\b")

# DVF line: Student Name: Last, First   (student name is always Last, First on DVF)
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

# Ignore surrounding admin context when scanning inside PDFs
CONTEXT_REJECT = re.compile(
    r"(Contact\s+Person'?s\s+Name|Guardian|Parent|Emergency|Student\s+Name\s*,?\s*Signature)"
    r"|(\bAddress\b|\bE-?mail\b|\bComments?\b|\bAbsence\b|\bInterim\b|\bBirth\b|\bCertificate)",
    re.I
)

# Strip trailing labels mistakenly glued to a name
LABEL_TAIL = re.compile(r"\s+(Grade|Date|Page|Lower|School|Data|Verification|Form)\b.*$", re.I)

# ------------------------- Small helpers --------------------------

def collapse_ws(s: str) -> str:
    s = s.replace("\u00A0"," ")
    s = s.replace("\u2010","-").replace("\u2011","-").replace("\u2013","-").replace("\u2014","-")
    return re.sub(r"\s+", " ", s)

def norm_cap(s: str) -> str:
    return " ".join(p[:1].upper() + p[1:].lower() if p else p for p in s.split())

def strip_labels(s: str) -> str:
    return LABEL_TAIL.sub("", s).strip()

def token_ok(tok: str) -> bool:
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
    toks = (last + " " + first).replace(",", " ").split()
    return all(token_ok(t) for t in toks)

def score_candidate(last: str, first: str, *, source: str, dvf_hit: bool, had_comma: bool, near_anchor: bool) -> int:
    score = 10
    if had_comma:   score += 6        # "Last, First" is best
    if near_anchor: score += 4        # near helpful label
    if source == "filename": score += 2
    if source == "pdf":      score += 1
    if dvf_hit:      score += 12       # DVF rule is very strong
    # Light preference for typical two-part names
    parts = len((last + " " + first).split())
    if parts <= 3:   score += 1
    if parts >= 5:   score -= 2
    # Prefer single-word surnames over multi-word (reduces "All Divisions")
    if " " in last:  score -= 2
    return score

# ------------------- Filename-based extraction --------------------

def candidates_from_filename(fname: str):
    base = collapse_ws(Path(fname).stem.replace("_"," ").replace("-"," "))
    # 1) Prefer "Last, First"
    for m in PAT_COMMA.finditer(base):
        last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
        if looks_like_name(last, first):
            yield (last, first, score_candidate(last, first, source="filename", dvf_hit=False, had_comma=True, near_anchor=False))
    # 2) Also allow "First Last"
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
    blob = pdf_text(pdf)
    if not blob:
        return []
    out = []
    # Strong DVF rule
    for m in PAT_DVF.finditer(blob):
        last = strip_labels(norm_cap(m.group("last")))
        first = strip_labels(norm_cap(m.group("first")))
        if looks_like_name(last, first):
            out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=True, had_comma=True, near_anchor=True)))
    # General fallbacks if needed
    for pat, had_comma in ((PAT_COMMA, True), (PAT_SEMI, True), (PAT_SPACE, False)):
        for m in pat.finditer(blob):
            if had_comma:
                last, first = norm_cap(m.group(1)), norm_cap(m.group(2))
            else:
                f, l = norm_cap(m.group(1)), norm_cap(m.group(2))
                last, first = l, f
            ctx = blob[max(0, m.start()-80): m.end()+80]
            if CONTEXT_REJECT.search(ctx):
                continue
            if looks_like_name(last, first):
                out.append((last, first, score_candidate(last, first, source="pdf", dvf_hit=False, had_comma=had_comma, near_anchor=False)))
    return out

# -------------------- Picking the best match ----------------------

def pick_best(cands):
    if not cands:
        return None
    best = {}
    for last, first, sc in cands:
        key = (last, first)
        if key not in best or sc > best[key]:
            best[key] = sc
    # Sort by (score desc, tokens asc, total length asc)
    def sort_key(kv):
        (last, first), sc = kv
        tokens = len((last + " " + first).split())
        total_len = len(last.replace(" ","")) + len(first.replace(" ",""))
        return (-sc, tokens, total_len)
    return sorted(best.items(), key=sort_key)[0][0]

# --------------------------- Renaming -----------------------------

def target_name(id8: str, last_first):
    last, first = last_first
    return f"{id8} - {last}, {first}"

def safe_rename(folder: Path, new_base: str):
    target = folder.with_name(new_base)
    if target.name == folder.name:
        print(f"[SAME] {folder.name}")
        return
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

# ---------------------------- Main --------------------------------

def derive_name_for_folder(folder: Path):
    # 1) Look at ALL file NAMES first
    file_name_cands = []
    for f in folder.iterdir():
        if f.is_file():
            file_name_cands.extend(candidates_from_filename(f.name))
    best_from_names = pick_best(file_name_cands)
    if best_from_names:
        return best_from_names

    # 2) Else look INSIDE ALL PDFs
    pdf_cands = []
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() == ".pdf":
            pdf_cands.extend(candidates_from_pdf(f))
    return pick_best(pdf_cands)

def main():
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
