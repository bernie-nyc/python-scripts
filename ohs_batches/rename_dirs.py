#!/usr/bin/env python3
# Rename <Legacy8>  →  <PID6>_<First-Last>_<Legacy8>
# Strict: act only when a top-level folder named exactly <Legacy8> exists.
# Row-driven, duplicate CSV rows auto-skip after first consumption.
# Robust CSV decode + control-char stripping to avoid underscores artifact.
# No suffixing; if target exists, skip.
import argparse, csv, io, re, sys
from pathlib import Path

# ---------- Windows safety ----------
WIN_ILLEGAL = r'<>:"/\\|?*\x00-\x1F'
_ILLEGAL_RE = re.compile(f"[{re.escape(WIN_ILLEGAL)}]")
_TRAILING_RE = re.compile(r"[\. ]+$")
_RESERVED = {"CON","PRN","AUX","NUL", *{f"COM{i}" for i in range(1,10)}, *{f"LPT{i}" for i in range(1,10)}}

def sanitize_component(s: str) -> str:
    s = _ILLEGAL_RE.sub("_", s)
    s = _TRAILING_RE.sub("", s)
    if not s: s = "_"
    if s.split(".",1)[0].upper() in _RESERVED: s += "_"
    return s

# ---------- text cleaning ----------
_CTRL_RE = re.compile(r"[\x00-\x1F]+")  # strip ASCII control chars
def strip_ctrl(s: str) -> str:
    return _CTRL_RE.sub("", s or "")

def norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def digits_only(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def pid_six(s: str) -> str:
    # remove all non-digits first, then take first 6 run
    d = digits_only(s)
    m = re.search(r"\d{6}", d)
    return m.group(0) if m else ""

def first_dash_last(name: str) -> str:
    n = norm_ws(strip_ctrl(name))
    if ", " in n:
        last, first = n.split(", ", 1)
        return f"{first}-{last}"
    return n

# ---------- CSV decoding ----------
def decode_csv(csv_path: Path) -> io.StringIO:
    b = csv_path.read_bytes()
    if b.startswith((b"\xff\xfe", b"\xfe\xff")):
        t = b.decode("utf-16")
    elif b.count(b"\x00") > 0:
        try: t = b.decode("utf-16-le")
        except UnicodeDecodeError: t = b.decode("utf-16", errors="strict")
    else:
        t = b.decode("utf-8-sig")
    return io.StringIO(t)

def hdrmap(fields):
    m = {}
    for h in fields or []:
        key = re.sub(r"\s+", " ", (h or "")).strip().lower()
        m[key] = h
    return m

# ---------- iterate CSV rows with 1-based indices ----------
def iter_rows(csv_path: Path):
    f = decode_csv(csv_path)
    r = csv.DictReader(f)
    h = hdrmap(r.fieldnames)
    c_legacy = h.get("legacy person id")
    c_pid    = h.get("person id")
    c_name   = h.get("full name")
    if not (c_legacy and c_pid and c_name):
        print("ERROR: CSV needs headers: Person ID, Full Name, Legacy Person ID", file=sys.stderr)
        return
    row_idx = 1  # header
    for row in r:
        row_idx += 1
        legacy = digits_only(strip_ctrl(row.get(c_legacy, "")))
        if len(legacy) != 8:
            yield (row_idx, None, None, None); continue
        pid6 = pid_six(row.get(c_pid, ""))
        if len(pid6) != 6:
            yield (row_idx, legacy, None, None); continue
        name = first_dash_last(row.get(c_name, ""))
        yield (row_idx, legacy, pid6, name)

def main():
    ap = argparse.ArgumentParser(description="Rename <Legacy8> → <PID6>_<First-Last>_<Legacy8> with indexed skip report.")
    ap.add_argument("--csv", required=True, help="Path to Person Query CSV")
    ap.add_argument("--apply", action="store_true", help="Apply changes (default dry run)")
    args = ap.parse_args()

    root = Path.cwd()
    csv_path = Path(args.csv)
    if not csv_path.exists():
        print(f"ERROR: CSV not found: {csv_path}"); return

    print(f"Working folder: {root}")
    print(f"Mode: {'APPLY' if args.apply else 'DRY RUN (no changes)'}\n")

    # snapshot available exact legacy folders; consume as we go
    available = {p.name for p in root.iterdir() if p.is_dir() and re.fullmatch(r"\d{8}", p.name)}
    planned_names = set()  # to catch target collisions in dry-run

    plan = []  # (src, dst, row_index)
    rows_total = 0
    skip_invalid_legacy = []
    skip_invalid_pid = []
    skip_no_dir = []
    skip_target_exists = []

    for row_index, legacy, pid6, name in iter_rows(csv_path):
        rows_total += 1

        if legacy is None:
            skip_invalid_legacy.append(row_index); continue
        if pid6 is None:
            skip_invalid_pid.append(row_index); continue

        # exact match only
        if legacy not in available:
            skip_no_dir.append(row_index); continue

        dst_name = f"{pid6}_{sanitize_component(name)}_{legacy}"
        if (root / dst_name).exists() or dst_name in planned_names:
            skip_target_exists.append(row_index); continue

        src = root / legacy
        dst = root / dst_name
        plan.append((src, dst, row_index))
        planned_names.add(dst_name)
        available.remove(legacy)  # consume so later duplicates skip

        if args.apply:
            try:
                src.rename(dst)
            except Exception as e:
                print(f"ERROR (row {row_index}): {src.name} -> {dst.name}: {e}")
                planned_names.discard(dst_name)
                available.add(legacy)

    print(f"CSV rows read:              {rows_total}")
    print(f"Planned renames:            {len(plan)}")
    print(f"Skipped (invalid legacy):   {len(skip_invalid_legacy)} -> {skip_invalid_legacy}")
    print(f"Skipped (invalid personid): {len(skip_invalid_pid)}    -> {skip_invalid_pid}")
    print(f"Skipped (no matching dir):  {len(skip_no_dir)}         -> {skip_no_dir}")
    print(f"Skipped (target exists):    {len(skip_target_exists)}  -> {skip_target_exists}\n")

    for src, dst, idx in plan[:50]:
        print(f"row {idx}: {src.name} -> {dst.name}")
    if len(plan) > 50:
        print(f"...and {len(plan)-50} more")

    if not plan:
        print("\nNothing to rename."); return
    if args.apply:
        print("\nDone.")
    else:
        print("\nDry run complete. Re-run with --apply to commit.")

if __name__ == "__main__":
    main()
