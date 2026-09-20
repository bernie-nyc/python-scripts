# ============================================================================
# missing_all_analysis.py
#
# WHAT THIS PROGRAM DOES (in plain words)
#
# Every year, each student gets grades for each class they took:
#   - an S1 grade (first half of the year)
#   - an S2 grade (second half of the year)
#   - an ALL grade (the whole-year grade)
#
# Some classes only last one semester. For those, the ALL grade is
# supposed to be the same as the one semester grade. But in our old
# grade files, some of those one-semester classes are missing their
# ALL row entirely.
#
# This program:
#   1. Opens the four old grade files (one per school year).
#   2. Finds every student + class combo that has NO "ALL" row.
#   3. Keeps only Upper School classes (drops Middle School).
#   4. Figures out what grade level (9th, 10th, etc.) each student was
#      in that year, using the student and alumni list.
#   5. Writes two files:
#        - a report of every missing row and how we figured out the grade level
#        - a draft "numeric grades" import file that Veracross can load
#
# HOW TO RUN IT
#   Put the CSV files in the folder named in the U variable below, then run:
#       python3 missing_all_analysis.py
# ============================================================================


# ---- Step 0: Load the tools we need ---------------------------------------
# "import" means "go get a toolbox so I can use its tools."
#   pandas  = a toolbox for working with spreadsheets/tables. We nickname it "pd".
#   glob    = a tool for finding files whose names match a pattern (like *.csv).
#   re      = a tool for finding patterns inside text (like "find the two digits
#             inside the quotes"). "re" stands for "regular expressions".
import pandas as pd
import glob
import re


# ---- Step 1: Say where the files live -------------------------------------
# U is just a nickname for the folder path so we don't retype it every time.
U = '/mnt/user-data/uploads/'

# Where we will save the results.
OUT = '/mnt/user-data/outputs/'


# ---- Step 2: Open the student and alumni list -----------------------------
# pd.read_csv opens a CSV file and turns it into a table (pandas calls a
# table a "DataFrame"). dtype=str means "treat every cell as plain text",
# so an ID like 100050 stays exactly 100050 and never gets turned into a
# number that might lose leading zeros or show up as 100050.0.
st = pd.read_csv(
    U + 'Students_and_Former_Stuudents_by_Enrollment_Status_and_Year__1908_records__2026-09-20.csv',
    dtype=str
)

# Some students have nothing in the "Roles" column. An empty cell shows up
# as a special "nothing" value called NaN, which makes text searches crash.
# fillna replaces those blanks with an empty piece of text ('') instead.
st = st.fillna({'Roles': ''})


# ---- Step 3: Figure out each person's graduation year ---------------------
# If we know the year someone graduates (or will graduate), we can work out
# what grade they were in during ANY year. Example:
#   graduates in 2025  ->  in school year 24-25 they were in 12th grade
#                      ->  in school year 21-22 they were in 9th grade
#
# The "Roles" column in the student list tells us this in a few different
# ways, so this function checks each way in order and uses the first one
# that works. A "function" is a small named recipe we can reuse.
def grad_year(r):
    # r is one row (one person) from the student list.
    roles = r['Roles']

    # Way 1: Roles looks like "Alum '25" or "Frmr Student 25-26 ('29)".
    # The pattern below means: an apostrophe, then two digits, then maybe
    # a closing parenthesis. re.search looks for that pattern in the text.
    m = re.search(r"'(\d\d)\)?", roles)
    if m:
        # m.group(1) is the two digits we found, like "25".
        # 2000 + 25 = 2025. We return the year AND a note saying where it came from.
        return 2000 + int(m.group(1)), 'Roles'

    # Way 2: Roles looks like "Student (11-D)" - a current student in 11th grade.
    # Today is school year 2026-27, so an 11th grader graduates in 2027 + (12 - 11) = 2028.
    m = re.search(r"Student \((\d+)-", roles)
    if m:
        return 2027 + (12 - int(m.group(1))), 'Roles current grade'

    # Way 3: Use the "Current Grade" or "Last Grade Enrolled At" column together
    # with "Last Year Enrolled". Example: last grade = "Grade 9",
    # last year = "2025-2026"  ->  graduates 2026 + (12 - 9) = 2029.
    for gcol, ycol, src in (
        ('Current Grade', 'Last Year Enrolled', 'Current Grade'),
        ('Last Grade Enrolled At', 'Last Year Enrolled', 'Last Grade Enrolled At'),
    ):
        g, y = r[gcol], r[ycol]
        # pd.notna means "this cell is NOT blank".
        if pd.notna(g) and pd.notna(y):
            # y[-4:] grabs the last 4 characters, e.g. "2026" from "2025-2026".
            # g.split()[-1] grabs the last word, e.g. "9" from "Grade 9".
            return int(y[-4:]) + (12 - int(g.split()[-1])), src

    # If none of the ways worked, we honestly say "we don't know".
    return None, None


# Run the grad_year recipe on every row of the student list.
# apply(..., axis=1) means "do this once per row". The result is two new
# columns: grad_year and grade_src (where the answer came from).
st[['grad_year', 'grade_src']] = st.apply(lambda r: pd.Series(grad_year(r)), axis=1)

# Make Person ID the "index" - like the label on a filing-cabinet drawer -
# so later we can look someone up instantly by their ID.
st = st.set_index('Person ID')


# ---- Step 4: Go through each year's grade file ----------------------------
# out is an empty list. We will add one entry to it for every missing ALL row.
out = []

# glob finds every file in the folder with "Archival" in its name.
# sorted puts them in year order (21-22, 22-23, ...).
for f in sorted(glob.glob(U + '*Archival*.csv')):

    # Open this year's grade file as a table.
    df = pd.read_csv(f, dtype=str)

    # Every row in the file has the same School Year, so read it off the
    # first row. It looks like "21-22".
    sy = df['School Year'].iloc[0]

    # The ending year as a full number: "21-22" -> 2022.
    # sy[-2:] is the last two characters ("22"); 2000 + 22 = 2022.
    end_year = 2000 + int(sy[-2:])

    # Group the rows by student + class. For each group, join together the
    # grading periods it has, e.g. "ALL+S1+S2" or just "S1".
    # This gives us one line per student-class combo showing which periods exist.
    g = (
        df.groupby(['Person ID', 'Class ID'])['Grading Period']
          .apply(lambda s: '+'.join(sorted(s)))
          .reset_index(name='periods_present')
    )

    # Keep only the combos whose period list does NOT contain "ALL".
    # The ~ symbol means "not".
    noall = g[~g.periods_present.str.contains('ALL')]

    # Go back to the full grade table and keep only the rows that belong to
    # those missing-ALL combos. "merge" matches rows by Person ID + Class ID.
    sub = df.merge(noall, on=['Person ID', 'Class ID'])

    # Now look at each of those rows one at a time.
    for _, r in sub.iterrows():
        pid = r['Person ID']

        # Look the student up in the student list. If they are not there,
        # s becomes None (nothing).
        s = st.loc[pid] if pid in st.index else None

        # Their graduation year, if we know it.
        gy = None if s is None else s['grad_year']

        # Convert graduation year into the grade they were in THIS year.
        # Example: graduates 2025, this year ends 2022 -> 12 - (2025 - 2022) = 9th grade.
        gl = None if pd.isna(gy) else int(12 - (gy - end_year))

        # Save everything we know about this row into a dictionary
        # (a dictionary is a list of labeled facts, like "course: Algebra I").
        out.append(dict(
            school_year=sy,
            person_id=pid,
            student=r['Student'],
            school_level=r['School Level'],
            class_id=r['Class ID'],
            internal_class_id=r['Course ID (Internal)'],
            course=r['Course'],
            periods_present=r['periods_present'],
            semester=r['Grading Period'],
            semester_period_id=r['GRADING PERIOD: ID'],
            status=r['Status'],
            posted_grade=r['Grade: Posted'],
            other_numeric=r['Other Numeric Grade'],
            exam_grade=r['Exam Grade'],
            in_student_csv=s is not None,
            name_match=(s is not None and s['Full Name'] == r['Student']),
            grad_year=gy,
            grade_level_extrapolated=gl,
            grade_source=None if s is None else s['grade_src'],
        ))


# ---- Step 5: Turn the collected rows into one table -----------------------
res = pd.DataFrame(out)

# Keep only Upper School classes. Middle School rows are thrown away here.
# reset_index just renumbers the rows 0, 1, 2, ... after the filter.
res = res[res.school_level == 'Upper School'].reset_index(drop=True)


# ---- Step 6: Fill in grade level from the course name, as a backup --------
# Some course names start with a grade number, like "08 English".
# If we still don't know the grade level, that number is a good clue.
# The pattern means: at the start of the text, maybe a 0, then 1 or 2 digits,
# then a space. extract pulls out just the digits.
cg = res.course.str.extract(r'^0?(\d{1,2}) ')[0]

# fb (fallback) is True for rows where grade level is blank AND the course
# name gave us a number.
fb = res.grade_level_extrapolated.isna() & cg.notna()

# For those rows, fill in the grade level and note where it came from.
res.loc[fb, 'grade_level_extrapolated'] = cg[fb].astype(int)
res.loc[fb, 'grade_source'] = 'Course name prefix'

# Save the full report. index=False means "don't add a column of row numbers".
res.to_csv(OUT + 'missing_ALL_rows_all_years.csv', index=False)


# ---- Step 7: Print a summary to the screen --------------------------------
print('== Enrollments (student x class) lacking an ALL row, by year ==')
# groupby + agg = "for each school year, count these things".
print(res.groupby('school_year').agg(
    enrollments=('class_id', 'size'),
    students=('person_id', 'nunique'),
    S1_only=('semester', lambda s: (s == 'S1').sum()),
    S2_only=('semester', lambda s: (s == 'S2').sum()),
    pass_status=('status', lambda s: (s == 'Pass').sum()),
    grade_level_resolved=('grade_level_extrapolated', lambda s: s.notna().sum()),
    not_in_student_csv=('in_student_csv', lambda s: (~s).sum()),
).to_string())
print()
print('grade level values:', res.grade_level_extrapolated.value_counts(dropna=False).sort_index().to_dict())
print('grade source:', res.grade_source.value_counts(dropna=False).to_dict())
print()

print('== Person IDs lacking ALL rows, by year ==')
for sy, grp in res.groupby('school_year'):
    print(sy, sorted(grp.person_id.unique()))
print()

print('== Unresolved grade level (person, year) ==')
u = res[res.grade_level_extrapolated.isna()][
    ['school_year', 'person_id', 'student', 'school_level', 'in_student_csv']
].drop_duplicates()
print(u.to_string(index=False))


# ---- Step 8: Build the draft import file ----------------------------------
# This makes a new table shaped exactly like the Veracross "numeric grades"
# template: same column names, same order. One row per missing ALL grade.
#
# Rules we follow:
#   - grading_period is always 50 (that is Veracross's ID for "ALL").
#   - school_year is the START year as a 4-digit number: "21-22" -> 2021.
#   - posted_grade is the one semester grade, because a one-semester class's
#     whole-year grade IS its semester grade.
#   - If the class was Pass/Fail, leave posted_grade blank and put "P" in status.
#   - Columns we have nothing for are left as empty text ('').
imp = pd.DataFrame({
    'person_id': res.person_id,
    'person_reference_type': '',
    'person_reference_value': '',
    'internal_class_id': res.internal_class_id,
    'class_id': res.class_id,
    'school_year': res.school_year.map(lambda s: 2000 + int(s[:2])),
    # Int64 (capital I) is a number type that allows blanks, so unknown
    # grade levels stay empty instead of crashing.
    'grade_level': res.grade_level_extrapolated.astype('Int64'),
    'grading_period': 50,
    'assignment_posted_grade': '',
    'exam_grade': '',
    'posted_grade': res.apply(lambda r: '' if r.status == 'Pass' else r.posted_grade, axis=1),
    'status': res.status.map(lambda s: 'P' if s == 'Pass' else ''),
    'other_grade_1': '',
    'other_grade_2': '',
    'comments': '',
})

imp.to_csv(OUT + 'numeric_grades_ALL_backfill_DRAFT.csv', index=False)
print('draft import rows', len(imp), 'blank grade_level', imp.grade_level.isna().sum())
