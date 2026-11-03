"""Snapshots on disk, academic categories, exports and new-hire diffs.

Data folder layout:
  <data_dir>/snapshots/YYYY-MM-DD.csv   phone book (all staff) as of that day, in the reduced
                                        public format (see crawler.SNAPSHOT_COLUMNS)
  <data_dir>/output/                    every Excel/CSV the command line produces
"""
import re
from pathlib import Path

import pandas as pd

from ku_staff.crawler import SNAPSHOT_COLUMNS
from ku_staff.matching import UserError, autosize, email_hash, style_header, unique_path  # noqa: F401

# Academic categories, matched in order against the job title (first match wins).
# "core" categories are employed academic staff; the others are affiliated people
# (guests, emeriti, student teaching jobs), only included when asked for.
CATEGORIES = [
    ("Emeritus", False, r"emerit"),
    ("Guest / visiting / external", False, r"guest|visit|external|affiliate|international researcher"),
    ("Assistant Professor", True, r"assistant professor|tenure track"),
    ("Associate Professor", True, r"associate professor"),
    ("Professor", True, r"professor"),
    ("PhD", True, r"ph\.?\s?d"),
    ("Postdoc", True, r"post\.?\s?doc|marie curie fellow"),
    ("Researcher / Research assistant", True,
     r"researcher|research fellow|research assistant|research staff|research officer|scientist|forsker"),
    ("Lecturer", True, r"lecturer|academic teacher|lektor|docent"),
    ("Academic staff (unspecified)", True, r"academic staff|\bvip\b"),
    ("Teaching assistant / instructor", False,
     r"teaching assistant|teacher.?s assistant|instructor|demonstrator|postgraduate fellow"),
]
CATEGORIES = [(name, core, re.compile(rx, re.I)) for name, core, rx in CATEGORIES]
# Support and technical roles, unless the title also says professor (e.g. "PhD Administrator").
NON_ACADEMIC = re.compile(
    r"secretar|administrat|technician|coordinator|consultant|adviser|councel|laboratory", re.I)

# Which rows an export or diff keeps.
SCOPES = {
    "academic": "core academic staff",
    "academic_incl_affiliated": "academic staff incl. guests, emeriti and teaching assistants",
    "all_staff": "all staff",
}

# Kept in CSV/JSON (the matcher needs them) but left out of Excel files.
MACHINE_COLUMNS = ["email_sha256"]
EXPORT_COLUMNS = ["person_id", "first_name", "last_name", "category", "title_en", "department",
                  "subunit", "faculty", "research_profile", "title_da", "unit_en", "email_sha256"]


def categorize(title):
    """Job title -> (category, is_core). ("", False) for non-academic titles."""
    if NON_ACADEMIC.search(title) and "professor" not in title.lower():
        return "", False
    for name, core, rx in CATEGORIES:
        if rx.search(title):
            return name, core
    return "", False


def snapshot_dir(data_dir):
    return Path(data_dir) / "snapshots"


def output_dir(data_dir):
    return Path(data_dir) / "output"


def snapshots(data_dir):
    return sorted(snapshot_dir(data_dir).glob("????-??-??.csv"))


def snapshot_on_or_before(data_dir, date=None):
    snaps = snapshots(data_dir)
    if not snaps:
        raise UserError("No staff list downloaded yet. Download it first.")
    if date:
        earlier = [s for s in snaps if s.stem <= date]
        if not earlier:
            raise UserError(f"No download on or before {date}; the earliest one is {snaps[0].stem}.")
        return earlier[-1]
    return snaps[-1]


def load(path):
    return pd.read_csv(path, dtype=str, keep_default_na=False)


def to_public_snapshot(df):
    """A snapshot in the original full format (with plain emails, phones, addresses) ->
    the reduced public format. Snapshots already in that format pass through unchanged."""
    if "email_sha256" not in df:
        df = df.assign(email_sha256=df.get("email", pd.Series("", index=df.index)).map(email_hash))
    for col in SNAPSHOT_COLUMNS:
        if col not in df:
            df[col] = ""
    return df[SNAPSHOT_COLUMNS]


def import_snapshot(data_dir, path, date=None):
    """Copy a snapshot CSV (any format) into data_dir/snapshots/ in the public format."""
    path = Path(path)
    found = re.search(r"\d{4}-\d{2}-\d{2}", date or path.stem)
    if not found:
        raise UserError(f"Can't tell the snapshot date from {path.name!r}; pass it as YYYY-MM-DD.")
    date = found.group()
    out = snapshot_dir(data_dir) / f"{date}.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    to_public_snapshot(load(path)).to_csv(out, index=False)
    return out


def select(df, scope="academic"):
    """Add a category column and keep the rows in scope."""
    if scope not in SCOPES:
        raise ValueError(f"unknown scope {scope!r}")
    title = df["title_en"].where(df["title_en"] != "", df["title_da"])
    cats = title.map(categorize)
    df = df.assign(category=[c for c, _ in cats])
    if scope == "all_staff":
        return df
    keep = [bool(c) and (core or scope == "academic_incl_affiliated") for c, core in cats]
    return df[keep]


def autosize(ws, max_width=60):
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col[:500])
        ws.column_dimensions[col[0].column_letter].width = min(max(width, 8), max_width)


def unique_path(path):
    """path, or "name (2).ext", "name (3).ext"... if it exists. Excel files are never
    overwritten because people annotate them (colours, comments, notes)."""
    path = Path(path)
    i = 2
    while path.exists():
        path = path.with_name(f"{path.stem.rsplit(' (', 1)[0]} ({i}){path.suffix}")
        i += 1
    return path


def write_table(df, base_path, sheet_name="staff", xlsx=True):
    """Write base_path.csv (replaced if present) and, if xlsx, a new base_path.xlsx with
    filters and a frozen header (never overwriting one). Returns the .xlsx path, or the
    .csv path when xlsx=False."""
    base_path = Path(base_path)
    base_path.parent.mkdir(parents=True, exist_ok=True)
    csv = base_path.with_name(base_path.name + ".csv")
    df.to_csv(csv, index=False)
    if not xlsx:
        return csv
    df = df.drop(columns=MACHINE_COLUMNS, errors="ignore")  # meaningless to a reader
    xlsx = unique_path(base_path.with_name(base_path.name + ".xlsx"))
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet_name)
        ws = xw.sheets[sheet_name]
        style_header(ws)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autosize(ws)
    return xlsx


def load_snapshot(path):
    return to_public_snapshot(load(path))


def academic_table(data_dir, scope="academic"):
    """Latest snapshot filtered to scope. -> (DataFrame, snapshot date)."""
    snap = snapshot_on_or_before(data_dir)
    return select(load_snapshot(snap), scope)[EXPORT_COLUMNS], snap.stem


def new_staff_table(data_dir, since, scope="academic"):
    """People in scope in the newest snapshot but not in the one taken on/before `since`.

    The in-scope sets are compared, so a former master student who is now a PhD fellow
    counts as new; previous_title shows what they were before.
    -> (DataFrame, old snapshot date, new snapshot date).
    """
    old_snap, new_snap = snapshot_on_or_before(data_dir, since), snapshot_on_or_before(data_dir)
    if old_snap == new_snap:
        raise UserError(f"Only one download available for that period ({new_snap.stem}). "
                        "Download the staff list again on a later day to find new hires.")
    old_all = load_snapshot(old_snap)
    old, new = select(old_all, scope), select(load_snapshot(new_snap), scope)
    new_hires = new[~new["person_id"].isin(set(old["person_id"]))][EXPORT_COLUMNS].copy()
    new_hires["first_seen"] = new_snap.stem
    previous = old_all.groupby("person_id")["title_en"].agg(lambda t: "; ".join(sorted(set(t))))
    new_hires["previous_title"] = new_hires["person_id"].map(previous).fillna("")
    return new_hires, old_snap.stem, new_snap.stem


def export(data_dir, scope="academic", xlsx=True):
    """Latest snapshot, filtered to scope -> output/<scope>_<date>.xlsx/.csv.
    Returns (xlsx path, or csv path if xlsx=False; rows)."""
    df, date = academic_table(data_dir, scope)
    return write_table(df, output_dir(data_dir) / f"{scope}_{date}", xlsx=xlsx), len(df)


def diff(data_dir, since, scope="academic", xlsx=True):
    """New staff since `since` -> output/new_<scope>_<old>_to_<new>.xlsx/.csv.
    Returns (xlsx path, or csv path if xlsx=False; rows; old date; new date)."""
    df, old, new = new_staff_table(data_dir, since, scope)
    path = write_table(df, output_dir(data_dir) / f"new_{scope}_{old}_to_{new}", "new_staff", xlsx=xlsx)
    return path, len(df), old, new
