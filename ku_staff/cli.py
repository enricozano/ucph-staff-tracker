"""Command line: ku-staff [--dir FOLDER] {crawl,export,diff,match,publish,import-snapshot}."""
import argparse
import re
import sys
from pathlib import Path

from ku_staff import __version__, matching, staff
from ku_staff.crawler import CrawlError, Crawler

DESCRIPTION = """\
UCPH Staff Tracker: snapshots of the University of Copenhagen phone book.

The phone book has no hire dates, so new hires are found by comparing dated
snapshots: anyone in a newer snapshot but not in an older one is new.

Commands:
  crawl                     download the phone book into snapshots/YYYY-MM-DD.csv
  export                    export the latest snapshot (core academic staff) to xlsx + csv
  diff --since YYYY-MM-DD   academic staff who are new since the snapshot taken on/before that date
  match CONTACTS.xlsx       find contacts in your database (xlsx/csv) who now work at UCPH
  publish SITE_DIR          build the website's data from the snapshots
  import-snapshot FILE      add an existing snapshot CSV (any format) to snapshots/

export and diff accept --scope academic | academic_incl_affiliated | all_staff.
"""


def progress_printer():
    tty = sys.stderr.isatty()
    last = [-1]

    def progress(fraction, text):
        if tty:
            print(f"\r{text:<90}", end="", file=sys.stderr, flush=True)
        elif int(fraction * 20) != last[0]:  # every 5% when not on a terminal
            last[0] = int(fraction * 20)
            print(text, file=sys.stderr, flush=True)

    return progress


def cmd_crawl(args):
    log = lambda msg: print(f"\n{msg}" if sys.stderr.isatty() else msg, file=sys.stderr)
    Crawler(log=log, progress=progress_printer()).crawl(staff.snapshot_dir(args.dir))


def cmd_export(args):
    xlsx, n = staff.export(args.dir, args.scope)
    print(f"Wrote {n:,} rows to {xlsx} and .csv")


def cmd_diff(args):
    xlsx, n, old, new = staff.diff(args.dir, args.since, args.scope)
    print(f"Comparing {old} -> {new}: {n:,} new entries. Wrote {xlsx} and .csv")


def cmd_match(args):
    ku_path = args.ku
    if ku_path is None:
        # Fresh export of the latest snapshot (CSV only).
        ku_path, _ = staff.export(args.dir, "academic", xlsx=False)
    elif not {"category"} <= set(staff.load(ku_path).columns):
        # A raw snapshot: categorize it first.
        df = staff.select(staff.load_snapshot(ku_path), "academic")
        ku_path = staff.output_dir(args.dir) / f"academic_{Path(ku_path).stem}.csv"
        ku_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(ku_path, index=False)
    stem = re.sub(r"[^\w-]+", "_", args.contacts.stem).strip("_")
    output = args.output or staff.output_dir(args.dir) / f"matches_{stem}_vs_{Path(ku_path).stem}.xlsx"
    columns = {role: getattr(args, f"{role}_col") for role in matching.COLUMN_GUESSES}
    matching.run(args.contacts, ku_path, output, args.sheet, {k: v for k, v in columns.items() if v})


def cmd_publish(args):
    from ku_staff.publish import publish

    publish(args.dir, args.site)


def cmd_import(args):
    out = staff.import_snapshot(args.dir, args.file, args.date)
    print(f"Imported {args.file} as {out}")


COLUMN_HELP = {"name": "full name (when there are no separate first/last columns)",
               "first": "first name", "last": "last name", "company": "company / university",
               "date": "signup / create date", "record_id": "CRM record ID", "phone": "phone number",
               "position": "position / job title", "field": "field / job category"}


def build_parser():
    p = argparse.ArgumentParser(prog="ku-staff", description=DESCRIPTION,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    p.add_argument("--dir", type=Path, default=Path.cwd(),
                   help="folder holding snapshots/ and output/ (default: current folder)")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("crawl", help="download the phone book").set_defaults(func=cmd_crawl)

    scope = dict(choices=list(staff.SCOPES), default="academic",
                 help="which staff to include (default: academic)")
    e = sub.add_parser("export", help="export academic staff from the latest snapshot")
    e.add_argument("--scope", **scope)
    e.set_defaults(func=cmd_export)

    d = sub.add_parser("diff", help="academic staff who are new since a date")
    d.add_argument("--since", required=True, help="YYYY-MM-DD baseline date")
    d.add_argument("--scope", **scope)
    d.set_defaults(func=cmd_diff)

    m = sub.add_parser("match", help="find contacts from your database who now work at UCPH",
                       description="Match a contact database against UCPH staff. "
                                   "See the 'method' sheet of the output for the rules.")
    m.add_argument("contacts", type=Path, help="your contacts file (.xlsx or .csv)")
    m.add_argument("--ku", type=Path,
                   help="UCPH file to match against: an export, a diff result or a snapshot "
                        "(default: academic staff in the latest snapshot)")
    m.add_argument("-o", "--output", type=Path, help="output .xlsx (default: output/matches_...xlsx)")
    m.add_argument("--sheet", help="sheet name in the contacts workbook (default: first sheet)")
    for role in matching.COLUMN_GUESSES:  # one option per column role, e.g. --record-id-col
        m.add_argument(f"--{role.replace('_', '-')}-col", dest=f"{role}_col",
                       help=f"column with the {COLUMN_HELP.get(role, role)} (default: guessed)")
    m.set_defaults(func=cmd_match)

    pub = sub.add_parser("publish", help="build the website's data from the snapshots")
    pub.add_argument("site", type=Path, help="the website folder (site/)")
    pub.set_defaults(func=cmd_publish)

    imp = sub.add_parser("import-snapshot", help="add an existing snapshot CSV to snapshots/")
    imp.add_argument("file", type=Path)
    imp.add_argument("--date", help="snapshot date YYYY-MM-DD (default: from the file name)")
    imp.set_defaults(func=cmd_import)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        args.func(args)
    except (matching.UserError, CrawlError) as e:
        sys.exit(f"Error: {e}")
    except KeyboardInterrupt:
        sys.exit("\nStopped.")


if __name__ == "__main__":
    main()
