"""Build the website's data from the snapshots.

Writes into <site>/ (everything here is generated; never edit it by hand):
  data/index.json                     what exists: snapshot dates, counts, file names
  data/academic-latest.json           academic staff (incl. affiliated) in the latest snapshot
  data/new-since-<date>.json          academic staff new since each earlier snapshot
  data/files/*.xlsx                   the same lists as Excel downloads
  py/matching.py                      the matching module, run in the browser by Pyodide
"""
import datetime as dt
import json
import shutil
from pathlib import Path

from ku_staff import __version__, staff

# Fields published per person (the site's tables, and the matcher in the browser).
PUBLIC_FIELDS = ["person_id", "first_name", "last_name", "category", "core", "title_en",
                 "department", "subunit", "faculty", "research_profile", "email_sha256"]
CORE = {name for name, core, _ in staff.CATEGORIES if core}


def table_json(df, extra=()):
    df = df.assign(core=df["category"].isin(CORE))
    cols = PUBLIC_FIELDS + list(extra)
    return {"columns": cols, "rows": df[cols].values.tolist()}


def write_json(path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def publish(data_dir, site_dir, log=print):
    site_dir = Path(site_dir)
    data = site_dir / "data"
    shutil.rmtree(data, ignore_errors=True)
    (data / "files").mkdir(parents=True)
    (site_dir / "py").mkdir(exist_ok=True)
    shutil.copy(Path(__file__).with_name("matching.py"), site_dir / "py" / "matching.py")

    snaps = [s.stem for s in staff.snapshots(data_dir)]
    index = {"generated": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
             "version": __version__, "snapshots": snaps, "latest": snaps[-1] if snaps else None,
             "downloads": {}, "new_since": {}}
    if not snaps:
        write_json(data / "index.json", index)
        log("No snapshots yet: published an empty site.")
        return index

    latest = snaps[-1]
    for scope in ("academic", "academic_incl_affiliated"):
        df, _ = staff.academic_table(data_dir, scope)
        xlsx = staff.write_table(df, data / "files" / f"{scope}_{latest}")
        xlsx.with_name(xlsx.stem + ".csv").unlink()
        index["downloads"][scope] = f"data/files/{xlsx.name}"
        index["count_" + scope] = len(df)
    write_json(data / "academic-latest.json", table_json(df))  # incl. affiliated; `core` flags the rest

    for old in snaps[:-1]:
        entry = {}
        for scope in ("academic", "academic_incl_affiliated"):
            df, _, _ = staff.new_staff_table(data_dir, old, scope)
            xlsx = staff.write_table(df, data / "files" / f"new_{scope}_{old}_to_{latest}", "new_staff")
            xlsx.with_name(xlsx.stem + ".csv").unlink()
            entry["download_" + scope] = f"data/files/{xlsx.name}"
            entry["count_" + scope] = len(df)
        entry["json"] = f"data/new-since-{old}.json"
        write_json(data / f"new-since-{old}.json", table_json(df, ["first_seen", "previous_title"]))
        index["new_since"][old] = entry

    write_json(data / "index.json", index)
    log(f"Published {len(snaps)} snapshot(s), latest {latest}: "
        f"{index['count_academic']:,} academic staff, {len(snaps) - 1} new-staff comparison(s).")
    return index
