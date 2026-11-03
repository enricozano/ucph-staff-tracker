"""Offline tests (no network). Run with: python -m pytest"""
from pathlib import Path

import pandas as pd
import pytest

from ku_staff import cli, matching, staff
from ku_staff.crawler import SNAPSHOT_COLUMNS, normalize, parse
from ku_staff.publish import publish


def snapshot_row(pid, first, last, title, email="", unit="Department of Biology, Faculty of Science"):
    return {"person_id": pid, "first_name": first, "last_name": last, "title_en": title,
            "title_da": "", "function_en": "", "unit_en": unit, "unit_da": "", "email": email,
            "phone": "", "mobile": "", "address": "", "website": "", "research_profile": "",
            "faculty": "Faculty of Science", "department": "Department of Biology", "subunit": ""}


def write_snapshot(data_dir, date, rows):
    d = staff.snapshot_dir(data_dir)
    d.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_csv(d / f"{date}.csv", index=False)


# ---- crawler ----------------------------------------------------------------
def test_parse_tolerates_raw_backslashes_and_control_characters():
    text = '{"root":{"employees":[{"ANSAT_WWW":"https://www.teol.ku.dk\\solitudes","X":"a\tb"}]}}'
    assert parse(text)[0]["ANSAT_WWW"] == "https://www.teol.ku.dk\\solitudes"


def test_normalize_splits_unit_and_extracts_person_id():
    row = normalize({
        "PERSON_FORNAVN": "Daniel Arthur", "PERSON_EFTERNAVN": "Blears",
        "STED_NAVN_SAMLET_ENG": "Department of Biology, Ecology and Evolution, Faculty of Science",
        "FOTOURL": "https://www2.adm.ku.dk/selv/pls/prt_www40.hentindhold_cms?p_personid=707614",
        "ANSAT_ARB_EMAIL": "DBlears@sund.ku.dk", "ANSAT_ARB_TLF": "+4535323710",
    })
    assert row["person_id"] == "707614"
    # Public snapshots: no plain email or phone, only the email's hash.
    assert set(row) == set(SNAPSHOT_COLUMNS)
    assert row["email_sha256"] == matching.email_hash("dblears@sund.ku.dk")
    assert (row["department"], row["subunit"], row["faculty"]) == (
        "Department of Biology", "Ecology and Evolution", "Faculty of Science")


# ---- categories ---------------------------------------------------------------
@pytest.mark.parametrize("title,expected", [
    ("Postdoc", ("Postdoc", True)),
    ("PhD Fellow", ("PhD", True)),
    ("Assistant Professor - Tenure Track", ("Assistant Professor", True)),
    ("Clinical Associate Professor", ("Associate Professor", True)),
    ("Professor, Head of PhD programme", ("Professor", True)),
    ("Head of Administration, Associate Professor", ("Associate Professor", True)),
    ("Guest Researcher", ("Guest / visiting / external", False)),
    ("Professor Emeritus", ("Emeritus", False)),
    ("Teaching Assistant", ("Teaching assistant / instructor", False)),
    ("Academic Officer", ("", False)),
    ("PhD Administrator", ("", False)),
    ("Biomedical Laboratory Scientist", ("", False)),
    ("Research Consultant", ("", False)),
])
def test_categorize(title, expected):
    assert staff.categorize(title) == expected


# ---- export / diff ----------------------------------------------------------
def test_diff_finds_new_people_and_role_changes(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [
        snapshot_row("1", "Ann", "Old", "Professor"),
        snapshot_row("2", "Bo", "Student", "Master Student"),
    ])
    write_snapshot(tmp_path, "2026-02-01", [
        snapshot_row("1", "Ann", "Old", "Professor"),
        snapshot_row("2", "Bo", "Student", "PhD Fellow"),
        snapshot_row("3", "Cy", "New", "Postdoc"),
        snapshot_row("4", "Di", "Admin", "Academic Officer"),
    ])
    xlsx, n, old, new = staff.diff(tmp_path, "2026-01-15")
    df = pd.read_csv(xlsx.with_name(xlsx.stem + ".csv"), dtype=str, keep_default_na=False)
    assert (n, old, new) == (2, "2026-01-01", "2026-02-01")
    assert dict(zip(df["person_id"], df["previous_title"])) == {"2": "Master Student", "3": ""}


def test_diff_needs_two_snapshots(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [snapshot_row("1", "Ann", "Old", "Professor")])
    with pytest.raises(staff.UserError):
        staff.diff(tmp_path, "2026-01-01")


def test_export_scopes(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [
        snapshot_row("1", "Ann", "A", "Professor"),
        snapshot_row("2", "Bo", "B", "Guest Researcher"),
        snapshot_row("3", "Cy", "C", "IT Officer"),
    ])
    assert staff.export(tmp_path, "academic")[1] == 1
    assert staff.export(tmp_path, "academic_incl_affiliated")[1] == 2
    assert staff.export(tmp_path, "all_staff")[1] == 3


# ---- matching -----------------------------------------------------------------
def test_canon_collapses_danish_spellings_and_accents():
    assert matching.canon("Søren Jørgensen") == matching.canon("Soeren Joergensen") == ["soren", "jorgensen"]
    assert matching.canon("Åse Rodríguez-Otero") == ["ase", "rodriguez", "otero"]


@pytest.mark.parametrize("a,b,expected", [
    ("rune aabenhus", "rune aabenhus", "full name"),
    ("rune aabenhus", "rune munck aabenhus", "partial name"),
    ("syed zeeshan ahmad", "syed salman ahmad", None),
])
def test_compare_names(a, b, expected):
    assert matching.compare_names(a.split(), b.split()) == expected


def test_match_end_to_end(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [
        snapshot_row("1", "Giulia", "Zampirolo", "Postdoc", "gz@sund.ku.dk"),
        snapshot_row("2", "Rune Munck", "Aabenhus", "Associate Professor", "runa@sund.ku.dk"),
        snapshot_row("3", "Syed Salman", "Ahmad", "Part-time Lecturer"),
        snapshot_row("4", "Edda", "Santello", "PhD Fellow", "edda@bric.ku.dk"),
        snapshot_row("5", "Katarina", "Dragicevic", "Postdoc", "kd@bric.ku.dk"),
    ])
    ku_csv, _ = staff.export(tmp_path, xlsx=False)
    contacts = tmp_path / "contacts.csv"
    pd.DataFrame({
        "First Name": ["Giulia", "Rune", "Syed Zeeshan", "Edda", "Katarina", "Nobody"],
        "Last Name": ["Zampirolo", "Aabenhus", "Ahmad", "Santello", "Dragicevic", "Atall"],
        "Email": ["g@gmail.com", "r@example.org", "s@gmail.com", "x@alumni.ku.dk",
                  "KD@bric.ku.dk", "n@x.com"],
    }).to_csv(contacts, index=False)
    out, stats = matching.run(contacts, ku_csv, tmp_path / "matches.xlsx", log=lambda m: None)
    new = pd.read_excel(out, sheet_name="possible_new_hires", dtype=str)
    already = pd.read_excel(out, sheet_name="already_ucph_at_signup", dtype=str)
    assert set(new["db: Last Name"]) == {"Zampirolo", "Aabenhus"}
    assert set(already["db: Last Name"]) == {"Santello", "Dragicevic"}
    assert stats["already_at_ucph"] == 2
    # Names side by side right after the confidence, as they appear in each source.
    assert list(new.columns[:3]) == ["confidence", "name_in_our_database", "name_at_ucph"]
    row = new.set_index("name_in_our_database").loc["Rune Aabenhus"]
    assert row["name_at_ucph"] == "Rune Munck Aabenhus"


def test_match_names_link_to_a_web_search(tmp_path):
    import openpyxl

    test_match_end_to_end(tmp_path)
    out = tmp_path / "matches.xlsx"
    ws = openpyxl.load_workbook(out)["possible_new_hires"]
    headers = [c.value for c in ws[1]]
    ours, ucph = headers.index("name_in_our_database") + 1, headers.index("name_at_ucph") + 1
    links = {ws.cell(r, ucph).value: ws.cell(r, ucph).hyperlink.target for r in range(2, ws.max_row + 1)}
    assert links["Giulia Zampirolo"] == (
        "https://www.google.com/search?q=%22Giulia+Zampirolo%22+University+of+Copenhagen")
    assert ws.cell(2, ours).hyperlink.target.startswith("https://www.google.com/search?q=%22")
    assert ws.freeze_panes == "D2"
    assert ws.cell(2, 1).fill.fgColor.rgb.endswith("C6EFCE")  # High = green


def test_excel_files_are_never_overwritten(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [snapshot_row("1", "Ann", "A", "Professor")])
    first, _ = staff.export(tmp_path)
    second, _ = staff.export(tmp_path)
    third, _ = staff.export(tmp_path)
    assert [p.name for p in (first, second, third)] == [
        "academic_2026-01-01.xlsx", "academic_2026-01-01 (2).xlsx", "academic_2026-01-01 (3).xlsx"]
    csv, _ = staff.export(tmp_path, xlsx=False)
    assert csv.name == "academic_2026-01-01.csv"


def test_import_snapshot_converts_old_full_format(tmp_path):
    old = tmp_path / "2026-10-23.csv"
    pd.DataFrame([{**snapshot_row("1", "Ann", "A", "Professor", "ann@ku.dk"),
                   "phone": "+45 1234", "address": "Nørregade 10"}]).to_csv(old, index=False)
    out = staff.import_snapshot(tmp_path / "data", old)
    df = pd.read_csv(out, dtype=str, keep_default_na=False)
    assert out.name == "2026-10-23.csv" and list(df.columns) == SNAPSHOT_COLUMNS
    assert df.loc[0, "email_sha256"] == matching.email_hash("ann@ku.dk")


def test_cli_match_uses_latest_snapshot_by_default(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [snapshot_row("1", "Giulia", "Zampirolo", "Postdoc")])
    contacts = tmp_path / "contacts.csv"
    contacts.write_text("Name;Email\nGiulia Zampirolo;g@gmail.com\n")
    cli.main(["--dir", str(tmp_path), "match", str(contacts)])
    out = next((tmp_path / "output").glob("matches_contacts_vs_academic_2026-01-01*.xlsx"))
    assert pd.read_excel(out, sheet_name="possible_new_hires")["name_at_ucph"].tolist() == ["Giulia Zampirolo"]


def test_publish_builds_site_data(tmp_path):
    import json

    write_snapshot(tmp_path, "2026-01-01", [snapshot_row("1", "Ann", "Old", "Professor", "ann@ku.dk")])
    write_snapshot(tmp_path, "2026-02-01", [
        snapshot_row("1", "Ann", "Old", "Professor", "ann@ku.dk"),
        snapshot_row("2", "Cy", "New", "Postdoc"),
        snapshot_row("3", "Gus", "Guest", "Guest Researcher"),
    ])
    site = tmp_path / "site"
    index = publish(tmp_path, site, log=lambda m: None)
    assert index["latest"] == "2026-02-01" and index["snapshots"] == ["2026-01-01", "2026-02-01"]
    assert index["count_academic"] == 2 and index["count_academic_incl_affiliated"] == 3
    entry = index["new_since"]["2026-01-01"]
    assert (entry["count_academic"], entry["count_academic_incl_affiliated"]) == (1, 2)
    for f in [*index["downloads"].values(), entry["download_academic"], entry["json"]]:
        assert (site / f).is_file(), f
    latest = json.loads((site / "data" / "academic-latest.json").read_text())
    assert "email" not in latest["columns"] and len(latest["rows"]) == 3
    assert (site / "py" / "matching.py").read_text() == Path(matching.__file__).read_text()


def test_excel_exports_leave_out_the_email_hash(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [snapshot_row("1", "Ann", "A", "Professor", "ann@ku.dk")])
    xlsx, _ = staff.export(tmp_path)
    csv, _ = staff.export(tmp_path, xlsx=False)
    assert "email_sha256" not in pd.read_excel(xlsx).columns
    assert "email_sha256" in pd.read_csv(csv).columns
