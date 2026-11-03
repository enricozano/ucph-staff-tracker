"""Match a contact database (xlsx/csv) against UCPH staff to find contacts who now work at UCPH.

Matching, per contact:
  email         same work email as a UCPH staff entry (compared as SHA-256 hashes)
  full name     all names identical after normalization (case, accents, ø/oe, å/aa, æ/ae, hyphens)
  partial name  same first given name and last surname, and every name part of the shorter
                name appears in the longer one (a middle name or second surname left out).
                Names that contradict each other (different middle names) are not matched.

Confidence of name matches (sheet "possible_new_hires"):
  High    full name, unique on both sides, and the name is distinctive: a rare surname,
          3+ name parts, or Danish/UCPH evidence (.dk email, country Denmark, UCPH company)
  Medium  full name, unique on both sides, common surname; or a partial name, unique on
          both sides, with a rare surname or Danish/UCPH evidence
  Low     everything else: the name is shared by several people on either side, or a
          partial match on a common surname

A surname is rare when at most RARE_SURNAME people carry it across both files.
Click a name in either name column to search for that person on the web.

Contacts who were already at UCPH when they signed up go to the "already_ucph_at_signup"
sheet instead: an email match, or any ku.dk address in the database (including student
addresses such as @alumni.ku.dk).
"""
# This module must run both in normal Python and in the browser (Pyodide): it only uses
# the standard library and openpyxl. Don't import pandas or other ku_staff modules here.
import csv
import datetime as dt
import hashlib
import io
import re
import unicodedata
import warnings
from pathlib import Path
from urllib.parse import quote_plus

RARE_SURNAME = 3
TRANSLIT = str.maketrans({"ø": "o", "æ": "ae", "ß": "ss", "ł": "l", "đ": "d", "ı": "i"})
UCPH_COMPANY = re.compile(r"copenhagen|københavn|kobenhavn|ucph|\bku\b", re.I)
UCPH_EMAIL = re.compile(r"[@.]ku\.dk$", re.I)
CONFIDENCE_ORDER = {"High": 0, "Medium": 1, "Low": 2}
MATCH_ORDER = {"email": 0, "full name": 1, "partial name": 2}

NAME_OURS = "name_in_our_database"
NAME_UCPH = "name_at_ucph"
SEARCH_OURS, SEARCH_UCPH = "_search_ours", "_search_ucph"  # link targets, not written as columns
SEARCH_URL = "https://www.google.com/search?q="
CONFIDENCE_FILL = {"High": "C6EFCE", "Medium": "FFEB9C", "Low": "FCE4D6"}
HEADER_FILL = "901A1E"  # UCPH red

# Column names tried, case-insensitively, when no column is given.
COLUMN_GUESSES = {
    "first": ["first name", "firstname", "first_name", "given name", "fornavn", "nome"],
    "last": ["last name", "lastname", "last_name", "surname", "family name", "efternavn", "cognome"],
    "name": ["name", "full name", "fullname", "full_name", "navn", "nome completo"],
    "email": ["email", "e-mail", "email address", "e-mail address", "mail"],
    "company": ["company name", "company", "associated company", "organization", "organisation",
                "institution", "affiliation"],
    "country": ["country/region", "country", "land"],
    "date": ["create date", "created", "created at", "signup date", "sign up date", "date"],
}


class UserError(Exception):
    """A problem the user can fix (missing file, bad column name...). Shown as-is."""


# ---- names ----------------------------------------------------------------------
def canon(s):
    """Lowercase ASCII tokens; Danish spellings collapsed so ø/oe, å/aa, æ/ae compare equal."""
    s = unicodedata.normalize("NFKD", str(s).lower().translate(TRANSLIT))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("oe", "o").replace("aa", "a").replace("ae", "a")
    return re.findall(r"[a-z]+", s)


def name_key(first, last):
    """(first given name, last surname token), or None for initials / one-letter names."""
    if not first or not last or len(first[0]) < 2 or len(last[-1]) < 2:
        return None
    return first[0], last[-1]


def compare_names(a, b):
    """'full name' if identical, 'partial name' if one is the other minus some middle
    names / extra surnames, None if they contradict each other."""
    if a == b:
        return "full name"
    short, long = sorted((a, b), key=len)
    return "partial name" if set(short) <= set(long) else None


def email_hash(email):
    """SHA-256 of the lowercased email (how UCPH work emails are stored in snapshots)."""
    email = (email or "").strip().lower()
    return hashlib.sha256(email.encode()).hexdigest() if email else ""


def search_url(name, context=""):
    """Web search for the exact name, narrowed by an institution/company if known."""
    return SEARCH_URL + quote_plus(f'"{name}" {context}'.strip())


# ---- reading the contacts file ---------------------------------------------------
def cell_text(value):
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M" if value.time() != dt.time() else "%Y-%m-%d")
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_table(data, filename, sheet=None):
    """Contacts file (bytes) -> (headers, rows). Rows are lists of strings."""
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook

        try:
            with warnings.catch_warnings():  # "Workbook contains no default style" (HubSpot)
                warnings.simplefilter("ignore", UserWarning)
                wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        except Exception as e:
            raise UserError(f"Could not open {filename} as an Excel file ({e}).") from e
        if sheet and sheet not in wb.sheetnames:
            raise UserError(f"No sheet called {sheet!r}. Sheets: {', '.join(wb.sheetnames)}")
        ws = wb[sheet] if sheet else wb.worksheets[0]
        # Some exporters (HubSpot) write a wrong sheet size; read the actual cells instead.
        ws.reset_dimensions()
        rows = [[cell_text(v) for v in r] for r in ws.iter_rows(values_only=True)]
        wb.close()
    elif name.endswith((".csv", ".txt")):
        text = data.decode("utf-8-sig", errors="replace")
        try:
            dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        rows = [[c.strip() for c in r] for r in csv.reader(io.StringIO(text), dialect)]
    elif name.endswith(".xls"):
        raise UserError("This is an old-style .xls file. Open it in Excel, save it as .xlsx, "
                        "and try again.")
    else:
        raise UserError(f"{filename} is not an Excel (.xlsx) or CSV file.")
    rows = [r for r in rows if any(r)]
    if not rows:
        raise UserError(f"{filename} is empty.")
    headers, seen = [], {}
    for i, h in enumerate(rows[0]):
        h = h or f"Unnamed: {i}"
        if h in seen:  # duplicate headers (HubSpot has two "Create Date") get a suffix
            seen[h] += 1
            h = f"{h}.{seen[h]}"
        else:
            seen[h] = 0
        headers.append(h)
    body = [(r + [""] * len(headers))[:len(headers)] for r in rows[1:]]
    return headers, body


def find_column(headers, role, given):
    if given:
        if given not in headers:
            raise UserError(f"Column {given!r} not found in the contacts file. "
                            f"Its columns are: {', '.join(headers)}")
        return given
    lower = {h.strip().lower(): h for h in headers}
    return next((lower[g] for g in COLUMN_GUESSES[role] if g in lower), None)


def load_contacts(data, filename, sheet=None, columns=None, log=print):
    """Read the contacts file. columns maps a role (first, last, name, email, company,
    country, date) to a column name; roles left out are guessed."""
    headers, rows = read_table(data, filename, sheet)
    columns = columns or {}
    cols = {role: find_column(headers, role, columns.get(role)) for role in COLUMN_GUESSES}
    idx = {h: i for i, h in enumerate(headers)}
    get = lambda row, role: row[idx[cols[role]]] if cols[role] else ""
    if cols["first"] and cols["last"]:
        names = [(canon(get(r, "first")), canon(get(r, "last")), f'{get(r, "first")} {get(r, "last")}'.strip())
                 for r in rows]
    elif cols["name"]:
        # Single full-name column: first word is the given name, the rest the surname(s).
        names = [(canon(get(r, "name"))[:1], canon(get(r, "name"))[1:], get(r, "name")) for r in rows]
    else:
        raise UserError("Could not find the name columns in the contacts file. It needs a "
                        "'First Name' and a 'Last Name' column, or a single 'Name' column. "
                        f"Its columns are: {', '.join(headers)}")
    log("Contacts file columns used: " + ", ".join(f"{r}={c!r}" for r, c in cols.items() if c))
    contacts = []
    for r, (first, last, display) in zip(rows, names):
        email = get(r, "email")
        contacts.append({
            "first": first, "last": last, "name": display,
            "email": email.lower(), "email_original": email, "company": get(r, "company"),
            "country": get(r, "country"), "date": get(r, "date"),
            "row": dict(zip(headers, r)),
        })
    return headers, contacts


# ---- UCPH people -----------------------------------------------------------------
def people_from_records(records):
    """UCPH rows (export, diff result, or published JSON) -> one dict per person.
    Someone with several affiliations has one row each; titles and units are joined."""
    people = {}
    for rec in records:
        pid = rec.get("person_id") or f'{rec.get("first_name")} {rec.get("last_name")}'
        p = people.get(pid)
        if p is None:
            p = people[pid] = {
                "person_id": rec.get("person_id", ""),
                "first_name": rec.get("first_name", ""), "last_name": rec.get("last_name", ""),
                "email_sha256": rec.get("email_sha256") or email_hash(rec.get("email")),
                "research_profile": rec.get("research_profile", ""),
                "first_seen": rec.get("first_seen", ""), "previous_title": rec.get("previous_title", ""),
                "category": [], "title_en": [], "department": [], "faculty": [],
            }
        for c in ("category", "title_en", "department", "faculty"):
            v = rec.get(c) or ""
            if v and v not in p[c]:
                p[c].append(v)
    out = []
    for p in people.values():
        for c in ("category", "title_en", "department", "faculty"):
            p[c] = "; ".join(p[c])
        p["first_tokens"], p["last_tokens"] = canon(p["first_name"]), canon(p["last_name"])
        out.append(p)
    return out


def read_ucph_csv(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


# ---- matching --------------------------------------------------------------------
def match(people, contacts):
    """-> (possible new hires, already at UCPH at signup): lists of row dicts, sorted."""
    ku_names = [p["first_tokens"] + p["last_tokens"] for p in people]
    db_names = [c["first"] + c["last"] for c in contacts]
    surname_count = {}
    for last in [p["last_tokens"] for p in people] + [c["last"] for c in contacts]:
        if last:
            surname_count[last[-1]] = surname_count.get(last[-1], 0) + 1
    ku_index = {}
    for i, p in enumerate(people):
        if k := name_key(p["first_tokens"], p["last_tokens"]):
            ku_index.setdefault(k, []).append(i)
    ku_by_email = {p["email_sha256"]: i for i, p in enumerate(people) if p["email_sha256"]}

    pairs = {}  # (contact index, person index) -> match kind
    for h, c in enumerate(contacts):
        i = ku_by_email.get(email_hash(c["email"])) if c["email"] else None
        if i is not None:
            pairs[(h, i)] = "email"
        for i in ku_index.get(name_key(c["first"], c["last"]), ()):
            kind = compare_names(db_names[h], ku_names[i])
            if kind and (h, i) not in pairs:
                pairs[(h, i)] = kind
    # Candidates per side, to flag names several people share.
    n_ku, n_db = {}, {}
    for (h, i), kind in pairs.items():
        if kind != "email":
            n_ku[h] = n_ku.get(h, 0) + 1
            n_db[i] = n_db.get(i, 0) + 1

    rows = []
    for (h, i), kind in pairs.items():
        c, k = contacts[h], people[i]
        evidence = []
        if c["email"].endswith(".dk") and not UCPH_EMAIL.search(c["email"]):
            evidence.append(".dk email")
        if c["country"].lower() in ("denmark", "danmark"):
            evidence.append("country Denmark")
        if UCPH_COMPANY.search(c["company"]):
            evidence.append("company UCPH/Copenhagen")
        surname = c["last"][-1] if c["last"] else ""
        rare = surname_count.get(surname, 0) <= RARE_SURNAME
        unique = kind == "email" or (n_ku[h] == 1 and n_db[i] == 1)
        if kind == "email":
            confidence = "High"
        elif kind == "full name" and unique and (rare or len(db_names[h]) >= 3 or evidence):
            confidence = "High"
        elif unique and (kind == "full name" or rare or evidence):
            confidence = "Medium"
        else:
            confidence = "Low"
        ucph_name = f'{k["first_name"]} {k["last_name"]}'.strip()
        row = {
            "already_at_ucph_reason": "same email" if kind == "email"
            else "ku.dk email in database" if UCPH_EMAIL.search(c["email"]) else "",
            "confidence": confidence,
            # The two names side by side, each a link to a web search (see style_sheet()).
            NAME_OURS: c["name"],
            NAME_UCPH: ucph_name,
            "ucph_title": k["title_en"],
            "ucph_department": k["department"],
            "ucph_faculty": k["faculty"],
            "match_type": kind,
            "evidence": ", ".join(evidence),
            "email_in_our_database": c["email_original"],
            "ucph_research_profile": k["research_profile"],
            "company_in_our_database": c["company"],
            "country_in_our_database": c["country"],
            "signup_date_in_our_database": c["date"],
            "db_row": h + 2,  # row number as shown in Excel (header is row 1)
            "surname_frequency": surname_count.get(surname, 0),
            "ucph_candidates_for_this_contact": n_ku.get(h, 1),
            "contacts_matching_this_ucph_person": n_db.get(i, 1),
            "ucph_category": k["category"],
            "ucph_person_id": k["person_id"],
        }
        if k["first_seen"]:
            row["ucph_first_seen"] = k["first_seen"]
            row["ucph_previous_title"] = k["previous_title"]
        row.update({f"db: {col}": v for col, v in c["row"].items()})
        row[SEARCH_OURS] = search_url(c["name"], c["company"])
        row[SEARCH_UCPH] = search_url(ucph_name, "University of Copenhagen")
        rows.append(row)
    rows.sort(key=lambda r: (CONFIDENCE_ORDER[r["confidence"]], MATCH_ORDER[r["match_type"]],
                             r["surname_frequency"], r[NAME_UCPH].lower()))
    already = [r for r in rows if r["already_at_ucph_reason"]]
    candidates = [r for r in rows if not r["already_at_ucph_reason"]]
    for r in candidates:
        del r["already_at_ucph_reason"]
    return candidates, already


def stats_of(contacts, people, candidates, already):
    count = lambda c: sum(r["confidence"] == c for r in candidates)
    return {
        "contacts": len(contacts), "ucph_people": len(people), "pairs": len(candidates),
        **{c: count(c) for c in CONFIDENCE_ORDER},
        "contacts_matched": len({r["db_row"] for r in candidates}),
        "already_at_ucph": len(already),
    }


# ---- Excel output ----------------------------------------------------------------
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


def style_header(ws):
    from openpyxl.styles import Font, PatternFill

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)


def write_rows(ws, rows, columns):
    ws.append(columns)
    for r in rows:
        ws.append([r.get(c, "") for c in columns])


def style_matches_sheet(ws, rows, columns):
    """Header colours, confidence colours, clickable names, frozen name columns."""
    from openpyxl.styles import PatternFill

    style_header(ws)
    idx = {c: i + 1 for i, c in enumerate(columns)}
    for r, row in enumerate(rows, start=2):
        ws.cell(r, idx["confidence"]).fill = PatternFill("solid", fgColor=CONFIDENCE_FILL[row["confidence"]])
        for name_col, url_col in ((NAME_OURS, SEARCH_OURS), (NAME_UCPH, SEARCH_UCPH)):
            cell = ws.cell(r, idx[name_col])
            if cell.value:
                cell.hyperlink = row[url_col]
                cell.style = "Hyperlink"
        profile = ws.cell(r, idx["ucph_research_profile"])
        if profile.value:
            profile.hyperlink = profile.value
            profile.style = "Hyperlink"
    # Keep the confidence and both names visible while scrolling right.
    ws.freeze_panes = ws.cell(2, idx[NAME_UCPH] + 1).coordinate
    ws.auto_filter.ref = ws.dimensions


def build_workbook(candidates, already, summary):
    """-> openpyxl Workbook with the four sheets."""
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in (("possible_new_hires", candidates), ("already_ucph_at_signup", already)):
        ws = wb.create_sheet(name)
        columns = [c for c in (rows[0] if rows else {NAME_OURS: "", NAME_UCPH: ""}) if not c.startswith("_")]
        write_rows(ws, rows, columns)
        if rows:
            style_matches_sheet(ws, rows, columns)
        else:
            style_header(ws)
        autosize(ws, 45)
    for name, rows, width in (("summary", summary, 70),
                              ("method", [["how matches are made"]] + [[l] for l in __doc__.strip().splitlines()], 95)):
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(list(r))
        style_header(ws)
        ws.freeze_panes = "A2"
        autosize(ws, width)
    return wb


def summary_rows(contacts_name, ucph_name, s):
    return [["metric", "value"], ["Contacts file", contacts_name], ["UCPH list", ucph_name],
            ["Contacts", s["contacts"]], ["UCPH people", s["ucph_people"]],
            ["Possible new hires (pairs)", s["pairs"]], ["  High", s["High"]],
            ["  Medium", s["Medium"]], ["  Low", s["Low"]],
            ["Possible new hires (distinct contacts)", s["contacts_matched"]],
            ["Already at UCPH at signup", s["already_at_ucph"]],
            ["Created", dt.datetime.now().strftime("%Y-%m-%d %H:%M")]]


def compare(data, filename, ucph_records, ucph_name, sheet=None, columns=None, log=print):
    """Core used by both the website and the command line.
    -> (openpyxl Workbook, stats, candidates, already)."""
    _, contacts = load_contacts(data, filename, sheet, columns, log)
    people = people_from_records(ucph_records)
    candidates, already = match(people, contacts)
    stats = stats_of(contacts, people, candidates, already)
    wb = build_workbook(candidates, already, summary_rows(filename, ucph_name, stats))
    log(f"{stats['contacts']:,} contacts vs {stats['ucph_people']:,} UCPH people ({ucph_name})")
    log(f"Possible new hires: {stats['pairs']} pairs (High {stats['High']}, "
        f"Medium {stats['Medium']}, Low {stats['Low']})")
    log(f"Already at UCPH at signup (same email or ku.dk address): {stats['already_at_ucph']}")
    return wb, stats, candidates, already


def run(database, ku_path, output, sheet=None, columns=None, log=print):
    """Command line: match the contacts file `database` against the UCPH CSV `ku_path`,
    write `output` (never overwriting). -> (output path, stats)."""
    database, ku_path = Path(database), Path(ku_path)
    try:
        data = database.read_bytes()
    except OSError as e:
        raise UserError(f"Could not read {database}: {e}") from e
    wb, stats, _, _ = compare(data, database.name, read_ucph_csv(ku_path), ku_path.name,
                              sheet, columns, log)
    out = unique_path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    log(f"Wrote {out}")
    return out, stats
