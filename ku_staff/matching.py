"""Match a contact database (xlsx/csv) against UCPH staff to find contacts who now work at UCPH.

Matching, per contact:
  email         same work email as a UCPH staff entry (compared as SHA-256 hashes)
  full name     all names identical after normalization (case, accents, ø/oe, å/aa, æ/ae, hyphens)
  partial name  same first given name and last surname, and every name part of the shorter
                name appears in the longer one (a middle name or second surname left out).
                Names that contradict each other (different middle names) are not matched.

Confidence of name matches (sheet "possible_new_hires"):
  High    full name, unique on both sides, and the name is distinctive: a rare surname,
          3+ name parts, or Danish evidence (.dk email, country Denmark, a Danish
          university or company in the CRM record)
  Medium  full name, unique on both sides, common surname; or a partial name, unique on
          both sides, with a rare surname or Danish/UCPH evidence
  Low     everything else: the name is shared by several people on either side, or a
          partial match on a common surname

A surname is rare when at most RARE_SURNAME people carry it across both files.
Click a name (UCPH or CRM side) to search for that person on the web.

Whether two records are really the same person, and whether they moved to Denmark from
abroad, is left to Claude ("Claude flag" / "Claude note") and to you ("Your verdict").

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
# A Danish institution in the CRM record: the contact was probably already in Denmark.
DANISH_INSTITUTION = re.compile(
    r"copenhagen|københavn|kobenhavn|ucph|\bku\b|aarhus|\bau\b|\bdtu\b|technical university of "
    r"denmark|southern denmark|\bsdu\b|aalborg|\baau\b|roskilde|\bruc\b|\bcbs\b|\bitu\b|"
    r"denmark|danmark|danish|dansk", re.I)
UCPH_EMAIL = re.compile(r"[@.]ku\.dk$", re.I)
CONFIDENCE_ORDER = {"High": 0, "Medium": 1, "Low": 2}
MATCH_ORDER = {"email": 0, "full name": 1, "partial name": 2}

SEARCH_URL = "https://www.google.com/search?q="
HEADER_FILL = "901A1E"  # UCPH red

# Column names tried, case-insensitively, when no column is given.
COLUMN_GUESSES = {
    "first": ["first name", "firstname", "first_name", "given name", "fornavn", "nome"],
    "last": ["last name", "lastname", "last_name", "surname", "family name", "efternavn", "cognome"],
    "name": ["name", "full name", "fullname", "full_name", "navn", "nome completo"],
    "email": ["email", "e-mail", "email address", "e-mail address", "mail"],
    "company": ["company name", "company", "shd associated university", "associated university",
                "university", "associated company", "organization", "organisation", "institution",
                "affiliation"],
    "country": ["country/region", "country", "land"],
    "date": ["create date", "created", "created at", "signup date", "sign up date", "date"],
    "record_id": ["record id - contact", "record id", "contact id", "id"],
    "phone": ["phone number", "phone", "mobile phone number", "telephone", "telefon"],
    "city": ["city", "by", "town"],
    "position": ["shd academic position", "academic position", "position", "job title", "jobtitle"],
    "field": ["shd job category", "job category", "research field", "field", "focus area",
              "shd other focus area"],
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
    """Contacts file (bytes) -> (headers, rows, label). Rows are lists of strings; label names
    the source (the sheet name for Excel, e.g. "SOD_ScienceHub_All Leads")."""
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
        label = ws.title
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
        label = Path(filename).stem[:40]
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
    return headers, body, label


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
    headers, rows, label = read_table(data, filename, sheet)
    columns = columns or {}
    cols = {role: find_column(headers, role, columns.get(role)) for role in COLUMN_GUESSES}
    idx = {h: i for i, h in enumerate(headers)}
    get = lambda row, role: row[idx[cols[role]]] if cols[role] else ""
    if cols["first"] and cols["last"]:
        names = [(get(r, "first"), get(r, "last")) for r in rows]
    elif cols["name"]:
        # Single full-name column: first word is the given name, the rest the surname(s).
        names = [tuple((get(r, "name").split(maxsplit=1) + [""])[:2]) for r in rows]
    else:
        raise UserError("Could not find the name columns in the contacts file. It needs a "
                        "'First Name' and a 'Last Name' column, or a single 'Name' column. "
                        f"Its columns are: {', '.join(headers)}")
    log("Contacts file columns used: " + ", ".join(f"{r}={c!r}" for r, c in cols.items() if c))
    contacts = []
    for i, (r, (first_name, last_name)) in enumerate(zip(rows, names)):
        email = get(r, "email")
        contacts.append({
            "first": canon(first_name), "last": canon(last_name),
            "first_name": first_name, "last_name": last_name,
            "name": f"{first_name} {last_name}".strip(),
            "email": email.lower(), "email_original": email,
            "record_id": get(r, "record_id"), "phone": get(r, "phone"),
            "company": get(r, "company"), "city": get(r, "city"), "country": get(r, "country"),
            "date": get(r, "date")[:10],  # date only
            "position": get(r, "position"), "field": get(r, "field"),
            "row_number": i + 2,  # as shown in Excel (header is row 1)
        })
    return headers, contacts, cols, label


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
            evidence.append("CRM university = UCPH")
        elif DANISH_INSTITUTION.search(c["company"]):
            evidence.append("CRM university in Denmark")
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
        rows.append({
            "already_reason": "same email" if kind == "email"
            else "ku.dk email in our CRM" if UCPH_EMAIL.search(c["email"]) else "",
            "confidence": confidence, "match_type": kind, "evidence": ", ".join(evidence),
            "surname_frequency": surname_count.get(surname, 0),
            "ucph_candidates": n_ku.get(h, 1), "contacts_matching": n_db.get(i, 1),
            "ucph_person_id": k["person_id"],
            "ucph_first_name": k["first_name"], "ucph_last_name": k["last_name"],
            "ucph_category": k["category"], "ucph_title": k["title_en"],
            "ucph_department": k["department"], "ucph_faculty": k["faculty"],
            "ucph_profile": k["research_profile"],
            "ucph_first_seen": k["first_seen"], "ucph_previous_title": k["previous_title"],
            "crm_record_id": c["record_id"], "crm_row": c["row_number"],
            "crm_first_name": c["first_name"], "crm_last_name": c["last_name"],
            "crm_email": c["email_original"], "crm_phone": c["phone"], "crm_company": c["company"],
            "crm_city_country": ", ".join(v for v in (c["city"], c["country"]) if v),
            "crm_date": c["date"], "crm_position": c["position"], "crm_field": c["field"],
            "search_ucph": search_url(ucph_name, "University of Copenhagen"),
            "search_crm": search_url(c["name"], c["company"]),
        })
    rows.sort(key=lambda r: (CONFIDENCE_ORDER[r["confidence"]], MATCH_ORDER[r["match_type"]],
                             r["surname_frequency"], r["ucph_last_name"].lower()))
    already = [r for r in rows if r["already_reason"]]
    candidates = [r for r in rows if not r["already_reason"]]
    return candidates, already


def stats_of(contacts, people, candidates, already):
    count = lambda c: sum(r["confidence"] == c for r in candidates)
    return {
        "contacts": len(contacts), "ucph_people": len(people), "pairs": len(candidates),
        **{c: count(c) for c in CONFIDENCE_ORDER},
        "contacts_matched": len({r["crm_row"] for r in candidates}),
        "already_at_ucph": len(already),
    }


# ---- Claude assessment -----------------------------------------------------------
# Judging whether two records are the same person, and whether they moved to Denmark from
# abroad, needs web research and judgement. The site can't call Claude itself, so the user
# uploads the Excel to a Claude chat with this prompt (shown on the site and in the file).
CLAUDE_PROMPT = """\
I've attached an Excel file of possible matches between contacts in our CRM and staff at the University of Copenhagen (UCPH), made by the UCPH Staff Tracker.

Goal: find the CRM contacts who now work at UCPH, and above all those who moved to Denmark from abroad to take the job.

For every row of the sheet "UCPH vs CRM", fill in the two empty columns "Claude flag" and "Claude note". Don't change anything else: keep all other columns, rows, order and formatting, and leave "Your verdict" and "Your comments" empty.

Claude flag, exactly one of:
- Likely same, from abroad: the same person, and they moved to Denmark from abroad
- Likely same, in Denmark: the same person, but already in Denmark before (e.g. studied or worked at a Danish university)
- Likely same: the same person, origin unclear
- Check: plausible, but needs a human look
- Likely different: probably two different people
- Ambiguous: several UCPH people share the name, so you can't tell which one

Claude note: one or two sentences with the evidence, e.g. how rare the name is; the CRM email; "University / Company", which is the institution the contact gave when they signed up (a Danish one, such as the University of Copenhagen, means they were probably already in Denmark; a foreign one suggests they moved from abroad); the CRM "Position" and "Field" (if present) compared with the UCPH title and department; the date added to the CRM (if present) compared with the UCPH position; the UCPH profile (the "Open profile" link); and public pages such as LinkedIn or university websites if you can search the web. Say where the person came from when you can tell.

When you're done, give me the completed Excel file to download, plus a short summary listing the people flagged "Likely same, from abroad".
"""


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
    """Simple one-row header (summary/method sheets, staff exports)."""
    from openpyxl.styles import Font, PatternFill

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)


# Layout of the match sheets: (group, header, row key or None for "fill in later", width).
# Group bands: dark colour for row 1, light colour for row 2.
GROUPS = {
    "num": ("595959", "F2F2F2", ""),
    "ucph": ("1F4E79", "DDEBF7", "UNIVERSITY OF COPENHAGEN (UCPH) STAFF"),
    "crm": ("375623", "E2EFDA", "CRM  ({crm})"),
    "assess": ("595959", "F2F2F2", "MATCH ASSESSMENT"),
    "review": ("7F6000", "FFF2CC", "YOUR REVIEW"),
}
NONE = "(none)"
LINK_BLUE = "0563C1"
FILLS = {"green": "C6EFCE", "yellow": "FFEB9C", "red": "FFC7CE", "blue": "DDEBF7", "grey": "F2F2F2"}
CONFIDENCE_FILL = {"High": FILLS["green"], "Medium": FILLS["yellow"], "Low": FILLS["red"]}
REVIEW_FILL = "FFFBEA"
CENTRED = {"#", "confidence", "surname_frequency", "ucph_candidates"}
VERDICTS = ["YES", "MID", "NO", "Unsure"]
# Cell colours that follow what Claude / the reviewer types (conditional formatting).
FLAG_COLOURS = [("Likely same", "green"), ("Check", "yellow"), ("Likely different", "red"),
                ("Ambiguous", "red"), ("Origin hint", "blue")]
VERDICT_COLOURS = [("YES", "green"), ("MID", "yellow"), ("NO", "red"), ("UNSURE", "grey")]


def match_columns(rows, new_staff, found=None):
    """Columns of the match sheets. `found`: contact-file roles that exist (CRM columns the
    file doesn't have are left out instead of showing "(none)" everywhere)."""
    has = lambda role: found is None or bool(found.get(role))
    cols = [("num", "#", "#", 4),
            ("ucph", "First name", "ucph_first_name", 14), ("ucph", "Last name", "ucph_last_name", 17),
            ("ucph", "Category", "ucph_category", 16), ("ucph", "Title", "ucph_title", 16),
            ("ucph", "Department", "ucph_department", 28), ("ucph", "Faculty", "ucph_faculty", 20),
            ("ucph", "Research profile", "ucph_profile", 12)]
    if new_staff:
        cols += [("ucph", "Before at UCPH", "ucph_previous_title", 16),
                 ("ucph", "First seen", "ucph_first_seen", 11)]
    cols += [("crm", "Record ID", "crm_record_id", 13) if has("record_id") else ("crm", "Row in file", "crm_row", 8),
             ("crm", "First name", "crm_first_name", 12), ("crm", "Last name", "crm_last_name", 16),
             ("crm", "Email", "crm_email", 30)]
    optional = [("phone", "Phone", "crm_phone", 10), ("company", "University / Company", "crm_company", 22),
                ("position", "Position", "crm_position", 14), ("field", "Field", "crm_field", 14),
                ("city_country", "City, Country", "crm_city_country", 16),
                ("date", "Date added to CRM", "crm_date", 12)]
    for role, header, key, width in optional:
        present = has("city") or has("country") if role == "city_country" else has(role)
        if present:
            cols.append(("crm", header, key, width))
    cols += [("assess", "Original confidence", "confidence", 11),
             ("assess", "Match type", "match_text", 20),
             ("assess", "Surname frequency", "surname_frequency", 10),
             ("assess", "UCPH people with this name", "ucph_candidates", 11),
             ("assess", "Claude flag", None, 16), ("assess", "Claude note", None, 48),
             ("review", "Your verdict", None, 13), ("review", "Your comments", None, 30)]
    return cols


def write_match_sheet(ws, rows, crm_label, new_staff, legend, found=None):
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    cols = match_columns(rows, new_staff, found)
    fill = lambda colour: PatternFill("solid", fgColor=colour)
    thin = Side(style="thin", color="D9D9D9")
    body_font = Font(name="Arial", size=10)
    wrap_top = Alignment(wrap_text=True, vertical="top")
    last_col = len(cols)

    # Row 1: group bands; row 2: column headers.
    start = 1
    for idx in range(1, last_col + 2):
        group = cols[idx - 1][0] if idx <= last_col else None
        if idx > 1 and group != cols[start - 1][0]:
            dark, _, title = GROUPS[cols[start - 1][0]]
            cell = ws.cell(1, start, title.format(crm=crm_label) or None)
            for c in range(start, idx):
                ws.cell(1, c).fill = fill(dark)
            cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if idx - 1 > start:
                ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=idx - 1)
            start = idx
    for c, (group, header, _, width) in enumerate(cols, start=1):
        cell = ws.cell(2, c, header)
        cell.font = Font(name="Arial", bold=True, size=10)
        cell.fill = fill(GROUPS[group][1])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium"))
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 34

    col_of = {header if key is None else key: i for i, (_, header, key, _) in enumerate(cols, start=1)}
    for n, row in enumerate(rows, start=1):
        r = n + 2
        values = dict(row, **{"#": n})
        values["match_text"] = row["already_reason"] or " · ".join(
            v for v in (row["match_type"], row["evidence"]) if v)
        for c, (group, header, key, _) in enumerate(cols, start=1):
            v = values.get(key) if key else None
            if key and key.startswith("crm_") and v in ("", None):
                v = NONE
            cell = ws.cell(r, c, v)
            cell.font, cell.alignment = body_font, wrap_top
            cell.border = Border(bottom=thin)
            if group == "review":
                cell.fill = fill(REVIEW_FILL)
            if key in CENTRED:
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center")
        # Names link to a web search for that person; last names in bold, as in the template.
        for side in ("ucph", "crm"):
            for part in ("first_name", "last_name"):
                cell = ws.cell(r, col_of[f"{side}_{part}"])
                if cell.value and cell.value != NONE:
                    cell.hyperlink = row[f"search_{side}"]
                    cell.font = Font(name="Arial", size=10, color=LINK_BLUE, underline="single",
                                     bold=part == "last_name")
        profile = ws.cell(r, col_of["ucph_profile"])
        if row["ucph_profile"]:
            profile.value, profile.hyperlink = "Open profile", row["ucph_profile"]
            profile.font = Font(name="Arial", size=10, color=LINK_BLUE, underline="single")
        else:
            profile.value = NONE
        ws.cell(r, col_of["confidence"]).fill = fill(CONFIDENCE_FILL[row["confidence"]])

    last_row = max(len(rows) + 2, 3)
    letter = lambda name: get_column_letter(col_of[name])
    flag, verdict = letter("Claude flag"), letter("Your verdict")
    dv = DataValidation(type="list", formula1='"' + ",".join(VERDICTS) + '"', allow_blank=True,
                        showErrorMessage=False)  # free text still allowed
    ws.add_data_validation(dv)
    dv.add(f"{verdict}3:{verdict}{last_row}")
    for column, rules in ((flag, FLAG_COLOURS), (verdict, VERDICT_COLOURS)):
        for text, colour in rules:
            formula = (f'ISNUMBER(SEARCH("{text}",{column}3))' if column == flag
                       else f'LEFT(UPPER(TRIM({column}3)),{len(text)})="{text}"')
            ws.conditional_formatting.add(f"{column}3:{column}{last_row}", Rule(
                type="expression", formula=[formula], stopIfTrue=True,
                dxf=DifferentialStyle(fill=PatternFill(bgColor=FILLS[colour]))))

    ws.freeze_panes = "D3"
    ws.auto_filter.ref = f"A2:{get_column_letter(last_col)}{last_row}"
    ws.sheet_view.zoomScale = 85
    for i, text in enumerate(legend):
        cell = ws.cell(last_row + 2 + i, 1, text)
        cell.font = Font(italic=True, size=9, color="595959")


def build_workbook(candidates, already, summary, crm_label, ucph_label, new_staff=False, found=None,
                   crm_file=None):
    """-> openpyxl Workbook: the two match sheets, Claude instructions, summary, method."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    created = dt.date.today().isoformat()
    legend = [
        "Legend: blue = UCPH staff, green = our CRM, grey = the automatic match assessment "
        "(confidence, match type, surname frequency) plus Claude's flag and note, yellow = fill in "
        "yourself. Your verdict: YES = same person, moved to Denmark from abroad; MID = same person, "
        "already in Denmark; NO = different person. Click a name to search for that person on the web.",
        f"Source: {crm_file or crm_label} compared with {ucph_label} on {created} (UCPH Staff Tracker). "
        "Claude flag and Claude note stay empty until you ask Claude: see the 'Instructions for Claude' sheet.",
    ]
    wb = Workbook()
    wb.remove(wb.active)
    write_match_sheet(wb.create_sheet("UCPH vs CRM"), candidates, crm_label, new_staff, legend, found)
    write_match_sheet(wb.create_sheet("Already at UCPH at signup"), already, crm_label, new_staff, [
        "These contacts used a UCPH email address in our CRM (the same as their UCPH work email, or "
        "another ku.dk address such as a student one), so they were already at UCPH when they signed up."],
        found)

    ws = wb.create_sheet("Instructions for Claude")
    steps = ["How to get Claude's assessment of the matches",
             "1. Open https://claude.ai and start a new chat.",
             "2. Attach this Excel file.",
             "3. Copy the prompt below (all the lines in grey), paste it into the chat and send it.",
             "4. Download the completed file Claude gives you, then fill in 'Your verdict' and 'Your comments'.",
             "", "Prompt:"]
    for line in steps:
        ws.append([line])
    for line in CLAUDE_PROMPT.splitlines():
        ws.append([line])
        ws.cell(ws.max_row, 1).font = Font(color="595959")
    ws["A1"].font = Font(bold=True, size=13)
    ws["A7"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 130
    for row in ws.iter_rows():
        row[0].alignment = Alignment(wrap_text=True, vertical="top")

    for name, rows, width in (("Summary", summary, 70),
                              ("Method", [["how matches are made"]] + [[l] for l in __doc__.strip().splitlines()], 95)):
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
    _, contacts, found, crm_label = load_contacts(data, filename, sheet, columns, log)
    people = people_from_records(ucph_records)
    candidates, already = match(people, contacts)
    stats = stats_of(contacts, people, candidates, already)
    new_staff = any(p["first_seen"] for p in people)
    wb = build_workbook(candidates, already, summary_rows(filename, ucph_name, stats),
                        crm_label, ucph_name, new_staff, found, crm_file=filename)
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
    wb, stats, _, _ = compare(data, database.name, read_ucph_csv(ku_path), ku_path.stem,
                              sheet, columns, log)
    out = unique_path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    log(f"Wrote {out}")
    log("Next: ask Claude to assess the matches (see the 'Instructions for Claude' sheet).")
    return out, stats
