"""Download the public UCPH phone book into a dated snapshot CSV.

The phone book page (https://about.ku.dk/contact/phone-book/) loads its results from a
JSON endpoint. That endpoint has a few quirks handled here:
  - at most 99 records per page;
  - a search term is required, so several terms are run and their results unioned;
  - raw control characters and unescaped backslashes inside JSON strings;
  - now and then an empty page in the middle of a result list.
"""
import datetime as dt
import json
import re
import time
from pathlib import Path

import pandas as pd
import requests

from ku_staff import __version__
from ku_staff.matching import email_hash

API = "https://www2.adm.ku.dk/selv/pls/!app_tlfbog_v2.soeg"
PAGE_SIZE = 99  # the API caps a page at 99 records
# Search terms whose result sets are unioned; "ku.dk" matches nearly everyone,
# the vowels catch entries without a ku.dk email.
QUERIES = ["ku.dk", "a", "e", "i", "o", "u"]
DELAY = 0.4  # seconds between requests, to keep the load on the UCPH server low
USER_AGENT = f"Mozilla/5.0 (ucph-staff-tracker/{__version__}; research use)"

# Snapshot column -> API field. Snapshots are published in a public repository, so they
# keep only what is needed: no phone numbers, addresses or plain email addresses.
FIELDS = {
    "first_name": "PERSON_FORNAVN",
    "last_name": "PERSON_EFTERNAVN",
    "title_en": "ANSAT_UOFF_STIL_TEKST_ENGELSK",
    "title_da": "ANSAT_UOFF_STIL_TEKST",
    "unit_en": "STED_NAVN_SAMLET_ENG",
    "research_profile": "ANSAT_PURE_UK",
}
SNAPSHOT_COLUMNS = ["person_id", "first_name", "last_name", "title_en", "title_da", "unit_en",
                    "department", "subunit", "faculty", "research_profile", "email_sha256"]



class CrawlError(Exception):
    pass


def parse(text):
    text = re.sub(r'\\(?!["\\/bfnrtu])', r"\\\\", text)
    return json.loads(text, strict=False)["root"]["employees"]


def normalize(raw):
    """One API record -> one snapshot row, with department/sub-unit/faculty split out."""
    row = {"person_id": ""}
    row.update({k: (raw.get(v) or "").strip() for k, v in FIELDS.items()})
    m = re.search(r"p_personid=(\d+)", raw.get("FOTOURL") or "")
    row["person_id"] = m.group(1) if m else ""
    row["email_sha256"] = email_hash(raw.get("ANSAT_ARB_EMAIL"))
    unit = row["unit_en"] or (raw.get("STED_NAVN_SAMLET") or "").strip()
    parts = [p.strip() for p in unit.split(",")]
    row["faculty"] = next((p for p in reversed(parts) if re.search(r"faculty|fakultet", p, re.I)), "")
    row["department"] = parts[0] if parts else ""
    row["subunit"] = ", ".join(p for p in parts[1:] if p != row["faculty"])
    return row


class Crawler:
    def __init__(self, log=print, progress=None, delay=DELAY):
        self.log = log
        self.progress = progress or (lambda fraction, text: None)
        self.delay = delay
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT

    def fetch_page(self, query, start, size=PAGE_SIZE):
        for attempt in range(5):
            try:
                r = self.session.post(
                    API,
                    data={"format": "json", "startrecord": start, "recordsperpage": size,
                          "searchstring": query},
                    timeout=30,
                )
                r.raise_for_status()
                break
            except requests.RequestException as e:
                wait = 2 ** attempt
                self.log(f"Connection problem ({e}); retrying in {wait}s")
                time.sleep(wait)
        else:
            raise CrawlError("Could not reach the UCPH phone book. Check the internet connection "
                             "and try again later.")
        try:
            return parse(r.text)
        except (ValueError, KeyError) as e:
            if size == 1:
                self.log(f"Skipping one unreadable record ({query!r} #{start}: {e})")
                return []
            # Isolate the malformed record by refetching this page one record at a time.
            rows = []
            for i in range(start, start + size):
                rows += self.fetch_page(query, i, size=1)
                time.sleep(self.delay / 2)
            return rows

    def fetch_query(self, query, rows, gaps, qi):
        start, total = 0, None
        while total is None or start < total:
            # The server occasionally returns an empty page mid-listing; retry before giving up.
            for attempt in range(4):
                page = self.fetch_page(query, start)
                if page:
                    break
                time.sleep(2 ** (attempt + 1))
            if not page:
                gaps.append((query, start))
                if total is None:
                    self.log(f"WARNING: no results at all for search {query!r}")
                    return
                self.log(f"WARNING: empty page for search {query!r} at record {start}, skipped")
                start += PAGE_SIZE
                continue
            total = int(page[0].get("TOTAL_RECORDS") or 0)
            for raw in page:
                row = normalize(raw)
                key = (row["person_id"] or row["email_sha256"] or f'{row["first_name"]} {row["last_name"]}',
                       row["unit_en"], row["title_en"])
                rows[key] = row
            start += PAGE_SIZE
            done = min(start, total)
            self.progress((qi + done / max(total, 1)) / len(QUERIES),
                          f"Search {qi + 1} of {len(QUERIES)}: {done:,}/{total:,} records, "
                          f"{len(rows):,} unique entries so far")
            time.sleep(self.delay)

    def crawl(self, snapshot_dir):
        """Download everything and save snapshot_dir/YYYY-MM-DD.csv. Returns (path, gaps)."""
        rows, gaps = {}, []
        for qi, query in enumerate(QUERIES):
            self.fetch_query(query, rows, gaps, qi)
        if not rows:
            raise CrawlError("The UCPH phone book returned no data. It may be down or may have "
                             "changed; try again later.")
        snapshot_dir = Path(snapshot_dir)
        snapshot_dir.mkdir(parents=True, exist_ok=True)
        df = pd.DataFrame(rows.values())[SNAPSHOT_COLUMNS].sort_values(["last_name", "first_name"])
        path = snapshot_dir / f"{dt.date.today().isoformat()}.csv"
        df.to_csv(path, index=False)
        if gaps:
            self.log(f"WARNING: {len(gaps)} page(s) could not be fetched; the snapshot may be "
                     "slightly incomplete. Consider downloading again later.")
        self.log(f"Saved {len(df):,} entries ({df['person_id'].nunique():,} people) to {path}")
        return path, gaps
