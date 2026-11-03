"""Browser test of the website: builds site data from small fake snapshots, serves the site
locally and drives it with Playwright (Chrome). Needs `pip install playwright` and either
Google Chrome or `playwright install chromium`, plus internet access (Pyodide loads from
its CDN). Skipped when those are missing. Run with: python -m pytest -m browser
"""
import functools
import http.server
import shutil
import threading
from pathlib import Path

import pandas as pd
import pytest

from ku_staff.publish import publish
from tests.test_core import snapshot_row, write_snapshot

SITE = Path(__file__).resolve().parent.parent / "site"
pytestmark = pytest.mark.browser


@pytest.fixture
def site_url(tmp_path):
    write_snapshot(tmp_path, "2026-01-01", [
        snapshot_row("1", "Rune Munck", "Aabenhus", "Associate Professor", "runa@sund.ku.dk"),
    ])
    write_snapshot(tmp_path, "2026-02-01", [
        snapshot_row("1", "Rune Munck", "Aabenhus", "Associate Professor", "runa@sund.ku.dk"),
        snapshot_row("2", "Giulia", "Zampirolo", "Postdoc", "gz@sund.ku.dk"),
        snapshot_row("3", "Katarina", "Dragicevic", "Postdoc", "kd@bric.ku.dk"),
    ])
    site = tmp_path / "site"
    shutil.copytree(SITE, site, ignore=shutil.ignore_patterns("data", "py"))
    publish(tmp_path, site, log=lambda m: None)
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(site))
    handler.log_message = lambda *a: None
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{server.server_address[1]}/"
    server.shutdown()


@pytest.fixture
def page():
    sync_api = pytest.importorskip("playwright.sync_api")
    with sync_api.sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome")
        except Exception:
            try:
                browser = p.chromium.launch()
            except Exception as e:
                pytest.skip(f"no browser available: {e}")
        page = browser.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        yield page
        browser.close()
        assert not errors, errors


def test_new_staff_tab(site_url, page):
    page.goto(site_url)
    page.wait_for_selector("#new-table tbody tr")
    assert page.text_content("#new-count") == "2"
    names = page.locator("#new-table tbody td.name").all_text_contents()
    assert sorted(names) == ["Giulia Zampirolo", "Katarina Dragicevic"]


def test_compare_in_browser(site_url, page, tmp_path):
    contacts = tmp_path / "contacts.csv"
    pd.DataFrame({"First Name": ["Giulia", "Rune", "Katarina"],
                  "Last Name": ["Zampirolo", "Aabenhus", "Dragicevic"],
                  "Email": ["g@gmail.com", "r@example.org", "kd@bric.ku.dk"]}).to_csv(contacts, index=False)
    page.goto(site_url + "#compare")
    page.set_input_files("#file", str(contacts))
    page.click("#run")
    page.wait_for_selector("#results:not([hidden])", timeout=180_000)
    stats = {k: page.text_content(f"#r-{k}") for k in ("contacts", "high", "medium", "already")}
    # Zampirolo: full name -> High; Aabenhus: middle name missing -> Medium;
    # Dragicevic: same work email -> already at UCPH when they signed up.
    assert stats == {"contacts": "2", "high": "1", "medium": "1", "already": "1"}
    with page.expect_download() as d:
        page.click("#results-download")
    out = tmp_path / "matches.xlsx"
    d.value.save_as(out)
    df = pd.read_excel(out, sheet_name="possible_new_hires")
    assert set(df["name_at_ucph"]) == {"Giulia Zampirolo", "Rune Munck Aabenhus"}
