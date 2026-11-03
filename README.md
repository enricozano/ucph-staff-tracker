# UCPH Staff Tracker

**Website: <https://enricozano.github.io/ucph-staff-tracker/>**

A website that answers two questions:

1. **Who are the new academic staff at the University of Copenhagen (UCPH)?**
2. **Have any of the people in our contact database been hired there?**

Nothing to install. Open the link in any browser (Chrome, Edge, Safari, Firefox).

> Technical details for developers are in [DEVELOPERS.md](DEVELOPERS.md).

---

## What's on the website

| Tab | What you can do |
|---|---|
| **New staff** | See who joined UCPH between two downloads of the staff list, and download it as Excel. |
| **All academic staff** | Search and filter everyone with an academic position at UCPH right now, and download the list as Excel. |
| **Compare with our contacts** | Choose an export of our contacts (e.g. from HubSpot) and see who now works at UCPH. |
| **Help** | A short version of this guide. |

"Academic staff" means professors (incl. associate, assistant and clinical professors),
postdocs, PhD students, researchers, research assistants and lecturers. Tick **"Include
guests, emeriti and teaching assistants"** to add those too.

---

## Comparing our contacts with UCPH

1. Export the contacts from HubSpot (or another system) as an **Excel (.xlsx) or CSV
   file**. It needs first and last names (one "Name" column also works). Emails, company
   and country make the results better.
2. Open the **Compare with our contacts** tab and drop the file onto the box, or click
   the box to choose it.
3. Choose what to compare with:
   - **All current academic staff**: everyone at UCPH now, or
   - **Only new academic staff since …**: only people who joined since the chosen
     download (the usual monthly check).
4. Click **Find matches**. The first time it takes 10–20 seconds to get ready, then the
   results appear.
5. Click **Download Excel** for the full results.

**Your file stays on your computer.** The comparison runs inside your browser window, and
the file is never uploaded anywhere.

### Reading the results

| | Meaning |
|---|---|
| **High** | The same full name, and an unusual one. Very likely the same person. |
| **Medium** | The same name, but a fairly common one. Probably right; check it. |
| **Low** | A common name shared by several people, or a slightly different name (e.g. a missing middle name). Check carefully. |
| **Already at UCPH when they signed up** | The contact gave a UCPH email address (staff or student) when joining our database, so they were already at UCPH. |

The Excel file shows **UCPH staff (blue) and our CRM (green) side by side**. Click a name
to search for that person on the web, or *Open profile* for their UCPH page. When the CRM
says the contact's university is UCPH or another Danish university, *Match type* says so:
those people were probably already in Denmark.

### Letting Claude assess the matches

Whether two records really are the same person, and whether they **moved to Denmark from
abroad**, takes some research. Claude can do it for you:

1. After the comparison, the website shows a **Next: let Claude assess the matches** box.
2. Click **Open Claude**, start a new chat and **attach the Excel file** you downloaded.
3. Click **Copy prompt**, paste it into the chat and send it.
4. Claude fills in the **Claude flag** and **Claude note** columns and gives you the file
   back. Download it.
5. Fill in **Your verdict**, which has a dropdown: **YES** = same person, moved to Denmark
   from abroad; **MID** = same person, already in Denmark; **NO** = different person. Add
   any notes under **Your comments**.

The same instructions are inside the Excel file, on the *Instructions for Claude* sheet.
Use a Claude plan that allows file uploads and web search.

**The same name doesn't always mean the same person, so check before acting on a match.**

---

## Updating the data (about once a month)

The UCPH phone book doesn't say *when* someone was hired. The website therefore keeps
copies of the phone book taken on different days ("downloads") and compares them. **Do a
new download about once a month.** It needs a GitHub account with access to the project.

1. Open the [**Update site**](https://github.com/enricozano/ucph-staff-tracker/actions/workflows/update-site.yml) page.
2. Click **Run workflow** (on the right), then the green **Run workflow** button.
3. Wait about 15 minutes. A green tick means it worked, and the website shows the new
   date in the top right ("Last download: …"). A red cross means the download failed,
   usually because the UCPH website is down. Try again later. Nothing changes on the
   website until a download succeeds.

After the update, the **New staff** tab shows who joined since the previous download.

---

## Questions

**Who can see the website?** Anyone who has the link. It's hidden from search engines,
and it only shows what the public UCPH phone book shows: names, job titles, departments
and links to UCPH profile pages. No phone numbers, addresses or email addresses.

**Is our contact data safe?** Yes. Contact files are only read inside your own browser.
They're never uploaded or stored.

**Privacy.** The downloaded Excel files contain personal data (names, job titles, and our
contacts' details). Treat them under the GDPR: keep them on work computers, and only
share them with people who need them.

**"Could not find the name columns".** The contacts file needs columns called
*First Name* and *Last Name* (or a single *Name* column). Rename the columns in Excel,
save, and try again.

**"This is an old-style .xls file".** Open it in Excel, choose *Save as → Excel Workbook
(.xlsx)*, and use the new file.

**Something else?** Contact the person who maintains the website.
