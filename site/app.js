"use strict";
// UCPH Staff Tracker website. Data comes from data/*.json, built by `ku-staff publish`.
// Everything shown is inserted as text (never as HTML), so names can't inject markup.

const REPO = "enricozano/ucph-staff-tracker";
const UPDATE_WORKFLOW = `https://github.com/${REPO}/actions/workflows/update-site.yml`;
const PAGE = 100; // table rows shown before "Show more"

const $ = (id) => document.getElementById(id);
const state = { index: null, cache: {}, worker: null, file: null, results: null, view: "candidates" };

// ---- helpers --------------------------------------------------------------------
function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "class") node.className = v;
    else if (k === "text") node.textContent = v;
    else node.setAttribute(k, v);
  }
  for (const c of children) if (c != null) node.append(c instanceof Node ? c : String(c));
  return node;
}

function link(text, href) {
  return href ? el("a", { href, target: "_blank", rel: "noopener", text }) : document.createTextNode(text);
}

function niceDate(iso) {
  return new Date(`${iso}T00:00:00`).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
}

async function getJSON(url) {
  if (!state.cache[url]) {
    state.cache[url] = fetch(url, { cache: "no-cache" }).then((r) => {
      if (!r.ok) throw new Error(`Could not load ${url} (${r.status})`);
      return r.json();
    });
  }
  return state.cache[url];
}

function records(table) {
  const cols = table.columns;
  return table.rows.map((r) => Object.fromEntries(cols.map((c, i) => [c, r[i]])));
}

function matchesSearch(row, query, fields) {
  if (!query) return true;
  const hay = fields.map((f) => row[f] || "").join(" ").toLowerCase();
  return query.toLowerCase().split(/\s+/).every((w) => hay.includes(w));
}

/** Render rows into a <table>, PAGE at a time, with a "Show more" button. */
function renderTable(table, moreBtn, columns, rows, emptyText) {
  let shown = 0;
  table.replaceChildren(el("thead", {}, el("tr", {}, ...columns.map((c) => el("th", { text: c.label })))));
  const body = el("tbody");
  table.append(body);
  const page = () => {
    for (const row of rows.slice(shown, shown + PAGE)) {
      body.append(el("tr", {}, ...columns.map((c) => el("td", { class: c.cls || "" }, c.cell(row)))));
    }
    shown = Math.min(rows.length, shown + PAGE);
    if (moreBtn) {
      moreBtn.hidden = shown >= rows.length;
      moreBtn.textContent = `Show more (${(rows.length - shown).toLocaleString("en")} left)`;
    }
  };
  if (!rows.length) body.append(el("tr", { class: "empty-row" }, el("td", { colspan: columns.length, text: emptyText })));
  else page();
  if (moreBtn) moreBtn.onclick = page;
}

const nameCell = (r) => link(`${r.first_name} ${r.last_name}`, r.research_profile);
const STAFF_COLUMNS = [
  { label: "Name", cls: "name", cell: nameCell },
  { label: "Position", cell: (r) => r.title_en },
  { label: "Department", cell: (r) => [r.department, r.subunit].filter(Boolean).join(" · ") },
  { label: "Faculty", cell: (r) => r.faculty },
];

function setDownload(a, href) {
  a.href = href || "#";
  a.classList.toggle("disabled", !href);
}

// ---- tabs -----------------------------------------------------------------------
function showTab(name) {
  if (!document.querySelector(`.tab[data-tab="${name}"]`)) name = "new";
  for (const t of document.querySelectorAll(".tab")) t.setAttribute("aria-selected", t.dataset.tab === name);
  for (const p of document.querySelectorAll(".panel")) p.hidden = p.id !== `tab-${name}`;
  if (location.hash !== `#${name}`) history.replaceState(null, "", `#${name}`);
  if (name === "all") showAll();
  if (name === "compare") startWorker(); // load the matching engine while the user picks a file
}

// ---- New staff ------------------------------------------------------------------
function setupNew() {
  const dates = Object.keys(state.index.new_since).sort().reverse();
  const latest = state.index.latest;
  if (!dates.length) {
    const box = $("new-empty");
    box.hidden = false;
    box.replaceChildren(
      el("strong", { text: latest ? "Only one download so far" : "No downloads yet" }),
      latest
        ? `The phone book was downloaded on ${niceDate(latest)}. New staff appear here after the next download, by comparing the two. See the Help tab for how to run it.`
        : "Run the first download from GitHub (see the Help tab).",
    );
    return;
  }
  $("new-body").hidden = false;
  $("new-since").replaceChildren(...dates.map((d) => el("option", { value: d, text: niceDate(d) })));
  for (const id of ["new-since", "new-affiliated"]) $(id).addEventListener("change", renderNew);
  $("new-search").addEventListener("input", renderNew);
  renderNew();
}

async function renderNew() {
  const since = $("new-since").value;
  const incl = $("new-affiliated").checked;
  const entry = state.index.new_since[since];
  const all = records(await getJSON(entry.json));
  const rows = all.filter((r) => incl || r.core);
  $("new-count").textContent = rows.length.toLocaleString("en");
  $("new-count-label").textContent =
    `new ${incl ? "academic staff, guests, emeriti and teaching assistants" : "academic staff"} between ${niceDate(since)} and ${niceDate(state.index.latest)}`;
  setDownload($("new-download"), entry[`download_${incl ? "academic_incl_affiliated" : "academic"}`]);
  const q = $("new-search").value.trim();
  const shown = rows.filter((r) => matchesSearch(r, q, ["first_name", "last_name", "title_en", "department", "faculty", "previous_title"]));
  renderTable($("new-table"), $("new-more"), [
    ...STAFF_COLUMNS,
    { label: "Before", cell: (r) => r.previous_title || "not at UCPH" },
  ], shown, q ? "Nobody matches the search." : "No new staff in this period.");
}

// ---- All academic staff ---------------------------------------------------------
let allReady = null;
function showAll() {
  if (!state.index.latest) {
    $("all-count").replaceChildren(noDataYet("The list appears after the first download."));
    return;
  }
  allReady ||= (async () => {
    const rows = records(await getJSON("data/academic-latest.json"));
    const uniq = (f) => [...new Set(rows.map((r) => r[f]).filter(Boolean))].sort();
    $("all-faculty").append(...uniq("faculty").map((f) => el("option", { value: f, text: f })));
    $("all-category").append(...uniq("category").map((c) => el("option", { value: c, text: c })));
    for (const id of ["all-faculty", "all-category", "all-affiliated"]) $(id).addEventListener("change", () => renderAll(rows));
    $("all-search").addEventListener("input", () => renderAll(rows));
    renderAll(rows);
  })().catch(showLoadError);
}

function renderAll(rows) {
  const q = $("all-search").value.trim();
  const fac = $("all-faculty").value, cat = $("all-category").value;
  const incl = $("all-affiliated").checked || (cat && !rows.find((r) => r.category === cat)?.core);
  const base = rows.filter((r) => incl || r.core);
  const shown = base.filter((r) => (!fac || r.faculty === fac) && (!cat || r.category === cat) &&
    matchesSearch(r, q, ["first_name", "last_name", "title_en", "department", "subunit", "faculty"]));
  $("all-count").textContent = `Showing ${shown.length.toLocaleString("en")} of ${base.length.toLocaleString("en")} entries` +
    " (people with several positions appear once per position).";
  renderTable($("all-table"), $("all-more"), STAFF_COLUMNS, shown, "Nobody matches the search.");
}

// ---- Compare --------------------------------------------------------------------
function startWorker() {
  if (!state.worker) {
    state.worker = new Worker("worker.js");
    state.worker.onmessage = onWorkerMessage;
    state.worker.postMessage({ cmd: "init" });
  }
  return state.worker;
}

function setupCompare() {
  const dates = Object.keys(state.index.new_since).sort().reverse();
  $("compare-since").replaceChildren(...dates.map((d) => el("option", { value: d, text: niceDate(d) })));
  if (!dates.length) {
    $("against-new").disabled = true;
    $("compare-since").disabled = true;
    $("against-new-label").classList.add("disabled");
    $("against-new-label").title = "Available after the second download";
  }
  $("compare-since").addEventListener("change", () => { $("against-new").checked = true; });

  const drop = $("drop");
  $("file").addEventListener("change", (e) => chooseFile(e.target.files[0]));
  for (const ev of ["dragenter", "dragover"]) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); });
  for (const ev of ["dragleave", "drop"]) drop.addEventListener(ev, () => drop.classList.remove("over"));
  drop.addEventListener("drop", (e) => { e.preventDefault(); chooseFile(e.dataTransfer.files[0]); });
  $("run").addEventListener("click", runCompare);
  for (const b of document.querySelectorAll(".subtab")) b.addEventListener("click", () => {
    state.view = b.dataset.view;
    for (const o of document.querySelectorAll(".subtab")) o.setAttribute("aria-selected", o === b);
    renderResults();
  });
  if (!state.index.latest) {
    $("run").title = "No staff list has been downloaded yet";
    $("drop").before(noDataYet("Comparing becomes available after the first download."));
  }
}

function noDataYet(text) {
  return el("div", { class: "empty" }, el("strong", { text: "No staff list yet" }),
    `${text} See the Help tab for how to run it.`);
}

function chooseFile(file) {
  if (!file) return;
  state.file = file;
  const kb = file.size < 1e6 ? `${Math.round(file.size / 1e3)} KB` : `${(file.size / 1e6).toFixed(1)} MB`;
  $("drop-title").textContent = `${file.name} (${kb})`;
  $("drop").classList.add("has-file");
  $("run").disabled = !state.index.latest;
  $("error").hidden = true;
}

function runCompare() {
  const incl = $("compare-affiliated").checked;
  const useNew = $("against-new").checked && !$("against-new").disabled;
  const since = $("compare-since").value;
  const latest = state.index.latest;
  const dataUrl = useNew ? state.index.new_since[since].json : "data/academic-latest.json";
  const listName = useNew
    ? `new academic staff ${niceDate(since)}–${niceDate(latest)}`
    : `academic staff ${niceDate(latest)}`;
  const listSlug = useNew ? `new_academic_${since}_to_${latest}` : `academic_${latest}`;
  state.pending = { listName, listSlug, fileName: state.file.name };
  state.running = true;
  $("run").disabled = true;
  $("error").hidden = true;
  $("results").hidden = true;
  setProgress("Reading your file…");
  state.file.arrayBuffer().then((buffer) => {
    startWorker().postMessage({ cmd: "compare", buffer, filename: state.file.name, dataUrl,
      includeAffiliated: incl, listName }, [buffer]);
  });
}

function setProgress(text) {
  $("progress").hidden = !text;
  if (text) $("progress-text").textContent = text;
}

function onWorkerMessage(e) {
  const m = e.data;
  if (!state.running) return; // engine preloading in the background: stay quiet
  if (m.type === "progress") {
    setProgress(m.text);
  } else if (m.type === "result") {
    state.running = false;
    setProgress(null);
    $("run").disabled = false;
    showResults(m);
  } else if (m.type === "error") {
    state.running = false;
    setProgress(null);
    $("run").disabled = !state.file;
    const box = $("error");
    box.hidden = false;
    box.textContent = m.message + (m.details ? `\n\nTechnical details: ${m.details}` : "");
  }
}

function showResults(m) {
  const { listName, listSlug, fileName } = state.pending;
  const s = m.stats;
  state.results = m;
  $("r-contacts").textContent = s.contacts_matched;
  $("r-high").textContent = s.High;
  $("r-medium").textContent = s.Medium;
  $("r-low").textContent = s.Low;
  $("r-already").textContent = s.already_at_ucph;
  $("results-sub").textContent =
    `${s.contacts.toLocaleString("en")} contacts in ${fileName} compared with ${s.ucph_people.toLocaleString("en")} people (${listName}).`;
  const stem = fileName.replace(/\.[^.]+$/, "").replace(/[^\w-]+/g, "_");
  const a = $("results-download");
  if (a.href.startsWith("blob:")) URL.revokeObjectURL(a.href);
  a.href = URL.createObjectURL(new Blob([m.xlsx], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  a.download = `matches_${stem}_vs_${listSlug}.xlsx`;
  for (const b of document.querySelectorAll(".subtab")) {
    const n = b.dataset.view === "candidates" ? m.candidates.length : m.already.length;
    b.textContent = `${b.dataset.view === "candidates" ? "Possible new hires" : "Already at UCPH at signup"} (${n})`;
  }
  $("results").hidden = false;
  renderResults();
  $("results").scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderResults() {
  const m = state.results;
  if (!m) return;
  const rows = state.view === "candidates" ? m.candidates : m.already;
  const columns = [
    { label: "Confidence", cls: "nowrap", cell: (r) => el("span", { class: `pill ${r.confidence}`, text: r.confidence }) },
    { label: "Name in our database", cls: "name pair", cell: (r) => link(r.name_in_our_database, r._search_ours) },
    { label: "Name at UCPH", cls: "name pair", cell: (r) => link(r.name_at_ucph, r._search_ucph) },
    { label: "Position at UCPH", cell: (r) => r.ucph_title },
    { label: "Department", cell: (r) => r.ucph_department },
    { label: "Why", cls: "nowrap", cell: (r) => [r.match_type, r.evidence].filter(Boolean).join(" · ") },
    { label: "Profile", cls: "nowrap", cell: (r) => link(r.ucph_research_profile ? "UCPH page" : "", r.ucph_research_profile) },
    { label: "Row in file", cls: "nowrap", cell: (r) => r.db_row },
  ];
  if (state.view === "already") columns.splice(1, 0, { label: "Reason", cell: (r) => r.already_at_ucph_reason });
  renderTable($("results-table"), null, columns, rows,
    state.view === "candidates" ? "No possible new hires found." : "Nobody in this group.");
}

// ---- start --------------------------------------------------------------------------
function showLoadError(err) {
  $("freshness").textContent = "Could not load the data";
  console.error(err);
}

async function main() {
  $("repo-link").href = `https://github.com/${REPO}`;
  $("actions-link").href = UPDATE_WORKFLOW;
  for (const t of document.querySelectorAll(".tab")) t.addEventListener("click", () => showTab(t.dataset.tab));
  try {
    state.index = await getJSON("data/index.json");
  } catch (err) {
    showLoadError(err);
    return;
  }
  const latest = state.index.latest;
  $("freshness").textContent = latest ? `Last download: ${niceDate(latest)}` : "No download yet";
  $("version").textContent = `Version ${state.index.version} · data built ${new Date(state.index.generated).toLocaleString("en-GB")}`;
  setDownload($("all-download"), state.index.downloads.academic);
  setDownload($("all-download-incl"), state.index.downloads.academic_incl_affiliated);
  if (latest) $("all-sub").textContent = `Everyone with an academic position at UCPH in the download of ${niceDate(latest)}.`;
  setupNew();
  setupCompare();
  showTab(location.hash.slice(1) || "new");
}

main();
