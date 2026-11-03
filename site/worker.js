"use strict";
// Background thread that compares a contacts file with UCPH staff, using the project's
// own Python matching module (py/matching.py) running in Pyodide. The file never leaves
// the browser.
//
// Messages in:  {cmd: "init"} | {cmd: "compare", buffer, filename, dataUrl, includeAffiliated, listName}
// Messages out: {type: "progress", text} | {type: "ready"}
//               {type: "result", stats, candidates, already, prompt, xlsx} | {type: "error", message, details}

const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/";
// Pure-Python wheels shipped with the site (Pyodide doesn't include openpyxl).
const WHEELS = ["vendor/et_xmlfile-2.0.0-py3-none-any.whl", "vendor/openpyxl-3.1.5-py2.py3-none-any.whl"];

const GLUE = `
import io, json, sys
sys.path.insert(0, "/home/pyodide")
import matching

def run(data, filename, staff_json, include_affiliated, list_name, log):
    staff = json.loads(staff_json)
    cols = staff["columns"]
    recs = [dict(zip(cols, r)) for r in staff["rows"]]
    if not include_affiliated:
        recs = [r for r in recs if r["core"]]
    log("Comparing names…")
    try:
        wb, stats, cand, already = matching.compare(data.to_bytes(), filename, recs, list_name,
                                                    log=lambda m: None)
    except matching.UserError as e:
        return json.dumps({"error": str(e)}), b""
    log("Creating the Excel file…")
    buf = io.BytesIO()
    wb.save(buf)
    out = {"stats": stats, "candidates": cand, "already": already, "prompt": matching.CLAUDE_PROMPT}
    return json.dumps(out), buf.getvalue()
`;

const post = (msg, transfer = []) => self.postMessage(msg, transfer);
const progress = (text) => post({ type: "progress", text });
let engine = null;

async function init() {
  progress("Loading the matching engine (only the first time, about 10 seconds)…");
  try {
    importScripts(`${PYODIDE}pyodide.js`);
  } catch (err) {
    throw new UserFacingError("Could not load the matching engine. Check the internet connection and try again.", err);
  }
  const py = await loadPyodide({ indexURL: PYODIDE });
  const base = new URL(".", self.location).href;
  await py.loadPackage(WHEELS.map((w) => base + w));
  const code = await (await fetch(`${base}py/matching.py`, { cache: "no-cache" })).text();
  py.FS.writeFile("/home/pyodide/matching.py", code);
  py.runPython(GLUE);
  return py;
}

class UserFacingError extends Error {
  constructor(message, cause) {
    super(message);
    this.cause = cause;
  }
}

function toBytes(value) {
  if (value instanceof Uint8Array) return value;
  const buf = value.getBuffer("u8");
  try {
    return buf.data.slice();
  } finally {
    buf.release();
    value.destroy();
  }
}

self.onmessage = async (e) => {
  const m = e.data;
  try {
    engine ||= init();
    const py = await engine;
    if (m.cmd === "init") {
      post({ type: "ready" });
      return;
    }
    progress("Loading the UCPH staff list…");
    const res = await fetch(m.dataUrl, { cache: "no-cache" });
    if (!res.ok) throw new UserFacingError(`Could not load the UCPH staff list (${res.status}).`);
    const staffJson = await res.text();
    progress("Reading your file (large files take up to a minute)…");
    // Let the progress message reach the page before the (blocking) Python work starts.
    await new Promise((r) => setTimeout(r, 30));
    const result = py.globals.get("run")(new Uint8Array(m.buffer), m.filename, staffJson,
      m.includeAffiliated, m.listName, progress);
    const [json, xlsxProxy] = [result.get(0), result.get(1)];
    result.destroy();
    const out = JSON.parse(json);
    if (out.error) {
      post({ type: "error", message: out.error });
      return;
    }
    const xlsx = toBytes(xlsxProxy);
    post({ type: "result", ...out, xlsx }, [xlsx.buffer]);
  } catch (err) {
    if (err instanceof UserFacingError) {
      engine = null; // allow a retry
      post({ type: "error", message: err.message });
    } else {
      const lines = String(err && err.message || err).trim().split("\n");
      post({ type: "error", message: "Something went wrong while comparing the files.",
             details: lines[lines.length - 1] });
    }
  }
};
