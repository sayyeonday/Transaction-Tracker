"use strict";

const $ = (id) => document.getElementById(id);
const buildBtn = $("build");
const statusEl = $("status");

let pyodide = null;

// Trusted markup only (spinners, the download link with a blob: URL).
function setStatus(msg, kind = "") {
  statusEl.className = kind;
  statusEl.innerHTML = msg;
}

// Plain text — use for anything that includes an exception message or other
// untrusted content, so it can never be interpreted as HTML.
function setStatusText(text, kind = "") {
  statusEl.className = kind;
  statusEl.textContent = text;
}

function readText(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsText(file);
  });
}

function readBytes(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(new Uint8Array(r.result));
    r.onerror = () => reject(r.error);
    r.readAsArrayBuffer(file);
  });
}

async function initPyodide() {
  try {
    setStatus('<span class="spinner"></span>Downloading Python engine (one time, ~10&nbsp;MB)…');
    pyodide = await loadPyodide();

    setStatus('<span class="spinner"></span>Loading data libraries…');
    await pyodide.loadPackage(["pandas", "micropip"]);
    await pyodide.runPythonAsync(
      'import micropip; await micropip.install("openpyxl")'
    );

    setStatus('<span class="spinner"></span>Loading workbook builder…');
    for (const f of ["clean_data.py", "build_excel.py"]) {
      const txt = await (await fetch(f)).text();
      pyodide.FS.writeFile("/home/pyodide/" + f, txt);
    }
    await pyodide.runPythonAsync(`
import sys
if "/home/pyodide" not in sys.path:
    sys.path.insert(0, "/home/pyodide")
import io, build_excel
`);

    buildBtn.disabled = false;
    buildBtn.innerHTML = "Build my workbook";
    setStatusText("Ready — your files stay on this device.", "ok");
  } catch (e) {
    console.error(e);
    setStatusText("Couldn't load the engine: " + e.message, "err");
  }
}

async function buildWorkbook() {
  const creditFiles = Array.from($("credit").files);
  const debitFiles = Array.from($("debit").files);
  const priorFile = $("prior").files[0] || null;

  if (creditFiles.length === 0 && debitFiles.length === 0) {
    setStatusText("Please choose at least one CSV file first.", "err");
    return;
  }

  buildBtn.disabled = true;
  setStatus('<span class="spinner"></span>Reading your files…');

  try {
    const specs = [];
    for (const f of creditFiles) specs.push([await readText(f), "credit"]);
    for (const f of debitFiles) specs.push([await readText(f), "debit"]);
    const priorBytes = priorFile ? await readBytes(priorFile) : null;

    setStatus('<span class="spinner"></span>Building your workbook…');
    pyodide.globals.set("js_specs", pyodide.toPy(specs));
    pyodide.globals.set("js_prior", priorBytes);

    const resultProxy = await pyodide.runPythonAsync(`
import io, build_excel
_specs = [(io.StringIO(t), s) for (t, s) in js_specs]
_pb = bytes(js_prior.to_py()) if js_prior is not None else None
build_excel.build_bytes(_specs, _pb)
`);
    const bytes = resultProxy.toJs();
    resultProxy.destroy();

    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    setStatus(
      'Done! <a class="download" href="' + url + '" download="finance.xlsx">Download finance.xlsx</a>',
      "ok"
    );
  } catch (e) {
    console.error(e);
    const msg = String(e.message || e);
    const clean = msg.includes("No transactions")
      ? "No transactions found — are these CIBC CSV exports?"
      : "Something went wrong: " + msg;
    setStatusText(clean, "err");
  } finally {
    buildBtn.disabled = false;
  }
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch((e) => console.warn("SW:", e));
  });
}

buildBtn.addEventListener("click", buildWorkbook);
initPyodide();
