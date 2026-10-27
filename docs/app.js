"use strict";
const $ = (id) => document.getElementById(id);
const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const PDF_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

let py = null, api = null, analysis = null;
let pdfWords = null;                       // positioned words when the vendor file is a PDF

const setStatus = (el, msg, err) => { el.textContent = msg || ""; el.classList.toggle("err", !!err); };
const tick = () => new Promise((r) => setTimeout(r, 30));      // let the browser repaint before heavy work

async function initEngine() {
  try {
    py = await loadPyodide({ indexURL: PYODIDE_URL });
    await py.loadPackage("micropip");
    await py.pyimport("micropip").install("openpyxl");
    py.FS.mkdirTree("/work/fcu"); py.FS.mkdirTree("/work/vent"); py.FS.mkdirTree("/work/vendor");
    py.FS.mkdirTree("/py"); py.FS.mkdirTree("/assets");
    const manifest = await (await fetch("py/manifest.json")).json();
    for (const f of manifest) {
      const text = await (await fetch("py/" + f)).text();
      const dir = f.includes("/") ? "/py/" + f.split("/")[0] : "/py";
      py.FS.mkdirTree(dir);
      py.FS.writeFile("/py/" + f, text);
    }
    const tpl = new Uint8Array(await (await fetch("assets/template.xlsx")).arrayBuffer());
    py.FS.writeFile("/assets/template.xlsx", tpl);
    py.runPython("import sys; sys.path.insert(0, '/py')");
    api = py.pyimport("web_api");
    $("engine").textContent = "Python engine ready"; $("engine").classList.add("ok");
    refreshButtons();
  } catch (e) {
    console.error(e);
    $("engine").textContent = "Could not load the Python engine: " + e.message; $("engine").classList.add("err");
  }
}

const safe = (n) => n.replace(/[\\/:*?"<>|]/g, "_");

async function writeFiles(input, dir) {
  const paths = [];
  for (const f of input.files) {
    const p = `/work/${dir}/${safe(f.name)}`;
    py.FS.writeFile(p, new Uint8Array(await f.arrayBuffer()));
    paths.push(p);
  }
  return paths;
}

async function extractPdfWords(file) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const words = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    for (const it of tc.items) {
      const t = (it.str || "").trim();
      if (!t) continue;
      const x0 = it.transform[4];
      words.push({ page: p, x0, x1: x0 + it.width, top: vp.height - it.transform[5] - (it.height || 0), text: t });
    }
  }
  return words;
}

function refreshButtons() {
  const ready = $("fcu").files.length > 0 && $("vent").files.length > 0;
  $("analyze").disabled = !(api && ready);
  $("fallbackBox").classList.toggle("hidden", $("vendor").files.length > 0);
}

$("fcu").addEventListener("change", refreshButtons);
$("vent").addEventListener("change", refreshButtons);
$("vendor").addEventListener("change", async () => {
  refreshButtons();
  pdfWords = null;
  $("sheetRow").classList.add("hidden");
  const f = $("vendor").files[0];
  if (!f || !api) return;
  if (/\.pdf$/i.test(f.name)) return;
  try {
    setStatus($("status"), "Reading vendor workbook…");
    await tick();
    const [p] = await writeFiles($("vendor"), "vendor");
    const info = JSON.parse(api.vendor_sheets(p));
    const sel = $("vendorSheet");
    sel.innerHTML = "";
    for (const s of info.sheets) sel.add(new Option(s, s, false, s === info.default));
    $("sheetRow").classList.remove("hidden");
    setStatus($("status"), "");
  } catch (e) { setStatus($("status"), "Vendor file error: " + e.message, true); }
});

const checked = (id) => {
  const o = {};
  document.querySelectorAll(`#${id} input`).forEach((c) => { o[c.value] = c.checked; });
  return o;
};

function options() {
  const apply = checked("apply");
  const fb = checked("fbApply"); fb.power = true;
  return {
    pct: parseFloat($("pct").value) || 0, direction: $("dir").value, apply_to: apply,
    fallback_pct: parseFloat($("fbPct").value) || 0, fallback_direction: $("fbDir").value,
    fallback_apply_to: fb, qty: parseFloat($("qty").value),
    header: { plot: $("plot").value.trim(), project: $("project").value.trim(), consultant: $("consultant").value.trim() },
    hide_empty: $("hideEmpty").checked,
  };
}

$("analyze").addEventListener("click", async () => {
  const st = $("status");
  $("review").classList.add("hidden");
  const hasVendor = $("vendor").files.length > 0;
  if (!$("vent").files.length) { setStatus(st, "The ventilation calculation is required.", true); return; }
  if (!hasVendor && !(parseFloat($("qty").value) > 0)) { setStatus(st, "Enter the First Selection quantity (no vendor file).", true); return; }
  $("analyze").disabled = true;
  Object.keys(overridesState).forEach((k) => delete overridesState[k]);
  try {
    setStatus(st, "Reading files…"); await tick();
    const fcu = await writeFiles($("fcu"), "fcu");
    const vent = await writeFiles($("vent"), "vent");
    let vendor = "", words = "";
    if (hasVendor) {
      const f = $("vendor").files[0];
      if (/\.pdf$/i.test(f.name)) {
        setStatus(st, "Extracting text from the vendor PDF…"); await tick();
        words = JSON.stringify(await extractPdfWords(f));
      } else { vendor = (await writeFiles($("vendor"), "vendor"))[0]; }
    }
    setStatus(st, "Parsing and matching… (can take a few seconds)"); await tick();
    analysis = JSON.parse(api.analyze(fcu, vent, vendor, hasVendor ? $("vendorSheet").value : "", words, 0.8, 0.4));
    renderReview();
    setStatus(st, `${analysis.units.length} units found.`);
  } catch (e) {
    console.error(e);
    setStatus(st, "Error: " + String(e.message).split("\n").filter(Boolean).pop(), true);
  } finally { $("analyze").disabled = false; }
});

const fmt = (v, d = 2) => (v === null || v === undefined ? "–" : Number(v).toFixed(d).replace(/\.?0+$/, ""));

function vendorCell(u) {
  if (!u.vendor) return "<td>–</td>";
  const v = u.vendor;
  let opts = `<option value="">${v.best === null ? "(no match)" : "auto: " + esc(label(v, v.best))}</option>`;
  for (const c of v.cands) opts += `<option value="${c.i}">${esc(c.label)} (${c.s})</option>`;
  opts += `<option value="-1">— none —</option>`;
  const note = v.best !== null && v.score < 1 ? `<span class="note">similarity ${v.score}</span>` : "";
  return `<td><select data-u="${u.i}" data-k="vendor">${opts}</select>${note}</td>`;
}
const label = (v, i) => (v.cands.find((c) => c.i === i) || { label: "#" + i }).label;

function ventCell(u) {
  const v = u.vent;
  const auto = v.hits.length ? "auto: " + v.hits.map((h) => h.label).join(" + ") : "(no match)";
  let opts = `<option value="">${esc(auto)}</option>`;
  for (const c of v.cands) opts += `<option value="${c.i}">${esc(c.label)} (${c.s})</option>`;
  opts += `<option value="-1">— none —</option>`;
  const note = v.hits.length && v.score < 1 ? `<span class="note">similarity ${v.score}</span>` : "";
  return `<td><select data-u="${u.i}" data-k="vent">${opts}</select>${note}</td>`;
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function renderReview() {
  const hasVendor = analysis.units.some((u) => u.vendor);
  let head = "<tr><th>Unit</th><th>Location</th><th>Area</th><th>Calc kW</th><th>Flow L/s</th>";
  if (hasVendor) head += "<th>Vendor match</th>";
  head += "<th>Ventilation match</th></tr>";
  let rows = "";
  for (const u of analysis.units) {
    rows += `<tr><td>${esc(u.ref)}</td><td>${esc(u.zone)}</td>` +
      `<td class="num">${fmt(u.area, 1)}</td><td class="num">${fmt(u.total)}</td><td class="num">${fmt(u.flow, 0)}</td>` +
      (hasVendor ? vendorCell(u) : "") + ventCell(u) + "</tr>";
  }
  $("table").innerHTML = head + rows;
  const noVent = analysis.units.filter((u) => !u.vent.hits.length).length;
  $("summary").textContent = `${analysis.units.length} units · vent spaces: ${analysis.vent_count}` +
    (noVent ? ` · ${noVent} without ventilation match` : "") +
    (hasVendor ? ` · vendor rows: ${analysis.vendor_count}` : " · no vendor file");
  $("review").classList.remove("hidden");
}

const overridesState = {};
$("table").addEventListener("change", (e) => {
  const s = e.target;
  if (s.tagName !== "SELECT") return;
  const o = (overridesState[s.dataset.u] = overridesState[s.dataset.u] || {});
  if (s.value === "") delete o[s.dataset.k]; else o[s.dataset.k] = parseInt(s.value, 10);
});
$("generate").addEventListener("click", async () => {
  const st = $("genStatus");
  $("generate").disabled = true;
  try {
    setStatus(st, "Building workbook…"); await tick();
    const opts = options();
    const res = JSON.parse(api.build(JSON.stringify(opts), JSON.stringify(overridesState), "/work/out.xlsx"));
    const bytes = py.FS.readFile("/work/out.xlsx");
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    const base = opts.header.project ? safe(opts.header.project) + "_" : "";
    a.href = URL.createObjectURL(blob);
    a.download = `${base}AC-Unit-Schedule.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setStatus(st, `Done: ${res.rows} rows written.`);
    $("warnings").innerHTML = res.warnings.map((w) => `<li>${esc(w)}</li>`).join("");
  } catch (e) {
    console.error(e);
    setStatus(st, "Error: " + String(e.message).split("\n").filter(Boolean).pop(), true);
  } finally { $("generate").disabled = false; }
});

refreshButtons();
initEngine();
