"use strict";
const $ = (id) => document.getElementById(id);
const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const PDF_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

let py = null, api = null, analysis = null;
const overridesState = {};
const reviewed = new Set();
const KINDS = ["fcu", "people", "vendor", "vent"];
const files = { fcu: [], people: [], vendor: [], vent: [] };      // [{id, file, path?, sheets?, sheet?}]
let nextId = 1;

const setStatus = (el, msg, err) => { el.textContent = msg || ""; el.classList.toggle("err", !!err); };
const tick = () => new Promise((r) => setTimeout(r, 30));      // let the browser repaint before heavy work
const safe = (n) => n.replace(/[\\/:*?"<>|]/g, "_");
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fmt = (v, d = 2) => (v === null || v === undefined ? "–" : Number(v).toFixed(d).replace(/\.?0+$/, ""));
const isPdf = (f) => /\.pdf$/i.test(f.name);

/* ---------------------------------------------------------------- Python engine */
async function initEngine() {
  try {
    py = await loadPyodide({ indexURL: PYODIDE_URL });
    await py.loadPackage("micropip");
    await py.pyimport("micropip").install("openpyxl");
    for (const d of ["/work/fcu", "/work/vent", "/work/vendor", "/py", "/assets"]) py.FS.mkdirTree(d);
    const manifest = await (await fetch("py/manifest.json")).json();
    for (const f of manifest) {
      const text = await (await fetch("py/" + f)).text();
      py.FS.mkdirTree(f.includes("/") ? "/py/" + f.split("/")[0] : "/py");
      py.FS.writeFile("/py/" + f, text);
    }
    py.FS.writeFile("/assets/template.xlsx", new Uint8Array(await (await fetch("assets/template.xlsx")).arrayBuffer()));
    py.runPython("import sys; sys.path.insert(0, '/py')");
    api = py.pyimport("web_api");
    $("engine").textContent = "Python engine ready"; $("engine").classList.add("ok");
    refreshButtons();
    for (const e of files.vendor) await loadSheets(e);          // files added while the engine was loading
  } catch (e) {
    console.error(e);
    $("engine").textContent = "Could not load the Python engine: " + e.message; $("engine").classList.add("err");
  }
}

async function writeEntry(dir, entry) {
  entry.path = `/work/${dir}/${entry.id}_${safe(entry.file.name)}`;
  py.FS.writeFile(entry.path, new Uint8Array(await entry.file.arrayBuffer()));
  return entry.path;
}

/* ---------------------------------------------------------------- file lists */
function renderGroup(kind) {
  const ul = document.querySelector(`.group[data-kind="${kind}"] ul.files`);
  ul.innerHTML = "";
  for (const e of files[kind]) {
    const li = document.createElement("li");
    const kb = Math.max(1, Math.round(e.file.size / 1024));
    li.innerHTML = `<span class="fname">${esc(e.file.name)}</span><span class="fsize">${kb} KB</span>`;
    if (kind === "vendor" && !isPdf(e.file)) {
      const sel = document.createElement("select");
      sel.className = "sheet";
      sel.title = "Sheet to read";
      if (e.sheets) for (const s of e.sheets) sel.add(new Option(s, s, false, s === e.sheet));
      else sel.add(new Option("reading sheets…", ""));
      sel.addEventListener("change", () => { e.sheet = sel.value; });
      li.appendChild(sel);
    }
    const x = document.createElement("button");
    x.type = "button"; x.className = "x"; x.textContent = "×"; x.title = "Remove";
    x.addEventListener("click", () => { files[kind] = files[kind].filter((f) => f !== e); renderGroup(kind); refreshButtons(); });
    li.appendChild(x);
    ul.appendChild(li);
  }
}

async function loadSheets(entry) {
  if (!api || entry.sheets || isPdf(entry.file)) return;
  try {
    const p = await writeEntry("vendor", entry);
    const info = JSON.parse(api.vendor_sheets(p));
    entry.sheets = info.sheets; entry.sheet = info.default;
  } catch (e) {
    entry.sheets = []; setStatus($("status"), `Vendor file ${entry.file.name}: ${String(e.message).split("\n").filter(Boolean).pop()}`, true);
  }
  renderGroup("vendor");
}

document.querySelectorAll(".group").forEach((g) => {
  const kind = g.dataset.kind, picker = g.querySelector(".picker");
  g.querySelector(".add").addEventListener("click", () => picker.click());
  picker.addEventListener("change", async () => {
    const added = [...picker.files].map((file) => ({ id: nextId++, file }));
    picker.value = "";                                           // allows adding the same file again
    files[kind].push(...added);
    renderGroup(kind); refreshButtons();
    if (kind === "vendor") for (const e of added) await loadSheets(e);
  });
});

document.querySelectorAll('input[name="pdef"]').forEach((r) => r.addEventListener("change", () => {
  $("pdefVal").disabled = document.querySelector('input[name="pdef"]:checked').value !== "value";
}));

function refreshButtons() {
  $("analyze").disabled = !(api && files.fcu.length > 0);
  $("fallbackBox").classList.toggle("hidden", files.vendor.length > 0);
}

/* ---------------------------------------------------------------- PDF text (pdf.js) */
async function pdfDocument(file) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
  return pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
}

async function extractPdfWords(file) {                           // vendor PDF: positioned words
  const pdf = await pdfDocument(file);
  const words = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    for (const it of (await page.getTextContent()).items) {
      const t = (it.str || "").trim();
      if (!t) continue;
      const x0 = it.transform[4];
      words.push({ page: p, x0, x1: x0 + it.width, top: vp.height - it.transform[5] - (it.height || 0), text: t });
    }
  }
  return words;
}

function pageLines(items) {                                      // text items -> lines (top to bottom, left to right)
  const rows = [];
  for (const it of items) {
    const t = it.str || "";
    if (!t.trim()) continue;
    const y = it.transform[5], x = it.transform[4];
    let row = rows.find((r) => Math.abs(r.y - y) < 2.5);
    if (!row) { row = { y, cells: [] }; rows.push(row); }
    row.cells.push({ x, t });
  }
  rows.sort((a, b) => b.y - a.y);
  return rows.map((r) => r.cells.sort((a, b) => a.x - b.x).map((c) => c.t).join(" ")).join("\n");
}

async function extractPdfPages(file, onProgress) {               // people PDF: text of every page
  const pdf = await pdfDocument(file);
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    pages.push(pageLines((await (await pdf.getPage(p)).getTextContent()).items));
    if (p % 10 === 0 || p === pdf.numPages) { onProgress(p, pdf.numPages); await tick(); }
  }
  return pages;
}

/* ---------------------------------------------------------------- options */
const checked = (id) => {
  const o = {};
  document.querySelectorAll(`#${id} input`).forEach((c) => { o[c.value] = c.checked; });
  return o;
};

function options() {
  const fb = checked("fbApply"); fb.power = true;
  const useVal = document.querySelector('input[name="pdef"]:checked').value === "value";
  const pv = parseFloat($("pdefVal").value);
  return {
    pct: parseFloat($("pct").value) || 0, direction: $("dir").value, apply_to: checked("apply"),
    fallback_pct: parseFloat($("fbPct").value) || 0, fallback_direction: $("fbDir").value,
    fallback_apply_to: fb, qty: parseFloat($("qty").value),
    people_default: useVal && !isNaN(pv) ? pv : null,
    header: { plot: $("plot").value.trim(), project: $("project").value.trim(), consultant: $("consultant").value.trim() },
    hide_empty: $("hideEmpty").checked,
  };
}

/* ---------------------------------------------------------------- analyze */
$("analyze").addEventListener("click", async () => {
  const st = $("status");
  $("review").classList.add("hidden");
  if (!files.fcu.length) { setStatus(st, "Add at least one FCU schedule.", true); return; }
  if (!files.vendor.length && !(parseFloat($("qty").value) > 0)) { setStatus(st, "Enter the First Selection quantity (no vendor file).", true); return; }
  $("analyze").disabled = true;
  Object.keys(overridesState).forEach((k) => delete overridesState[k]);
  try {
    setStatus(st, "Reading files…"); await tick();
    const fcu = [], vent = [];
    for (const e of files.fcu) fcu.push(await writeEntry("fcu", e));
    for (const e of files.vent) vent.push(await writeEntry("vent", e));
    const people = [];
    for (const e of files.people) {
      setStatus(st, `Reading ${e.file.name}…`); await tick();
      const pages = await extractPdfPages(e.file, (p, n) => setStatus(st, `Reading ${e.file.name}: page ${p} of ${n}`));
      people.push({ pages, name: e.file.name });
    }
    const vendor = [];
    for (const e of files.vendor) {
      if (isPdf(e.file)) {
        setStatus(st, `Extracting text from ${e.file.name}…`); await tick();
        vendor.push({ words: await extractPdfWords(e.file), name: e.file.name });
      } else {
        if (!e.path) await writeEntry("vendor", e);
        vendor.push({ path: e.path, sheet: e.sheet || null });
      }
    }
    setStatus(st, "Parsing and matching… (can take a few seconds)"); await tick();
    analysis = JSON.parse(api.analyze(fcu, vent, JSON.stringify(vendor), JSON.stringify(people), 0.8, 0.4));
    renderReview();
    setStatus(st, `${analysis.units.length} units found.`);
  } catch (e) {
    console.error(e);
    setStatus(st, "Error: " + String(e.message).split("\n").filter(Boolean).pop(), true);
  } finally { $("analyze").disabled = false; }
});

/* ---------------------------------------------------------------- review table */
const label = (v, i) => (v.cands.find((c) => c.i === i) || { label: "#" + i }).label;

function vendorCell(u) {
  if (!u.vendor) return "<td>–</td>";
  const v = u.vendor;
  let opts = `<option value="">${v.best === null ? "(no match)" : "auto: " + esc(label(v, v.best))}</option>`;
  for (const c of v.cands) opts += `<option value="${c.i}">${esc(c.label)} (${c.s})</option>`;
  opts += `<option value="-1">— none —</option>`;
  const note = v.best !== null && v.score < 1 ? `<span class="note">similarity ${v.score}</span>` : "";
  return `<td><select data-u="${u.i}" data-k="vendor">${opts}</select>${note}</td>`;
}

function peopleCell(u) {
  const p = u.people;
  if (!p) return `<td class="note">${hasDefault() ? "default value" : "–"}</td>`;
  const auto = p.hits.length ? "auto: " + p.hits.map((h) => h.label).join(" + ") : "(no match)";
  let opts = `<option value="">${esc(auto)}</option>`;
  for (const c of p.cands) opts += `<option value="${c.i}">${esc(c.label)} (${c.s})</option>`;
  opts += `<option value="-1">— none —</option>`;
  return `<td><select data-u="${u.i}" data-k="people">${opts}</select><span class="note" data-pnote="${u.i}">${peopleNote(u)}</span></td>`;
}

const hasDefault = () => options().people_default !== null;

function peopleNote(u) {
  const o = overridesState[u.i] || {};
  if (o.people !== undefined) {
    if (o.people < 0) return "no people value";
    const c = u.people.cands.find((c) => c.i === o.people);
    return c ? `= ${fmt(c.p)} people` : "";
  }
  if (u.people && u.people.hits.length) return `= ${fmt(u.people.value)} people` + (u.people.hits.some((h) => h.s < 1) ? ` · similarity ${Math.min(...u.people.hits.map((h) => h.s))}` : "");
  return hasDefault() ? `default ${fmt(options().people_default)} people will be used` : "";
}

function actionText(u) {
  const o = overridesState[u.i] || {};
  const msgs = [];
  if (u.vendor && o.vendor === undefined) {
    if (u.vendor.best === null) msgs.push("No vendor match – choose one");
    else msgs.push(u.vendor.score < 1 ? "Check vendor match (fuzzy " + u.vendor.score + ")" : "Confirm vendor match");
  }
  if (u.people && o.people === undefined) {
    if (!u.people.hits.length) msgs.push(hasDefault() ? "No people match – default value used" : "No people match – choose one");
    else msgs.push(u.people.hits.some((h) => h.s < 1) ? "Check people match" : "Confirm people match");
  }
  if (!msgs.length) msgs.push(u.people || u.vendor ? "Confirm values" : "Confirm unit");
  return msgs;
}

const actionHtml = (u) => actionText(u).map((m) => `<b>${esc(m)}</b>`).join("");

function renderReview() {
  reviewed.clear();
  const hasVendor = analysis.units.some((u) => u.vendor);
  const hasPeople = analysis.units.some((u) => u.people);
  let head = '<tr><th class="chk"><input type="checkbox" id="allOk" title="Mark all as reviewed"></th><th>Action required</th><th>Unit</th><th>Location</th><th>Area</th><th>Calc kW</th><th>Flow L/s</th>';
  if (hasVendor) head += "<th>Vendor match</th>";
  head += "<th>No. of People</th></tr>";
  let rows = "";
  for (const u of analysis.units) {
    rows += `<tr class="review" data-row="${u.i}"><td class="chk"><input type="checkbox" class="ok" data-u="${u.i}" title="Mark as reviewed"></td>` +
      `<td class="action">${actionHtml(u)}</td>` +
      `<td>${esc(u.ref)}${u.subs ? `<span class="note">${u.subs} sub rows</span>` : ""}</td><td>${esc(u.zone)}</td>` +
      `<td class="num">${fmt(u.area, 1)}</td><td class="num">${fmt(u.total)}</td><td class="num">${fmt(u.flow, 0)}</td>` +
      (hasVendor ? vendorCell(u) : "") + (hasPeople ? peopleCell(u) : `<td class="note">${hasDefault() ? "default value" : "–"}</td>`) + "</tr>";
  }
  $("table").innerHTML = head + rows;
  const w = $("fcuWarnings");
  w.innerHTML = analysis.warnings.map((m) => `<li>${esc(m)}</li>`).join("");
  w.classList.toggle("hidden", !analysis.warnings.length);
  updateSummary();
  $("review").classList.remove("hidden");
}

function setReviewed(i, on) {
  const tr = document.querySelector(`tr[data-row="${i}"]`);
  if (!tr) return;
  tr.classList.toggle("review", !on);
  tr.querySelector("input.ok").checked = on;
  tr.querySelector("td.action").innerHTML = on ? "Reviewed" : actionHtml(analysis.units[i]);
  if (on) reviewed.add(i); else reviewed.delete(i);
}

function updateSummary() {
  const hasVendor = analysis.units.some((u) => u.vendor);
  const noPeople = analysis.units.filter((u) => u.people && !u.people.hits.length && (overridesState[u.i] || {}).people === undefined).length;
  const left = analysis.units.length - reviewed.size;
  $("summary").textContent = `${analysis.units.length} units · ${left} still to review` +
    (analysis.people_count ? ` · people sources: ${analysis.pdf_people_count} HAP + ${analysis.people_count - analysis.pdf_people_count} vent` : "") +
    (noPeople ? ` · ${noPeople} without people match` : "") +
    (hasVendor ? ` · vendor rows: ${analysis.vendor_count}` : " · no vendor file");
}

$("table").addEventListener("change", (e) => {
  const s = e.target;
  if (s.tagName !== "SELECT") return;
  const i = parseInt(s.dataset.u, 10);
  const o = (overridesState[i] = overridesState[i] || {});
  if (s.value === "") delete o[s.dataset.k]; else o[s.dataset.k] = parseInt(s.value, 10);
  const n = document.querySelector(`[data-pnote="${i}"]`);
  if (n) n.textContent = peopleNote(analysis.units[i]);
  setReviewed(i, true);                                          // picking a value counts as reviewing the row
  updateSummary();
});

$("table").addEventListener("click", (e) => {
  const c = e.target;
  if (c.tagName !== "INPUT" || c.type !== "checkbox") return;
  if (c.id === "allOk") analysis.units.forEach((u) => setReviewed(u.i, c.checked));
  else if (c.classList.contains("ok")) setReviewed(parseInt(c.dataset.u, 10), c.checked);
  updateSummary();
});

/* ---------------------------------------------------------------- generate */
$("generate").addEventListener("click", async () => {
  const st = $("genStatus");
  $("generate").disabled = true;
  try {
    setStatus(st, "Building workbook…"); await tick();
    const opts = options();
    const res = JSON.parse(api.build(JSON.stringify(opts), JSON.stringify(overridesState), "/work/out.xlsx"));
    const blob = new Blob([py.FS.readFile("/work/out.xlsx")], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${opts.header.project ? safe(opts.header.project) + "_" : ""}AC-Unit-Schedule.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setStatus(st, `Done: ${res.rows} rows written.`);
    $("warnings").innerHTML = res.warnings.map((w) => `<li>${esc(w)}</li>`).join("");
  } catch (e) {
    console.error(e);
    setStatus(st, "Error: " + String(e.message).split("\n").filter(Boolean).pop(), true);
  } finally { $("generate").disabled = false; }
});

KINDS.forEach(renderGroup);
refreshButtons();
initEngine();
