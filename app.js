// ----------------------------
// GitHub Pages-safe absolute URLs to /data/*
// ----------------------------
const BASE = new URL(".", window.location.href).href; // always ends with "/"
const DEFAULT_FILES = {
  targetsCsv: new URL("data/targets.csv", BASE).toString(),
  pipelineCsv: new URL("data/pipeline.csv", BASE).toString(),
  workTicketsXlsx: new URL("data/work_tickets.xlsx", BASE).toString(),
};

// Status words
const TICKET_ACTIVE_STATUS_WORDS = ["open", "scheduled"];
const WON_STATUS_WORDS = ["won", "closed won", "sold"];

// ----------------------------
// Helpers
// ----------------------------
function setPathAndPill(pathEl, pillEl, ok, text) {
  pathEl.textContent = text.path;
  pillEl.textContent = text.status;
  pillEl.style.background = ok ? "rgba(34,197,94,0.15)" : "rgba(239,68,68,0.15)";
}

function isConstructionDivision(name) {
  return String(name || "").toLowerCase().includes("construction");
}

function isMaintenanceDivision(name) {
  const s = String(name || "").toLowerCase();
  return s.includes("maintenance") || s.includes("irrigation") || s.includes("lighting");
}

function isTicketActive(status) {
  const s = String(status || "").trim().toLowerCase();
  return TICKET_ACTIVE_STATUS_WORDS.some(w => s.includes(w));
}

function isWonPipelineStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  return WON_STATUS_WORDS.some(w => s.includes(w));
}

function parseCurrency(v) {
  if (v == null) return 0;
  const s = String(v).replace(/[\s,$]/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function parseNumberLoose(v) {
  if (v == null) return 0;
  const s = String(v).replace(/[\s,]/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function parseDateAny(v) {
  if (!v) return null;
  if (v instanceof Date && Number.isFinite(v.getTime())) return v;

  const s = String(v).trim();

  // ISO: YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const yy = Number(m[1]);
    const mm = Number(m[2]) - 1;
    const dd = Number(m[3]);
    const d = new Date(yy, mm, dd);
    return Number.isFinite(d.getTime()) ? d : null;
  }

  // US: MM/DD/YY or MM/DD/YYYY
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (m) {
    const mm = Number(m[1]) - 1;
    const dd = Number(m[2]);
    let yy = Number(m[3]);
    if (m[3].length === 2) yy += 2000;
    const d = new Date(yy, mm, dd);
    return Number.isFinite(d.getTime()) ? d : null;
  }

  return null;
}

// Excel date serials
function parseExcelDate(v) {
  if (!v) return null;
  if (v instanceof Date && Number.isFinite(v.getTime())) return v;

  if (typeof v === "number" && Number.isFinite(v)) {
    const o = XLSX.SSF.parse_date_code(v);
    if (!o) return null;
    return new Date(o.y, o.m - 1, o.d);
  }

  return parseDateAny(v);
}

function monthKeyFromDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function monthKeyFromTargetsLabel(label) {
  const s = String(label || "").trim();

  // Jan-26
  let m = s.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (m) {
    const mon = m[1].toLowerCase();
    const yy = Number(m[2]) + 2000;
    const idx = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"].indexOf(mon);
    if (idx >= 0) return `${yy}-${String(idx + 1).padStart(2, "0")}`;
  }

  // 2026-03 or 2026/03
  m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) {
    const yy = Number(m[1]);
    const mm = Number(m[2]);
    if (yy >= 2000 && mm >= 1 && mm <= 12) return `${yy}-${String(mm).padStart(2, "0")}`;
  }

  return null;
}

async function fetchText(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`Fetch failed ${res.status} for ${url}`);
  return await res.text();
}

// ----------------------------
// Loaders
// ----------------------------
async function loadTargets(url) {
  const text = await fetchText(url);

  const parsed = Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    delimiter: "", // auto detect (tab or comma)
    transformHeader: h => String(h || "").trim(),
  });

  const fields = parsed.meta?.fields || [];
  const findCol = (want) => fields.find(f => String(f).trim().toLowerCase() === want.toLowerCase()) || want;

  const COL = {
    Month: findCol("Month"),
    CH: findCol("Construction Hours"),
    CR: findCol("Construction Revenue"),
    MH: findCol("Maintenance Hours"),
    MR: findCol("Maintenance Revenue"),
  };

  const monthKeys = [];
  const constrHours = {};
  const constrRev = {};
  const maintHours = {};
  const maintRev = {};

  for (const r of parsed.data || []) {
    const mk = monthKeyFromTargetsLabel(r?.[COL.Month]);
    if (!mk) continue;
    monthKeys.push(mk);
    constrHours[mk] = parseNumberLoose(r[COL.CH]);
    constrRev[mk] = parseCurrency(r[COL.CR]);
    maintHours[mk] = parseNumberLoose(r[COL.MH]);
    maintRev[mk] = parseCurrency(r[COL.MR]);
  }

  monthKeys.sort();
  if (!monthKeys.length) throw new Error("No months recognized in targets.csv");

  return { monthKeys, constrHours, constrRev, maintHours, maintRev };
}

async function loadPipeline(url) {
  const text = await fetchText(url);
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true, delimiter: "" });

  const rows = (parsed.data || []).filter(r => r && Object.values(r).some(v => String(v ?? "").trim() !== ""));

  return rows.map(r => ({
    division: String(r["Division Name"] ?? "").trim(),
    status: String(r["Opp Status"] ?? "").trim(),
    startDate: parseDateAny(r["Start Date"]),
    weightedPipeline: parseCurrency(r["Weighted Pipeline"]),
    estimatedDollars: parseCurrency(r["Estimated $"]),
    weightedHours: parseNumberLoose(r["Weighted Hours"]),
  })).filter(r => r.startDate);
}

async function loadWorkTickets(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`Fetch failed ${res.status} for ${url}`);
  const arrayBuffer = await res.arrayBuffer();

  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

  if (!json.length) return [];

  const sample = json[0] || {};
  const keys = Object.keys(sample);

  // normalize header keys (remove periods + collapse spaces)
  const norm = (s) => String(s || "").trim().toLowerCase().replace(/\./g, "").replace(/\s+/g, " ");
  const pick = (...candidates) => {
    const cand = candidates.map(norm);
    const found = keys.find(k => cand.includes(norm(k)));
    return found || null;
  };

  // Your headers: Status, Sched Date, Est Hrs, Division
  const COL_STATUS = pick("status", "abr status", "ticket status", "work status");
  const COL_DATE   = pick("sched date", "scheduled date", "schedule date", "start date", "due date", "date");
  const COL_HRS    = pick("est hrs", "est h", "estimated hours", "hours", "labor hours", "est hrs");
  const COL_DIV    = pick("division", "division name", "department", "service");

  return json.map(r => ({
    status: String(COL_STATUS ? r[COL_STATUS] : "").trim(),
    schedDate: parseExcelDate(COL_DATE ? r[COL_DATE] : null),
    estHrs: parseNumberLoose(COL_HRS ? r[COL_HRS] : 0),
    division: String(COL_DIV ? r[COL_DIV] : "").trim(),
  })).filter(t => t.schedDate && t.estHrs > 0);
}

// ----------------------------
// Actuals (localStorage)
// ----------------------------
const ACTUALS_KEY = "dashboard_actuals_v2";

function loadActualsStore() {
  try {
    const raw = localStorage.getItem(ACTUALS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveActualsStore(store) {
  localStorage.setItem(ACTUALS_KEY, JSON.stringify(store));
}

function getActualsForMonth(store, mk) {
  return store[mk] || { constrRev: 0, constrHrs: 0, maintRev: 0, maintHrs: 0 };
}

function setActualsForMonth(store, mk, vals) {
  store[mk] = {
    constrRev: Number(vals.constrRev || 0),
    constrHrs: Number(vals.constrHrs || 0),
    maintRev: Number(vals.maintRev || 0),
    maintHrs: Number(vals.maintHrs || 0),
  };
}

function loadMonthActualsIntoInputs(store, mk) {
  const a = getActualsForMonth(store, mk);

  document.getElementById("actConstrRev").value = a.constrRev || "";
  document.getElementById("actConstrHrs").value = a.constrHrs || "";
  document.getElementById("actMaintRev").value = a.maintRev || "";
  document.getElementById("actMaintHrs").value = a.maintHrs || "";

  document.getElementById("actualsLoaded").textContent =
    `Month ${mk} • Construction: $${a.constrRev || 0} / ${a.constrHrs || 0} hrs • Maintenance: $${a.maintRev || 0} / ${a.maintHrs || 0} hrs`;
}

// ----------------------------
// Charts
// ----------------------------
let chartConstr = null;
let chartMaint = null;
let chartRevenueYear = null;
let chartMonthTracking = null;

function destroy(chart) { if (chart) chart.destroy(); return null; }

function bucketSumByMonth(items, getDate, getValue, monthKeys) {
  const out = {};
  for (const k of monthKeys) out[k] = 0;

  for (const it of items) {
    const d = getDate(it);
    if (!d) continue;
    const mk = monthKeyFromDate(d);
    if (!(mk in out)) continue;
    out[mk] += Number(getValue(it) || 0);
  }
  return out;
}

function sumMonths(seriesByMonth, monthKeys) {
  return monthKeys.reduce((acc, k) => acc + (seriesByMonth[k] || 0), 0);
}

function buildHoursChart(canvas, labels, targetLine, barsTickets, barsPipeline) {
  return new Chart(canvas.getContext("2d"), {
    data: {
      labels,
      datasets: [
        {
          type: "line",
          label: "Target Hours",
          data: targetLine,
          borderWidth: 2,
          pointRadius: 2,
          tension: 0.2,
          yAxisID: "y",
        },
        {
          type: "bar",
          label: "Work Tickets (Open/Scheduled)",
          data: barsTickets,
          stack: "stack1",
          yAxisID: "y",
        },
        {
          type: "bar",
          label: "Opportunities (Pipeline Weighted Hours)",
          data: barsPipeline,
          stack: "stack1",
          yAxisID: "y",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: {
        legend: { position: "top" },
        tooltip: { mode: "index", intersect: false },
      },
      scales: {
        x: { stacked: true },
        y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } },
      },
    },
  });
}

function buildRevenueYearChart(canvas, target, pipeUnweighted, pipeWeighted) {
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: ["Year"],
      datasets: [
        { label: "Target Revenue", data: [target] },
        { label: "Pipeline $ (Unweighted)", data: [pipeUnweighted] },
        { label: "Pipeline $ (Weighted)", data: [pipeWeighted] },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: { legend: { position: "top" } },
      scales: { y: { beginAtZero: true } },
    },
  });
}

function buildMonthTrackingChart(canvas, targetVals, actualVals) {
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: ["Revenue", "Hours"],
      datasets: [
        { label: "Target", data: targetVals },
        { label: "Actual", data: actualVals },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: { legend: { position: "top" } },
      scales: { y: { beginAtZero: true } },
    },
  });
}

// ----------------------------
// Rendering
// ----------------------------
function populateMonthSelect(monthKeys) {
  const sel = document.getElementById("asOfMonth");
  sel.innerHTML = "";
  for (const mk of monthKeys) {
    const opt = document.createElement("option");
    opt.value = mk;
    opt.textContent = mk;
    sel.appendChild(opt);
  }
  sel.value = monthKeys[monthKeys.length - 1];
}

function getScope(view) {
  if (view === "construction") {
    return {
      targetYearRev: (t) => sumMonths(t.constrRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.constrRev[mk] || 0,
      targetMonthHrs: (t, mk) => t.constrHours[mk] || 0,
      actualMonthRev: (a) => a.constrRev || 0,
      actualMonthHrs: (a) => a.constrHrs || 0,
      pipeFilter: (p) => isConstructionDivision(p.division),
    };
  }
  if (view === "maintenance") {
    return {
      targetYearRev: (t) => sumMonths(t.maintRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.maintRev[mk] || 0,
      targetMonthHrs: (t, mk) => t.maintHours[mk] || 0,
      actualMonthRev: (a) => a.maintRev || 0,
      actualMonthHrs: (a) => a.maintHrs || 0,
      pipeFilter: (p) => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division),
    };
  }
  return {
    targetYearRev: (t) => {
      let s = 0;
      for (const mk of t.monthKeys) s += (t.constrRev[mk] || 0) + (t.maintRev[mk] || 0);
      return s;
    },
    targetMonthRev: (t, mk) => (t.constrRev[mk] || 0) + (t.maintRev[mk] || 0),
    targetMonthHrs: (t, mk) => (t.constrHours[mk] || 0) + (t.maintHours[mk] || 0),
    actualMonthRev: (a) => (a.constrRev || 0) + (a.maintRev || 0),
    actualMonthHrs: (a) => (a.constrHrs || 0) + (a.maintHrs || 0),
    pipeFilter: (p) => isConstructionDivision(p.division) || isMaintenanceDivision(p.division),
  };
}

function renderAll(state) {
  const view = document.getElementById("viewToggle").value;
  const mk = document.getElementById("asOfMonth").value;
  const scope = getScope(view);

  const monthKeys = state.targets.monthKeys;

  // Work ticket and pipeline buckets
  const activeTickets = state.tickets.filter(t => isTicketActive(t.status));
  const pipePotential = state.pipeline.filter(p => !isWonPipelineStatus(p.status)); // exclude won for hours chart

  const ticketConstr = bucketSumByMonth(
    activeTickets.filter(t => isConstructionDivision(t.division)),
    t => t.schedDate,
    t => t.estHrs,
    monthKeys
  );

  const ticketMaint = bucketSumByMonth(
    activeTickets.filter(t => isMaintenanceDivision(t.division) && !isConstructionDivision(t.division)),
    t => t.schedDate,
    t => t.estHrs,
    monthKeys
  );

  const pipeConstr = bucketSumByMonth(
    pipePotential.filter(p => isConstructionDivision(p.division)),
    p => p.startDate,
    p => p.weightedHours,
    monthKeys
  );

  const pipeMaint = bucketSumByMonth(
    pipePotential.filter(p => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division)),
    p => p.startDate,
    p => p.weightedHours,
    monthKeys
  );

  // Build chart arrays
  const constrTarget = monthKeys.map(m => state.targets.constrHours[m] || 0);
  const maintTarget  = monthKeys.map(m => state.targets.maintHours[m] || 0);

  const constrTickets = monthKeys.map(m => ticketConstr[m] || 0);
  const maintTickets  = monthKeys.map(m => ticketMaint[m] || 0);

  const constrPipe = monthKeys.map(m => pipeConstr[m] || 0);
  const maintPipe  = monthKeys.map(m => pipeMaint[m] || 0);

  chartConstr = destroy(chartConstr);
  chartMaint = destroy(chartMaint);

  chartConstr = buildHoursChart(
    document.getElementById("chartConstruction"),
    monthKeys,
    constrTarget,
    constrTickets,
    constrPipe
  );

  chartMaint = buildHoursChart(
    document.getElementById("chartMaintenance"),
    monthKeys,
    maintTarget,
    maintTickets,
    maintPipe
  );

  // Year revenue KPI (view-specific)
  const inScopePipe = state.pipeline.filter(scope.pipeFilter);
  const pipeYearUnweighted = inScopePipe.reduce((acc, p) => acc + (p.estimatedDollars || 0), 0);
  const pipeYearWeighted = inScopePipe.reduce((acc, p) => acc + (p.weightedPipeline || 0), 0);
  const targetYearRev = scope.targetYearRev(state.targets);

  chartRevenueYear = destroy(chartRevenueYear);
  chartRevenueYear = buildRevenueYearChart(
    document.getElementById("chartRevenueYear"),
    targetYearRev,
    pipeYearUnweighted,
    pipeYearWeighted
  );

  // Month tracking (target vs actual)
  const actualsStore = loadActualsStore();
  const a = getActualsForMonth(actualsStore, mk);

  const tMonthRev = scope.targetMonthRev(state.targets, mk);
  const tMonthHrs = scope.targetMonthHrs(state.targets, mk);
  const aMonthRev = scope.actualMonthRev(a);
  const aMonthHrs = scope.actualMonthHrs(a);

  chartMonthTracking = destroy(chartMonthTracking);
  chartMonthTracking = buildMonthTrackingChart(
    document.getElementById("chartMonthTracking"),
    [tMonthRev, tMonthHrs],
    [aMonthRev, aMonthHrs]
  );

  document.getElementById("lastRefresh").textContent = new Date().toLocaleString();
}

// ----------------------------
// Load all data + UI wires
// ----------------------------
async function loadAllData(state) {
  const tPath = document.getElementById("targetsPath");
  const pPath = document.getElementById("pipelinePath");
  const wPath = document.getElementById("ticketsPath");

  const tPill = document.getElementById("targetsStatus");
  const pPill = document.getElementById("pipelineStatus");
  const wPill = document.getElementById("ticketsStatus");

  // Targets
  tPath.textContent = DEFAULT_FILES.targetsCsv;
  try {
    state.targets = await loadTargets(DEFAULT_FILES.targetsCsv);
    tPill.textContent = `Loaded (${state.targets.monthKeys.length} months)`;
    tPill.style.background = "rgba(34,197,94,0.15)";
    populateMonthSelect(state.targets.monthKeys);
  } catch (e) {
    console.error("Targets failed:", e);
    tPill.textContent = `Failed: ${e?.message || e}`;
    tPill.style.background = "rgba(239,68,68,0.15)";
    throw e;
  }

  // Pipeline
  pPath.textContent = DEFAULT_FILES.pipelineCsv;
  try {
    state.pipeline = await loadPipeline(DEFAULT_FILES.pipelineCsv);
    pPill.textContent = `Loaded (${state.pipeline.length})`;
    pPill.style.background = "rgba(34,197,94,0.15)";
  } catch (e) {
    console.error("Pipeline failed:", e);
    state.pipeline = [];
    pPill.textContent = `Failed: ${e?.message || e}`;
    pPill.style.background = "rgba(239,68,68,0.15)";
  }

  // Work Tickets
  wPath.textContent = DEFAULT_FILES.workTicketsXlsx;
  try {
    state.tickets = await loadWorkTickets(DEFAULT_FILES.workTicketsXlsx);
    wPill.textContent = `Loaded (${state.tickets.length})`;
    wPill.style.background = "rgba(34,197,94,0.15)";
  } catch (e) {
    console.error("Work tickets failed:", e);
    state.tickets = [];
    wPill.textContent = `Failed: ${e?.message || e}`;
    wPill.style.background = "rgba(239,68,68,0.15)";
  }

  loadMonthActualsIntoInputs(loadActualsStore(), document.getElementById("asOfMonth").value);
}

function wireControls(state) {
  document.getElementById("refreshBtn").onclick = async () => {
    await loadAllData(state);
    renderAll(state);
  };

  document.getElementById("viewToggle").onchange = () => renderAll(state);

  document.getElementById("asOfMonth").onchange = () => {
    loadMonthActualsIntoInputs(loadActualsStore(), document.getElementById("asOfMonth").value);
    renderAll(state);
  };
}

function wireActualsButtons(state) {
  document.getElementById("saveActualsBtn").onclick = () => {
    const mk = document.getElementById("asOfMonth").value;
    const store = loadActualsStore();

    setActualsForMonth(store, mk, {
      constrRev: document.getElementById("actConstrRev").value,
      constrHrs: document.getElementById("actConstrHrs").value,
      maintRev: document.getElementById("actMaintRev").value,
      maintHrs: document.getElementById("actMaintHrs").value,
    });

    saveActualsStore(store);
    loadMonthActualsIntoInputs(store, mk);
    renderAll(state);
  };

  document.getElementById("exportActualsBtn").onclick = () => {
    const store = loadActualsStore();
    const blob = new Blob([JSON.stringify(store, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "actuals.json";
    a.click();
    URL.revokeObjectURL(url);
  };

  document.getElementById("importActualsInput").onchange = async (ev) => {
    const file = ev.target.files?.[0];
    if (!file) return;
    try {
      const obj = JSON.parse(await file.text());
      saveActualsStore(obj);
      loadMonthActualsIntoInputs(obj, document.getElementById("asOfMonth").value);
      renderAll(state);
    } catch {
      alert("Invalid actuals.json");
    } finally {
      ev.target.value = "";
    }
  };

  document.getElementById("clearActualsBtn").onclick = () => {
    if (!confirm("Clear ALL actuals stored in this browser?")) return;
    localStorage.removeItem(ACTUALS_KEY);
    loadMonthActualsIntoInputs(loadActualsStore(), document.getElementById("asOfMonth").value);
    renderAll(state);
  };
}

// ----------------------------
// Init (NO top-level await)
// ----------------------------
(async function init() {
  const state = { targets: null, pipeline: [], tickets: [] };

  wireControls(state);
  wireActualsButtons(state);

  await loadAllData(state);
  renderAll(state);
})();

