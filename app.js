// ----------------------------
// FILE PATHS (GitHub Pages-friendly)
// ----------------------------
const DEFAULT_FILES = {
  targetsCsv: "./data/targets.csv",
  pipelineCsv: "./data/pipeline.csv",
  workTicketsXlsx: "./data/work_tickets.xlsx",
};

// Work ticket statuses treated as "active"
const TICKET_ACTIVE_STATUS_WORDS = ["open", "scheduled"];

// Pipeline statuses treated as "won" (excluded from the HOURS bars)
const WON_STATUS_WORDS = ["won", "approved", "sold", "closed won"];

// ----------------------------
// HELPERS
// ----------------------------
function isWonPipelineStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  if (!s) return false;
  return WON_STATUS_WORDS.some(w => s.includes(w));
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

function parseCurrency(v) {
  if (v == null) return 0;
  const s = String(v).replace(/[\s,$]/g, "").replace(/^\-$/, "0");
  if (s === "" || s === "-" || s === "—") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function parseNumberLoose(v) {
  if (v == null) return 0;
  const s = String(v).replace(/[\s,]/g, "").replace(/^\-$/, "0");
  if (s === "" || s === "-" || s === "—") return 0;
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

// Excel often stores dates as serial numbers
function parseExcelDate(v) {
  if (!v) return null;
  if (v instanceof Date && Number.isFinite(v.getTime())) return v;

  if (typeof v === "number" && Number.isFinite(v)) {
    const o = XLSX.SSF.parse_date_code(v);
    if (!o) return null;
    return new Date(o.y, o.m - 1, o.d);
  }

  return parseDateMDY(v);
}

function monthKeyFromDate(d) {
  if (!d) return null;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`; // YYYY-MM
}

function formatMoney(n) {
  return Number(n || 0).toLocaleString(undefined, { style: "currency", currency: "USD" });
}

function setPill(el, ok, text) {
  el.textContent = text;
  el.style.background = ok ? "rgba(34,197,94,0.15)" : "rgba(239,68,68,0.15)";
}

async function fetchText(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`Failed to fetch ${url} (${res.status})`);
  return await res.text();
}

// ----------------------------
// TARGETS LOADER (tab OR comma delimited)
// ----------------------------
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

  // 2026-01 or 2026/01
  m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) {
    const yy = Number(m[1]);
    const mm = Number(m[2]);
    if (yy >= 2000 && mm >= 1 && mm <= 12) return `${yy}-${String(mm).padStart(2, "0")}`;
  }

  return null;
}

async function loadTargets(url) {
  const text = await fetchText(url);

  const parsed = Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    delimiter: "", // auto-detect: fixes tab-delimited exports
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
  if (!monthKeys.length) throw new Error("targets.csv loaded but no months recognized.");

  return { monthKeys, constrHours, constrRev, maintHours, maintRev };
}

// ----------------------------
// PIPELINE LOADER
// ----------------------------
async function loadPipeline(url) {
  const text = await fetchText(url);
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true, delimiter: "" });

  const rows = (parsed.data || []).filter(r => r && Object.values(r).some(v => String(v ?? "").trim() !== ""));

  // Expected columns in pipeline.csv:
  // Division Name, Opp Status, Start Date, Weighted Pipeline, Estimated $, Weighted Hours
  return rows.map(r => ({
    division: String(r["Division Name"] ?? "").trim(),
    status: String(r["Opp Status"] ?? "").trim(),
    startDate: parseDateany(r["Start Date"]),
    weightedPipeline: parseCurrency(r["Weighted Pipeline"]),
    estimatedDollars: parseCurrency(r["Estimated $"]),
    weightedHours: parseNumberLoose(r["Weighted Hours"]),
  })).filter(r => r.startDate);
}

// ----------------------------
// WORK TICKETS LOADER (robust header mapping)
// ----------------------------
async function loadWorkTickets(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`Failed to fetch ${url} (${res.status})`);
  const arrayBuffer = await res.arrayBuffer();

  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
  console.log("XLSX raw rows:", json.length, "first row:", json[0]);
  const sample = json[0] || {};
  const keys = Object.keys(sample);
  const norm = (s) =>
  String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\./g, "")      // remove periods
    .replace(/\s+/g, " ");   // collapse whitespace
  const pick = (...candidates) => {
    const c = candidates.map(norm);
    const found = keys.find(k => c.includes(norm(k)));
    return found || null;
  };
const COL_STATUS = pick(
  "Abr Status", "Status", "Ticket Status", "Work Status", "WO Status"
);

const COL_DATE = pick(
  "Sched Date", "Schedule Date", "Scheduled Date", "Sched. Date", "Start Date", "Due Date", "Date"
);

const COL_HRS = pick(
  "Est Hrs", "Est. Hrs", "Estimated Hours", "Est Hours", "Estimated Hrs", "Hours", "Labor Hours"
);

const COL_DIV = pick(
  "Division", "Division Name", "Service", "Department", "Work Type", "Category"
);


  return json.map(r => ({
    status: String(COL_STATUS ? r[COL_STATUS] : "").trim(),
    schedDate: parseExcelDate(COL_DATE ? r[COL_DATE] : null),
    estHrs: parseNumberLoose(COL_HRS ? r[COL_HRS] : 0),
    division: String(COL_DIV ? r[COL_DIV] : "").trim(),
  })).filter(t => t.schedDate && t.estHrs > 0);
}

// ----------------------------
// ACTUALS (browser-only)
// ----------------------------
const ACTUALS_KEY = "dashboard_actuals_final_v1";

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
  return store;
}

function loadMonthActualsIntoInputs(store, mk) {
  const a = getActualsForMonth(store, mk);
  document.getElementById("actConstrRev").value = a.constrRev || "";
  document.getElementById("actConstrHrs").value = a.constrHrs || "";
  document.getElementById("actMaintRev").value = a.maintRev || "";
  document.getElementById("actMaintHrs").value = a.maintHrs || "";

  document.getElementById("actualsLoaded").textContent =
    `Month ${mk} • Construction: ${formatMoney(a.constrRev)} / ${a.constrHrs || 0} hrs • Maintenance: ${formatMoney(a.maintRev)} / ${a.maintHrs || 0} hrs`;
}

// ----------------------------
// CHARTS
// ----------------------------
let chartConstr = null;
let chartMaint = null;
let chartRevenueYear = null;
let chartMonthTracking = null;

function destroy(c) { if (c) c.destroy(); return null; }

function bucketSumByMonth(items, getDate, getValue, monthKeys) {
  const out = {};
  for (const k of monthKeys) out[k] = 0;
  for (const it of items) {
    const mk = monthKeyFromDate(getDate(it));
    if (!mk || !(mk in out)) continue;
    out[mk] += Number(getValue(it) || 0);
  }
  return out;
}

function sumMonths(seriesByMonth, monthKeys) {
  return monthKeys.reduce((acc, k) => acc + (seriesByMonth[k] || 0), 0);
}

// Mixed chart: Target line + stacked bars
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
// UI
// ----------------------------
function populateAsOfMonthSelect(monthKeys) {
  const sel = document.getElementById("asOfMonth");
  sel.innerHTML = "";
  for (const mk of monthKeys) {
    const opt = document.createElement("option");
    opt.value = mk;
    opt.textContent = mk;
    sel.appendChild(opt);
  }
  // default to latest month in targets
  sel.value = monthKeys[monthKeys.length - 1];
}

function getScope(view) {
  if (view === "construction") {
    return {
      inPipe: (p) => isConstructionDivision(p.division),
      targetYearRev: (targets) => sumMonths(targets.constrRev, targets.monthKeys),
      targetMonthRev: (targets, mk) => targets.constrRev[mk] || 0,
      targetMonthHrs: (targets, mk) => targets.constrHours[mk] || 0,
      actualMonthRev: (a) => a.constrRev || 0,
      actualMonthHrs: (a) => a.constrHrs || 0,
    };
  }
  if (view === "maintenance") {
    return {
      inPipe: (p) => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division),
      targetYearRev: (targets) => sumMonths(targets.maintRev, targets.monthKeys),
      targetMonthRev: (targets, mk) => targets.maintRev[mk] || 0,
      targetMonthHrs: (targets, mk) => targets.maintHours[mk] || 0,
      actualMonthRev: (a) => a.maintRev || 0,
      actualMonthHrs: (a) => a.maintHrs || 0,
    };
  }
  // all
  return {
    inPipe: (p) => isConstructionDivision(p.division) || isMaintenanceDivision(p.division),
    targetYearRev: (targets) => {
      let s = 0;
      for (const mk of targets.monthKeys) s += (targets.constrRev[mk] || 0) + (targets.maintRev[mk] || 0);
      return s;
    },
    targetMonthRev: (targets, mk) => (targets.constrRev[mk] || 0) + (targets.maintRev[mk] || 0),
    targetMonthHrs: (targets, mk) => (targets.constrHours[mk] || 0) + (targets.maintHours[mk] || 0),
    actualMonthRev: (a) => (a.constrRev || 0) + (a.maintRev || 0),
    actualMonthHrs: (a) => (a.constrHrs || 0) + (a.maintHrs || 0),
  };
}

// ----------------------------
// RENDER
// ----------------------------
function renderAll(state) {
  const view = document.getElementById("viewToggle").value;
  const asOfKey = document.getElementById("asOfMonth").value;

  const scope = getScope(view);
  const actualsStore = loadActualsStore();
  const a = getActualsForMonth(actualsStore, asOfKey);

  // ---------- Hours by month ----------
  const monthKeys = state.targets.monthKeys;

  const activeTickets = state.tickets.filter(t => isTicketActive(t.status));
  const pipePotential = state.pipeline.filter(p => !isWonPipelineStatus(p.status)); // exclude won for hours

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
    p => p.weightedHours, // weighted hours
    monthKeys
  );

  const pipeMaint = bucketSumByMonth(
    pipePotential.filter(p => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division)),
    p => p.startDate,
    p => p.weightedHours, // weighted hours
    monthKeys
  );

  const constrTargetArr = monthKeys.map(m => state.targets.constrHours[m] || 0);
  const maintTargetArr  = monthKeys.map(m => state.targets.maintHours[m] || 0);

  const constrTicketsArr = monthKeys.map(m => ticketConstr[m] || 0);
  const maintTicketsArr  = monthKeys.map(m => ticketMaint[m] || 0);

  const constrPipeArr = monthKeys.map(m => pipeConstr[m] || 0);
  const maintPipeArr  = monthKeys.map(m => pipeMaint[m] || 0);

  chartConstr = destroy(chartConstr);
  chartMaint = destroy(chartMaint);

  chartConstr = buildHoursChart(
    document.getElementById("chartConstruction"),
    monthKeys,
    constrTargetArr,
    constrTicketsArr,
    constrPipeArr
  );

  chartMaint = buildHoursChart(
    document.getElementById("chartMaintenance"),
    monthKeys,
    maintTargetArr,
    maintTicketsArr,
    maintPipeArr
  );

  // ---------- Revenue Year chart ----------
  const targetYearRev = scope.targetYearRev(state.targets);

  const inScopePipe = state.pipeline.filter(scope.inPipe);
  const pipeYearUnweighted = inScopePipe.reduce((acc, p) => acc + (p.estimatedDollars || 0), 0);
  const pipeYearWeighted   = inScopePipe.reduce((acc, p) => acc + (p.weightedPipeline || 0), 0);

  chartRevenueYear = destroy(chartRevenueYear);
  chartRevenueYear = buildRevenueYearChart(
    document.getElementById("chartRevenueYear"),
    targetYearRev,
    pipeYearUnweighted,
    pipeYearWeighted
  );

  // ---------- Tracking this month ----------
  const targetMonthRev = scope.targetMonthRev(state.targets, asOfKey);
  const targetMonthHrs = scope.targetMonthHrs(state.targets, asOfKey);
  const actualMonthRev = scope.actualMonthRev(a);
  const actualMonthHrs = scope.actualMonthHrs(a);

  chartMonthTracking = destroy(chartMonthTracking);
  chartMonthTracking = buildMonthTrackingChart(
    document.getElementById("chartMonthTracking"),
    [targetMonthRev, targetMonthHrs],
    [actualMonthRev, actualMonthHrs]
  );

  document.getElementById("lastRefresh").textContent = new Date().toLocaleString();
}

// ----------------------------
// LOAD + WIRES
// ----------------------------
async function loadAllData(state) {
  // Targets
 try {
  state.tickets = await loadWorkTickets(DEFAULT_FILES.workTicketsXlsx);
  setPill(document.getElementById("ticketsStatus"), true, `Loaded (${state.tickets.length})`);
} catch (e) {
  console.error("Work tickets failed:", e);
  state.tickets = [];
  const msg = (e && e.message) ? e.message : String(e);
  setPill(document.getElementById("ticketsStatus"), false, `Failed: ${msg}`);
}
    throw e;
  }

  // Pipeline
  try {
    state.pipeline = await loadPipeline(DEFAULT_FILES.pipelineCsv);
    setPill(document.getElementById("pipelineStatus"), true, `Loaded (${state.pipeline.length})`);
  } catch (e) {
    console.error(e);
    state.pipeline = [];
    setPill(document.getElementById("pipelineStatus"), false, "Missing/Failed");
  }

  // Tickets
  try {
    state.tickets = await loadWorkTickets(DEFAULT_FILES.workTicketsXlsx);
    setPill(document.getElementById("ticketsStatus"), true, `Loaded (${state.tickets.length})`);
  } catch (e) {
    console.error(e);
    state.tickets = [];
    setPill(document.getElementById("ticketsStatus"), false, "Missing/Failed");
    console.log("Work tickets loaded:", state.tickets.length, state.tickets.slice(0,3));
  }

  // Actuals inputs
  const mk = document.getElementById("asOfMonth").value;
  loadMonthActualsIntoInputs(loadActualsStore(), mk);
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

(async function init() {
  const state = { targets: null, pipeline: [], tickets: [] };
  wireControls(state);
  wireActualsButtons(state);

  await loadAllData(state);
  renderAll(state);

})();

