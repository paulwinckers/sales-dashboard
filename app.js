// ----------------------------
// CONFIG — EDIT THESE ONLY
// ----------------------------

// 1) Put your published Google Sheet CSV URL here (Publish to web → CSV)
const LOGBOOK_CSV_URL = "PASTE_YOUR_PUBLISHED_CSV_URL_HERE";

// 2) Tell the parser which columns to use from that CSV.
// If you already have separate columns for Construction/Maintenance actual revenue,
// set MODE="separate_columns" and configure those column names.
//
// If your logbook is transactional (many rows), set MODE="rows_with_division".
const LOGBOOK_MODE = "rows_with_division"; // "rows_with_division" or "separate_columns"

// Common: a Date column like 2026-03-15 OR a Month column like 2026-03
const LOGBOOK_DATE_COL = "Date";      // used in rows_with_division mode
const LOGBOOK_MONTH_COL = "Month";    // if you have yyyy-mm directly, use this

// Transactional mode:
const LOGBOOK_AMOUNT_COL = "Amount";
const LOGBOOK_DIVISION_COL = "Division"; // values containing "Construction" or maintenance keywords

// Separate-columns mode:
const LOGBOOK_CONSTR_ACTUAL_COL = "Construction Actual Revenue";
const LOGBOOK_MAINT_ACTUAL_COL  = "Maintenance Actual Revenue";

// Maintenance keywords used when parsing division strings
const MAINT_KEYWORDS = ["maintenance", "commercial maintenance", "residential maintenance", "irrigation", "lighting"];

// ----------------------------
// Static file paths (GitHub Pages safe)
// ----------------------------
const BASE = new URL(".", window.location.href).href;
const FILES = {
  targetsCsv: new URL("data/targets.csv", BASE).toString(),
  pipelineCsv: new URL("data/pipeline.csv", BASE).toString(),
  workTicketsXlsx: new URL("data/work_tickets.xlsx", BASE).toString(),
  capacityCsv: new URL("data/capacity.csv", BASE).toString(),
};

const TICKET_ACTIVE_STATUS_WORDS = ["open", "scheduled"];
const WON_STATUS_WORDS = ["won", "closed won", "sold"];

// ----------------------------
// Helpers
// ----------------------------
function isConstructionDivision(name) {
  return String(name || "").toLowerCase().includes("construction");
}
function isMaintenanceDivision(name) {
  const s = String(name || "").toLowerCase();
  return MAINT_KEYWORDS.some(k => s.includes(k));
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

  // ISO YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return Number.isFinite(d.getTime()) ? d : null;
  }

  // MM/DD/YY or MM/DD/YYYY
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (m) {
    let yy = Number(m[3]);
    if (m[3].length === 2) yy += 2000;
    const d = new Date(yy, Number(m[1]) - 1, Number(m[2]));
    return Number.isFinite(d.getTime()) ? d : null;
  }

  return null;
}
function monthKeyFromDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
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
function projectionForMonth(mk, actualMtd) {
  const now = new Date();
  const currentMk = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  if (mk !== currentMk) return actualMtd;

  const day = now.getDate();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  if (day <= 0) return actualMtd;
  return (actualMtd / day) * daysInMonth;
}

// ----------------------------
// Loaders
// ----------------------------
async function loadTargets(url) {
  const text = await fetchText(url);

  const parsed = Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    delimiter: "",
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

async function loadCapacity(url, monthKeys) {
  const text = await fetchText(url);
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true, delimiter: "" });

  const constCap = {};
  const maintCap = {};
  for (const mk of monthKeys) { constCap[mk] = 0; maintCap[mk] = 0; }

  for (const r of parsed.data || []) {
    const mk = String(r["month"] || "").trim();
    if (!mk || !(mk in constCap)) continue;
    constCap[mk] = parseNumberLoose(r["constcap"]);
    maintCap[mk] = parseNumberLoose(r["maintcap"]);
  }

  return { constCap, maintCap };
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

  const keys = Object.keys(json[0] || {});
  const norm = (s) => String(s || "").trim().toLowerCase().replace(/\./g, "").replace(/\s+/g, " ");
  const pick = (...cands) => {
    const cand = cands.map(norm);
    return keys.find(k => cand.includes(norm(k))) || null;
  };

  const COL_STATUS = pick("status", "abr status", "ticket status", "work status");
  const COL_DATE   = pick("sched date", "scheduled date", "schedule date", "start date", "due date", "date");
  const COL_HRS    = pick("est hrs", "estimated hours", "hours", "labor hours");
  const COL_DIV    = pick("division", "division name", "department", "service");

  return json.map(r => ({
    status: String(COL_STATUS ? r[COL_STATUS] : "").trim(),
    schedDate: parseDateAny(COL_DATE ? r[COL_DATE] : null) || parseExcelDate(COL_DATE ? r[COL_DATE] : null),
    estHrs: parseNumberLoose(COL_HRS ? r[COL_HRS] : 0),
    division: String(COL_DIV ? r[COL_DIV] : "").trim(),
  })).filter(t => t.schedDate && t.estHrs > 0);
}

// ---- Google Sheet (published CSV) actual revenue loader ----
async function loadActualRevenueFromLogbook(url, monthKeys) {
  const out = {};
  for (const mk of monthKeys) out[mk] = { constr: 0, maint: 0 };

  if (!url || url.includes("PASTE_YOUR_PUBLISHED")) return out;

  const text = await fetchText(url);
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true, delimiter: "" });

  const rows = parsed.data || [];

  if (LOGBOOK_MODE === "separate_columns") {
    for (const r of rows) {
      const mk = String(r[LOGBOOK_MONTH_COL] || "").trim();
      if (!mk || !(mk in out)) continue;
      out[mk].constr += parseCurrency(r[LOGBOOK_CONSTR_ACTUAL_COL]);
      out[mk].maint  += parseCurrency(r[LOGBOOK_MAINT_ACTUAL_COL]);
    }
    return out;
  }

  // rows_with_division mode
  for (const r of rows) {
    let mk = String(r[LOGBOOK_MONTH_COL] || "").trim();
    if (!mk) {
      const d = parseDateAny(r[LOGBOOK_DATE_COL]);
      if (d) mk = monthKeyFromDate(d);
    }
    if (!mk || !(mk in out)) continue;

    const amt = parseCurrency(r[LOGBOOK_AMOUNT_COL]);
    const div = String(r[LOGBOOK_DIVISION_COL] || "").toLowerCase();

    if (isConstructionDivision(div)) out[mk].constr += amt;
    else if (isMaintenanceDivision(div)) out[mk].maint += amt;
    else {
      // if no clear division, ignore (or change to allocate somewhere)
    }
  }

  return out;
}

// ----------------------------
// Charts
// ----------------------------
let chartConstr = null;
let chartMaint = null;
let chartRevenueYear = null;
let chartRevenuePace = null;

function destroy(chart) { if (chart) chart.destroy(); return null; }

function buildHoursChart(canvas, labels, targetLine, capLine, barsTickets, barsPipeline) {
  return new Chart(canvas.getContext("2d"), {
    data: {
      labels,
      datasets: [
        { type: "line", label: "Target Hours", data: targetLine, borderWidth: 2, pointRadius: 2, tension: 0.2 },

        // ✅ NEW capacity line
        { type: "line", label: "Capacity", data: capLine, borderWidth: 2, pointRadius: 0, tension: 0.2 },

        // ✅ Bright green tickets
        {
          type: "bar",
          label: "Work Tickets (Open/Scheduled)",
          data: barsTickets,
          stack: "stack1",
          backgroundColor: "rgba(34, 197, 94, 0.9)",
        },

        // ✅ Bright yellow opportunities
        {
          type: "bar",
          label: "Opportunities (Pipeline Weighted Hours)",
          data: barsPipeline,
          stack: "stack1",
          backgroundColor: "rgba(250, 204, 21, 0.9)",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: { legend: { position: "top" }, tooltip: { mode: "index", intersect: false } },
      scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } },
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
    options: { responsive: true, maintainAspectRatio: false, animation: false, scales: { y: { beginAtZero: true } } },
  });
}

function buildRevenuePaceChart(canvas, targetMonthRev, actualMonthRev, projectedFullMonthRev) {
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: ["Revenue"],
      datasets: [
        { label: "Target (Full Month)", data: [targetMonthRev] },
        { label: "Actual (From Logbook)", data: [actualMonthRev] },
        { label: "Projected (Full Month)", data: [projectedFullMonthRev] },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, animation: false, scales: { y: { beginAtZero: true } } },
  });
}

// ----------------------------
// UI & rendering
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

  const now = new Date();
  const currentMk = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  sel.value = monthKeys.includes(currentMk) ? currentMk : monthKeys[monthKeys.length - 1];
}

function getScope(view) {
  if (view === "construction") {
    return {
      targetYearRev: (t) => sumMonths(t.constrRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.constrRev[mk] || 0,
      actualMonthRev: (a) => a?.constr || 0,
      pipeFilter: (p) => isConstructionDivision(p.division),
    };
  }
  if (view === "maintenance") {
    return {
      targetYearRev: (t) => sumMonths(t.maintRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.maintRev[mk] || 0,
      actualMonthRev: (a) => a?.maint || 0,
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
    actualMonthRev: (a) => (a?.constr || 0) + (a?.maint || 0),
    pipeFilter: (p) => isConstructionDivision(p.division) || isMaintenanceDivision(p.division),
  };
}

function renderAll(state) {
  if (!state.targets?.monthKeys?.length) return;

  const view = document.getElementById("viewToggle").value || "all";
  const mk = document.getElementById("asOfMonth").value || state.targets.monthKeys[state.targets.monthKeys.length - 1];
  const scope = getScope(view);
  const monthKeys = state.targets.monthKeys;

  // KPI: Year revenue vs pipeline
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

  // KPI: Revenue pace (uses Logbook actuals)
  const targetMonthRev = scope.targetMonthRev(state.targets, mk);
  const actualMonthRev = scope.actualMonthRev(state.actualRevenueByMonth[mk]);
  const projectedRev = projectionForMonth(mk, actualMonthRev);

  chartRevenuePace = destroy(chartRevenuePace);
  chartRevenuePace = buildRevenuePaceChart(
    document.getElementById("chartRevenuePace"),
    targetMonthRev,
    actualMonthRev,
    projectedRev
  );

  // Hours by month: tickets + pipeline (exclude won)
  const activeTickets = state.tickets.filter(t => isTicketActive(t.status));
  const pipePotential = state.pipeline.filter(p => !isWonPipelineStatus(p.status));

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

  const constrTarget = monthKeys.map(m => state.targets.constrHours[m] || 0);
  const maintTarget  = monthKeys.map(m => state.targets.maintHours[m] || 0);

  const constrCap = monthKeys.map(m => state.capacity.constCap[m] || 0);
  const maintCap  = monthKeys.map(m => state.capacity.maintCap[m] || 0);

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
    constrCap,
    constrTickets,
    constrPipe
  );

  chartMaint = buildHoursChart(
    document.getElementById("chartMaintenance"),
    monthKeys,
    maintTarget,
    maintCap,
    maintTickets,
    maintPipe
  );

  document.getElementById("lastRefresh").textContent = new Date().toLocaleString();
}

// ----------------------------
// Data load + wiring
// ----------------------------
async function loadAllData(state) {
  const setStatus = (id, ok, msg) => {
    const pill = document.getElementById(id);
    if (!pill) return;
    pill.textContent = msg;
    pill.style.background = ok ? "rgba(34,197,94,0.15)" : "rgba(239,68,68,0.15)";
  };

  // Paths
  document.getElementById("targetsPath").textContent = FILES.targetsCsv;
  document.getElementById("pipelinePath").textContent = FILES.pipelineCsv;
  document.getElementById("ticketsPath").textContent = FILES.workTicketsXlsx;

  // Targets
  try {
    state.targets = await loadTargets(FILES.targetsCsv);
    populateMonthSelect(state.targets.monthKeys);
    setStatus("targetsStatus", true, `Loaded (${state.targets.monthKeys.length} months)`);
  } catch (e) {
    console.error(e);
    setStatus("targetsStatus", false, `Failed: ${e.message || e}`);
    throw e;
  }

  // Capacity (depends on monthKeys)
  try {
    state.capacity = await loadCapacity(FILES.capacityCsv, state.targets.monthKeys);
  } catch (e) {
    console.error("Capacity failed:", e);
    // still render with zeros
    state.capacity = { constCap: {}, maintCap: {} };
  }

  // Pipeline
  try {
    state.pipeline = await loadPipeline(FILES.pipelineCsv);
    setStatus("pipelineStatus", true, `Loaded (${state.pipeline.length})`);
  } catch (e) {
    console.error(e);
    state.pipeline = [];
    setStatus("pipelineStatus", false, `Failed: ${e.message || e}`);
  }

  // Work tickets
  try {
    state.tickets = await loadWorkTickets(FILES.workTicketsXlsx);
    setStatus("ticketsStatus", true, `Loaded (${state.tickets.length})`);
  } catch (e) {
    console.error(e);
    state.tickets = [];
    setStatus("ticketsStatus", false, `Failed: ${e.message || e}`);
  }

  // Logbook actual revenue
  try {
    state.actualRevenueByMonth = await loadActualRevenueFromLogbook(LOGBOOK_CSV_URL, state.targets.monthKeys);
  } catch (e) {
    console.error("Logbook failed:", e);
    // default to zeros
    state.actualRevenueByMonth = {};
    for (const mk of state.targets.monthKeys) state.actualRevenueByMonth[mk] = { constr: 0, maint: 0 };
  }
}

function wireControls(state) {
  document.getElementById("refreshBtn").addEventListener("click", async () => {
    await loadAllData(state);
    renderAll(state);
  });

  document.getElementById("viewToggle").addEventListener("change", () => renderAll(state));
  document.getElementById("asOfMonth").addEventListener("change", () => renderAll(state));
}

// ----------------------------
// Init
// ----------------------------
(async function init() {
  const state = {
    targets: null,
    capacity: { constCap: {}, maintCap: {} },
    pipeline: [],
    tickets: [],
    actualRevenueByMonth: {},
  };

  wireControls(state);
  await loadAllData(state);
  renderAll(state);
})();




