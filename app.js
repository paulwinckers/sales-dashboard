// ------------------------------------------------------------
// CONFIG
// ------------------------------------------------------------
const LOGBOOK_URL =
  "https://script.google.com/macros/s/AKfycbxemOcHaO8jJL2JNvr6G3INrHOSahH3-1QYcsrb5IV19DG77lPUPtDkco_s9r8RFwmI/exec";

// SalesAct headers (you provided):
// Month, ActualConstRevMTD, ActualMaintRevMTD, ActualConstrHoursMTD, ActualMaintHoursMTD, UpdatedAt, UpdatedBy

// ------------------------------------------------------------
// Static file paths (GitHub Pages safe)
// ------------------------------------------------------------
const BASE = new URL(".", window.location.href).href;
const FILES = {
  targetsCsv: new URL("data/targets.csv", BASE).toString(),
  pipelineCsv: new URL("data/pipeline.csv", BASE).toString(),
  workTicketsXlsx: new URL("data/work_tickets.xlsx", BASE).toString(),
  capacityCsv: new URL("data/capacity.csv", BASE).toString(),
};

const TICKET_ACTIVE_STATUS_WORDS = ["open", "scheduled"];
const WON_STATUS_WORDS = ["won", "closed won", "sold"];

// Chart colors
const BLUE = "rgba(54, 162, 235, 1)";
const GREEN = "rgba(34, 197, 94, 0.9)";      // bright green
const YELLOW = "rgba(250, 204, 21, 0.9)";    // bright yellow

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------
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

  // Already yyyy-mm
  if (/^\d{4}-\d{2}$/.test(s)) return s;

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

// Spread maintenance pipeline weighted hours from Start Date month -> November inclusive
function bucketMaintenancePipelineSpread(pipelineRows, monthKeys) {
  const out = {};
  for (const k of monthKeys) out[k] = 0;

  for (const p of pipelineRows) {
    if (!p.startDate) continue;

    const y = p.startDate.getFullYear();
    const startM = p.startDate.getMonth() + 1;
    const endM = 11; // November
    const total = Number(p.weightedHours || 0);

    if (startM > endM) {
      const mk = monthKeyFromDate(p.startDate);
      if (mk in out) out[mk] += total;
      continue;
    }

    const monthsCount = (endM - startM + 1);
    const perMonth = total / monthsCount;

    for (let m = startM; m <= endM; m++) {
      const mk = `${y}-${String(m).padStart(2, "0")}`;
      if (mk in out) out[mk] += perMonth;
    }
  }

  return out;
}

// ------------------------------------------------------------
// Loaders
// ------------------------------------------------------------
async function loadTargets(url) {
  const text = await fetchText(url);

  const parsed = Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    delimiter: "",
    transformHeader: h => String(h || "").trim(),
  });

  const fields = parsed.meta?.fields || [];
  const findCol = (want) =>
    fields.find(f => String(f).trim().toLowerCase() === want.toLowerCase()) || want;

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
    const mk = String(r.month || r.Month || "").trim();
    if (!mk || !(mk in constCap)) continue;
    constCap[mk] = parseNumberLoose(r.constcap ?? r.ConstCap ?? r.Constcap);
    maintCap[mk] = parseNumberLoose(r.maintcap ?? r.MaintCap ?? r.Maintcap);
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
    schedDate: parseExcelDate(COL_DATE ? r[COL_DATE] : null),
    estHrs: parseNumberLoose(COL_HRS ? r[COL_HRS] : 0),
    division: String(COL_DIV ? r[COL_DIV] : "").trim(),
  })).filter(t => t.schedDate && t.estHrs > 0);
}

async function loadSalesActMonthly(baseUrl, monthKeys) {
  // Returns: byMonth[yyyy-mm] = { constrRevMTD, maintRevMTD, constrHrsMTD, maintHrsMTD, updatedAt, updatedBy }
  const out = {};
  for (const mk of monthKeys) {
    out[mk] = { constrRevMTD: 0, maintRevMTD: 0, constrHrsMTD: 0, maintHrsMTD: 0, updatedAt: "", updatedBy: "" };
  }

  const url = `${baseUrl}?tab=SalesAct`;
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`SalesAct HTTP ${res.status}`);

  const payload = await res.json();
  const rows = payload?.rows || [];

  for (const r of rows) {
    const mk = String(r.Month || "").trim();
    if (!mk || !(mk in out)) continue;

    out[mk].constrRevMTD = parseCurrency(r.ActualConstRevMTD);
    out[mk].maintRevMTD = parseCurrency(r.ActualMaintRevMTD);
    out[mk].constrHrsMTD = parseNumberLoose(r.ActualConstrHoursMTD);
    out[mk].maintHrsMTD = parseNumberLoose(r.ActualMaintHoursMTD);
    out[mk].updatedAt = r.UpdatedAt ? String(r.UpdatedAt) : "";
    out[mk].updatedBy = r.UpdatedBy ? String(r.UpdatedBy) : "";
  }

  return out;
}

// ------------------------------------------------------------
// Charts
// ------------------------------------------------------------
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
        // ✅ Target line BLUE
        {
          type: "line",
          label: "Target Hours",
          data: targetLine,
          borderWidth: 2,
          pointRadius: 2,
          tension: 0.2,
          borderColor: BLUE,
          pointBackgroundColor: BLUE,
        },
        // ✅ Capacity line dashed BLUE
        {
          type: "line",
          label: "Capacity",
          data: capLine,
          borderWidth: 2,
          pointRadius: 0,
          tension: 0.2,
          borderColor: BLUE,
          borderDash: [6, 6],
        },
        // ✅ Work tickets bright green
        {
          type: "bar",
          label: "Work Tickets (Open/Scheduled)",
          data: barsTickets,
          stack: "stack1",
          backgroundColor: GREEN,
        },
        // ✅ Opportunities bright yellow
        {
          type: "bar",
          label: "Opportunities (Pipeline Weighted Hours)",
          data: barsPipeline,
          stack: "stack1",
          backgroundColor: YELLOW,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: { legend: { position: "top" }, tooltip: { mode: "index", intersect: false } },
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
      scales: { y: { beginAtZero: true } },
    },
  });
}

function buildRevenuePaceChart(canvas, targetMonthRev, actualMonthRev, projectedFullMonthRev) {
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: ["Revenue"],
      datasets: [
        { label: "Target (Full Month)", data: [targetMonthRev] },
        { label: "Actual (MTD from SalesAct)", data: [actualMonthRev] },
        { label: "Projected (Full Month)", data: [projectedFullMonthRev] },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      scales: { y: { beginAtZero: true } },
    },
  });
}

// ------------------------------------------------------------
// UI helpers
// ------------------------------------------------------------
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
      actualMonthRev: (a) => a?.constrRevMTD || 0,
      pipeFilter: (p) => isConstructionDivision(p.division),
    };
  }
  if (view === "maintenance") {
    return {
      targetYearRev: (t) => sumMonths(t.maintRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.maintRev[mk] || 0,
      actualMonthRev: (a) => a?.maintRevMTD || 0,
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
    actualMonthRev: (a) => (a?.constrRevMTD || 0) + (a?.maintRevMTD || 0),
    pipeFilter: (p) => isConstructionDivision(p.division) || isMaintenanceDivision(p.division),
  };
}

// ------------------------------------------------------------
// Render
// ------------------------------------------------------------
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

  // KPI: Revenue pace (SalesAct MTD)
  const targetMonthRev = scope.targetMonthRev(state.targets, mk);
  const salesAct = state.salesActByMonth?.[mk] || {};
  const actualMonthRev = scope.actualMonthRev(salesAct);
  const projectedRev = projectionForMonth(mk, actualMonthRev);

  chartRevenuePace = destroy(chartRevenuePace);
  chartRevenuePace = buildRevenuePaceChart(
    document.getElementById("chartRevenuePace"),
    targetMonthRev,
    actualMonthRev,
    projectedRev
  );

  // Hours by month: tickets + pipeline weighted hours (exclude won)
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

  // Maintenance pipeline spread start month -> Nov
  const maintPipeRows = pipePotential.filter(
    p => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division)
  );
  const pipeMaintMap = bucketMaintenancePipelineSpread(maintPipeRows, monthKeys);

  const constrTarget = monthKeys.map(m => state.targets.constrHours[m] || 0);
  const maintTarget  = monthKeys.map(m => state.targets.maintHours[m] || 0);

  const constrCap = monthKeys.map(m => state.capacity.constCap[m] || 0);
  const maintCap  = monthKeys.map(m => state.capacity.maintCap[m] || 0);

  const constrTickets = monthKeys.map(m => ticketConstr[m] || 0);
  const maintTickets  = monthKeys.map(m => ticketMaint[m] || 0);

  const constrPipe = monthKeys.map(m => pipeConstr[m] || 0);
  const maintPipe  = monthKeys.map(m => pipeMaintMap[m] || 0);

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

  const lr = document.getElementById("lastRefresh");
  if (lr) lr.textContent = new Date().toLocaleString();
}

// ------------------------------------------------------------
// Load all data + wiring
// ------------------------------------------------------------
async function loadAllData(state) {
  const tPath = document.getElementById("targetsPath");
  const pPath = document.getElementById("pipelinePath");
  const wPath = document.getElementById("ticketsPath");

  const tPill = document.getElementById("targetsStatus");
  const pPill = document.getElementById("pipelineStatus");
  const wPill = document.getElementById("ticketsStatus");

  if (tPath) tPath.textContent = FILES.targetsCsv;
  if (pPath) pPath.textContent = FILES.pipelineCsv;
  if (wPath) wPath.textContent = FILES.workTicketsXlsx;

  // Targets
  try {
    state.targets = await loadTargets(FILES.targetsCsv);
    if (tPill) {
      tPill.textContent = `Loaded (${state.targets.monthKeys.length} months)`;
      tPill.style.background = "rgba(34,197,94,0.15)";
    }
    populateMonthSelect(state.targets.monthKeys);
  } catch (e) {
    console.error("Targets failed:", e);
    if (tPill) {
      tPill.textContent = `Failed: ${e?.message || e}`;
      tPill.style.background = "rgba(239,68,68,0.15)";
    }
    throw e;
  }

  // Capacity (depends on months)
  try {
    state.capacity = await loadCapacity(FILES.capacityCsv, state.targets.monthKeys);
  } catch (e) {
    console.error("Capacity failed:", e);
    state.capacity = { constCap: {}, maintCap: {} };
  }

  // Pipeline
  try {
    state.pipeline = await loadPipeline(FILES.pipelineCsv);
    if (pPill) {
      pPill.textContent = `Loaded (${state.pipeline.length})`;
      pPill.style.background = "rgba(34,197,94,0.15)";
    }
  } catch (e) {
    console.error("Pipeline failed:", e);
    state.pipeline = [];
    if (pPill) {
      pPill.textContent = `Failed: ${e?.message || e}`;
      pPill.style.background = "rgba(239,68,68,0.15)";
    }
  }

  // Work Tickets
  try {
    state.tickets = await loadWorkTickets(FILES.workTicketsXlsx);
    if (wPill) {
      wPill.textContent = `Loaded (${state.tickets.length})`;
      wPill.style.background = "rgba(34,197,94,0.15)";
    }
  } catch (e) {
    console.error("Work tickets failed:", e);
    state.tickets = [];
    if (wPill) {
      wPill.textContent = `Failed: ${e?.message || e}`;
      wPill.style.background = "rgba(239,68,68,0.15)";
    }
  }

  // SalesAct (Actuals)
  try {
    state.salesActByMonth = await loadSalesActMonthly(LOGBOOK_URL, state.targets.monthKeys);
  } catch (e) {
    console.error("SalesAct load failed:", e);
    state.salesActByMonth = {};
    for (const mk of state.targets.monthKeys) {
      state.salesActByMonth[mk] = { constrRevMTD: 0, maintRevMTD: 0, constrHrsMTD: 0, maintHrsMTD: 0 };
    }
  }
}

function wireControls(state) {
  const refreshBtn = document.getElementById("refreshBtn");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      await loadAllData(state);
      renderAll(state);
    });
  }

  const viewToggle = document.getElementById("viewToggle");
  if (viewToggle) viewToggle.addEventListener("change", () => renderAll(state));

  const asOfMonth = document.getElementById("asOfMonth");
  if (asOfMonth) asOfMonth.addEventListener("change", () => renderAll(state));
}

// ------------------------------------------------------------
// Init (no top-level await)
// ------------------------------------------------------------
(async function init() {
  const state = {
    targets: null,
    capacity: { constCap: {}, maintCap: {} },
    pipeline: [],
    tickets: [],
    salesActByMonth: {},
  };

  wireControls(state);
  await loadAllData(state);
  renderAll(state);
})();
