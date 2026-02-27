// ----------------------------
// GitHub Pages-safe absolute URLs to /data/*
// ----------------------------
const BASE = new URL(".", window.location.href).href;
const DEFAULT_FILES = {
  targetsCsv: new URL("data/targets.csv", BASE).toString(),
  pipelineCsv: new URL("data/pipeline.csv", BASE).toString(),
  workTicketsXlsx: new URL("data/work_tickets.xlsx", BASE).toString(),
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

function sumMonths(seriesByMonth, monthKeys) {
  return monthKeys.reduce((acc, k) => acc + (seriesByMonth[k] || 0), 0);
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

// ✅ NEW: spread maintenance pipeline weighted hours from Start month -> November
function bucketMaintenancePipelineSpread(pipelineRows, monthKeys) {
  const out = {};
  for (const k of monthKeys) out[k] = 0;

  for (const p of pipelineRows) {
    if (!p.startDate) continue;

    const y = p.startDate.getFullYear();
    const startM = p.startDate.getMonth() + 1; // 1-12
    const endM = 11; // November

    const total = Number(p.weightedHours || 0);

    if (startM > endM) {
      // If start month is after November, just bucket to start month
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

// Projection for current month only
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
    delimiter: "", // auto-detect tab/comma
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

  const norm = (s) => String(s || "").trim().toLowerCase().replace(/\./g, "").replace(/\s+/g, " ");
  const pick = (...candidates) => {
    const cand = candidates.map(norm);
    const found = keys.find(k => cand.includes(norm(k)));
    return found || null;
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

// ----------------------------
// Actuals (localStorage)
// ----------------------------
const ACTUALS_KEY = "dashboard_actuals_v5";

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

  const el = document.getElementById("actualsLoaded");
  if (el) {
    el.textContent =
      `Month ${mk} • Construction: $${a.constrRev || 0} / ${a.constrHrs || 0} hrs • Maintenance: $${a.maintRev || 0} / ${a.maintHrs || 0} hrs`;
  }
}

// ----------------------------
// Charts
// ----------------------------
let chartConstr = null;
let chartMaint = null;
let chartRevenueYear = null;
let chartRevenuePace = null;
let chartCoverageConstr = null;
let chartCoverageMaint = null;

function destroy(chart) { if (chart) chart.destroy(); return null; }

function buildHoursChart(canvas, labels, targetLine, barsTickets, barsPipeline) {
  return new Chart(canvas.getContext("2d"), {
    data: {
      labels,
      datasets: [
        { type: "line", label: "Target Hours", data: targetLine, borderWidth: 2, pointRadius: 2, tension: 0.2 },
        { type: "bar", label: "Work Tickets (Open/Scheduled)", data: barsTickets, stack: "stack1" },
        { type: "bar", label: "Opportunities (Pipeline Weighted Hours)", data: barsPipeline, stack: "stack1" },
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

function buildRevenuePaceChart(canvas, targetMonthRev, actualMtdRev, projectedFullMonthRev) {
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: ["Revenue"],
      datasets: [
        { label: "Target (Full Month)", data: [targetMonthRev] },
        { label: "Actual (MTD)", data: [actualMtdRev] },
        { label: "Projected (Full Month)", data: [projectedFullMonthRev] },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, animation: false, scales: { y: { beginAtZero: true } } },
  });
}

// Color thresholds vs target:
// red: < 0.75, yellow: < 0.90, green: >= 0.90
function coverageColor(work, target) {
  if (!target || target <= 0) return "rgba(156,163,175,0.7)";
  const r = work / target;
  if (r < 0.75) return "rgba(239,68,68,0.75)";
  if (r < 0.90) return "rgba(234,179,8,0.75)";
  return "rgba(34,197,94,0.75)";
}

function buildCoverageChart(canvas, labels, targetLine, workBars) {
  const colors = labels.map((_, i) => coverageColor(workBars[i], targetLine[i]));
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
          // ✅ Blue so it's visible on dark background
          borderColor: "rgba(54, 162, 235, 1)",
          pointBackgroundColor: "rgba(54, 162, 235, 1)",
        },
        {
          type: "bar",
          label: "Work Tickets Hours (Open/Scheduled)",
          data: workBars,
          backgroundColor: colors,
        }
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: { legend: { position: "top" } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } } },
    },
  });
}

// ----------------------------
// UI + rendering
// ----------------------------
function populateMonthSelect(monthKeys) {
  const sel = document.getElementById("asOfMonth");
  const prior = sel.value;

  sel.innerHTML = "";
  for (const mk of monthKeys) {
    const opt = document.createElement("option");
    opt.value = mk;
    opt.textContent = mk;
    sel.appendChild(opt);
  }

  // ✅ default to current month if present
  const now = new Date();
  const currentMk = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

  if (monthKeys.includes(currentMk)) {
    sel.value = currentMk;
  } else if (monthKeys.includes(prior)) {
    sel.value = prior;
  } else {
    sel.value = monthKeys[monthKeys.length - 1];
  }
}

function getScope(view) {
  if (view === "construction") {
    return {
      targetYearRev: (t) => sumMonths(t.constrRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.constrRev[mk] || 0,
      actualMonthRev: (a) => a.constrRev || 0,
      pipeFilter: (p) => isConstructionDivision(p.division),
    };
  }
  if (view === "maintenance") {
    return {
      targetYearRev: (t) => sumMonths(t.maintRev, t.monthKeys),
      targetMonthRev: (t, mk) => t.maintRev[mk] || 0,
      actualMonthRev: (a) => a.maintRev || 0,
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
    actualMonthRev: (a) => (a.constrRev || 0) + (a.maintRev || 0),
    pipeFilter: (p) => isConstructionDivision(p.division) || isMaintenanceDivision(p.division),
  };
}

function renderAll(state) {
  if (!state.targets?.monthKeys?.length) return;

  const view = document.getElementById("viewToggle").value || "all";
  const mk = document.getElementById("asOfMonth").value || state.targets.monthKeys[state.targets.monthKeys.length - 1];
  const scope = getScope(view);
  const monthKeys = state.targets.monthKeys;

  // --- KPI charts ---
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

  const actualsStore = loadActualsStore();
  const a = getActualsForMonth(actualsStore, mk);
  const targetMonthRev = scope.targetMonthRev(state.targets, mk);
  const actualMtdRev = scope.actualMonthRev(a);
  const projectedRev = projectionForMonth(mk, actualMtdRev);

  chartRevenuePace = destroy(chartRevenuePace);
  chartRevenuePace = buildRevenuePaceChart(
    document.getElementById("chartRevenuePace"),
    targetMonthRev,
    actualMtdRev,
    projectedRev
  );

  // --- Hours by month ---
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

  // ✅ Maintenance pipeline spread (Start month -> Nov)
  const maintPipeRows = pipePotential.filter(
    p => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division)
  );
  const pipeMaint = bucketMaintenancePipelineSpread(maintPipeRows, monthKeys);

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

  // --- Next 3 months coverage (targets vs work tickets only) ---
  const startIdx = Math.max(0, monthKeys.indexOf(mk));
  const next3 = monthKeys.slice(startIdx, startIdx + 3);

  const next3ConstrTarget = next3.map(m => state.targets.constrHours[m] || 0);
  const next3MaintTarget  = next3.map(m => state.targets.maintHours[m] || 0);

  const next3ConstrWork = next3.map(m => ticketConstr[m] || 0);
  const next3MaintWork  = next3.map(m => ticketMaint[m] || 0);

  chartCoverageConstr = destroy(chartCoverageConstr);
  chartCoverageMaint = destroy(chartCoverageMaint);

  chartCoverageConstr = buildCoverageChart(
    document.getElementById("chartCoverageConstr"),
    next3,
    next3ConstrTarget,
    next3ConstrWork
  );

  chartCoverageMaint = buildCoverageChart(
    document.getElementById("chartCoverageMaint"),
    next3,
    next3MaintTarget,
    next3MaintWork
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

  const mk = document.getElementById("asOfMonth").value || state.targets.monthKeys[state.targets.monthKeys.length - 1];
  loadMonthActualsIntoInputs(loadActualsStore(), mk);
}

function wireControls(state) {
  document.getElementById("refreshBtn").addEventListener("click", async () => {
    await loadAllData(state);
    renderAll(state);
  });

  document.getElementById("viewToggle").addEventListener("change", () => {
    renderAll(state);
  });

  document.getElementById("asOfMonth").addEventListener("change", () => {
    const mk = document.getElementById("asOfMonth").value;
    loadMonthActualsIntoInputs(loadActualsStore(), mk);
    renderAll(state);
  });
}

function wireActualsButtons(state) {
  document.getElementById("saveActualsBtn").addEventListener("click", () => {
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
  });

  document.getElementById("exportActualsBtn").addEventListener("click", () => {
    const store = loadActualsStore();
    const blob = new Blob([JSON.stringify(store, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "actuals.json";
    a.click();
    URL.revokeObjectURL(url);
  });

  document.getElementById("importActualsInput").addEventListener("change", async (ev) => {
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
  });

  document.getElementById("clearActualsBtn").addEventListener("click", () => {
    if (!confirm("Clear ALL actuals stored in this browser?")) return;
    localStorage.removeItem(ACTUALS_KEY);
    loadMonthActualsIntoInputs(loadActualsStore(), document.getElementById("asOfMonth").value);
    renderAll(state);
  });
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

