// ------------------------------------------------------------
// CONFIG
// ------------------------------------------------------------
const LOGBOOK_URL =
  "https://script.google.com/macros/s/AKfycbxemOcHaO8jJL2JNvr6G3INrHOSahH3-1QYcsrb5IV19DG77lPUPtDkco_s9r8RFwmI/exec";

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
const LOST_STATUS_WORDS = ["lost", "closed lost", "no sale", "cancel"];

// Colors
const BLUE = "rgba(54, 162, 235, 1)";
const GREEN_SOLID = "rgba(34, 197, 94, 0.85)";
const GREEN_SHADE = "rgba(34, 197, 94, 0.25)";
const PURPLE_SOLID = "rgba(168, 85, 247, 0.80)";
const PURPLE_SHADE = "rgba(168, 85, 247, 0.25)";
const TICKETS_GREEN = "rgba(34, 197, 94, 0.90)";
const OPPS_YELLOW = "rgba(250, 204, 21, 0.90)";

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
function isLostPipelineStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  return LOST_STATUS_WORDS.some(w => s.includes(w));
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

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return Number.isFinite(d.getTime()) ? d : null;
  }

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
function normalizeMonthKey(mk) {
  const s = String(mk || "").trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const d = parseDateAny(s);
  return d ? monthKeyFromDate(d) : "";
}
function monthKeyFromTargetsLabel(label) {
  const s = String(label || "").trim();

  let m = s.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (m) {
    const mon = m[1].toLowerCase();
    const yy = Number(m[2]) + 2000;
    const idx = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"].indexOf(mon);
    if (idx >= 0) return `${yy}-${String(idx + 1).padStart(2, "0")}`;
  }

  m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) {
    const yy = Number(m[1]);
    const mm = Number(m[2]);
    if (yy >= 2000 && mm >= 1 && mm <= 12) return `${yy}-${String(mm).padStart(2, "0")}`;
  }

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
    const endM = 11; // Nov
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

// Coverage KPI colors
function coverageStatus(ratio) {
  if (!Number.isFinite(ratio)) return { label: "—", bg: "rgba(148,163,184,0.20)" };
  if (ratio < 0.75) return { label: `${Math.round(ratio * 100)}%`, bg: "rgba(239,68,68,0.30)" };
  if (ratio < 0.90) return { label: `${Math.round(ratio * 100)}%`, bg: "rgba(250,204,21,0.35)" };
  return { label: `${Math.round(ratio * 100)}%`, bg: "rgba(34,197,94,0.30)" };
}
function renderCoverageKpi(containerId, months, targets, works) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = "";
  for (let i = 0; i < months.length; i++) {
    const t = targets[i] || 0;
    const w = works[i] || 0;
    const ratio = t > 0 ? (w / t) : NaN;
    const s = coverageStatus(ratio);

    const pill = document.createElement("div");
    pill.className = "kpi-pill";
    pill.style.background = s.bg;
    pill.textContent = `${months[i]}: ${s.label}`;
    el.appendChild(pill);
  }
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

  let count = 0;
  for (const r of parsed.data || []) {
    const mk = normalizeMonthKey(r.month ?? r.Month);
    if (!mk || !(mk in constCap)) continue;
    constCap[mk] = parseNumberLoose(r.constcap ?? r.ConstCap);
    maintCap[mk] = parseNumberLoose(r.maintcap ?? r.MaintCap);
    count++;
  }

  return { constCap, maintCap, _count: count };
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
  const out = {};
  for (const mk of monthKeys) {
    out[mk] = { constrRevMTD: 0, maintRevMTD: 0, constrHrsMTD: 0, maintHrsMTD: 0, updatedAt: "", updatedBy: "" };
  }

  const url = `${baseUrl}?tab=SalesAct&_=${Date.now()}`;
  const res = await fetch(url, { cache: "no-store" });
  const raw = await res.text();
  const t = raw.trim();

  if (t.startsWith("<!DOCTYPE") || t.startsWith("<html") || t.includes("<body")) {
    throw new Error("SalesAct returned HTML, not JSON. Re-deploy Apps Script Web App with access: Anyone / Anyone with link.");
  }

  let payload;
  try {
    payload = JSON.parse(t);
  } catch {
    throw new Error("SalesAct response is not valid JSON.");
  }

  if (payload?.ok === false) {
    throw new Error(`SalesAct ok:false (${payload.error || "unknown error"})`);
  }

  const rows = payload?.rows || [];
  for (const r of rows) {
    const mk = normalizeMonthKey(r.Month);
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
let chartRevenueYear = null;
let chartRevenuePace = null;
let chartConstr = null;
let chartMaint = null;
let chartCoverageConstr = null;
let chartCoverageMaint = null;

function destroy(chart) { if (chart) chart.destroy(); return null; }

// Revenue Year: Target + Pipeline stacks (Won solid, Remaining shaded)
function buildRevenueYearChart(canvas, target, unwWon, unwRem, wWon, wRem) {
  return new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: ["Year"],
      datasets: [
        { label: "Target Revenue", data: [target], backgroundColor: "rgba(54,162,235,0.45)" },

        { label: "Pipeline $ (Unweighted) - Won", data: [unwWon], stack: "unw", backgroundColor: GREEN_SOLID },
        { label: "Pipeline $ (Unweighted) - Remaining", data: [unwRem], stack: "unw", backgroundColor: GREEN_SHADE },

        { label: "Pipeline $ (Weighted) - Won", data: [wWon], stack: "w", backgroundColor: PURPLE_SOLID },
        { label: "Pipeline $ (Weighted) - Remaining", data: [wRem], stack: "w", backgroundColor: PURPLE_SHADE },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: { legend: { position: "top" } },
      scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } },
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

// Hours chart: bars stacked; lines separate axis (prevents capacity adding to target)
function buildHoursChart(canvas, labels, targetLine, capLine, barsTickets, barsPipeline) {
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
          borderColor: BLUE,
          pointBackgroundColor: BLUE,
          yAxisID: "y2",
        },
        {
          type: "line",
          label: "Capacity",
          data: capLine,
          borderWidth: 2,
          pointRadius: 0,
          tension: 0.2,
          borderColor: BLUE,
          borderDash: [6, 6],
          yAxisID: "y2",
        },
        {
          type: "bar",
          label: "Work Tickets (Open/Scheduled)",
          data: barsTickets,
          stack: "hours",
          backgroundColor: TICKETS_GREEN,
          yAxisID: "y",
        },
        {
          type: "bar",
          label: "Opportunities (Pipeline Weighted Hours)",
          data: barsPipeline,
          stack: "hours",
          backgroundColor: OPPS_YELLOW,
          yAxisID: "y",
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
        y2: { stacked: false, beginAtZero: true, ticks: { precision: 0 }, grid: { drawOnChartArea: false } },
      },
    },
  });
}

function buildCoverageChart(canvas, labels, targetLine, workBars) {
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
          borderColor: BLUE,
          pointBackgroundColor: BLUE,
        },
        {
          type: "bar",
          label: "Work Tickets Hours (Open/Scheduled)",
          data: workBars,
          backgroundColor: TICKETS_GREEN,
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

// ------------------------------------------------------------
// UI helpers
// ------------------------------------------------------------
function populateMonthSelect(monthKeys) {
  const sel = document.getElementById("asOfMonth");
  if (!sel) return;
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

  const monthKeys = state.targets.monthKeys;
  const view = document.getElementById("viewToggle")?.value || "all";

  let mk = document.getElementById("asOfMonth")?.value || monthKeys[monthKeys.length - 1];
  if (!monthKeys.includes(mk)) mk = monthKeys[monthKeys.length - 1];

  const scope = getScope(view);

  // --- Revenue year pipeline: won + remaining (exclude lost from remaining) ---
  const inScopePipe = state.pipeline.filter(scope.pipeFilter);

  const wonPipe = inScopePipe.filter(p => isWonPipelineStatus(p.status));
  const remPipe = inScopePipe.filter(p => !isWonPipelineStatus(p.status) && !isLostPipelineStatus(p.status));

  const unwWon = wonPipe.reduce((a, p) => a + (p.estimatedDollars || 0), 0);
  const unwRem = remPipe.reduce((a, p) => a + (p.estimatedDollars || 0), 0);

  const wWon = wonPipe.reduce((a, p) => a + (p.weightedPipeline || 0), 0);
  const wRem = remPipe.reduce((a, p) => a + (p.weightedPipeline || 0), 0);

  const targetYearRev = scope.targetYearRev(state.targets);

  chartRevenueYear = destroy(chartRevenueYear);
  chartRevenueYear = buildRevenueYearChart(
    document.getElementById("chartRevenueYear"),
    targetYearRev,
    unwWon, unwRem,
    wWon, wRem
  );

  // --- Revenue pace (SalesAct actuals) ---
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

  // SalesAct stamp
  const stamp = document.getElementById("salesActStamp");
  if (stamp) {
    stamp.textContent = salesAct?.updatedAt
      ? `SalesAct updated ${salesAct.updatedAt}${salesAct.updatedBy ? ` by ${salesAct.updatedBy}` : ""}`
      : "";
  }

  // --- Hours by month (tickets + pipeline weighted hours, exclude WON) ---
  const activeTickets = state.tickets.filter(t => isTicketActive(t.status));
  const pipePotential = state.pipeline.filter(p => !isWonPipelineStatus(p.status));

  const ticketConstrMap = bucketSumByMonth(
    activeTickets.filter(t => isConstructionDivision(t.division)),
    t => t.schedDate,
    t => t.estHrs,
    monthKeys
  );

  const ticketMaintMap = bucketSumByMonth(
    activeTickets.filter(t => isMaintenanceDivision(t.division) && !isConstructionDivision(t.division)),
    t => t.schedDate,
    t => t.estHrs,
    monthKeys
  );

  const pipeConstrMap = bucketSumByMonth(
    pipePotential.filter(p => isConstructionDivision(p.division)),
    p => p.startDate,
    p => p.weightedHours,
    monthKeys
  );

  const maintPipeRows = pipePotential.filter(
    p => isMaintenanceDivision(p.division) && !isConstructionDivision(p.division)
  );
  const pipeMaintMap = bucketMaintenancePipelineSpread(maintPipeRows, monthKeys);

  const constrTarget = monthKeys.map(m => state.targets.constrHours[m] || 0);
  const maintTarget  = monthKeys.map(m => state.targets.maintHours[m] || 0);

  const constrCap = monthKeys.map(m => state.capacity?.constCap?.[m] || 0);
  const maintCap  = monthKeys.map(m => state.capacity?.maintCap?.[m] || 0);

  const constrTickets = monthKeys.map(m => ticketConstrMap[m] || 0);
  const maintTickets  = monthKeys.map(m => ticketMaintMap[m] || 0);

  const constrPipe = monthKeys.map(m => pipeConstrMap[m] || 0);
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

  // --- Next 3 months coverage (based on work tickets vs target) ---
  let startIdx = monthKeys.indexOf(mk);
  if (startIdx < 0) startIdx = Math.max(0, monthKeys.length - 3);
  const next3 = monthKeys.slice(startIdx, startIdx + 3);

  const next3ConstrTarget = next3.map(m => state.targets.constrHours[m] || 0);
  const next3MaintTarget  = next3.map(m => state.targets.maintHours[m] || 0);

  const next3ConstrWork = next3.map(m => ticketConstrMap[m] || 0);
  const next3MaintWork  = next3.map(m => ticketMaintMap[m] || 0);

  renderCoverageKpi("kpiCoverageConstr", next3, next3ConstrTarget, next3ConstrWork);
  renderCoverageKpi("kpiCoverageMaint",  next3, next3MaintTarget,  next3MaintWork);

  chartCoverageConstr = destroy(chartCoverageConstr);
  chartCoverageMaint = destroy(chartCoverageMaint);

  const cc = document.getElementById("chartCoverageConstr");
  const cm = document.getElementById("chartCoverageMaint");
  if (cc) chartCoverageConstr = buildCoverageChart(cc, next3, next3ConstrTarget, next3ConstrWork);
  if (cm) chartCoverageMaint  = buildCoverageChart(cm, next3, next3MaintTarget,  next3MaintWork);

  const lr = document.getElementById("lastRefresh");
  if (lr) lr.textContent = new Date().toLocaleString();
}

// ------------------------------------------------------------
// Load all data + wiring
// ------------------------------------------------------------
async function loadAllData(state) {
  const setPill = (id, ok, msg) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = msg;
    el.style.background = ok ? "rgba(34,197,94,0.15)" : "rgba(239,68,68,0.15)";
  };

  // Paths
  const setPath = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };
  setPath("targetsPath", FILES.targetsCsv);
  setPath("pipelinePath", FILES.pipelineCsv);
  setPath("ticketsPath", FILES.workTicketsXlsx);
  setPath("capacityPath", FILES.capacityCsv);

  // Targets
  try {
    state.targets = await loadTargets(FILES.targetsCsv);
    populateMonthSelect(state.targets.monthKeys);
    setPill("targetsStatus", true, `Targets: Loaded (${state.targets.monthKeys.length} months)`);
  } catch (e) {
    console.error("Targets failed:", e);
    setPill("targetsStatus", false, `Targets: Failed (${e?.message || e})`);
    throw e;
  }

  // Pipeline
  try {
    state.pipeline = await loadPipeline(FILES.pipelineCsv);
    setPill("pipelineStatus", true, `Pipeline: Loaded (${state.pipeline.length})`);
  } catch (e) {
    console.error("Pipeline failed:", e);
    state.pipeline = [];
    setPill("pipelineStatus", false, `Pipeline: Failed (${e?.message || e})`);
  }

  // Work tickets
  try {
    state.tickets = await loadWorkTickets(FILES.workTicketsXlsx);
    setPill("ticketsStatus", true, `Work tickets: Loaded (${state.tickets.length})`);
  } catch (e) {
    console.error("Work tickets failed:", e);
    state.tickets = [];
    setPill("ticketsStatus", false, `Work tickets: Failed (${e?.message || e})`);
  }

  // Capacity
  try {
    state.capacity = await loadCapacity(FILES.capacityCsv, state.targets.monthKeys);
    setPill("capacityStatus", true, `Capacity: Loaded (${state.capacity._count || 0} rows)`);
  } catch (e) {
    console.error("Capacity failed:", e);
    state.capacity = { constCap: {}, maintCap: {}, _count: 0 };
    setPill("capacityStatus", false, `Capacity: Failed (${e?.message || e})`);
  }

  // SalesAct actuals
  try {
    state.salesActByMonth = await loadSalesActMonthly(LOGBOOK_URL, state.targets.monthKeys);
    setPill("salesActStatus", true, "SalesAct: Loaded");
  } catch (e) {
    console.error("SalesAct load failed:", e);
    state.salesActByMonth = {};
    for (const mk of state.targets.monthKeys) {
      state.salesActByMonth[mk] = { constrRevMTD: 0, maintRevMTD: 0, constrHrsMTD: 0, maintHrsMTD: 0 };
    }
    setPill("salesActStatus", false, `SalesAct: Failed (${e?.message || e})`);
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
// Init
// ------------------------------------------------------------
(async function init() {
  const state = {
    targets: null,
    pipeline: [],
    tickets: [],
    capacity: { constCap: {}, maintCap: {}, _count: 0 },
    salesActByMonth: {},
  };

  wireControls(state);
  await loadAllData(state);
  renderAll(state);
})();

