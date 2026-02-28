/* Ops Dashboard (GitHub Pages)
   Shared daily entries pulled from Google Sheet via Apps Script (read-only)
   Targets/calendar loaded from repo /data
*/

const CONFIG = {
  // Paste your Google Apps Script Web App URL here (ends with /exec)
  SHEET_API_URL: "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE",

  // If you add an API key in Apps Script, put it here (optional)
  API_KEY: "",

  // Repo data files (Option 2 shared /data)
  CALENDAR_URL: "../data/calendar.csv",
  TARGETS_URL: "../data/targets.csv",
  SHEET_TAB_NAME: "OpsDailyLog"
};

/* ---------------- helpers ---------------- */
const el = (id) => document.getElementById(id);
const safeNum = (v) => {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
};
const fmt1 = (n) => (Number.isFinite(n) ? n.toFixed(1) : "—");
const fmt0 = (n) => (Number.isFinite(n) ? String(Math.round(n)) : "—");

function toast(msg){
  const t = el("toast");
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"), 2600);
}

function ymd(d){
  const pad = (x)=> String(x).padStart(2,"0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}
function monthKeyFromDate(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}
function monthStartEnd(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  const start = ymd(new Date(d.getFullYear(), d.getMonth(), 1));
  const end = ymd(new Date(d.getFullYear(), d.getMonth()+1, 0));
  return { start, end };
}
function yearStart(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  return `${d.getFullYear()}-01-01`;
}
function isWorkday(d){ const day=d.getDay(); return day!==0 && day!==6; }
function workdaysInMonth(year, monthIndex){
  let c=0; const d=new Date(year, monthIndex, 1);
  while(d.getMonth()===monthIndex){ if(isWorkday(d)) c++; d.setDate(d.getDate()+1); }
  return c;
}
function countWorkdaysElapsed(year, monthIndex, through){
  let c=0; const d=new Date(year, monthIndex, 1);
  const end=new Date(through.getFullYear(), through.getMonth(), through.getDate());
  while(d.getMonth()===monthIndex && d<=end){ if(isWorkday(d)) c++; d.setDate(d.getDate()+1); }
  return c;
}

function setBoxStatus(node, state){
  node.classList.remove("good","warn","bad");
  if(state) node.classList.add(state);
}
function statusByPct(pct){
  if(pct==null) return null; // neutral for 0 target
  if(pct<0.95) return "bad";
  if(pct<1.0) return "warn";
  return "good";
}
function tagByPct(pct){
  if(pct==null) return { cls:"", txt:"No target" };
  if(pct<0.95) return { cls:"bad", txt:"Behind" };
  if(pct<1.0) return { cls:"warn", txt:"Close" };
  return { cls:"good", txt:"On Track" };
}

/* ---------------- CSV parsing (basic) ---------------- */
function parseDelimited(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim().length);
  if(!lines.length) return { headers:[], rows:[] };
  const delim = lines[0].includes("\t") ? "\t" : ",";
  const splitLine = (line)=>{
    const out=[]; let cur="", inQ=false;
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(ch === '"'){
        if(inQ && line[i+1] === '"'){ cur+='"'; i++; }
        else inQ=!inQ;
      } else if(ch===delim && !inQ){
        out.push(cur); cur="";
      } else cur += ch;
    }
    out.push(cur);
    return out.map(x=>x.trim());
  };
  const headers = splitLine(lines[0]).map(h=>h.replace(/\uFEFF/g,""));
  const rows = [];
  for(let i=1;i<lines.length;i++) rows.push(splitLine(lines[i]));
  return { headers, rows };
}

/* ---------------- targets fallback ---------------- */
let targets = {};        // targets[YYYY-MM] = { maintenanceMonthly, constructionMonthly }
let targetsLoaded = false;

function toMonthKey(raw){
  const s=(raw??"").toString().trim();
  if(!s) return null;

  // supports "26-Jan" or "Jan-26" or date strings
  const monthMap={jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};

  let m=s.match(/^(\d{2})[-\s]?([A-Za-z]{3,})\.?$/); // 26-Jan
  if(m){
    const yy=String(m[1]).padStart(2,"0");
    const mon=m[2].slice(0,3).toLowerCase();
    const mm=monthMap[mon];
    if(mm) return `20${yy}-${mm}`;
  }

  m=s.match(/^([A-Za-z]{3,})[-\s]?(\d{2})$/); // Jan-26
  if(m){
    const mon=m[1].slice(0,3).toLowerCase();
    const yy=String(m[2]).padStart(2,"0");
    const mm=monthMap[mon];
    if(mm) return `20${yy}-${mm}`;
  }

  const m2=s.match(/^(\d{4})-(\d{2})$/);
  if(m2) return `${m2[1]}-${m2[2]}`;

  const d=new Date(s);
  if(!isNaN(d)) return monthKeyFromDate(d);

  return null;
}

function parseTargets(text){
  const { headers, rows } = parseDelimited(text);
  const lower = headers.map(h=>h.toLowerCase());

  const monthIdx = lower.findIndex(h => h==="month" || h.includes("month"));
  const constIdx = lower.findIndex(h => h.includes("construction") && h.includes("hours"));
  const maintIdx = lower.findIndex(h => h.includes("maintenance") && h.includes("hours"));
  if(monthIdx < 0 || (constIdx < 0 && maintIdx < 0)) return {};

  const parseHours=(v)=>{
    const s=(v??"").toString().replaceAll(",","").trim();
    if(!s || s==="-" ) return 0;
    const n=Number(s);
    return Number.isFinite(n) ? n : 0;
  };

  const out = {};
  for(const row of rows){
    const mk = toMonthKey(row[monthIdx]);
    if(!mk) continue;
    out[mk] = {
      maintenanceMonthly: maintIdx>=0 ? parseHours(row[maintIdx]) : 0,
      constructionMonthly: constIdx>=0 ? parseHours(row[constIdx]) : 0
    };
  }
  return out;
}

function computeFallbackDailyTargets(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  const mk=monthKeyFromDate(d);
  const t=targets[mk] || {maintenanceMonthly:0, constructionMonthly:0};
  const wd=workdaysInMonth(d.getFullYear(), d.getMonth());
  const dailyMaint = wd>0 ? safeNum(t.maintenanceMonthly)/wd : 0;
  const dailyConst = wd>0 ? safeNum(t.constructionMonthly)/wd : 0;
  return {
    mk, wd,
    monthMaint: safeNum(t.maintenanceMonthly),
    monthConst: safeNum(t.constructionMonthly),
    dailyMaint, dailyConst, dailyTotal: dailyMaint+dailyConst
  };
}

/* ---------------- calendar targets ---------------- */
let calendar = {};        // calendar[YYYY-MM-DD] = { maintTarget, constTarget }
let calendarLoaded = false;

function parseCalendar(text){
  const { headers, rows } = parseDelimited(text);
  const lower = headers.map(h => h.toLowerCase());

  const dateIdx = lower.findIndex(h => h === "date" || h.includes("date"));
  const maintIdx = lower.findIndex(h => h.includes("maintenancehours") || (h.includes("maintenance") && h.includes("hours")));
  const constIdx = lower.findIndex(h => h.includes("constructionhours") || (h.includes("construction") && h.includes("hours")));
  if (dateIdx < 0 || (maintIdx < 0 && constIdx < 0)) return {};

  const toDateKey = (v) => {
    const s = (v ?? "").toString().trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    const d = new Date(s);
    if (!isNaN(d)) return ymd(d);
    return null;
  };

  const toNum = (v) => {
    const s = (v ?? "").toString().replaceAll(",", "").trim();
    if (!s || s === "-") return 0;
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  };

  const out = {};
  for (const row of rows){
    const key = toDateKey(row[dateIdx]);
    if (!key) continue;
    out[key] = {
      maintTarget: maintIdx >= 0 ? toNum(row[maintIdx]) : 0,
      constTarget: constIdx >= 0 ? toNum(row[constIdx]) : 0
    };
  }
  return out;
}

function getCalendarTargetForDate(dateKey){
  const r = calendar[dateKey];
  if (!r) return { maint: 0, cons: 0, total: 0, has: false };
  const maint = safeNum(r.maintTarget);
  const cons  = safeNum(r.constTarget);
  return { maint, cons, total: maint + cons, has: true };
}

function sumCalendarTargetsByDivision(startYmd, endYmd){
  let maint = 0, cons = 0;
  for (const [k, r] of Object.entries(calendar)){
    if (k < startYmd || k > endYmd) continue;
    maint += safeNum(r.maintTarget);
    cons  += safeNum(r.constTarget);
  }
  return { maint, cons, total: maint + cons };
}

function getDayTargets(dateKey){
  const fb = computeFallbackDailyTargets(dateKey);
  if(calendarLoaded){
    const c = getCalendarTargetForDate(dateKey);
    if(c.has) return { maint:c.maint, cons:c.cons, total:c.total, source:"Calendar", fallback:fb };
    return { maint:fb.dailyMaint, cons:fb.dailyConst, total:fb.dailyTotal, source:"Calendar (missing date → avg)", fallback:fb };
  }
  return { maint:fb.dailyMaint, cons:fb.dailyConst, total:fb.dailyTotal, source:"Monthly average", fallback:fb };
}
function getMonthTargetsSum(monthStart, monthEnd){
  if(calendarLoaded) return sumCalendarTargetsByDivision(monthStart, monthEnd);
  const fb = computeFallbackDailyTargets(monthStart);
  const d = new Date(monthStart+"T00:00:00");
  const wd = workdaysInMonth(d.getFullYear(), d.getMonth());
  return { maint: fb.dailyMaint*wd, cons: fb.dailyConst*wd, total: fb.dailyTotal*wd };
}
function getMTDTargetsSum(monthStart, through){
  if(calendarLoaded) return sumCalendarTargetsByDivision(monthStart, through);
  const fb = computeFallbackDailyTargets(through);
  const d = new Date(through+"T00:00:00");
  const elapsed = countWorkdaysElapsed(d.getFullYear(), d.getMonth(), d);
  return { maint: fb.dailyMaint*elapsed, cons: fb.dailyConst*elapsed, total: fb.dailyTotal*elapsed };
}
function getYTDTargetsSum(yearStartYmd, through){
  if(calendarLoaded) return sumCalendarTargetsByDivision(yearStartYmd, through);
  // fallback: sum whole-year monthly targets
  const d=new Date(through+"T00:00:00");
  const year=d.getFullYear();
  let tM=0, tC=0;
  for(let mi=0; mi<12; mi++){
    const mk = `${year}-${String(mi+1).padStart(2,"0")}`;
    const t = targets[mk];
    if(!t) continue;
    tM += safeNum(t.maintenanceMonthly);
    tC += safeNum(t.constructionMonthly);
  }
  return { maint:tM, cons:tC, total:tM+tC };
}

/* ---------------- shared daily entries from Google ---------------- */
let daily = {}; // daily[YYYY-MM-DD] = { actualMaint, actualConst, ticketsMissed, safetyIncidents, notes }
let sheetLoadedCount = 0;

function getDayRec(dateKey){
  return daily[dateKey] || { actualMaint:"", actualConst:"", ticketsMissed:"", safetyIncidents:"", notes:"" };
}

function sumActualsByDivision(startYmd, endYmd){
  let maint=0, cons=0;
  for(const [k, rec] of Object.entries(daily)){
    if(k < startYmd || k > endYmd) continue;
    maint += safeNum(rec.actualMaint);
    cons  += safeNum(rec.actualConst);
  }
  return { maint, cons, total: maint+cons };
}

async function loadDailyFromGoogle(startYmd, endYmd){
  if(!CONFIG.SHEET_API_URL || CONFIG.SHEET_API_URL.includes("PASTE_")){
    throw new Error("Missing Apps Script URL in CONFIG.SHEET_API_URL");
  }

  const u = new URL(CONFIG.SHEET_API_URL);
  u.searchParams.set("tab", CONFIG.SHEET_TAB_NAME);
  if(startYmd) u.searchParams.set("start", startYmd);
  if(endYmd) u.searchParams.set("end", endYmd);
  if(CONFIG.API_KEY) u.searchParams.set("key", CONFIG.API_KEY);

  const res = await fetch(u.toString(), { cache: "no-store" });
  if(!res.ok) throw new Error(`Google endpoint HTTP ${res.status}`);
  const data = await res.json();

  // expected: { ok:true, rows:[{Date:"YYYY-MM-DD", ActualMaint:..., ...}] }
  const map = {};
  const rows = Array.isArray(data.rows) ? data.rows : [];
  for(const r of rows){
    const date = String(r.Date || r.date || "").slice(0,10);
    if(!date.match(/^\d{4}-\d{2}-\d{2}$/)) continue;

    map[date] = {
      actualMaint: r.ActualMaint ?? r.ActualMaintHours ?? r.ActualMaintenance ?? r.ActualMaintHours ?? "",
      actualConst: r.ActualConst ?? r.ActualConstHours ?? r.ActualConstruction ?? r.ActualConstHours ?? "",
      ticketsMissed: r.MissedTickets ?? r.TicketsMissed ?? r.Missed ?? "",
      safetyIncidents: r.SafetyIncidents ?? r.Incidents ?? r.Safety ?? "",
      notes: r.Notes ?? r.Note ?? ""
    };
  }

  daily = map;
  sheetLoadedCount = rows.length;

  el("rowsLoaded").textContent = String(sheetLoadedCount);
  el("lastFetch").textContent = new Date().toLocaleTimeString();
  el("sheetTag").className = "tag good";
  el("sheetTag").textContent = "Connected";
}

/* ---------------- Month table render ---------------- */
function renderMonthTable(mStart, mEnd, selected){
  const tbody = el("monthTbody");
  tbody.innerHTML = "";

  const start = new Date(mStart+"T00:00:00");
  const end = new Date(mEnd+"T00:00:00");

  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){
    const key = ymd(d);
    const rec = getDayRec(key);

    const t = getDayTargets(key);
    const aM = safeNum(rec.actualMaint);
    const aC = safeNum(rec.actualConst);
    const aT = aM + aC;

    const pct = t.total>0 ? (aT/t.total) : null;
    const tagState = tagByPct(pct);
    const deltaPct = (t.total>0) ? (((aT/t.total)-1)*100).toFixed(1)+"%" : "—";

    const tickets = safeNum(rec.ticketsMissed);
    const safety = safeNum(rec.safetyIncidents);
    const notes = (rec.notes ?? "").toString();

    const tr = document.createElement("tr");
    tr.className = "clickRow";
    if(key === selected) tr.style.outline = "1px solid rgba(55,233,255,.35)";

    tr.innerHTML = `
      <td class="mono">${key}</td>
      <td><span class="tag ${tagState.cls}">${tagState.txt}</span></td>
      <td class="mono">${fmt1(t.maint)}</td>
      <td class="mono">${fmt1(t.cons)}</td>
      <td class="mono">${fmt1(t.total)}</td>
      <td class="mono">${fmt1(aM)}</td>
      <td class="mono">${fmt1(aC)}</td>
      <td class="mono">${fmt1(aT)}</td>
      <td class="mono">${deltaPct}</td>
      <td class="mono">${fmt0(tickets)}</td>
      <td class="mono">${fmt0(safety)}</td>
      <td>${escapeHtml(notes)}</td>
    `;

    tr.addEventListener("click", ()=>{
      setDate(key);
      window.scrollTo({ top: 0, behavior: "smooth" });
    });

    tbody.appendChild(tr);
  }
}

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

/* ---------------- Export month CSV ---------------- */
function exportMonthCSV(){
  const sel = el("datePicker").value;
  const { start: mStart, end: mEnd } = monthStartEnd(sel);

  const rows = [];
  rows.push([
    "Date","TargetMaint","TargetConst","TargetTotal",
    "ActualMaint","ActualConst","ActualTotal",
    "DeltaPct","MissedTickets","SafetyIncidents","Notes"
  ]);

  const start = new Date(mStart+"T00:00:00");
  const end = new Date(mEnd+"T00:00:00");

  for(let d=new Date(start); d<=end; d.setDate(d.getDate()+1)){
    const key = ymd(d);
    const rec = getDayRec(key);
    const t = getDayTargets(key);

    const aM = safeNum(rec.actualMaint);
    const aC = safeNum(rec.actualConst);
    const aT = aM + aC;
    const deltaPct = t.total>0 ? (((aT/t.total)-1)*100) : "";

    rows.push([
      key, t.maint, t.cons, t.total,
      aM, aC, aT,
      (t.total>0 ? deltaPct.toFixed(1)+"%" : ""),
      safeNum(rec.ticketsMissed),
      safeNum(rec.safetyIncidents),
      (rec.notes ?? "")
    ]);
  }

  const esc = (v)=> {
    const s = String(v ?? "");
    return (s.includes(",") || s.includes('"') || s.includes("\n")) ? `"${s.replaceAll('"','""')}"` : s;
  };
  const csv = rows.map(r=>r.map(esc).join(",")).join("\n");
  const blob = new Blob([csv], { type:"text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `ops-month-export-${mStart.slice(0,7)}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---------------- Render scoreboard + summaries ---------------- */
function render(){
  const sel = el("datePicker").value;
  if(!sel) return;

  const selDate = new Date(sel+"T00:00:00");
  const mk = monthKeyFromDate(selDate);
  el("displayDate").textContent = sel;
  el("monthLabel").textContent = mk;

  // target debug labels
  el("targetsLabel").textContent = targetsLoaded ? "Yes" : "No";
  el("calendarLabel").textContent = calendarLoaded ? "Yes" : "No";
  el("calendarDebug").textContent = CONFIG.CALENDAR_URL;
  el("targetsDebug").textContent = CONFIG.TARGETS_URL;

  // day targets + day actuals (from Google)
  const t = getDayTargets(sel);
  const rec = getDayRec(sel);

  const aM = safeNum(rec.actualMaint);
  const aC = safeNum(rec.actualConst);
  const aT = aM + aC;

  el("dailyTargetsLine").textContent = `${fmt1(t.maint)} / ${fmt1(t.cons)} / ${fmt1(t.total)}`;
  el("dailyActualsLine").textContent = `${fmt1(aM)} / ${fmt1(aC)} / ${fmt1(aT)}`;
  el("targetSource").textContent = t.source;

  let delta = "—";
  const pct = t.total>0 ? (aT/t.total) : null;
  if(t.total>0){
    delta = `${(((aT/t.total)-1)*100).toFixed(1)}%`;
  }
  el("deltaPct").textContent = delta;

  // Daily production status
  setBoxStatus(el("prodStatus"), statusByPct(pct));
  const prodTag = tagByPct(pct);
  el("prodTag").className = `tag ${prodTag.cls}`;
  el("prodTag").textContent = prodTag.txt;

  // Tickets + safety
  const tickets = safeNum(rec.ticketsMissed);
  const safety = safeNum(rec.safetyIncidents);

  el("ticketsValue").textContent = fmt0(tickets);
  setBoxStatus(el("ticketsStatus"), tickets>0 ? "bad" : "good");
  el("ticketsTag").className = `tag ${tickets>0 ? "bad" : "good"}`;
  el("ticketsTag").textContent = tickets>0 ? "Action" : "Good";

  el("safetyValue").textContent = fmt0(safety);
  setBoxStatus(el("safetyStatus"), safety>0 ? "bad" : "good");
  el("safetyTag").className = `tag ${safety>0 ? "bad" : "good"}`;
  el("safetyTag").textContent = safety>0 ? "Incident" : "Good";

  // Summary periods
  const { start: mStart, end: mEnd } = monthStartEnd(sel);
  const yStart = yearStart(sel);

  const mtdActual = sumActualsByDivision(mStart, sel);
  const ytdActual = sumActualsByDivision(yStart, sel);

  const mtdTarget = getMTDTargetsSum(mStart, sel);
  const ytdTarget = getYTDTargetsSum(yStart, sel);

  el("mtdMaintLine").textContent = `${fmt1(mtdActual.maint)} / ${fmt1(mtdTarget.maint)}`;
  el("mtdConstLine").textContent = `${fmt1(mtdActual.cons)} / ${fmt1(mtdTarget.cons)}`;
  el("ytdMaintLine").textContent = `${fmt1(ytdActual.maint)} / ${fmt1(ytdTarget.maint)}`;
  el("ytdConstLine").textContent = `${fmt1(ytdActual.cons)} / ${fmt1(ytdTarget.cons)}`;

  const view = el("kpiView").value;
  const mA = (view==="maint") ? mtdActual.maint : (view==="const") ? mtdActual.cons : mtdActual.total;
  const mT = (view==="maint") ? mtdTarget.maint : (view==="const") ? mtdTarget.cons : mtdTarget.total;

  const yA = (view==="maint") ? ytdActual.maint : (view==="const") ? ytdActual.cons : ytdActual.total;
  const yT = (view==="maint") ? ytdTarget.maint : (view==="const") ? ytdTarget.cons : ytdTarget.total;

  // MTD
  const mPct = mT>0 ? mA/mT : null;
  const mVar = mA - mT;
  setBoxStatus(el("mtdStatus"), statusByPct(mPct));
  el("mtdHeadline").textContent = (mT>0) ? `${fmt1(mA)} / ${fmt1(mT)} hrs` : "No target";
  el("mtdTarget").textContent = fmt1(mT);
  el("mtdActual").textContent = fmt1(mA);
  el("mtdVar").textContent = (mT>0||mA>0) ? `${mVar>=0?"+":""}${mVar.toFixed(1)}` : "—";
  el("mtdPct").textContent = (mPct!=null) ? `${(mPct*100).toFixed(1)}%` : "—";
  el("mtdBar").style.width = `${Math.min(100, Math.max(0, (mPct||0)*100))}%`;

  // YTD
  const yPct = yT>0 ? yA/yT : null;
  const yVar = yA - yT;
  setBoxStatus(el("ytdStatus"), statusByPct(yPct));
  el("ytdHeadline").textContent = (yT>0) ? `${fmt1(yA)} / ${fmt1(yT)} hrs` : "No target";
  el("ytdTarget").textContent = fmt1(yT);
  el("ytdActual").textContent = fmt1(yA);
  el("ytdVar").textContent = (yT>0||yA>0) ? `${yVar>=0?"+":""}${yVar.toFixed(1)}` : "—";
  el("ytdPct").textContent = (yPct!=null) ? `${(yPct*100).toFixed(1)}%` : "—";
  el("ytdBar").style.width = `${Math.min(100, Math.max(0, (yPct||0)*100))}%`;

  // Monthly summary
  const monthActual = sumActualsByDivision(mStart, mEnd);
  const monthTarget = getMonthTargetsSum(mStart, mEnd);
  const pctMonth = monthTarget.total>0 ? (monthActual.total/monthTarget.total) : null;

  const fb = computeFallbackDailyTargets(sel);
  el("monthFallbackTarget").textContent = fmt1(fb.monthMaint + fb.monthConst);

  const tag = tagByPct(pctMonth);
  const tagEl = el("monthSummaryTag");
  tagEl.className = `tag ${tag.cls}`;
  tagEl.textContent = `${tag.txt} (${pctMonth==null ? "—" : (pctMonth*100).toFixed(1)+"%"})`;
  setBoxStatus(el("monthSummaryBox"), statusByPct(pctMonth));

  const fill = (tId, aId, vId, pId, tVal, aVal) => {
    el(tId).textContent = fmt1(tVal);
    el(aId).textContent = fmt1(aVal);
    const v = aVal - tVal;
    el(vId).textContent = `${v>=0?"+":""}${v.toFixed(1)}`;
    const p = tVal>0 ? (aVal/tVal) : null;
    el(pId).textContent = p==null ? "—" : `${(p*100).toFixed(1)}%`;
  };
  fill("msTMaint","msAMaint","msVMaint","msPMaint", monthTarget.maint, monthActual.maint);
  fill("msTConst","msAConst","msVConst","msPConst", monthTarget.cons,  monthActual.cons);
  fill("msTTotal","msATotal","msVTotal","msPTotal", monthTarget.total, monthActual.total);

  // Month table
  renderMonthTable(mStart, mEnd, sel);
}

/* ---------------- data loading from repo ---------------- */
async function fetchText(url){
  const res = await fetch(url, { cache: "no-store" });
  if(!res.ok) throw new Error(`${url} HTTP ${res.status}`);
  return res.text();
}
async function loadCalendarFromRepo(){
  const text = await fetchText(CONFIG.CALENDAR_URL);
  calendar = parseCalendar(text);
  calendarLoaded = Object.keys(calendar).length > 0;
}
async function loadTargetsFromRepo(){
  const text = await fetchText(CONFIG.TARGETS_URL);
  targets = parseTargets(text);
  targetsLoaded = Object.keys(targets).length > 0;
}
async function reloadRepoTargets(){
  try{ await loadCalendarFromRepo(); } catch(e){ console.warn(e); calendarLoaded=false; }
  try{ await loadTargetsFromRepo(); } catch(e){ console.warn(e); targetsLoaded=false; }
}

/* ---------------- UI actions ---------------- */
function setDate(ymdStr){
  el("datePicker").value = ymdStr;
  render();
}
function addDays(ymdStr, delta){
  const d=new Date(ymdStr+"T00:00:00");
  d.setDate(d.getDate()+delta);
  return ymd(d);
}

async function refreshAll(){
  const sel = el("datePicker").value || ymd(new Date());
  const { start: mStart, end: mEnd } = monthStartEnd(sel);

  // Load targets/calendar from repo + daily actuals from Google for month range
  await reloadRepoTargets();

  try{
    el("sheetTag").className = "tag warn";
    el("sheetTag").textContent = "Loading…";
    await loadDailyFromGoogle(mStart, mEnd);
  } catch(err){
    console.error(err);
    el("sheetTag").className = "tag bad";
    el("sheetTag").textContent = "Offline";
    toast("Could not load Google Sheet data. Check Apps Script URL.");
  }

  render();
}

/* ---------------- init ---------------- */
(function init(){
  const today = ymd(new Date());
  el("datePicker").value = today;

  el("btnPrevDay").addEventListener("click", ()=> setDate(addDays(el("datePicker").value, -1)));
  el("btnNextDay").addEventListener("click", ()=> setDate(addDays(el("datePicker").value, +1)));
  el("btnMonthStart").addEventListener("click", ()=>{
    const { start } = monthStartEnd(el("datePicker").value);
    setDate(start);
  });

  el("btnRefresh").addEventListener("click", ()=> refreshAll());
  el("btnExportMonth").addEventListener("click", exportMonthCSV);

  el("datePicker").addEventListener("change", ()=> refreshAll());
  el("kpiView").addEventListener("change", render);

  // First load
  refreshAll();
})();
