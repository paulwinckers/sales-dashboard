const CONFIG = {
  SHEET_API_URL: "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE", // ends with /exec
  API_KEY: "",
  SHEET_TAB_NAME: "OpsDailyLog",
  CALENDAR_URL: "../data/calendar.csv",
  TARGETS_URL: "../data/targets.csv"
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
  setTimeout(()=>t.classList.remove("show"), 2400);
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
function yearEnd(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  return `${d.getFullYear()}-12-31`;
}
function addDays(ymdStr, delta){
  const d=new Date(ymdStr+"T00:00:00");
  d.setDate(d.getDate()+delta);
  return ymd(d);
}

function setBoxStatus(node, state){
  node.classList.remove("good","warn","bad");
  if(state) node.classList.add(state);
}
function statusByPct(pct){
  if(pct==null) return null;
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
function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

/* ---------------- CSV parsing ---------------- */
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
let targets = {};
let targetsLoaded = false;

function toMonthKey(raw){
  const s=(raw??"").toString().trim();
  if(!s) return null;

  const monthMap={jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};

  // 26-Jan
  let m=s.match(/^(\d{2})[-\s]?([A-Za-z]{3,})\.?$/);
  if(m){
    const yy=String(m[1]).padStart(2,"0");
    const mon=m[2].slice(0,3).toLowerCase();
    const mm=monthMap[mon];
    if(mm) return `20${yy}-${mm}`;
  }

  // Jan-26
  m=s.match(/^([A-Za-z]{3,})[-\s]?(\d{2})$/);
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

function isWorkdayByDayType(dayType){
  const dt = String(dayType || "").trim().toLowerCase();
  // customize as you like:
  // treat "stat"/"holiday"/"off" as non-workday for proration
  if(!dt) return null; // unknown
  if(["stat","holiday","off","closed"].includes(dt)) return false;
  // weekday/weekend are both "work-eligible" in your world
  if(["weekday","weekend","overtime"].includes(dt)) return true;
  return null;
}

function isWeekday(d){
  const day=d.getDay();
  return day!==0 && day!==6;
}

function countWorkdaysInRange(startYmd, endYmd){
  let c=0;
  for(let d=new Date(startYmd+"T00:00:00"); d<=new Date(endYmd+"T00:00:00"); d.setDate(d.getDate()+1)){
    const key = ymd(d);
    const rec = daily[key];
    const byType = isWorkdayByDayType(rec?.dayType);
    if(byType === true) { c++; continue; }
    if(byType === false) { continue; }
    // fallback if DayType unknown
    if(isWeekday(d)) c++;
  }
  return c;
}

function computeFallbackDailyTargets(selYmd){
  const d=new Date(selYmd+"T00:00:00");
  const mk=monthKeyFromDate(d);
  const t=targets[mk] || {maintenanceMonthly:0, constructionMonthly:0};

  const wd = (() => {
    // If daily log has DayType populated for the month, use that for workday count
    const { start, end } = monthStartEnd(selYmd);
    const c = countWorkdaysInRange(start, end);
    return c > 0 ? c : 0;
  })();

  const dailyMaint = wd>0 ? safeNum(t.maintenanceMonthly)/wd : 0;
  const dailyConst = wd>0 ? safeNum(t.constructionMonthly)/wd : 0;
  return { mk, wd, monthMaint:safeNum(t.maintenanceMonthly), monthConst:safeNum(t.constructionMonthly), dailyMaint, dailyConst, dailyTotal:dailyMaint+dailyConst };
}

/* ---------------- calendar targets ---------------- */
let calendar = {};
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

/* ---------------- daily log from Google ---------------- */
let daily = {}; // daily[YYYY-MM-DD] = { actualMaint, actualConst, ticketsMissed, safetyIncidents, notes, dayType, targetMaint, targetConst }
let sheetLoadedCount = 0;

function getDayRec(dateKey){
  return daily[dateKey] || { actualMaint:"", actualConst:"", ticketsMissed:"", safetyIncidents:"", notes:"", dayType:"", targetMaint:"", targetConst:"" };
}

async function loadDailyFromGoogle(rangeStart, rangeEnd){
  if(!CONFIG.SHEET_API_URL || CONFIG.SHEET_API_URL.includes("PASTE_")){
    throw new Error("Missing Apps Script URL in CONFIG.SHEET_API_URL");
  }

  const u = new URL(CONFIG.SHEET_API_URL);
  u.searchParams.set("tab", CONFIG.SHEET_TAB_NAME);
  if(rangeStart) u.searchParams.set("start", rangeStart);
  if(rangeEnd) u.searchParams.set("end", rangeEnd);
  if(CONFIG.API_KEY) u.searchParams.set("key", CONFIG.API_KEY);

  const res = await fetch(u.toString(), { cache: "no-store" });
  if(!res.ok) throw new Error(`Google endpoint HTTP ${res.status}`);
  const data = await res.json();

  const map = {};
  const rows = Array.isArray(data.rows) ? data.rows : [];
  for(const r of rows){
    const date = String(r.Date || r.date || "").slice(0,10);
    if(!date.match(/^\d{4}-\d{2}-\d{2}$/)) continue;

    map[date] = {
      actualMaint: r.ActualMaint ?? r.ActualMaintHours ?? r.MaintActual ?? "",
      actualConst: r.ActualConst ?? r.ActualConstHours ?? r.ConstActual ?? "",
      ticketsMissed: r.MissedTickets ?? r.TicketsMissed ?? r.Missed ?? "",
      safetyIncidents: r.SafetyIncidents ?? r.Incidents ?? r.Safety ?? "",
      notes: r.Notes ?? "",
      dayType: r.DayType ?? r.Daytype ?? "",
      // these are OPTIONAL: add them to sheet for exact targets
      targetMaint: r.TargetMaint ?? r.TargetMaintenance ?? "",
      targetConst: r.TargetConst ?? r.TargetConstruction ?? ""
    };
  }

  daily = map;
  sheetLoadedCount = rows.length;

  el("sheetStatus").textContent = "Connected";
}

/* ---------------- targets selection logic ---------------- */
function getDayTargets(dateKey){
  const rec = getDayRec(dateKey);

  // 1) Prefer daily targets from SHEET if provided
  const sM = safeNum(rec.targetMaint);
  const sC = safeNum(rec.targetConst);
  if((sM > 0) || (sC > 0)){
    return { maint:sM, cons:sC, total:sM+sC, source:"Sheet Targets" };
  }

  // 2) Calendar.csv
  if(calendarLoaded){
    const c = getCalendarTargetForDate(dateKey);
    if(c.has) return { maint:c.maint, cons:c.cons, total:c.total, source:"Calendar" };
  }

  // 3) Fallback monthly proration
  const fb = computeFallbackDailyTargets(dateKey);
  return { maint:fb.dailyMaint, cons:fb.dailyConst, total:fb.dailyTotal, source:"Monthly avg (prorated)" };
}

function sumTargetsForRange(startYmd, endYmd){
  // If sheet targets exist for any days in range, sum them (most accurate)
  let hasSheetTargets = false;
  let sm=0, sc=0;

  for(let d=new Date(startYmd+"T00:00:00"); d<=new Date(endYmd+"T00:00:00"); d.setDate(d.getDate()+1)){
    const k = ymd(d);
    const rec = getDayRec(k);
    const tm = safeNum(rec.targetMaint);
    const tc = safeNum(rec.targetConst);
    if(tm>0 || tc>0){
      hasSheetTargets = true;
      sm += tm; sc += tc;
    }
  }
  if(hasSheetTargets) return { maint:sm, cons:sc, total:sm+sc, source:"Sheet Targets" };

  // Else: calendar sums if available
  if(calendarLoaded){
    const c = sumCalendarTargetsByDivision(startYmd, endYmd);
    return { ...c, source:"Calendar" };
  }

  // Else: monthly avg prorated by workdays in range
  // We compute per-month totals based on targets.csv and prorate by counted workdays in that month/range
  let maint=0, cons=0;
  for(let d=new Date(startYmd+"T00:00:00"); d<=new Date(endYmd+"T00:00:00"); d.setDate(d.getDate()+1)){
    const k = ymd(d);
    const rec = getDayRec(k);
    const byType = isWorkdayByDayType(rec.dayType);
    const workEligible = (byType === true) ? true : (byType === false) ? false : isWeekday(d);
    if(!workEligible) continue;

    const fb = computeFallbackDailyTargets(k);
    maint += fb.dailyMaint;
    cons  += fb.dailyConst;
  }
  return { maint, cons, total: maint+cons, source:"Monthly avg (prorated)" };
}

/* ---------------- actuals sums ---------------- */
function sumActualsByDivision(startYmd, endYmd){
  let maint=0, cons=0;
  for(const [k, rec] of Object.entries(daily)){
    if(k < startYmd || k > endYmd) continue;
    maint += safeNum(rec.actualMaint);
    cons  += safeNum(rec.actualConst);
  }
  return { maint, cons, total: maint+cons };
}

/* ---------------- month table ---------------- */
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
      <td>${escapeHtml(rec.notes ?? "")}</td>
    `;
    tr.addEventListener("click", ()=>{
      el("datePicker").value = key;
      render();
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    tbody.appendChild(tr);
  }
}

/* ---------------- export ---------------- */
function exportMonthCSV(){
  const sel = el("datePicker").value;
  const { start: mStart, end: mEnd } = monthStartEnd(sel);

  const rows = [];
  rows.push(["Date","TargetMaint","TargetConst","TargetTotal","ActualMaint","ActualConst","ActualTotal","DeltaPct","MissedTickets","SafetyIncidents","Notes"]);

  for(let d=new Date(mStart+"T00:00:00"); d<=new Date(mEnd+"T00:00:00"); d.setDate(d.getDate()+1)){
    const key = ymd(d);
    const rec = getDayRec(key);
    const t = getDayTargets(key);

    const aM = safeNum(rec.actualMaint);
    const aC = safeNum(rec.actualConst);
    const aT = aM + aC;
    const deltaPct = t.total>0 ? (((aT/t.total)-1)*100).toFixed(1)+"%" : "";

    rows.push([key, t.maint, t.cons, t.total, aM, aC, aT, deltaPct, safeNum(rec.ticketsMissed), safeNum(rec.safetyIncidents), (rec.notes ?? "")]);
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

/* ---------------- repo loads ---------------- */
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

/* ---------------- summary range presets ---------------- */
function setSummaryPreset(preset){
  const today = ymd(new Date());
  const selected = el("datePicker").value || today;
  const d = new Date(selected+"T00:00:00");

  let start = today, end = today;

  if(preset === "thisMonth"){
    start = ymd(new Date(d.getFullYear(), d.getMonth(), 1));
    end = today;
  } else if(preset === "lastMonth"){
    const lmStart = new Date(d.getFullYear(), d.getMonth()-1, 1);
    const lmEnd = new Date(d.getFullYear(), d.getMonth(), 0);
    start = ymd(lmStart);
    end = ymd(lmEnd);
  } else if(preset === "thisYear"){
    start = `${d.getFullYear()}-01-01`;
    end = today;
  } else if(preset === "custom"){
    // keep user values
    start = el("summaryStart").value || `${d.getFullYear()}-01-01`;
    end = el("summaryEnd").value || today;
  }

  el("summaryStart").value = start;
  el("summaryEnd").value = end;

  const disable = preset !== "custom";
  el("summaryStart").disabled = disable;
  el("summaryEnd").disabled = disable;
}

function getSummaryRange(){
  const start = el("summaryStart").value;
  const end = el("summaryEnd").value;
  return { start, end };
}

/* ---------------- render ---------------- */
function render(){
  const sel = el("datePicker").value;
  if(!sel) return;

  el("displayDate").textContent = sel;
  el("monthLabel").textContent = monthKeyFromDate(new Date(sel+"T00:00:00"));

  el("targetsLabel").textContent = targetsLoaded ? "Yes" : "No";
  el("calendarLabel").textContent = calendarLoaded ? "Yes" : "No";

  const rec = getDayRec(sel);
  el("dayTypeLabel").textContent = rec.dayType ? String(rec.dayType) : "—";

  const t = getDayTargets(sel);
  const aM = safeNum(rec.actualMaint);
  const aC = safeNum(rec.actualConst);
  const aT = aM + aC;

  el("dailyTargetsLine").textContent = `${fmt1(t.maint)} / ${fmt1(t.cons)} / ${fmt1(t.total)}`;
  el("dailyActualsLine").textContent = `${fmt1(aM)} / ${fmt1(aC)} / ${fmt1(aT)}`;
  el("targetSource").textContent = t.source;

  const pct = t.total>0 ? (aT/t.total) : null;
  el("deltaPct").textContent = (t.total>0) ? `${(((aT/t.total)-1)*100).toFixed(1)}%` : "—";

  setBoxStatus(el("prodStatus"), statusByPct(pct));
  const prodTag = tagByPct(pct);
  el("prodTag").className = `tag ${prodTag.cls}`;
  el("prodTag").textContent = prodTag.txt;

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

  // month table always uses selected month
  const { start: mStart, end: mEnd } = monthStartEnd(sel);
  renderMonthTable(mStart, mEnd, sel);

  // summary uses summary range controls
  const { start: sStart, end: sEnd } = getSummaryRange();
  const view = el("kpiView").value;

  const sActual = sumActualsByDivision(sStart, sEnd);
  const sTarget = sumTargetsForRange(sStart, sEnd);

  const mA = (view==="maint") ? sActual.maint : (view==="const") ? sActual.cons : sActual.total;
  const mT = (view==="maint") ? sTarget.maint : (view==="const") ? sTarget.cons : sTarget.total;

  const mPct = mT>0 ? mA/mT : null;
  const mVar = mA - mT;

  // Use the same summary values for MTD and YTD tiles (since you requested range-driven)
  // If you want two different ranges (MTD vs YTD), we can split them later.
  setBoxStatus(el("mtdStatus"), statusByPct(mPct));
  el("mtdHeadline").textContent = (mT>0) ? `${fmt1(mA)} / ${fmt1(mT)} hrs` : "No target";
  el("mtdTarget").textContent = fmt1(mT);
  el("mtdActual").textContent = fmt1(mA);
  el("mtdVar").textContent = (mT>0||mA>0) ? `${mVar>=0?"+":""}${mVar.toFixed(1)}` : "—";
  el("mtdPct").textContent = (mPct!=null) ? `${(mPct*100).toFixed(1)}%` : "—";
  el("mtdBar").style.width = `${Math.min(100, Math.max(0, (mPct||0)*100))}%`;

  setBoxStatus(el("ytdStatus"), statusByPct(mPct));
  el("ytdHeadline").textContent = (mT>0) ? `${fmt1(mA)} / ${fmt1(mT)} hrs` : "No target";
  el("ytdTarget").textContent = fmt1(mT);
  el("ytdActual").textContent = fmt1(mA);
  el("ytdVar").textContent = (mT>0||mA>0) ? `${mVar>=0?"+":""}${mVar.toFixed(1)}` : "—";
  el("ytdPct").textContent = (mPct!=null) ? `${(mPct*100).toFixed(1)}%` : "—";
  el("ytdBar").style.width = `${Math.min(100, Math.max(0, (mPct||0)*100))}%`;

  // Monthly summary tile now also uses summary range (as requested)
  const pctMonth = mT>0 ? (mA/mT) : null;
  const tag = tagByPct(pctMonth);
  el("monthSummaryTag").className = `tag ${tag.cls}`;
  el("monthSummaryTag").textContent = `${tag.txt} (${pctMonth==null ? "—" : (pctMonth*100).toFixed(1)+"%"})`;
  setBoxStatus(el("monthSummaryBox"), statusByPct(pctMonth));

  // fallback month target display (still useful)
  const fb = computeFallbackDailyTargets(sel);
  el("monthFallbackTarget").textContent = fmt1(fb.monthMaint + fb.monthConst);

  // Fill the summary table using the summary range
  const fill = (tId, aId, vId, pId, tVal, aVal) => {
    el(tId).textContent = fmt1(tVal);
    el(aId).textContent = fmt1(aVal);
    const v = aVal - tVal;
    el(vId).textContent = `${v>=0?"+":""}${v.toFixed(1)}`;
    const p = tVal>0 ? (aVal/tVal) : null;
    el(pId).textContent = p==null ? "—" : `${(p*100).toFixed(1)}%`;
  };

  fill("msTMaint","msAMaint","msVMaint","msPMaint", sTarget.maint, sActual.maint);
  fill("msTConst","msAConst","msVConst","msPConst", sTarget.cons,  sActual.cons);
  fill("msTTotal","msATotal","msVTotal","msPTotal", sTarget.total, sActual.total);
}

/* ---------------- refresh ---------------- */
async function refreshAll(){
  const sel = el("datePicker").value || ymd(new Date());

  try { await loadCalendarFromRepo(); calendarLoaded = Object.keys(calendar).length>0; } catch { calendarLoaded=false; }
  try { await loadTargetsFromRepo(); targetsLoaded = Object.keys(targets).length>0; } catch { targetsLoaded=false; }

  // IMPORTANT: load the whole year so Jan shows even if you're in Feb
  try{
    await loadDailyFromGoogle(yearStart(sel), yearEnd(sel));
    el("sheetStatus").textContent = `Loaded (${sheetLoadedCount})`;
  } catch(err){
    console.error(err);
    el("sheetStatus").textContent = "Offline";
    toast("Google Sheet fetch failed. Check Apps Script URL.");
  }

  render();
}

/* ---------------- init ---------------- */
(function init(){
  const today = ymd(new Date());
  el("datePicker").value = today;

  // summary defaults
  el("summaryPreset").value = "thisMonth";
  setSummaryPreset("thisMonth");

  el("btnPrevDay").addEventListener("click", ()=> { el("datePicker").value = addDays(el("datePicker").value, -1); refreshAll(); });
  el("btnNextDay").addEventListener("click", ()=> { el("datePicker").value = addDays(el("datePicker").value, +1); refreshAll(); });
  el("btnMonthStart").addEventListener("click", ()=>{
    const { start } = monthStartEnd(el("datePicker").value);
    el("datePicker").value = start;
    refreshAll();
  });

  el("btnRefresh").addEventListener("click", refreshAll);
  el("btnExportMonth").addEventListener("click", exportMonthCSV);

  el("datePicker").addEventListener("change", refreshAll);

  el("summaryPreset").addEventListener("change", ()=>{
    setSummaryPreset(el("summaryPreset").value);
    render();
  });
  el("summaryStart").addEventListener("change", render);
  el("summaryEnd").addEventListener("change", render);
  el("kpiView").addEventListener("change", render);

  refreshAll();
})();
